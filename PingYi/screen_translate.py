# -*- coding: utf-8 -*-
"""
屏译 · 悬浮翻译器
用法：双击「启动悬浮翻译.vbs」运行。
- 工具栏常驻屏幕置顶：选源/目标语言、识别语言、透明度。
- 点「框选翻译」→ 拖动框选屏幕文字 → 自动 OCR + 翻译 → 译文悬浮在原处（透明度可调）。
- 译文卡片可拖动、可复制、可关闭。
"""

import os
import sys
import time
import math
import ctypes
import threading
import traceback
import html as _html
import json
import queue
from concurrent.futures import ThreadPoolExecutor

# 先设置 DPI 感知，保证框选坐标与屏幕像素一致
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

try:
    import requests
    import mss
    import numpy as np
    from PIL import Image
except Exception:
    _log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.log")
    try:
        with open(_log, "w", encoding="utf-8") as _f:
            _f.write("缺少运行依赖：\n" + traceback.format_exc())
    except Exception:
        pass
    try:
        _r = tk.Tk()
        _r.withdraw()
        messagebox.showerror("屏译启动失败", "缺少运行依赖，详情见同目录 error.log")
        _r.destroy()
    except Exception:
        pass
    raise SystemExit(1)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(APP_DIR, "error.log")

try:
    _RESAMPLE = Image.Resampling.LANCZOS
except AttributeError:
    _RESAMPLE = Image.LANCZOS

# ---------------- 语言表 ----------------
# 内置语言（代码, 显示名）。要加更多语言：程序里“语言包·教程”可一键添加，
# 或直接编辑 langpacks.json 追加 {"code":"sv","label":"瑞典语"}，重启生效。
BUILTIN_LANGS = [
    ("auto", "自动检测"), ("zh-CN", "中文（简体）"), ("zh-TW", "中文（繁体）"),
    ("en", "英语"), ("ja", "日语"), ("ko", "韩语"), ("fr", "法语"), ("de", "德语"),
    ("es", "西班牙语"), ("ru", "俄语"), ("pt", "葡萄牙语"), ("it", "意大利语"),
    ("ar", "阿拉伯语"), ("th", "泰语"), ("vi", "越南语"), ("id", "印尼语"),
    ("tr", "土耳其语"), ("nl", "荷兰语"), ("pl", "波兰语"),
    ("sv", "瑞典语"), ("da", "丹麦语"), ("fi", "芬兰语"), ("cs", "捷克语"),
    ("el", "希腊语"), ("hu", "匈牙利语"), ("ro", "罗马尼亚语"), ("uk", "乌克兰语"),
    ("he", "希伯来语"), ("hi", "印地语"),
]
LANGPACK_PATH = os.path.join(APP_DIR, "langpacks.json")


def _load_langpacks():
    try:
        with open(LANGPACK_PATH, "r", encoding="utf-8") as f:
            extra = json.load(f) or []
        out = []
        for it in extra:
            c = (it or {}).get("code")
            l = (it or {}).get("label")
            if c and l:
                out.append((str(c), str(l)))
        return out
    except Exception:
        return []


LANGS = list(BUILTIN_LANGS)
for _c, _l in _load_langpacks():
    if _c != "auto" and _c not in [c for c, _ in LANGS]:
        LANGS.append((_c, _l))

SRC_LABELS = [lbl for _, lbl in LANGS]
SRC_CODES = [code for code, _ in LANGS]
TGT_LANGS = [(c, l) for c, l in LANGS if c != "auto"]
TGT_LABELS = [lbl for _, lbl in TGT_LANGS]
TGT_CODES = [code for code, _ in TGT_LANGS]
LABEL_TO_CODE = {lbl: code for code, lbl in LANGS}


def add_langpack(code, label):
    """运行时新增一个翻译语言，并写入 langpacks.json（重启后仍在）。"""
    code = (code or "").strip()
    label = (label or "").strip()
    if not code or not label:
        raise ValueError("语言代码和名称都要填")
    if code == "auto":
        raise ValueError("auto 是保留代码，不能添加")
    if code in [c for c, _ in LANGS]:
        raise ValueError("该语言代码已存在")
    LANGS.append((code, label))
    SRC_LABELS.append(label)
    SRC_CODES.append(code)
    TGT_LANGS.append((code, label))
    TGT_LABELS.append(label)
    TGT_CODES.append(code)
    LABEL_TO_CODE[label] = code
    packs = _load_langpacks()
    packs.append({"code": code, "label": label})
    with open(LANGPACK_PATH, "w", encoding="utf-8") as f:
        json.dump(packs, f, ensure_ascii=False, indent=2)

# OCR 识别语言（RapidOCR 默认模型同时覆盖中英文；日/韩可另行加装模型）
OCR_OPTIONS = ["中文+英文", "英文"]
OCR_VALUE_TO_ENGINE = {"中文+英文": "ch", "英文": "en"}


# ---------------- 翻译 ----------------
def chunk_text(text, maxlen):
    out, cur = [], ""
    for line in text.split("\n"):
        while len(line) > maxlen:
            out.append(line[:maxlen])
            line = line[maxlen:]
        if cur and len(cur) + len(line) + 1 > maxlen:
            out.append(cur)
            cur = ""
        cur = line if not cur else cur + "\n" + line
    if cur:
        out.append(cur)
    return out


_engine_order = ["mymemory", "google"]
_preferred_engine = "mymemory"
_engine_config_path = os.path.join(APP_DIR, "config.json")


def _load_preferred_engine():
    global _preferred_engine
    try:
        with open(_engine_config_path, "r", encoding="utf-8") as f:
            eng = (json.load(f) or {}).get("engine")
            if eng in _engine_order:
                _preferred_engine = eng
    except Exception:
        pass


def _save_preferred_engine(eng):
    global _preferred_engine
    _preferred_engine = eng
    try:
        with open(_engine_config_path, "w", encoding="utf-8") as f:
            json.dump({"engine": eng}, f)
    except Exception:
        pass


_session = requests.Session()
_session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

_trans_cache = {}
_trans_cache_order = []
_TRANS_CACHE_MAX = 200


def _cache_get(key):
    return _trans_cache.get(key)


def _cache_put(key, res):
    _trans_cache[key] = dict(res)
    _trans_cache_order.append(key)
    if len(_trans_cache_order) > _TRANS_CACHE_MAX:
        _trans_cache.pop(_trans_cache_order.pop(0), None)


def google_translate(text, src, tgt):
    def one(c):
        url = ("https://translate.googleapis.com/translate_a/single"
               "?client=gtx&sl={}&tl={}&dt=t&dj=1&q={}").format(src, tgt, requests.utils.quote(c))
        r = _session.get(url, timeout=6)
        r.raise_for_status()
        j = r.json()
        return ("".join(s.get("trans", "") for s in j.get("sentences", [])), j.get("src"))

    detected, parts = None, []
    with ThreadPoolExecutor(max_workers=4) as ex:
        for p, det in ex.map(one, chunk_text(text, 4500)):
            parts.append(p)
            if det and not detected:
                detected = det
    out = "".join(parts)
    if not out:
        raise RuntimeError("Google 返回为空")
    return {"text": out, "provider": "Google 在线", "detected": detected}


def mymemory_translate(text, src, tgt):
    mmsrc = "Autodetect" if src == "auto" else src

    def one(c):
        url = "https://api.mymemory.translated.net/get?q={}&langpair={}".format(
            requests.utils.quote(c), requests.utils.quote("{}|{}".format(mmsrc, tgt)))
        r = _session.get(url, timeout=10)
        r.raise_for_status()
        j = r.json()
        if j.get("quotaFinished"):
            raise RuntimeError("MyMemory 今日免费额度已用完")
        status = j.get("responseStatus")
        if status and int(status) >= 400:
            raise RuntimeError("MyMemory {}".format(status))
        rd = j.get("responseData") or {}
        return (_html.unescape(rd.get("translatedText", "")), rd.get("detectedLanguage"))

    detected, parts = None, []
    with ThreadPoolExecutor(max_workers=4) as ex:
        for p, det in ex.map(one, chunk_text(text, 450)):
            parts.append(p)
            if det and not detected:
                detected = det
    out = "".join(parts)
    if not out:
        raise RuntimeError("MyMemory 返回为空")
    return {"text": out, "provider": "MyMemory 在线",
            "detected": detected or (None if src == "auto" else src)}


def _base_lang(code):
    if not code:
        return ""
    return str(code).split("-")[0].lower()


def _same_lang(a, b):
    ba, bb = _base_lang(a), _base_lang(b)
    return bool(ba) and ba == bb


def _detect_lang_simple(s):
    """本地粗判语言，用于同框多语言分区翻译。"""
    has_kana = has_hangul = has_cjk = has_cyr = has_arab = has_thai = has_latin = False
    for ch in s:
        o = ord(ch)
        if 0x3040 <= o <= 0x30FF:
            has_kana = True
        elif 0xAC00 <= o <= 0xD7A3 or 0x1100 <= o <= 0x11FF:
            has_hangul = True
        elif 0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF:
            has_cjk = True
        elif 0x0400 <= o <= 0x04FF:
            has_cyr = True
        elif 0x0600 <= o <= 0x06FF:
            has_arab = True
        elif 0x0E00 <= o <= 0x0E7F:
            has_thai = True
        elif (65 <= o <= 90) or (97 <= o <= 122):
            has_latin = True
    if has_kana:
        return "ja"
    if has_hangul:
        return "ko"
    if has_cjk:
        return "zh"
    if has_cyr:
        return "ru"
    if has_arab:
        return "ar"
    if has_thai:
        return "th"
    if has_latin:
        return "en"
    return None


def _translate_with_engines(text, src, tgt):
    order = [_preferred_engine] + [e for e in _engine_order if e != _preferred_engine]
    errors = []
    for eng in order:
        try:
            if eng == "google":
                res = google_translate(text, src, tgt)
            else:
                res = mymemory_translate(text, src, tgt)
            _save_preferred_engine(eng)
            return res
        except Exception as e:
            errors.append("{}: {}".format(eng, e))
    raise RuntimeError(" | ".join(errors) or "无可用翻译引擎")


def _translate_mixed(text, tgt):
    """同框多语言：按行粗判语言，按连续同语言分组翻译，并保持原有顺序。"""
    lines = text.split("\n")
    groups = []
    for ln in lines:
        lang = _detect_lang_simple(ln)
        if groups and groups[-1][0] == lang:
            groups[-1][1].append(ln)
        else:
            groups.append([lang, [ln]])

    providers, detected, out_lines, errs = [], [], [], []
    for lang, ls in groups:
        joined = "\n".join(ls)
        if not joined.strip():
            out_lines.extend(ls)
            continue
        # 该段语言就是目标语言：保留原文
        if lang and _same_lang(lang, tgt):
            out_lines.extend(ls)
            continue
        try:
            req_src = "auto" if (lang is None or lang == "en") else lang
            res = _translate_with_engines(joined, req_src, tgt)
            providers.append(res.get("provider", ""))
            if res.get("detected"):
                detected.append(res["detected"])
            if res.get("detected") and _same_lang(res["detected"], tgt):
                out_lines.extend(ls)
            else:
                out_lines.extend(res["text"].split("\n"))
        except Exception as e:
            errs.append(str(e))
            out_lines.extend(ls)  # 该段失败保留原文，不影响其它段

    if not providers and errs:
        raise RuntimeError("翻译失败：{}".format(" | ".join(errs[:3])))
    uni_prov = set(p for p in providers if p)
    if not providers:
        provider = "同语言·已跳过"
    elif len(uni_prov) == 1:
        provider = uni_prov.pop()
    else:
        provider = "在线引擎 · 分段"
    uni_det = set(d for d in detected if d)
    det = (uni_det.pop() if len(uni_det) == 1 else ("multi" if uni_det else None))
    return {"text": "\n".join(out_lines), "provider": provider, "detected": det}


def translate_text(text, src, tgt):
    # 界面下拉框存的是“显示名”，这里统一转成代码
    src = LABEL_TO_CODE.get(src, src) or "auto"
    tgt = LABEL_TO_CODE.get(tgt, tgt) or "zh-CN"
    text = (text or "").strip()
    if not text:
        raise ValueError("没有可翻译的文字")
    key = (src, tgt, text)
    hit = _cache_get(key)
    if hit:
        return dict(hit)

    # 明确源语言且等于目标：直接返回原文，不请求
    if src != "auto" and _same_lang(src, tgt):
        res = {"text": text, "provider": "同语言·已跳过", "detected": src}
        _cache_put(key, res)
        return res

    if src == "auto":
        res = _translate_mixed(text, tgt)
    else:
        res = _translate_with_engines(text, src, tgt)
        if res.get("detected") and _same_lang(res["detected"], tgt):
            res = {"text": text, "provider": res.get("provider", "") + " · 与目标同语言",
                   "detected": res["detected"]}

    _cache_put(key, res)
    return res


# ---------------- OCR ----------------
ocr_lock = threading.Lock()
_ocr_engine = None
_ocr_ready = False


def get_ocr_engine():
    global _ocr_engine, _ocr_ready
    with ocr_lock:
        if _ocr_engine is None:
            from rapidocr_onnxruntime import RapidOCR
            custom = os.path.join(APP_DIR, "models", "rec.onnx")
            try:
                if os.path.isfile(custom):
                    _ocr_engine = RapidOCR(rec_model_path=custom)
                else:
                    _ocr_engine = RapidOCR()
            except Exception:
                _ocr_engine = RapidOCR()
        _ocr_ready = True
        return _ocr_engine


def warmup_ocr():
    try:
        get_ocr_engine()
    except Exception as e:
        print("OCR warmup failed:", e)


def do_ocr_lines(img):
    """逐行识别，返回 [{'text':..., 'box':(x0,y0,x1,y1)}]，box 为截图内像素坐标。"""
    if not _ocr_ready:
        get_ocr_engine()
    img = img.convert("RGB")
    w, h = img.size
    scale = 1.0
    maxside = max(w, h)
    if maxside > 1280:
        scale = 1280.0 / maxside
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), _RESAMPLE)
    arr = np.array(img)[:, :, ::-1].copy()  # RGB -> BGR 连续数组
    result, _ = get_ocr_engine()(arr)
    out = []
    if not result:
        return out
    inv = (1.0 / scale) if scale else 1.0
    for item in result:
        try:
            box, t = item[0], item[1]
        except Exception:
            continue
        if not t:
            continue
        try:
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
        except Exception:
            continue
        x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
        out.append({"text": str(t).strip(),
                    "box": (x0 * inv, y0 * inv, x1 * inv, y1 * inv)})
    out.sort(key=lambda d: (round(d["box"][1] / 8.0), d["box"][0]))
    return out


def do_ocr(img):
    return "\n".join(d["text"] for d in do_ocr_lines(img))


# ---------------- 主程序 ----------------
FONT = ("Microsoft YaHei", 10)
FONT_SMALL = ("Microsoft YaHei", 9)
FONT_BIG = ("Microsoft YaHei", 12)
FONT_TITLE = ("Microsoft YaHei", 12, "bold")

# 配色
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
PANEL = "#1e293b"
PANEL_2 = "#0f172a"
BORDER = "#334155"
MUTED = "#94a3b8"

# 收起小球
BALL_SIZE = 52
BALL_X = 6
TRANSPARENT_KEY = "#010203"


def screen_physical_size():
    try:
        u = ctypes.windll.user32
        return u.GetSystemMetrics(0), u.GetSystemMetrics(1)
    except Exception:
        return None


class OverlayApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("屏译")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.96)
        self.root.configure(bg=PANEL)

        self.opacity = 0.85
        self.result_windows = []
        self.src_var = tk.StringVar(value="auto")
        self.tgt_var = tk.StringVar(value="zh-CN")
        self.ocr_var = tk.StringVar(value=OCR_OPTIONS[0])

        self._ball_y = None
        self._collapsed = False
        self._toolbar_pos = None
        self._adv_visible = False
        self._mv = None
        self.overlay_ttl_ms = 10000
        self._overlay_after = None
        self._overlay_fg = None
        self._check_fg_after = None

        self._build_toolbar()
        self.root.attributes("-alpha", self.opacity)
        self.root.update_idletasks()
        self._place_expanded()
        self._build_ball()

        try:
            self.root.bind_all("<Escape>", lambda e: self._clear_overlays())
        except Exception:
            pass

        # 后台预热 OCR 模型（首次会联网下载，约十几秒）
        threading.Thread(target=warmup_ocr, daemon=True).start()

        # 结果队列 + 主线程轮询（避免后台线程直接触碰 Tk）
        self._result_queue = queue.Queue()
        self._pending_bbox = None
        self.root.after(100, self._poll_queue)

    # ---------- 工具栏：紧凑（默认）+ 展开 ----------
    def _build_toolbar(self):
        outer = tk.Frame(self.root, bg=BORDER)
        outer.pack(fill="both", expand=True)
        bar = tk.Frame(outer, bg=PANEL)
        bar.pack(fill="both", expand=True, padx=1, pady=1)

        accent = tk.Frame(bar, bg=ACCENT, height=3)
        accent.pack(fill="x")

        # 紧凑行：正常状态只显示“框选翻译”+“展开”
        compact = tk.Frame(bar, bg=PANEL)
        compact.pack(fill="x", padx=10, pady=8)
        self._mk_btn(compact, "框选翻译", self.start_selection, primary=True).pack(side="left")
        self.expand_btn = self._mk_btn(compact, "展开 ▾", self.toggle_adv, primary=False)
        self.expand_btn.pack(side="left", padx=(8, 0))
        self._mk_btn(compact, "收起", self.collapse, primary=False).pack(side="left", padx=(8, 0))

        # 展开面板（默认隐藏）
        self.adv = tk.Frame(bar, bg=PANEL)

        row2 = tk.Frame(self.adv, bg=PANEL)
        row2.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(row2, text="源", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left")
        self.src_combo = ttk.Combobox(row2, textvariable=self.src_var, values=SRC_LABELS,
                                      state="readonly", width=10, font=FONT)
        self.src_combo.current(0)
        self.src_combo.pack(side="left", padx=(4, 6))
        tk.Label(row2, text="→", bg=PANEL, fg="#60a5fa", font=FONT_BIG).pack(side="left")
        self.tgt_combo = ttk.Combobox(row2, textvariable=self.tgt_var, values=TGT_LABELS,
                                      state="readonly", width=10, font=FONT)
        self.tgt_combo.current(TGT_CODES.index("zh-CN"))
        self.tgt_combo.pack(side="left", padx=(6, 10))
        tk.Label(row2, text="识别", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left")
        self.ocr_combo = ttk.Combobox(row2, textvariable=self.ocr_var, values=OCR_OPTIONS,
                                      state="readonly", width=8, font=FONT)
        self.ocr_combo.current(0)
        self.ocr_combo.pack(side="left", padx=(4, 0))

        row3 = tk.Frame(self.adv, bg=PANEL)
        row3.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(row3, text="透明", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left")
        self.opacity_scale = tk.Scale(row3, from_=0.20, to=1.0, resolution=0.05,
                                      orient="horizontal", length=76, showvalue=False,
                                      command=self._on_opacity, bg=PANEL, fg="#ffffff",
                                      highlightthickness=0, troughcolor="#374151",
                                      activebackground=ACCENT, sliderlength=12)
        self.opacity_scale.set(self.opacity)
        self.opacity_scale.pack(side="left", padx=(4, 10))
        tk.Label(row3, text="停留", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left")
        self.ttl_scale = tk.Scale(row3, from_=3, to=60, resolution=1,
                                  orient="horizontal", length=60, showvalue=True,
                                  command=self._on_ttl, bg=PANEL, fg="#ffffff",
                                  highlightthickness=0, troughcolor="#374151",
                                  activebackground=ACCENT, sliderlength=12)
        self.ttl_scale.set(10)
        self.ttl_scale.pack(side="left", padx=(4, 10))

        row4 = tk.Frame(self.adv, bg=PANEL)
        row4.pack(fill="x", padx=10, pady=(0, 8))
        self._mk_btn(row4, "其他语言", self.open_langpack, primary=False).pack(side="left")
        self._mk_btn(row4, "×", self.quit, primary=False).pack(side="right")

        # 右键拖动可改变工具栏位置
        self._enable_move()

    def toggle_adv(self):
        if self._adv_visible:
            self.adv.pack_forget()
            self.expand_btn.configure(text="展开 ▾")
            self._adv_visible = False
        else:
            self.adv.pack(fill="x")
            self.expand_btn.configure(text="收起面板 ▴")
            self._adv_visible = True
        self._clamp_window()

    def _clamp_window(self):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()
        x = max(0, min(sw - w, self.root.winfo_x()))
        y = max(0, min(sh - h, self.root.winfo_y()))
        self.root.geometry("+{}+{}".format(int(x), int(y)))

    def _enable_move(self):
        def walk(w):
            try:
                w.bind("<Button-3>", self._mv_start)
                w.bind("<B3-Motion>", self._mv_drag)
                w.bind("<ButtonRelease-3>", self._mv_end)
            except Exception:
                pass
            for c in w.winfo_children():
                walk(c)
        walk(self.root)

    def _mv_start(self, e):
        self._mv = {"dx": e.x_root - self.root.winfo_x(),
                    "dy": e.y_root - self.root.winfo_y()}

    def _mv_drag(self, e):
        if not self._mv:
            return
        x = e.x_root - self._mv["dx"]
        y = e.y_root - self._mv["dy"]
        self.root.geometry("+{}+{}".format(int(x), int(y)))

    def _mv_end(self, e):
        self._toolbar_pos = (self.root.winfo_x(), self.root.winfo_y())
        self._mv = None

    def refresh_combos(self):
        try:
            self.src_combo["values"] = SRC_LABELS
            self.tgt_combo["values"] = TGT_LABELS
        except Exception:
            pass

    def open_langpack(self):
        w = tk.Toplevel(self.root)
        w.title("其他语言")
        w.configure(bg=PANEL)
        w.attributes("-topmost", True)
        w.resizable(False, False)
        w.geometry("+{}+{}".format(self.root.winfo_x() + 20, self.root.winfo_y() + 40))

        txt = ("一、翻译成更多语言（改完立即生效）\n"
               "  · 在“目标”下拉框里选择要翻译成的语言。\n"
               "  · 常用代码：en 英 / ja 日 / ko 韩 / fr 法 / de 德 / es 西 / ru 俄 / ar 阿 /\n"
               "    it 意 / pt 葡 / th 泰 / vi 越 / nl 荷 / pl 波 / sv 瑞典 / da 丹麦 / fi 芬兰。\n"
               "  · 还想加别的？在下面填“代码+名称”点添加，会写进 langpacks.json，重启仍在。\n\n"
               "二、识别更多文字（OCR）\n"
               "  · 内置识别模型只认“中文+英文”，日文/韩文/俄文等默认识别不出。\n"
               "  · 要识别这些文字：去 RapidOCR 项目下载对应语言的识别模型（.onnx + 字典），\n"
               "    放进程序目录的 models 文件夹，重启即可（README 有详细步骤）。")
        tk.Label(w, text=txt, bg=PANEL, fg="#cbd5e1", font=FONT_SMALL, justify="left",
                 anchor="w", padx=14, pady=10, wraplength=560).pack(fill="x")

        form = tk.Frame(w, bg=PANEL)
        form.pack(fill="x", padx=14, pady=(0, 8))
        tk.Label(form, text="代码", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left")
        code_var = tk.StringVar()
        tk.Entry(form, textvariable=code_var, width=7, font=FONT, bg="#0f172a",
                 fg="#e2e8f0", insertbackground="#e2e8f0", relief="flat").pack(
                 side="left", padx=(4, 10), ipady=2)
        tk.Label(form, text="名称", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left")
        label_var = tk.StringVar()
        tk.Entry(form, textvariable=label_var, width=12, font=FONT, bg="#0f172a",
                 fg="#e2e8f0", insertbackground="#e2e8f0", relief="flat").pack(
                 side="left", padx=(4, 10), ipady=2)

        def do_add():
            try:
                add_langpack(code_var.get(), label_var.get())
                self.refresh_combos()
                self.toast("已添加语言包，下拉框已更新")
                w.destroy()
            except Exception as e:
                self.toast("添加失败：{}".format(e), 6000)

        self._mk_btn(w, "添加到语言包", do_add, primary=True).pack(pady=(0, 12))

    def _mk_btn(self, parent, text, cmd, primary=False):
        bg = ACCENT if primary else "#334155"
        abg = ACCENT_DARK if primary else "#475569"
        return tk.Button(parent, text=text, command=cmd, bg=bg, fg="#ffffff",
                         activebackground=abg, activeforeground="#ffffff",
                         relief="flat", bd=0, font=FONT, cursor="hand2",
                         padx=12, pady=3)

    def _place_expanded(self):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()
        if getattr(self, "_toolbar_pos", None):
            x, y = self._toolbar_pos
        else:
            x, y = (sw - w) // 2, 14
        x = max(0, min(sw - w, int(x)))
        y = max(0, min(sh - h, int(y)))
        self.root.geometry("+{}+{}".format(int(x), int(y)))

    # ---------- 收起：左侧球形（仅可上下移动） ----------
    def _build_ball(self):
        self.ball = tk.Toplevel(self.root)
        self.ball.overrideredirect(True)
        self.ball.attributes("-topmost", True)
        try:
            self.ball.attributes("-transparentcolor", TRANSPARENT_KEY)
        except Exception:
            pass
        self.ball.configure(bg=TRANSPARENT_KEY)
        self.ball.geometry("{}x{}+{}+{}".format(BALL_SIZE, BALL_SIZE, BALL_X, self._ball_default_y()))

        cv = tk.Canvas(self.ball, width=BALL_SIZE, height=BALL_SIZE, bg=TRANSPARENT_KEY,
                       highlightthickness=0, bd=0)
        cv.pack()
        pad = 3
        cv.create_oval(pad, pad, BALL_SIZE - pad, BALL_SIZE - pad,
                       fill=ACCENT, outline=ACCENT_DARK, width=2)
        cv.create_text(BALL_SIZE // 2, BALL_SIZE // 2, text="译", fill="#ffffff",
                       font=("Microsoft YaHei", 17, "bold"))
        cv.bind("<Button-1>", self._ball_press)
        cv.bind("<B1-Motion>", self._ball_move)
        cv.bind("<ButtonRelease-1>", self._ball_release)
        cv.bind("<Button-3>", lambda e: self.expand())
        self._ball_canvas = cv
        self.ball.withdraw()

    def _ball_default_y(self):
        return int(self.root.winfo_screenheight() * 0.42)

    def _ball_press(self, e):
        self._ball_drag = {"y0": e.y_root, "wy": self.ball.winfo_y(), "moved": False}

    def _ball_move(self, e):
        d = getattr(self, "_ball_drag", None)
        if not d:
            return
        dy = e.y_root - d["y0"]
        if abs(dy) > 4:
            d["moved"] = True
        sh = self.root.winfo_screenheight()
        newy = max(0, min(sh - BALL_SIZE, d["wy"] + dy))
        self.ball.geometry("+{}+{}".format(BALL_X, newy))

    def _ball_release(self, e):
        d = getattr(self, "_ball_drag", None)
        if not d:
            return
        self._ball_y = self.ball.winfo_y()
        self._ball_drag = None
        if not d["moved"]:
            self.expand()

    def collapse(self):
        self._collapsed = True
        self._clear_overlays()
        sh = self.root.winfo_screenheight()
        y = self._ball_y if self._ball_y is not None else self._ball_default_y()
        y = max(0, min(sh - BALL_SIZE, y))
        self.root.withdraw()
        self.ball.geometry("+{}+{}".format(BALL_X, y))
        self.ball.deiconify()
        self.ball.lift()

    def expand(self):
        self._collapsed = False
        try:
            self.ball.withdraw()
        except Exception:
            pass
        self.root.deiconify()
        self.root.lift()
        self._place_expanded()

    def _on_ttl(self, val):
        try:
            self.overlay_ttl_ms = int(float(val)) * 1000
        except Exception:
            pass

    def _on_opacity(self, val):
        self.opacity = float(val)
        try:
            self.root.attributes("-alpha", self.opacity)
        except Exception:
            pass
        for w in self.result_windows:
            try:
                w.set_opacity(self.opacity)
            except Exception:
                pass

    # ---------- 框选 ----------
    def start_selection(self):
        self._clear_overlays()
        self.root.withdraw()
        self.root.update_idletasks()
        time.sleep(0.08)

        sel = tk.Toplevel(self.root)
        sel.overrideredirect(True)
        sel.attributes("-topmost", True)
        sel.attributes("-alpha", 0.30)
        sw, sh = sel.winfo_screenwidth(), sel.winfo_screenheight()
        sel.geometry("{}x{}+0+0".format(sw, sh))

        canvas = tk.Canvas(sel, bg="#000000", highlightthickness=0, cursor="crosshair")
        canvas.pack(fill="both", expand=True)
        canvas.create_text(sw // 2, 48, text="拖动框选要翻译的区域，按 Esc 取消",
                           fill="#ffffff", font=("Microsoft YaHei", 15, "bold"))
        state = {"start": None, "rect": None}

        def on_down(e):
            state["start"] = (e.x, e.y)
            if state["rect"]:
                canvas.delete(state["rect"])
            state["rect"] = canvas.create_rectangle(e.x, e.y, e.x, e.y,
                                                    outline="#3b82f6", width=2)

        def on_move(e):
            if state["start"]:
                x0, y0 = state["start"]
                canvas.coords(state["rect"], x0, y0, e.x, e.y)

        def finish(e):
            if not state["start"]:
                return
            x0, y0 = state["start"]
            x1, y1 = e.x, e.y
            sel.destroy()
            self.root.deiconify()
            self._place_expanded()
            if abs(x1 - x0) < 8 or abs(y1 - y0) < 8:
                self.toast("区域太小，请重新框选")
                return
            logical = (x0, y0, x1, y1)
            phys = self._to_physical(logical, sw, sh)
            self.process(logical, phys)

        def cancel(_=None):
            sel.destroy()
            self.root.deiconify()
            self._place_expanded()

        canvas.bind("<Button-1>", on_down)
        canvas.bind("<B1-Motion>", on_move)
        canvas.bind("<ButtonRelease-1>", finish)
        sel.bind("<Escape>", cancel)
        canvas.focus_set()

    def _to_physical(self, logical, sw, sh):
        ps = screen_physical_size()
        x0, y0, x1, y1 = logical
        if ps:
            pw, ph = ps
            sx, sy = pw / sw, ph / sh
            x0, y0, x1, y1 = x0 * sx, y0 * sy, x1 * sx, y1 * sy
        return tuple(int(v) for v in (x0, y0, x1, y1))

    # ---------- 处理流水线 ----------
    def process(self, logical_bbox, phys_bbox):
        self._show_processing(logical_bbox)
        # 在主线程读取界面变量，避免后台线程触碰 Tk
        src = self.src_var.get() or "auto"
        tgt = self.tgt_var.get() or "zh-CN"
        self._pending_bbox = logical_bbox

        def work():
            img = self._capture(phys_bbox)
            items = do_ocr_lines(img)
            if not items:
                raise RuntimeError("未识别到文字，可尝试更换识别语言或重新框选")

            px1, py1, px2, py2 = phys_bbox
            px1, px2 = sorted((px1, px2))
            py1, py2 = sorted((py1, py2))

            trans = self._translate_lines([it["text"] for it in items], src, tgt)

            lines, provs = [], []
            for it, tr in zip(items, trans):
                b = it["box"]
                lines.append({"text": it["text"], "trans": tr.get("text", ""),
                              "bbox": (px1 + b[0], py1 + b[1], px1 + b[2], py1 + b[3])})
                if tr.get("provider"):
                    provs.append(tr["provider"])
            uniq = set(p for p in provs if p)
            provider = (uniq.pop() if len(uniq) == 1 else ("在线引擎 · 分段" if uniq else ""))
            return {"lines": lines, "provider": provider}

        self._run_async(work)

    def _translate_lines(self, texts, src, tgt):
        """逐行翻译：先整块翻一次，行数吻合就拆分对应；否则并行逐行翻。"""
        n = len(texts)
        if n == 0:
            return []
        try:
            whole = translate_text("\n".join(texts), src, tgt)
            tl = whole["text"].split("\n")
            if len(tl) == n and all(x.strip() for x in tl):
                return [{"text": tl[i], "provider": whole.get("provider", ""),
                         "detected": whole.get("detected")} for i in range(n)]
        except Exception:
            pass

        def one(t):
            try:
                return translate_text(t, src, tgt)
            except Exception as e:
                return {"text": t, "provider": "", "detected": None, "_err": str(e)}

        with ThreadPoolExecutor(max_workers=8) as ex:
            results = list(ex.map(one, texts))
        if results and all(r.get("_err") for r in results):
            raise RuntimeError(results[0]["_err"])
        return results

    def _capture(self, bbox):
        x1, y1, x2, y2 = bbox
        x1, x2 = sorted((x1, x2))
        y1, y2 = sorted((y1, y2))
        w, h = max(1, x2 - x1), max(1, y2 - y1)
        with mss.mss() as sct:
            shot = sct.grab({"left": x1, "top": y1, "width": w, "height": h})
        return Image.frombytes("RGB", shot.size, shot.rgb)

    def _run_async(self, fn):
        def wrapper():
            try:
                res = fn()
                self._result_queue.put((None, res))
            except Exception as e:
                self._log_error("后台任务失败", e)
                self._result_queue.put((e, None))
        threading.Thread(target=wrapper, daemon=True).start()

    def _log_error(self, what, e):
        try:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write("[{}] {}\n".format(time.strftime("%H:%M:%S"), what))
                f.write("".join(traceback.format_exception(type(e), e, e.__traceback__)))
                f.write("\n")
        except Exception:
            pass

    def _poll_queue(self):
        try:
            while True:
                err, res = self._result_queue.get_nowait()
                self._on_result(err, res)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_result(self, err, res):
        self._hide_processing()
        bbox = self._pending_bbox or (0, 0, 0, 0)
        self._pending_bbox = None
        if err:
            self._log_error("翻译流程失败", err)
            self.toast("翻译失败：{}".format(err), 9000)
            return
        self.show_result(res, bbox)

    # ---------- 处理中提示 ----------
    def _show_processing(self, bbox):
        self._processing = tk.Toplevel(self.root)
        self._processing.overrideredirect(True)
        self._processing.attributes("-topmost", True)
        self._processing.attributes("-alpha", 0.92)
        tk.Label(self._processing, text="正在识别并翻译…", bg="#1f2937", fg="#ffffff",
                 font=FONT_BIG, padx=16, pady=10).pack()
        x = int((bbox[0] + bbox[2]) / 2) - 90
        y = int((bbox[1] + bbox[3]) / 2) - 24
        self._processing.geometry("+{}+{}".format(max(0, x), max(0, y)))

    def _hide_processing(self):
        w = getattr(self, "_processing", None)
        if w:
            try:
                w.destroy()
            except Exception:
                pass
            self._processing = None

    # ---------- 原位译文浮层 ----------
    def _clear_overlays(self):
        for w in list(self.result_windows):
            try:
                w.destroy()
            except Exception:
                pass
        self.result_windows = []
        if getattr(self, "_overlay_after", None):
            try:
                self.root.after_cancel(self._overlay_after)
            except Exception:
                pass
            self._overlay_after = None
        if getattr(self, "_check_fg_after", None):
            try:
                self.root.after_cancel(self._check_fg_after)
            except Exception:
                pass
            self._check_fg_after = None

    def show_result(self, res, bbox):
        self._clear_overlays()
        shown = 0
        for ln in (res.get("lines") or []):
            tr = (ln.get("trans") or "").strip()
            txt = (ln.get("text") or "").strip()
            if not tr or tr == txt:
                continue  # 无译文，或与原文相同（同语言）则不加浮层
            ov = TransOverlay(self, ln)
            self.result_windows.append(ov)
            shown += 1
        if shown == 0:
            self.toast("没有需要翻译的内容（或与目标语言相同）")
            return
        self._start_overlay_lifetime()

    def _foreground_hwnd(self):
        try:
            return int(ctypes.windll.user32.GetForegroundWindow())
        except Exception:
            return 0

    def _own_hwnds(self):
        ids = set()
        try:
            ids.add(int(self.root.winfo_id()))
        except Exception:
            pass
        for w in self.result_windows:
            try:
                ids.add(int(w.winfo_id()))
            except Exception:
                pass
        return ids

    def _start_overlay_lifetime(self):
        # 记录当前前台窗口（通常是浏览器/页面），切换或重载后自动清除
        self._overlay_fg = self._foreground_hwnd()
        try:
            self._overlay_after = self.root.after(int(self.overlay_ttl_ms), self._clear_overlays)
        except Exception:
            self._overlay_after = None
        self._check_overlay_fg()

    def _check_overlay_fg(self):
        self._check_fg_after = None
        if not self.result_windows:
            return
        fg = self._foreground_hwnd()
        own = self._own_hwnds()
        if fg and self._overlay_fg and fg != self._overlay_fg and fg not in own:
            # 前台窗口变了（切到别的窗口、页面重载等）→ 收起译文
            self._clear_overlays()
            return
        self._check_fg_after = self.root.after(400, self._check_overlay_fg)

    def toast(self, msg, ms=2600):
        w = tk.Toplevel(self.root)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        w.attributes("-alpha", 0.95)
        tk.Label(w, text=str(msg), bg="#111827", fg="#ffffff", font=FONT,
                 padx=16, pady=8, wraplength=440).pack()
        w.update_idletasks()
        sw = self.root.winfo_screenwidth()
        x = max(0, (sw - w.winfo_reqwidth()) // 2)
        w.geometry("+{}+{}".format(int(x), 86))
        w.after(ms, w.destroy)

    def quit(self):
        try:
            self.root.destroy()
        except Exception:
            pass


class TransOverlay(tk.Toplevel):
    """贴在原文字位置上的单行译文浮层（不截图、不聚成一个大卡片）。"""
    def __init__(self, app, item):
        super().__init__(app.root)
        self.app = app
        self.item = item
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        try:
            self.attributes("-alpha", app.opacity)
        except Exception:
            pass

        bx0, by0, bx1, by1 = item.get("bbox", (0, 0, 200, 24))
        wpx = max(40, int(bx1 - bx0))
        hpx = max(12, int(by1 - by0))
        fs = max(9, min(20, int(hpx * 0.72)))
        wrap = max(140, min(640, wpx + 60))

        self.configure(bg="#fff7c2")
        lbl = tk.Label(self, text=item.get("trans", ""), bg="#fff7c2", fg="#111827",
                       font=("Microsoft YaHei", fs), justify="left", anchor="w",
                       wraplength=wrap, padx=4, pady=2, bd=0)
        lbl.pack()
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = min(max(0, int(bx0)), max(0, sw - w))
        y = min(max(0, int(by0)), max(0, sh - h))
        self.geometry("+{}+{}".format(x, y))

        for wd in (self, lbl):
            wd.bind("<Button-1>", self._copy)
            wd.bind("<Button-3>", self._close)

    def _copy(self, e=None):
        try:
            self.app.root.clipboard_clear()
            self.app.root.clipboard_append(self.item.get("trans", ""))
            self.app.root.update()
            self.app.toast("已复制")
        except Exception:
            pass
        return "break"

    def _close(self, e=None):
        try:
            if self in self.app.result_windows:
                self.app.result_windows.remove(self)
        except Exception:
            pass
        self.destroy()
        return "break"

    def set_opacity(self, v):
        try:
            self.attributes("-alpha", v)
        except Exception:
            pass


def main():
    try:
        _load_preferred_engine()
        app = OverlayApp()
        app.root.mainloop()
    except Exception:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        try:
            messagebox.showerror("屏译启动失败", "程序出错，详情见同目录 error.log")
        except Exception:
            pass


if __name__ == "__main__":
    main()
