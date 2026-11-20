' PingYi - Screen Translate launcher (silent, double-click this file)
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
script = Chr(34) & dir & "\screen_translate.py" & Chr(34)

py = "C:\Python314\pythonw.exe"
If Not fso.FileExists(py) Then
    py = "pythonw.exe"
End If

sh.Run Chr(34) & py & Chr(34) & " " & script, 0, False
