Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 500
  For i = 1 To 10
  WshShell.SendKeys "f"
Next