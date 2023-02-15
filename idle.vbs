'this script prevents the pc from sleeping.
'double taps Numlock every 5 minutes.

Dim objResult

Set objShell = WScript.CreateObject("WScript.Shell")    

Do While True
  objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
  Wscript.Sleep (300000)
Loop