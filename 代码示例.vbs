str="��Ϣ�����滻"

Set Wshshell=WScript.CreateObject("WScript.Shell")
WshShell.run"mshta vbscript:clipboardData.SetData("+""""+"text"+""""+","+""""&str&""""+")(close)",0,true

WshShell.run "Ⱥ�������滻"   

WScript.Sleep 1500
WshShell.SendKeys"^v"
WScript.Sleep 1500
WshShell.SendKeys "%s"
WScript.Sleep 1500
WshShell.SendKeys"%{F4}"


