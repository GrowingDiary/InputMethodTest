
' Get install file path and install it
Dim ExeFileName
RunAction "GetFileNameByRegular", oneIteration, "C:\Users\Administrator\Desktop\QTP要测试的版本\a\", "2345pinyin_v[0-9]\.[0-9]\.[0-9]*\.exe", ExeFileName
If Len(ExeFileName) = 0 Then
	Print("ERROR " & Environment("ActionName") & " - Get install file path failed")
	ExitActionIteration()
End If
RunAction "InstallInputMethod", oneIteration, ExeFileName

Dim VerifyFileName
RunAction "GetFileNameByRegular", oneIteration, "C:\Users\Administrator\Desktop\QTP要测试的版本\", "2345pinyin_v[0-9]\.[0-9]\.[0-9]*\.txt", VerifyFileName

Dim WshShell, autoSaveTime,TXTFileName
AutoSaveTime = 10
Set WshShell = CreateObject("WScript.Shell")

TXTFileName = "cikuyunxingjieguo"
WshShell.Run "notepad"
wait 2
WshShell.AppActivate "词库 - 记事本"

WshShell.SendKeys "agaxiang"
WshShell.SendKeys (" ")

WshShell.SendKeys "aozhouyaogun"
WshShell.SendKeys (" ")

WshShell.SendKeys "meisaidesibenchi"
WshShell.SendKeys (" ")

WshShell.SendKeys "^s"


wait 2
WshShell.SendKeys TXTFileName
WshShell.SendKeys (" ")
WshShell.SendKeys "%s"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
wait 2
WshShell.SendKeys ("{LEFT}")
WshShell.SendKeys ("{Enter}")
wait 2

WshShell.SendKeys ("%{F4}")



FileContent("词库运行结果.txt").Check CheckPoint("词库运行结果.txt")

Dim UninstallExeFileName
RunAction "GetFileNameByRegular", oneIteration, "C:\Program Files (x86)\2345Soft\2345Pinyin\", "[0-9]\.[0-9]\.[0-9]\.[0-9]*", UninstallExeFileName
RunAction "UninstallInputMethod", oneIteration, UninstallExeFileName & "\Uninstall.exe"
 @@ hightlight id_;_396190_;_script infofile_;_ZIP::ssf67.xml_;_

