'for i=1 to 10

Dim ExeFileName
RunAction "GetFileNameByRegular", oneIteration, "C:\Users\Administrator\Desktop\QTP要测试的版本\", "2345pinyin_v[0-9]\.[0-9]\.[0-9]*\.exe", ExeFileName
RunAction "InstallInputMethod", oneIteration, ExeFileName

Dim WshShell, autoSaveTime,TXTFileName
AutoSaveTime=10
Set WshShell=CreateObject("WScript.Shell")

TXTFileName="cikuyunxingjieguo"
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
'FileContent(TXTFileName).Check CheckPoint("sdf.txt")
wait 2

WshShell.SendKeys ("%{F4}")



FileContent("词库运行结果.txt").Check CheckPoint("词库运行结果.txt")

Dim UninstallExeFileName
RunAction "GetFileNameByRegular", oneIteration, "C:\Program Files (x86)\2345Soft\2345Pinyin\", "[0-9]\.[0-9]\.[0-9]\.[0-9]*", UninstallExeFileName
RunAction "UninstallInputMethod", oneIteration, UninstallExeFileName & "\Uninstall.exe"
 @@ hightlight id_;_396190_;_script infofile_;_ZIP::ssf67.xml_;_





'next

'wait 2
'WshShell.SendKeys TXTFileName
'WshShell.SendKeys "%s"
'wait AutoSaveTime
'
'While WshShell.AppActivate(TXTFileName)=true
'WshShell.SendKeys "s"
'wait AutoSaveTime
'Wend
'Set WshShell = Nothing

'set WshShell = WScript.CreateObject("WScript.Shell")
'WshShell.SendKeys "ABCD"
'Systemutil.run"C:\Program Files (x86)\2345chrome\2345chrome.exe"

'Window("2345加速浏览器").WinObject("Chrome_RenderWidgetHostHWND").Click 250,147
'Window("2345加速浏览器").WinObject("Chrome_RenderWidgetHostHWND").Type "哈哈哈你好啊"


'Systemutil.run"C:\Users\Administrator\Desktop\2345pinyin_kUID_v3.0.1580.exe"
'
'Dialog("2345王牌输入法v3.0安装").Click 510,303
'Dialog("2345王牌输入法v3.0安装").Click 316,316
''Dialog("2345王牌输入法v3.0安装").Click 71,381
'wait 20
'
'Dialog("2345王牌输入法v3.0安装").Click 16,378
'Dialog("2345王牌输入法v3.0安装").Click 329,317
'
''Dialog("2345王牌输入法v3.0安装").Click 25,382
'Dialog("2345王牌输入法设置向导").WinComboBox("ComboBox").Select "5"
'Dialog("2345王牌输入法设置向导").WinComboBox("ComboBox_2").Select "中"
'Dialog("2345王牌输入法设置向导").Click 412,286
'Dialog("2345王牌输入法设置向导").Click 551,403
'Dialog("2345王牌输入法设置向导").WinObject("RCWizardLittleSkins").Click 293,55
'Dialog("2345王牌输入法设置向导").Click 541,397
'Dialog("2345王牌输入法设置向导").Click 155,343
'Dialog("2345王牌输入法设置向导").Click 155,343
'Dialog("2345王牌输入法设置向导").Click 547,406
'
'Systemutil.run"E:\Notepad++\notepad++.exe"
'
''Window("Notepad++ [Administrator]").WinObject("ꛧ붅ꛧ붅敨楬潣瑰牥붦藥붦藥㖽敨楬潣瑰牥붦藥붦藥㖽").Click 363,107
''Window("Notepad++ [Administrator]").WinObject("ꛧ붅ꛧ붅敨楬潣瑰牥붦藥붦藥㖽敨楬潣瑰牥붦藥붦藥㖽").Type "啦啦啦2亲手心情烽火走走走"
''Systemutil.run"C:\Program Files (x86)\2345chrome\2345chrome.exe"
''Window("2345加速浏览器").WinObject("Chrome_RenderWidgetHostHWND").Click 250,147
''Window("2345加速浏览器").WinObject("Chrome_RenderWidgetHostHWND").Type "哈哈哈你好啊"
'
'
''Systemutil.run"C:\Program Files (x86)\2345Soft\2345Pinyin\3.0.1.1580\Uninstall.exe"
'
'
''Dialog("2345王牌输入法v3.0卸载").Click 181,336
''Dialog("2345王牌输入法v3.0卸载").Click 420,345
''Dialog("2345王牌输入法v3.0卸载").Dialog("2345好压卸载程序").Click 242,167
''wait 15
''Dialog("2345王牌输入法v3.0卸载").Click 534,339
'
'
'
'
'
'
'
'
'
























