'for i=1 to 10

Function GetFileNameByRegular(baseDir, regExp)
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Pattern = regExp
	
	Dim oFolders
	Dim oRoot, oFolder
    Dim oFiles, oFile
    Dim colMatches
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oRoot = FSO.GetFolder(baseDir)
    Set oFolders = oRoot.SubFolders
    
    For Each oFolder In oFolders
    	Set colMatches = objRegEx.Execute(oFolder.Path)
		If colMatches.Count > 0 Then
			GetFileNameByRegular = oFolder.Path
			Exit Function
		End If
    Next
	
	Set oFiles = oRoot.Files
	For Each oFile In oFiles
    	Set colMatches = objRegEx.Execute(oFile.Name)
		If colMatches.Count > 0 Then
			GetFileNameByRegular = oFile.Path
			Exit Function
		End If
    Next
	
	GetFileNameByRegular = baseDir
End Function

Systemutil.run GetFileNameByRegular("E:\Documents\Desktop\", "2345pinyin_v[0-9]\.[0-9]\.[0-9]*\.exe")

Dialog("2345王牌输入法v*.*安装").Click 341,293 @@ hightlight id_;_395960_;_script infofile_;_ZIP::ssf49.xml_;_
Dialog("2345王牌输入法v*.*安装").Click 275,306 @@ hightlight id_;_395960_;_script infofile_;_ZIP::ssf52.xml_;_
wait 30
Dialog("2345王牌输入法v*.*安装").Click 329,302 @@ hightlight id_;_397222_;_script infofile_;_ZIP::ssf68.xml_;_

Dialog("2345王牌输入法设置向导").WinComboBox("ComboBox").Select "8" @@ hightlight id_;_592570_;_script infofile_;_ZIP::ssf53.xml_;_
Dialog("2345王牌输入法设置向导").WinComboBox("ComboBox_2").Select "大" @@ hightlight id_;_789932_;_script infofile_;_ZIP::ssf54.xml_;_
Dialog("2345王牌输入法设置向导").Click 526,410 @@ hightlight id_;_985772_;_script infofile_;_ZIP::ssf55.xml_;_
Dialog("2345王牌输入法设置向导").Click 526,410 @@ hightlight id_;_985772_;_script infofile_;_ZIP::ssf56.xml_;_

Dialog("2345王牌输入法设置向导").Click 61,355 @@ hightlight id_;_985772_;_script infofile_;_ZIP::ssf57.xml_;_
Dialog("2345王牌输入法设置向导").Click 540,409 @@ hightlight id_;_985772_;_script infofile_;_ZIP::ssf58.xml_;_








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


Systemutil.run GetFileNameByRegular("C:\Program Files (x86)\2345Soft\2345Pinyin\", "[0-9]\.[0-9]\.[0-9]\.[0-9]*") & "\Uninstall.exe"



 @@ hightlight id_;_1880172800_;_script infofile_;_ZIP::ssf62.xml_;_
Dialog("2345王牌输入法v*.*卸载").Click 30,356 @@ hightlight id_;_396190_;_script infofile_;_ZIP::ssf63.xml_;_
Dialog("2345王牌输入法v*.*卸载").Click 69,363 @@ hightlight id_;_396190_;_script infofile_;_ZIP::ssf64.xml_;_
Dialog("2345王牌输入法v*.*卸载").Click 420,349 @@ hightlight id_;_396190_;_script infofile_;_ZIP::ssf65.xml_;_
Dialog("2345王牌输入法v*.*卸载").Dialog("2345好压卸载程序").Click 235,169 @@ hightlight id_;_658498_;_script infofile_;_ZIP::ssf66.xml_;_
wait 15
Dialog("2345王牌输入法v*.*卸载").Click 521,350 @@ hightlight id_;_396190_;_script infofile_;_ZIP::ssf67.xml_;_




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


























