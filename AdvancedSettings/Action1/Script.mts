' Get install file path and install it
Dim VersionPath
RunAction "GetDirNameByRegular [UtilityActions]", oneIteration, "C:\Program Files (x86)\2345Soft\2345Pinyin\", "[0-9]\.[0-9]\.[0-9]\.[0-9]*", VersionPath @@ hightlight id_;_199146_;_script infofile_;_ZIP::ssf19.xml_;_
If Len(VersionPath) = 0 Then
	Print("ERROR " & Environment("ActionName") & " - Get install file path failed")
	ExitActionIteration()
End If

Systemutil.run VersionPath & "\2345PinyinConfig.exe"


Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

Dialog("2345王牌输入法设置").Click 66,214 @@ hightlight id_;_657146_;_script infofile_;_ZIP::ssf1.xml_;_

Dialog("2345王牌输入法设置").Click 205,137 @@ hightlight id_;_657146_;_script infofile_;_ZIP::ssf2.xml_;_
Dialog("2345王牌输入法设置").Click 361,139 @@ hightlight id_;_657146_;_script infofile_;_ZIP::ssf3.xml_;_
wait 1
'TODO:
Dialog("2345王牌输入法设置").Dialog("模糊音设置").Click 56,212 @@ hightlight id_;_984614_;_script infofile_;_ZIP::ssf4.xml_;_
Dialog("2345王牌输入法设置").Dialog("模糊音设置").Click 336,217 @@ hightlight id_;_984614_;_script infofile_;_ZIP::ssf5.xml_;_
Dialog("2345王牌输入法设置").Click 390,268 @@ hightlight id_;_657146_;_script infofile_;_ZIP::ssf6.xml_;_


'清空用户词库
Dialog("2345王牌输入法设置").Click 78,144 @@ hightlight id_;_526250_;_script infofile_;_ZIP::ssf23.xml_;_
Dialog("2345王牌输入法设置").Click 478,146 @@ hightlight id_;_526250_;_script infofile_;_ZIP::ssf24.xml_;_
Dialog("2345王牌输入法设置").Dialog("模糊音设置").Click 243,164 @@ hightlight id_;_526926_;_script infofile_;_ZIP::ssf25.xml_;_
Dialog("2345王牌输入法设置").Dialog("模糊音设置").Click 194,163 @@ hightlight id_;_592462_;_script infofile_;_ZIP::ssf26.xml_;_
Dialog("2345王牌输入法设置").Click 420,458


'添加自定义短语 @@ hightlight id_;_2098834_;_script infofile_;_ZIP::ssf3.xml_;_
'Dialog("自定义短语设置").Click 74,262 @@ hightlight id_;_461848_;_script infofile_;_ZIP::ssf4.xml_;_
Dialog("自定义短语设置").Dialog("添加自定义短语").WinObject("缩写").Click 104,13
WshShell.SendKeys "jita"
'WshShell.SendKeys ("{Enter}")


'Dialog("自定义短语设置").Dialog("添加自定义短语").Click 181,146 @@ hightlight id_;_268760_;_script infofile_;_ZIP::ssf20.xml_;_
Dialog("自定义短语设置").Dialog("添加自定义短语").WinComboBox("候选位置").Select "3" @@ hightlight id_;_268846_;_script infofile_;_ZIP::ssf21.xml_;_

Dialog("自定义短语设置").Dialog("添加自定义短语").DblClick 61,26 @@ hightlight id_;_268760_;_script infofile_;_ZIP::ssf19.xml_;_


' Dialog("自定义短语设置").Dialog("添加自定义短语").WinObject("自定义短语").Type "" @@ hightlight id_;_399288_;_script infofile_;_ZIP::ssf6.xml_;_
Dialog("自定义短语设置").Dialog("添加自定义短语").WinObject("自定义短语").Click 82,17 @@ hightlight id_;_4199240_;_script infofile_;_ZIP::ssf7.xml_;_


WshShell.SendKeys "jiajia"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")


'Dialog("自定义短语设置").Dialog("添加自定义短语").WinObject("ATL:006572E8_2").Type "" @@ hightlight id_;_4199240_;_script infofile_;_ZIP::ssf8.xml_;_
Dialog("自定义短语设置").Dialog("添加自定义短语").Click 292,326 @@ hightlight id_;_200564_;_script infofile_;_ZIP::ssf9.xml_;_
Dialog("自定义短语设置").Click 395,265 @@ hightlight id_;_461848_;_script infofile_;_ZIP::ssf10.xml_;_



'Dialog("2345王牌输入法设置").Click 405,442 @@ hightlight id_;_4068404_;_script infofile_;_ZIP::ssf22.xml_;_



'
'Dim WshShell
'
'Set WshShell = CreateObject("WScript.Shell")

'已通过测试维护运行更新



'
'Dim WshShell
'
'Set WshShell = CreateObject("WScript.Shell")


'验证输入内容是不是生效

TXTFileName = "gaoji"
WshShell.Run "notepad"
wait 1
WshShell.AppActivate "词库 - 记事本"

'验证模糊音
WshShell.SendKeys "zidao"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")


WshShell.SendKeys "cifan"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")

WshShell.SendKeys "side"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "niaojie"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "hengzi"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "lengran"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "qianhan"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "fengzi"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "qintian"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "qianqiang"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "quangjia"

WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")

'验证纠错

WshShell.SendKeys "lign"

WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "homg"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "qiou"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "tuei"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "kuen"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
WshShell.SendKeys "jita"
WshShell.SendKeys ("3")
WshShell.SendKeys ("{Enter}")


'保存


WshShell.SendKeys "^s"
wait 1
WshShell.SendKeys TXTFileName
WshShell.SendKeys (" ")
WshShell.SendKeys "%s"
WshShell.SendKeys (" ")
WshShell.SendKeys ("{Enter}")
wait 1
WshShell.SendKeys ("{LEFT}")
WshShell.SendKeys ("{Enter}")
wait 1
WshShell.SendKeys ("%{F4}")
'检查点
FileContent("高级.txt").Check CheckPoint("高级.txt")


' 恢复默认状态
Systemutil.run VersionPath & "\2345PinyinConfig.exe" @@ hightlight id_;_68006_;_script infofile_;_ZIP::ssf20.xml_;_
Dialog("2345王牌输入法设置").Click 78,210 @@ hightlight id_;_135926_;_script infofile_;_ZIP::ssf21.xml_;_
Dialog("2345王牌输入法设置").Click 76,452 @@ hightlight id_;_135926_;_script infofile_;_ZIP::ssf22.xml_;_
Dialog("2345王牌输入法设置").Click 338,268 @@ hightlight id_;_203068_;_script infofile_;_ZIP::ssf14.xml_;_
Dialog("自定义短语设置").WinObject("ATL:00657498").Click 135,32 @@ hightlight id_;_334054_;_script infofile_;_ZIP::ssf15.xml_;_
Dialog("自定义短语设置").Click 278,265 @@ hightlight id_;_399602_;_script infofile_;_ZIP::ssf16.xml_;_
Dialog("自定义短语设置").Dialog("添加自定义短语").Click 235,160 @@ hightlight id_;_334048_;_script infofile_;_ZIP::ssf17.xml_;_
Dialog("自定义短语设置").Click 445,255 @@ hightlight id_;_399602_;_script infofile_;_ZIP::ssf18.xml_;_
Dialog("2345王牌输入法设置").Click 396,452 @@ hightlight id_;_135926_;_script infofile_;_ZIP::ssf23.xml_;_