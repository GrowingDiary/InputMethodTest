' Log Parameter
Print("INFO " & Environment("ActionName") & " - Parameter BaseDir: " & parameter("BaseDir"))
Print("INFO " & Environment("ActionName") & " - Parameter RegExp: " & parameter("RegExp"))

Dim oRoot
Dim oFiles, oFile
Dim colMatches

parameter("FileName") = ""

' If the BaseDir is exist
Set FSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set oRoot = FSO.GetFolder(parameter("BaseDir"))
If Err.number <> 0 Then
	Print("ERROR " & Environment("ActionName") & " - Not Exist " & parameter("BaseDir"))
	ExitActionIteration()
End If

' Find dir that match the regExp
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = parameter("RegExp")
Set oFiles = oRoot.Files
For Each oFile In oFiles
	Set colMatches = objRegEx.Execute(oFile.Name)
	If colMatches.Count > 0 Then
		parameter("FileName") = oFile.Path
	End If
Next

