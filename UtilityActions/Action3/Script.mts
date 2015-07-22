' Log Parameter
Print("INFO " & Environment("ActionName") & " - Parameter BaseDir: " & parameter("BaseDir"))
Print("INFO " & Environment("ActionName") & " - Parameter RegExp: " & parameter("RegExp"))

Dim oFolders
Dim oRoot, oFolder
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
Set oFolders = oRoot.SubFolders

' Find dir that match the regExp
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = parameter("RegExp")
For Each oFolder In oFolders
	Set colMatches = objRegEx.Execute(oFolder.Path)
	If colMatches.Count > 0 Then
		parameter("FileName") = oFolder.Path
	End If
Next

