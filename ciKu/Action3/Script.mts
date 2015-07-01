'Log Parameter
Print("INFO " & Environment("ActionName") & " - Parameter BaseDir: " & parameter("BaseDir"))
Print("INFO " & Environment("ActionName") & " - Parameter RegExp: " & parameter("RegExp"))

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = parameter("RegExp")

Dim oFolders
Dim oRoot, oFolder
Dim oFiles, oFile
Dim colMatches

Set FSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set oRoot = FSO.GetFolder(parameter("BaseDir"))
If Err.number <> 0 Then
	Print("ERROR " & Environment("ActionName") & " - Not Exist " & parameter("BaseDir"))
	parameter("FileName") = ""
	ExitActionIteration()
End If
Set oFolders = oRoot.SubFolders

For Each oFolder In oFolders
	Set colMatches = objRegEx.Execute(oFolder.Path)
	If colMatches.Count > 0 Then
		parameter("FileName") = oFolder.Path
		ExitActionIteration()
	End If
Next

Set oFiles = oRoot.Files
For Each oFile In oFiles
	Set colMatches = objRegEx.Execute(oFile.Name)
	If colMatches.Count > 0 Then
		parameter("FileName") = oFile.Path
		ExitActionIteration()
	End If
Next

parameter("FileName") = ""