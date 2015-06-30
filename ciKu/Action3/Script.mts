
Print("BaseDir: " & parameter("BaseDir"))
Print("RegExp: " & parameter("RegExp"))

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = parameter("RegExp")

Dim oFolders
Dim oRoot, oFolder
Dim oFiles, oFile
Dim colMatches

Set FSO = CreateObject("Scripting.FileSystemObject")
Set oRoot = FSO.GetFolder(parameter("BaseDir"))
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

parameter("FileName") = parameter("BaseDir")