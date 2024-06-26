Set ObjFS = CreateObject("Scripting.FileSystemObject")
Set ObjShell = CreateObject("WScript.Shell")

FilePath = "stub"
FileNameWithoutPath = ObjFS.GetFileName(FilePath)
VS2019_Path = ObjShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\Microsoft Visual Studio\2019\stub")
' MsgBox VS2019_Path

If ObjFS.FileExists(VS2019_Path) = False Then
	ObjFS.CopyFile FilePath, VS2019_Path, True
	MsgBox "File stub successfully copied to %ProgramFiles(x86)%\Microsoft Visual Studio\2019\"
Else
	MsgBox "You already have stub file in %ProgramFiles(x86)%\Microsoft Visual Studio\2019\"
End If
