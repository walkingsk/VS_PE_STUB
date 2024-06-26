Set ObjFS = CreateObject("Scripting.FileSystemObject")
Set ObjShell = CreateObject("WScript.Shell")

FilePath = "stub"
FileNameWithoutPath = ObjFS.GetFileName(FilePath)
VS2017_Path = ObjShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\Microsoft Visual Studio\2017\stub")
' MsgBox VS2017_Path

If ObjFS.FileExists(VS2017_Path) = False Then
	ObjFS.CopyFile FilePath, VS2017_Path, True
	MsgBox "File stub successfully copied to %ProgramFiles(x86)%\Microsoft Visual Studio\2017\"
Else
	MsgBox "You already have stub file in %ProgramFiles(x86)%\Microsoft Visual Studio\2017\"
End If
