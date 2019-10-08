Dim objFileSys
Dim objWshShell
Dim objExcel
Dim objAddin
Dim installPath

Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")

On Error Resume Next

installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
objFileSys.CopyFile "XLR8R.xlam", installPath & "XLR8R.xlam", True
If Dir(installPath & "XLR8R.ini") = "" Then
    objFileSys.CopyFile "XLR8R.ini", installPath & "XLR8R.ini", True
End If

objExcel.Workbooks.Add
Set objAddin = objExcel.AddIns.Add(installPath & "XLR8R.xlam", True)
objAddin.Installed = True
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing
Set objFileSys = Nothing
Set objWshShell = Nothing

If Err.Number = 0 Then
    MsgBox "Installation complete.", vbInformation, "Install"
Else
    MsgBox "Installation failed.", vbExclamation, "Install"
End If
