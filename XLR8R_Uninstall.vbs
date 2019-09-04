Dim objFileSys
Dim objWshShell
Dim objExcel
Dim objAddin
Dim installPath

Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")

On Error Resume Next

objExcel.Workbooks.Add
For i = 1 To objExcel.AddIns.Count
    Set objAddin = objExcel.AddIns.Item(i)
    If objAddin.Name = "XLR8R.xlam" Then
        objAddin.Installed = False
    End If
Next
objExcel.Quit
Set objAddin = Nothing

installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
objFileSys.DeleteFile installPath & "XLR8R.xlam", True
objFileSys.DeleteFile installPath & "XLR8R.ini", True

Set objExcel = Nothing
Set objFileSys = Nothing
Set objWshShell = Nothing

If Err.Number = 0 Then
    MsgBox "Uninstallation complete.", vbInformation, "Uninstall"
Else
    MsgBox "Uninstallation failed.", vbExclamation, "Uninstall"
End If
