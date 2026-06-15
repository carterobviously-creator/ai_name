Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Folder where THIS script lives (vbs/)
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)

' Go up one level to the main folder
mainFolder = fso.GetParentFolderName(scriptFolder)

' Full path to pkg.html
pkgPath = fso.BuildPath(mainFolder, "pkg.html")

If fso.FileExists(pkgPath) Then
    shell.Run """" & pkgPath & """"
Else
    MsgBox "pkg.html not found in: " & mainFolder, 48, "Not Found"
End If
