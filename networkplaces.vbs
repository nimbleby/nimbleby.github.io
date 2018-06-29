Set WSHShell = CreateObject("WScript.Shell")

Const NETHOOD = &H13&

Set objWSHShell = CreateObject("Wscript.Shell")

Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.Namespace(NETHOOD)
Set objFolderItem = objFolder.Self
strNetHood = objFolderItem.Path

strPath = objFolderItem.Path & "\*.*"
On Error Resume Next
objFSO.DeleteFile strPath, true

strShortcutName = "ACCOUNTING"
strShortcutPath = "\\servername\share"

Set objShortcut = objWSHShell.CreateShortcut _
(strNetHood & "\" & strShortcutName & ".lnk")
objShortcut.TargetPath = strShortcutPath
objShortcut.Save

strShortcutName = "COMMERCIAL DEPT"
strShortcutPath = "\\servername\sharename"

Set objShortcut = objWSHShell.CreateShortcut _
(strNetHood & "\" & strShortcutName & ".lnk")
objShortcut.TargetPath = strShortcutPath
objShortcut.Save

strShortcutName = "HUMAN RESOURCES"
strShortcutPath = "\\servername\sharename"

Set objShortcut = objWSHShell.CreateShortcut _
(strNetHood & "\" & strShortcutName & ".lnk")
objShortcut.TargetPath = strShortcutPath
objShortcut.Save

strShortcutName = "PROCUREMENT"
strShortcutPath = "\\servername\sharename"

WSCript.Quit

'As you can see you will need to replace the names with friendly names of yoru own your users will 'understand and recongnise. You mus then replace \\servername\sharename with the name of the file 'server and the share on it.

'Save the script as networkplaces.vbs
