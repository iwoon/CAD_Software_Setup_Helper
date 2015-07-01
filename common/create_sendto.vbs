strShortcutFile =Wscript.Arguments.Item(0) & ".lnk"
set objShell = WScript.CreateObject("WScript.Shell" )
set objLink = objShell.CreateShortcut(objShell.SpecialFolders("SendTo") & "\" & strShortcutFile) 
target=Wscript.Arguments.Item(1)
desc=Wscript.Arguments.Item(2)
icon_loc=Wscript.Arguments.Item(3)
objLink.TargetPath = target
objLink.Description = desc
objLink.IconLocation = icon_loc
objLink.WorkingDirectory = Wscript.Arguments.Item(1)
objLink.Save