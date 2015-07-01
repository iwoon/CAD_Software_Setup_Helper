strShortcutFile =Wscript.Arguments.Item(0) & ".lnk"
set objShell = WScript.CreateObject("WScript.Shell" )
set objLink = objShell.CreateShortcut(objShell.SpecialFolders("AllUsersDesktop") & "\" & strShortcutFile) 
url=Wscript.Arguments.Item(1)
desc=Wscript.Arguments.Item(2)
icon_loc=Wscript.Arguments.Item(3)
com_art=Wscript.Arguments.Item(4)
IE_ver=Wscript.Arguments.Item(5)
If com_art = 32 Then
	objLink.TargetPath = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\Internet Explorer\iexplore.exe"
	objLink.WorkingDirectory = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\Internet Explorer\"
else If com_art = 64 Then
	objLink.TargetPath = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") & "\Internet Explorer\iexplore.exe"
	objLink.WorkingDirectory = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") & "\Internet Explorer\"
End If
End If
If IE_ver < 8 Then
	objLink.arguments =" -extoff " & url
Else
	objLink.arguments =" -private " & url
End If
objLink.Description = desc
objLink.IconLocation = icon_loc
objLink.Save
