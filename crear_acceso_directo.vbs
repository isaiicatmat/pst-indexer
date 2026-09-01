Set oWS = WScript.CreateObject("WScript.Shell")
strDesktop = oWS.SpecialFolders("Desktop")
strAppPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

Set oLink = oWS.CreateShortcut(strDesktop & "\Buscador de Correos.lnk")
oLink.TargetPath = strAppPath & "\ejecutar.bat"
oLink.WorkingDirectory = strAppPath
oLink.Description = "Buscador de Correos Outlook"
oLink.IconLocation = "shell32.dll,149"
oLink.Save

MsgBox "✓ Acceso directo creado en el escritorio" & vbCrLf & vbCrLf & "Puedes doble-click para ejecutar la aplicación", vbInformation, "Éxito"
