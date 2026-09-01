' Script VBA para exportar correos de Outlook a archivos .msg
' Guarde esto como exportar_correos.vbs en el escritorio
' Luego ejecute: cscript.exe exportar_correos.vbs

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")

' Pedir carpeta al usuario
Set objFolder = objOutlook.Session.PickFolder()

If objFolder Is Nothing Then
    WScript.Echo "Se canceló la operación"
    WScript.Quit
End If

' Crear carpeta destino
strDestPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Correos_Exportados"
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FolderExists(strDestPath) Then
    fso.CreateFolder(strDestPath)
End If

' Exportar todos los correos
intCount = 0
For Each objItem In objFolder.Items
    If objItem.Class = 43 Then ' 43 = Mail message
        strFileName = strDestPath & "\" & Format(objItem.ReceivedTime, "yyyy-mm-dd_hhmm_") & intCount & ".msg"
        objItem.SaveAs strFileName, 3 ' 3 = Outlook message format (.msg)
        intCount = intCount + 1
        WScript.Echo "Exportado: " & objItem.Subject
    End If
Next

WScript.Echo "Exportación completada: " & intCount & " correos en " & strDestPath
