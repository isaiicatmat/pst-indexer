' Script VBA para exportar correos de Outlook a archivos .msg
' Ejecute desde: cmd > cscript.exe exportar_correos.vbs

On Error Resume Next

Set objOutlook = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo abrir Outlook"
    WScript.Echo "Asegúrate de que Outlook está instalado y ejecutándose"
    WScript.Quit 1
End If

Set objNamespace = objOutlook.GetNamespace("MAPI")

WScript.Echo "Selecciona la carpeta de Outlook que quieres exportar..."
Set objFolder = objOutlook.Session.PickFolder()

If objFolder Is Nothing Then
    WScript.Echo "Se canceló la operación"
    WScript.Quit
End If

WScript.Echo "Carpeta seleccionada: " & objFolder.Name

' Crear carpeta destino
strDestPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Correos_Exportados"
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FolderExists(strDestPath) Then
    fso.CreateFolder(strDestPath)
    WScript.Echo "Carpeta creada: " & strDestPath
End If

' Exportar correos recursivamente
intCount = 0
intSubfolders = 0
ExportarCarpeta objFolder, strDestPath, intCount, intSubfolders

WScript.Echo ""
WScript.Echo "======================================"
WScript.Echo "Exportación completada:"
WScript.Echo "Correos exportados: " & intCount
WScript.Echo "Carpetas procesadas: " & intSubfolders
WScript.Echo "Guardados en: " & strDestPath
WScript.Echo "======================================"

If intCount = 0 Then
    WScript.Echo "ADVERTENCIA: No se encontraron correos"
    WScript.Echo "Verifica que seleccionaste una carpeta con correos"
End If

WScript.Echo "Presiona Enter para cerrar..."
WScript.StdIn.ReadLine()

' Subrutina para exportar de forma recursiva
Sub ExportarCarpeta(objFolder, strPath, ByRef intCount, ByRef intSubfolders)
    Dim objItem, objSubfolder, strFileName

    WScript.Echo "Procesando: " & objFolder.Name & " (" & objFolder.Items.Count & " items)"

    ' Exportar correos de esta carpeta
    For Each objItem In objFolder.Items
        If objItem.Class = 43 Then ' 43 = Mail message
            On Error Resume Next
            strFileName = strPath & "\" & intCount & "_" & Left(Replace(objItem.Subject, "/", "_"), 100) & ".msg"
            objItem.SaveAs strFileName, 3 ' 3 = Outlook message format
            If Err.Number = 0 Then
                intCount = intCount + 1
                If intCount Mod 100 = 0 Then
                    WScript.Echo "  " & intCount & " correos exportados..."
                End If
            Else
                WScript.Echo "  ERROR exportando: " & objItem.Subject
            End If
            Err.Clear
            On Error Resume Next
        End If
    Next

    ' Procesar subcarpetas
    For Each objSubfolder In objFolder.Folders
        intSubfolders = intSubfolders + 1
        ExportarCarpeta objSubfolder, strPath, intCount, intSubfolders
    Next
End Sub
