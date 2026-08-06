' ============================================================
'   CRIA ATALHO DO SISTEMA RH NA AREA DE TRABALHO
'   Com icone vermelho da coroa!
' ============================================================

Option Explicit

Dim WshShell, fso, objShellLink
Dim desktopPath, scriptPath, targetPath, iconPath

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
desktopPath = WshShell.SpecialFolders("Desktop")
targetPath = scriptPath & "\Iniciar_SistemaRH_Silencioso.vbs"
iconPath = scriptPath & "\app_icon.ico"

' Verifica se o arquivo VBS existe
If Not fso.FileExists(targetPath) Then
    MsgBox "Arquivo nao encontrado:" & vbCrLf & targetPath & vbCrLf & vbCrLf & _
           "Certifique-se de que o Criar_Atalho_Desktop.vbs esta na mesma pasta do sistema.", _
           vbExclamation, "Erro"
    WScript.Quit 1
End If

' Verifica se o icone existe
If Not fso.FileExists(iconPath) Then
    iconPath = scriptPath & "\app_icon.png"
    If Not fso.FileExists(iconPath) Then
        iconPath = ""
    End If
End If

' Cria o atalho
Set objShellLink = WshShell.CreateShortcut(desktopPath & "\Sistema RH.lnk")
objShellLink.TargetPath = targetPath
objShellLink.WorkingDirectory = scriptPath
objShellLink.Description = "Sistema de RH"
objShellLink.IconLocation = iconPath & ",0"
objShellLink.WindowStyle = 7 ' Minimizado
objShellLink.Save

Set objShellLink = Nothing
Set WshShell = Nothing
Set fso = Nothing

MsgBox "Atalho criado com sucesso!" & vbCrLf & vbCrLf & _
       "Nome: Sistema RH" & vbCrLf & _
       "Local: Area de Trabalho" & vbCrLf & _
       "Icone: Vermelho (coroa)" & vbCrLf & vbCrLf & _
       "Agora e so clicar duas vezes no icone vermelho!", _
       vbInformation, "Sucesso"
