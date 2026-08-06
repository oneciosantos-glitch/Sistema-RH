' ============================================================
'   SISTEMA RH - INICIALIZACAO SILENCIOSA
'   Sem janela preta do CMD!
' ============================================================

Option Explicit

Dim WshShell, fso, scriptPath, pythonPath
Dim objWMIService, colProcesses, objProcess
Dim isRunning

Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
Set WshShell = CreateObject("WScript.Shell")

' Verifica se o Streamlit ja esta rodando
isRunning = False
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'streamlit.exe' OR Name = 'python.exe'")

For Each objProcess in colProcesses
    If InStr(objProcess.CommandLine, "app_rh.py") > 0 Then
        isRunning = True
        Exit For
    End If
Next

If isRunning Then
    ' So abre o navegador
    WshShell.Run "http://localhost:8501", 1, False
    WScript.Quit 0
End If

' Verifica se o Python esta instalado
On Error Resume Next
pythonPath = WshShell.RegRead("HKLM\SOFTWARE\Python\PythonCore\3.11\InstallPath\ExecutablePath")
If Err.Number <> 0 Then
    pythonPath = WshShell.RegRead("HKLM\SOFTWARE\Python\PythonCore\3.10\InstallPath\ExecutablePath")
End If
If Err.Number <> 0 Then
    pythonPath = WshShell.RegRead("HKLM\SOFTWARE\Python\PythonCore\3.9\InstallPath\ExecutablePath")
End If
If Err.Number <> 0 Then
    pythonPath = "python"
End If
On Error GoTo 0

' Executa o launcher.py SILENCIOSAMENTE (sem janela preta)
' 0 = oculto, False = nao espera terminar
WshShell.Run "cmd /c cd /d """ & scriptPath & """ && """ & pythonPath & """ launcher.py""", 0, False

' Pequena pausa para o Streamlit iniciar
WScript.Sleep 3000

' Abre o navegador
WshShell.Run "http://localhost:8501", 1, False

Set WshShell = Nothing
Set fso = Nothing
