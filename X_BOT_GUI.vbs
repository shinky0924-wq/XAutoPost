Option Explicit
Dim fso, shell, baseDir, pyCmd, guiPy, cmdLine

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = baseDir

Function FindPythonGui()
    Dim localApp, progFiles, candidates, i, pathEnv, dirs, d, candidate, pyw
    localApp = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    progFiles = shell.ExpandEnvironmentStrings("%ProgramFiles%")

    candidates = Array( _
        "C:\Users\PC_User\AppData\Local\Programs\Python\Python312\pythonw.exe", _
        "C:\Users\PC_User\AppData\Local\Programs\Python\Python312\python.exe", _
        localApp & "\Programs\Python\Python312\pythonw.exe", _
        localApp & "\Programs\Python\Python312\python.exe", _
        localApp & "\Programs\Python\Python311\pythonw.exe", _
        localApp & "\Programs\Python\Python311\python.exe", _
        localApp & "\Programs\Python\Python313\pythonw.exe", _
        localApp & "\Programs\Python\Python313\python.exe", _
        localApp & "\Programs\Python\Python310\pythonw.exe", _
        localApp & "\Programs\Python\Python310\python.exe", _
        progFiles & "\Python312\pythonw.exe", _
        progFiles & "\Python312\python.exe" _
    )

    For i = 0 To UBound(candidates)
        If fso.FileExists(candidates(i)) Then
            FindPythonGui = candidates(i)
            Exit Function
        End If
    Next

    pathEnv = shell.ExpandEnvironmentStrings("%PATH%")
    dirs = Split(pathEnv, ";")
    For Each d In dirs
        d = Trim(d)
        If d <> "" And InStr(1, LCase(d), "windowsapps") = 0 Then
            pyw = fso.BuildPath(d, "pythonw.exe")
            If fso.FileExists(pyw) Then
                FindPythonGui = pyw
                Exit Function
            End If
            candidate = fso.BuildPath(d, "python.exe")
            If fso.FileExists(candidate) Then
                FindPythonGui = candidate
                Exit Function
            End If
        End If
    Next

    FindPythonGui = ""
End Function

guiPy = fso.BuildPath(baseDir, "x_bot_gui.py")
pyCmd = FindPythonGui()

If Not fso.FileExists(guiPy) Then
    MsgBox "x_bot_gui.py was not found in:" & vbCrLf & baseDir, vbExclamation, "X_AutoTweetBot GUI"
ElseIf pyCmd = "" Then
    MsgBox "Python was not found. Please install Python 3.10-3.12.", vbExclamation, "X_AutoTweetBot GUI"
Else
    cmdLine = Chr(34) & pyCmd & Chr(34) & " " & Chr(34) & guiPy & Chr(34)
    shell.Run cmdLine, 0, False
End If
