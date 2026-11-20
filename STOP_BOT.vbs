Option Explicit
Dim fso, shell, baseDir, stopFlag, pidFile, wmi, procs, p, stoppedCount, cmdLow, nameLow, waitLoop, stillRunning

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
stopFlag = fso.BuildPath(baseDir, "stop_bot.flag")
pidFile = fso.BuildPath(baseDir, "bot_running.pid")

Function IsBotRunning(wmiObj)
    Dim ps, proc, nLow, cLow
    IsBotRunning = False
    Set ps = wmiObj.ExecQuery("SELECT Name, CommandLine FROM Win32_Process")
    For Each proc In ps
        nLow = LCase(proc.Name & "")
        cLow = LCase(proc.CommandLine & "")
        If nLow = "x_autotweetbot.exe" Or _
           ((nLow = "python.exe" Or nLow = "pythonw.exe") And InStr(1, cLow, "auto_tweet_playwright.py") > 0) Then
            IsBotRunning = True
            Exit Function
        End If
    Next
End Function

Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

If Not IsBotRunning(wmi) Then
    If fso.FileExists(pidFile) Then
        On Error Resume Next
        fso.DeleteFile pidFile, True
        On Error GoTo 0
    End If
    MsgBox "No running X_AutoTweetBot process was found.", vbInformation, "STOP X_AutoTweetBot"
    WScript.Quit 0
End If

' First request graceful stop so Playwright closes its own browser and Excel cleanly
On Error Resume Next
Dim ts
Set ts = fso.CreateTextFile(stopFlag, True)
ts.WriteLine "stop"
ts.Close
On Error GoTo 0

stoppedCount = 0
stillRunning = True
For waitLoop = 1 To 18
    WScript.Sleep 300
    If Not IsBotRunning(wmi) Then
        stillRunning = False
        stoppedCount = 1
        Exit For
    End If
Next

' Force terminate only if still unresponsive; never touches manually opened chrome.exe
If stillRunning Then
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process")
    For Each p In procs
        nameLow = LCase(p.Name & "")
        cmdLow = LCase(p.CommandLine & "")
        If nameLow = "x_autotweetbot.exe" Or _
           ((nameLow = "python.exe" Or nameLow = "pythonw.exe") And InStr(1, cmdLow, "auto_tweet_playwright.py") > 0) Then
            On Error Resume Next
            p.Terminate()
            If Err.Number = 0 Then
                stoppedCount = stoppedCount + 1
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
End If

Dim activeProfFile
activeProfFile = fso.BuildPath(baseDir, "bot_active_profile.txt")
On Error Resume Next
If fso.FileExists(stopFlag) Then fso.DeleteFile stopFlag, True
If fso.FileExists(pidFile) Then fso.DeleteFile pidFile, True
If fso.FileExists(activeProfFile) Then fso.DeleteFile activeProfFile, True
On Error GoTo 0

If stoppedCount > 0 Then
    MsgBox "Safely stopped X_AutoTweetBot." & vbCrLf & "Manual Chrome windows and GUI remain open.", vbInformation, "STOP X_AutoTweetBot"
Else
    MsgBox "No running X_AutoTweetBot process was found.", vbInformation, "STOP X_AutoTweetBot"
End If
