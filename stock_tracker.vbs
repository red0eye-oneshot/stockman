Option Explicit

' ============================================================
' stock_tracker.vbs
' - Server running  : refresh browser (Ctrl+Shift+R)
' - Server stopped  : kill stale Python -> restart silently
'                     -> wait for ready -> open browser
' ============================================================

Dim oShell, oFSO, scriptDir
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
scriptDir  = oFSO.GetParentFolderName(WScript.ScriptFullName)

' --- Log file ---
Dim fLog
On Error Resume Next
Set fLog = oFSO.CreateTextFile(scriptDir & "\vbs_log.txt", True)
If Err.Number <> 0 Then
    Err.Clear
    Set fLog = oFSO.CreateTextFile( _
        oShell.ExpandEnvironmentStrings("%TEMP%") & "\stock_tracker.log", True)
End If
On Error GoTo 0

Sub Log(msg)
    On Error Resume Next
    If Not (fLog Is Nothing) Then fLog.WriteLine Now() & " " & msg
    On Error GoTo 0
End Sub

Log "=== START ==="
Log "scriptDir=" & scriptDir

Dim srvPath : srvPath = scriptDir & "\stock_server.py"
If Not oFSO.FileExists(srvPath) Then
    Log "FAIL: stock_server.py missing"
    ClosLog
    MsgBox "stock_server.py not found:" & vbCrLf & scriptDir, vbCritical, "Stock Tracker"
    WScript.Quit 1
End If

' ── Case 1: server already running -> just refresh browser ──
If ServerAlive() Then
    Log "Server alive - refreshing browser"
    ClosLog
    RefreshOrOpen
    WScript.Quit 0
End If

' ── Case 2: server not running -> kill stale, restart silently ──
Log "Server not alive - killing stale Python"
oShell.Run "taskkill /F /IM python.exe",  0, True
oShell.Run "taskkill /F /IM pythonw.exe", 0, True
oShell.Run "taskkill /F /IM pyw.exe",     0, True
oShell.Run "taskkill /F /IM py.exe",      0, True
WScript.Sleep 1200

StartServer

' ── Wait up to 35 s for server to come up ──
Dim i
For i = 1 To 35
    WScript.Sleep 1000
    If ServerAlive() Then
        Log "Server up after " & i & "s"
        ClosLog
        OpenBrowser
        WScript.Quit 0
    End If
    Log "waiting " & i & "s"
Next

Log "TIMEOUT"
ClosLog
MsgBox "Server did not start within 35s." & vbCrLf & _
       "Run diagnose.bat for details.", vbExclamation, "Stock Tracker"
WScript.Quit 1


' ============================================================
' Sub / Function definitions
' ============================================================

' -- Start server without a console window --
Sub StartServer()
    Dim started : started = False

    ' 1) Try pyw (Python launcher, no window)
    On Error Resume Next
    oShell.Run "pyw """ & srvPath & """", 0, False
    If Err.Number = 0 Then started = True
    Err.Clear
    On Error GoTo 0
    If started Then Log "Started via pyw" : Exit Sub

    ' 2) Try pythonw.exe from common install paths
    Dim pyBase : pyBase = ""
    Dim pyDirs(5)
    pyDirs(0) = "C:\Python313"
    pyDirs(1) = "C:\Python314"
    pyDirs(2) = "C:\Python312"
    pyDirs(3) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313"
    pyDirs(4) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python314"
    pyDirs(5) = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312"
    Dim d
    For Each d In pyDirs
        If oFSO.FileExists(d & "\pythonw.exe") Then
            pyBase = d : Exit For
        End If
    Next
    If pyBase <> "" Then
        Log "Started via pythonw: " & pyBase
        oShell.Run """" & pyBase & "\pythonw.exe"" """ & srvPath & """", 0, False
        Exit Sub
    End If

    ' 3) Fallback: run_server.bat (may show a brief window)
    Dim batPath : batPath = scriptDir & "\run_server.bat"
    If oFSO.FileExists(batPath) Then
        Log "Fallback: run_server.bat"
        oShell.Run "cmd /c """ & batPath & """", 0, False
    Else
        Log "Fallback: py (last resort)"
        oShell.Run "py """ & srvPath & """", 0, False
    End If
End Sub

' -- Open browser with fresh URL (cache-busting) --
Sub RefreshOrOpen()
    OpenBrowser
End Sub

' -- Open browser with cache-busting timestamp URL --
Sub OpenBrowser()
    Dim ts : ts = Year(Now) & _
                  Right("0" & Month(Now),  2) & _
                  Right("0" & Day(Now),    2) & _
                  Right("0" & Hour(Now),   2) & _
                  Right("0" & Minute(Now), 2) & _
                  Right("0" & Second(Now), 2)
    Dim url : url = "http://127.0.0.1:5555/?_t=" & ts
    On Error Resume Next
    oShell.Run "cmd /c start " & url, 0, False
    If Err.Number <> 0 Then
        Err.Clear
        oShell.Run "rundll32 url.dll,FileProtocolHandler " & url, 1, False
    End If
    On Error GoTo 0
End Sub

' -- Check if server is responding --
Function ServerAlive()
    ServerAlive = False
    On Error Resume Next
    Dim exitCode
    exitCode = oShell.Run( _
        "cmd /c curl -s --max-time 2 http://127.0.0.1:5555/api/ping >nul 2>nul", _
        0, True)
    If exitCode = 0 Then ServerAlive = True
    On Error GoTo 0
End Function

Sub ClosLog()
    On Error Resume Next
    If Not (fLog Is Nothing) Then fLog.Close
    On Error GoTo 0
End Sub
