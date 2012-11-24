VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsDOS - used to run console programs and return results
'Code by:
'   William Moeur
'   http://moeur.net
'
'   Joacim Andersson, Brixoft Software
'   http://www.brixoft.net

Option Explicit

Public Event Complete(stdOut As String, stdErr As String)
Private hSTDOUT As Long, hSTDERR As Long
Private Const BUFSIZE = 1024 ' pipe buffer size

'Run the console program specified
'returns zero on success, otherwise returns error code
Public Function RunCommand(hwnd As Long, strCommand As String) As Long
    Dim hPipeWrite1 As Long, hPipeWrite2 As Long
    Dim hCurProcess As Long
    Dim sa As CmdOutput.SECURITY_ATTRIBUTES
    Dim si As CmdOutput.STARTUPINFO
    Dim pi As CmdOutput.PROCESS_INFORMATION
    ReDim baOutput(BUFSIZE - 1) As Byte

    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With

'create stdout pipe
    If CreatePipe(hSTDOUT, hPipeWrite1, sa, BUFSIZE) = 0 Then
        RunCommand = Err.LastDllError
        Exit Function
    End If

'create stderr pipe
    If CreatePipe(hSTDERR, hPipeWrite2, sa, BUFSIZE) = 0 Then
        RunCommand = Err.LastDllError
        Call CmdOutput.CloseHandle(hPipeWrite1)
        Exit Function
    End If
    
    hCurProcess = CmdOutput.GetCurrentProcess()

    ' Replace our inheritable read handles with an non-inheritable.
    'Not that it seems to be necessary in this case, but the docs
    'say we should.
    Call CmdOutput.DuplicateHandle(hCurProcess, hSTDOUT, hCurProcess, hSTDOUT, 0&, _
                         0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)

    Call CmdOutput.DuplicateHandle(hCurProcess, hSTDERR, hCurProcess, hSTDERR, 0&, _
                         0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
    
    'startup info
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
         ' hide the window since all output is redirected anyway
         .wShowWindow = SW_HIDE
        .hStdOutput = hPipeWrite1
        .hStdError = hPipeWrite2
    End With

    If CmdOutput.CreateProcess(vbNullString, strCommand, ByVal 0&, ByVal 0&, 1, 0&, _
        ByVal 0&, vbNullString, si, pi) = 0 Then
        RunCommand = Err.LastDllError
        Call CmdOutput.CloseHandle(hPipeWrite1)
        Call CmdOutput.CloseHandle(hPipeWrite2)
        Exit Function
    End If

        ' Close thread handle - we don't need it
        Call CmdOutput.CloseHandle(pi.hThread)
        
        ' Also close our handle(s) to the write end of the pipe. This is _
        ' important, since ReadFile will *not* return until all write _
        ' handles are closed or the buffer is full.
        Call CmdOutput.CloseHandle(hPipeWrite1)
        Call CmdOutput.CloseHandle(hPipeWrite2)

        RunCommand = 0
        'start our timer to poll for stdout and stderr messages
        CmdOutput.StartTimer hwnd, ObjPtr(Me), 100

End Function

'timer event, polls for stdout and stderr messages
'once any message arrives, our process is finished so
'shut down the timer and return the results in an event
Friend Sub CLSTimerProc(ByVal hwnd As Long, ByVal TimerID As Long, _
    ByVal dwTime As Long)
    Dim baOutput(BUFSIZE - 1) As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    Dim TotalBytesAvail As Long
    Dim BytesLeftThisMessage As Long
    Dim KillFlag As Boolean
    Dim strSTDOUT As String
    Dim strSTDERR As String
    
    Debug.Print "CLSTimerProc"
    'see if stdout has anything for us
    Call CmdOutput.PeekNamedPipe(hSTDOUT, baOutput(0), BUFSIZE, lBytesRead, TotalBytesAvail, _
        BytesLeftThisMessage)
    If lBytesRead > 0 Then
        strSTDOUT = ReadPipe(hSTDOUT)
        KillFlag = True
    End If
    
    'see if stderr has anything for us
    Call CmdOutput.PeekNamedPipe(hSTDERR, baOutput(0), BUFSIZE, lBytesRead, TotalBytesAvail, _
        BytesLeftThisMessage)
    If lBytesRead > 0 Then
        strSTDERR = ReadPipe(hSTDERR)
        KillFlag = True
        'maybe stdout has something for us now if it didn't before
        If strSTDOUT = "" Then
            Call CmdOutput.PeekNamedPipe(hSTDOUT, baOutput(0), BUFSIZE, lBytesRead, _
                TotalBytesAvail, BytesLeftThisMessage)
            If lBytesRead > 0 Then strSTDOUT = ReadPipe(hSTDOUT)
        End If
    End If
    
    If KillFlag Then
    'were done here
        CmdOutput.KillTimer hwnd, TimerID
        Call CmdOutput.CloseHandle(hSTDOUT)
        Call CmdOutput.CloseHandle(hSTDERR)
        RaiseEvent Complete(strSTDOUT, strSTDERR)
    End If
    
End Sub

Private Function ReadPipe(hPipe As Long) As String
    Dim baOutput(BUFSIZE - 1) As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    Dim CommandOutput As String
    
    CommandOutput = ""
    Do While ReadFile(hPipe, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) <> 0
        ' convert from "DOS" to "Windows" characters
        sNewOutput = String$(lBytesRead, 0)
        Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
        CommandOutput = CommandOutput & sNewOutput
    Loop
    ReadPipe = CommandOutput
End Function

