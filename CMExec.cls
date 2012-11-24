VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

' dwLanguageId
Private Const LANG_USER_DEFAULT = &H400&

Private WithEvents m_console As clsConsole
Attribute m_console.VB_VarHelpID = -1

Private mStdOutput As String, mComplete As Boolean
Public Function GetAPIErrStr(dwErrCode As Long) As String
  Dim sErrDesc As String * 256   ' max string resource len
  Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS Or _
                                 FORMAT_MESSAGE_MAX_WIDTH_MASK, _
                                 ByVal 0&, dwErrCode, LANG_USER_DEFAULT, _
                                 ByVal sErrDesc, 256, 0)
  GetAPIErrStr = GetStrFromBufferA(sErrDesc)
End Function
Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function
Public Function Exec(ByVal Strcmd As String) As String

    If m_console Is Nothing Then Set m_console = New clsConsole
    mStdOutput = ""
    mComplete = False
    If m_console.RunCommand(0, Strcmd) = 0 Then
        Do Until mComplete
            DoEvents
        Loop
    Else
        Exec = "Execution of " & Strcmd & " failed;" & GetAPIErrStr(Err.LastDllError)
    
    End If
    Exec = mStdOutput



End Function

Private Sub m_console_Complete(stdOut As String, stdErr As String)
    mComplete = True
    mStdOutput = stdOut
End Sub
