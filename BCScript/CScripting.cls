VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScripting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mconsole As New CConsole
'Private handleparser As Object
Public Sub Exec(ByVal ScriptFile As String)
    ModBCScript.ExecuteScript ScriptFile, mMainParser



End Sub



Public Sub WriteLn(ByVal Strwrite)
    'Con.WriteLine Strwrite, True
    Dim writeme As String

    If IsObject(Strwrite) Or IsArray(Strwrite) Then
   
    
    
        writeme = mMainParser.ResultToString(Strwrite)
    Else
        writeme = Strwrite
    End If
    Debug.Print "writeLn called"
   
    Con.WriteLine writeme
End Sub
Public Function ReadLn(Optional ByVal Prompt As String = "", Optional ByVal NumChars As Long = 0, Optional ByVal pwCharacter As String = "") As String
    Dim readIt As String
    If Con.IsCompiled Then
    If pwCharacter = "" Then
    readIt = Con.ReadLine(Prompt, NumChars)
    Else
        readIt = Con.ReadPassword(Prompt, pwCharacter)
    
    End If
    Else
        readIt = InputBox(Prompt, "ReadLn...")
    
    End If
    ReadLn = readIt

End Function
Public Sub FlashWindow(Optional ByVal Count As Long = 2, Optional ByVal Delay As Long = 250)
    Con.FlashWindow Count, Delay
End Sub
Public Sub Maximize()
    Con.FullScreen = True
End Sub
Public Sub Restore()
    Con.FullScreen = False
End Sub
Public Sub Pause(ByVal Prompt As String)
    Con.PressAnyKey Prompt
End Sub
Public Property Let CursorVisible(ByVal Vdata As Boolean)
    Con.CursorVisible = Vdata
End Property
Public Property Get CursorVisible() As Boolean
    CursorVisible = Con.CursorVisible
End Property
Public Property Let Title(ByVal Vdata As String)
Con.Title = Vdata
End Property
Public Property Get Title() As String
    Title = Con.Title
End Property

Public Property Let CurrentX(ByVal Vdata As Long)
    Con.CurrentX = Vdata
End Property
Public Property Get CurrentX() As Long
    CurrentX = Con.CurrentX
End Property
Public Property Let CurrentY(ByVal Vdata As Long)
    Con.CurrentY = Vdata
End Property
Public Property Get CurrentY() As Long
    CurrentY = Con.CurrentY
End Property
Public Property Let CursorHeight(ByVal Vdata As Long)
    Con.CursorHeight = Vdata
End Property
Public Property Get CursorHeight() As Long
    CursorHeight = Con.CursorHeight
End Property

Public Property Get ForeColor() As Integer
    ForeColor = Con.ForeColor
End Property
Public Property Let ForeColor(ByVal Vdata As Integer)
    Con.ForeColor = Vdata
End Property

Public Property Get BackColor() As Integer
    BackColor = Con.BackColor
End Property
Public Property Let BackColor(ByVal Vdata As Integer)
    Con.BackColor = Vdata
    
End Property




'Public Property Get Console() As CConsole
'    Set Console = mconsole
'End Property


'Private Sub handleparser_Error(ParserError As BASeParserXP.CParserError, RecoveryConst As BASeParserXP.ParserErrorRecoveryConstants)
''
'
'WriteLine "Parser Error in Script."
'WriteLine "position:" & ParserError.Position & _
'vbCrLf & """" & ParserError.Description & """"
'RecoveryConst = PERR_RESUMENEXT
'End Sub
'
'Private Sub handleparser_ExecuteComplete(valret As Variant)
''
'End Sub
'
'Private Sub handleparser_ExpressionChange()
''
'End Sub
'
'Private Sub handleparser_ParseComplete()
''
'End Sub
'
'Private Sub handleparser_PluginAfterLoad(PluginObject As BASeParserXP.IEvalEvents, cancel As Boolean)
''
''If VerboseFlag Then WriteLn "plugin loaded, """ & progID & """."
'End Sub
'
'Private Sub handleparser_PluginBeforeLoad(progID As String, cancel As Boolean)
''
'If VerboseFlag Then WriteLn "loading plugin """ & progID & """..."
'End Sub
Private Sub Class_Initialize()
    Set mconsole = New CConsole
End Sub
