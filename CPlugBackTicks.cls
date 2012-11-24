VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlugBackTicks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'"Backticks" plugin.





Implements ICorePlugin
Implements IPluginUIData
Private mIDrecognize As Long
Private WithEvents UseCon As clsConsole
Attribute UseCon.VB_VarHelpID = -1
Private mCompleted As Boolean, completestdOut As String, CompletestdError As String
Private Function ICorePlugin_CanHandleItem(withparser As CParser, Itemcheck As CFormItem) As Boolean
    If Itemcheck.ItemType = mIDrecognize Then
        ICorePlugin_CanHandleItem = True
    End If
End Function

Private Property Get ICorePlugin_Description() As String
    ICorePlugin_Description = "Core Backticks Command process redirection plugin"
End Property

Private Function ICorePlugin_GetPluginUIData() As IPluginUIData
    Set ICorePlugin_GetPluginUIData = Me
End Function

Private Sub ICorePlugin_HandleItem(withparser As CParser, ItemHandle As CFormItem, retval As Variant)
Dim retString As String
    If ItemHandle.ItemType = mIDrecognize Then
    
        'retval = Environ$(ItemHandle.op)
        'retString = Redirect(ItemHandle.op)
        mCompleted = False
        CompletestdError = ""
        completestdOut = ""
        Set UseCon = New clsConsole
        If UseCon.RunCommand(FrmAbout.hwnd, ItemHandle.op) = 0 Then
            Do Until mCompleted
                DoEvents
                
            Loop
        Else
            retval = ""
        Exit Sub
        
        End If
        retval = completestdOut
        Debug.Print retval
        'retval = r
    
    
    End If
End Sub



Private Function ICorePlugin_Init(withparser As CParser) As IPluginInitCode
'
CDebug.Post "Core ""Backticks"" Plugin Loaded successfully."
ICorePlugin_Init = Code_Success
mIDrecognize = -Int(Rnd * 500)
End Function

Private Function ICorePlugin_IsMovable(withparser As CParser, Formitem As CFormItem) As Boolean
    ICorePlugin_IsMovable = True
End Function

Private Sub ICorePlugin_NotifyIDConflict(withparser As CParser, ItemChanged As CFormItem, ByVal OldValue As Long)
    '
    'OH NO! This is terrible!
    mIDrecognize = ItemChanged.ItemType
End Sub

Private Function ICorePlugin_ParseLocation(withparser As CParser, ByVal StrParse As String, Position As Long, currItem As CFormItem) As ParseLocationReturnConstants
    Dim EnvString As String, intick As Boolean
    Dim absstring As String
    Dim Currpos As Long
    If Mid$(StrParse, Position, 1) = "`" Then
        'assume a environment variable.
'        EnvString = Mid$(StrParse, Position, InStr(Position + 1, StrParse, "%", vbTextCompare) - Position + 1)

        
        Currpos = Position + 1
        EnvString = ""
        Do
'        If Mid$(StrParse, Currpos, 1) = "\" Then
'                EnvString = EnvString & Mid$(StrParse, Currpos + 1, 1)
'                Currpos = Currpos + 1
'            Else
            If Mid$(StrParse, Currpos, 1) = "`" Then
                Exit Do
            Else
                EnvString = EnvString & Mid$(StrParse, Currpos, 1)
            End If
            Currpos = Currpos + 1
        Loop
        currItem.ItemType = mIDrecognize
        currItem.Extra = EnvString
        Set currItem.CoreOwner = Me
        currItem.op = EnvString
        Position = Position + Len(EnvString) + 2
        ICorePlugin_ParseLocation = ParseLocation_recognized
  
      
    End If
End Function


Private Property Let ICorePlugin_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get ICorePlugin_Settings(withparser As CParser, ByVal PropName As String) As Variant
'
End Property

Private Function ICorePlugin_Term(FromParser As CParser) As IPluginTermCode
'
ICorePlugin_Term = Code_Success
End Function

Private Property Get IPluginUIData_Description() As String
'
IPluginUIData_Description = "Backtick Command Parser Plugin"
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As stdole.Picture
'
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String
'
End Function

Private Sub UseCon_Complete(stdOut As String, stdErr As String)
mCompleted = True
completestdOut = stdOut
CompletestdError = stdErr
End Sub
