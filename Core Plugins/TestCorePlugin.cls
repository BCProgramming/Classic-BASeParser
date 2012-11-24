VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlugEnvString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Core plugin to handle Environment variables.
'I mean Core as in both- it is both part of the Core Parser, as well as a ICorePlugin Implementor.
'Makes ya think... OK, it doesn't.




Implements ICorePlugin
Implements IPluginUIData
Private mIDrecognize As Long

Private Function ICorePlugin_CanHandleItem(withparser As CParser, Itemcheck As CFormItem) As Boolean
    If Itemcheck.ItemType = mIDrecognize Then
        ICorePlugin_CanHandleItem = True
    End If
End Function

Private Property Get ICorePlugin_Description() As String
    ICorePlugin_Description = "Core environment Variable Parser Plugin"
End Property

Private Function ICorePlugin_GetPluginUIData() As IPluginUIData
    Set ICorePlugin_GetPluginUIData = Me
End Function

Private Sub ICorePlugin_HandleItem(withparser As CParser, ItemHandle As CFormItem, retval As Variant)
    If ItemHandle.ItemType = mIDrecognize Then
    
        retval = Environ$(ItemHandle.op)
    
    
    End If
End Sub



Private Function ICorePlugin_Init(withparser As CParser) As IPluginInitCode
'
CDebug.Post "Core Environment Variable Hook Plugin Loaded successfully."
ICorePlugin_Init = Code_Success

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
    Dim EnvString As String
    If Mid$(StrParse, Position, 1) = "%" Then
        'assume a environment variable.
        EnvString = Mid$(StrParse, Position, InStr(Position + 1, StrParse, "%", vbTextCompare) - Position + 1)
        currItem.ItemType = mIDrecognize
        currItem.Extra = EnvString
        Set currItem.CoreOwner = Me
        currItem.op = Mid$(EnvString, 2, Len(EnvString) - 2)
        Position = Position + Len(EnvString)
        
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
IPluginUIData_Description = "Environment Variable String Parser Plugin"
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As stdole.Picture
'
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String
'
End Function
