VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICorePlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICorePlugin
Private mCol As Collection
Private mParser As CParser
Private loopobject As ICorePlugin
'delegator class that delegates calls on the ICorePlugin interface of this class to
'the appropriate class that can actually handle the operation.
Friend Property Set Parser(Vdata As CParser)
    Set mParser = Vdata
End Property
Friend Property Get Parser() As CParser
    Set Parser = mParser
End Property
Public Function Add(PluginObject As ICorePlugin)
'right now, assumed to be the actual object.

mCol.Add PluginObject
If PluginObject.Init(mParser) Then
    'load failure

End If
End Function
'Public Function Remove(PluginObject
Public Property Get Col() As Collection
    Set Col = mCol
End Property
Public Function GetFirstHandler(withparser As CParser, ForItem As CFormItem) As ICorePlugin
    For Each loopobject In mCol
        If loopobject.CanHandleItem(withparser, ForItem) Then
            Set GetFirstHandler = loopobject
            Exit Function
        End If
    Next
    

End Function
Public Sub Clear()
    For Each loopobject In mCol
        loopobject.Term mParser
    Next
    Set mCol = New Collection
End Sub
Private Sub Class_Initialize()
    Set mCol = New Collection
    
End Sub

Private Function ICorePlugin_CanHandleItem(withparser As CParser, Itemcheck As CFormItem) As Boolean
    For Each loopobject In mCol
        If loopobject.CanHandleItem(withparser, Itemcheck) Then
            ICorePlugin_CanHandleItem = True
            Exit Function
        End If
    Next
End Function

Private Function ICorePlugin_GetPluginUIData() As IPluginUIData
'
End Function

Private Sub ICorePlugin_HandleItem(withparser As CParser, ItemHandle As CFormItem, retval As Variant)
    For Each loopobject In mCol
        If loopobject.CanHandleItem(withparser, ItemHandle) Then
            Call loopobject.HandleItem(withparser, ItemHandle, retval)
            Exit Sub
        End If
    Next loopobject
End Sub

Private Function ICorePlugin_Init(withparser As CParser) As IPluginInitCode
'

End Function



Private Sub ICorePlugin_NotifyIDConflict(withparser As CParser, ItemChanged As CFormItem, ByVal OldValue As Long)
'STUB!
End Sub

Private Function ICorePlugin_ParseLocation(withparser As CParser, ByVal StrParse As String, Position As Long, currItem As CFormItem) As ParseLocationReturnConstants
Dim savedpos As Long
    savedpos = Position
    For Each loopobject In mCol
        savedpos = Position
        If loopobject.ParseLocation(withparser, StrParse, savedpos, currItem) = ParseLocation_recognized Then
            Position = savedpos
            ICorePlugin_ParseLocation = ParseLocation_recognized
            Exit Function
        Else
            
        End If
    
        
    Next
End Function

Private Property Let ICorePlugin_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get ICorePlugin_Settings(withparser As CParser, ByVal PropName As String) As Variant
'
End Property

Private Function ICorePlugin_Term(FromParser As CParser) As IPluginTermCode
'
End Function
