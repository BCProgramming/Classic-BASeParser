VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Functions Class- holds, CFunction objects.
'designed to be a aggregate of CParser
Public Event FunctionAdded(FuncAdded As CFunction)
Public Event FunctionRemoved(FuncRemoved As CFunction)

Private mCol As Collection
Private mParser As CParser
Private loopobject As CFunction
Friend Property Get Parser() As CParser
    Set Parser = mParser
End Property
Friend Property Set Parser(Vdata As CParser)
    Set mParser = Vdata
End Property
Public Function FindByName(ByVal funcName As String) As CFunction
    'looks for the given name in the collection. obviously a simple Mcol.Item, but
    'nobody needs to know that. Well, maybe you, dear reader.
    Set FindByName = mCol.Item(funcName)
End Function

Public Function Add(ByVal FuncStr As String, FuncKey As String) As CFunction

    Dim AddFunc As CFunction
    Set AddFunc = New CFunction
    AddFunc.Name = FuncKey
    AddFunc.InitParser mParser
    On Error GoTo ParseError
    AddFunc.Expression = FuncStr
    mCol.Add AddFunc
    RaiseEvent FunctionAdded(AddFunc)
    Set Add = AddFunc
    Exit Function
ParseError:
    Err.Raise 5, "CFunctions::Add", "Refusing to add Function, """ & FuncStr & """, due to parse error in expression."
End Function
Public Function AddDirect(AddFunc As CFunction) As CFunction
    Set AddDirect = AddFunc
    mCol.Add AddFunc, AddFunc.Name
    AddFunc.InitParser mParser
End Function
Public Property Get Count() As Long
    Count = mCol.Count
End Property
Private Function GetIndex(ForVar As Variant) As Long
Dim i As Long, ForVarCast As CFunction
    If IsObject(ForVar) Then
        If TypeOf ForVar Is CFunction Then
        'search through the collection
        Set ForVarCast = ForVar
        For i = 1 To mCol.Count
            Set loopobject = mCol.Item(i)
            If loopobject Is ForVarCast Then
                GetIndex = i
                Exit Function
            End If
        Next
        
        End If
    ElseIf VBA.IsNumeric(ForVar) Then
        'ordinal.
        GetIndex = ForVar
    ElseIf VarType(ForVar) = vbString Then
        'a string.
        'For Each loopObject In mCol
        For i = 1 To mCol.Count
            Set loopobject = mCol.Item(i)
            If loopobject.Name = ForVar Then
                GetIndex = i
                Exit Function
            End If
        Next
        'if we get here, synthetically raise an error.
        GetIndex = -1  'YES! FAIL!
    End If
    

End Function
Public Function Item(ByVal Vntindexkey As Variant) As CFunction
    'returns a CFunction object given either the object itself
    '(which makes this function kind of moot) or
    'a Ordinal, or Key. The key will search through the Name properties.
    If IsObject(Vntindexkey) Then
        If TypeOf Vntindexkey Is CFunction Then
            Set Item = Vntindexkey
        End If
    Else
        Set Item = mCol.Item(GetIndex(Vntindexkey))
    End If
    
End Function
Public Function Remove(ByVal Vntindexkey As Variant) As CFunction
    Dim remindex As Long, removed As CFunction
    remindex = GetIndex(Vntindexkey)
    Set removed = mCol.Item(remindex)
    
    Set Remove = removed
    mCol.Remove remindex
    RaiseEvent FunctionRemoved(removed)
End Function
Public Sub Clear()
    Set mCol = New Collection
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
