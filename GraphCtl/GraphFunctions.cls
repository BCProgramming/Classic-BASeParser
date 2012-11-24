VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'GraphFunctions
'Collection of CXGraphFunction objects.


Private mCol As Collection
Private mvarControlowner As FGraph
Friend Property Set Owner(Vdata As FGraph)
    Set mvarControlowner = Vdata
End Property
Friend Property Get Owner() As FGraph
    Set Owner = mvarControlowner
End Property
Public Sub DrawAll(ByVal OnDC As Long)
    Dim LoopIt As CGraphFunction
    For Each LoopIt In mCol
        LoopIt.DrawFunction OnDC
    Next


End Sub
Public Function Add(ByVal ObjAdd As Variant)
    Dim newfunc As CGraphFunction
    If IsObject(ObjAdd) Then
        If TypeOf ObjAdd Is CGraphFunction Then
        Set newfunc = ObjAdd
            mCol.Add ObjAdd
            Set Add = ObjAdd
        ElseIf TypeOf ObjAdd Is BASeParserXP.Cfunction Then
            'allow them to add a Cfunction directly too.
            Set newfunc = New CGraphFunction
            Set newfunc.Cfunction = ObjAdd
            mCol.Add newfunc, newfunc.Name
            Set Add = newfunc
        End If
    Else
        Set newfunc = New CGraphFunction
        newfunc.Expression = ObjAdd
        mCol.Add newfunc, "F" & Trim$(Str$(mCol.Count))
        Set Add = newfunc
    End If
    Set newfunc.Parent = Me
End Function
Public Function Remove(ByVal VntIndexkey As Variant) As CGraphFunction
    Set Remove = Me.Item(VntIndexkey)
    mCol.Remove VntIndexkey
End Function


Public Function Item(ByVal VntIndexkey As Variant)
Attribute Item.VB_UserMemId = 0
    Set Item = mCol.Item(VntIndexkey)
End Function
Public Function newEnum() As IUnknown
Attribute newEnum.VB_UserMemId = -4
    Set newEnum = mCol.[_NewEnum]
End Function
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Dim ReadFunc As CGraphFunction
    Dim I As Long
    Dim totalCount As Long
    
    
    With PropBag
        totalCount = .ReadProperty("FunctionCount", 0)
        For I = 1 To totalCount
            Set ReadFunc = .ReadProperty("Function" & Trim$(Str$(I)), Nothing)
            If Not ReadFunc Is Nothing Then
                mCol.Add ReadFunc
            End If
        Next I
        
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
Dim LoopFunc As CGraphFunction
Dim I As Long
With PropBag
    .WriteProperty "FunctionCount", mCol.Count, 0
    I = 1
    For Each LoopFunc In Me
        .WriteProperty "Function" & Trim$(Str$(I)), LoopFunc, Nothing
        I = I + 1
    Next
End With
End Sub
