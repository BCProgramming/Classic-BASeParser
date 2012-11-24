VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CValStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'holds Values during a parse.
Public Event PushItem(ByRef ValPushed As Variant)
Private WithEvents mStack As DataStack
Attribute mStack.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mStack = New DataStack
End Sub
Public Sub Push(ByVal cFormitem As Variant)
If IsObject(cFormitem) Then
    If TypeOf cFormitem Is CVariable Then
        Debug.Assert False
    End If
    
End If
    mStack.Push cFormitem
End Sub
Public Function Peek() As Variant
    If IsObject(mStack.Peek) Then
        Set Peek = mStack.Peek
    Else
        Peek = mStack.Peek
    End If
End Function
Public Function Pop() As Variant
     If IsObject(mStack.Peek) Then
        Set Pop = mStack.Pop
    Else
        Pop = mStack.Pop
    End If
End Function

Private Sub mStack_ItemPush(ItemVal As Variant)
    RaiseEvent PushItem(ItemVal)
End Sub
