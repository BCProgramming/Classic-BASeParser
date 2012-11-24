VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COpStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'holds operators during a parse.
Private mStack As DataStack

Private Sub Class_Initialize()
    Set mStack = New DataStack
End Sub
Public Sub Push(ByVal Operator As String)
    mStack.Push Operator
End Sub
Public Function Peek() As String
    Peek = mStack.Peek
End Function
Public Function Pop() As String
    Pop = mStack.Pop
End Function
Public Function Count() As Long
    Count = mStack.Count
End Function
