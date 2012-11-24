VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'outputhooks: holds a collection of IParserOutput implementors.

Private mCol As Collection


Private Function FindIndex(OfObject As Variant) As Long
    
    Dim I As Long
    For I = 1 To mCol.Count
        If mCol(I) Is OfObject Then
            FindIndex = I
            Exit Function
        End If
    Next I
    FindIndex = -1


End Function
Public Sub Add(ObjAdd As IParserOutput)
    mCol.Add ObjAdd
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub Remove(Vntindexkey As Variant)
    If FindIndex(Vntindexkey) <= 0 Then
        Exit Sub
    Else
        mCol.Remove Vntindexkey
    End If
End Sub
Public Sub Message(withparser As CParser, ByVal Strmessage As String, Optional Verbosity As Integer = 0)
    Dim Loopoutput As IParserOutput
    For Each Loopoutput In mCol
        Loopoutput.Message withparser, Strmessage, Verbosity
    
    Next Loopoutput




End Sub
