VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IDynamicDispatch
'Represents an Enumeration- IE: set of constants.
'Very basic implementation...
'two basic properties- a Name, and a group of variables.
Private mNames() As String, mValues() As Variant
Private mvarName As String


Public Property Get Name() As String
    Name = mvarName
End Property

Public Sub InitEnum(ByVal Name As String, ConstNames() As String, ConstValues() As Variant)
    mvarName = Name
    mNames = ConstNames
    mValues = ConstValues
    



End Sub

Private Sub IDynamicDispatch_InvokeMember(withparser As CParser, ByVal Strmember As String, Params As Variant, ByVal PreferProperty As Boolean, returnValue As Variant)
'
'loop through our Names array. If any name equals the member, return it.

Dim I As Long
For I = LBound(mNames) To UBound(mNames)
    If StrComp(mNames(I), Strmember, vbTextCompare) = 0 Then
        If IsObject(mValues(I)) Then Set returnValue = mValues(I) Else returnValue = mValues(I)
        Exit Sub
    
    End If

Next I
End Sub
