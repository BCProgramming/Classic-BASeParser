VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUnitValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CUnitValue: represents a Number that is expressed in some unit.


Private mvarValue As Variant    'the number/value.
Private mvarUnit As String



Public Property Let Value(ByVal Vdata As Variant)
    mvarValue = Vdata
End Property
Public Property Get Value() As Variant
    Value = mvarValue
End Property
Public Property Let Unit(ByVal Vdata As String)
    mvarUnit = Vdata
End Property
Public Property Get Unit() As String
    Unit = mvarUnit
End Property
