VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumberUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarBase As String  'For example, Metre.
                            'mostly for display
                            
Private mvarUnit As String
Private mvarValue As Variant
'CNumberUnit: Represents a value given a unit suffix
Public Property Let Unit(ByVal Vdata As String)
    mvarUnit = Vdata
End Property
Public Property Get Unit() As String
    Unit = mvarUnit
End Property
Public Property Let Base(ByVal Vdata As String)
    mvarBase = Vdata
End Property
Public Property Get Base() As String
    Base = mvarBase
End Property
Public Property Get Value() As Variant
    Value = mvarValue
End Property
Public Property Let Value(ByVal Vdata As Variant)
    mvarValue = Vdata
End Property
Public Function ToString() As String

    ToString = Str$(mvarValue) & mvarUnit


End Function
'requires Module mConversions


'Metre (Base Unit,No multiplier)
Private Sub InitData()

    mConversions.InitModule




'10 -1   deci-   d
'10 -2   centi-  c
'10 -3   milli-  m
'10 -6   micro-
'10 -9   nano-   n
'10 -12  pico-   p
'10 -15  femto-  f
'10 -18  atto-   a
'10 -21  zepto-  z
'10 -24  yocto-  y
End Sub
