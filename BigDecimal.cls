VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'requires: cStringBuilder.cls
Public Enum SignConstants
    Sign_negative = -1
    sign_zero = 0
    Sign_positive = 1
End Enum

Private mvarValues() As Byte
Private mvarSign As SignConstants
'each value represents two hex digits.

'all internal arithmetic will be done via the hex of the bytes (IE base 255 rather then 10).

'also, the first array items are the least significant, that is the first item will be the "ones" place
Public Property Get Sign() As SignConstants
    Sign = mvarSign
End Property
Public Property Let Sign(ByVal vdata As SignConstants)
    mvarSign = vdata
End Property
Public Function SetBytes(Setit() As Byte)
    mvarValues = Setit
End Function
Public Function GetBytes() As Byte
    GetBytes = mvarValues
End Function
Public Function Add(OtherInt As BigInteger) As BigInteger
    Dim retme As BigInteger
    Dim opA() As Byte, OpB() As Byte, result() As Byte
    Dim Iterate As Long, Carry As Integer
    Dim Loopto As Long, IntSum As Long, newsign As SignConstants
    Set retme = New BigInteger
    opA = mvarValues
    OpB = OtherInt.GetBytes
    If UBound(opA) > UBound(OpB) Then Loopto = UBound(opA) Else Loopto = UBound(OpB)
    'redim ...
    
    ReDim Preserve opA(1 To Loopto)
    ReDim Preserve OpB(1 To Loopto)
    ReDim Preserve result(1 To Loopto)
    For Iterate = 1 To Loopto
        IntSum = (opA(Iterate) * Me.Sign) + (OpB(Iterate) * OtherInt.Sign)
        newsign = Sgn(IntSum)
        IntSum = Abs(IntSum)
        
        IntSum = IntSum + Carry
        If IntSum > 255 Then
            Carry = IntSum Mod 255
            IntSum = 255
        Else
            Carry = 0
            
        End If
    
    
    
    
        result(Iterate) = IntSum
    
    Next Iterate
    
    
    
    
    


End Function
Public Function toString() As String
    'returns a string representation....
    Dim I As Long
    Dim retstr As String
    For I = LBound(mvarValues) To UBound(mvarValues)
        retstr = retstr & Trim$(Hex$(mvarValues(I)))
    Next I
    toString = retstr

End Function
Private Sub Class_Initialize()
    
    
End Sub
