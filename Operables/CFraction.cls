VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *****************************************************************
'  Copyright ©1994-2000, Zane Thomas, http://www.mabry.com
'  Modified by Karl E. Peterson, May 2000, for publication
'  in VBPJ and at http://www.mvps.org/vb/
' *****************************************************************
'  Authors grant royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' ******************************************************************
'  http://www.dictionary.com/cgi-bin/dict.pl?term=Fractional%20number
' ******************************************************************
'Modification by Michael Burgwin (BASeCamp corporation)
'for BASeParser XP: now implements BASeParser XP interfaces
'to support IOperable.
'Also changed the name to CFraction, since CRat was stupid. Sounds like it
'should be in some Animal Library.
#If BASEPARSER Then
    Implements IOperable
#End If



Option Explicit

Private m_Numer As Long
Private m_Denom As Long
Private m_AutoReduce As Boolean

' ************************************************
'  Initialize
' ************************************************
Private Sub Class_Initialize()
   m_Numer = 0
   m_Denom = 1
   m_AutoReduce = True
End Sub

' ************************************************
'  Public Properties
' ************************************************
Public Property Let AutoReduce(ByVal NewVal As Boolean)
   m_AutoReduce = NewVal
End Property

Public Property Get AutoReduce() As Boolean
   AutoReduce = m_AutoReduce
End Property

Public Property Let Denominator(ByVal NewVal As Long)
   If NewVal = 0 Then
      ' Denominator cannot be zero
      Err.Raise Number:=11, Source:="CFraction.Denominator", _
         Description:="Division by zero."
   Else
      m_Denom = NewVal
   End If
End Property

Public Property Get Denominator() As Long
   Denominator = m_Denom
End Property

Public Property Let Numerator(ByVal NewVal As Long)
   m_Numer = NewVal
End Property

Public Property Get Numerator() As Long
   Numerator = m_Numer
End Property

' ************************************************
'  Public Methods - Math
' ************************************************
Public Sub Reduce()
   Dim Denom As Long
   ' Calc GCD and divide both to reduce.
   Denom = Gcd(m_Numer, m_Denom)
   m_Numer = m_Numer / Denom
   m_Denom = m_Denom / Denom
End Sub

Public Function Add(aFraction As CFraction) As CFraction
   Dim Result As New CFraction
   '
   ' Add a Fraction, return result as a new Fraction.
   '
   ' n1   n2   n1d2 + n2d1
   ' -- + -- = -----------
   ' d1   d2       d1d2
   '
   Result.Numerator = m_Numer * aFraction.Denominator + aFraction.Numerator * m_Denom
   Result.Denominator = m_Denom * aFraction.Denominator
   If m_AutoReduce Then Result.Reduce
   Set Add = Result
End Function

Public Function Subtract(aFraction As CFraction) As CFraction
   Dim Result As New CFraction
   '
   ' Subtract a Fraction, return result as a new Fraction
   '
   ' n1   n2   n1d2 - n2d1
   ' -- - -- = -----------
   ' d1   d2       d1d2
   '
   Result.Numerator = m_Numer * aFraction.Denominator - aFraction.Numerator * m_Denom
   Result.Denominator = m_Denom * aFraction.Denominator
   If m_AutoReduce Then Result.Reduce
   Set Subtract = Result
End Function

Public Function Multiply(aFraction As CFraction) As CFraction
   Dim Result As New CFraction
   '
   ' Multiply a Fraction, return result as a new Fraction
   '
   ' n1   n2   n1d2
   ' -- * -- = ----
   ' d1   d2   d1d2
   '
   Result.Numerator = m_Numer * aFraction.Numerator
   Result.Denominator = m_Denom * aFraction.Denominator
   If m_AutoReduce Then Result.Reduce
   Set Multiply = Result
End Function
Public Sub Test()
Dim X As CFraction, Y As CFraction
Set X = New CFraction
Set Y = New CFraction
X.SetValues 1, 2
Y.SetValues 1, 2
'Debug.Print X.Multiply(Y).AsString
End Sub
Public Function Divide(aFraction As CFraction) As CFraction
   Dim Result As New CFraction
   '
   ' Divide by a Fraction, return result as a new Fraction
   '
   ' n1/d1   n1d2
   ' ----- = ----
   ' n2/d2   d1n2
   '
   Result.Numerator = m_Numer * aFraction.Denominator
   Result.Denominator = m_Denom * aFraction.Numerator
   If m_AutoReduce Then Result.Reduce
   Set Divide = Result
End Function

' ************************************************
'  Public Methods - Utility/Display
' ************************************************
Public Function AsFloat() As Double
   AsFloat = m_Numer / m_Denom
End Function

Public Function AsString(Optional ByVal ShowWhole As Boolean = False) As String
   If m_AutoReduce Then Me.Reduce
   
   If m_Numer Mod m_Denom = 0 Then
      ' Output should be like "2"
      AsString = m_Numer \ m_Denom
   
   ElseIf Abs(m_Numer) > Abs(m_Denom) And ShowWhole = True Then
      ' Output should be like "1-3/8"
      AsString = CStr(m_Numer \ m_Denom) & "-" & _
         CStr(Abs(m_Numer Mod m_Denom)) & "/" & _
         CStr(Abs(m_Denom))
   
   Else
      ' Output should be like "11/8"
      AsString = m_Numer & "/" & m_Denom
   End If
End Function

Public Function MakeFraction(ByVal Numer As Long, ByVal Denom As Long) As CFraction
   Dim Result As New CFraction
   ' Provide a Fraction for on-the-fly calcs
   Result.SetValues Numer, Denom
   Set MakeFraction = Result
End Function

Public Sub SetValue(ByVal NewVal As Double, Optional ByVal HighDenom As Long = 64)
   Dim Numer As Long
   Dim Denom As Long
   Dim Whole As Long
   ' Create a Fraction from a single floating-point value,
   ' and the largest value acceptable for denom.
   Whole = Abs(Fix(NewVal))
   Denom = Abs(HighDenom)
   Numer = Denom * (Abs(NewVal) - Whole)
   If Numer Then
      m_Numer = ((Whole * Denom) + Numer) * Sgn(NewVal)
   Else
      m_Numer = (Whole * Denom) * Sgn(NewVal)
   End If
   Me.Denominator = Denom * Sgn(HighDenom)
End Sub

Public Sub SetValues(ByVal Numer As Long, ByVal Denom As Long)
   ' Create a Fraction from passed numer and denom.
   m_Numer = Numer
   Me.Denominator = Denom
End Sub

' ************************************************
'  Private Methods
' ************************************************
Private Function Gcd(ByVal A As Long, ByVal B As Long) As Long
   ' Find greatest common denominator.
   If (B = 0) Then
      Gcd = A
   Else
      Gcd = Gcd(B, A Mod B)
   End If
End Function


Private Function IOperable_PerformFunc(withparser As CParser, ByVal FuncName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants
IOperable_PerformFunc = Perform_Unknown
End Function

Private Sub IOperable_PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, retval As Variant)
    '
    'Err.Raise 5, "CFraction.IOperable::PerformOp", "PerformOp Not implemented. for fractions yet."
    
    'Operations on fractions:
    'this is actually a pretty simple thing to do! just get the Variant Decimal result of division and
    'then perform the given operation on the two values!
    Dim Values(1 To 2) As Variant
    Dim retdecimal As Variant, retfraction As CFraction
    Values(1) = CDec(Me.Numerator) / CDec(Me.Denominator)
    
    If IsObject(OtherOperand) Then
            If TypeOf OtherOperand Is CFraction Then
                Values(2) = CDec(OtherOperand.Numerator) / CDec(OtherOperand.Denominator)
            Else
                'we'll be calling performoperation with our decimal fraction representation and
                'this other item, so if that object implements IOperable and knows how to perform the operation
                'with a number, we'll be just fine.
                Set Values(2) = OtherOperand
            End If
            
    ElseIf IsNumeric(OtherOperand) Then
        'well, by golly! if it is a number, we are just dandy!
        Values(2) = OtherOperand
    Else
    
        On Error Resume Next
        Values(2) = Val(OtherOperand)
        If Err <> 0 Then
            Values(0) = 0
        End If
    End If
    Call withparser.PerformOperation(Operation, Values(1), Values(2), retdecimal)
    Set retfraction = New CFraction
    retfraction.SetValue retdecimal
    Set retval = retfraction


End Sub

Private Function IOperable_toString(withparser As CParser) As String
'
IOperable_toString = Me.AsString(False)
End Function
