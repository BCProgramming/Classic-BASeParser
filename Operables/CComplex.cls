VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComplex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CComplex
'Complex Number Support Class. Designed for Use by my BASeParser Library.
'may work in other contexts.


'I must apologize, this class is tightly coupled with BASeParser XP. If you want
'to use it in another project by itself, you'll need to replace
'all the PerformOperation method calls and the WithParser parameters.

'You caught me! I converted some Perl Core code. So sue me. OK, please don't.



Implements IOperable        'Implement BASeParser Operable Interface.

Public Enum ValueType
    value_real
    Value_Complex
End Enum
Private mRealPart As Variant        'why variant? Like I said, it was designed with BASeParser in mind.
                        'you might want to change it :). Mostly there for CDec().
Private mImagPart As Variant
Private mRho As Variant
Private mTheta As Variant
Private mFormatExpression As String
Private mFormatFunction As CFunction

Private mlastparser As CParser
'stole some stuff from Perl modules.
'What? doesn't everybody browse the source?

'return $self->{'cartesian'} = [$r * CORE::cos($t), $r * CORE::sin($t)];
Public Sub UpdateCartesian()
    mRealPart = mRho * Cos(mTheta)
    mImagPart = mRho * Sin(mTheta)
End Sub
Private Function Atan2(ByVal X, ByVal Y)
    Atan2 = Atn(Y / X)
End Function
Private Sub UpdatePolar()
    'Assign values to Polar form variables- A and b.
    'return $self->{'polar'} = [0, 0] if $x == 0 && $y == 0;
    'return $self->{'polar'} = [CORE::sqrt($x*$x + $y*$y), CORE::atan2($y, $x)];
    Dim X, Y
    X = mRealPart
    Y = mImagPart
    If X = 0 Or Y = 0 Then
        mRho = 0
        mTheta = 0
    Else
        mRho = Sqr(X ^ 2 + Y ^ 2)
        mTheta = Atan2(Y, X)
    
    End If
    
End Sub

Public Sub SetValue(ByVal Realpart As Double, ByVal ImagPart As Double)
    mRealPart = Realpart
    mImagPart = ImagPart
End Sub
Public Property Let Value(ByVal ValType As ValueType, ByVal Vdata As Variant)
    If ValType = value_real Then mRealPart = Vdata Else mImagPart = Vdata
End Property
Public Property Get Value(ByVal ValType As ValueType) As Variant
    If ValType = value_real Then Value = mRealPart Else Value = mImagPart
End Property
Public Function IsReal() As Boolean
Attribute IsReal.VB_Description = "Returns wether this complex number has no imaginary component."
    IsReal = IsEmpty(mImagPart) Or mImagPart = 0
End Function
Public Property Let Realpart(ByVal Vdata As Variant)
Attribute Realpart.VB_Description = "Sets/Returns the Real Component of this Complex Number."
    mRealPart = Vdata
    UpdatePolar
End Property
Public Property Get Realpart() As Variant
    Realpart = mRealPart
End Property
Public Property Let ImagPart(ByVal Vdata As Variant)
Attribute ImagPart.VB_Description = "Sets/Returns the Imaginary Part of this Complex Number."
    mImagPart = Vdata
    UpdatePolar
End Property
Public Property Get ImagPart() As Variant
    ImagPart = mImagPart
End Property
Public Property Get StringForm() As String
'a + bi

Dim strbuffer As String
Dim strImag As String




    If mImagPart = 0 Then
        strImag = ""
    ElseIf mImagPart = 1 Then
        'neither. it is simply I.
        strImag = "i"
    Else
        strImag = Trim$(Abs(mImagPart)) & "i"
    End If
strImag = Trim$(strImag)
If mRealPart = 0 Then
    'no real portion.
    strbuffer = IIf(Sgn(mImagPart) = -1, "-", "+") & strImag
Else
    If Sgn(mImagPart) = -1 Then
        strbuffer = Trim$(CStr(mRealPart)) & " - " & strImag
    ElseIf strImag = "" Then
        strbuffer = Trim$(CStr(mRealPart))
    Else
        'positive imaginary.
        strbuffer = Trim$(CStr(mRealPart)) & " + " & strImag
        
    End If


End If
If left$(strbuffer, 1) = "+" Then
    'remove a +, since this means it is + then the imagpart.
    strbuffer = Mid$(strbuffer, 2)
End If


StringForm = strbuffer



End Property
Public Property Let StringForm(ByVal Vdata As String)
    'can I ever escape from String parsing :-P
    'well, if the last character is an i, then this has at least an imagpart.
    If StrComp(right$(Vdata, 1), "i") = 0 Then
    'sigh.
        Realpart = Val(Vdata)
        'MODIFY:
        'I realized this wouldn't work for negative reals. So I made sure to use InstrRev
        'and check that it is not the first character.
        If InStrRev(Vdata, "-") > 1 Then
            ImagPart = Val(Mid$(Vdata, InStr(Vdata, "-")))
        ElseIf InStr(Vdata, "+") <> 0 Then
            ImagPart = Val(Mid$(Vdata, InStr(Vdata, "+")))
        
        End If
    Else
        'no Ipart.
        Realpart = Val(Vdata)
    End If


End Property
Public Function Clone() As CComplex
Attribute Clone.VB_Description = "Makes a Copy of the Complex number."
    Dim retval As CComplex
    Set retval = New CComplex
    retval.Realpart = mRealPart
    retval.ImagPart = mImagPart
    Set Clone = retval
End Function
Public Function Add(ValAdd As Variant) As CComplex
Attribute Add.VB_Description = "Adds either another Complex number or a Real number to this class and returns the resulting Complex number."
Dim retcplx As CComplex, Casted As CComplex
Set retcplx = Me.Clone

    If IsObject(ValAdd) Then
        If TypeOf ValAdd Is CComplex Then
            Set Casted = ValAdd
            retcplx.Realpart = retcplx.Realpart + Casted.Realpart
            
            retcplx.ImagPart = retcplx.ImagPart + Casted.ImagPart
        
        End If
    ElseIf IsNumeric(ValAdd) Then
        retcplx.Realpart = mRealPart + ValAdd
        retcplx.ImagPart = mImagPart
    End If

    Set Add = retcplx



End Function

'Subtract
Public Function Subtract(ValDiff As Variant) As CComplex
    If IsObject(ValDiff) Then
        Set Subtract = Add(ValDiff.Negate)
    ElseIf IsNumeric(ValDiff) Then
        Set Subtract = Complex(mRealPart - ValDiff, mImagPart)
    End If
End Function

Public Function compare(CompWith As Variant, Optional ByVal FlDescending As Boolean = False) As Integer
Attribute compare.VB_Description = "compares the Complex number with another complex/Real number."
    'compare real first.
    Dim Ret As Integer
    If IsObject(CompWith) Then
        Ret = Sgn(mRealPart - CompWith.Realpart)
            'equal.
            If Ret = 0 Then
                Ret = Sgn(mImagPart - CompWith.ImagPart)
            End If
            
    ElseIf IsNumeric(CompWith) Then
        Ret = compare(Complex(CompWith, 0), FlDescending)
    End If
    If FlDescending Then
        Ret = -Ret
    End If
    compare = Ret
End Function
'multiply

Public Function Multiply(arg As Variant) As CComplex
Attribute Multiply.VB_Description = "Multiplies this Complex number by another Complex Or real Number."
    If IsObject(arg) Then
        Set Multiply = Complex(mRealPart * arg.Realpart - mImagPart * arg.ImagPart, mRealPart * arg.ImagPart + mImagPart * arg.Realpart)
    ElseIf IsNumeric(arg) Then
        Set Multiply = Multiply(Complex(arg, 0))
    End If
End Function


'divide
Public Function Divide(arg As Variant) As CComplex
Attribute Divide.VB_Description = "Divides this Complex number by another Complex or Real Number."
    Dim out As CComplex
    Dim D As Variant
    Dim u, v
    If IsObject(arg) Then
        'translated from the perl math.pm.
    '        $d = $x2*$x2 + $y2*$y2;
        D = (arg.Realpart ^ 2) + (arg.ImagPart ^ 2)
    '        _divbyzero "$z1/0" if $d == 0;
    '        my $u = ($x1*$x2 + $y1*$y2)/$d;
    '        my $v = ($y1*$x2 - $x1*$y2)/$d;
    'No special error detection- VB's default will find a handler,
    'which will fire off a ParserError Object.
        u = ((mRealPart * arg.Realpart) + (mImagPart * arg.ImagPart)) / D
        v = ((mImagPart * arg.Realpart) - (mRealPart * arg.ImagPart)) / D
       Set Divide = Complex(u, v)
    ElseIf IsNumeric(arg) Then
        Set Divide = Divide(Complex(arg, 0))
    End If
    
'
'        Dim sum As Double
'    ' create the sum only once
'    sum = Real * Real + Imaginary * Imaginary
'    ' evaluate the real and imaginary parts
'    Divide2.Real = (Me.Real * Real + Me.Imaginary * Imaginary) / sum
'    Divide2.Imaginary = (Me.Imaginary * Real - Me.Real * Imaginary) / sum

End Function
'Power
Public Function Negate() As CComplex
Attribute Negate.VB_Description = "Negates the real and imaginary components of this complex number, and returns the result."
    
    Set Negate = Complex(-mRealPart, -mImagPart)
End Function
'grr. can't do powers yet. grrr

Public Function Power(PowerOf As Variant) As CComplex
Attribute Power.VB_Description = "Returns this Complex number raised to a Complex Power."

'z1**z2 = exp(z2 * log z1)).

  
    Set Power = PowerOf.Multiply(Me.Logarithm).EExponential


End Function
Public Function Conjugate() As CComplex
Attribute Conjugate.VB_Description = "Returns the Complex Conjugate of this Complex number."
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = mRealPart
    retcplx.ImagPart = -mImagPart
    Set Conjugate = retcplx



End Function
Public Function EExponential() As CComplex
Attribute EExponential.VB_Description = "Returns the Complex Exponential of this Complex number."
    'exponential-
    'e^(x+i*y)= (e^x)*(Cos(y)+ i(e^x)*(sin(y))
    Dim retcplx As CComplex, Casted As CComplex

    Set retcplx = Me.Clone
    
    With retcplx
        'real part is (e^x) * (Cos(y)
        .Realpart = (e ^ mRealPart) * (Cos(mImagPart))
        .ImagPart = (e ^ mRealPart) * (Sin(mImagPart))
    
    
    End With
    
    
    Set EExponential = retcplx
    
    
End Function



'return (ref $z)->make(CORE::log($r), $t);




Public Function ExpCplx() As CComplex
    'take this
    Set ExpCplx = Complex(Exp(mRealPart), mImagPart)
    'return (ref $z)->emake(CORE::exp($x), $y);

End Function
Public Function Sine() As CComplex
    Dim retcplx As CComplex
    Set retcplx = Me.Clone
    With retcplx
        'Sin(x)Cosh(y) + icos(x)*sinh(y)
        .Realpart = Sin(mRealPart) * hCos(mImagPart)
        .ImagPart = Cos(mRealPart) * hSin(mImagPart)
    
    End With

    Set Sine = retcplx


End Function
'X is real part, Y is imaginary part.

'Cosine
'
'cos(x)cosh(y)
'-isin(x)sinh(y)
Public Function Cosine() As CComplex
Attribute Cosine.VB_Description = "Returns the Complex Cosine of this Complex number."
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = Cos(mRealPart) * hCos(mImagPart)
    retcplx.ImagPart = -(Sin(mRealPart) * hSin(mImagPart))
    Set Cosine = retcplx
    
End Function
Public Function CoTangent() As CComplex
    'Cot (x) = Cos(z)/sin(Z)
    
    Set CoTangent = Me.Cosine.Divide(Me.Sine)


End Function
Public Function Secant() As CComplex
    'sec=1/cos(z)
    Set Secant = Complex(1, 0).Divide(Me.Cosine)
End Function
Public Function CoSecant() As CComplex
    'sec=1/cos(z)
    Set CoSecant = Complex(1, 0).Divide(Me.Sine)
End Function
Public Function Logarithm() As CComplex
Attribute Logarithm.VB_Description = "Returns the Complex Logarithm of this Complex Number."
'complex number to the base e.
'ie:
'log base e of us.
    Set Logarithm = Complex(Log(mRho), mTheta)




End Function
'return (ref $z)->make(CORE::log($r), $t);
'hyperbolic Sine
'
'sinh(x)cos(y)+icosh(x)sin(y)
Public Function Sinh() As CComplex
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = hSin(mRealPart) * Cos(mImagPart)
    retcplx.ImagPart = hCos(mRealPart) * Sin(mImagPart)
    Set Sinh = retcplx
End Function
'Hyperbolic Cosine
'
'cosh(x)cos(y)+isinh(x)sin(y)
Public Function Cosh() As CComplex
Attribute Cosh.VB_Description = "Returns the Complex Hyperbolic Cosine Of this Complex number."
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = hCos(mRealPart) * Cos(mImagPart)
    retcplx.ImagPart = hSin(mRealPart) * Sin(mImagPart)
    Set Cosh = retcplx
End Function

'i*ln(2-sqrt(3))  or i*ln(2+sqrt(3))


'Natural Logarithm
'
'(1/2)ln(x^2+Y^2)+i(atn(y/x)+2kpi)
'(k=0,+-1,+-2,+-....)
Public Function Ln(ByVal K As Long) As CComplex
Dim X, Y
X = mRealPart
Y = mImagPart
Set Ln = Complex((1 / 2) * Log(X ^ 2 + Y ^ 2), (Atn(Y / X) + 2 * K * pi))



End Function
'Tangent
'
'(sin(2x)/cos(2x)_cosh(2y))) +
'isinh(2y)/(cos(2x)+cosh(2y)))
Public Function Tangent() As CComplex
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = (Sin(2 * mRealPart) / (Cos(2 * mRealPart) + hCos(2 * mImagPart)))
    retcplx.ImagPart = hSin(2 * mImagPart) / (Cos(2 * mImagPart) + hCos(2 * mImagPart))
    Set Tangent = retcplx
End Function
'Hyperbolic Tangent
'
'(sinh(2x)/cosh(2x)+Cos(2y))) +
'isin(2y)/(cosh(2x)+cos(2y))
Public Function Tanh() As CComplex
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = (Sin(2 * mRealPart) / (Cos(2 * mRealPart) + hCos(2 * mImagPart)))
    retcplx.ImagPart = (hSin(2 * mImagPart) / Cos(2 * mRealPart) + hCos(2 * mImagPart))
    Set Tanh = retcplx
End Function
'Cotangent
'
'(Sin(2x))
'-isinh(2y))/cosh(2y)-cos(2x))


'Hyperbolic coTangent
'sinh(2x)
'-isin(2y))/(cosh(2x) - cos(2y))





'DERIVED MATH FUNCTIONS.
'placed here to make CCOmplex more encapsulated. Not that it makes a difference, it's
'already tightly coupled with BASeParser (IE, the CParser method calls)
Private Function ArcSin(ByVal radial As Variant) As Variant
On Error Resume Next
ArcSin = Atn(CDec(radial) / Sqr(-radial * radial + 1))
End Function
Private Function ArcCos(ByVal radial As Variant) As Variant
ArcCos = Atn(CDec(radial) / Sqr(-radial * radial + 1)) + 1.5708
End Function
Private Function Sec(ByVal radial As Variant) As Variant
Sec = 1 / Cos(CDec(radial))
End Function
Private Function Cosec(ByVal radial As Variant) As Variant
Cosec = 1 / Sin(radial)
End Function
Private Function Cotan(ByVal radial As Variant) As Variant
Cotan = 1 / Tan(radial)
End Function
Private Function ArcSec(ByVal radial As Variant) As Variant
ArcSec = Atn(radial / Sqr(-radial * radial + 1)) + Sgn(Sgn(radial) - 1) * 1.5708
End Function
Private Function ArcCoSec(ByVal radial As Variant) As Variant
ArcCoSec = Atn(radial / Sqr(radial * radial - 1)) + (Sgn(radial) - 1) * 1.5708
End Function
Private Function ArcCoTan(ByVal radial As Variant) As Variant
ArcCoTan = Atn(radial) * 1.5708
End Function
Private Function hSin(ByVal radial As Variant) As Variant
hSin = (Exp(radial) - Exp(-radial)) / 2
End Function

Private Function hCos(ByVal radial As Variant) As Variant
hCos = (Exp(radial) + Exp(-radial)) / 2
End Function
Private Function hTan(ByVal radial As Variant) As Variant
hTan = (Exp(radial) - Exp(-radial)) / (Exp(radial) + Exp(-radial))
End Function
Private Function HSec(ByVal radial As Variant) As Variant
HSec = 2 / (Exp(radial) + Exp(-radial))
End Function
Private Function HCoSec(ByVal radial As Variant) As Variant
HCoSec = 2 / (Exp(radial) - Exp(-radial))
End Function
Private Function HCoTan(ByVal radial As Variant) As Variant
HCoTan = (Exp(radial) + Exp(-radial)) / (Exp(radial) - Exp(-radial))
End Function
Private Function HArcSin(ByVal radial As Variant) As Variant
HArcSin = Log(radial + Sqr(radial * radial + 1))
End Function

Private Function HArcCos(ByVal radial As Variant) As Variant
HArcCos = Log(radial + Sqr(radial * radial - 1))
End Function
Private Function HArcTan(ByVal radial As Variant) As Variant
HArcTan = Log(1 + radial) / (1 - radial) / 2
End Function
Private Function HArcSec(ByVal radial As Variant) As Variant
HArcSec = Log((Sqr(-radial * radial + 1) = 1) / radial)
End Function
Private Function HArcCoSec(ByVal radial As Variant) As Variant
HArcCoSec = Log((Sgn(radial) * Sqr(radial * radial + 1) + 1) / radial)
End Function
Private Function HArcCoTan(ByVal radial As Variant) As Variant
HArcCoTan = Log(radial + 1) / (radial - 1) / 2
End Function
Private Function Complex(ByVal Realpart As Variant, ByVal ImagPart As Variant) As CComplex
    Dim retcplx As CComplex
    Set retcplx = New CComplex
    retcplx.Realpart = Realpart
    retcplx.ImagPart = ImagPart
    Set Complex = retcplx
End Function
Public Function Operation(ByVal op As String, Otherval As Variant) As Variant
Attribute Operation.VB_Description = "Performs a operation given one of several operands (+,-,==,<=, etc...)"
''" - + * \ / ^ ** == != <> <= =< >= >= "
If IsObject(Otherval) Then
    If TypeOf Otherval Is CComplex Then
    Select Case UCase$(op)
        Case "+"
            Set Operation = Add(Otherval)
        Case "-"
            Set Operation = Subtract(Otherval)
        Case "*"
            Set Operation = Multiply(Otherval)
        Case "\"
            Set Operation = Divide(Otherval)
        Case "/"
            Set Operation = Divide(Otherval)
        Case "=="
            Operation = compare(Otherval) = 0
        Case "!=", "<>"
            Operation = compare(Otherval) <> 0
        Case "<"
            Operation = compare(Otherval) = -1
        Case ">"
            Operation = compare(Otherval) = 1
        Case "^", "**"
            Set Operation = Me.Power(Otherval)
        Case Else
            'Uh oh!
            Err.Raise 13, "CComplex::Operation", "Operation """ & op & """ Not defined for Complex Numbers."
    End Select
    Else
        Err.Raise 13
    End If
ElseIf IsNumeric(Otherval) Then
    Set Operation = Operation(op, Complex(Otherval, 0))
End If
End Function

Public Function Root(ByVal n As Long) As CComplex()
Attribute Root.VB_Description = "Returns the Nth root of this Complex number."
'The roots are given by (for k = 0..n-1):
'#
'z = us
'n = root
'n must be positive integer.
'# z^(1/n) = r^(1/n) (cos ((t+2 k pi)/n) + i sin ((t+2 k pi)/n))
Dim Ret() As CComplex, K As Long
ReDim Ret(0 To n)
For K = 0 To n

'root = r^(1/n) * ((t+2 k pi)/n) + i Sin((t+2 k pi)/n))

Set Ret(K) = Complex(mRho ^ (1 / n) * (mTheta + 2 * K * pi) / n, Sin((mTheta + 2 * K * pi) / n))



Next K
'exp(z2 * log z1))
Root = Ret


End Function
Private Function pi() As Variant
Attribute pi.VB_Description = "Returns PI."
    pi = CDec(Atn(1)) * 4
End Function

Private Function IOperable_PerformFunc(withparser As CParser, ByVal funcName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants
    '
    'functions "subclassed" by IOperable:
    'anything. How? you ask?
    'well, we subclass Sin,Cos, Tan, Atn, etc. However, any other function is
    'performed on the Real and imaginary parts separately.
    Dim IArgs() As Variant
    Dim RArgs() As Variant
    Dim rvalI As Variant, RValR As Variant
    Dim retme As CComplex
    IOperable_PerformFunc = Perform_Success
    Select Case UCase$(funcName)
        Case "ARRAY"
            retval = AllArguments
        Case "STR", "STR$"
           retval = Me.StringForm
        Case "SIN"
            Set retval = Me.Sine
        Case "COS"
            Set retval = Me.Cosine
        Case "TAN"
            Set retval = Me.Tangent
            'add others here.
            'to lazy now-- Friday May 04 2007
            
            'ArcTan,ArcSin,ArcCosine:
        Case "ARCSIN"
            
            
        Case "CONJ"
            Set retval = Me.Conjugate
        Case "COSH"
            Set retval = Me.Cosh
        Case Else
            'trickster tactic, I am very used to those.
            'recreate the parameter lists for the real and imaginary portions of
            'the Complex number we represent.
            
            IArgs = AllArguments
            RArgs = AllArguments
            'since the first argument should be the reference to us, we need to
            'replace it with the real and imaginary components and use the return values
            'to fill the respective fields of a new CComplex class.
            Set retval = New CComplex
            RArgs(LBound(RArgs)) = Me.Realpart
            IArgs(LBound(IArgs)) = Me.ImagPart
            
            
            'BUGFIX:
            'feb 21st 2009
            
            
            
            'now, we simply call (ha, simply- good joke) the parser
            'to do the function call again.
            Call withparser.EvalListener.Self.HandleFunction(withparser, funcName, RArgs(), RValR)
            Call withparser.EvalListener.Self.HandleFunction(withparser, funcName, IArgs(), rvalI)
            
            
            Set retme = New CComplex
            retme.Realpart = RValR
            retme.ImagPart = rvalI
            'there. simple enough.
            Set retval = retme
    End Select
    
        


End Function
Public Function Equals(Othervalue As CComplex) As Boolean
    
        If Me.Realpart = Othervalue.Realpart And Me.ImagPart = Othervalue.ImagPart Then Equals = True Else Equals = False
   




End Function
Private Sub IOperable_PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, retval As Variant)
    'CComplex Class IOperable implementation.
    'First, this class can only work with two Complex numbers. So of the given operand (assuming it isn't a unary operation) is
    'not a Complex number, convert it.
    Dim convComplex As CComplex
    Dim PerformHack As CComplex
    
    Dim NewReal, NewImag
    
    Dim OtherCplx As CComplex
    Dim i As Long
    Dim RetArray()
    'STEP ONE:
    
    If IsArray(OtherOperand) Then
        'If the "OtherOperand" is an Array, we will recursively call PerformOp (not our performop, mind you, the parsers) to create a new result and return that:
        
        'Also, if OtherOperand is really an array of Arrays, that's fine as well, since the recursed procedures will execute this same code.
        
        ReDim RetArray(LBound(OtherOperand) To UBound(OtherOperand))
        For i = LBound(OtherOperand) To UBound(OtherOperand)
            If DoUnary Then
                Debug.Assert False
            Else
        
        
                Call withparser.PerformOperation(Operation, Me, OtherOperand, RetArray(i))
            End If
        
        Next i
        
        retval = RetArray
        
        
        
        
    
        
    
    
    
    ElseIf DoUnary Then
        'good, a unary op.
        'firstly, check for those operations we do support.
        Select Case Operation
            Case "-"
                'return the "negate"
                Set retval = Me.Negate
            Case "~"
                'return Complex conjugate.
                Set retval = Me.Conjugate
            
            Case Else
                'uh oh- but all is not lost.
                'clone this Complex number, and ask the Cparser class to perform a unary operation
                'on each one.
                
                'Set PerformHack = Me.Clone()
                withparser.PerformUnaryOp Operation, mRealPart, DoUnary = UnaryOp_Suffix, NewReal
                'now with Imaginary part.
                withparser.PerformUnaryOp Operation, mImagPart, DoUnary = UnaryOp_Suffix, NewImag
                
                'if either of the returned values is NOT numeric, raise an appropriate error.
                If Not IsNumeric(NewReal) Then
                    Err.Raise 5, "CComplex::IOperable.PerformOp", "Recursive Unary Scalar operation returned a Non-numeric entity for the Real Component."
                ElseIf Not IsNumeric(NewImag) Then
                    Err.Raise 5, "CComplex::IOperable.PerformOp", "Recursive Unary Scalar operation returned a Numeric entity for the imaginary component."
                Else
                    'whew.
                    Set PerformHack = Complex(NewReal, NewImag)
                    
                    Set retval = PerformHack
                
                End If
                
                
        End Select
    Else
        'Binary
        'CCOmplex Supports +,-,\,/
        'But first- we need to make the operand an instance of Ccomplex.
        If Not (TypeOf OtherOperand Is CComplex) Then
            
                'It's an array.
            
            'cast it as such. is it numeric? of so,
            'simply create a new complex number with the given real component.
            If IsNumeric(OtherOperand) Then
                Set OtherCplx = Complex(OtherOperand, 0)
            ElseIf VarType(OtherOperand) = vbString Then
                'CComplex can't deal with strings.
                Err.Raise 5, "CComplex::Ioperable.PerformOp", "CComplex Ioperable Implementation cannot handle Strings."
                
            ElseIf IsObject(OtherOperand) Then
                'remember- if we get here:
                Debug.Assert Not (TypeOf OtherOperand Is IOperable)
                'it will NOT be an IOperable interface.
                'that is what I assert.
                'the BPCoreOpFunc delegation code should have ensured it isn't- and in the case it is, it will
                'be CComplex.
                'oh yeah. This class doesn't support it.
                Err.Raise 5, "CComplex::Ioperable.PerformOp", "CComplex Class cannot handle other operands that are objects."
            
            End If
        Else
            Set OtherCplx = OtherOperand
        
        
        
        End If
        
    
    'OK, there is an OtherOperand.
        Select Case Operation
        Case "+"
            Set retval = Me.Add(OtherCplx)
        Case "-"
            Set retval = Me.Subtract(OtherCplx)
        Case "*"
            Set retval = Me.Multiply(OtherCplx)
        Case "/", "\"
            Set retval = Me.Divide(OtherCplx)
        Case "="
            retval = Me.Equals(OtherCplx)
        Case Else
            Err.Raise 5, "CComplex::PerformOp", "CComplex does now how to " & Operation
        End Select
    
    End If
    
    
    
        
End Sub

Private Function IOperable_toString(withparser As CParser) As String
    Set mlastparser = withparser
    IOperable_toString = Me.StringForm
End Function
