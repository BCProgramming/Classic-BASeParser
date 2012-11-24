VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParserError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CParserError:
'represents a error that occured during the parse of an expression. We want
'to allow for clients to find out exactly where in the expression the error occured, since
'the parser does have this information when it encounters an error.


Implements IOperable
Private mInParser As CParser
'the actual parser object whose RipFormula method fired then error.
'interesting to note it may not be the same as the one that fired the event
Private mPosition As Long      'position in the string where the error occured.
Private mCode As Long          'usually a Parser Error Code, but watch out for VB errors
                                    'that might pilfer through from the plugins- doubtful, however.
Private mSource As String
Private mDescription As String


Friend Property Set InParser(Vdata As CParser)
    Set mInParser = Vdata
End Property
Public Property Get InParser() As CParser
    Set InParser = mInParser
End Property

Friend Property Let Position(ByVal Vdata As Long)
    mPosition = Vdata
End Property
Public Property Get Position() As Long
    Position = mPosition
End Property
Friend Property Let Code(ByVal Vdata As Long)
    mCode = Vdata
End Property
Public Property Get Code() As Long
    Code = mCode
End Property
Friend Property Let Description(ByVal Vdata As String)
    mDescription = Vdata
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
Friend Property Let Source(ByVal Vdata As String)
    mSource = Vdata
End Property
Public Property Get Source() As String
    Source = mSource
End Property
Public Function ToString() As String
    Dim StrBuild As String
    If mCode <> 0 Then
        StrBuild = " Code #" & mCode
    End If
    If mDescription <> "" Then
        StrBuild = StrBuild & "(" & mDescription & ")"
    End If
    If mSource <> "" Then
        StrBuild = StrBuild & " occured in " & mSource & " "
        
    End If
    If mPosition > 0 Then
        StrBuild = StrBuild & " At position " & mPosition
    End If
    ToString = StrBuild
End Function
Public Function Throw() As ParserErrorRecoveryConstants
    Dim returnValue As ParserErrorRecoveryConstants
    'throws this error. not how you'd expect.
    If mInParser Is Nothing Then
        returnValue = PERR_FATAL
    Else
        
        mInParser.RaiseParseError Me, returnValue
    End If
    'return the raised stuff.
    Throw = returnValue
End Function


Private Function IOperable_PerformFunc(withparser As CParser, ByVal FuncName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants
IOperable_PerformFunc = Perform_Unknown
End Function

Private Sub IOperable_PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, retval As Variant)
    '
    Err.Raise 5, "CParserError.IOperable::PerformOp", "CParserError does not support PerformOp."

End Sub

Private Function IOperable_toString(withparser As CParser) As String
'
IOperable_toString = "!Error #" & mCode & "(""" & mDescription & """)!"
End Function
