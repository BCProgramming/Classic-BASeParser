VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOpFuncInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IOpFuncInfo:
'used to retrieve help/assistory information on implemented functions and operators.
'the information retrieved from this class can be presented to the user or parsed for some kind of retarded dynamic call.
'to present this information to the user, the CParser class will implement
'a "ParseFunctionInformation" method, which looks at the
'given FUNCTIONINFORMATION and creates a string not unlike that presented
'by Visual Basic's Object browser.

'generally implementing this class isn't necessary for proper operation, and under most circumstances
'simply adds polish.
Public Enum ParameterHandlingConstants
    Parameter_ByVal = 0
    Parameter_ByRef = 1
    
End Enum
Public Type PARAMETERINFORMATION
    Name As String
    DataType As String
    isOptional As Boolean
    isParamArray As Boolean
    Description As String
    HandleType As ParameterHandlingConstants
End Type
Public Type FUNCTIONINFORMATION
    StrFunctionName As String
    ParameterCount As Long
    ParameterInfo() As PARAMETERINFORMATION
    ReturnType As String
    ReturnDescription As String
    StrDescription As String
    StrSample   As String   'a Sample Expression, if available
    StrHelpHTML As String       'OPTIONAL- a string that is the HTML help text for the given function.
    'In my BPCOREOPFUNC class, I plan to implement these through resource files, whose
    'ID is equivalent to the name of the function.
End Type




Public Sub GetFunctionInformation(withparser As CParser, ByVal FuncName As String, ByRef FuncInfo As FUNCTIONINFORMATION)
    'populates the given FuncInfo structure with information on the given function.
End Sub
