VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBPGUIKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IEvalEvents

Private Const HandledFunctions As String = " NEWFORM "

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_AssignSubScript(withparser As BASeParserXP.CParser, OnValue As Variant, subscripts As Variant, AssignThis As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeFunctionCall(withparser As BASeParserXP.CParser, FuncCalling As String, WithHandler As BASeParserXP.IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As BASeParserXP.CParser, ByVal funcName As String, NoParsedArgs() As Boolean) As Boolean
'
IEvalEvents_CanHandleFunction = InStr(1, HandledFunctions, " " & funcName & " ", vbTextCompare) > 0
End Function

Private Function IEvalEvents_CanHandleOperator(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'
IEvalEvents_CanHandleOperator = False
End Function

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
IEvalEvents_GetFunctionFlags = Function_Implemented
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As BASeParserXP.CParser) As String
'
IEvalEvents_GetHandledFunctionString = HandledFunctions
End Function

Private Function IEvalEvents_GetHandledOperatorsString(withparser As BASeParserXP.CParser) As String
'
End Function

Private Function IEvalEvents_GetInfoObject() As BASeParserXP.IOpFuncInfo
'
End Function

Private Sub IEvalEvents_GetOperation(withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Function IEvalEvents_GetOperatorFlags(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As BASeParserXP.OperatorFlags
'
End Function

Private Function IEvalEvents_GetPluginUIData() As BASeParserXP.IPluginUIData
'
End Function

Private Sub IEvalEvents_GetUnaryOperation(withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Sub IEvalEvents_HandleFunction(withparser As BASeParserXP.CParser, ByVal funcName As String, FuncParams() As Variant, ReturnVal As Variant)
'
Dim ret As Variant
If UCase$(funcName) = "NEWFORM" Then
    Set ret = New FrmRuntimeDialog
    
End If
    Assign ReturnVal, ret

End Sub

Private Function IEvalEvents_HandleSubScript(withparser As BASeParserXP.CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
'
IEvalEvents_Init = Code_Success
End Function

Private Property Get IEvalEvents_Name() As String
'
IEvalEvents_Name = "BPGUIKit"
End Property

Private Function IEvalEvents_OnInvokeObject(withparser As BASeParserXP.CParser, ObjInvoke As Variant, currItem As BASeParserXP.CFormItem, Result As Variant) As Boolean
'
End Function

Private Sub IEvalEvents_ParseComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Sub IEvalEvents_PreExecute(withparser As BASeParserXP.CParser, FirstItem As BASeParserXP.CFormItem)
'
End Sub

Private Sub IEvalEvents_PreParse(Parser As BASeParserXP.CParser, Expression As String)
'
End Sub

Private Property Get IEvalEvents_Self() As Object
'
Set IEvalEvents_Self = Me
End Property

Private Property Let IEvalEvents_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get IEvalEvents_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String) As Variant
'
End Property

Private Function IEvalEvents_Term(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginTermCode
'
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function
