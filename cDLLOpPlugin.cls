VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDLLPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'cDLLPlugin
'will contain implementation of code for interfacing with a run of the mill DLL on disk. I will likely need to
'trim down the set of required exports from those required in the IEvalEvents Interface, however.
'The tricky part is that this Plugin will need to support the use of multiple Libraries (Via settings) since the Parser can't have two plugins that are exactly the same (OK, it can, but It seems a bit verbose to me.)

Implements IEvalEvents
'implements the IEvalEvents interface- Duh, we are a Parser Plugin!

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'NO!
End Function

Private Function IEvalEvents_CanHandleOperator(withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As Boolean
'maybe...

End Function

Private Sub IEvalEvents_ExecuteComplete(ParserDone As CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As CParser, ByVal FuncString As String) As FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As CParser) As String
'
End Function

Private Function IEvalEvents_GetHandledOperatorsString(withparser As CParser) As String
'
End Function

Private Function IEvalEvents_GetInfoObject() As IOpFuncInfo
'
End Function

Private Sub IEvalEvents_GetOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval As Variant)
'
End Sub

Private Function IEvalEvents_GetOperatorFlags(withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As OperatorFlags
'
End Function

Private Function IEvalEvents_GetPluginUIData() As IPluginUIData
'
End Function

Private Sub IEvalEvents_GetUnaryOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As CFormItem, retval As Variant)
'
End Sub

Private Sub IEvalEvents_HandleFunction(withparser As CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'
End Sub

Private Function IEvalEvents_HandleSubScript(OnValue As Variant, Subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(withparser As CParser) As IPluginInitCode
'
End Function

Private Property Get IEvalEvents_Name() As String
'
End Property

Private Function IEvalEvents_OnInvokeObject(withparser As CParser, ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
'
End Function

Private Sub IEvalEvents_ParseComplete(ParserDone As CParser)
'
End Sub

Private Sub IEvalEvents_PreExecute(withparser As CParser, FirstItem As CFormItem)
'
End Sub

Private Sub IEvalEvents_PreParse(Parser As CParser, Expression As String)
'No preparse code.
End Sub

Private Property Get IEvalEvents_Self() As Object
Set IEvalEvents_Self = Me
End Property

Private Function IEvalEvents_Term(withparser As CParser) As IPluginTermCode
'
'
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function
