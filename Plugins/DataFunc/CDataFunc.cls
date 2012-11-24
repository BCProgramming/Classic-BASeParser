VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implement the BASeParserXP Plugin interface(s).

Implements IEvalEvents
Implements IOpFuncInfo
Implements IPluginUIData
Private mCon As ADODB.Connection


Private Sub Class_Initialize()
    'initialize: get the database connection.
    'this plugin is actually pretty easy from my perspective (I did write the Core
    'Operators and functions, which was a LOT of code...)
    'all it does it get the connection at start-up, and then, every time it
    'connects to a parser, it simply loads the respective Functions as CFunction Objects, into
    'the CParsers Functions collection.
    
    'What is the database name though?
    Set mCon = GetConnection()
    'Also, New additions to the Database will not be visible in parsers that have already connected.
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    mCon.Close
    Set mCon = Nothing
End Sub
Private Function IEvalEvents_AssignSubScript(Withparser As CParser, OnValue As Variant, subscripts As Variant, assignthis As Variant) As Boolean
'
End Function
Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_BeforeInvokeObject(ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
'Still nothing here too.
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(Withparser As BASeParserXP.CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'
'we don't need this- instead we add our functions to the respective
'Withparser Argument, then the FunctionHandler plug-in (if enabled) will use them
'of course, if the FunctionHandler plugin ISN'T loaded, we raised an error in the initialization anyway.
End Function

Private Function IEvalEvents_CanHandleOperator(Withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'
End Function

Private Property Get IEvalEvents_Description() As String
'
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(Withparser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(Withparser As BASeParserXP.CParser) As String
'
End Function

Private Function IEvalEvents_GetHandledOperatorsString(Withparser As BASeParserXP.CParser) As String
'
End Function

Private Function IEvalEvents_GetInfoObject() As BASeParserXP.IOpFuncInfo
'
Set IEvalEvents_GetInfoObject = Me
End Function

Private Sub IEvalEvents_GetOperation(Withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Function IEvalEvents_GetOperatorFlags(Withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As BASeParserXP.OperatorFlags
'
End Function

Private Function IEvalEvents_GetPluginUIData() As BASeParserXP.IPluginUIData
'
Set IEvalEvents_GetPluginUIData = Me
End Function

Private Sub IEvalEvents_GetUnaryOperation(Withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Sub IEvalEvents_HandleFunction(Withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'
End Sub

Private Function IEvalEvents_HandleSubScript(Withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(Withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
    'Here is a neat-o usage of the result code.
    'if our Database connection failed, return a error in the initialization.
    'You know, I don't really remember learning about ADO...
    Dim recOpen As ADODB.Recordset
    'if the passed in parser has unloaded the Functionhander plugin, we don't fail- rather, we
    'output a warning to the debug.
    If Not Withparser.PluginLoaded("BASeParserXP.FunctionHandler") Then
        'Egads! It isn't loaded.
        IEvalEvents_Init = Code_InitFailure
    End If
    
    If mCon Is Nothing Then
        IEvalEvents_Init = Code_InitFailure
    Else
        'retrieve the recordset, peruse it,  add the Cfunction objects for each record,
        'then close the recordset.
        Set recOpen = New ADODB.Recordset
        Call recOpen.Open("SELECT * FROM TblFunctions", mCon, adOpenForwardOnly, adLockOptimistic)
        'if it actually is open...
        If recOpen.State = adStateOpen Then
        recOpen.MoveFirst
        Debug.Print "iterating through recordset..."
        Do Until recOpen.EOF
        'add it.
            With Withparser.Functions.Add(recOpen.Fields("Expression").Value, recOpen.Fields("Name").Value)
                Debug.Print "added function, Name=" & .Name & " , Expression=" & .Expression
            
            End With
            
        'and move on.
            recOpen.MoveNext
        
        
        Loop
        
        
        
        End If
    
    
    
    End If
End Function

Private Property Get IEvalEvents_Name() As String
IEvalEvents_Name = "ADO data source to Functions BASeParser Plugin."
End Property

Private Function IEvalEvents_OnInvokeObject(Withparser As CParser, ObjInvoke As Variant, currItem As BASeParserXP.CFormItem, Result As Variant) As Boolean
'
End Function

Private Sub IEvalEvents_ParseComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Sub IEvalEvents_PreExecute(Withparser As BASeParserXP.CParser, FirstItem As BASeParserXP.CFormItem)
'
End Sub

Private Sub IEvalEvents_PreParse(Parser As BASeParserXP.CParser, Expression As String)
'
End Sub

Private Property Get IEvalEvents_Self() As Object
Set IEvalEvents_Self = Me
End Property

Private Function IEvalEvents_Term(Withparser As BASeParserXP.CParser) As BASeParserXP.IPluginTermCode
'
'hmm- do we destroy it yet?

End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function

Private Sub IOpFuncInfo_GetFunctionInformation(Withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncInfo As BASeParserXP.FUNCTIONINFORMATION)
'
End Sub

Private Property Get IPluginUIData_Description() As String
IPluginUIData_Description = "BASeParserXP Function Database Handler"
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As BASeParserXP.PluginPictureSizeConstants) As stdole.Picture
'
'no pic
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String
'
End Function
