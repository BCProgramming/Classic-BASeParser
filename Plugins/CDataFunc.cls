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
'Current(old) implementation

'The NEW implementation:
'rather then get somebody else to do the work, I'll create a new flag on the CFunction class,
'which will cause that Class instance to Call into another class that implements IEvalEvents to perform the
'function. Sure, It requires modification to a core class, CFunction, but, it'll be useful in the future I'm sure.


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
    
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    mCon.Close
    Set mCon = Nothing
End Sub
Private Function IEvalEvents_AssignSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, assignthis As Variant) As Boolean
'
End Function
Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_BeforeFunctionCall(withparser As CParser, FuncCalling As String, WithHandler As BASeParserXP.IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeInvokeObject(ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
'Still nothing here too.
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean

If mCon Is Nothing Then
    'no connection, no functions, no dice.
RECOPENERROR:
    IEvalEvents_CanHandleFunction = False
    
    
    Exit Function
Else
    IEvalEvents_CanHandleFunction = (InStr(1, IEvalEvents_GetHandledFunctionString(withparser), " " & FuncName & " ", vbTextCompare) > 0)
    


End If


End Function

Private Function IEvalEvents_CanHandleOperator(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'
End Function

Private Property Get IEvalEvents_Description() As String
'
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As BASeParserXP.CParser) As String
'    '
'    'quick! iterate through the recordset and retrieve the strings!
'    Dim retval As cStringBuilder
'    Static RsetLook As ADODB.Recordset
'    Set retval = New cStringBuilder
'    On Error GoTo NOFUNCTIONS
'
'    If RsetLook Is Nothing Then
'        Set RsetLook = New ADODB.Recordset
'        RsetLook.Open "Select * From TblFunctions", mCon, adOpenForwardOnly, adLockOptimistic
'    End If
'
'    RsetLook.MoveFirst
'
'    Do Until RsetLook.EOF
'
'        retval.Append " "
'        retval.Append RsetLook.Fields("Name").Value
'        retval.Append " "
'        'important! :)
'        RsetLook.MoveNext
'
'    Loop
'    IEvalEvents_GetHandledFunctionString = retval.ToString
'    Exit Function
'NOFUNCTIONS:
'    IEvalEvents_GetHandledFunctionString = ""
End Function

Private Function IEvalEvents_GetHandledOperatorsString(withparser As BASeParserXP.CParser) As String
'
End Function

Private Function IEvalEvents_GetInfoObject() As BASeParserXP.IOpFuncInfo
'
Set IEvalEvents_GetInfoObject = Me
End Function

Private Sub IEvalEvents_GetOperation(withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Function IEvalEvents_GetOperatorFlags(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As BASeParserXP.OperatorFlags
'
End Function

Private Function IEvalEvents_GetPluginUIData() As BASeParserXP.IPluginUIData
'
Set IEvalEvents_GetPluginUIData = Me
End Function

Private Sub IEvalEvents_GetUnaryOperation(withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Sub IEvalEvents_HandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'
End Sub

Private Function IEvalEvents_HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
  'Here is a neat-o usage of the result code.
    'if our Database connection failed, return a error in the initialization.
    'You know, I don't really remember learning about ADO...
    Dim recOpen As ADODB.Recordset
    'if the passed in parser has unloaded the Functionhander plugin, we don't fail- rather, we
    'output a warning to the debug.
    If Not withparser.PluginLoaded("BASeParserXP.FunctionHandler") Then
        'Egads! It isn't loaded.
        IEvalEvents_Init = Code_InitFailure
        Exit Function
    End If
    
    If mCon Is Nothing Then
        IEvalEvents_Init = Code_InitFailure
    Else
        'retrieve the recordset, peruse it,  add the Cfunction objects for each record,
        'then close the recordset.
        'ideally we should keep it open the whole time,
        Set recOpen = New ADODB.Recordset
        Call recOpen.Open("SELECT * FROM TblFunctions", mCon, adOpenForwardOnly, adLockOptimistic)
        'if it actually is open...
        If recOpen.State = adStateOpen Then
        recOpen.MoveFirst
        Debug.Print "iterating through recordset..."
        Do Until recOpen.EOF
        'add it.
            With withparser.functions.Add(recOpen.Fields("Expression").Value, recOpen.Fields("Name").Value)
                Debug.Print "added function, Name=" & .Name & " , Expression=" & .Expression
                'Set .FunctionCallBack = Me
            End With
            
        'and move on.
            recOpen.MoveNext
        
        
        Loop
        
        
        
        End If
    
    
    
    End If
    recOpen.Close
End Function


Private Property Get IEvalEvents_Name() As String
IEvalEvents_Name = "ADO data source to Functions BASeParser Plugin."
End Property

Private Function IEvalEvents_OnInvokeObject(withparser As CParser, ObjInvoke As Variant, currItem As BASeParserXP.CFormItem, Result As Variant) As Boolean
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
'hmm- do we destroy it yet?

End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function

Private Sub IOpFuncInfo_GetFunctionInformation(withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncInfo As BASeParserXP.FUNCTIONINFORMATION)
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
