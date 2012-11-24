VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEvalEventSinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEvalEvents
Implements IOpFuncInfo

'Parser Event sinks.
'TODO:\\ make this into a collection class (Item,RemoveImplementor,etc..
'right now this is actually- the PluginManager for the library.


'//December 9th 2008

'Optimization plan:

'as it stands at this moment, the methods all simply delegate to the collection.

'BUT: since this class exists, it may as well perform another useful function, no?

'So, I propose that it caches return values; for example, the GethandledFunctionString, GethandledOperatorString, etc.
'that is fairly straightforward- simply cache the string for subsequent returns flipping a boolean when the add method is called

'problem being the whole idea was to be dynamic about it.



Public Event PluginBeforeAdd(ByVal ObjAdd As IEvalEvents, ByRef cancel As Boolean)
Public Event PluginAfterAdd(objAdded As IEvalEvents)
Public Event PluginAddUnknown(ObjAdd As Object)
Public Event PluginBeforeRemove(ObjRemoving As IEvalEvents, ByRef cancel As Boolean)
Public Event PluginAfterRemove(ObjRemoved As IEvalEvents)
Private mParser As CParser
Private mCol As Collection, loopobject As IEvalEvents

Private Function IEvalEvents_AssignSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, AssignThis As Variant) As Boolean
'

    For Each loopobject In mCol
    
        'loopobject.AfterResolveVariable ParserUse, VarName, VarValue
        If loopobject.AssignSubScript(withparser, OnValue, subscripts, AssignThis) Then
            Exit For
        End If
    
    Next

End Function
Friend Property Set Parser(Vdata As CParser)
    Set mParser = Vdata
End Property
Friend Property Get Parser() As CParser
    Set Parser = mParser
End Property
Friend Function GetCollection() As Collection
    Set GetCollection = mCol
End Function
Public Function Add(ObjAdd As Variant) As IEvalEvents
    Dim bCancel As Boolean, ValEvents As IEvalEvents
    Dim ObjCast As Object
    On Error GoTo Error_Add
    'if it is a ProgID, we should Create it, then verify it is a IEvalEvents Implementor.
    If VarType(ObjAdd) = vbString Then
        Set ValEvents = createObject(CStr(ObjAdd))
    ElseIf VarType(ObjAdd) = vbObject Then
    'if it is IevalEvents, good- otherwise, raise the AddUnknown event.
        If TypeOf ObjAdd Is IEvalEvents Then
            'Set ObjCast = ValEvents
            Set ValEvents = ObjCast
        Else
            Set ObjCast = ObjAdd
            RaiseEvent PluginAddUnknown(ObjCast)
            Set ValEvents = ObjAdd
        End If
    End If
    
    
    If ObjAdd Is Nothing Then Exit Function
    'raise the before event.
    RaiseEvent PluginBeforeAdd(ValEvents, bCancel)
    'call the Init() method.
    
    If bCancel Then
        'TODO:\\ probably SHOULD check to make sure
        'the core handler loading isn't being cancelled...
        'we haven't actually added it, so just call term.
        ValEvents.Term mParser
    Else
        'AddImplementor objAdd
        mCol.Add ObjAdd
    End If
    'raise the After event.
    RaiseEvent PluginAfterAdd(ValEvents)
    Set Add = ObjAdd
    Exit Function
Error_Add:
    Debug.Assert False
End Function
Public Function FindIndex(ByVal ObjFind As IEvalEvents) As Long
    Dim looper As Long
    looper = 0
    For looper = mCol.Count To 1 Step -1
        If mCol.Item(looper) Is ObjFind Then
            FindIndex = looper
            Exit Function
        End If
    
    Next looper
    
End Function
Public Function Remove(VarRemove As Variant) As IEvalEvents
    'remove it, and return it.
    Dim IndRemove As Variant, IEvCast As IEvalEvents
    Dim flCancel As Boolean
    If IsObject(VarRemove) Then
        Set IEvCast = VarRemove
        IndRemove = FindIndex(IEvCast)
    Else
        IndRemove = VarRemove
    End If
    'before we remove, we should call the Interface method, Term(), so the plugin can cleanup it's
    'resources if it used any.
    'check for nothing. if it isn't nothing, then their is no use setting it to itself. Besides
    'indexing into the collection for no reason is wasteful.
    If IEvCast Is Nothing Then Set IEvCast = mCol.Item(IndRemove)
    If IEvCast.Term(mParser) = Code_ForceRemain Then
        'if forced itself to stay. Hmm.
        'pretend Remove was never called. Also, return nothing as a flag.
        Set Remove = Nothing
        
        Exit Function
    
    
    End If
    Set Remove = IEvCast
    RaiseEvent PluginBeforeRemove(IEvCast, flCancel)
    If flCancel Then
        Set Remove = Nothing
    Else
        mCol.Remove IndRemove
        RaiseEvent PluginAfterRemove(IEvCast)
    End If
End Function
Public Sub AddImplementor(ObjAdd As Object)
    On Error GoTo AddError
    'CDebug.Post "adding implementor.(in addimplementor)" & GetObjectProgID(ObjAdd)

    'ParserSettings.VerifyPlugin GetObjectProgID(ObjAdd)
    'verification was buggy and crashed when compiled; very strange.
    Add ObjAdd
    
    Exit Sub
AddError:
    CDebug.Post "Error occured in addImplementor: " & Err.Description & "," & Err.Source
    'Stop   'force halt
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub
Public Function Self() As IEvalEvents
    Set Self = Me
End Function
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()

    Set mCol = Nothing
End Sub

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As CParser, ByVal VarName As String, VarValue As Variant)
'
    For Each loopobject In mCol
    
        loopobject.AfterResolveVariable ParserUse, VarName, VarValue
    
    Next
End Sub

Private Function IEvalEvents_BeforeFunctionCall(withparser As CParser, FuncCalling As String, WithHandler As IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
Dim doeshandle As Boolean, originalhandler As IEvalEvents
Set originalhandler = WithHandler
    For Each loopobject In mCol
        
        doeshandle = loopobject.BeforeFunctionCall(withparser, FuncCalling, WithHandler, Argarray, ForceResult)
        If doeshandle Or Not (originalhandler Is WithHandler) Then
            ForceResult = True
            Exit Function
        
        End If
    Next
End Function

Private Function IEvalEvents_GetOperatorFlags(withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As OperatorFlags
    Dim gotflags As OperatorFlags
    For Each loopobject In mCol
        gotflags = loopobject.GetOperatorFlags(withparser, OpCheck, UnaryOp_None)
        If CBool(gotflags) Then
            Exit For
        End If
    
    Next
    IEvalEvents_GetOperatorFlags = gotflags
    
End Function

Private Function IEvalEvents_GetPluginUIData() As IPluginUIData
    Set IEvalEvents_GetPluginUIData = Nothing
        'This Delegates to all the plugins. In order to get to them...
End Function

Private Function IEvalEvents_HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
    Dim Retbool As Boolean
    For Each loopobject In mCol
        Retbool = loopobject.HandleSubScript(withparser, OnValue, subscripts, retval)
        If Retbool Then
            IEvalEvents_HandleSubScript = Retbool
            Exit Function
        End If
    Next
    
End Function

Private Function IEvalEvents_OnInvokeObject(withparser As CParser, ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
    Dim retLoop As Boolean
    For Each loopobject In mCol
        retLoop = loopobject.OnInvokeObject(withparser, ObjInvoke, currItem, Result)
        If retLoop Then
            'if they returned, "True", then there is no need to
            'call the next one in the "Chain".
            IEvalEvents_OnInvokeObject = retLoop
            Exit Function
        End If
    
    
    
    Next loopobject
    IEvalEvents_OnInvokeObject = False
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As CParser, ByVal VarName As String, ByRef cancel As Boolean) As Variant
'
    For Each loopobject In mCol
        On Error Resume Next
        IEvalEvents_BeforeResolveVariable = loopobject.BeforeResolveVariable(ParserUse, VarName, cancel)
        If Err <> 0 Then
            Set IEvalEvents_BeforeResolveVariable = loopobject.BeforeResolveVariable(ParserUse, VarName, cancel)
        End If
        If cancel Then Exit For
    Next
End Function



Private Function IEvalEvents_CanHandleFunction(withparser As CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
 Dim LoopObj As IEvalEvents
 'if more than one Implementor can handle the function, only the first one will be dealt with.
    For Each LoopObj In mCol
        'redimension the no parsed args to erase it.
        ReDim NoParsedArgs(LBound(NoParsedArgs) To UBound(NoParsedArgs))
        If LoopObj.CanHandleFunction(withparser, FuncName, NoParsedArgs) Then
            IEvalEvents_CanHandleFunction = True
            Exit Function
        End If
    Next LoopObj
End Function



Private Function IEvalEvents_CanHandleOperator(withparser As CParser, ByVal OpCheck As String, ByRef Unary As IEvalUnaryOpFlags) As Boolean
   Dim loopitem As IEvalEvents, Ret As Long
        For Each loopitem In mCol
           ' NoParseArgsA = False
            'NoParseArgsB = False
            Ret = loopitem.CanHandleOperator(withparser, OpCheck, Unary)
        
            If Ret Then Exit For
                
        Next loopitem
        IEvalEvents_CanHandleOperator = Ret
End Function


Private Property Get IEvalEvents_Description() As String
IEvalEvents_Description = "BASeParser IEvalEvents integrated Delegator"
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As CParser)
'
For Each loopobject In mCol
    loopobject.ExecuteComplete ParserDone
Next
End Sub



Private Function IEvalEvents_GetFunctionFlags(withparser As CParser, ByVal FuncString As String) As FunctionFlags
   Dim tempresult As Long
        For Each loopobject In mCol
            tempresult = loopobject.GetFunctionFlags(withparser, FuncString)
            
            If tempresult <> Function_Implemented Then
            IEvalEvents_GetFunctionFlags = tempresult
            Exit Function
            End If
            

    Next
    IEvalEvents_GetFunctionFlags = Function_Implemented
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As CParser) As String
    Dim Accum As cStringBuilder
    Set Accum = New cStringBuilder
    For Each loopobject In mCol
        Accum.Append loopobject.GetHandledFunctionString(withparser)
    
    
    Next

    IEvalEvents_GetHandledFunctionString = Accum.ToString


End Function

Private Function IEvalEvents_GetHandledOperatorsString(withparser As CParser) As String
    '
    Dim BuildStr As cStringBuilder
    Set BuildStr = New cStringBuilder
    For Each loopobject In mCol
        BuildStr.Append loopobject.GetHandledOperatorsString(withparser)
        
    
    Next loopobject
    IEvalEvents_GetHandledOperatorsString = BuildStr.ToString
End Function




Private Function IEvalEvents_GetInfoObject() As IOpFuncInfo
    Set IEvalEvents_GetInfoObject = Me
End Function



Private Sub IEvalEvents_GetOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval As Variant)
  Dim getop As IEvalEvents
     'IEvalEvents_GetOperation = CVErr(GPRV_NOTHANDLED)
        For Each getop In mCol
            If getop.CanHandleOperator(withparser, OpUse, False) Then
                 Call getop.GetOperation(withparser, OpUse, OpA, OpB, currItem, retval)
                 Exit For
            End If
        Next getop
    'IEvalEvents_GetOperation = getop


End Sub

Private Sub IEvalEvents_GetUnaryOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As CFormItem, retval As Variant)
 For Each loopobject In mCol
        If loopobject.CanHandleOperator(withparser, OpUse, True) Or loopobject.CanHandleOperator(withparser, OpUse, UnaryOp_Suffix) Then
            Call loopobject.GetUnaryOperation(withparser, OpUse, OpA, currItem, retval)
        
        End If
    Next

End Sub

Private Sub IEvalEvents_HandleFunction(withparser As CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
    Dim retval As Variant
    Dim looper As IEvalEvents, tempparsed() As Boolean
    'retval = CVErr(GPRV_NOTHANDLED)
    On Error Resume Next
    ReDim tempparsed(UBound(FuncParams))
    On Error GoTo 0
    For Each looper In mCol
        If looper.CanHandleFunction(withparser, FuncName, tempparsed) Then
            Call looper.HandleFunction(withparser, FuncName, FuncParams, retval)
            Exit For
        End If
    Next
    If IsObject(retval) Then
        Set ReturnVal = retval
    Else
        ReturnVal = retval
    End If
End Sub

Private Function IEvalEvents_Init(withparser As CParser) As IPluginInitCode
 'mcol.Add New BPCoreOpFunc     'the VERY core object is the interface. hee hee.
Set mParser = withparser
End Function

Private Property Get IEvalEvents_Name() As String
IEvalEvents_Name = "IEvalEventSinks"
End Property

Private Sub IEvalEvents_ParseComplete(ParserDone As CParser)
'
For Each loopobject In mCol
    loopobject.ParseComplete ParserDone
Next
End Sub

Private Sub IEvalEvents_PreExecute(withparser As CParser, FirstItem As CFormItem)
    For Each loopobject In mCol
        loopobject.PreExecute withparser, FirstItem
    Next
End Sub

Private Sub IEvalEvents_PreParse(Parser As CParser, Expression As String)
'

For Each loopobject In mCol
    loopobject.PreParse Parser, Expression
Next
End Sub

Private Property Get IEvalEvents_Self() As Object
Set IEvalEvents_Self = Me
End Property



Private Property Let IEvalEvents_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get IEvalEvents_Settings(withparser As CParser, ByVal PropName As String) As Variant
'
End Property

Private Function IEvalEvents_Term(withparser As CParser) As IPluginTermCode
    Dim termit As IEvalEvents
    If Not mCol Is Nothing Then
        For Each termit In mCol
            termit.Term withparser
        Next
        Set mCol = Nothing
    End If
End Function


Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
    For Each loopobject In mCol
        If loopobject.VariableChanged(VarName, OldValue, VarValue) Then
            Exit For
        End If
    
    Next
End Function


Private Sub IOpFuncInfo_GetFunctionInformation(withparser As CParser, ByVal FuncName As String, FuncInfo As FUNCTIONINFORMATION)
'delegate to first class that handles funcname, and then break.
Dim TempArr() As Boolean, infoObj As IOpFuncInfo
ReDim TempArr(50)
    For Each loopobject In mCol
        If loopobject.CanHandleFunction(withparser, FuncName, TempArr()) Then
            Set infoObj = loopobject.GetInfoObject
                If Not infoObj Is Nothing Then
                    infoObj.GetFunctionInformation withparser, FuncName, FuncInfo
                
                End If
            Exit Sub
        End If
    
    
    Next




End Sub

'Private Property Get IPluginUIData_Description() As String
''
'
'End Property
'
'Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As stdole.Picture
''
'End Function
'
'Private Function IPluginUIData_GetSettingsPageProgID() As String
''
'End Function
