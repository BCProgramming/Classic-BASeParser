VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEvalEvents
'FuncDelegate BASeParser XP Plugin
'Delegates function calls to "global" objects that implement those functions.
'this makes it way to easy to implement new functions.
'since the given library doesn't even need to implement the Plugin interface.
'for example, if an object Obj exposed a method "FUNCTEST", and that object had it's progID
'listed as a section under "Software\BASeCamp\BASeParser\Plugins\BPCoreFunc.FuncDelegate\DelegateProgIDs",
'Section has changed:
'Software\BASeCamp\BASeParser\Configuration Data\<Section>\Plugins\BPCoreFunc.FuncDelegate\DelegateProgIDs"

'then the Functest method would be suddenly available via the Parser.
'shoot- now that I think about it, this can expose a whole wackload of functions,
'and may make the core Operators/Functions obsolete. Oh well.
Private Type ObjectInfo
    IsAvailable As Boolean
    ProgID As String
    mObject As Object
    mInterfaceInfo As TLI.InterfaceInfo
End Type
Private Type DelegateData
    NumObjects As Long
    Objects() As ObjectInfo
End Type
Private mkeyLoad As String
Private mObjectCount  As Long
Private mObjData As DelegateData
Private mreguse As cRegistry

Private Sub Class_Initialize()
     Set mreguse = New cRegistry
    
End Sub
Private Function IEvalEvents_AssignSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, assignthis As Variant) As Boolean
'
End Function
Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_BeforeFunctionCall(withparser As BASeParserXP.CParser, FuncCalling As String, WithHandler As BASeParserXP.IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'
Dim handler As Long
handler = GetHandlerForFunction(FuncName)
If handler > -1 Then IEvalEvents_CanHandleFunction = True

End Function

Private Function IEvalEvents_CanHandleOperator(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'
End Function

Private Property Get IEvalEvents_Description() As String
'
IEvalEvents_Description = "BASeParser Core Global Delegator"
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As BASeParserXP.CParser) As String
'
'Oh boy, they asked for it now :)
Dim LoopStruct As Long
Dim LoopMember As MemberInfo
Static Retrieved As Boolean, StrRetrieved As String

If Retrieved Then
    IEvalEvents_GetHandledFunctionString = StrRetrieved
Else

    StrRetrieved = ""
    For LoopStruct = 1 To mObjData.NumObjects
        With mObjData.Objects(LoopStruct)
            'If .IsAvailable Then
                For Each LoopMember In .mInterfaceInfo.Members
         'add to the return.
                StrRetrieved = StrRetrieved & " " & LoopMember.Name & " "
                Next
            'End If
        End With
    Next
    Retrieved = True
    IEvalEvents_GetHandledFunctionString = StrRetrieved
End If
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
''
End Sub

Private Sub IEvalEvents_HandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'




End Sub

Private Function IEvalEvents_HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
'
Debug.Print "FuncDelegate Init()"
mkeyLoad = withparser.Settings.SetKey(withparser.ConfigSet)
If Not Right$(mkeyLoad, 1) = "\" Then
    mkeyLoad = mkeyLoad & "\"
End If
'append our progID.
mkeyLoad = mkeyLoad & "Plugins\BPCoreFunc.FuncDelegate"
'there.
LoadSettings

End Function

Private Property Get IEvalEvents_Name() As String
'
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
Set IEvalEvents_Self = Me
End Property

Private Property Get IEvalEvents_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String) As Variant
'
End Property

Private Property Let IEvalEvents_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Function IEvalEvents_Term(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginTermCode
'
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function


Private Sub LoadSettings()
    'load the settings- stored in mkeyload.
    Dim LoadIDs() As String, IDcount As Long
    Dim I As Long
    With mreguse
        'we will find the data we want within mkeyload.
        'ie: mkeyload\Globals\<progID sections>
        'where each ProgID is a section- I make them sections
        'just in case I add specific configuration data for each one.
'        Private Type ObjectInfo
'            mObject As Object
'            mInterfaceInfo As TLI.InterfaceInfo
'        End Type
'        Private Type DelegateData
'            NumObject As Long
'            Objects As ObjectInfo
'        End Type
        mreguse.Classkey = HHKEY_CURRENT_USER
        mreguse.SectionKey = mkeyLoad & "\Globals\"
        Call mreguse.EnumerateSections(LoadIDs, IDcount)
        'redimension the array as appropriate.
        ReDim mObjData.Objects(1 To IDcount)
        mObjData.NumObjects = IDcount
        For I = 1 To IDcount
            LoadobjectData mObjData.Objects(I), LoadIDs(I)
        
        
        Next I
    
    End With
    
    
    
    
    
End Sub

Private Function GetHandlerForFunction(ByVal FuncName As String) As Long
    'returns the index of the ObjectInfo item in the Array.
    '-1 if no handler.
    Dim LoopI As Long
    Dim testmember As MemberInfo
    For LoopI = 1 To mObjData.NumObjects
        Set testmember = InterfaceFunction(mObjData.Objects(LoopI).mInterfaceInfo, FuncName)
        If Not testmember Is Nothing Then
            GetHandlerForFunction = LoopI
            
        End If
            
    
    Next
    '-1 for failure.
    GetHandlerForFunction = -1
End Function
Private Function InterfaceFunction(ByVal Interface As TLI.InterfaceInfo, ByVal FuncName As String) As MemberInfo
'mObjData.Objects(LoopI).mInterfaceInfo
'determines wether the given "InterfaceInfo" item has the given function (method).
Dim LoopMember As MemberInfo
For Each LoopMember In Interface.Members
'if the names match in a case insensitive comparision, go ahead and return it.

    'loopmember.name
    If StrComp(FuncName, LoopMember.Name, vbTextCompare) = 0 Then
        'return the item.
        Set InterfaceFunction = LoopMember
        Exit Function
    Else
        'No dice :-(.
        
    End If
    
    
    

        
    
    
    

Next


'uh oh-
Set InterfaceFunction = Nothing

End Function
Private Sub LoadobjectData(ByRef ObjData As ObjectInfo, IDLoad As String)
    'load object data for IDload.
    'if the load fails, set all the values as appropriate.
    Dim ObjLoad As Object
    On Error GoTo CreateFailure
    Set ObjLoad = CreateObject(IDLoad)
    'OK- now load interface info.
    Set ObjData.mInterfaceInfo = InterfaceInfoFromObject(ObjLoad)
    'Set ObjData.mInterfaceInfo.VTableInterface = ObjData.mInterfaceInfo
    'all loaded.



Exit Sub
CreateFailure:
    ObjData.IsAvailable = False
    ObjData.ProgID = ""
    Set ObjData.mObject = Nothing
    Set ObjData.mInterfaceInfo = Nothing
    

End Sub
