VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEvalEvents
Private Type LoadObjectInfo
    Objectvariable As String        'name of variable for this object.
    ObjectObject As Variant          'err- do the three "Object"'s in there not
                                    'say it all? The object itself.
End Type
Private Type ObjectsLoaded
    NumItems As Long
    Items() As LoadObjectInfo
End Type
Private mreguse As New cRegistry
Private mObjects As ObjectsLoaded
Private mkeyLoad As String
Private Sub AddObjectVariables(WithParser As CParser)
    'add the object variables to the withparser.
    Dim I As Long
    With WithParser.Variables
        For I = 1 To mObjects.NumItems
            .Add mObjects.Items(I).Objectvariable, mObjects.Items(I).ObjectObject
        
        Next I
    
    End With
End Sub
Private Sub LoadobjectData(IntoStruct As LoadObjectInfo, ByVal ProgID As String)
    On Error Resume Next
    With mreguse
        
        Set IntoStruct.ObjectObject = CreateObject(ProgID)
        If Err <> 0 Then
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
        IntoStruct.Objectvariable = .ValueEx(HHKEY_CURRENT_USER, mkeyLoad & ProgID, "VariableName", RREG_SZ, "")
        If IntoStruct.Objectvariable = "" Then
            Err.Raise 5, "GlobalObjects::LoadObjectData", "Error: No variable Name to assign to object type " & ProgID & "."
        End If
    
    End With



End Sub
Private Sub LoadObjects()
    'load the objects from the registry.
    'enumerate mkeyload.
    'sections are the ProgIDs, and inside each the "VariableName" value will be the name of the variable to give it.
    Dim LoadStr() As String, LoadCount As Long
    Dim CurrLoad As Long
    With mreguse
        .ClassKey = HHKEY_CURRENT_USER
        .SectionKey = mkeyLoad
        .EnumerateSections LoadStr, LoadCount
        mObjects.NumItems = LoadCount
        If LoadCount > 0 Then
            ReDim mObjects.Items(1 To LoadCount)
            For CurrLoad = 1 To LoadCount
                LoadobjectData mObjects.Items(CurrLoad), LoadStr(CurrLoad)
            
            Next CurrLoad
        End If
    
    End With
End Sub
Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_AssignSubScript(WithParser As BASeParserXP.CParser, OnValue As Variant, subscripts As Variant, assignthis As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeFunctionCall(WithParser As CParser, FuncCalling As String, WithHandler As BASeParserXP.IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(WithParser As BASeParserXP.CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'
End Function

Private Function IEvalEvents_CanHandleOperator(WithParser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'
End Function

Private Property Get IEvalEvents_Description() As String
'
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(WithParser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(WithParser As BASeParserXP.CParser) As String
'
End Function

Private Function IEvalEvents_GetHandledOperatorsString(WithParser As BASeParserXP.CParser) As String
'
End Function

Private Function IEvalEvents_GetInfoObject() As BASeParserXP.IOpFuncInfo
'
End Function

Private Sub IEvalEvents_GetOperation(WithParser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Function IEvalEvents_GetOperatorFlags(WithParser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As BASeParserXP.OperatorFlags
'
End Function

Private Function IEvalEvents_GetPluginUIData() As BASeParserXP.IPluginUIData
'
End Function

Private Sub IEvalEvents_GetUnaryOperation(WithParser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Sub IEvalEvents_HandleFunction(WithParser As BASeParserXP.CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'
End Sub

Private Function IEvalEvents_HandleSubScript(WithParser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(WithParser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
    mkeyLoad = WithParser.Settings.SetKey(WithParser.ConfigSet)
    If Right$(mkeyLoad, 1) <> "\" Then mkeyLoad = mkeyLoad & "\"
    mkeyLoad = mkeyLoad & "Plugins\BPCoreFunc.GlobalObjects\Variables\"
    LoadObjects
    AddObjectVariables WithParser
    
End Function

Private Property Get IEvalEvents_Name() As String
'
End Property

Private Function IEvalEvents_OnInvokeObject(WithParser As BASeParserXP.CParser, ObjInvoke As Variant, currItem As BASeParserXP.CFormItem, Result As Variant) As Boolean
'
End Function

Private Sub IEvalEvents_ParseComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Sub IEvalEvents_PreExecute(WithParser As BASeParserXP.CParser, FirstItem As BASeParserXP.CFormItem)
'
End Sub

'Parser Plugin- GlobalObjects
'Implements "Global Objects"
Private Sub IEvalEvents_PreParse(Parser As BASeParserXP.CParser, Expression As String)

End Sub

Private Property Get IEvalEvents_Self() As Object
Set IEvalEvents_Self = Me
End Property

Private Property Let IEvalEvents_Settings(WithParser As BASeParserXP.CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get IEvalEvents_Settings(WithParser As BASeParserXP.CParser, ByVal PropName As String) As Variant
'
End Property

Private Function IEvalEvents_Term(WithParser As BASeParserXP.CParser) As BASeParserXP.IPluginTermCode
'
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function
