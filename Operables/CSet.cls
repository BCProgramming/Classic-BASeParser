VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CSet class. represents a Set. In Pascal, sets are created using the square brackets. as such:
'[1,5,2] + [2,3,5] would return the union, or [1,5,2,3] as another set. y original implementation of sets used Arrays. Thus the use if "In"
'This class implements IOperable and sub-classes the use of the In Operator to work with it.
Implements IEvalEvents
Implements ICorePlugin
Implements IOperable
Implements IPluginUIData
'NOTE: I don't need to implement a new IevalEvents plugin since the default execution code
'will discover that this object supports IOperable.

'a "Set" is simply a group of values. For example:
'« »
'«1,2,3,4,5»
'is a set.

'set arithmetic includes Union,Difference,Intersection, etc.

'the Union of two sets will contain elements found in either or both set. (but only one of each)
'the Intersection of two sets will contain elements found only in both sets.
'the difference of two sets will contain items found in either one, but NOT both, sets.



'Some older languages represent sets via constants, making them seem more like some
'strange mutation of enumerated types.

'Also, another somewhat trivial note- the most important thing to do is to assign
'the appropriate value to the Parser Property, giving the Set object a context to work in.
'
Private Const SET_PREFIX = "SetOf"
Private Const IT_SET = -51
Private mSetContents As Variant
Private mvParser As CParser
Private Const SetStart As String = "«"
Private Const SetEnd As String = "»"
'uses the Array_ routines to perform operations on the set. the main purpose of the class
'is to allow the easy manipulation of Sets both within and outside baseparser, in addition to implementing
'special features such as the toString() method of IOperable.

Private Property Let ICorePlugin_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get ICorePlugin_Settings(withparser As CParser, ByVal PropName As String) As Variant
'
End Property

'when I say "outside baseparser" I mean referencing the BASeParser library, just not in the Parsed Strings.



'Set Arithmetic.

'Arithmetic operations:
'+ (Union)
'   all items from both sets. Of course, no more than one each.)
'- (Difference)
'   items  in the first set not in the second set.
'* (intersection)
'   items in both sets.
'\  (exclusion)
'   items in one and only one of the two sets.
Private Function IEvalEvents_AssignSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, AssignThis As Variant) As Boolean
'
End Function
Public Function Union(ByVal Withset As CSet) As CSet

    'perform an Array_Append Operation. Since our current code automatically
    'strips out multiples, no need to repeat it here.
    Dim newset As CSet
    Set newset = New CSet
    'assign the parser to ours.
    Set newset.Parser = Me.Parser
    
    newset.SetElements Array_Append(mSetContents, Withset.GetElements)
    Set Union = newset


End Function
Public Function Difference(Withset As CSet) As CSet
    Dim newset As CSet
    Set newset = New CSet
    Set newset.Parser = Me.Parser
    
    newset.SetElements Array_Difference(mParser, mSetContents, Withset.GetElements)
    Set Difference = newset
    
End Function
Public Function InterSection(Withset As CSet) As CSet
    Dim newset As CSet
    Set newset = New CSet
    Set newset.Parser = Me.Parser
    newset.SetElements Array_Intersection(mParser, mSetContents, Withset.GetElements)
    Set InterSection = newset
End Function
Public Function Exclusion(Withset As CSet) As CSet
    Dim newset As CSet
    Set newset = New CSet
    Set newset.Parser = Me.Parser
    
    newset.SetElements Array_Exclusion(mParser, mSetContents, Withset.GetElements)
    Set Exclusion = newset
End Function
Public Property Set Parser(Vdata As CParser)
    Set mParser = Vdata
End Property
Public Property Get Parser() As CParser
    Set Parser = mParser
End Property

Public Sub SetElements(Elements As Variant)
    
    mSetContents = Elements
    On Error Resume Next
    StripMultiples Elements, mParser
End Sub
Friend Function GetElements() As Variant
    GetElements = mSetContents
End Function
Public Sub SetValue(ParamArray SetContents())
    Dim workaround As Variant
    workaround = SetContents
    
    Call Me.SetElements(workaround)
End Sub
Public Function ToString() As String
   
    ToString = " Set Of «" & ModParser.JoinEx(mParser, mSetContents, ",") & "»"
End Function
Public Function CSet(ParamArray Elements()) As CSet
    Dim retset As CSet
    Dim copied As Variant
    Set retset = New CSet
    copied = Elements
    retset.SetElements copied
    Set CSet = retset


End Function
Private Property Get mParser() As CParser
    If mvParser Is Nothing Then
        Set mvParser = New CParser
        mvParser.Create
    End If
    Set mParser = mvParser
End Property
Private Property Set mParser(Vdata As CParser)
    Set mvParser = Vdata
End Property
Private Sub Class_Initialize()
    'Set mParser = New CParser
    'Call mParser.Create
End Sub

Private Function ICorePlugin_CanHandleItem(withparser As CParser, Itemcheck As CFormItem) As Boolean
'
'ICorePlugin_CanHandleItem = Itemcheck.ItemType = -51
End Function

Private Function ICorePlugin_GetPluginUIData() As IPluginUIData
'not yet...
End Function

Private Sub ICorePlugin_HandleItem(withparser As CParser, ItemHandle As CFormItem, retval As Variant)
'    '
'    '
'    'Debug.Assert False
'    Dim RetSet As CSet, ElementsMake As Variant
'    Dim I As Long, CurrParser As CParser
'    Dim ParserCol As Variant
'    Set RetSet = New CSet
'    'that Assert had been there a VERY long time. I really need to stop forgetting what code I was working on. Oh well.
'    'So what happens here, you ask? Simple- just create and return a Set object, using the value of the ItemHandle Formula Item.
'    ParserCol = ItemHandle.Value
'    If Not IsArray(ParserCol) Then
'        CDebug.Post "CSet Class Could not handle Custom Item as Item Data was not an array Of CParser Objects."
'
'        Exit Sub
'    Else
'        'It is an array.
'        ReDim ElementsMake(0 To UBound(ParserCol) - LBound(ParserCol))
'        For I = LBound(ParserCol) To UBound(ParserCol)
'            ElementsMake(I + LBound(ParserCol)) = ParserCol(I).Execute
'
'        Next I
'    End If
'    RetSet.SetElements ElementsMake
'    Set retval = RetSet
End Sub

Private Function ICorePlugin_Init(withparser As CParser) As IPluginInitCode
'
CDebug.Post "CSet initialized for use as ICorePlugin."
End Function
Private Sub StringArgsToParserArray(StrArgs() As String, Parsers() As CParser, currItem As CFormItem)
Dim I As Long
 On Error Resume Next
    ReDim Parsers(UBound(StrArgs))
    If Err <> 0 Then
        'assume that Setentries is an unallocated array- which means, that there weren't any arguments.
        'as such, we can simply give the formula item a straight-up set.
        Set currItem.Value = New CSet
        currItem.op = SetStart
        
    Else
        'good- it does have items.
        'and we did allocate the array of parsers.
        For I = 0 To UBound(Parsers)
            'allocate...
            Set Parsers(I) = New CParser
            'assign the expression to the corresponding item...
            Parsers(I).Expression = StrArgs(I)
            
        Next I
        'there. now assign to the formula item...
        currItem.Value = Parsers()
        
    End If

End Sub
Private Function ICorePlugin_ParseLocation(withparser As CParser, ByVal StrParse As String, Position As Long, currItem As CFormItem) As ParseLocationReturnConstants

'new Parse code: REALLY simple compared to the old one!
'Rather then expect a "hard coded" Array style list, simply wait until the Item is handled and grab the next item off of the stack! This essentially makes SetOf a Unary Prefix Operator, and we will simply make it handled as such-
'COMPLETELY forego-ing the need for ANY Core Plugin implementation at all :|







'BEGIN OLD PARSE CODE
''
''Sets are available in BASeParser via the use of guillemets.
''"«»"
'Dim setentries() As String
'Dim newset As CSet, listStart As Long
'Dim Parsers() As CParser, I As Long
'Dim AfterPrefix As String
'
'If Mid$(StrParse, Position, 1) = SetStart Then
'    'use the parser's "ParseArguments" routine.
'    setentries = WithParser.ParseArguments(StrParse, ParserSettings.DefaultBracketBegin & SetStart, ParserSettings.DefaultBracketStop & SetEnd, , Position)
'    'now, create a new set.
'    'OK, now we essentially duplicate the code
'    'used for functions- that is, we modify the given formula item, and give the value
'    'a array of CParser objects.
'    '====
'    curritem.ItemType = -51
'    ICorePlugin_ParseLocation = ParseLocation_recognized
'    StringArgsToParserArray setentries, Parsers(), curritem
'
'    '====
'
'ElseIf StrComp(Mid$(StrParse, Position, Len(SET_PREFIX)), SET_PREFIX, vbTextCompare) = 0 Then
'    'if the next Item after the "Set Of" is a "{"...
'    AfterPrefix = Mid$(StrParse, Position + Len(SET_PREFIX))
'    If left$(Trim$(AfterPrefix), 1) = "{" Then
'        curritem.ItemType = -51
'        ICorePlugin_ParseLocation = ParseLocation_recognized
'        listStart = InStr(AfterPrefix, "{")
'        listStart = listStart + 1
'        setentries = WithParser.ParseArguments(AfterPrefix, "{", "}", , listStart)
'
'        StringArgsToParserArray setentries(), Parsers(), curritem
'        curritem.Value = Parsers
'        curritem.ExprPos = Position
'        Position = listStart + Len(SET_PREFIX)
'    End If
'
'
'    'all done...
'    'Set newset = New CSet
'
'End If
'END OLD PARSE CODE
End Function

Private Function ICorePlugin_Term(FromParser As CParser) As IPluginTermCode
'
CDebug.Post "CSet Core plugin terminating."
End Function

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_BeforeFunctionCall(withparser As CParser, FuncCalling As String, WithHandler As IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'
If InStr(1, IEvalEvents_GetHandledFunctionString(withparser), FuncName, vbTextCompare) <> 0 Then
    IEvalEvents_CanHandleFunction = True

End If
End Function

Private Function IEvalEvents_CanHandleOperator(withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As Boolean
'
If InStr(1, IEvalEvents_GetHandledOperatorsString(withparser), OpCheck, vbTextCompare) <> 0 Then
    IEvalEvents_CanHandleOperator = True
End If
End Function

Private Sub IEvalEvents_ExecuteComplete(ParserDone As CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As CParser, ByVal FuncString As String) As FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As CParser) As String
'
IEvalEvents_GetHandledFunctionString = " Set "
End Function

Private Function IEvalEvents_GetHandledOperatorsString(withparser As CParser) As String
IEvalEvents_GetHandledOperatorsString = " SETOF "
End Function

Private Function IEvalEvents_GetInfoObject() As IOpFuncInfo
'
End Function

Private Sub IEvalEvents_GetOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval As Variant)
'
End Sub

Private Function IEvalEvents_GetOperatorFlags(withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As OperatorFlags
'
If StrComp(OpCheck, " SetOf ", vbTextCompare) = 0 Then
    Unary = UnaryOp_Prefix

End If
End Function

Private Function IEvalEvents_GetPluginUIData() As IPluginUIData
'
'no UI data yet....
End Function

Private Sub IEvalEvents_GetUnaryOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As CFormItem, retval As Variant)
'
'yes :)
Dim retset As CSet
Set retset = New CSet
Select Case UCase$(Trim$(OpUse))
    Case "SETOF"
        retset.SetElements OpA
        Set retval = retset
End Select
End Sub

Private Sub IEvalEvents_HandleFunction(withparser As CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'
Dim retThis As CSet
Dim P As Variant
P = FuncParams
Select Case UCase$(FuncName)
    Case "SET"
        Set retThis = New CSet
        retThis.SetElements FuncParams
    
End Select

End Sub

Private Function IEvalEvents_HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
'not yet....
End Function

Private Function IEvalEvents_Init(withparser As CParser) As IPluginInitCode
'
CDebug.Post "CSet IEvalEvents_Init Called.", Severity_Information
End Function

Private Property Get IEvalEvents_Name() As String
'
IEvalEvents_Name = "CSet IEvalEvents Implementation"
End Property

Private Function IEvalEvents_OnInvokeObject(withparser As CParser, ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
'
'no concern here.
End Function

Private Sub IEvalEvents_ParseComplete(ParserDone As CParser)
'
End Sub

Private Sub IEvalEvents_PreExecute(withparser As CParser, FirstItem As CFormItem)
'
End Sub

Private Sub IEvalEvents_PreParse(Parser As CParser, Expression As String)
'
End Sub

Private Property Get IEvalEvents_Self() As Object
'
Set IEvalEvents_Self = Me
End Property

Private Property Let IEvalEvents_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get IEvalEvents_Settings(withparser As CParser, ByVal PropName As String) As Variant
'
End Property

Private Function IEvalEvents_Term(withparser As CParser) As IPluginTermCode
'
    IEvalEvents_Term = TermCode_Success
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function


Private Function IOperable_PerformFunc(withparser As CParser, ByVal FuncName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants
'
'soon!
End Function

Private Sub IOperable_PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, retval As Variant)
'we support relatively few operators.
'oh, and only for Sets.
On Error GoTo setassert
Dim rme As CSet
 Dim otherset As CSet
 If Not TypeOf OtherOperand Is CSet Then
    Err.Raise 6, "CSet", "Cset currently does not support non-Set operands."
 
 End If
 Set otherset = OtherOperand

Select Case Operation
Case "+"    'union
    Set rme = Me.Union(otherset)
Case "-"    'difference
    Set rme = Me.Difference(otherset)
Case "*"    'intersection
    Set rme = Me.InterSection(otherset)
Case "\", "/"   'exclusion.
    Set rme = Me.Exclusion(otherset)

End Select
Set retval = rme
Exit Sub
setassert:
Debug.Assert False
Resume

End Sub

Private Function IOperable_toString(withparser As CParser) As String
'
IOperable_toString = Me.ToString
End Function

Private Property Get IPluginUIData_Description() As String
IPluginUIData_Description = "Core Plugin for Set Parsing."
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As stdole.Picture
'
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String
'
End Function
