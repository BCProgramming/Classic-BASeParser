VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FunctionHandler:
'BASeParser Plugin That makes use of the CParser Objects "Functions" collection, which is otherwise unused. THis plugin
'makes any function added to that collection usable within expressions.

Implements IEvalEvents
Implements IPluginUIData
Private Type SunkEventsData
    SunkEventName As String
    SunkEventsFunction As CFunction 'function to call
End Type
Private Type EventSinkData
    SunkObject As Object
    SunkFunctionLookupCount As Long
    SunkFunctionLookup() As SunkEventsData
End Type


Private mSinkData() As EventSinkData, mSinkDataCount As Long
Const OurFuncs As String = " FUNCTION REMFUNCTION DLLFUNCTION SINKEVENT "
Private mParser As CParser
Private WithEvents mextendercol As ExtenderCollection
Attribute mextendercol.VB_VarHelpID = -1



Private Function IEvalEvents_AssignSubScript(Withparser As CParser, OnValue As Variant, SubScripts As Variant, AssignThis As Variant) As Boolean
'
End Function
Private Sub IEvalEvents_AfterResolveVariable(ParserUse As CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_BeforeFunctionCall(Withparser As CParser, FuncCalling As String, WithHandler As IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_CanHandleFunction(Withparser As CParser, ByVal funcName As String, NoParsedArgs() As Boolean) As Boolean

'this Plugin illustrates how function handling can be
'determined dynamically, from some external source- in this case,
'the Functions Collection exposed by BASeParser.
Dim TestFunc As CFunction
Dim createfunc As CFunction
'Firstly- see if it is one of OUR functions, that is- a Function-handling function.
'right now, we have two- FUNCTION, and REMFUNCTION
'DEFFUNCTION
If InStr(1, OurFuncs, funcName, vbTextCompare) <> 0 Then
    'yep- we can handle it.
    'I don't really think it is necessary to describe here, but I use Instr tests against strings
    'containing valid function names to inspect for an appropriate function quickly.
    
    Select Case UCase$(funcName)
    Case "FUNCTION"
    'FUNCTION("FUNCNAME","FUNCEXPRESSION")
    'defines a new function, and returns the resulting CFunction object.
    'ad to withparser.functions collection...
    
    'Set createfunc = withparser.Functions.Add(FuncName, FuncName)
        'Set createfunc = withparser.Functions.Add(P(0), P(1))
        
    Case "DLLFUNCTION"
        'DLLFUNCTION("LIBNAME","PROCNAME","REALNAME","PARAMS","RETURNTYPE")
     
    
    
    
    
    
    Case "REMFUNCTION"
    'REMFUNCTION("FUNCNAME")
    
    
    End Select
    
    IEvalEvents_CanHandleFunction = True
    Exit Function

End If
On Error GoTo FUNCNOTTHERE
    Set TestFunc = Withparser.Functions.Item(funcName)
    'No error, cool. The function exists, so we can handle it.
     IEvalEvents_CanHandleFunction = True
     Exit Function
FUNCNOTTHERE:
'that's a shame.



End Function

Private Function IEvalEvents_GetOperatorFlags(Withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As OperatorFlags
'
End Function

Private Function IEvalEvents_GetPluginUIData() As IPluginUIData
    Set IEvalEvents_GetPluginUIData = Me
End Function

Private Function IEvalEvents_HandleSubScript(Withparser As CParser, OnValue As Variant, SubScripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_OnInvokeObject(Withparser As CParser, ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
'nothing- Yet.
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function


Private Function IEvalEvents_CanHandleOperator(Withparser As CParser, ByVal OpCheck As String, ByRef Unary As IEvalUnaryOpFlags) As Boolean
'
'false.- Well, unless I make some kind of freaky function syntax,
'or decide to have functions with operator names.
IEvalEvents_CanHandleOperator = False
End Function



Private Property Get IEvalEvents_Description() As String
'
IEvalEvents_Description = "BASeParser Custom Function handling Plugin."
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As CParser)
'
End Sub


Private Function IEvalEvents_GetFunctionFlags(Withparser As CParser, ByVal FuncString As String) As FunctionFlags
IEvalEvents_GetFunctionFlags = 0
End Function

Private Function IEvalEvents_GetHandledFunctionString(Withparser As CParser) As String
'
Dim StrBuild As String, loopobject As CFunction
For Each loopobject In Withparser.Functions
    StrBuild = StrBuild & " " & loopobject.Name
Next
StrBuild = StrBuild & " "
IEvalEvents_GetHandledFunctionString = OurFuncs & StrBuild


End Function

Private Function IEvalEvents_GetHandledOperatorsString(Withparser As CParser) As String
'
End Function

Private Function IEvalEvents_GetInfoObject() As IOpFuncInfo
'
End Function

Private Sub IEvalEvents_GetOperation(Withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval As Variant)
'
End Sub

Private Sub IEvalEvents_GetUnaryOperation(Withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As CFormItem, retval As Variant)
'
End Sub
Private Sub DoSinkEvent(ObjSink As Object, ByVal EventName As String, useFunction As CFunction)
'''
    'steps: look in our structure, see if this object is already sunk.
    'If so, check if the event is also sunk.
    'if it is, change it to the new function.
    'otherwise, add a new entry.
    'if the object is not sunk, then add it. as well as the function entry.
Dim lookobj As Long, lookevent As Long
For lookobj = 1 To mSinkDataCount
    If mSinkData(lookobj).SunkObject Is ObjSink Then
        'the object has event sinks...
        'search for eventname...
        For lookevent = 1 To mSinkData(lookobj).SunkFunctionLookupCount
            If StrComp(mSinkData(lookobj).SunkFunctionLookup(lookevent).SunkEventName, lookevent, vbTextCompare) = 0 Then
                'same event....
                Set mSinkData(lookobj).SunkFunctionLookup(lookevent).SunkEventsFunction = useFunction
                'assigned!
                Exit Sub
            End If
        
        
        Next
    'if we get here, then the event is not sunk... add it.
    mSinkData(lookobj).SunkFunctionLookupCount = mSinkData(lookobj).SunkFunctionLookupCount + 1
    ReDim Preserve mSinkData(lookobj).SunkFunctionLookup(mSinkData(lookobj).SunkFunctionLookupCount)
    With mSinkData(lookobj).SunkFunctionLookup(mSinkData(lookobj).SunkFunctionLookupCount)
        .SunkEventName = EventName
        Set .SunkEventsFunction = useFunction
        'since the object is already "hooked" no need to mess with that.
        Exit Sub
    End With
    
    End If

Next lookobj
'if we get here, then the object is NOT hooked.

mSinkDataCount = mSinkDataCount + 1
ReDim Preserve mSinkData(1 To mSinkDataCount)
With mSinkData(mSinkDataCount)
    Set .SunkObject = ObjSink
    .SunkFunctionLookupCount = 1
    ReDim .SunkFunctionLookup(1 To 1)
    .SunkFunctionLookup(1).SunkEventName = EventName
    Set .SunkFunctionLookup(1).SunkEventsFunction = useFunction
    Dim newextender As BCControlExtender
    'don't add ObjSink if fiIsControl is true; instead, create our BCControlExtender and use that.
    Set newextender = New BCControlExtender
    newextender.Init ObjSink
    mextendercol.Add newextender

End With




End Sub

Private Sub IEvalEvents_HandleFunction(Withparser As CParser, ByVal funcName As String, FuncParams() As Variant, ReturnVal As Variant)
'
'actually handle the function.
Dim paramcount As Long
Dim FuncGet As CFunction
On Error Resume Next
    paramcount = UBound(FuncParams)
    If paramcount < 0 Or Err <> 0 Then
        paramcount = 0
    '
    End If
'hold on- we should check to see if it is a function handling function.
If InStr(1, OurFuncs, funcName, vbTextCompare) <> 0 Then
    Select Case UCase$(funcName)
        Case "SINKEVENT"
            'syntax:
            'SINKEVENT(Object,Eventname,Function)
            Dim castcfunc As CFunction, castobj As Object
            Set castcfunc = FuncParams(2)
            Set castobj = FuncParams(0)
            DoSinkEvent castobj, FuncParams(1), castcfunc
        
        
        
        Case "FUNCTION", "DLLFUNCTION"
            'Define a new function.
            'Arguments:
              'FUNCTION("FUNCNAME","FUNCEXPRESSION")
              '"DLLFUNCTION("LIBRARY","FUNCNAME",["STDCALL"|"CDECL"],[ReturnType])
              'Oh yeah- the name to call is, right now, the name of the function.
              'IE- no Aliasing yet.
              'Notes On DLLFUNCTION:
              'if calling type is omitted, defaults to STDCALL.
              'if returntype is omitted, defaults to Number.
              '(not yet implemented anyways....)
              
              If UCase$(funcName) = "DLLFUNCTION" Then
                   'DLLFUNCTION("LIBNAME","PROCNAME","REALNAME","PARAMS","RETURNTYPE")
                    Set FuncGet = New CFunction
                    On Error GoTo 0
                    FuncGet.InitAPIFunction FuncParams(0), FuncParams(1), FuncParams(2), FuncParams(3), FuncParams(4)
                    Withparser.Functions.AddDirect FuncGet
              
              
              Else
                
                Set FuncGet = Withparser.Functions.Add(FuncParams(1), CStr(FuncParams(0)))
                
                Set ReturnVal = FuncGet
                
    
            End If
        
        Case "REMFUNCTION"
            'remove a function by name
            Withparser.Functions.Remove FuncParams(0)
            
    'REMFUNCTION("FUNCNAME")
    End Select
    Exit Sub
End If


Set FuncGet = Withparser.Functions.Item(funcName)
Call FuncGet.CallFuncByRef(ReturnVal, FuncParams)
'wooh. that was easy.

End Sub

Private Function IEvalEvents_Init(Withparser As CParser) As IPluginInitCode
'
    Set mParser = Withparser
    CDebug.Post "FunctionHandler Class Init()"
    Set mextendercol = New ExtenderCollection
    InitFunctions Withparser, Withparser.configset
End Function

Private Property Get IEvalEvents_Name() As String
'
IEvalEvents_Name = "BASeParser FunctionHandler Core Plugin"
End Property

Private Sub IEvalEvents_ParseComplete(ParserDone As CParser)
'
End Sub

Private Sub IEvalEvents_PreExecute(Withparser As CParser, FirstItem As CFormItem)
'
End Sub

'FunctionHandler- Implements the IEvalEvents interface
'to expose those functions in a CParser's "Functions" Collection.
'This handler is intrinsic and is loaded before BPCoreOpFunc on the Listener Stack.

Private Sub IEvalEvents_PreParse(Parser As CParser, Expression As String)
'
End Sub

Private Property Get IEvalEvents_Self() As Object
Set IEvalEvents_Self = Me
End Property

Private Property Let IEvalEvents_Settings(Withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get IEvalEvents_Settings(Withparser As CParser, ByVal PropName As String) As Variant
'
End Property

Private Function IEvalEvents_Term(Withparser As CParser) As IPluginTermCode
'
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function

Private Property Get IPluginUIData_Description() As String
IPluginUIData_Description = "BASeParser XP Core FunctionHandler Plugin."
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As stdole.Picture
'
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String
'
End Function
Private Sub InitFunctions(ForParser As CParser, ByVal Setname As String)
    'the NEW InitFunctions Procedure.
    'The Old routine would inspect the ParserSettings.SetKey(setName) & "\" & "Functions\" key for
    'Function Names and Values. I have no idea WHY  I did it that way. The new method is to store File names referencing
    'either XML or INI files, and then essentially importing the routines defined therein.
    'By default, a XML file named "FUNCTIONS.XML" is grabbed from the application directory.
    '(turns out that ISN'T app.path. who knew.)
    Dim FileNames() As String, FileCount As Long, flsuccess As Boolean
    Dim mregistry As cRegistry, ExprGrab As String
    Dim I As Long
    Set mregistry = New cRegistry
    mregistry.ClassKey = HHKEY_CURRENT_USER
    mregistry.SectionKey = ParserSettings.SetKey(Setname) & "Functions\"
    flsuccess = mregistry.EnumerateValues(FileNames(), FileCount)
    'add another XML file reference here to.
    'but only if we don't find a value, "NoDefaultFileCheck"
    If Not mregistry.ValueEx(HHKEY_CURRENT_USER, ParserSettings.SetKey(Setname) & "Functions\Options", "NoDefaultFileCheck", RREG_DWORD, False) Then
        FileCount = FileCount + 1
        ReDim Preserve FileNames(1 To FileCount)
        FileNames(FileCount) = IIf(right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "FUNCTIONS.XML"
    End If
    For I = 1 To UBound(FileNames)
    
        If right$(FileNames(I), 4) = ".XML" Then
            ParseFunctionsXML ForParser, FileNames(I)
        
        
        Else
            CDebug.Post "Current Implementation of InitFunctions in FunctionHandler Class does not support INI files for defining Functions."
        
        End If
    
    Next I
End Sub
Private Function ParseFunctionsXML(ForParser As CParser, ByVal Filename As String) As Boolean
    'XML file DTD is the same as that used for Help information.
    'that is, "<FUNCTION>" tags enclosing "PARAMETER" tags. we look for a "Expression" attribute within the function itself.
    'then, for each function we find that has an Expression attribute, we add a new item.
    'for example:
    '<FUNCTION NAME="CUBEROOT" EXPRESSION="P[1]^(1/3)">
    
    'it is generally recommended to use the P[] array for functions added in this way. Actually,
    'it is the only way to do it from here.
    
    'in any case, I think I'll remove the older parameter list information. Instead, the function can use whatever parameters it wants to.
    '(of course the parameters will still be documented in the XML files)
    Dim DOMInfo As MSXML2.DOMDocument30, CurrNameSpace As String
    Dim ParentNode As MSXML2.IXMLDOMNode
    Dim LoopChild As MSXML2.IXMLDOMNode
    Dim FuncNodes As MSXML2.IXMLDOMNode
    
    Set DOMInfo = New DOMDocument
    If Not DOMInfo.Load(Filename) Then
        ParseFunctionsXML = False
        CDebug.Post "Failed to load XML, " & Filename
        Exit Function
    Else
        CDebug.Post "Successful load of XML in file " & Filename
        
    End If
    'alright, go ahead and parse the bugger.
    'Iterate through- we should find a "FUNCTIONLIST" or whatever it was called.
    'oh, ""
    Set ParentNode = DOMInfo.firstChild
    Do Until ParentNode Is Nothing
        If StrComp(ParentNode.baseName, "FUNCTIONLIST") = 0 Then
            'Alright, we found a Functionlist definition.
            'before we begin the loop, retrieve any defined NameSpace for this list of functions.
            On Error Resume Next
            If ParentNode.Attributes.getNamedItem("NAMESPACE").nodeTypedValue <> "" Then
                'alright- so a namespace was defined. good for it.
                CurrNameSpace = ParentNode.Attributes.getNamedItem("NAMESPACE").nodeTypedValue
                'now iterate through each child. Look for "<FUNCTION> tags....
                Set LoopChild = ParentNode.firstChild
                Do Until LoopChild Is Nothing
                    If StrComp(LoopChild.baseName, "FUNCTION", vbTextCompare) = 0 Then
                        'function tag.
                        'alright, check for an "Expression" attribute.
                        On Error Resume Next
                        If LoopChild.Attributes.getNamedItem("EXPRESSION").nodeTypedValue = "" Then
                            'error or no "Expression" attribute.
                        Else
                            With ForParser.Functions.Add(LoopChild.Attributes.getNamedItem("EXPRESSION").nodeTypedValue, LoopChild.Attributes.getNamedItem("NAME").nodeTypedValue)
                                'debugging stuff....
                                CDebug.Post "Successful creation of Function Name=""" & .Name & """, Expression=""" & .Expression & """."
                                                            
                            End With
                            
                        End If
                
                    End If
                
                    Set LoopChild = LoopChild.nextSibling
                    
                Loop
            
            End If
        
        End If
        Set ParentNode = ParentNode.nextSibling
    Loop
    
    
End Function
'Private Sub InitFunctions(ForParser As CParser, ByVal Setname As String)
'    'Initializes the functions collection of a Cparser Object using those
'    'functions defined in the configuration data for SetName.
'
'
'
'    Dim FuncNames() As String, FuncCount As Long
'    Dim mregistry As cRegistry, ExprGrab As String
'    Set mregistry = New cRegistry
'    Dim I As Long
'    mregistry.SectionKey = ParserSettings.SetKey(Setname) & "\" & "Functions\"
'    'every key in the section has a name equal to the name of a function. the "Expression" value
'    'within that section is the expression to be assigned to that function.
'    mregistry.EnumerateSections FuncNames, FuncCount
'    If FuncCount = 0 Then
'        CDebug.Post "No Functions found in registry for ConfigSet """ & Setname & """"
'        Exit Sub
'    End If
'    For I = 1 To UBound(FuncNames)
'        'funcnames is a key.
'        mregistry.SectionKey = ParserSettings.SetKey(Setname) & "\" & "Functions\" & FuncNames(I)
'
'        'get the value.
'        mregistry.ValueKey = "Expression"
'        ExprGrab = Trim$(mregistry.Value)
'
'        'if ExprGrab is an empty string, then there is no expression. However, if we were to allow a file reference here, though, it would be strange,
'        'since then the file itself is the name of a function. Rather, it will indicate the name of the set of functions.
'        'However, we will discard that information.
'        If ExprGrab = "" Then
'            'OK- lets try FileName and File respectively.
'            mregistry.ValueKey = "FileName"
'            ExprGrab = Trim$(mregistry.Value)
'            If ExprGrab = "" Then
'                mregistry.ValueKey = "File"
'                ExprGrab = Trim$(mregistry.Value)
'
'
'            End If
'                    'now, recheck, see if the previous if block fixed it.
'            If ExprGrab = "" Then
'                CDebug.Post FuncNames(I) & " registry key could not be interpolated as either a expression or a file reference. Entry ignored.", Severity_Warning
'                GoTo CONTINUEBLOCK
'
'
'            Else
'                'Either the File or FileName checks succeeded, so we need to call another function
'                'to determine the type of file and load the functions therein as appropriate.
'
'            End If
'
'
'        End If
'
'        CDebug.Post "added function """ & FuncNames(I) & " whose expression is " & ExprGrab
'
'        ForParser.Functions.Add mregistry.Value, FuncNames(I)
'CONTINUEBLOCK:
'    Next I
'
'
'
'
'
'
'
'End Sub
Public Function LoadFunctionsFromFile(ByVal Filename As String, AddTo As CFunctions, Optional ByVal replaceOnConflict As Boolean = False) As Boolean
    'loads the functions defined in the given file. The file can be:
    'a XML file
    'a Text file
    'the XML format should be:
    '<FUNCTION NAME="NAME" DESCRIPTION="WHAT IT DOES" EXPRESSION="X">
    '<PARAMETER NAME="X" TYPE="NUMBER">
    '</FUNCTION>
    'in a curious twist, the same XML file used to extract the actual Function "Code" can then
    'be used to retrieve the Information via  FunctionInformationFromXML.
    
End Function


Private Function ParametersToVariant(params() As ParameterInfo) As Variant

    Dim RetArray As Variant, I As Long
    ReDim RetArray(0 To UBound(params))
    
    For I = 0 To UBound(params)
        Call Assign(RetArray(I), params(I).ParameterValue)
    
    Next I

End Function
Private Sub mextendercol_ObjectEvent(FromExtender As BCControlExtender, ObjectEvt As Object, pEventInfo As EventInfo)
    Debug.Print "Event raised:" & TypeName(ObjectEvt) & " Event name " & pEventInfo.EventName
    
    'easy... we look in our structure and find the right function object to call.
    Dim LoopSink As Long, loopevents As Long
    Dim funccall As CFunction, paramsuse As Variant
    
    For LoopSink = 1 To mSinkDataCount
        With mSinkData(LoopSink)
            If .SunkObject Is FromExtender Then
                For loopevents = 1 To .SunkFunctionLookupCount
                    With .SunkFunctionLookup(loopevents)
                        If StrComp(.SunkEventName, pEventInfo.EventName, vbTextCompare) = 0 Then
                        
                            'call the given function.
                            'this is messy- we need to pass along the appropriate parameters as well.
                            Set funccall = .SunkEventsFunction
                                If pEventInfo.ParameterCount = 0 Then
                                    funccall.CallFunc
                                Else
                                    paramsuse = ParametersToVariant
                                    funccall.CallFuncArray paramsuse
                                End If
                            
                            
                            
                            
                        
                        End If
            
                    End With
                Next
            End If
    
    
    
    
    
    
End Sub
