VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'HAHA HA!
'a Dynamic CObject class. Allows you to add and remove methods- Of course, it only
'works for BASeParser- I haven't yet figured out how to use the IDispatchEx interface from VB.
'I thought it would be easy- create the typelib and AWAY I GO! but no. Oh well. I'm sure someone else figured
'it out. probably that COM genius, Edanmo. Hopefully before he BECAME A TRAITOR to VB6 :-) ha ha. just kidding.
'but seriously...


'CObject is essentially a base-level implementation for custom objects in BASeParser.
'currently, I cheap out and use CFunction objects for the implementation of the methods and properties.
'Eventually, I'll make them more versatile. or Maybe, somebody else can implement IDynamicDispatch from
'BASeParser XP themselves. Why do <I> have to implement all these interfac... oh yeah- I'm the only person who knows this library
'even exists. OK, you're all off the hook, but be warned, since you're reading this, you know the library exists and
'thus have the ability to make it better. GIVE ME SUGGESTIONS. WHY IS THIS STUFF IN THE COMMENTS
'FOR THE COBJECT CLASS? I DON'T KNOW. WHY AN I YELLING? I DON'T KNOW. :)

'Ideally, I'll hunt down some way to implement IDispatchEx from a VB class.
'Using that class, you could make other classes that implement IDispatchEx and simply delegate to the
'class, which will support members to add Extended Members.
' And then you can actually use those members directly in VB or whatever lang you use!
Implements IDynamicDispatch     'implement
Implements IOperable

'the BASeParser IDispatchEx kludge.

'Events.
'essentially delegated from our function and variable collections, but the
'names are changed to protect the guilty.
Public Event PropertyChanged(Strmember As String, OldValue As Variant, NewValue As Variant)
Attribute PropertyChanged.VB_Description = "Event raised when a property of the object is changed."
Public Event PropertyAdded(PropVar As CVariable)

Public Event MethodAdded(MethodFunc As CFunction)
Public Event MethodRemoved(MethodRemove As CFunction)


'add more Events !

'here is something I find VERY interesting.
'each object will have it's own "private" variables, correct? well, since we don't want
'anyone else to change them, and we kind of want to isolate them as well, we simply use our own
'CParser object! Of course, it needs to instantiated dynamically when first needed, since we want
'to clone everything from the parent other then variables. The neat thing is, we don't need access to
'the variables of our parent anyway- we will be receiving expressions whose variable values have been parsed already and replaced
'with the equivalent values! great!?
Private mParser As CParser
Private WithEvents mVariables As BASeParserXP.CVariables
Attribute mVariables.VB_VarHelpID = -1
Private WithEvents mFunctions As BASeParserXP.CFunctions
Attribute mFunctions.VB_VarHelpID = -1
Private Const DVARIABLE As Integer = 1
Private Const DFUNC As Integer = 0
Private mADORS(DFUNC To DVARIABLE) As Object
'Private WithEvents mFunctions As BASeParserXP.CFunctions
    'TIDBIT: the mparser we use is private, and, conveniently enough, it has private data (variables) and
            'public accessors/methods (CFunction objects), so use those.
            'also, when we enhance the CFunction objects to do something less, err- single-lined,
            'we'll be able to make un-single-lined methods on these objects.

Friend Property Set ParserObj(Vdata As CParser)
    Set mParser = Vdata
End Property
Public Property Get ParserObj() As CParser
    Set ParserObj = mParser
End Property

            'Properties for users of the Cobject.
            
            
            
            
Public Property Get Variables() As CVariables
Attribute Variables.VB_Description = "returns the Public Properties of this object."
    If mVariables Is Nothing Then Set mVariables = New CVariables
    Set Variables = mVariables
End Property
Public Property Get Functions() As CFunctions
Attribute Functions.VB_Description = "Returns the collection of CFunction Objects associated with this object. These are the methods of the object."
    If mFunctions Is Nothing Then Set mFunctions = New CFunctions
    Set Functions = mFunctions
End Property
Private Sub CreateDefaultData(Onparser As CParser)
    With Onparser
        'clear the variables and functions.
        'now they should not be referencing the collections of the cloned parser.
        '.Variables.Clear
        '.Functions.Clear
        Set .Variables = mVariables
        Set .Functions = mFunctions
        
        'one of the default variables is "Self" which refers
        'to this Object.
        .Variables.Add "Self", Me
    End With
End Sub
Public Function InvokeMember(withparser As BASeParserXP.CParser, ByVal Strmember As String, Params As Variant) As Variant
Attribute InvokeMember.VB_Description = "Invokes a member of this object."
    Dim returnValue As Variant
    IDynamicDispatch_InvokeMember withparser, Strmember, Params, False, returnValue
    InvokeMember = returnValue
End Function


'Private Sub Class_GetDataMember(DataMember As String, Data As Object)
''We are a data source, however, we only have ONE- yes, ONE, recordset.
''the one we created in the Initialize Event.
'Select Case DataMember
'
'    Case "Function", "Method", "Methods"
'        Set Data = mADORS(DFUNC)
'    Case Else
'        Set Data = mADORS(DVARIABLE)
'    End Select
'
'End Sub

Private Sub Class_Initialize()
'    On Error GoTo InitFail
'    Set mADORS(0) = createObject("ADODB.RecordSet")
'    Set mADORS(1) = createObject("ADODB.RecordSet")
'
'    With mADORS(0)
'    Call .Fields.Append("Name", adBSTR, 80, adFldKeyColumn)
'    Call .Fields.Append("Expression", adBSTR, 256)
'      .CursorType = adOpenStatic
'      .LOCKTYPE = adLockOptimistic
'      .Open
'    End With
'    With mADORS(1)
'        Call .Fields.Append("Name", adBSTR, 80, adFldKeyColumn)
'        Call .Fields.Append("Value", adBSTR, 256)
'        .CursorType = adOpenStatic
'        .LOCKTYPE = adLockOptimistic
'        .Open
'    End With
'
'    Exit Sub
'InitFail:
'    If Err <> 0 Then
'        Set mADORS(0) = Nothing    'flag for error- means we will ignore DataMember requests.
'        Set mADORS(1) = Nothing
'    End If
End Sub

Private Sub IDynamicDispatch_InvokeMember(withparser As CParser, ByVal Strmember As String, Params As Variant, ByVal PreferProperty As Boolean, returnValue As Variant)
  'step one: if our private parser is not yet instantiated,
    'we need to "Clone()" it from the provided argument.
    'also, if the provided WithParser is different then the parent of
    'our current parser, we need to create a new one, transferring our current variables and
    'methods to the new copy.
    Dim FuncAttempt As CFunction, CacheErr As Long
    Dim VarAttempt As CVariable, paramcount As Long
    Dim DoVariable As Boolean
    Dim CastFunc As CFunction, ForceFunc As Boolean
    On Error Resume Next
    paramcount = (UBound(Params) - LBound(Params)) + 1
    'Error- no parameters where passed.
    If Err <> 0 Then
        paramcount = 0
        Err.Clear
    End If
    On Error GoTo 0
    If mParser Is Nothing Then
        'easiest problem case.
        Set mParser = withparser.Clone()
        'we don't want the variables collection from the parent, so assign the variables and functions
        'collections to two new Variables, which we keep as private variables in this class.
        'I am so sneaky.
        If mVariables Is Nothing Then Set mVariables = New BASeParserXP.CVariables
        If mFunctions Is Nothing Then Set mFunctions = New BASeParserXP.CFunctions
        Set mParser.Variables = mVariables
        Set mParser.Functions = mFunctions
        'also, create the default Functions/Variables.
        CreateDefaultData mParser
    End If
        On Error Resume Next
'        'kludgy sort of.
'        If paramcount > 0 Then
'            If IsObject(Params(0)) Then
'                If TypeOf Params(0) Is CFunction Then
'                    ForceFunc = True
'                End If
'            End If
'        End If
            
            
        Set VarAttempt = mParser.Variables.GetExisting(Strmember)
        Set FuncAttempt = mParser.Functions.FindByName(Strmember)
        'CREATION-------------------------------------------------
        '--------------------------------------------------------------------------------------------------'
        '--------------------------------------------------------------------------------------------------
        
        If VarAttempt Is Nothing And FuncAttempt Is Nothing Then
            'Woop- this means that they MUST be creating it.
            'since it is currently not a method or property.
            'a Function is created by specifying a SINGLE argument,
            'a CFunction Object.
            On Error Resume Next
            If paramcount = 0 Then
                'since there is no property and there is no function, raise the error.
                Err.Raise 438, "CObject::InvokeMember", "No Member Variable/Method defined for " & Strmember & "."
            End If
            Set CastFunc = Params(0)
            If Err = 0 And Not PreferProperty Then
                'good- no error.
                'we are creating a function- unless the flag "PreferProperty" is passed
                'as True, in which case Properties are more important.
                Set returnValue = mParser.Functions.AddDirect(CastFunc)
                'there.
                Exit Sub
                
                
                
                
            Else
                'an error(or preferProperty) we must be making a variable.
                Set VarAttempt = mParser.Variables.Add(Strmember, Params(0))
                Call Assign(returnValue, VarAttempt.Value)
                Exit Sub
            End If
        Else
            'one or both AREN'T nothing- we will end up retrieving ONE of them, at least...
            If VarAttempt Is Nothing Then
                'the func.
                DoVariable = False
            ElseIf FuncAttempt Is Nothing Then
                DoVariable = True
            Else
                'neither one is nothing.
                'As such, we should use the preferproperty:
                DoVariable = PreferProperty
            End If
                'also: in both of these cases, we should account for assignment.
                'a Property, for example, is always assigned if a parameter is passed.
                'a Method, on the other hand, is only assigned if there is only one arg and
                'it is a CFunction object.
                
            If DoVariable Then
                'check for assignment.
                'assignment occurs if a parameter is passed.
                If paramcount > 0 Then
                    'parameter- assign.
                    
                    If IsObject(Params(0)) Then
                        
                        Set VarAttempt.Value = Params(0)
                        
                    Else
                    
                        VarAttempt.Value = Params(0)
                    
                    End If
                    'return the new value.
                    Call Assign(returnValue, VarAttempt.Value)
                Else
                    'No parameter, so retrieve the value.
                    Assign returnValue, VarAttempt.Value
                End If
                
            Else
            'a Function.
            Dim DoAssignFunc As Boolean
                DoAssignFunc = False
                
                If paramcount = 1 Then
                    If IsObject(Params(0)) Then
                        If TypeOf Params(0) Is CFunction Then
                            DoAssignFunc = True
                        End If
                    End If
                End If
                If DoAssignFunc Then
                    'assign to the function the CFunction at Params(0)
                    'remove the current one.
                    mParser.Functions.Remove FuncAttempt
                    Set FuncAttempt = Params(0)
                    mParser.Functions.AddDirect FuncAttempt
                    Set returnValue = FuncAttempt
                Else
                    'invoke the function.
                    'finally- some action.
                    FuncAttempt.CallFuncByRef returnValue, Params
                    
                    
                    
                    
                End If
            
            End If
        
        
        
        End If
       Err.Clear
        
End Sub


Private Function IOperable_PerformFunc(withparser As CParser, ByVal funcName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants

'Not yet implemented. Will use specifically named methods. Such as- PerformFunc.
IOperable_PerformFunc = Perform_Unknown
End Function

Private Sub IOperable_PerformOp(withparser As BASeParserXP.CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As BASeParserXP.IEvalUnaryOpFlags, retval As Variant)
    Err.Raise 5, "CObject::IOperable.PerformOp", "CObject's Cannot be """ & Operation & """." & vbCrLf & _
        "Well, actually, I am to lazy to do anything like that with it-"
    'ha! for long.
    
    'How it is done:
    'CObject's overload operators the same way C++ classes can-
    'via a function named Operator<operation>, for example, Operator+
    'the CObject should take a single parameter- the object to
    'add. that is why it will probably need conditionals....
    
End Sub

Private Function IOperable_toString(withparser As BASeParserXP.CParser) As String
    'HA! the Cobject ToString implementation is much more diverse then the crappy
    'ResultToString Methods object to string algorithm.
    'create something like this:
    Dim StrBuild As cStringBuilder
    Dim LoopMember As CVariable, LoopFunc As CFunction
    Set StrBuild = New cStringBuilder
    On Error Resume Next
    
    'Object {
    'Member <VarName>{
    '   Value:
    '   (use resulttostring)
    '   }
    'etc for variables...
    'Function <FuncName>{
    '   Return <functionexpression>
    '}
    '}
    StrBuild.Append "Object {" & vbCrLf & vbTab
    'loop through each member.
    If mVariables Is Nothing Then
        StrBuild.Append "No Variables."
    Else
        For Each LoopMember In mVariables
            StrBuild.Append vbTab & "Member """ & LoopMember.Name & """{" & vbCrLf
            StrBuild.Append vbTab & vbTab & "Value:" & vbCrLf
            'note the clever replacement to indent the Resulting string.
            'Doesn't work, though, so I don't think "clever" is particularly appropriate...
            
            If Not LoopMember.Value Is Me Then
                StrBuild.Append vbTab & vbTab & Replace$(withparser.ResultToString(LoopMember.Value), vbCrLf, vbCrLf & vbTab & vbTab)
            Else
                StrBuild.Append vbTab & vbTab & "<Self reference>"
            End If
            StrBuild.Append vbCrLf & vbTab & "}" & vbCrLf
        Next
    End If
    StrBuild.Append vbCrLf
    'now, iterate through the functions.
    If mFunctions Is Nothing Then
        StrBuild.Append "No Functions."
    Else
        For Each LoopFunc In mFunctions
            StrBuild.Append vbTab & "Function " & LoopFunc.Name & "{" & vbCrLf
            StrBuild.Append vbTab & vbTab & "Return (" & LoopFunc.Expression & ")" & vbCrLf
            StrBuild.Append vbTab & "}" & vbCrLf & vbCrLf
        Next LoopFunc
    End If
    StrBuild.Append "}"
    
    IOperable_toString = StrBuild.ToString
End Function



Private Sub mFunctions_FunctionAdded(FuncAdded As CFunction)
    'Raise Our corresponding Event. Hey, wait- isn't that what the CFUnction class does?
    RaiseEvent MethodAdded(FuncAdded)
End Sub

Private Sub mFunctions_FunctionRemoved(FuncRemoved As CFunction)
    RaiseEvent MethodRemoved(FuncRemoved)
End Sub

Private Sub mVariables_VarAdded(ByVal VarAdded As CVariable)
'
End Sub

Private Sub mVariables_VarChanged(ByVal VarChanged As BASeParserXP.CVariable, ByVal OldValue As Variant)
    RaiseEvent PropertyChanged(VarChanged.Name, OldValue, VarChanged.Value)
End Sub




