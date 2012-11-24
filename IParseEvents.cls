VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEvalEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Fires events from a Cparser.





'What gets me so excited about using this interface is that I can add new functions-
'ANYONE, can add new functions, either specific to their app or globally. For example,
'a function handling plugin could be created that handles DLL function calls with a
'DLLFunctionCall() method. (I better not have to write it- am I trying to ask someone else to?)
'or, a GetRunningTasks() Function that returns the number of tasks.

'Right now, I'm concentrating on functions that are mostly mathematical and scientific, as well
'as intrinsic VB and C++ functions.

'General rules for Parser Plugins:
'When dealing with parameters passed into your handled functions (For example, if you implement a function that multiplies parameter A by C, then divides that result
'by parameter B/A, you can't write this:
'RetVal = (P(0)*P(2))/(P(1)/P(0))
'rather, each of those operations needs to be implemented via a call to the Parser's Special
'"PerformOperation" Function.

'*** well, actually, you don't HAVE to use the performoperation routine, but if you don't
'your plugin will only work with intrinsic types (not CComplex or any other)


'a few general coding guidelines for plugins and their use.

'No plugin is loaded until the CParser Objects "Create()" method is called. When called,
'the parser will load all plugins described in "HKEY_CURRENT_USER\Software\BASeCamp\BASeParser\Configuration Data\<Configset>"
'where "Configset" is the name passed into the Create method. Note that unless explicitly disabled, every Core plugin will be loaded.

'Here's a biggie: No plugin can be loaded more than once. That is, no two plugins in the collection can
'have the same ProgID. This is implemented simply because such a scenario can exist where a single plugin is loaded so many times the
'machine simply runs out of memory and resources and random operations start failing. If a Plugin would have normally been coded in such a fashion, it
'should be designed to perform such decision paths within it's own code.(aggregation and Delegation)


'Note that through COM Interop, it is possible for .NET Classes to be used as IEvalEvents Sinkers. I'll experiment
'with this soon, and possibly add code to detect this type of Plugin. Not sure how I'd optimize it though...

Public Enum IPluginInitCode
    Code_Success = 0    'good. it succeeded.
    Code_InitFailure = 1 'failed to initialize.

End Enum
Public Enum IPluginTermCode
    TermCode_Success = 0    'termination successful.
    TermCode_Failure = 1    'Termination not successful. Plugin NOT removed.
                        'This return value is ignored when the Force parameter is specified.
    TermCode_NotFound       'the Cparser could not remove it because it was not found.
    'Not documented: Code_ForceRemain=-700. This is used by the internal code.
    

End Enum
'changed this enum to a bit field so the handlers can be queried for the set of any three that they
'implement.
Public Enum IEvalUnaryOpFlags
    UnaryOp_None
    UnaryOp_Prefix = 2  'for example- "Not I", "SizeOf "TEST" (-1 for coercion from boolean)
    UnaryOp_Suffix = 4  'for example, n!,++ (for variables)
End Enum
Public Enum OperatorFlags
    Operator_Implemented = 1
    Operator_Optimizable = 2
End Enum
Public Enum FunctionFlags
    'functionflags- use by GetFunctionFlags as a return value.
    'Replaced CanOptimizeFunction.
    Function_Implemented = 1
    Function_Optimizable = 2
    Function_NoParse = 4
    'no other flags yet.

End Enum
Public Enum SettingsErrorsEnum
    Settings_ReadOnly = ParserErrorBase + 2048 + 17
    Settings_Unrecognized = ParserErrorBase + 2048 + 19
End Enum
Public Enum GenericPluginReturnValues
    GPRV_NOTHANDLED = ParserErrorBase + 1
    'GRPV_OPFUNCERROR        'error occured within the called function/operator, such as "divide by zero" or invalid parameters or something.
    
End Enum

Public Sub PreParse(Parser As CParser, ByRef Expression As String)
    'fired immediately when Parsing begins.
    'to define special extensions, change those extensions
    'into something the Parser will understand. Note that
    'this can include any plugin stuff the client added the the CParser too.
    'for example, I might add, in the BPCoreOpFunc Class, the ability
    'to define "Macros" like in C/C++. These macros would then become global, much like variables.
    '(actually, it would probably be best to implement a new Macros collection, like with Variables.
    
End Sub

'Note: these two functions can help time the difference between parsing and execution.
'note that.
Public Sub ParseComplete(ParserDone As CParser)
    'fired when a CParser object completes the parsing portion.
    'used mostly for asynchronous execution, when I finally add it.
    
End Sub
Public Sub ExecuteComplete(ParserDone As CParser)
    'fired when a CParser object completes execution.
    'used mostly for asynchronous execution, when I finally add it.
End Sub
Public Function VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, ByRef VarValue As Variant) As Boolean
    'fired when a variables value is assigned.
    'This could be through some code, another Plugin, or something even more goofy.
    'even this implementor could be responsible, if it calls Variables.Value
    



End Function
Public Function BeforeResolveVariable(ParserUse As CParser, ByVal VarName As String, ByRef cancel As Boolean) As Variant
    'fired before the parser attempts to resolve a variable reference.
    'this can be used to implement Constants.
    'in Cparser, this is fired in CollapseStack, immediately before a IT_VARIABLE type is
    'resolved by a method call to the Cvariables collection.
    'Take note that even though it says it is before it RESOLVES a variable, it is
    'not before resolving that the token is a variable- that is, you can be pretty sure that the given
    'variable name already exists in ParserUse's Variables Collection.
End Function
'TODO:\\ If/when Macro implementation is complete- Before/After Resolve Macro, too.
'macros are defined in the Pre-Parse stage.
'Also, this interface will have WAY too many methods for
'common use that would end up as stubs.
Public Sub AfterResolveVariable(ParserUse As CParser, ByVal VarName As String, ByRef VarValue As Variant)
    'fired After the parser finds out the value of a variable, but before it returns it.
    
    
End Sub


Public Function Init(withparser As CParser) As IPluginInitCode
'
'return Code_InitFailure to signal initialization failure. The Plugin manager will
'take note of this and remove it- this means that Term() will follow shortly- expect it!
End Function
Public Function Term(withparser As CParser) As IPluginTermCode
'terminates from a parser. Returns a term code. Either way, it will
'be terminated irregardless.
End Function
Public Function CanHandleOperator(withparser As CParser, ByVal OpCheck As String, ByRef Unary As IEvalUnaryOpFlags) As Boolean
'
'returns wether it can handle the operator. note that if Unary is true that must be taken into account as well.


End Function
Public Function GetOperatorFlags(withparser As CParser, ByVal OpCheck As String, ByRef Unary As IEvalUnaryOpFlags) As OperatorFlags
    'return the flags for the given operator.
End Function
Public Function GetHandledOperatorsString(withparser As CParser) As String
    'Should return a String of the operators, separated by spaces.
    'during getpriority, these are appended to each other to make the string to test against.
    'also, the string should be in reverse order of importance.
    'Private Const operators As String = " ( ) ^ / \ * + - MOD "
    ' - + * \ / ^ == <>....
    'make a mental note to include the space at the beginning and end of the string-
    'although it might not be necessary, you can never tell
    'if your string was placed at the front or end of the built string to look in.
    
End Function
Public Function GetHandledFunctionString(withparser As CParser) As String
    'return functions that are handled.
    'the larger function names(in length) come first.
    
    'For Example, if you implement, say, a Function "FOO" and a Function "FOOBAR" in your
    'plugin, FOOBAR should come first. Otherwise, you will not be able to use FOOBAR, since the parser
    'detects FOO as a function name. Of course, ideally, it should make sure that the function name is followed by
    'a "(", so this errant behaviour may be fixed now that I realize that it is err- fixable.
    
    





End Function
Public Sub GetUnaryOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As CFormItem, ByRef retval As Variant)
    'perform a Unary operation, OpUse, on OpA.

    'NOTE: there is not need to specify wether the operator is a suffix- because
    'it passes in the current cFormItem, whose Tag will be "SUFFIX" for a suffix-type unary operator.
    

End Sub

Public Sub GetOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, ByRef retval As Variant)
'
    'Perform the operation, OpUse, on the given Operands. If required, The parser also passes in the FormItem object, taking it on faith that you
    'won't be a jerk and modify it. freaky stuff could happen if you modify it- while it probably won't cause an error, it will definitely give
    'unexpected results to the user. I hate that guy.
    
End Sub
Public Function CanHandleFunction(withparser As CParser, ByVal funcName As String, ByRef NoParsedArgs() As Boolean) As Boolean
   '
   'Can this plugin handle the given function, given in FuncName?
   'if so, return True.
   
   
   
   'the ByRef NoParsedArgs Array is predimensioned by the parser to the size of the argument list that was passed.
   'I recommend only setting those arguments you wish parsed to false, and settings to True any other elements.
   'it is imperative that implementors of this function
   'return false for those arguments they want parsed. If so, the expressions will be parsed,
   'and optimized for later Executions.
   
   
End Function
'moved the single method in IEvalEventsEx to this one.
Public Sub PreExecute(withparser As CParser, FirstItem As CFormItem)
'fired immediately before an execute (collapseStack)
'occurs with the parseStack FirstItem.
End Sub
Public Sub HandleFunction(withparser As CParser, ByVal funcName As String, FuncParams() As Variant, ByRef ReturnVal As Variant)
   '
    'handle the given function.
    'mostly self explanatory.
    'However, if the function was called without arguments then FuncParams will not have been
    'initialized and accessing it will cause Subscript out of Range.
    'By default, all parameters are By Reference, and any changes you make to the passed array are propagated back to
    ' the variable used for that parameter, if any. Note that  this doesn't apply for properties- they are always By Value.
End Sub

   
   'the CanOptimizeFunction has been deprecated- use GetFunctionFlags(FuncString) to get bit fields.
    'the only time CanOptimizeFunction should return false is with functions that, given the same set
    'of parameters, won't always return the same result. (the RANDOM function works this way). Otherwise, given that
    'the function has constant arguments, the parser will optimize the call by replacing the
    'IT_FUNCTION formitem with a IT_VALUE formitem, equal to the value returned by the function.
    'a way for the user to tell the parser not to optimize the function is to include any variable within the parameter list.
    'extra parameters are ignored. This won't work for arguments with optional parameters unless all those
    'parameters are also specified.
    'if the implementor does not understand the function that was passed, it should return TRUE.
    'if it returns false, then no more implementors will be checked, and the function will return false.
    

Public Function GetFunctionFlags(withparser As CParser, ByVal FuncString As String) As FunctionFlags
'
End Function
Public Function OnInvokeObject(withparser As CParser, ObjInvoke As Variant, currItem As CFormItem, ByRef Result As Variant) As Boolean
'NEW as of 02-09-2007: ability to "subclass" object method/property invocations. This is for
'if you wish to Define object-like behaviours for, say- numbers.(ObjInvoke will be the expression
'that occured before the @, evaluated.

'Note that this is a plugin-level implementation of the "IDynamicDispatch" Interface.
'the "IDynamicDispatch" interface can be implemented by objects to define their own
'Dynamic Invocations. For example, a Script Object might respond to different method calls depending
'on the scripts contained within it- the imagination is the boundary.

'I have a funny feeling I'll be the only person to create plugins and such for BASeParser.
'I kind of hope it gets popular, but- who knows. It's could just as easily flop completely.




End Function





Public Function GetInfoObject() As IOpFuncInfo
    Set GetInfoObject = Me
End Function
Public Property Get Name() As String
    'returns a name for this. IE- the core is bpCoreOpFunc
End Property
Public Function AssignSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, AssignThis As Variant) As Boolean
    'AssignSubscript- new interface method for assigning items.
    




End Function
Public Function HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, ByRef retval As Variant) As Boolean
    'HandleSubScript.
    'When the Parser encounters subscripts after a value type that it cannot identify,
    'it will call this function on all it's plugins, using the retVal from the first one
    'that returns true, if any. If none return true, an appropriate error will be raised.
    'IE: No Handler installed for Array-Type access to values of type "Typename"
    



End Function
Public Function GetPluginUIData() As IPluginUIData
    'returns a IPluginUIData control or Nothing, if there isn't one.
    'the best way to use this is the implement IPluginUIData in the Plugin itself,
    'and simply return a reference to it here.
End Function
Public Property Get Self() As Object
    'return a reference to the object itself.
    'Set IEvalEvents_Self = Me
End Property


Public Function BeforeFunctionCall(withparser As CParser, FuncCalling As String, ByRef WithHandler As IEvalEvents, Argarray As Variant, ByRef ForceResult As Variant) As Boolean
'Called immediately before a function is being called (via "handleFunction")
'passes the handler that will be used, arguments, function name, etc...


'Return values:
'If True is returned, the value in ForceResult will be used instead of calling that function.
'If False, the call will proceed as normal.
'NOTE: Withhandler can be changed to refer to another Handler. Obviously a bit dangerous if the new handler doesn't support the function...




End Function

'New Property, "Settings"

Public Property Let Settings(withparser As CParser, ByVal PropName As String, ByVal Vdata As Variant)
'
End Property

Public Property Get Settings(withparser As CParser, ByVal PropName As String) As Variant
'
End Property
