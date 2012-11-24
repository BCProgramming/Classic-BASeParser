VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : CParser
' DateTime  : 1/5/2007 18:30
' Author    : Michael Burgwin
' Purpose   : Core class of this entire library. used to parse and execute expressions.
'             Raises Events for Parsing and execution, so the client doesn't need to implement the
'entire plugin interface.


'Although I didn't copy any of his code, my previous work on Konrad Rudolph's Expression Evaluation Library
'can be called a learning experience. I didn't lose the code (Although I lost about the last three months if work on it)

'I did use his implementation of Stacks.  I originally wanted to use a Parse Tree. And,- surprisingly, that is exactly what I have now.
'yes, it is strange- the items in the linked list contain other parser objects for parentheses and arguments, which in turn have stacks, and so on.
Implements IOperable
Implements IReturnFormatter

Const MODULE_NAME = "Cparser"
Const MODULE_FILE_NAME = "Cparser.cls"
Const STATEMENTSEPARATOR = ";"   'stolen from C#,which stole it from java which stole it from C++ which had it from C

'BASeParserXP is used by first instantiating a object of this type.

'---------------------------------------------------------------------------------------

Option Explicit

Public Type ParserVersionInformation
    Major As Long
    Minor As Long
    Build As Long
End Type
    

Public Enum FireEventReturnCodeConstants
    FireEvent_NoHandler
    FireEvent_Handled
    

End Enum
'Errors raised during Parsing of an expression. These CAN occur during Execution if a
'subexpression had not yet been Parsed yet.
Public Enum ParseErrors
    Parse_Object_Not_Initialized = ParserErrorBase + 50
    Parse_Badsyntax = ParserErrorBase + 51
    Parse_InvalidTerm = ParserErrorBase + 52
    Parse_MalformedExpression = ParserErrorBase + 53
End Enum
Public Enum ExecuteErrors
    Exec_VariableAlreadyExists = ParserErrorBase + 100
    Exec_UnsupportedOperation = ParserErrorBase + 101
    Exec_UnSupportedFunction = ParserErrorBase + 102
    Exec_GenericExecError = ParserErrorBase + 150       'This simply adds the VB error code
                                                        'and then raises it in the client.
End Enum
Public Enum NumericNotationConstants
    NumNotation_Standard = 0        '50
    NumNotation_Scientific       '5.00E+01
   
    NumNotation_Hexadecimal    '32
    NumNotation_Octal        '62

End Enum
'NOTE: this Type def is not used yet,
'however, when I decide it is worth it, I'll use it to try to get the implementation
'for sub-expression optimization (IE- to remove the parenthesized requirement.)
'Hey,  have the data structure set-up and the prototype code that is DEFfed out...
Public Type SubExprOData
    StackItem As CFormItem
    Position As Long

End Type
Public Enum FunctionHelpType
    FH_QUICK    'just a function definition- SIN would be "Sin(Number)"
                'optional items will be denoted with
                'square brackets.
    FH_EXTENSIVE
        'Extensive- lists all the data relevant.
        'the sin function would be:
        
        'Sin(Number)
        'Returns the Sine of a number.
        'PARAMETERS:
        'NUMBER - The number whose Sine is to be calculated.
        
    

End Enum



'Argument separator.
'this is my first step to actually using constants, instead of
'filling up the EXE with repeated references to equal strings, like "(" and ",".
Const ARGSEP = ","
Const STRLISTOPEN = "{", STRLISTCLOSE = "}"
Private WithEvents UnarySink As DataStack
Attribute UnarySink.VB_VarHelpID = -1
Private mDisabledFunctions As String
'Private mUnaryCache As cFormitem
Private mMultiStatement  As Boolean     'is the given statement Multiple statements?
                    '(if so, we will need to use the "mMultiParsers" collection of CParsers in
                    'sequence, and return ANS.
Private mMultiParsers As Collection 'Of CParser. Hey, that would be a cool
    'syntax feature. I think it is VB .NET, too. sigh. (As Collection Of <Type> is the feature I mean)
    
Private mUnaryCache As New DataStack
Public Enum ParserErrorRecoveryConstants
    PERR_IGNORE = 0     'ignores the error. the result from that particular operation will be
                        'inaccurate.
    PERR_RETURN = 1     'returns the error information as the result.
    PERR_RESUME = 2     'try again.
    PERR_RESUMENEXT = PERR_IGNORE   '
    PERR_FATAL = ParserErrorBase + 600 'Crash, I suppose....

End Enum
Public Enum ParserObjectStateConstants
    ParseState_Idle = 0 'the parser object is idle.
    ParseState_Parsing = 1  'In RipFormula
    ParseState_Parsed = 2   'Idle, but has already parsed the current expression.
    ParseState_Executing = 3    'In CollapseStack
End Enum
'Public Events raised by the Parser.
Private mCorePlugins As ICorePlugins
'Private mCorePlugin As ICorePlugin
Public Event ExecuteComplete(valret As Variant)
Attribute ExecuteComplete.VB_Description = "Fired when the Parser finishes Executing a Parse Stack."
Public Event ParseComplete()
Attribute ParseComplete.VB_Description = "Fired when the ParseStack is built (or re-built)"
Public Event ExpressionChange() 'when expression changes.
Attribute ExpressionChange.VB_Description = "Event Fired when the Expression Changes."
'when an error occurs. "really? I thought is was for whenever the start button was clicked!" you say. :)
'Public Event StartClicked()
'no THIS is fired when the start button is clicked, regardless of how redundant it is in the context of an expression parsing library class.
Public Event Error(ParserError As CParserError, ByRef RecoveryConst As ParserErrorRecoveryConstants)
'When an Error occurs in the parser.

Public Event PluginBeforeLoad(progID As String, ByRef cancel As Boolean)
    'fired before a plugin loads.
    'set cancel, and it won't be loaded.
    
Public Event PluginAfterLoad(PluginObject As IEvalEvents, ByRef cancel As Boolean)   'fired Before a plugin is loaded.
                'if Cancel is set to true when the event returns, then
                'the plugin will immediately be terminated. of course, if it returns
                'the secret code, it won't. This all occurs within a call to the Add() method
                'of the EventSinks collection.

Private WithEvents mVariables As CVariables
Attribute mVariables.VB_VarHelpID = -1
Private WithEvents mFunctions As CFunctions 'our functions- passed on, by reference, to all cloned Parsers.
Attribute mFunctions.VB_VarHelpID = -1



'NOTE: the BPCoreOpFunc cares nothing for these functions- and, the
'FunctionHandler Class is actually Higher in the event stack, and so you can
'add Functions that have the same name as functions that are intrinsic.
'Specifying Intrinsic functions in that case involves simply prepending the name with
'a "BPCoreOpFunc."


Private mFirstItem As CFormItem 'first item of our current formula.
Private mvarExpression As String    'our expression.
Private mAllowInterrupt As Boolean
Private mEvalListener As IEvalEvents
    'don't be fooled. We automatically assign this to a new Instance of
    'IEvalEventSinks and initialize it with the given Set in Create().
Private mParentParser As CParser
    'If this Cparser instance is being used as a helper instance for another
    'Cparser (IE- for brackets), this will be the parent parser.
Private msink As IEvalEventSinks
    'same as the mEvalListener variable, but holds it as the actual type.
Private mSetName As String
    'name of the Configuration set we loaded in Create. Usually "Default"
Private mState As ParserObjectStateConstants
Private mIsConstant As Boolean
Private mReturnFormatter As IReturnFormatter
Private mvarConstantValue As Variant    'if the Expression is constant.
    'if this class contains no references to Variables or dynamic functions,
    'mvarConstantValue will be the already executed value and mIsConstant will be true,
    'flagging Execute to simply return that.
    'the beauty of this architecture is that even the smallest bits will be optimized, as in:
    '(4+N)*(Sqr(-1*2))
    'will not optimize 4+N, but will optimize Sqr(-1*2), preventing a call to the
    'BPCoreOpFunc class every time it is executed. (this optimization is especially useful in
    'large expressions containing functions that are constant but take a long time to execute)
    'however- it may be necessary to turn optimizations off.
    
    
Private UniqueID As String      'GUID used within the library to uniquely identify
'a CParser object.



'UPDATE: OK, the above comment is moot now, since I don't use the guid to determine the top-level
'parser. I use the obvious "mparentParser Is Nothing" test. Much easier.
'Another plus to this method is somebody won't come along and make me look silly by
'pointing it  out. :P


'Error events.
Private mImpliedObject As Object
Private mcancel As Boolean
Private mNotation As NumericNotationConstants
Private mOutputhooks As OutputHooks
Private mInit As Boolean 'Create must be called.
'ImpliedObject: when a Object access operator is found without a preceding object, it is assumed to refer to this object instead.
Public Property Get FirstItem() As CFormItem
    Set FirstItem = mFirstItem
End Property
Public Property Set FirstItem(vdata As CFormItem)
    Set mFirstItem = vdata
End Property
Public Property Set ImpliedObject(vdata As Object)
    Set mImpliedObject = vdata
End Property
Public Property Get ImpliedObject() As Object
    Set ImpliedObject = mImpliedObject
End Property
'Note: These should be changed to As Object for release build.

Public Property Let DisabledFunctions(ByVal vdata As String)
    mDisabledFunctions = UCase$(vdata)
End Property
Public Property Get DisabledFunctions() As String
    DisabledFunctions = mDisabledFunctions
End Property
Public Property Get OutputHooks() As OutputHooks
    Set OutputHooks = mOutputhooks
End Property
Public Function GetNestedLevel() As Long
Attribute GetNestedLevel.VB_Description = "Returns the Nested Level of this parser Object."
'Mr. Owl, how many parents do we have to go through before we get to Nothing?

    If mParentParser Is Nothing Then
        GetNestedLevel = 0
    Else
        GetNestedLevel = 1 + mParentParser.GetNestedLevel
    End If
End Function

Public Function Execute() As Variant
Attribute Execute.VB_Description = "Executes The current Expression. Optionally, include a new expression as a argument."
    Dim ExecRet As Variant
    
    
    
    
    
    ExecuteByRef ExecRet
    If IsObject(ExecRet) Then
        Set Execute = ExecRet
    Else
        Execute = ExecRet
    End If
    'also, add this to the array of last results.
    mcancel = False
    


End Function

Public Function ExecuteByRef(ByRef retval As Variant) As Variant
Attribute ExecuteByRef.VB_Description = "Same as the Execute method, but makes it both faster to run as well as possible to retrieve actual Objects rather then their default values."

    
    If mIsConstant Then
        If IsObject(mvarConstantValue) Then
            Set retval = mvarConstantValue
        Else
            retval = mvarConstantValue
        End If
        RaiseEvent ExecuteComplete(retval)
        Exit Function
    
    End If
    
    If mFirstItem Is Nothing Then
        
        ParseInfix mvarExpression
        'hee hee. This looks REALLY silly.
        'when the parseInfix function detects
        'that the expression is constant, it will
        'automagically execute the statement( Before we know it is constant)
        'and assign the result to mvarConstantValue.
        'from then on- until our expression changes- we return that.
       
        
    End If
    If mvarExpression = "" Then
    
        retval = ""
        Exit Function
    End If
    Dim valret As Variant
    
    
     If mIsConstant Then
            'Fixed a Intermittent bug where variables used in function arguments will cause
            'those arguments to become vbempty. oddly enough, a second execution will return the correct
            'result. This was puzzling at first. As I thought about it (for a minute or two)
            'I realized that this helps me- since the stack was built the first time
            'and executed proper afterwards. As Such, I turned to my optimization code. Turned out
            'that if it was the first Execute (mFirstItem is nothing) And it isn't constant,
            'we were simply exiting the function.
            '(this mIsConstant Block was previously within the "mfirstitem is nothing" block).
            
            'If InStr(Me.Expression, "SEQ") <> 0 Then Stop
            Assign retval, mvarConstantValue
            'retval = mvarConstantValue
            Exit Function
            
        End If
        
        
    Call CollapseStack(mFirstItem, valret)
    RaiseEvent ExecuteComplete(valret)
    Assign mvarConstantValue, valret
    mEvalListener.ExecuteComplete Me
    
  
    If IsObject(valret) Then
        Set retval = valret
        Set mvarConstantValue = valret
    Else
        retval = valret
        mvarConstantValue = valret
    End If
End Function

Public Function Clone() As CParser
Attribute Clone.VB_Description = "Clones this Cparser object. Well, actually, it makes a Child parser."
    'copy everything, except the expression.
    Dim newparser As CParser
    Set newparser = New CParser
    With newparser
        Set .Variables = mVariables
        Set .EvalListener = mEvalListener
        Set .Functions = mFunctions
        Set .ReturnFormatter = mReturnFormatter
        Set .CorePlugins = mCorePlugins
        Set .ImpliedObject = mImpliedObject
        .Notation = mNotation
        
                    'with the above assignment, there may appear to be a bug. What bug?
                    'Well, what if this parser wasn't given a formatter? Then our
                    'mReturnFormatter will actually be US.
                    'but, that is what we want. otherwise, the child parsers
                    'will format values as they please.
                    'If we are a child parser, then the returnformatter will be a reference
                    'to the parentParser.
        .SetChild Me
        
    End With
    Set Clone = newparser
End Function

Public Function ParseArguments(ByVal FromString As String, Optional ByVal BracketStart As String = DefaultBracketStart, Optional ByVal Bracketend As String = DefaultBracketEnd, Optional ByVal ARGUMENTSEP As String = ARGSEP, Optional ByRef startpos As Long = 1) As String()
Attribute ParseArguments.VB_Description = "Parses a Set of Arguments out of the specified string. ignores quotes and other defined special constructs, such as nested parentheses."

    'NOTE: this routine REALLY needs to be rewritten!
    'full of hacks and kludges...

    'ParseArguments Function:
    'assumes the given string is a parameter lists.
    'keeping track of parentheses and quotes, it finds the next ARGSEP that is not part of some other token.
    
    'Added May 06 2008 10:50 PM:
    
    'recognize the use of a "To" keyword between two arguments.
    
    Dim Currpos As Long
    Dim SplMake() As String
    Dim CurrArgument As Long, Char As String
    Dim inQuote As Boolean, intick As Boolean, bracketcount As Long
    Dim ArgStart As Long
    'If there isn't even a separator in the string, at all, return the string as the single argument
    If InStr(FromString, ARGUMENTSEP) = 0 Then
        ReDim SplMake(0)
        SplMake(0) = FromString
        ParseArguments = SplMake
    Exit Function
    
    End If
    
    
    If Len(FromString) <= 0 Then
        Erase SplMake()
       
        ParseArguments = SplMake
        Exit Function
    End If
    
    
    
    CurrArgument = 0
    Currpos = startpos
    ArgStart = Currpos
    Do
        Char = Mid$(FromString, Currpos, 1)
        
        Select Case True
        Case Char = "`"
            intick = Not intick
        Case Char = """"
            inQuote = Not inQuote
    
        Case (inQuote Or intick) And Currpos < Len(FromString)
            'IGNORE!>
            'ignore as long as we are in a string.
            
        Case InStr(BracketStart, Char) <> 0 And Char <> ""
            bracketcount = bracketcount + 1
        Case InStr(Bracketend, Char) <> 0 And Char <> ""
        
            bracketcount = bracketcount - 1
               If bracketcount < 0 Then
                   'Sigh.
                   bracketcount = 0
                    
               End If
               
         
        Case InStr(Currpos, FromString, ARGUMENTSEP) = 0 And bracketcount = 0 And Trim$(Mid$(FromString, Currpos)) <> ""
'            Stop
            ReDim Preserve SplMake(CurrArgument)
            SplMake(CurrArgument) = Mid$(FromString, ArgStart)
            'Possible bugfix: was using "currpos" rather then "Argstart" here. caused some confusion with argument parsing.
            Exit Do
        Case Char = ARGUMENTSEP, (Currpos >= Len(FromString))
            'Stop
            'If ((Not Inquote) Or _
             (CurrPos >= Len(FromString))) And BracketCount = 0 Then
             If ((Not inQuote) And bracketcount = 0) Or (Currpos >= Len(FromString)) Then
                'wow, that is very strange, heh?
                'This will only be true when inquote=false (0) and bracketcount=0 as well.
                ReDim Preserve SplMake(CurrArgument)
                SplMake(CurrArgument) = Mid$(FromString, ArgStart, Abs(Currpos - ArgStart))
                If right$(SplMake(CurrArgument), 1) = ARGUMENTSEP Then
                    SplMake(CurrArgument) = Mid$(SplMake(CurrArgument), 1, Len(SplMake(CurrArgument)) - 1)
                End If
                CurrArgument = CurrArgument + 1
                ArgStart = Currpos + 1
                
                If Currpos >= Len(FromString) Then
'                    If InStr(1, Bracketend, right$(SplMake(CurrArgument - 1), 1), vbTextCompare) <> 0 And BracketCount <= 0 Then
'                        SplMake(CurrArgument - 1) = Mid$(SplMake(CurrArgument - 1), 1, Len(SplMake(CurrArgument - 1)) - 1)
'                    End If
                    
                
                    Exit Do
                    
                End If
            End If
        
        End Select
    
        Currpos = Currpos + 1
    Loop
    If UBound(SplMake) = 0 Then
        If SplMake(0) = "" Then
            'there weren't any arguments- return an actual "Empty" array to denote this to the caller.
            Erase SplMake
        End If
    End If
    startpos = Currpos
    ParseArguments = SplMake
End Function

Public Function ParseValue(ByVal ExprParse As String, ByVal ParseStart As Long, Optional ByVal IsObjAccess As Boolean = False) As String
Attribute ParseValue.VB_Description = "Parses out a token at the given position. This token may be a function or number."
    'starting at Pos, parses out the value that starts there. for example:
    '43 + 52 / 71 * 400 / "String"
    'if the parse were started at the " in String, String would be returned.
    'the value returned should run on until the end of the token.
    'this is delimited by " " or ")"
    
    'UPDATE:
    'support Date literals enclosed in "#"
    'will ignore such literals if the interior date is not valid.
    '(that way there can be operators that use #)
    
    'UPDATE:
    'to support the use of the period in specifying which handler should perform a function, I have adopted the period.
    'for  this reason, the period must be checked to see wether it is actually a part of a number (a number is either before or after it)
    
    'UPDATE:
    'support "Smart quote" style quotes. in Courier, these are “”.
    
    'UPDATE:
    'integrate and clean-up implementations of Hex and other number systems
    'for example, unless preceded with a system identifier that allows it, do not allow alphabetic characters
    'within the value. (well, unless it is a Variable or string constant, I am referring to
    'numbers here)
    Const SmartQuoteStart As String = "“"
    Const SmartQuoteEnd As String = "”"
    
    'NOTE: add recognition of Modulename.Function specification. right now something like
    'BPCoreOpFunc.Sin is recognized improperly and the In part is parsed out.
    'sigh. The solution is pretty easy  though- just check for the module name followed by a period, then find the
    'next ( after that.
    CDebug.PushP "ParseValue", MODULE_NAME, MODULE_FILE_NAME, 0, "Parse Value at " & ParseStart & " in " & ExprParse & "."
    
    Dim inQuote As Boolean, bracketcount As Long
    Dim Char As String, I As Long
    Dim StrRunner As String, NumFound As Boolean
    Dim TokenEnded As Long
    Dim FoundStart As Boolean
   Dim IsNumber As Boolean
    
    If IsObjAccess Then
        'simple. Object Accesses are required to have the function/method/property
        'end, for non-argument versions, with at least a (), No exceptions.
        'the normal parsing routines would extract any embedded operators from it, which is, to say the least, not
        'desirable.
        ParseValue = Mid$(ExprParse, ParseStart, InStr(ParseStart, ExprParse, "(", vbTextCompare) - ParseStart + 1)
        CDebug.PopP "ParseValue"
        Exit Function
    
    End If
    
    
    
    If HasFunctionAtPos(ExprParse, ParseStart) <> "" Then
        ParseValue = HasFunctionAtPos(ExprParse, ParseStart)
        CDebug.PopP "ParseValue"
        Exit Function
    ElseIf hasVariableAtPos(ExprParse, ParseStart) <> "" Then
        
        ParseValue = hasVariableAtPos(ExprParse, ParseStart)
        CDebug.PopP "ParseValue"
        Exit Function
    End If
                'a func, we want to avoid parsing any operators within this function
                'name and thus making parsevalue come back inproper.
                I = ParseStart
    Do
        Dim ValPos As Long, StrExtract As String, DateCheck As Date
        
        
        If I = ParseStart Then
      
            'determine wether the first character is numeric. If so, we can safely assume
            'that the term WILL end when the value stops being a number or period.
            If IsNumeric(Mid$(ExprParse, I, 1)) Then
                IsNumber = True
            End If
            
            
         
        
        End If
        Char = Mid$(ExprParse, I, 1)
        Select Case True
            Case IsNumber
                If InStr("1234567890.", Char) <> 0 Then
                    StrRunner = StrRunner & Char
                Else
                    'end of number.
                    TokenEnded = I
                    Exit Do
                
                End If
            
            Case Char = "#" And Not inQuote
                'possible date literal.
                'check-
                'grab the next "#" and use the interior string in a call to DateValue()
                ValPos = InStr(I + 1, ExprParse, "#")
                StrExtract = Mid$(ExprParse, I + 1, ValPos - I - 1)
                'replace any "#" we somehow picked up.
                'strextract = replaceString(strextract, "#")
                On Error Resume Next
                DateCheck = DateValue(StrExtract)
                
                If Err <> 0 Then
                    'an error. Sigh, not a date.
                    'hmm, but they wanted to try to insert a date, so-
                    ParseValue = ""
                Else
                    'Modification:--6/23/2007 @ 02:22--
                    'it could just be a time. However, a Time will have been properly parsed by DateValue, but the return from that function will not have the time.
                    'The solution is to check to see if TimeValue returns something larger then the TimeValue of the attempted Date parse.
                    'not here though- in the execution.
                    'If TimeValue(strextract) > TimeValue(DateCheck) Then
                    
                    
                    'Else
                
                    'it's a date.
                    'return the whole chunk with the literals.
                    ParseValue = Mid$(ExprParse, I, ValPos)
                    CDebug.PopP "ParseValue"
                    Exit Function
                End If
            Case Char = "(" And Not inQuote
                FoundStart = True
                bracketcount = bracketcount + 1
                TokenEnded = I + 1
                Exit Do
            Case Char = ")" And Not inQuote
                FoundStart = True
                bracketcount = bracketcount - 1
                TokenEnded = I
                Exit Do
                'If the bracketcount is 0, we have reached the end of the token.No.
            Case Char = """"
                FoundStart = True
                
                inQuote = Not inQuote
                If Not inQuote And bracketcount = 0 Then
                    'it has just become false.
                    'The string has ended.
                    'First- check to see if there is another quote immediately after this one.
                    On Error Resume Next
                    If Mid$(ExprParse, I + 1, 1) = """" And Err = 0 Then
                        'No change-
                        StrRunner = StrRunner & Char
                        'we want to skip the extra ""-
                        I = I + 1
                        inQuote = True
                    Else
                    
                    TokenEnded = I + 1
                    Exit Do
                    'add one, we want to include the quote symbol.
                    End If
                    
                End If
            
                
        
            Case inQuote
                'if we are in quotes, append this char to the running value.
                StrRunner = StrRunner & Char
'            Case InStr("1234567890.", char) <> 0
'                StrRunner = StrRunner & char
'                NumFound = True
'                FoundStart = True
            
      
            ' removed condition from following case.
            
            Case Char = " " And FoundStart And bracketcount = 0, GetOperator(ExprParse, I) <> "" And bracketcount = 0, Char = "@" And bracketcount = 0
                'here we have it, the end of the value.
                'don't include the operator or space.
                TokenEnded = I
                Exit Do
            'TODO:\\ this is where a certain thing I'd like to call a "ValueParse Plugin" will be implemented.
            'in all fairness however it could be done with ICorePlugin so I guess I'm just getting a bit ambitious.
            'Case Mid$(ExprParse, I, 2) = "&H"
            Case Else
                'add to the strrunner anyway, for Variable names.
                'check for a period...

                StrRunner = StrRunner & Char
        End Select
            
        
        
        
        I = I + 1
        
    
    Loop Until I > Len(ExprParse)
    
    Dim parsedvalue As String
    
    If TokenEnded = 0 Then
    'TokenEnded = Len(ExprParse)
    parsedvalue = Mid$(ExprParse, ParseStart)
    Else
    parsedvalue = Mid$(ExprParse, ParseStart, TokenEnded - ParseStart)
    End If
    'If Left$(parsedvalue, 1) = """" Then
        'assume it was a string, and replace all "" with, "-
        'parsedvalue = Replace$(parsedvalue, """""", """")
    If left$(parsedvalue, 1) = """" Then
    'here we have a dilemma.
    'we want to support C-type escapes, such as \n in the string. This is just a test....
'    Escape parsedvalue
    
    
    
    
    End If
    
    'End If
    ParseValue = parsedvalue
    CDebug.PopP "ParseValue"
End Function

Public Function ResultToString(Result As Variant) As String
Attribute ResultToString.VB_Description = "Converts an already executed and retrieved value to a string. Also works on user-created values to create user-friendly display."
    'Converts the result returned from Execute into a String.
    'This means that arrays are joined, objects ObjPtr is retrieved, etc.
    Dim BuildStr As String, I As Long
    Dim test As Long, strbuffer As String
    Dim Handled As Boolean
    'before we do ANYTHING- call our IReturnFormatter implementor, and see if they handle it (via the second argument...)
    BuildStr = mReturnFormatter.ResultToString(Result, Handled)
    
    
    If Handled Then
        CDebug.Post "ResultToString Aborting: ReturnFormatter implementation handler used.", Severity_Information
    End If
    
    On Error Resume Next
    BuildStr = Result
    Err.Clear
    'is it an error?
    If IsError(Result) Then
        ResultToString = "!" & CStr(Result) & "!"
        Exit Function
    End If
    
    'can return string and be object- (default property)
    If VarType(Result) = vbString And Not IsObject(Result) Then
        'put quotes around it to make sure it is clear that
        'the value is a string. Oh, and to make it parsable again,
        'as my own rules state :).
        BuildStr = """" & Replace$(Result, """""", """") & """"
    
    
    ElseIf IsNull(Result) Then
        'convert Null to "Null".
        BuildStr = "Null"
        Exit Function
    
    
    ElseIf IsObject(Result) Then
        'it is an object.
        'since we are clueless as to how to work with the object, we will
        'simply display its type and address. If anybody wants to make it more flexible,
        'they can Create an IReturnFormatter implementation that takes their particular classes into account.
        Dim castoperable As IOperable
        'OK, so I lied. there are a few object types that we recognize.
        'Well, one, really- IOperable.
        If TypeOf Result Is IOperable Then
            'case to interface.
            Set castoperable = Result
            On Error Resume Next
            BuildStr = castoperable.ToString(Me)
        
        
        ElseIf InStr(1, TypeName(Result), "MatchCollection", vbTextCompare) <> 0 Then
            Dim LoopMatch As Object
            For Each LoopMatch In Result
                BuildStr = BuildStr & ResultToString(LoopMatch) & vbCrLf
            Next
        ElseIf InStr(1, TypeName(Result), "Match", vbTextCompare) <> 0 Then
        BuildStr = "Match Position=" & Result.FirstIndex & " Value=" & Result.Value
       ' Stop
        
        ElseIf StrComp(TypeName(Result), "Match", vbTextCompare) = 0 Then
            Set LoopMatch = Result
            BuildStr = "$match at " & LoopMatch.FirstIndex & ", Length," & LoopMatch.Length + 1 & " Value=""" & LoopMatch.Value & """.$"
       
        
        Else
        
            'does it have a "NewEnum" property?
            
        
        
         BuildStr = TypeName(Result) & "@" & ObjPtr(Result)
        End If

    
    ElseIf IsNumeric(Result) Then
        'numeric.
        
        If mNotation = NumNotation_Standard Then
           BuildStr = Result
        ElseIf mNotation = NumNotation_Scientific Then
                    BuildStr = Format$(Result, "Scientific")
                
                
        ElseIf Result = Int(Result) Then    'make sure it is whole.
            Err.Clear
            test = CLng(Result)
            If Err = 0 Then
            'Hexadecimal and octal are valid.
                If mNotation = NumNotation_Hexadecimal Then
                    BuildStr = "&H" & Hex$(Result)
                ElseIf mNotation = NumNotation_Octal Then
                     BuildStr = "&O" & Oct$(Result)
                End If
            
            Else
                'it doesn't fit in a long.
                'only scientific.
                
                
            
            
            End If
        End If
        
    
    
    
           
    
    ElseIf IsArray(Result) Then
        'use the "{" to delimit, so the resulting string is compatible to be parsed
        'again.
        Dim BuilderStr As cStringBuilder
        Set BuilderStr = New cStringBuilder
        
        BuilderStr.Append "{"
        On Error Resume Next
            For I = LBound(Result) To UBound(Result)
            'recursively convert each element to a string as well.
                If Err.Number <> 0 Then Exit For
                'BuildStr = BuildStr & ResultToString(Result(I))
                BuilderStr.Append ResultToString(Result(I))
                If I <> UBound(Result) Then BuilderStr.Append ARGSEP
            
            Next I
            BuilderStr.Append "}"
       BuildStr = BuilderStr.ToString
        
    
    End If
    
    
    
    
    ResultToString = BuildStr
End Function

Public Function IsConstant() As Boolean
Attribute IsConstant.VB_Description = "Returns wether the Expression contained within this Parser is Constant (No variables, no dynamic functions, etc.)"
    'determines if this parser is only filled with static data. For example, 5-4 or what-not.
    Dim FirstItem As CFormItem, pCurrItem As CFormItem
    Dim Retbool As Boolean, looper As Long
    'parse, if necessary.
      If ParserSettings.Optimize = False Then
        CDebug.Post "Optimizations Disabled."
        IsConstant = False
        Exit Function
    End If
    If mFirstItem Is Nothing Then ParseInfix Me.Expression
    Set FirstItem = mFirstItem
    'there.
    'idea:
    'add ability to customize this function, via both implemented procedures (such as, "IsFunctionConstant" or "IsOperatorConstant")
    'as well as registry settings, such as "AllowOptimize" and what-not.
    Retbool = True
    Set pCurrItem = FirstItem
    'first- check our ParserSettings to see if
    'optimization is disabled.

    Do Until pCurrItem Is Nothing
        'IT_VALUE types are always constant.
       
           Set pCurrItem = pCurrItem.Next_
            If IsFormItemConstant(pCurrItem) = False Then
                IsConstant = False
                Exit Function
            End If
    Loop

    IsConstant = Retbool

End Function

Public Function EvalAsString(StrString As String) As String
Attribute EvalAsString.VB_Description = "Evaluates an Expression, returning a String converted from the result."
    Me.Expression = StrString
    EvalAsString = ResultAsString
End Function

Public Function hasVariableAtPos(ByVal StrCheck As String, ByVal PosCheck As Long) As String
Attribute hasVariableAtPos.VB_Description = "Determines wether the given expression has a variable  at the given position and returns it if true."
    Dim VarNames() As String, I As Long
    Dim CurrMatch As String     'current match.
    Dim ParsedVal As String
    Dim retval As String
    VarNames = Me.Variables.GetNames
    If Me.Variables.Count = 0 Then Exit Function
    'Note: we need to find the LONGEST matching variable name.
    'sadly, this means we need to go through the entire array.
    'sigh.
    
    'actually- now that I think about it, this is a pretty simple function, and we don't require
    'the loop!!:
    'algorithm:
    'retrieve the parsevalue of the given location.
'
'    ParsedVal = ParseValue(StrCheck, PosCheck, False)
'    ParsedVal = Trim$(ParsedVal)
'    'if the returned parsevalue is empty, return ""
'    If ParsedVal = "" Then
'        retval = ""
'
'    'if the returned parsevalue ends in "(", return ""
'    ElseIf Right$(ParsedVal, 1) = "(" Then
'        retval = ""
'    'if the returned parsevalue starts with a quote, return ""
'    ElseIf Right$(ParsedVal, 1) = """" Then
'        retval = ""
'    ElseIf Val(ParsedVal) <> 0 Then
'        retval = ""
'    Else
'        'OK! we're good!
'        retval = ParsedVal
'    End If
    'if the returned value has a "VAL" that is NOT 0, return ""
    'if none of the above tests return, return the string.
    
    
    
    
    For I = 0 To UBound(VarNames)
        'If VarNames(I) = "X" Then Stop
        If VarNames(I) <> "" Then
            If StrComp(VarNames(I), _
                Mid$(StrCheck, PosCheck, Len(VarNames(I))), vbTextCompare) = 0 Then
                    'ha. a variable name.
                    'if it is longer then our current entry, then
                    'it is eligible.
                    If Len(VarNames(I)) > Len(CurrMatch) Then
                        CurrMatch = VarNames(I)
                    End If





            End If
        End If
    Next I



    hasVariableAtPos = CurrMatch
'hasVariableAtPos = retval
End Function

Public Function HasFunctionAtPos(ByVal StrCheck As String, ByVal PosCheck As Long) As String
Attribute HasFunctionAtPos.VB_Description = "Determines wether the given expression has a function at the given position and returns it if true."
    Dim SplitStrings() As String
    Dim FullString As String, I As Long
    Dim ModulePrefix As String
    Dim testMod As String
    Dim PlugNames() As String
    Dim LoopTest As Boolean
    Dim lRunner As Long
    Dim NameRunner As String
    CDebug.PushP "HasFunctionAtPos", "CParser", "CParser.cls", ObjPtr(Me)
    FullString = mEvalListener.GetHandledFunctionString(Me)
    CDebug.Post "Handled Function string=" & FullString
    SplitStrings = Split(FullString, " ")
    'for the case of modulename.function notation, we don't verify that the function exists or anything-
    'in fact, to be honest, I don't care if it exists. This makes it possible for a routine to exist in
    'a plugin that cannot be accessed Without using the dot notation. Although Right now that seems dumb.
    'firstly, check for "module" names.
    PlugNames = GetLoadedPluginNames
    CDebug.Post "starting to loop through plugins..."
    
    '//PRB
    'Currently crashes here in compiled apps.
    
    For I = 1 To UBound(PlugNames)
    PlugNames(I) = Trim$(PlugNames(I))
    CDebug.Post "examining plugin, " & PlugNames(I)
    If InStr(PlugNames(I), " ") = 0 Then
        
        
            'if there is a space in the name of the plugin, we ignore all requests
            'via the dot-notation. Although it is possible to make it work, the spaces
            'really contort the situation, and would cause grievious errors abound later on down the
            'road I'm sure. For now, then, disallow such attempts.
            
        If StrComp(Mid$(StrCheck, PosCheck, Len(PlugNames(I))), PlugNames(I), vbTextCompare) = 0 Then
        'hey, guess what we Found!
        'determine wether this plugin name is followed by a period. Or a @, I suppose.
        NameRunner = PlugNames(I) & "."
        Select Case Mid$(StrCheck, PosCheck + Len(PlugNames(I)), 1)
            Case ".", "@"
            'alright- so it does access a function name.
            lRunner = 0
            Do Until LoopTest
                Select Case Mid$(StrCheck, PosCheck + Len(PlugNames(I)) + 1 + lRunner, 1)
                    Case "(", " "
                    
                        LoopTest = True
                    
                    Case Else
                        NameRunner = NameRunner + Mid$(StrCheck, PosCheck + Len(PlugNames(I)) + 1 + lRunner, 1)
                End Select
                lRunner = lRunner + 1
            Loop
            
                HasFunctionAtPos = NameRunner & "("
                CDebug.PopP "HasFunctionAtPos"
                Exit Function
            Case Else
                'aww. oh well.
                'to be honest, I mean- seriously, for what other reason would a plugin name appear in the expression?
                CDebug.Post "Detected use of Module name """ & PlugNames(I) & """. What not followed by @ or . and so was ignored."
                
        
        End Select
        
        
        
        
        
        
        
        End If
        
    End If
    Next I
    
    
    
    For I = 0 To UBound(SplitStrings)
    'If SplitStrings(I) = "SIN" Then Stop
    
        SplitStrings(I) = Trim$(SplitStrings(I))
        If SplitStrings(I) <> "" Then
        'add a parentheses
        SplitStrings(I) = SplitStrings(I) + "("
        'If SplitStrings(I) = "DIR$" Then Stop
        If StrComp(Mid$(StrCheck, PosCheck, Len(SplitStrings(I))) & "(", SplitStrings(I) & "(", vbTextCompare) = 0 Then
            'Equal.
            HasFunctionAtPos = SplitStrings(I)
            CDebug.PopP "HasFunctionAtPos"
            Exit Function
        End If
        End If
    Next I
            HasFunctionAtPos = ""
            CDebug.PopP "HasFunctionAtPos"
End Function

'Retrieves the FunctionInformation for the requested item.
'to be frank, I REALLY need to implement the UDT's as Classes- Although, the data is read-only,
'and they only serve to store data anyway.
Public Function GetHandlersForFunction(ByVal FunctionName As String, Optional ByRef FoundCount As Long) As IEvalEvents()
Attribute GetHandlersForFunction.VB_Description = "Retrieves a collection of all handlers that return true from their CanHandleFunction() routine for the specified Function."
    'retrieves an array of loaded plugins (IEvalEvents Handlers) that return true
    'when queried with "CanHandleFunction" The order in the array determines the order in which they were found traversing
    'the collection.
    'The Array returned is 1-based.
    Dim LoopSink As IEvalEvents
    Dim retArr() As IEvalEvents
    Dim tempargs() As Boolean

    FoundCount = 0
    For Each LoopSink In Me.EvalListener.GetCollection
        If LoopSink.CanHandleFunction(Me, FunctionName, tempargs()) Then
            'if it can handle the function, add it to the array to return.
            FoundCount = FoundCount + 1
            ReDim Preserve retArr(1 To FoundCount)
            Set retArr(FoundCount) = LoopSink
        
        End If
    
    
    Next LoopSink



    GetHandlersForFunction = retArr
    'We used the ByRef FoundCount for the array increment tracking so we don't need to
    're-assign it.
End Function

Public Function GetHandlersForOperator(ByVal Operator As String, Optional ByRef FoundCount As Long) As IEvalEvents()
Attribute GetHandlersForOperator.VB_Description = "Returns a collection of all IEvalEvents Handlers that return true for their CanHandleOperator Routine."
    'retrieves an array of loaded plugins (IEvalEvents Handlers) that return true
    'when queried with "CanHandleOperator" The order in the array determines the order in which they were found traversing
    'the collection.
    'The Array returned is 1-based.
    Dim LoopSink As IEvalEvents
    Dim retArr() As IEvalEvents
    Dim tempargs() As Boolean

    FoundCount = 0
    For Each LoopSink In Me.EvalListener.GetCollection
    
        If LoopSink.CanHandleOperator(Me, Operator, UnaryOp_None) Or _
        LoopSink.CanHandleOperator(Me, Operator, UnaryOp_Prefix) Or _
        LoopSink.CanHandleOperator(Me, Operator, UnaryOp_Suffix) Then
            'if it can handle the function, add it to the array to return.
            FoundCount = FoundCount + 1
            ReDim Preserve retArr(1 To FoundCount)
            Set retArr(FoundCount) = LoopSink
        
        End If
    
    
    Next LoopSink



    GetHandlersForOperator = retArr
    'We used the ByRef FoundCount for the array increment tracking so we don't need to
    're-assign it.
End Function

Public Function GetFunctionInformation(ByVal ForFunc As String, Optional FromSink As IEvalEvents = Nothing) As FUNCTIONINFORMATION
Attribute GetFunctionInformation.VB_Description = "Retrieves the FUNCTIONINFORMATION structure for the specified function."
    Dim Ret As FUNCTIONINFORMATION
    If Not FromSink Is Nothing Then
        FromSink.GetInfoObject.GetFunctionInformation Me, ForFunc, Ret
    
    
    Else
        mEvalListener.GetInfoObject.GetFunctionInformation Me, ForFunc, Ret
    End If
    GetFunctionInformation = Ret
End Function

'The above function's return UDT can be used in the following function to
'create an appropriate string that can be displayed to the user.
Public Function FormatFunctionInformation(FuncInfo As FUNCTIONINFORMATION, Optional ByVal EFormat As FunctionHelpType = FH_QUICK) As String
Attribute FormatFunctionInformation.VB_Description = "Given the appropriate structure creates formatted output ready to show the user or write to a file."
    Dim strbuffer As String
    Dim I As Long
    With FuncInfo
        strbuffer = .StrFunctionName & "("
        For I = 0 To .ParameterCount - 1
            With .ParameterInfo(I)
            If .isOptional Then
                strbuffer = strbuffer & "[" & .Name & "]"
            Else
                strbuffer = strbuffer & .Name
            End If
            
            End With
            If I < .ParameterCount - 1 Then
                strbuffer = strbuffer & ","
            End If
        Next I

    strbuffer = strbuffer & ")"
    End With
    If EFormat = FH_EXTENSIVE Then
        'DARN! they want more information on the function.
        'OK- I guess I'll give it to them...
        strbuffer = strbuffer & vbCrLf & _
            FuncInfo.StrDescription & vbCrLf & vbCrLf & _
            "PARAMETERS:" & vbCrLf & vbCrLf
            For I = 0 To FuncInfo.ParameterCount - 1
            
            With FuncInfo.ParameterInfo(I)
                strbuffer = strbuffer & vbTab & _
                    .Name & "(" & .DataType & ")" & "-" & vbTab & _
                    .Description & vbCrLf

            
            End With
            Next I
            
        
    
    End If


    FormatFunctionInformation = strbuffer

End Function

Public Function SignalEvent(ByVal EventName As String, Params() As Variant, ByRef Ret As Variant) As FireEventReturnCodeConstants
    'SignalEvent: Handles the Event-Signalling Sub-system of BASeParserXP. Yes, a Subsystem IS inside a plugin, but SO WHAT!
    
    'was originally in FunctionHandler, until I realized that FunctionHandler simply uses CParser's Collection of CFunction objects.
    On Error Resume Next
    Dim FuncInvoke As CFunction
    'prepend "EVENT_" so that we know that the function was added with the
    'intention of being used for this purpose.
    Set FuncInvoke = mFunctions.FindByName("EVENT_" & EventName)
    If Err <> 0 Then
        'no "Handler" for this routine.
        CDebug.Post "No Handler found for event """ & EventName & """."
    Else
        'good- we found a handler.
        'invoke it.
        Ret = FuncInvoke.CallFuncArray(Params)
    End If
    


End Function

Public Function PluginLoaded(ByVal progID As String) As Boolean
    'returns wether the specified plugin is loaded.
    Dim looper As Object
    For Each looper In EvalListener.GetCollection
        If StrComp(GetObjectProgID(looper), progID, vbTextCompare) = 0 Then
            PluginLoaded = True
            Exit Function
        
        End If
    
    
    Next
End Function

Public Property Get CorePlugins() As ICorePlugins
Attribute CorePlugins.VB_Description = "Currently loaded set of Core Plugins."
    Set CorePlugins = mCorePlugins
    
End Property

Public Property Get Notation() As NumericNotationConstants
Attribute Notation.VB_Description = "specifies the notation to use when the parser performs it's own ReturnFormatting."
    Notation = mNotation
End Property

Public Property Get ReturnFormatter() As IReturnFormatter
    Set ReturnFormatter = mReturnFormatter
End Property

Public Property Get Functions() As CFunctions
Attribute Functions.VB_Description = "Returns the current collection of User-Defined Functions."
    Set Functions = mFunctions
End Property

Public Property Get Settings() As ParserSettings
    Set Settings = ParserSettings
End Property

Public Property Get configset() As String
Attribute configset.VB_Description = "Determines the current Configuration Set this CParser object is using."
    configset = mSetName
End Property

Public Property Get EvalListener() As IEvalEventSinks
    Set EvalListener = mEvalListener
End Property

Public Property Get Variables() As CVariables
Attribute Variables.VB_Description = "Returns the collection of Variables currently defined within expressions. "
    'make sure this parser is initialized.
    If mEvalListener Is Nothing Then ParserSettings.InitializeParser Me, mSetName

'    If Not mParentParser Is Nothing Then
'        Set Variables = mParentParser.Variables
'        Set mVariables = mParentParser.Variables
'    Else
        Set Variables = mVariables
    'End If
End Property

Public Property Get Expression() As Variant
Attribute Expression.VB_Description = "Assigns the current expression. If the expression differs from the current expression, a implicit ParseInFix() will occur."
Attribute Expression.VB_UserMemId = 0
    Expression = mvarExpression
End Property

Public Property Get LastResult() As Variant
Attribute LastResult.VB_Description = "Returns the last computed result."
    If IsObject(mvarConstantValue) Then
        Set LastResult = mvarConstantValue
    Else
        LastResult = mvarConstantValue
    End If
    
End Property

Public Property Get ResultAsString() As String
Attribute ResultAsString.VB_Description = "Returns the result as a string."
    'Executes, but returns a string instead of a Variant.
    'I added this so I could implement the Command-line evaluator.
    'Which, BTW- doesn't work (Naturally) in DOS alone.
    On Error Resume Next
    ResultAsString = Me.ResultToString(Me.Execute)
    'If Err <> 0 Then
        


End Property

Public Property Get State() As ParserObjectStateConstants

    State = mState

End Property

Public Property Let Notation(ByVal vdata As NumericNotationConstants)
    mNotation = vdata
End Property

Public Property Let Expression(ByVal vdata As Variant)
    
    'see if the string given is an easter egg.
    'easter eggs are neat-o.
    If CheckEasterEgg(vdata) Then
    'set it to a constant,
        mIsConstant = True
        'and make the value the changed vdata.
        mvarConstantValue = vdata
    End If
    mvarConstantValue = vbEmpty
    mIsConstant = False
    mvarExpression = vdata
    RaiseEvent ExpressionChange
    'make our stack empty.
    'the stack being empty is a sort of "dirty" flag to the parser, indicating that
    'it should parse again.
    Set mFirstItem = Nothing
    
    'ParseInfix mvarExpression
End Property

Public Property Set CorePlugins(vdata As ICorePlugins)
    Set mCorePlugins = vdata
   ' Set mCorePlugin = mCorePlugins
End Property

Public Property Set ReturnFormatter(vdata As IReturnFormatter)
    'cannot have no return formatter implementor.
    'if an attempt is made to do so, assign it to us.
    If vdata Is Nothing Then Set vdata = Me
    Set mReturnFormatter = vdata
End Property

Public Property Set Functions(vdata As CFunctions)
    Set mFunctions = vdata
End Property

Public Property Set Variables(vdata As CVariables)
    'the variables and Functions should only be assigned
    'when you know what your doing with them. this usually means your me. :P
     
    
        Set mVariables = vdata
    

End Property

Public Sub cancel()
    mcancel = True
End Sub

'Public Sub ShowAbout(Optional FParent, Optional showmode As Integer = vbModeless)
''display the about dialog.
'    If Not ((TypeOf FParent Is Form) Or TypeOf FParent Is MDIForm) Then
'        Set FParent = Nothing
'
'    End If
'
'    If Not FParent Is Nothing Then
'        FrmAbout.Show showmode, FParent
'    Else
'        FrmAbout.Show showmode
'    End If
'End Sub

Public Sub CallDynamicDispatch(withparser As CParser, OnObject As Variant, ByVal Strmember As String, Params As Variant, ByRef retval As Variant)
Attribute CallDynamicDispatch.VB_Description = "Attempts to cast the given Object to the IDynamicDispatch Interface and call it's InvokeMember Method."
    Dim castme As IDynamicDispatch
    
    Set castme = OnObject
    CDebug.Post "Object verified as implementing the IDynamicDispatch Interface."
    'you have got to be kidding me- it actually worked! Well,
    'let's get to work, then.
    'Since we want errors that occur in the InvokeMember method, even invalid function/property errors-
    'to propagate, we will reset the error handler to our CollapseStack:
   
    'OK: here goes.
    'i'm gonna do it.
    'I'm going to call the method.
    'just wait.
    'you'll see.
    'I'll call it.
    'No.
    'I'm too scared.
    'Wait...
    'OK,
    'Fine.
    'But you'll be sorry.
    Call castme.InvokeMember(withparser, Strmember, Params, False, retval)
    
End Sub

Public Sub SetExpression(ByVal vdata As String)
Attribute SetExpression.VB_Description = "Assigns the expression. does not automatically Parse the expression."
    mvarExpression = vdata
End Sub

Public Sub PerformUnaryOp(ByVal Operation As String, ByVal op As Variant, ByVal IsSuffix As Boolean, ByRef retval)
    Dim mDummy As CFormItem
    Dim mdummyfirst As CFormItem
    'performs a unary operation on the given Op.
    'create a dummy stack, just in case the plugin that implements
    'this operation needs it. doubtful, but just in case...
    Set mdummyfirst = mDummy
    Set mDummy = New CFormItem
    mDummy.ItemType = IT_OPUNARY
    If IsSuffix Then
        'it's a suffix, so add "SUFFIX"
        mDummy.Tag = "SUFFIX"
    End If
    mDummy.op = Operation
    Set mDummy.Next_ = New CFormItem
    Set mDummy = mDummy.Next_
        'the next item- the value.
        mDummy.ItemType = IT_VALUE
        If Not IsObject(op) Then mDummy.Value = op Else Set mDummy.Value = op
        'good- now call the method.
        If mEvalListener.CanHandleOperator(Me, Operation, IIf(IsSuffix, UnaryOp_Suffix, UnaryOp_Prefix)) Then
        
        
            Call mEvalListener.GetUnaryOperation(Me, Operation, op, mdummyfirst, retval)
        
        Else
            'Error- Unsupported Operator.
            Err.Raise 5, "PerformUnaryOp", "Unary Operation """ & Operation & """ not supported on type """ & TypeName(op) & """."
        
        End If
        
End Sub

Public Sub PerformOperation(ByVal Operation As String, ByVal OpA, ByVal OpB, ByRef retval)
Attribute PerformOperation.VB_Description = "Performs the given operation on the provided operands."
    'Performs the given operation directly.
    'how? well, we create a small three-element dummy Linked list.
    On Error GoTo MyHandler
    Dim mDummy As CFormItem, FirstItem As CFormItem
    Set mDummy = New CFormItem
    
    mDummy.ItemType = IT_VALUE
    If Not IsObject(OpA) Then mDummy.Value = OpA Else Set mDummy.Value = OpA
    'first item is first op...
    'make next item...
    Set mDummy.Next_ = New CFormItem
    Set FirstItem = mDummy
    Set mDummy = mDummy.Next_
    'assign to OpB:
    mDummy.ItemType = IT_VALUE
    If Not IsObject(OpB) Then mDummy.Value = OpB Else Set mDummy.Value = OpB
    'Now, make a operator item.
    Set mDummy.Next_ = New CFormItem
    Set mDummy = mDummy.Next_
    mDummy.ItemType = IT_OPERATOR
    mDummy.op = Operation
    'OK- the finale- delegate to GetOperation.
    GetOperation Operation, OpA, OpB, FirstItem, retval
    'erase mdummy. Added after I revised CFormItem to use
    'a Doubly-linked list.
'    mDummy.Destroy
    Set mDummy = Nothing

    Exit Sub
MyHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

'Public Sub ExecuteByRef(ByRef DestLoc As Variant)
'    DestLoc = Me.Execute
'End Sub
Public Sub CallFunctionByName(ByVal FunctionName As String, AllArguments() As Variant, ByRef retval As Variant)
Attribute CallFunctionByName.VB_Description = "Calls the specified Function with the passed arguments. An Error occurs if the function does not exist."
    'Calls the given function with the given arguments.
    'cheap, actually. we simply delegate
    Call mEvalListener.HandleFunction(Me, FunctionName, AllArguments, retval)



End Sub

Public Sub Create(Optional ByVal configset As String = "Default")
Attribute Create.VB_Description = "Initializes this CParser Object. Must be called first."
On Error Resume Next
CDebug.PushP "Create", "CParser", "CParser.cls", ObjPtr(Me)

If ParserSettings Is Nothing Then Set ParserSettings = New ParserSettings
CDebug.Post "Initializing Parser object... Configset=" & configset, Severity_Information
    ParserSettings.InitializeParser Me, configset
 CDebug.Post "loaded plugin count=" & Me.EvalListener.GetCollection.Count & "error code=" & Err.Description
 CDebug.Post "loaded Core plugin count=" & Me.CorePlugins.Col.Count
'    Open "C:\didfinish.txt" For Output As #9
'    Print #9, "finished."
'    Close #9
'
    mSetName = configset
    mInit = True
    
 
 
'    Dim assignIt As IEvalEventSinks
'    Set assignIt = New IEvalEventSinks
'    Set mVariables = New CVariables
'    Set mEvalListener = assignIt
'    mEvalListener.Init Me
'    assignIt.AddImplementor New BPCoreOpFunc
CDebug.PopP "Create"
CDebug.Post "parser initialization completed.", Severity_Information
End Sub

'Private Function IsFormItemConstant(FuncItem As cFormitem) As Boolean
''Returns wether a given item contains constant data.
''for example, all IT_VALUE cformitems are constant, and IT_OPERATOR as well.
''IT_FUNCTION and IT_SUBEXPRESSION, require a recursive call to this function in their respective parsers.
'Dim Retbool As Boolean, pcurritem As cFormitem
'Dim NoParseIt() As Boolean
'Retbool = True
'Set pcurritem = FuncItem
'If FuncItem Is Nothing Then
'    IsFormItemConstant = True
'    Exit Function
'End If
'  Select Case pcurritem.ItemType
'            Case IT_VALUE
'            Case IT_Operator, IT_OPUNARY
'            'these are constant.- for now.
'                If (mEvalListener.GetOperatorFlags(Me, FuncItem.op, UnaryOp_None + IIf(FuncItem.Tag = "SUFFIX", UnaryOp_Suffix, UnaryOp_Prefix)) And Operator_Optimizable) = Operator_Optimizable Then
'
'                Else
'                    IsFormItemConstant = False
'                    Exit Function
'                End If
'
'
'
'
'            Case IT_FUNCTION, IT_SUBEXPRESSION, IT_ARRAYACCESS
'            'these are constant only if their arguments/contents are constant.
'
'            Dim parr() As Variant, I As Long
'            Dim parseCast As CParser
'            If pcurritem.ItemType <> IT_SUBEXPRESSION Then
'            ReDim NoParseIt(UBound(pcurritem.Value))
'
'            If pcurritem.ItemType = IT_FUNCTION Then
'
'                If Not ((mEvalListener.GetFunctionFlags(Me, pcurritem.op) And Function_Optimizable) = Function_Optimizable) Then
'                    IsFormItemConstant = False
'                    Exit Function
'                End If
'
'                Call mEvalListener.CanHandleFunction(Me, pcurritem.op, NoParseIt())
'
'            End If
'            If IsArray(pcurritem.Value) Then
'                parr = pcurritem.Value
'                For I = LBound(parr) To UBound(parr)
'                    Set parseCast = parr(I)
'
'                    If Not NoParseIt(I) Then
'
'                        If Not parseCast.IsConstant Then
'                            'not constant-
'                            'there is a variable or something there.
'                            IsFormItemConstant = False
'                            Exit Function
'                        End If
'                    End If
'
'
'
'                Next I
'
'            Else
'                Set parseCast = pcurritem.Value
'                If Not parseCast.IsConstant Then
'                    IsFormItemConstant = False
'                    Exit Function
'                End If
'            End If
'        Case IT_VARIABLE
'            IsFormItemConstant = False
'            Exit Function
'
'
'        End Select
'
'    IsFormItemConstant = Retbool
'
'
'End Function
'delegates to the Sinker.
Public Sub AddEventSink(SinkAdd As IEvalEvents)
Attribute AddEventSink.VB_Description = "Adds a  IEvalEvents Implementor to the Stack of Implementors."

    Me.GetSinker.AddImplementor SinkAdd
    SinkAdd.Init Me
   
End Sub

'Friend Sub CheckStack(FirstItem As CFormItem)
'   'First, set up a Collection, which we will use to find circular references.
'   Dim ColCheck As Collection
'   Dim loopitem As CFormItem
'   Set ColCheck = New Collection
'    Set loopitem = FirstItem
'    Do Until loopitem Is Nothing
'
'
'
'
'        Set loopitem = loopitem.Next_
'    Loop
'End Sub
Public Sub PurgeBuffer()
    'Purges the history buffer of all entries.
    Set mFirstItem = Nothing
    OptimizedStacks.Clear
    Class_Terminate
    Create
End Sub

Public Sub Configure(Optional Owner As Object = Nothing)
Attribute Configure.VB_Description = "Displays the BASeParser Configuration Dialog."
    'Displays the configuration dialog found in the BPXPGUI ActiveX control project.
    'the createobject will probably fail in the IDE, though...
    Static AboutShow As Object
    On Error GoTo noconfigure
    If Not AboutShow Is Nothing Then
        
    Else
        Set AboutShow = VBA.createObject("BPXPGUI.BPXPConfig")
    End If
    If Not Owner Is Nothing Then
        AboutShow.Show vbModal, Owner
    Else
        AboutShow.Show vbModal
    End If
    Exit Sub
noconfigure:
'GASP! an error.
End Sub

Public Function ParseMulti(ByVal Expression As String) As Boolean
    Dim Statements() As String
    Dim I As Long
    Dim currItem As CFormItem, Formitem As CFormItem
    Dim newchild As CParser
    'Parses Multiple statement expressions.
    'step One: Use "ParseArguments" function to parse out the arguments.
    'change the optional arg to statement separators.
    'On Error Resume Next
    'there appears to be a bug in parsearguments...
    Statements = ParseArguments(Expression, , , STATEMENTSEPARATOR)
    'Stop
    'now, iterate on this array, creating a new CParser object for
    'each element.
    'remember, these new CParser objects our our children, inheriting our collections and such.
    On Error Resume Next
    If LBound(Statements) = UBound(Statements) Then
    If Err <> 0 Then Exit Function
        'HEY! whats the big idea?
        'there weren't any statement separators.
        'fine then, never mind.
        Set mMultiParsers = Nothing
        mMultiStatement = False
        ParseMulti = False
        'set formitem = nothing
        Exit Function
    
    End If
    Set mMultiParsers = New Collection
    Set Formitem = New CFormItem
    Set currItem = Formitem
    For I = LBound(Statements) To UBound(Statements)
        Set newchild = Me.Clone
        'multiparsers.Add NewChild
        currItem.ItemType = IT_STATEMENT
        newchild.Expression = Statements(I)
        'gawd! that is a lot of statements to add a friggin' item.
        currItem.op = newchild.Expression
        Set currItem.Value = newchild
        Set currItem.Next_ = New CFormItem
        Set currItem.Next_.Prev = currItem
        Set currItem = currItem.Next_
        'there we go.
    
    Next I
    'replace any existing stack.
    Set mFirstItem = Formitem
    'useless debug crap. OK, maybe not useless.
    'CDebug.Post "Parsed Multiple statements:" & Formitem.GetString(Me)
    ParseMulti = True
End Function

Private Function CheckStatement(ByVal StrCheck As String) As String
    Const Statements As String = "LISTVARS"
    Dim SplSplit() As String
    SplSplit = Split(Statements, ",")
    Dim I As Long
    For I = 0 To UBound(SplSplit)
        If StrComp(SplSplit(I), left$(StrCheck, Len(SplSplit(I))), vbTextCompare) = 0 Then
        'indeed- it i a statement.
            CheckStatement = UCase$(SplSplit(I))
            Exit Function
        End If
    Next I
    



End Function

Private Function CheckImplicit(FirstItem As CFormItem, NextItem As CFormItem) As Boolean

    Dim InsertImplicit As Boolean
    Dim IT(1 To 2) As ItemTypeConstants
    Dim I As Long
    Dim K As Long
    Dim RetTest(1 To 2) As Boolean
    Dim CheckValues(1 To 4) As ItemTypeConstants
    'idiot check- checks to see if firstitem.next_ is nextitem and nextitem.prev is firstitem.
    If Not (FirstItem.Next_ Is NextItem And NextItem.Prev Is FirstItem) Then
        CDebug.Post "FirstItem's NextItem is not NextItem or NextItem's PrevItem is not FirstItem In CheckImplicit"
        CheckImplicit = False
        Exit Function
    End If
    CheckValues(1) = IT_VALUE
    CheckValues(2) = IT_VARIABLE
    CheckValues(3) = IT_SUBEXPRESSION
    CheckValues(4) = IT_FUNCTION
    
    'value,variable,subexpression or function.
    IT(1) = FirstItem.ItemType
    IT(2) = NextItem.ItemType
    For K = 1 To 2
        RetTest(K) = False
        For I = 1 To 4
            If IT(K) = CheckValues(I) Then
                RetTest(K) = True
                Exit For
            End If
        Next I
        
    Next K
'    If RetTest(1) And RetTest(2) Then
'        'YEP! interject a IT_OPERATOR type.
'        CDebug.Post "Inserting implicit * between " & FirstItem.op & " And " & NextItem.op
'        Set FirstItem.Next_ = New CFormItem
'        With FirstItem.Next_
'            .ItemType = IT_OPERATOR
'            .op = "*"
'            .ExprPos = NextItem.ExprPos
'            Set .Next_ = NextItem
'            Set .Prev = FirstItem
'            Set .Next_.Prev = FirstItem.Next_
'        End With
'        CheckImplicit = True
'    End If
        

End Function

'Utility functions.


Private Function ExtractBrackets(ByVal FromString As String, Optional ByVal ParseStart As Long = 1, Optional ByVal BracketStart As String = "([{", Optional ByVal Bracketend As String = ")]}") As String

    'the first item pointed to should be the starting delimiter, as in the ( right after a function name.
    Dim inQuote As Boolean
    Dim bracketcount As Long
    Dim Char As String
    Dim I As Long
    I = ParseStart
    Do Until I > Len(FromString)
        Char = Mid$(FromString, I, 1)
        Select Case True
            Case Char = """"
                inQuote = Not inQuote
            Case inQuote
                'ignore.
            Case InStr(BracketStart, Char) <> 0
            'start of a bracket.
                bracketcount = bracketcount + 1
            Case InStr(Bracketend, Char) <> 0
                bracketcount = bracketcount - 1
                If bracketcount = 0 Then
                    'there it is, the end.
                    Exit Do
                
                End If
    
    
        End Select
        I = I + 1
    Loop
    ExtractBrackets = Mid$(FromString, ParseStart + 1, I - ParseStart - 1)
End Function

Private Function RemoveExtraneousSpaces(ByVal FromString As String) As String
    Dim I As Long, inQuote As Boolean, intick As Boolean
    Dim parsework As String, Char As String
    Dim CreateBuffer As cStringBuilder, ignorechar As Boolean
    Set CreateBuffer = New cStringBuilder
    
    Const MessyChars As String = vbCrLf & vbNullChar & vbTab
    CDebug.PushP "RemoveExtraneousSpaces", MODULE_NAME, MODULE_FILE_NAME, 0, "Removing Extraneous Spaces from """ & FromString & """."
    For I = 1 To Len(FromString)
    Char = Mid$(FromString, I, 1)
        Select Case True
            Case Char = "`"
                intick = Not intick
            Case Char = """"
            
                inQuote = Not inQuote
            Case Char = " "
                If Not inQuote And Not intick Then
                        ignorechar = True
                
                End If
            Case inQuote Or intick   'ignore!
                ignorechar = False
            Case InStr(MessyChars, Char) <> 0 And Not inQuote
            'Moved this CASE above the GetOperator case.
            'just in case. :P
                'ignore this character, since it is defined as a "messy char"
                ignorechar = True
            
            Case GetOperator(FromString, I) <> ""
                parsework = GetOperator(FromString, I)
                'unary operators might get mixed up with numbers.
                'so check if this is a unary operator.
                If EvalListener.Self.CanHandleOperator(Me, parsework, UnaryOp_None) Then
                    'if so, add it, and a space.
                    'CreateBuffer = CreateBuffer & parsework & " "
                    CreateBuffer.Append parsework
                    I = I + (Len(parsework) - 1)
                    ignorechar = True
            
                End If
          
        End Select
        If Not ignorechar Then CreateBuffer.Append Mid$(FromString, I, 1)
        ignorechar = False
    Next I
    

    RemoveExtraneousSpaces = CreateBuffer.ToString
    CDebug.PopP "RemoveExtraneousSpaces"

End Function

Private Function GetEndBracket(ByVal CharFor As String) As String
    Select Case CharFor
        Case "("
            GetEndBracket = ")"
        Case "{"
            GetEndBracket = "}"
        Case "["
            GetEndBracket = "]"
    End Select
End Function

Private Function MakeParsable(ByVal StrString As String) As String


    'MakeParsable- makes a already properly formed expression parsable by RipFormula by
    'removing all whitespace and other unnecessary garbage from the expression. doesn't magically make an expression parsable
    'that consists of garbage, though- GIGO. Garbage In Garbage Out.
    Dim OpsSplit() As String, I As Long
    Dim Char As String
    Dim opGet As String, BracketInterior As String
    Dim inQuote As Boolean
    Dim TempItem As CFormItem
    
    Dim strbuffer As String, tempfunc As String
    Dim appendChunk As String, InDatelit As Boolean
    CDebug.Post "MakeParsable entered."
    Set TempItem = New CFormItem
    StrString = Replace$(StrString, vbCrLf, "")
    StrString = RemoveExtraneousSpaces(StrString)
    
    'For I = 1 To Len(StrString)
    I = 1
    Dim priorPos As Long
    Do Until I > Len(StrString)
        Char = Mid$(StrString, I, 1)
        'If Char = "%" Then Stop
        priorPos = I
       Select Case True
        Case Char = """"
            inQuote = Not inQuote
        Case Char = "#" And Not inQuote
             InDatelit = Not InDatelit
        Case inQuote Or InDatelit
            'ignore-
        Case InStr("[{(", Char) And Not inQuote
        'ignore brackets.
            BracketInterior = ExtractBrackets(StrString, I, Char, GetEndBracket(Char))
            'Now, we need to put this stuff on Strbuffer.
            'I feel like I wrote this code around a billion times.
            'and I did.
            'Oh well.
            appendChunk = Char & BracketInterior & GetEndBracket(Char)
            strbuffer = strbuffer & appendChunk
            'as usual, subtract one so that the I=I+1 at the end of the loop wil take us to
            'the char after the ending bracket.
            I = I + Len(appendChunk) - 1
            Char = ""
        
            Dim Throw As String
            Dim Chargrab As String
            Dim quickextract As String
        Case Char = "@"
                'lookie here, we found a Object access operator.
                'because this stupid library F'ed up everything,
                'I am assigned the simple task of un-f-ing everything up.
                'how delightful.
                'now, what is next is either a Method call or property access
                'irregardless, ParseValue might not work, because it's a piece of shit.
                'so, instead, I need to do it myself, and find the bracket- or a space.
                'if they don't put a space after the method name, that's their fault, and they can go screw themselves.
                'of course, we are faced with the problem that this function removed all the spaces. GRRRR!
                'Not a space then- look for the next operator that contains a non-alphabet character as it's first character.
                'OK.
                'oh, and if we find a bracket, that is even easier.
                'so:
                'a bracket, or a operator. HOWEVER- the operator must NOT have a alphabetic character as the first
                'character.
                strbuffer = strbuffer & "@"
                Do
                    I = I + 1
                    
                    Chargrab = Mid$(StrString, I, 1)
                    If Chargrab = "(" Then
                        'ok- function access. cool.
                        'WAIT! even better. properties are bad.
                        'no properties. they will be accessed via parentheses. NO EXCEPTIONS.
                        'this saves be the agonizing quest to find the end of the property name, without
                        'getting stuck.
                        'extract the brackets, and move to the other side.
                        quickextract = ExtractBrackets(StrString, I)
                        strbuffer = strbuffer & "(" & quickextract & ")"
                        I = I + Len(quickextract) + 1
                        
                        Exit Do
                    End If
                    strbuffer = strbuffer & Chargrab
                    If I > Len(StrString) Then
                        Err.Raise Parse_Badsyntax, "CParser::MakeParsable", "Object Operator @ method identifier must be suffixed with ""()""."
                    End If
                Loop
                Char = ""
                'I = I - 1
            
        Case Mid$(StrString, I, 1) = "@"
            '
                    I = I - 1
        'OK- check to see if our Core plugin understands and can use the
        'string that starts at this location.
        'note that we create a cFormItem beforehand but (perhaps erroneously) don't give it
        'a prior or next value...
        Case mCorePlugin.ParseLocation(Me, StrString, I, TempItem) = ParseLocation_recognized
        'the new position variable will have been updated by the Coreplugin, and as such we require no code here.
                
        'OK, I lied- we need to
        'place the correct portions of the string.
        'tempitem.extra is the location designated
        'for this purpose.
        
        'strbuffer = strbuffer &
        strbuffer = strbuffer & Mid$(StrString, priorPos, I - priorPos)
        Char = ""
        I = I - 1
        CDebug.Post "MakeParsable:A Core Plugin has recognized a token. Ignoring portions of the expression."
        CDebug.Post "The Core plugin recognized """ & Mid$(StrString, priorPos, I - priorPos + 1) & """"
       ' I = I + 1
        Case Else
                'if we didn't meet any of the special conditions.
                
                
                
             
                'NOTE: this code INTENTIONALLY executes
                'even if the Function or variable tests proved true. (well, at least NOW it
                'is intentional). This saves us another iteration.
                opGet = GetOperator(StrString, I)
                If opGet <> "" Then
                    'we found an operator, but put it on the buffer with spaces.
                    'yeah, so what, I use one on each side.
                    strbuffer = strbuffer & " " & opGet & " "
                    Char = ""
                    I = I + (Len(opGet) - 1)    'only add one, because of the +1 at the Loop end.
                'Else
                    'strbuffer = strbuffer & Char
                Else
                    
                       'if there is a function at the position, move the current
                    'position to the other side of the function, so we don't
                    'insert spaces that separate it.
                    'code moved after operator test.
                    tempfunc = HasFunctionAtPos(StrString, I)
                    If tempfunc <> "" Then
                
                        'add it to the buffer.
                        Throw = tempfunc & ExtractBrackets(StrString, I + Len(tempfunc) - 1) & ")"
                        
                        strbuffer = strbuffer & Throw
                        I = I + Len(Throw) - 1
                        'oh, and empty out char, so we don't get an extra one.
                        Char = ""
                    Else
                        'skip variables, too.
                        tempfunc = hasVariableAtPos(StrString, I)
                        If tempfunc <> "" Then
                            I = I + Len(tempfunc) - 1
                            'add it to the buffer.
                            strbuffer = strbuffer & tempfunc
                            'oh, and empty out char, so we don't get an extra one.
                            Char = ""
                        End If
                        
                    End If 'tempfunc <> ""
                    
                End If  'opget <> ""
            End Select
            
            
        I = I + 1
        strbuffer = strbuffer & Char
    Loop
    
   ' Next I
    
    
    MakeParsable = strbuffer


    CDebug.Post "makeparsable exited."
End Function

Private Function GetPriority(ByVal OpCheck As String) As Long


    
    OpCheck = Trim$(OpCheck)
    OpCheck = " " & OpCheck & " "
    CDebug.PushP "GetPriority", "CParser", "Cparser.cls", ObjPtr(Me)
    GetPriority = InStr(1, mEvalListener.GetHandledOperatorsString(Me), OpCheck, vbTextCompare)
    CDebug.PopP "GetPriority"
End Function

Private Function GetOperator(ByVal Expression As String, ByVal Currpos As Long) As String


'keep grabbing another character out of the expression from currpos, until we hit a space, or a separator object.
'we cannot use getValue since that will cause a mutual recursion problem.
Dim OpRunner As String
Dim Oppos As Long
Dim OpLength As Long
Dim LenReduce As Long, getop As String
CDebug.PushP "GetOperator", MODULE_NAME, MODULE_FILE_NAME, 0
Oppos = Currpos
OpLength = 1
If HasFunctionAtPos(Expression, Currpos) <> "" Then
    CDebug.PopP "GetOperator"
    Exit Function
End If

OpRunner = Mid$(Expression, Currpos, 1)
'If OpRunner = "*" Then Stop
'loop until we get a space. ALL operators are defined to require a space before and after.
'the @ operator will be used for calling by name via the TypeLib InvokeArray function.
'as such, if we have that right now, return it.
If InStr("1234567890.", OpRunner) Or OpRunner = """" Then
    GetOperator = ""
    CDebug.PopP "GetOperator"
    Exit Function

End If
Select Case OpRunner
   ' Case "@"
     '   GetOperator = "@"
     '   Exit Function
End Select
'anyway, continue until we hit, a space
'if we hit a ")" or the end of the string, we have a problem.
'ignore it for now, I just want this to work in a rudimentary fashion.
OpLength = 0
Do
 Dim checkchar As String
 checkchar = Mid$(Expression, Oppos + OpLength, 1)
 Select Case True
  
    Case checkchar = ")"
    'shit.
        OpRunner = Mid$(Expression, Oppos, OpLength)
        Exit Do
    
    Case checkchar = "(", checkchar = " "
        'operator ended.
        OpRunner = Mid$(Expression, Oppos, OpLength)
        Exit Do
        'PRB: shouldn't numbers be allowed in variables?....
    Case InStr("123456789.", checkchar) <> 0 Or checkchar = """"
    'Case checkchar = "." Or checkchar = """"
        OpRunner = Mid$(Expression, Oppos, OpLength)
        Exit Do

 End Select
 OpLength = OpLength + 1
Loop
Dim LastValid As String


For LenReduce = OpLength To 1 Step -1
    If CBool(GetPriority(Mid$(Expression, Oppos, LenReduce))) Then
        getop = Mid$(Expression, Oppos, LenReduce)
        Exit For
    
    End If


Next
'OK, this is the verification phase.
'we need to verify, with all the ParseEvents handlers, that
'this is indeed an operator. If any of them return true, we return it.
'Otherwise, we return "".
Dim GetRet As Boolean
'//TODO: optimize this so we only loop through the handlers ONCE- this function
'iterates through them all and we call it THREE TIMES!
If mEvalListener.CanHandleOperator(Me, getop, UnaryOp_None) Then
    GetOperator = getop

ElseIf mEvalListener.CanHandleOperator(Me, getop, UnaryOp_Prefix) Then

    GetOperator = getop
ElseIf mEvalListener.CanHandleOperator(Me, getop, UnaryOp_Suffix) Then
    GetOperator = getop
Else

    GetOperator = ""


End If





    'if we make it here, it is NOT an operator.










    CDebug.PopP "GetOperator"







End Function



Private Function IOperable_PerformFunc(withparser As CParser, ByVal funcName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants
'
End Function

Private Sub IOperable_PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, retval As Variant)
'
End Sub

Private Function IOperable_toString(withparser As CParser) As String
Execute
If Not mFirstItem Is Nothing Then
IOperable_toString = "PARSER- stack=" & mFirstItem.GetString(withparser)
Else
IOperable_toString = "PARSER- <Empty stack>"

End If
End Function

Private Function IReturnFormatter_ResultToString(ByVal VarResult As Variant, Handled As Boolean) As String
    'no implementation.
    Handled = False
End Function

'Public Function IsConstant() As Boolean
'    IsConstant = False
'End Function
Private Function IsFormItemConstant(pCurrItem As CFormItem) As Boolean
Dim Retbool As Boolean, looper As Long
If pCurrItem Is Nothing Then
    IsFormItemConstant = True
    Exit Function
End If
IsFormItemConstant = True
     Select Case True
                Case IT_NULL
                    IsFormItemConstant = True
                Case pCurrItem.ItemType = IT_VARIABLE
                    'a variable present in the expression means that it can NEVER be optimized.
                    IsFormItemConstant = False
                    Exit Function

                Case pCurrItem.ItemType = IT_ARRAYACCESS, pCurrItem.ItemType = IT_FUNCTION
                    'in either of these two cases, the Value of pCurritem is a array of Parser objects.
                    On Error GoTo ReturnFalse
                    For looper = LBound(pCurrItem.Value) To UBound(pCurrItem.Value)
                        If Not pCurrItem.Value(looper).IsConstant Then
                            IsFormItemConstant = False
                            Exit Function
                        End If
                    Next
                    If pCurrItem.ItemType = IT_FUNCTION Then
                        'if all the values are constant, check if the function is too.
                        'If Not mEvalListener.CanOptimizeFunction(pcurritem.Op) Then
                        If Not ((mEvalListener.GetFunctionFlags(Me, pCurrItem.op) And Function_Optimizable) = Function_Optimizable) Then
                            'nope- we cannot.
                            IsFormItemConstant = False
                            Exit Function
                        End If

                    End If


                Case pCurrItem.ItemType = IT_SUBEXPRESSION
                    If Not pCurrItem.Value.IsConstant Then
                        IsFormItemConstant = False
                        Exit Function
                    End If

                Case pCurrItem.ItemType = IT_OPUNARY, pCurrItem.ItemType = IT_OPERATOR
                    'Debug.Assert False
                   If (mEvalListener.GetOperatorFlags(Me, pCurrItem.op, UnaryOp_None + IIf(pCurrItem.Tag = "SUFFIX", UnaryOp_Suffix, UnaryOp_Prefix)) And Operator_Optimizable) = Operator_Optimizable Then
                    '
                    Else
                        IsFormItemConstant = False
                        Exit Function
                    End If
                    'Exit Do
                Case pCurrItem.ItemType = IT_ObjAccess
                    IsFormItemConstant = False

                End Select
IsFormItemConstant = Retbool
Exit Function
ReturnFalse:
    IsFormItemConstant = False
End Function

'OK, Once again- I LOVE VB, and I am starting to hate FreeBasic for
'sending my innocent VB DLL unicode characters.
Private Function IsSubStringofFunction(ByVal StrCheck As String) As String
    Dim SplitStrings() As String
    Dim FullString As String, I As Long
    FullString = mEvalListener.GetHandledOperatorsString(Me)
    SplitStrings = Split(FullString, " ")
    For I = 0 To UBound(SplitStrings)
        If InStr(SplitStrings(I), StrCheck) <> 0 Then
            IsSubStringofFunction = SplitStrings(I)
        End If
    Next I
End Function

Private Function GetLoadedPluginNames() As String()
    Dim retArr() As String, LoopSink As IEvalEvents, I As Long
    CDebug.PushP "GetLoadedPluginNames", "CParser", "Cparser.cls", ObjPtr(Me)
    For Each LoopSink In EvalListener.GetCollection
        I = I + 1
        ReDim Preserve retArr(1 To I)
        retArr(I) = TypeName(LoopSink)
        
    
    Next

    GetLoadedPluginNames = retArr()
    CDebug.PopP "GetLoadedPluginNames"
End Function

Private Function OldMakeParsable(ByVal StrString As String) As String
    Dim OpsSplit() As String, I As Long
    Dim Char As String
    Dim opGet As String, BracketInterior As String
    Dim inQuote As Boolean
    Dim strbuffer As String, tempfunc As String
    Dim appendChunk As String, InDatelit As Boolean
    StrString = Replace$(StrString, vbCrLf, "")
    StrString = RemoveExtraneousSpaces(StrString)
    
    'For I = 1 To Len(StrString)
    I = 1
    Do Until I > Len(StrString)
        Char = Mid$(StrString, I, 1)
       

        If InStr("[{(", Char) And Not inQuote Then
        'ignore brackets.
            BracketInterior = ExtractBrackets(StrString, I, Char, GetEndBracket(Char))
            'Now, we need to put this stuff on Strbuffer.
            'I feel like I wrote this code around a billion times.
            'and I did.
            'Oh well.
            appendChunk = Char & BracketInterior & GetEndBracket(Char)
            strbuffer = strbuffer & appendChunk
            'as usual, subtract one so that the I=I+1 at the end of the loop wil take us to
            'the char after the ending bracket.
            I = I + Len(appendChunk)
        
        Else
            Dim Throw As String
            Dim Chargrab As String
            Dim quickextract As String
            If Char = "@" And Not inQuote And Not InDatelit Then
                'lookie here, we found a Object access operator.
                'because this stupid library F'ed up everything,
                'I am assigned the simple task of un-f-ing everything up.
                'how delightful.
                'now, what is next is either a Method call or property access
                'irregardless, ParseValue might not work, because it's a piece of shit.
                'so, instead, I need to do it myself, and find the bracket- or a space.
                'if they don't put a space after the method name, that's their fault, and they can go screw themselves.
                'of course, we are faced with the problem that this function removed all the spaces. GRRRR!
                'Not a space then- look for the next operator that contains a non-alphabet character as it's first character.
                'OK.
                'oh, and if we find a bracket, that is even easier.
                'so:
                'a bracket, or a operator. HOWEVER- the operator must NOT have a alphabetic character as the first
                'character.
                strbuffer = strbuffer & "@"
                Do
                    I = I + 1
                    
                    Chargrab = Mid$(StrString, I, 1)
                    If Chargrab = "(" Then
                        'ok- function access. cool.
                        'WAIT! even better. properties are bad.
                        'no properties. they will be accessed via parentheses. NO EXCEPTIONS.
                        'this saves be the agonizing quest to find the end of the property name, without
                        'getting stuck.
                        'extract the brackets, and move to the other side.
                        quickextract = ExtractBrackets(StrString, I)
                        strbuffer = strbuffer & "(" & quickextract & ")"
                        I = I + Len(quickextract) + 1
                        
                        Exit Do
                    End If
                    strbuffer = strbuffer & Chargrab
                    If I > Len(StrString) Then
                        Err.Raise Parse_Badsyntax, "CParser::MakeParsable", "Object Operator @ method identifier must be suffixed with ""()""."
                    End If
                Loop
                Char = ""
            End If
            Select Case Char
            Case "#"
                If Char = "#" Then InDatelit = Not InDatelit
            
                If Char = """" Then inQuote = Not inQuote
            
                'if there is a function at the position, move the current
                'position to the other side of the function, so we don't
                'insert spaces that separate it.
                
                tempfunc = HasFunctionAtPos(StrString, I)
                If tempfunc <> "" Then
            
                    'add it to the buffer.
                    Throw = tempfunc & ExtractBrackets(StrString, I + Len(tempfunc) - 1) & ")"
                    
                    strbuffer = strbuffer & Throw
                    I = I + Len(Throw)
                    'oh, and empty out char, so we don't get an extra one.
                    Char = ""
                Else
                    'skip variables, too.
                    tempfunc = hasVariableAtPos(StrString, I)
                    If tempfunc <> "" Then
                        I = I + Len(tempfunc) - 1
                        'add it to the buffer.
                        strbuffer = strbuffer & tempfunc
                        'oh, and empty out char, so we don't get an extra one.
                        Char = ""
                    End If
                End If
            
                opGet = GetOperator(StrString, I)
                If opGet <> "" And Not inQuote And Not InDatelit Then
                    'we found an operator, but put it on the buffer with spaces.
                    'yeah, so what, I use one on each side.
                    'Like Ripformula cares. Actually, I could probably remove this whole "MakeParsable" stuff.
                    strbuffer = strbuffer & " " & opGet & " "
                    I = I + (Len(opGet) - 1)    'only add one, because of the +1 at the Loop end.
                Else
                    strbuffer = strbuffer & Char
                End If
                
                
                'HACK ALERT.
                
            End Select
            If Mid$(StrString, I, 1) = "@" And Not inQuote Then
                    I = I - 1
            End If
        I = I + 1
        End If
    Loop
    
   ' Next I
    
    
    OldMakeParsable = strbuffer



End Function

Private Property Get mCorePlugin() As ICorePlugin
    Set mCorePlugin = CorePlugins
End Property

Private Sub AppendItem(ByRef ToItem As CFormItem)
    Set ToItem.Next_ = New CFormItem
    Set ToItem = ToItem.Next_
End Sub

Private Sub CollapseStack(startitem As CFormItem, ByRef valret As Variant)
    'OH MY GOD THE STACK IS FALLING!
    
    '"Collapses" the currently held stack. Well, it isn't really a stack-
    'the structure seems more like a linked list.
       Dim rettemp As Variant
    Dim rVal As Variant
    Dim currItem    As CFormItem    ' current command item
    Dim ParseStack  As DataStack ' result stack
    Dim OpA As Variant, OpB As Variant, retval
    Dim PopTemp As Variant, tempvar As Variant
    Dim cancel As Boolean, I As Long
    Dim castme As IDynamicDispatch  'let's give this a shot :)
    
    'through rigorous testing, I have determined that collapseStack is executed the highest number of
    'times. As such, it is where I will place the "Cancel" test.
    If mcancel Then
        Err.Raise 18        'user interrupt occured.
    End If
    
    
    CDebug.PushP "CollapseStack", MODULE_NAME, MODULE_FILE_NAME, 0
    mState = ParseState_Executing
    'Clear out our Unary Cache.
    mUnaryCache.Clear
    On Local Error GoTo Err_CollapseStack
    Set ParseStack = New DataStack
    Set currItem = startitem
    
    
    
    
    'set error handler
    
    Do
    
       On Error GoTo Err_CollapseStack
       'moved this On Error into the loop.
        Select Case currItem.ItemType
            Case IT_NULL
                'nothing. thus the term, "Null"
                'Not sure WHY I added this itemtype.
                'I am sure of one thing, I REALLY like this itemtype architecture...
            Case IT_ENDOFLIST
                Exit Do
           
            
            Case IT_STATEMENT
            
                'execute that value. Oh, and assign it the the result.
                'Stop
                Call Assign(valret, currItem.Value.Execute)
              
            Case IT_BLOCK
            '
            'IT_BLOCK is a somewhat "special" type of Item. It is functionally similiar to IT_SUBEXPRESSION, however it is separate from it
            'to prevent ambiguity, since IT_BLOCK has both a different structure definition as well as a different usage context.
            
            'IT_BLOCK defines
            
            Case IT_VALUE
                ParseStack.Push currItem.Value
            Case IT_VARIABLE
             
                cancel = False
                On Error Resume Next
                rettemp = mEvalListener.BeforeResolveVariable(Me, currItem.op, cancel)
                If Err <> 0 Then
                    Set rettemp = mEvalListener.BeforeResolveVariable(Me, currItem.op, cancel)
                End If
                If cancel Then
                    ParseStack.Push rettemp
                Else
                    'BUGFIX:
                    'finally decided to find out why the parser sometimes actually
                    'returned a CVariable Object.
                    'turns out I forgot to say .Value here, and so the actual CVariable
                    'object was being Pushed onto the stack. Well I'll be.
                    If Me.Variables.Exists(currItem.op) Then
                        ParseStack.Push Me.Variables(currItem.op).Value
                    End If
                End If
            Case IT_FUNCTION
                
                PerformFunction currItem, tempvar
                ParseStack.Push tempvar
            Case IT_ARRAYACCESS
                'pop the last value. it should be an array.
                'however, it can also be a string.
                'here is the tricky part-
                'we can also have multiple items.
                Dim strcast As String
                Dim subscripts As Variant
                Dim GetScripts As Variant
                PopTemp = ParseStack.Pop
                subscripts = currItem.Value
                If IsArray(PopTemp) Then
                    'right now, it seems I have been backed into a corner and need to find some other way
                    'to access the proper item, using a subscript reference.
                    'Basically, I cannot at the moment retrieve an element
                    'from the array given another array that contains the subscripts.
                    
                    ReDim GetScripts(LBound(subscripts) To UBound(subscripts))
                    For I = LBound(subscripts) To UBound(subscripts)
                    On Error Resume Next
                        'Array subscripts. They kind of need
                        'to be numbers.
                        
                        GetScripts(I) = CLng(subscripts(I).Execute)
                        
                    
                    Next I
                    On Error GoTo Err_CollapseStack
                    'ARRAY_STANDARD can be defined.
                    'not sure why, though. oh well.
                #If ARRAY_STANDARD Then
                    ParseStack.Push PopTemp(currItem.Value(0).Execute)
                #Else
                    'ParseStack.Push GetArrayElement(PopTemp, GetScripts)
                    'changed. didn't realize I already wrote a getelement for arrays.
                    'but the new one was better anyway....
                    ParseStack.Push Array_GetElement(Me, PopTemp, GetScripts)
                
                #End If
                Else
                    
                    
                    'CHANGE:
                    'if it isn't an Array, we can't handle it.
                    'Find out if any of our plugins can, before we bitch to the client.
                    If mEvalListener.HandleSubScript(Me, PopTemp, subscripts, rVal) Then
                        'whew- push rval onto the stack.
                        ParseStack.Push rVal
                    Else
                    'uh oh, we don't know how to do it.
                    Err.Raise 5, "Cparser::CollapseStack", "No defined handler to perform subscript operation on value of type """ & TypeName(PopTemp) & """."
                    
                    
                    End If
                    
                    

'                    ParseStack.Push strcast
                
                End If
            
            
            Case IT_SUBEXPRESSION
                'STOP!-
                'OK, I lied. I need to do one thing before I execute this as a sub-expression.
                'I need to see if the caller wants this expression to be used as the parameter list to a
                'objects default method. Oddly enough- this seems pretty easy.
                'use Peek to see if the previous item on the stack is an object.
                'if so, perform an invokedynamic call on it.
                
                'UPDATE: removed the default method crap.
                'it simply screws up too much other stuff.
                Dim HoldPrevious As Variant
                Dim argsparsed() As String, ParseObjs, argsevaluated()
                    On Error Resume Next
                 
                        ParseStack.Push currItem.Value.Execute
              
            Case IT_OPUNARY
                'unary operation.
                
                
                
                If currItem.Tag = "SUFFIX" Then

                    'it is a suffix unary operator, this means
                    'that there IS already a value on the parsestack.
                    'pop it, ask to evaluate, and we're finished with this iteration.
                    If IsObject(ParseStack.Peek) Then Set rVal = ParseStack.Pop
                    If IsEmpty(rVal) Then rVal = ParseStack.Pop
                    'OK, now evaluate...
                    'Stop
                    If Not mEvalListener.CanHandleOperator(Me, currItem.op, UnaryOp_Suffix) Then
                        'descriptive error.
                        Debug.Assert False
                    
                    End If
                    GetUnaryOperation currItem.op, rVal, currItem, retval
                    ParseStack.Push retval
                
                Else
                    'Unary Prefix. This is actually MORE complicated, due to the
                    'way the parsestack is built.
                    'the key here is that the unary operation must be performed on the VERY next item that gets pushed to the stack-
                    'as such, we set up a temporary Event sink for the DataStack to this Cparser.
                    'and cache the appropriate values- curritem, to be exact.
                    'the DataStack will then raise the event, allowing us to change the pushed value.
                    'Set mUnaryCache = Curritem
                    If Not EvalListener.Self.CanHandleOperator(Me, currItem.op, UnaryOp_Prefix) Then
                        ParserError(Me, Parse_MalformedExpression, "Operator """ & currItem.op & " is not a unary operator.").Throw
                    End If
                    mUnaryCache.Push currItem
                    Set UnarySink = ParseStack
                    'now it is a waiting game...
                    'rVal = ParseStack.Pop
                    'Call GetUnaryOperation(Curritem.Op, rVal, retval)
                    
                    'ParseStack.Push retval
                End If
            Case IT_OPERATOR
                '
                'pop the last two items, then use them to perform the op.
                'the first item popped is also on the left side.
                'Dim rVal As Variant
                'NOTE: here is where we would add support for
                'unary operators by pushing OpB back onto the stack and
                'using OpA against the unary...
                Dim tempAssign As CFormItem
                'perform a quick check of the stack- if we only have one item,
                'then we are missing an operand.
                If ParseStack.Count < 2 Then
                    'sigh.
                    'mdlparse.parsererror()
                    ParseError(Me, Parse_MalformedExpression, "Binary Operator """ & currItem.op & """ used with insufficient number of operands.", "CollapseStack").Throw
                End If
                If IsObject(ParseStack.Peek) Then
                    Set OpB = ParseStack.Pop
                Else
                    OpB = ParseStack.Pop
                End If
                If IsObject(ParseStack.Peek) Then
                   Set OpA = ParseStack.Pop
                Else
                    OpA = ParseStack.Pop
                End If
                'Set tempAssign = Curritem.Tag
                On Error GoTo Err_CollapseStack
                GetOperation currItem.op, OpA, OpB, currItem, rVal
                'ParseStack.Push GetOperation(Curritem.Op, opA, OpB)
                ParseStack.Push rVal
          Case IT_ObjAccess
            'Pop off the object from the stack.
            'Now, this is ALSO valid with Arrays- however,
            'we need to perform explicit handling of the array methods.
            'UPDATE: the array handling logic will be moved into BPCoreOpFunc.
            'Thanks to the new Interface method, OnInvokeObject.
            Dim CreateParams() As Variant
            Dim boolflag As Boolean
            If ParseStack.Count > 0 Then
                If IsObject(ParseStack.Peek) Then
                    Set OpA = ParseStack.Pop
                 '   boolflag = True
                    
                    
                Else
                'boolflag = False
                OpA = ParseStack.Pop
                End If
            Else
             If Not mImpliedObject Is Nothing Then
                Set OpA = mImpliedObject
             Else
                ParserError(Me, 424, "Object Required: " & currItem.op, , currItem.ExprPos).Throw
             End If
            End If
            'If TypeOf OpA Is CObject Then Stop
            If Not IsObject(OpA) Then
                
                'OpA = ParseStack.Pop
                'Since it isn't an object, we (CParser) can no longer deal with it.
                'we delegate any other such tasks to the event handler. since BPCoreOpFunc should always be there,
                'Array support is still guaranteed.
                If Not mEvalListener.OnInvokeObject(Me, OpA, currItem, retval) Then
                    'Err.Raise 5, "Cparser::CollapseStack", "No Object Access ""@"" behaviour defined for " & TypeName(OpA)
                 
                    Err.Raise 5, "Cparser::CollapseStack", "No Object Access ""@"" behaviour defined for " & TypeName(OpA) & " and no Implied Object set."
                 
                Else
                    ParseStack.Push retval
                    GoTo NEXTITERATION
                End If
                
                'whew, at least one of our plugins had a clue.
                'pretend it was us that did it :P
                'hey- we are the manager, that is what managers are supposed to do.
                
                
                 
            End If
            On Error Resume Next
            If Not IsArray(OpA) Then  'Skip if it is an array.
            'If TypeOf OpA Is CObject Then Stop
            ReDim CreateParams(LBound(currItem.Value) To UBound(currItem.Value))
            If Err = 0 Then
            Err.Clear
                For I = LBound(currItem.Value) To UBound(currItem.Value)
                    On Error GoTo Err_CollapseStack
                    Assign CreateParams(I), currItem.Value(I).Execute
                
                Next I
                'OK,- now we call the method on it.
                'IT_ObjAccess is supplemented by the Function parsing part,
                'so we have the same info as a IT_FUNCTION- that is, OP is the
                'actual function name, and value is an array of CParsers.
                'Also, check for an array.
          
                If Err <> 0 Then
                
                
                End If
                Err.Clear
                On Error Resume Next
                Call InvokeDynamic(OpA, currItem.op, CreateParams, retval)
                If Err = 0 Then
                    ParseStack.Push retval
 
                Else
                    
                   
                   
                        Err.Clear
                        CallDynamicDispatch Me, OpA, currItem.op, CreateParams, retval
                         If Err <> 0 Then
                            'raise a different error message, instead of type mismatch or whatever VB says.
                            On Error GoTo 0
                            Err.Raise 5, "CParser::CollapseStack", "Method or Property, " & currItem.op & " Not defined for object type " & TypeName(OpA)
                    
                        End If
                        ParseStack.Push retval
                   
                
                
                End If
          
            Else
                'if an error occured, the call was without arguments, and thus no parameters.
                'no args- pass vbempty.
                Err.Clear
                On Error Resume Next
                Call InvokeDynamic(OpA, currItem.op, vbEmpty, retval)
                'push it onto the stack.
                
                If Err = 0 Then
                    'it succeeded, - good. Less work for us.
                      ParseStack.Push retval
                Else
                    Err.Clear
                    'clear out the error.
                   
                    
                    
                    'OK- this is our last chance.
                    'This was added Sunday Feb 18 2007 5:21 PM.
                    'it's an object- we know that much. however, if we can cast the object
                    'to a IDynamicDispatch, we can then call it's
                    'member function, InvokeMember. This works for BASeParser the same
                    'way IDispatchEx works for COM.
                    'if an error still occurs, it should be allowed to propagate.
                    
                    Set castme = OpA
                    
                    
                 
                        
                        CallDynamicDispatch Me, OpA, currItem.op, vbEmpty, retval
                        If Err <> 0 Then
                            'raise a different error message, instead of type mismatch or whatever VB says.
                            On Error GoTo Err_CollapseStack
                            Err.Raise 5, "CParser::CollapseStack", "Method or Property, " & currItem.op & " Not defined for object type " & TypeName(OpA)
                        End If
                        ParseStack.Push retval
                    
                
                End If
                

            End If
        End If
        Case IT_BACKTICKS
            'Backticks. Used for interpolated strings.
            'or something. Actually, not sure yet.
            'in the making...
            'huh. Oh, Backticks, eh? never mind those...
            
            'Debug.Assert False
        
        
        
        Case Is < 0
            'probably a custom plugin.
            If mCorePlugin.CanHandleItem(Me, currItem) Then
                CDebug.Post "Core plugin returned ""True"" for canhandleItem."
                mCorePlugin.HandleItem Me, currItem, retval
                'push it's result onto the stack.
                
                ParseStack.Push retval
            
            End If
            
        Case IT_FLOWCHANGE
        Dim DoTerm As Boolean
            'a flow change.
            'this is pretty simple to implement(well, while collapsing- parsing it into the tree is a different story altogether):
            'First, check condition...
            '(in curritem.Extra)
            '(should be a CParser object, we evaluate it, and if false, go on. (it is the termination expression, BTW)
        Dim FoundLabel As CFormItem
        On Error Resume Next
        Set FoundLabel = Labelled.Item(currItem.Value)
        If Err.Number <> 0 Then
            ParserError(Me, 5, "Label not defined, """ & currItem.Value & """", , currItem.ExprPos).Throw
        Else
            Set currItem = FoundLabel.Next_
            
        
        
        
        End If
        

        
        
        
End Select
NEXTITERATION:
        Set currItem = currItem.Next_
DIRECTNEXT:
    Loop Until currItem Is Nothing
    On Error Resume Next
    Dim grabit As Variant
    If Not ParseStack.Count = 0 Then
        If IsObject(ParseStack.Peek) Then
            Set grabit = ParseStack.Pop
        Else
            grabit = ParseStack.Pop
        End If
        If Err <> 0 Then
            Err.Clear
            grabit = Null
            
        Else
            If IsObject(grabit) Then
                If TypeOf grabit Is CVariable Then
                    
                    valret = grabit.Value
                Else
                    Set valret = grabit
                End If
            Else
                valret = grabit
            End If
        
        End If
    Else
    valret = 0
    End If 'count=0
    
    'Add the "ANS" last-answer variable now.
    'tada!
    
    If mParentParser Is Nothing Then
        mVariables.Add "ANS", valret
    End If
    'CollapseStack = ParseStack.Pop
    CDebug.Post "CollapseStack has completed executing the expression, " & mvarExpression
    CDebug.Post "The result is:" & ResultToString(valret)
    'If Trim$(ResultToString(valret)) = "5" Then Stop
    CDebug.PopP "CollapseStack"
    'If Trim$(ResultToString(valret)) = "8i" Then Stop
    '(5-(8*i))
    mState = ParseState_Idle
Exit Sub
Err_CollapseStack:
'nothing yet.
    'we should interpret Out of data errors as something more specific.
    'like "a required term was not found"
    'Error: we will create a more context specific error- somehow.
    'for example, if they do something like take the Sqr(-1), then
    'the current value of the curritem variable should be
    'a reference to the IT_FUNCTION type that caused the error.
    'in addition, if they do something illegal with operators- (divide by zero, for example)
    'it will still reference the "\" operator curritem.
    'TODO:// add checking code to make sure it doesn't have the parserErrorBase already.

'changed! now it raised the ParserError.
    Dim errNum As Long
    Dim errdesc As String
    Dim errsrc As String
    
    Dim howret As ParserErrorRecoveryConstants
    Dim rParserError As CParserError
    errsrc = Err.Source
    errdesc = Err.Description
    errNum = Err.Number
    'Resume
    Set rParserError = ModParser.ParserError(Me, errNum, errdesc & " in op " & currItem.op, "CParser::CollapseStack," & errsrc, I)
    howret = rParserError.Throw
    If howret = PERR_RETURN Then
        'return the error, converted via CVErr.
        'CollapseStack = CVErr(errNum)
        Set valret = rParserError
    ElseIf howret = PERR_RESUME Then
        Resume
    ElseIf howret = PERR_IGNORE Then
        Resume Next
    ElseIf howret = PERR_FATAL Then
        'sigh.
        mState = ParseState_Idle
        Err.Raise Err.Number, Err.Source
    End If
        
        
        
    
    
    
    
    
    
    
    
    
    'Err.Raise Err.Number + Exec_GenericExecError, "CParser::CollapseStack", """" & Err.Description & """ in op " & currItem.op
    
End Sub

Private Sub NewItem(mCurrItem As CFormItem, mPrevItem As CFormItem, mnextitem As CFormItem, Currpos As Long, Expression As String)
    
    If Not mCurrItem Is Nothing Then
        Set mPrevItem = mCurrItem
    End If
    If Currpos < Len(Expression) Then
    Set mnextitem = New CFormItem
    Set mCurrItem.Next_ = mnextitem
    Set mnextitem.Prev = mCurrItem
    
    Set mCurrItem = mnextitem
    
    
    Set mPrevItem = mCurrItem.Prev
    
    End If



End Sub

Private Sub GetOperation(ByVal op As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval As Variant, Optional ForceHandler As IEvalEvents)
'call the interface collection. Yes, I know the interface is really complicated-
'but who said it wasn't?
'the fact is, I wrote it, so I should know how to use it.

'Parameters/Arguments:
'Op[in]: Operation to perform.
'OpA[in]: First Operand.
'OpB[in]: Second operand.
'CurrItem[in,out]: The current item in the LinkedList of the parsed expression.
'retval:[out]: value returned. Why don't I use a function? Short answer, because we can return anything, and testing for an object
'to see wether a "Set" statement is required is WAY too ridiculous.
'ForceHandler: I think this is interesting. Forces the use of that specific handler to handle the operation.
'Of course, the client could simply call it itself, but hey, why not encapsulate all the logic here.

    If Not ForceHandler Is Nothing Then
        'OK, this is somewhat important- make sure it actually supports it.
        If Not (ForceHandler.CanHandleOperator(Me, op, UnaryOp_None)) Then
            'Oh! throw a error, it doesn't support the operator they passed.
            'also, remember to include the fact that the operation was forced
            'onto that specific handler.
            ParserError(Me, Exec_UnsupportedOperation, "Operator """ & op & """ not handled by forced Handler """ & ForceHandler.Name & """.", "CParser::GetOperation{Forced}").Throw
        
        End If
    
    Else

    
        Call mEvalListener.GetOperation(Me, op, OpA, OpB, currItem, retval)
    End If


End Sub

Private Sub GetUnaryOperation(ByVal op As String, ByVal Val As Variant, currItem As CFormItem, ByRef Ret As Variant)
    mEvalListener.GetUnaryOperation Me, op, Val, currItem, Ret
End Sub

Private Sub PerformFunction(ByVal cItem As CFormItem, retval As Variant)
    'given a IT_FUNCTION cformItem type, evaluates it.
    'returning a VARIANT as appropriate.
    'Dim Evaluated() As Variant  'evaluated arguments
    Dim Arguments As Variant  'CParser Array. retrieved from the cFormitem.
    Dim PassedItems() As Variant
    Dim originalpassed() As Variant
    Dim NoParsed() As Boolean
    Dim I As Long, Casted As IEvalEventSinks
    Dim LoopListener As IEvalEvents
    Dim castoperable As IOperable
    Dim NameUse As String, LoopPlug As IEvalEvents
    Dim StrFuncName As String
    Dim Useplugin As IEvalEvents, ForcedHandler As String 'set to value if a period was detected and thus a
                                'handler was attempted to force. Name of handler is in it, btw.
    StrFuncName = cItem.op
    Set Useplugin = mEvalListener
    If Not IsEmpty(cItem.Value) Then
        Arguments = cItem.Value
        On Local Error GoTo NOARGUMENTS
        ReDim NoParsed(UBound(Arguments))
        ReDim PassedItems(UBound(Arguments))
    Else
      ReDim Arguments(-1 To -1)
      ReDim NoParsed(-1 To -1)
      ReDim PassedItems(-1 To -1)

    End If
    'added May 03, 2007
    'Realized that certain functions (Sin,Cos,Tan) don't work on Complex numbers, even though the CComplex
    'class includes code to do so.
    'ReDim Evaluated(UBound(Arguments))
    
    
    
    
    
    
    'NOTE: The following block was commented out,
    'because it broke calling functions with more then one argument with UI type calls.
    'For example, "STORE(X,MSGBOX("TEST"))" would show the messagebox twice.
    
    
'    If UBound(Arguments) >= 0 Then
'        'a single argument! woopee!
'        'check it. if it is a object.
'       ' if
'        Arguments(0).ExecuteByRef PassedItems(0)
'
'        'check it. Is it an object?
'
'        'BUG: this is called even with "Lists" which implicitly create a new "ARRAY" function.
'        'while "somewhat" desirable, we should instead handle an error condition by proceeding with default handling, instead of raising an error.
'        If IsObject(PassedItems(0)) Then
'            'continue- it's an object. does it support IOperable, though?
'            If TypeOf PassedItems(0) Is IOperable Then
'                'ha HA!
'                Set castoperable = PassedItems(0)
'                'set error handling.
'                On Error Resume Next
'                originalpassed = PassedItems
'                If castoperable.PerformFunc(Me, StrFuncName, PassedItems, retval) = Perform_Success Then
'                    'OK-
'                    'perform Byref synch check.
'
'                    PropagateChangedArguments Me, cItem.Value, originalpassed, PassedItems
'
'                    'Ha. retval should be assigned to the appropriate return value, so we
'                    'can break out now.
'                    Exit Sub
'
'
'                Else
'                   'Continue with default processing.
'                    CDebug.Post "object type """ & TypeName(PassedItems(0)) & """ doesn't implement special handling for Function """ & StrFuncName & """ continuing default processing."
'
'
'                End If
'
'
'            Else
'            'aww, oh well.
'            'we could do a CallByName/Invoke for it though....
'                    On Error Resume Next
'                    InvokeDynamic PassedItems(0), StrFuncName, PassedItems, retval
'            End If
'
'        End If
'
'    End If
    
    'I thought I was going to have a nasty bug, but no- since the
    'first implementor encountered inside the Collection will be the one who is called.
    
    'New- use of period separator to access Functions in a specific plugin.
    'the name to use will be the returned "Name" value. if that "Name" value has a space, obviously
    'it will not be available for this purpose.
    'this means I need to change the logic for this area.
    'use a different variable- UsePlugin.
    If InStr(1, StrFuncName, ".", vbBinaryCompare) <> 0 Then
        'if a period occurs anywhere within the string, we may need to change the
        'UsePlugin variable. However, let's not get too eager.
        'we of course, first need to name to grab.
        NameUse = Mid$(StrFuncName, 1, InStr(StrFuncName, ".") - 1)
        'alright, now the fun part. Look through our current group of
        'Plugins, inspecting the name and comparing it to NameUse...
        For Each LoopPlug In EvalListener.GetCollection
            If StrComp(NameUse, TypeName(LoopPlug), vbTextCompare) = 0 Then
                'close enough.
                ForcedHandler = NameUse
                CDebug.Post "Forced usage of Plugin Named """ & NameUse & """."
            
                'also- we need to strip the leading name and period from the
                'function name.
                Set Useplugin = LoopPlug
                StrFuncName = Mid$(StrFuncName, InStr(StrFuncName, ".") + 1)
                Exit For
            End If
            
        
        
        Next
    End If
    
        If InStr(1, mDisabledFunctions, " " & StrFuncName & " ", vbTextCompare) > 0 Then
            ParserError(Me, ParserErrorBase + 13, "Function Named " & StrFuncName & " Is Disabled.").Throw
            Exit Sub
        End If
        If Useplugin.CanHandleFunction(Me, StrFuncName, NoParsed()) Then
        'On Error Resume Next

        
       
        
        
        
            For I = 0 To UBound(Arguments)
                If NoParsed(I) = False Then
                    'I was going to optimize here by preventing an execute,
                    'but that wouldn't make a difference. OK, it would,
                    'but i can add this later.
                   ' If StrFuncName = "MID$" Then Stop
                    'assign(PassedItems(I) = Arguments(I).Execute)
                    'don't reparse if there is only one argument.
                    
                    
                    
                    'ARRRRGH! This causes Msgbox, Inputbox, and so forth to ALL be executed TWICE!
                    'If UBound(Arguments) > 0 Then
                        Arguments(I).ExecuteByRef PassedItems(I)
                    'End If
                 
                    
                    
                Else
                    'otherwise, the Listener would like me to NOT to parse it.
                    PassedItems(I) = Arguments(I).Expression
                End If
                'If IsEmpty(PassedItems(I)) Then Stop
            Next I
            
            If UBound(NoParsed) >= 0 Then
            If NoParsed(0) = False Then
            
                'EDIT: moved the Object test code here- I'll just rewrite it for this context, however.
                
                'Is the first parameter an object?
                    If IsObject(PassedItems(0)) Then
                        'OK- it's an object. See if it implements IOperable...
                        If TypeOf PassedItems(0) Is IOperable Then
                            Set castoperable = PassedItems(0)
                            'Dear me...
                            If castoperable.PerformFunc(Me, StrFuncName, PassedItems, retval) = Perform_Success Then
                                'successful!
                                'synch by reference arguments...
                                 PropagateChangedArguments Me, cItem.Value, originalpassed, PassedItems
                                Exit Sub
                            
                            End If
                    
                        End If
                    End If
                
                
                End If
            End If
            
            'End If 'I=0
        
            
            
            originalpassed = PassedItems
                   
                'Call useplugin.BeforeFunctionCall
            Call Useplugin.HandleFunction(Me, StrFuncName, PassedItems, retval)
             PropagateChangedArguments Me, cItem.Value, originalpassed, PassedItems
            'OK, there we go.
            'Stop
        Else
            If ForcedHandler <> "" Then
                'if it is a forced handler, say so.
                ParserError(Me, Exec_UnSupportedFunction, "Function """ & StrFuncName & """ not supported by module """ & ForcedHandler & """.").Throw
            End If
        End If
    Exit Sub
NOARGUMENTS:
    'there weren't any arguments-
    'pass an empty passitems array.
    
    'BUG:
    'found with the SIN() function, but reproducible with any single item function that doesn't perform a fairly adequate check of it's arguments.
    'I thought I had previously coded the ability for the trig functions to be used with an array argument, as in Sin({1,2,3}). Turns out I was wrong, and I think I removed it.
    'in any case, I discovered a assumption that wasn't necessarily true here.
    'the real Sin() function was  then called with an array argument in the HandleFunction() Routine of BPCoreOpFunc, at which point a Type mismatch occured. since this is the last
    'active error handler, it goes here. For some stupid reason, I assumed that if we got here, there weren't enough arguments. kind of silly. but anyway, it went
    'about removing the elements and calling again, causing the opposite extreme, of no arguments.
    'the funny thing is the error handler was really more intended for the loca....
    'just changed to On Local Error. not sure if this will help.
    
    'step one:
    'check for type mistmatch. if we have a type mismatch AND our parameter is an array- it's likely we don't accept array arguments.
    'on the other hand, plugin code should be as robust as possible.
    
    
    If Err.Number = 13 Then
        'if ANY of our arguments are arrays....
        For I = LBound(PassedItems) To UBound(PassedItems)
            If IsArray(PassedItems(I)) Then
                Dim Perr As CParserError
                Set Perr = ParserError(Me, 13, "Function """ & StrFuncName & """ does not appear to support Array arguments, or a Type Mismatch error occured while processing one of it's other arguments.")
                Select Case Perr.Throw
                
                Case PERR_FATAL
                    'Crap. Return vbempty.
                    retval = vbEmpty
                Case PERR_IGNORE, PERR_RESUMENEXT
                    Resume Next
                Case PERR_RESUME
                
                    'not sure what good this would do.
                    Resume
                Case PERR_RETURN
                    'good, good. this one actually makes sense.
                    Set retval = Perr
                    Exit Sub
                
                End Select
            
            End If
        Next I
    
        ParserError Me, 13, "Type Mismatch"
    End If
    
    
    
    Erase PassedItems
    Call Useplugin.HandleFunction(Me, StrFuncName, PassedItems, retval)
    
    'uh oh
    
    
    
'    If IsObject(retval) Then
'        Set PerformFunction = retval
'    Else
'
'        PerformFunction = retval
'    End If
    
    
End Sub

Private Sub Class_Initialize()
'we are are own return formatter.
'in addition, any Clones() created will actually have US as a implementor of THEIR
'returnformatter interface! Pretty nifty, eh?
    mAllowInterrupt = True
    mSetName = "Default"
    Set mReturnFormatter = Me
    UniqueID = GenerateGUID     'create our unique ID.
   Set mCorePlugins = New ICorePlugins       'init to default
   'Set mCorePlugin = mCorePlugins
   CParserCount = CParserCount + 1
    Set mOutputhooks = New OutputHooks
        
End Sub

Private Sub Class_InitProperties()
    'There is only ONE property right now.
    'the expression property.
    mvarExpression = ""
    Set mVariables = New CVariables
    Set mFunctions = New CFunctions
    Set mFunctions.Parser = Me
    
    'Set mEvalListener = New IEvalEventSinks
    
End Sub



Private Sub Class_Terminate()
    ParserSettings.DeinitializeParser Me, mSetName
'    If Not mEvalListener Is Nothing Then
'        mEvalListener.Term Me
'    End If
'the above was removed because a Clone()'d copy of a CParser object
'was destroying the original CParser object instances Collection of EvalEvents Sinks.

    Set mEvalListener = Nothing
    If Not mFirstItem Is Nothing Then mFirstItem.Destroy
    Set mFirstItem = Nothing
    CParserCount = CParserCount - 1
    If CParserCount = 0 Then
        CDebug.Post "Last CParser is being Terminated.", Severity_Information
    
    End If
End Sub



Private Sub IReturnFormatter_FormatResult(Value As Variant)
    'TODO:// add different properties to
    'the CParser as to how it implements it's Return Formatter class.
    'at the very least, have different formats similiar to the TI-83.
    'don't ask why.
    
    'default behaviour formats a few miscellaneous objects and so forth.
    
End Sub

Private Sub mFunctions_FunctionAdded(FuncAdded As CFunction)
    'when a function is added to this object, we create a new variable that is a reference
    'to the added CItem, and whose name is the name of the function.
    'this is SUPER neat-o.
    mVariables.Add FuncAdded.Name, FuncAdded
        
End Sub

Private Sub mFunctions_FunctionRemoved(FuncRemoved As CFunction)
    'hmm, I wonder what this does :)
    mVariables.Remove FuncRemoved.Name
End Sub

Private Sub mVariables_GetOwner(ParserGet As CParser)
    Set ParserGet = Me
End Sub

Private Sub mVariables_VarAdded(ByVal VarAdded As CVariable)
'add a variable.
If Not mParentParser Is Nothing Then Exit Sub
CDebug.Post "Variable added," & VarAdded.Name
End Sub

Private Sub mVariables_VarChanged(ByVal VarChanged As CVariable, ByVal OldValue As Variant)
'changed a variable.
'WE don't care about the argument.
If Not mParentParser Is Nothing Then Exit Sub
    CDebug.Post "Variable " & VarChanged.Name & " Has changed."
    Call mEvalListener.VariableChanged(VarChanged.Name, OldValue, VarChanged.Value)
End Sub

Private Sub mVariables_VarReplace(Replacing As CVariable, Allow As Boolean)
'
Allow = True
End Sub

Private Sub mVariables_VarRetrieved(ByVal VarRetrieved As CVariable)
'retrieved a variable.
If Not mParentParser Is Nothing Then Exit Sub
    Call mEvalListener.AfterResolveVariable(Me, VarRetrieved.Name, VarRetrieved.Value)

End Sub

Private Sub UnarySink_PushItem(ValPushed As Variant, ByVal ItemKey As String)
    'HA! a unary operation must be performed on this value.
    'simply call the mevalevents.
    'however, all is not well.
    'we must loop until we empty the stack, performing the unary operation on each item.
    'the most recent unary operator first, of course.
    Dim returnValue As Variant
    Dim NoEvent As DataStack, poppedunary
    Dim PopCast As CFormItem
    Do Until mUnaryCache.Count = 0
        Set poppedunary = mUnaryCache.Pop
        Set PopCast = poppedunary
            'we need to explicitly cast it to a cformitem type
            'or else VB will whine in the next function call, here.
          
        Call EvalListener.Self.GetUnaryOperation(Me, poppedunary.op, ValPushed, PopCast, ValPushed)
    Loop
   
        Set UnarySink = Nothing
 
    'NoEvent.Push ValPushed
    
    
End Sub

Friend Function SetInitialized()
    mInit = True
End Function

Friend Function GetSinker() As IEvalEventSinks
    Set GetSinker = mEvalListener

End Function

Friend Function BuildStack_Infix(ByRef Token As CFormItem) As CFormItem
'rearranges the Stack created by RipFormula so that CollapseStack can handle it.

    ' rearrangement schema:
    ' - always push literals
    ' - do NOT push operators but put them on a separate stack
    ' - if an operator's priority is inferior to the one of the
    ' last operator on the operator stack then push all
    ' operators on the operator stack onto the parse stack
    ' and put the current operator on the operator stack
    ' - at brackets call this procedure recursively
    ' - finally push all remaining commands
    Dim operators As COpStack
    Dim CurrToken As CFormItem
    Dim lastprior As Long
    Dim MakeStack As CFormItem
    Dim storeUnary As CFormItem
    Dim pFirstItem As CFormItem, mPrevItem As CFormItem
    Set MakeStack = New CFormItem
    Set pFirstItem = MakeStack
    
    Set operators = New COpStack
    Set CurrToken = Token
    lastprior = -1  'default to -1.
    Do Until CurrToken Is Nothing
        Set mPrevItem = CurrToken       'constant- won't change even after appenditem
        Select Case CurrToken.ItemType
            
                
            Case ItemTypeConstants.IT_VALUE
                'CDebug.Post "IT_VALUE:" & CurrToken.Value
                MakeStack.ItemType = CurrToken.ItemType
                If IsObject(CurrToken.Value) Then
                    Set MakeStack.Value = CurrToken.Value
                Else
                    MakeStack.Value = CurrToken.Value
                End If
                MakeStack.op = CurrToken.op
                Call Assign(MakeStack.Extra, CurrToken.Extra)
                Call Assign(MakeStack.Tag, CurrToken.Tag)
                
                'Set MakeStack.Next_ = New cFormitem
                'Set MakeStack = MakeStack.Next_
                AppendItem MakeStack
            
            Case ItemTypeConstants.IT_VARIABLE
                CDebug.Post "IT_VARIABLE:" & CurrToken.op
                MakeStack.ItemType = IT_VARIABLE
                MakeStack.Value = CurrToken.Value
                MakeStack.op = CurrToken.op
                AppendItem MakeStack
            
            Case ItemTypeConstants.IT_CLOSEBRACKET
                'Exit Do
                'like I said before. IT_CLOSEBRACKET no longer exists.
                'if we somehow encounter it, then, by golly.
                'humm-
                'than it is a very curious occurence indeed.
                'Stop
                'actually, it can occur quite innocently if the string was simply malformed.
                'in this case, raise a malformed Expression error.
                ParserError(Me, ParseErrors.Parse_MalformedExpression, "Malformed Expression- Closing bracket without starting bracket.").Throw
            Case ItemTypeConstants.IT_FUNCTION
            'ha, no extra work here, acts like a IT_VALUE.
            'although, currtoken.value is really an array of CParser objects.
            
                MakeStack.ItemType = IT_FUNCTION
                MakeStack.Value = CurrToken.Value
                MakeStack.op = CurrToken.op
                'Set MakeStack.Next_ = New cFormitem
                'Set MakeStack = MakeStack.Next_
                AppendItem MakeStack
            Case ItemTypeConstants.IT_OPENBRACKET
                MakeStack.ItemType = IT_SUBEXPRESSION
                Set MakeStack.Value = CurrToken.Value
                MakeStack.op = CurrToken.op
                
                AppendItem MakeStack
                
            Case IT_ARRAYACCESS
                'simply do as we do with everything else.
                CDebug.Post "IT_ARRAYACCESS"
                MakeStack.ItemType = IT_ARRAYACCESS
                MakeStack.Value = CurrToken.Value
                MakeStack.op = CurrToken.op
                'Set MakeStack.Next_ = New cFormitem
                'Set MakeStack = MakeStack.Next_
                AppendItem MakeStack
                Set storeUnary = Nothing
            Case IT_ObjAccess
                CDebug.Post "IT_ObjAccess"
                MakeStack.ItemType = IT_ObjAccess
                MakeStack.Value = CurrToken.Value
                MakeStack.op = CurrToken.op
                MakeStack.Tag = CurrToken.Tag
                AppendItem MakeStack
                Set storeUnary = Nothing
            Case IT_OPUNARY
                'a unary operation must be immediately AFTER the next item.
                'as such, we store it, temporarily, until the next item rolls around.
                'Set storeUnary = CurrToken
                CDebug.Post "IT_OPUNARY"
                MakeStack.ItemType = IT_OPUNARY
                MakeStack.op = CurrToken.op
                MakeStack.Value = CurrToken.Value
                MakeStack.Tag = CurrToken.Tag
                AppendItem MakeStack
            Case ItemTypeConstants.IT_OPERATOR
                'operator logic.
                Dim priority As Long
                'get the priority of this operator.
                priority = GetPriority(CurrToken.op)
                
                If priority < lastprior Then
                    Do While CBool(operators.Count)
                        MakeStack.ItemType = IT_OPERATOR
                        MakeStack.op = operators.Pop
                        'Set MakeStack.Next_ = New cFormitem
                        'Set MakeStack = MakeStack.Next_
                        AppendItem MakeStack
                    Loop ' While CBool(Operators.Count)
                ElseIf priority = lastprior Then
                    MakeStack.ItemType = IT_OPERATOR
                    MakeStack.op = operators.Pop
                    'Set MakeStack.Next_ = New cFormitem
                    'Set MakeStack = MakeStack.Next_
                    AppendItem MakeStack
                
                
                End If ' Priority = LastPrior
                
                Call operators.Push(CurrToken.op) 'push the operator.
                lastprior = priority
                Set storeUnary = Nothing
            Case Else
                'any other item is simply placed as if it where a value.
                
                MakeStack.ItemType = CurrToken.ItemType
                If IsObject(CurrToken.Value) Then
                    Set MakeStack.Value = CurrToken.Value
                Else
                    MakeStack.Value = CurrToken.Value
                End If
                MakeStack.op = CurrToken.op
                Call Assign(MakeStack.Extra, CurrToken.Extra)
                Call Assign(MakeStack.Tag, CurrToken.Tag)
                
                'Set MakeStack.Next_ = New cFormitem
                'Set MakeStack = MakeStack.Next_
                AppendItem MakeStack
                
                
    
            End Select
'         If CurrToken.ItemType <> IT_OPUNARY And Not (storeUnary Is Nothing) And CurrToken.ItemType <> 0 Then
'            'now, we can place this one onto it.
'            'that is, as currtoken.next:
'            'there is a important reason for mPrevItem.
'
'          '  Set mPrevItem.Next_ = storeUnary
'          '  Set mPrevItem.Next_.Next_ = New cFormitem
'
'            'Set mPrevItem = MakeStack.Next_.Next_
'            'currtoken shouldn't be an operator, either.
'            'I mean, -+, what?
'            Select Case CurrToken.ItemType
'                Case Is <> IT_Operator
'                    With MakeStack
'                        .ItemType = IT_OPUNARY
'                        .Value = storeUnary.Value
'                        .Op = storeUnary.Op
'                    End With
'                    'tada
'                    AppendItem MakeStack
'                    Set storeUnary = Nothing        'delete the flag object.
'
'            End Select
'
'        End If
        Set CurrToken = CurrToken.Next_
    Loop
    
     Do While CBool(operators.Count)
        MakeStack.ItemType = IT_OPERATOR
        MakeStack.op = operators.Pop
        Set MakeStack.Next_ = New CFormItem
        Set MakeStack = MakeStack.Next_
    Loop
    
    
    
Set BuildStack_Infix = pFirstItem

End Function

Friend Function ParamListToParserArr(ParamList() As String) As Variant
    Dim CreateArray As Variant, I As Long
    ReDim CreateArray(LBound(ParamList) To UBound(ParamList))
    For I = LBound(ParamList) To UBound(ParamList)
    'create a new Cparser for that location.
    'we want it to have access to the same variables and such as this one- so we clone ourself.
        Set CreateArray(I) = Me.Clone
        
        CreateArray(I).Expression = ParamList(I)
        
    Next I

    ParamListToParserArr = CreateArray


End Function
Private Sub FindLabels(InExpression As String)






End Sub
Friend Function RipFormula(ByVal Expression As String) As CFormItem

    'RipFormula- converts an InFix expression into
    'a un-fixed list.
    'BuildStack_Infix should be run on the resulting linked list.

    'returns the first item in a Ripped Formula. the returned
    'value is the first cFormItem in the formula list.
    'all operators within the expression must be delimited by spaces.
    'IE:
    '5 + 2 + "HI I'm a string"
    'the ParseInFix calls MakeParsable to do this.
    
    
    'NOTES:
    'I just profiled this code with a large number of nested parentheses, a Seq call, the works.
    'turns out that this function was being executed longer, and was also
    'called less- go figure!
    
    'TODO:\\
    'segregate the code that actually looks at the token
    'from the code that adds items to the stack.
    'that is, make it possible to easily implement Prefix And PostFix ripping
    'as well as InFix. right now, it isn't really high on the list of priorities, since
    'not many people will use RPN directly. If I release this code and people request it though,
    'I imagine I'll be forced to add it.
    
    'WAIT: I got an idea- right now RipFormula performs a fair amount of work. If I want to implement RPN, I could do
    'many things:
    
    'Implements a RipRPN function. The thing is, this requires a way to tell the parser that it has a RPN
    
    'implement the above, but it only get's called implicitly via some implicit
    'operator. (I don't know what to use, I'm out of punctuation marks!- I could use some silly RPN() function-
    
    
    'don't support RPN. This is pretty easy.
    
    'TODO:\\ This function is executed around 90% of the time. Optimizations in this procedure
    'or the procedures it calls will likely speed up the program by considerable amounts.
        
        
        
   
    Dim currItem    As CFormItem    ' current command item
    Dim ParseStack  As DataStack  ' numeric result stack
    Dim IsObjAccess As Boolean
    Dim Currpos As Long
    Dim Char As String
    Dim mCurrItem As CFormItem
    Dim tempeval As Variant
    Dim holdstack As CFormItem
    Dim FirstItem As CFormItem
    Dim mnextitem As CFormItem
    Dim mPrevItem As CFormItem
    Dim OpUse As String
    Dim ExprLoop As Variant
    Dim makeitem As SubExprOData
    Dim StatementStart As Long
    
    
    
    
    StatementStart = 1  'initializes to the first character.
    Currpos = 1
         
    Dim OptimizeSubstrs As Collection
   Set OptimizeSubstrs = New Collection
    'On Error GoTo RiPError
    'if we are constant, then we don't need to rip.
    If mIsConstant And Not FirstItem Is Nothing Then
        Set RipFormula = FirstItem
        
        Exit Function
    End If
    Set FirstItem = New CFormItem
    
    Set mCurrItem = FirstItem
    'tell the class about it's new CFormitem.
    'Set mFirstItem = NewItem
    mEvalListener.PreParse Me, Expression
    CDebug.PushP "RipFormula", MODULE_NAME, MODULE_FILE_NAME, 0, "Parsed Expression:" & Expression
    CDebug.Post "Ripping formula:" & Expression
    'NEW CODE: --5/25/2007 @ 17:32--
                'optimize substrings.
                'steps (for this stage of parsing)
    'build the table of fix-ups to apply to the expression.
    'the following psuedocode:
    'iterate through the entire collection of cached expressions.
    'with each expression...
    '       determine if this expression is contained within the expression being parsed as a sub-string.
    '       if so...
    '       store away the location that the substring occurs as well as the Stack that should be inserted when we get there.
    'next expression
    'On Error Resume Next
    
    #If ENABLESUBSTROPTIMIZATION Then
        For Each holdstack In OptimizedStacks.Collection
        If holdstack Is Nothing Then Exit For       'kludge....
            If InStr(1, Expression, holdstack.Extra, vbTextCompare) <> 0 Then
                'woah! we found a match!
                'place it in optimizesubstrs.
                'firstly, make sure the item actually exists.
                If Err <> 0 Then Exit For
                On Error Resume Next
    
                
                If Err.Number = 0 Then
                    makeitem.Position = InStr(1, Expression, holdstack.Extra, vbTextCompare)
                    Set makeitem.StackItem = holdstack
                    
                    OptimizeSubstrs.Add makeitem, "LOC" & Trim$(str$(makeitem.Position))
                End If
            End If
    
        Next
    #End If
    '5+2*3/2+2/3
    
    
    '------------
    Do Until Currpos > Len(Expression)
    'can't seem to get sub-string optimization to work properly.
    
        #If ENABLESUBSTROPTIMIZATION Then
            On Error Resume Next
             
            makeitem = OptimizeSubstrs.Item("LOC" & Trim$(str$(Currpos)))
           
            If Err = 0 Then
                'ha HA!
                'insert the stack right here.
                'This is the tricky part, to be honest.
                'First off, it doesn't seem to work- I think it is being inserted improperly.
                
                'EUREKA! of course!
                'the end of the inserted stack will have IT_ENDOFLIST, which will cause the
                'Execution logic to break out of it's loop!
                'as such, we must remove the last item of the stack.
                'get the last item, and then set it's previous item's next to nothing.
                'Oh, but only if it actually IS IT_ENDOFLIST. It should be.
                If makeitem.StackItem.GetLast.ItemType = IT_ENDOFLIST Or makeitem.StackItem.GetLast.ItemType = IT_NULL Then
                    Set makeitem.StackItem.GetLast.Prev.Next_ = Nothing
                End If
                If mCurrItem Is Nothing Then
                    mFirstItem = makeitem.StackItem
                    Set mCurrItem = mFirstItem.GetLast
                    Set mPrevItem = mCurrItem.Prev
                    NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
                Else
                    'firstitem
                    'itemn
                    'mcurritem
                    '
                    
                    mCurrItem.InsertAfter makeitem.StackItem
                    Set mCurrItem.Next_.Prev = mCurrItem.Prev
                    Set mCurrItem.Prev.Next_ = mCurrItem.Next_
                    Set mCurrItem = mCurrItem.GetLast
                    Set mPrevItem = mCurrItem.Prev
                    NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
    
                End If
                GoTo NEXTITERATION
            Else
                On Error GoTo 0
            End If
        #End If

        '
        'build the CFormItem Linked list.
        
        
        
        Char = Mid$(Expression, Currpos, 1)
        'If Char = "%" Then Stop
       ' Debug.Assert Not InStr(Expression, "Script")

Select Case True
   
        
    Case mCorePlugin.ParseLocation(Me, Expression, Currpos, mCurrItem) = ParseLocation_recognized
      
            'if core plugins are enabled, then
            ' we will ask our ICorePlugin implementor to
            'see if it can subvert us.
            'If mCorePlugin.ParseLocation(Me, Expression, CurrPos, mCurrItem) = ParseLocation_recognized Then
            
            'the Implementor is responsible for storing something in the extra property of the formitem
            'to be able to quickly know that it is designed for it.
            'well, actually, it can set the itemtype to ANY negative value...
            'it has already assigned appropriate values in mCurrItem for when we call
            'it's handleitem procedure in collapseStack, so we can just continue.
            'Also notice that the CurrPos Parameter is purposely passed By reference, so that
            'the plugin can tell us how long the token is that it recognized, and thus where we should continue.
            CDebug.Post "The ICorePlugin Implementor recognized a token.", Severity_Information
             NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
            
            
            'End If
        
        
        Dim ParsedVal As String
        
    
    
    Case Char = " "

    
    Case Char = "("
       
        
    
        Dim SubExpression As String
        
        
        'implicit multiplication
        'if the VERY previous operator was a IT_SUBEXPRESSION type, then add a multiplication type between it and this new item.
       
        'compatible previous items for implicit multiplication:
        'IT_VALUE,IT_SUBEXPRESSION,IT_VARIABLE
        If Not mPrevItem Is Nothing Then
            Select Case mPrevItem.ItemType
                Case IT_SUBEXPRESSION, IT_VALUE, IT_VARIABLE, IT_FUNCTION
                    'insert a Multiplication operator here
                    With mCurrItem
                        .op = "*"
                        .Value = .op
                        .ItemType = IT_OPERATOR
                        .ExprPos = Currpos
                    End With
                    NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
            End Select
        
        End If
        mCurrItem.ItemType = IT_OPENBRACKET
        mCurrItem.op = Char
        mCurrItem.ExprPos = Currpos
        'the value will be the bracket contents, assigned to another CParser.
        
        
        'THIS IS OLD: don't listen to it! I keep it for, err- historical reasons.
        'Complex number Support has been implemented much easier since this implementation.
        'COMPLEX NUMBER SUPPORT:
        'I have made a compromise. instead of
        'forcing the use of the COMPLEX() function, I will check to see
        'if the interior of parentheses is a complex number.
        'a complex number will have a single + or minus, and an i at the end.
        'actually, they all have an I on the end. but we cannot find multiple operators afterward.
        
        
        'TODO:\\ Implicit multiplication support.
        'that is, check here for a previous item that is IT_SUBEXPRESSION (yes, a (- it includes the ).
        
        Set mCurrItem.Value = Me.Clone
        SubExpression = ExtractBrackets(Expression, Currpos)
        CDebug.Post "RIPFORMULA: create SubParser for SubExpression, """ & SubExpression & """."
        mCurrItem.Value.SetExpression SubExpression
        'we will treat the entire bracket set as a single object. this means that, technically, there is
        'no IT_CLOSEBRACKET, because we prevent the ripformula from ever encountering it.
        'as such, move to the end of the extraction. the extraction doesn't include the parentheses, so don't forget to
        'move past those.
        Dim tempval As Variant
        On Error GoTo 0
        'mCurrItem.Value.ExecuteByRef tempval
        'CDebug.Post "RIPFORMULA: created sub parser for expression, " & """" & SubExpression & """." & " RESULT was:" & ResultToString(tempval)
        
        Currpos = Currpos + Len(SubExpression) + 1 'ok, I lied. I only add one, thanks to the loop adding another. ;)
        NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
    
    Case Char = ")"
       ' Debug.Assert False
        'OK, this itemtype no longer exists- all
        'Bracketed expressions are dealt with in a separate CParser.
        'Now, that Parser is optimized, just like we are, but only when the CollapseStack function
        'encounters the IT_SUBEXPRESSION (equal to IT_OPENBRACKET) in
        'the formula list. It then Parses and executes it, using the returned value.
        'if we get here, then, we have a problem with the expression.
        'a ")" without opening "("
        
        
        'as such, create and throw a new CParserError Object.
        On Error GoTo 0
        'Err.Raise Parse_Badsyntax, "CParser::RipFormula", "Unexpected use of "")""."
        Select Case ParserError(Me, Parse_Badsyntax, "Unexpected use of "")"".", "CParser::RipFormula").Throw
        
        
        Case PERR_FATAL
        Case PERR_IGNORE
        Case PERR_RESUME
        Case PERR_RESUMENEXT
        Case PERR_RETURN
            'return back to the client. OK- I guess so.
        
        End Select
        mCurrItem.ItemType = IT_CLOSEBRACKET
        mCurrItem.op = Char
       
        NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
    Case Char = STRLISTOPEN
        ''the list open character.
        'Extract the arguments. Then, replace the expression with a new one, where the list reference
        'is a ARRAY call.
        
        
        
        
        
        Dim ListContents As String
        ListContents = ExtractBrackets(Expression, Currpos, "{", "}")
        'create a new IT_FUNCTION Item.
        'the function? well, simple, our friend- the ARRAY function, handled by the BASeParser CORE functions/operators class.
        mCurrItem.ItemType = IT_FUNCTION
        mCurrItem.op = "ARRAY"
        mCurrItem.ExprPos = Currpos
        'it expects an array of CParser objects in it's
        'Value property. of course, if the array is empty, there will be no arguments.
        On Error Resume Next
        mCurrItem.Value = ParamListToParserArr(ParseArguments(ListContents))
        If Err <> 0 Then
            'create a empty array.
            ReDim tempval(0)
            Erase tempval
            mCurrItem.Value = tempval
        
        End If
        NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
        Currpos = Currpos + Len(ListContents) + 1
        
        'There. Simple enough.
        
        
        
    Case Char = "@"
        'object access operator.
        'the previous item should have evaluated to a object, and, to the right,
        'we have what would appear to be a function call.
        'we must parse out this call, and store it as the items value (op will of course be @,
        'the key here is to find the end of the next VALUE- the parseValue function comes to mind.
        'However, ParseValue will not properly retrieve a entire method call, rather,
        'it will retrieve up to and including the first bracket.
        'as such, we perform the same testing as a function
        'so we set the flag, IsObjAccess. when this loop finds the next character, it should properly extract
        'either the property or method call. However, it would normally call it a function. here we must
        'set the flag.
        'however, if the flag is already set, we encountered
        'two @'s, or something screwy occured.
        
        '--6/15/2007 @ 18:33--
        'I have had a lot of fun in previous debug sessions with the Object access operator.
        
        
        
        
        
        If IsObjAccess Then
            'sigh....
            'we might be ok....
            'we could raise an error, like, "Unexpected Object Access Operator"
            'This likely occurs when there are two @ operators immediately after each other.
            'likely a typo on the part of the client or user.
            'Generally we can ignore it, but we'll place a warning in the log.
            CDebug.Post "Object Access operator found while a previous Object access operator was looking for the object's member name.", Severity_Warning
        End If
        IsObjAccess = True
    Case Char = STATEMENTSEPARATOR
        'if we encounter a statement separator, then we need to do a few things:
        'First, we need to grab everything from StatementStart to the current location,
        'create a new Parser Object (Cloneing this one), and assign it that "statement"
        'reset statementstart to the current location +1, and then continue.
        'the cFormItem, of course, becomes an IT_STATEMENT type.
        
        
'        Dim StrGrab As String, newchild As CParser
'        StrGrab = Mid$(Expression, StatementStart, CurrPos - StatementStart)
'        Set newchild = Me.Clone
'        'multiparsers.Add NewChild
'        newchild.Expression = StrGrab
'        mCurrItem.ItemType = IT_STATEMENT
'        'NewChild.Expression = Statements(I)
'        'gawd! that is a lot of statements to add a friggin' item.
'        currItem.op = newchild.Expression
'        Set mCurrItem.Value = newchild
'        Set mCurrItem.Next_ = New CFormItem
'        Set mCurrItem.Next_.Prev = mCurrItem
'        Set mCurrItem = mCurrItem.Next_
'        'reset statementstart...
'        StatementStart = CurrPos + 1
    Case Char = "["
        'Array Access Operator.
        'use the ExtractBrackets Function, with optional params to get this
        'entry.
        
        
        
        CDebug.Post "IT_ARRACCESS"
        Dim subscripts() As String, I As Long
        Dim SubScript As String
        Dim Parsers As Variant
        ReDim Parsers(0)
        mCurrItem.ItemType = IT_ARRAYACCESS
        mCurrItem.op = "["
        
        SubScript = ExtractBrackets(Expression, Currpos, "[", "]")
        'gotcha- now we parse out the arguments, and, once again, create an array of parsers.
        
        'UPDATE:
        'changed some code here: why? well, within the plugins HandleSubScript interface routines it should be
        'easy to find empty array access subscripts. however, due to the failure to recognize the special case
        'of an empty string here, it was recieving a single empty argument.
        If Trim$(SubScript) = "" Then   'any subscripts at all?
            Erase Parsers
            Erase subscripts
            
        Else
        
            subscripts = ParseArguments(SubScript)
            
            ReDim Parsers(LBound(subscripts) To UBound(subscripts))
            For I = LBound(subscripts) To UBound(subscripts)
                Set Parsers(I) = Me.Clone
                Parsers(I).SetExpression subscripts(I)
            
            Next I
        End If
        'assign it the array.
        mCurrItem.Value = Parsers
        'mCurrItem.Value.SetExpression subscript
        'also, move the position pointer.
         mCurrItem.ExprPos = Currpos
        Currpos = Currpos + Len(SubScript) + 1
        NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
        
    Case GetOperator(Expression, Currpos) <> "" And Not IsObjAccess
        'HEE HEE! added: the Not IsObjAccess condition, so we
        'don't fall into this Operator portion when we are
        'processing an object method.
        'There is an operator.
        'add it to the stack.
        'first, we need to check if it is a unary operator.
        'if so, we add it as IT_UNARY.
        'a normal op has a expression on both sides-
        'and we have those funny Binary operators that are more of a kludge to support ternary operators,
        'such as the : and ? operators.
        Dim getop As String
        Dim FindUnary As Boolean
        Dim mcachesuffix As Boolean, mcacheprefix As Boolean
       OpUse = GetOperator(Expression, Currpos)
       mcachesuffix = mEvalListener.CanHandleOperator(Me, OpUse, UnaryOp_Suffix)
       mcacheprefix = (mEvalListener.CanHandleOperator(Me, OpUse, UnaryOp_Prefix))
        If mcachesuffix And Not mcacheprefix Then
            'is a suffix-
            'MUST follow a value.
            If mCurrItem Is FirstItem Then
                'gasp. a suffix operator in the BAD PLACE.
                Debug.Assert False
                Err.Raise Parse_Badsyntax, "CParser::RipFormula", "Bad syntax in expression- Unary Suffix operator """ & OpUse & """ without preceding value to operate on."
            Else
            'whew- well, at least we have the previous value-
            'since the value will have been parsed and pushed into the stack by the time
            'we get to the operator in CollapseStack, we'll have no need for the current
            'method of hacking around with events (which actually works quite elegantly).
            mCurrItem.ItemType = IT_OPUNARY
            mCurrItem.Tag = "SUFFIX"
            mCurrItem.op = OpUse
            mCurrItem.ExprPos = Currpos
            'ta-da.
            FindUnary = True
            End If
        
        ElseIf mcachesuffix Then
            'they are both true, decide which one to use.
            If mCurrItem Is FirstItem Or mCurrItem.Prev Is Nothing Then
                'well, it can't be a postfix/suffix operator.
                FindUnary = True
                mCurrItem.Tag = "PREFIX"
                mCurrItem.ItemType = IT_OPUNARY
                mCurrItem.op = OpUse
                mCurrItem.ExprPos = Currpos
            
            Else
            
                FindUnary = True
                mCurrItem.Tag = "SUFFIX"
                mCurrItem.ItemType = IT_OPUNARY
                mCurrItem.op = OpUse
                mCurrItem.ExprPos = Currpos
            End If
        
     

      
        Else
        
            'MODIFICATION:

            'if there was a previous item, we must ensure that it is
            'NOT an value, since that would imply that
            'this op is for two operands.
            If Not mPrevItem Is Nothing Then
            Select Case mPrevItem.ItemType

                Case IT_VALUE, IT_FUNCTION, IT_OPENBRACKET, IT_VARIABLE
                    FindUnary = False
                Case IT_OPERATOR, IT_OPUNARY
                    FindUnary = True



            End Select
            Else
                FindUnary = True
            
            End If
        End If
        
        Dim TempOpItem As CFormItem
        
        
        If FindUnary Then
        'we discovered it should be unary.
        
        'perform a check HERE, rather then collapsestack, for parse errors.
        'what kind of parse errors? well, say, a operator that isn't unary being used
        'in a unary context.
            If Not mEvalListener.CanHandleOperator(Me, OpUse, UnaryOp_Prefix) Then
            On Error GoTo 0
                ParserError(Me, Parse_MalformedExpression, "Operator """ & OpUse & """ is not a Unary operator.").Throw
            End If
        
            mCurrItem.ItemType = IT_OPUNARY
        Else
            mCurrItem.ItemType = IT_OPERATOR
        End If
        'UPDATE: OK, I want to extend the parse support so that
        'such things as an "assignment" operator could be added.
        'in order for this to work, however, it needs to know the previous and next item.
        
        'In CollapseStack we'll have easy access to the
        'Next Item- however, we will have no such access to the previous item.
        'which we need to inspect for a variable reference.
        'at some point, to, I'll need to add some god-forsaken support
        'for Property Let's.


      
            
        mCurrItem.op = OpUse
        mCurrItem.ExprPos = Currpos
        Currpos = Currpos + (Len(OpUse))
        NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
        'not sure why this is a compiler directive-
        'should be a global setting stored with the ParserSettings class.
        'Actually, why not have both :)!
        
    
    
    Case Store(ParseValue(Expression, Currpos, IsObjAccess), ParsedVal) <> ""
        'we found a token.
        'CHANGE:
        'variables and function names containing characters that make up an operator
        'where not getting through.
        'since ParseValue will
        
 
        
'        ParsedVal = ParseValue(Expression, CurrPos, IsObjAccess)
       'Stop
       mCurrItem.ExprPos = Currpos
       'check for a date literal, somewhere in this mess...
       
       
       
       If right$(ParsedVal, 1) = "(" Then
        ParsedVal = Mid$(ParsedVal, 1, Len(ParsedVal) - 1)
        End If
        'Now, this could be a variable as well.
        'as such, we should check to see if it is.
        'Oh yeah, and if it is NOT numeric, then we are to
        'determine if it is a function or a variable, and act appropriately.
        
        If IsNumeric(ParsedVal) And ParsedVal <> "" Then
            'create a new IT_VALUE value.
      
            'to do so, we need to plop the number and then a * onto the stack.
            
            mCurrItem.ItemType = IT_VALUE
            mCurrItem.Value = Val(ParsedVal)
            
            NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
            Currpos = Currpos + Len((ParsedVal)) - 1
'            If right$(ParsedVal, 1) = "(" Then
'                'if there is a parentheses following it, create a new item, the Multiplication one.
'                mCurrItem.ItemType = IT_OPERATOR
'                mCurrItem.Value = "*"
'                mCurrItem.op = "*"
'                'we need to move back one to pick up the bracket we find here.
'                mCurrItem.ExprPos = CurrPos - 1
'                CurrPos = CurrPos - 1
'                NewItem mCurrItem, mPrevItem, mnextitem, CurrPos, Expression
''
'            End If
            
            
        ElseIf left$(ParsedVal, 1) = "#" And right$(ParsedVal, 1) = "#" Then
            'it's a date literal.
            mCurrItem.ItemType = IT_VALUE
            mCurrItem.Value = DateValue(Mid$(ParsedVal, 2, Len(ParsedVal) - 2))
            Currpos = Currpos + Len(RTrim$(ParsedVal)) - 1
            NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
        Else
        'not numeric. Either a variable,String, or otherwise.
        If left$(ParsedVal, 1) = """" Then
            
            'ha a quote.
            mCurrItem.ItemType = IT_VALUE
                 
            'we don't want the actual quotes in the value.
            mCurrItem.Value = Mid$(ParsedVal, 2, Len(ParsedVal) - 2)
            Currpos = Currpos + Len(ParsedVal) - 1
            mCurrItem.op = """"
            NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
        Else
            Dim createParser As CParser, grabargs() As String
            Dim looparg As Long
            'variable or function. if the name is immediately followed by "(" then it
            'is a function. otherwise, assume it is a variable.
            If Mid$(Expression, Currpos + Len(ParsedVal), 1) = "(" Then
                Dim argumentList As String, Arguments() As String
            
                'ha HA!
                'the key here is to make sure that the other logic excludes the parameter
                'list of this function. We do this by setting the current location in the string
                'to be that character immediately after the closing ")". Of course, if we don't find it,
                'then we have a "mismatched parentheses error"
                
                'We also need to do a little more work if we find a period or a "@" within a function name. Since
                'the other logic will detect a @ as a Object access,  we are faced with one scenario: a Module.Function notation.
                'Select Case True
                '    Case InStr(ParsedVal, "@") <> 0 Or InStr(ParsedVal, ".") <> 0
                    
                    
                'End Select
                
                
                
                
                argumentList = ExtractBrackets(Expression, Currpos + Len(ParsedVal))
                mCurrItem.ItemType = IT_FUNCTION
                'possible additional optimization:
                'optimize function calls on constant expressions.
                'this involves telling the parser to parse, then actually
                'retrieving the Formula stack from it.
                'then, it should determine wether all items within it are constant.
                
                
                If IsObjAccess Then
                    'an object call, huh.
                    'well, specify it as such.
                    'there, and now Op will be the method,
                    'and Value will be the parameters.
                    '(IT_OBJACCESS is exclusive for @)
                    mCurrItem.ItemType = IT_ObjAccess
                    
                End If
                'the curritem.value value will be an array, that is a group of CParser objects.
                'this is what SetExpression Let is for.
                
                'mCurrItem.Value = argumentList
                'If parsedval = "QUIT" Then Stop
                Arguments = ParseArguments(argumentList)
                'for each of the args, we clone ourself, and then give the new parser
                'a new expression directly, ready to parse when first used.
                'This is somewhat interesting, since it means that if any of the argument expression
                'has its stack already cached, it will optimize automatically.
                On Error Resume Next
                ReDim Parsers(0 To UBound(Arguments))
                If argumentList <> "" Then
                    For looparg = 0 To UBound(Arguments)
                        If Err <> 0 Then Exit For
                        Set Parsers(looparg) = Me.Clone
                        'assign the expression directly.
                        Parsers(looparg).SetExpression Arguments(looparg)
                        'Keep in mind that our stack will have no real mention of the actual arguments
                        'if this function, just the fact that there is a function there.
                        Err.Clear
            
                    Next looparg
                    If Err = 0 Then
                        mCurrItem.Value = Parsers
                    End If
                End If
                Err.Clear
                'On Error GoTo RiPError
        
                mCurrItem.op = Trim$(ParsedVal)     'Op is the function name
                'lastly, set the pointer to exclude the extracted brackets.
                'we want to set it to the actual position of the ), because of the addition of one at the loops end.
                If Not IsObjAccess Then
                    'object methods cannot be optimized, because you can never know what pops out.
                    'if it isn't an object access, we should now determine via interface method calls
                    'wether the function can be optimized. If so, and if the IT_FUNCTION type we created turns out to be constant,
                    'evaluate the expression and change mcurritem into a IT_VALUE with that value.
                    'If IsFormItemConstant(mCurrItem) Then
                    If IsFormItemConstant(mCurrItem) Then
                         CDebug.Post "Optimized function call for " & mCurrItem.op
                        'replace with value.
                        'Add code to check for default variable object method calls.
                        'these would look like functions.
                        PerformFunction mCurrItem, tempeval
                        mCurrItem.ItemType = IT_VALUE
                        mCurrItem.op = ""
                        If IsObject(tempeval) Then
                            Set mCurrItem.Value = tempeval
                        Else
                            mCurrItem.Value = tempeval
                        End If
                    End If
                    
                    
                    
                    
                    
                
                End If
                Currpos = Currpos + Len(ParsedVal) + Len(argumentList) + 1
                NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
                'there.
            Else
                'it is a variable.
                'TODO:\\ add support to disable the current implementation of implicit variable definition.
                'if we get here and the variable doesn't exist (when Implicit is disabled) then we
                'need to raise a "Variable Not Defined" Error.
                
                mCurrItem.ItemType = IT_VARIABLE
                If IsObjAccess Then
                    'an object call, huh.
                    'well, specify it as such.
                    'there, and now Op will be the method,
                    'and Value will be the parameters.
                    '(IT_OBJACCESS is exclusive for @)
                    mCurrItem.ItemType = IT_ObjAccess
                    IsObjAccess = False
                End If
                mCurrItem.op = Trim$(ParsedVal)
                Currpos = Currpos + Len(RTrim$(ParsedVal)) - 1
                NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
            End If
        End If
        End If
        IsObjAccess = False
    Case Else
    'this was a tough decision that I might change.
    'couldn't think of the proper place to put the inspection calls to the Core plugin.
    'if I put them at the top, I'd allow "sub-classing" of any and all internally handled items, such as quotes and parentheses. However, each
    'plugin would then need to perform checks to ensure they weren't in quotes and such.
    'so I decided to put it here.
    
          
    End Select
    
    
    'TEST: check to see if we should add a IT_OPERATOR * between
    'this new token and the previous one.
    'NewItem will have been called, so mCurrItem points to the item that will be set on the next loop.
    'we want to compare mcurritem.prev.prev.itemtype and mcurritem.prev.itemtype.
    'CheckImplicit
    
    On Error Resume Next
    'If (CheckImplicit(mCurrItem.Prev, mPrevItem)) Then
    
    'End If
    Err.Clear
    
    
    
        Currpos = Currpos + 1
NEXTITERATION:
    Loop
    
    
    'check for errors.
    'CheckStack FirstItem
    InsertImplicit FirstItem
    Set RipFormula = FirstItem
'    CDebug.Post "Stack created:" & FirstItem.GetString(Me), Severity_Information
    CDebug.PopP "RipFormula"
Exit Function



'NewItem:
    'shut UP! I know these are slow.- GoSubs, I mean.
    'REALLY slow. But only if I compile to native code...
'    If Not mCurrItem Is Nothing Then
'        Set mPrevItem = mCurrItem
'    End If
'    If Currpos < Len(Expression) Then
'    Set mNextItem = New cFormitem
'    Set mCurrItem.Next_ = mNextItem
'    Set mCurrItem = mNextItem
'    End If
   ' NewItem mCurrItem, mPrevItem, mnextitem, Currpos, Expression
    
  '  Return

'RiPError:
'Stop
'Resume
'Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Friend Property Get ParentParser() As CParser
Set ParentParser = mParentParser
End Property


Friend Property Set EvalListener(vdata As IEvalEventSinks)
    Set mEvalListener = vdata
End Property

Friend Sub SetChild(ChildOf As CParser)
        Set mParentParser = ChildOf
End Sub

Friend Sub ParseInfix(ByVal StrParse As String, Optional ByVal NoCache As Boolean = False)
'IMPORTANT INFORMATION TO EVERYBODY WHO READS THIS:
'I encountered what some might call a bug, but is better called inconsistent.
'If the passed Expression is a constant value, calling the ParseInFix() Function will cause the
'Expression to be executed (when it stored it in the private variable)- However, if it isn't
'a constant expression, the Execution WILL NEVER BE PERFORMED (unless you call execute). This is important
'if you want your code to work the same regardless. However, if you call the Execute after every call to ParseInFix,
'you'll end up recieving two Execute events, and two Parse Events. Unless Optimizations are turned off via the registry (or the settings, I'll make that)

    
    
    
    CDebug.Post "ParseInfix parsing " & StrParse
    Dim grabitem As CFormItem
    Dim tempGrab As Variant, TestCast As CFormItem
    'first, make sure create was called. I COULD make it automatic, but
    'since some objects will then be automatically changed again...
    'set the current state for events and asynchronous operations.
    
    mState = ParseState_Parsing
'    If ParseMulti(StrParse, mFirstItem) Then
'        'good. we delegated to some child parsers- we simply sequentially call them.
'        mState = ParseState_Idle
'        Exit Sub
'    End If
    If mEvalListener Is Nothing Then
        'humm, it is nothing. Well, that means we haven't yet been initialized, so...
        ParserSettings.InitializeParser Me, "Default"
    
    End If
    
    'Secondly, do we even need to parse?
    'if we can cache and the parse stack is already created...
    If Not (NoCache Or (mFirstItem Is Nothing)) Then
        'we have the stack already created.
        
        'check, see if the expressions are equivalent.
        'mvarExpression will have been "made parsable" long ago.
        If StrComp(MakeParsable(StrParse), mvarExpression, vbTextCompare) = 0 Then
            'pfft. the two parsed expressions are the same, so, why are we doing in here, wasting
            'time.
            CDebug.Post "ParseInfix: Hmm, we already parsed the expression """ & StrParse & "."
            Exit Sub

        End If
'
    End If
    'if we don't have a local cache of the ParseStack, what about the Library-Wide optimized Stack? Let's see if the expression we seek is there.
    '(we won't find it if the expression contained fewer then 6 parse items)
 StrParse = MakeParsable(StrParse)
  On Error Resume Next
            If OptimizedStacks.Collection(StrParse) Is Nothing Then
                'StrParse = MakeParsable(StrParse)
                mvarExpression = StrParse
                CDebug.IncrementDebugVariable "Parse Cache Misses"
            Else
                'check the stack for a previous item.
                On Error Resume Next
                
                'ha HA! so, you thought you could escape.
                'grab this stack.
                If Err = 0 Then
                    'CDebug.Post "ParseInfix:- Hey, lookee here! some friendly fellow left me a parsestack for " & StrParse & ". How nice!"
                    'increment the cache hit's value.
                    CDebug.IncrementDebugVariable "Parse Cache Hits"
                    CDebug.IncrementDebugVariable ("Parse Cache Hits")
                    Set TestCast = OptimizedStacks.PopKey(StrParse)
                    'push it back on.
                    OptimizedStacks.Push TestCast, StrParse
                    Set mFirstItem = TestCast
                    mState = ParseState_Idle
                    Exit Sub
            
            'Else
            
            
            
            'register a cache miss. Sigh.
            
            
                End If
            
            
            
            
            End If
            
    On Error GoTo 0
    Set grabitem = RipFormula(StrParse)
    
    
    Set mFirstItem = New CFormItem
    
    Set mFirstItem = BuildStack_Infix(grabitem)
    
    
    'added Wednesday january 03 3:51 AM: optimization.
        'If InStr(Me.Expression, "{") <> 0 Then Stop
    If Me.IsConstant Then
        'if we are a constant expression, no use wasting time with a stack.
        'we'll retrieve our value, set the variables, and delete it to save memory.
        CDebug.Post "optimized parser call for " & Me.Expression
       'FIX: a STUPID bug where I was setting mIsConstant before the call to execute.
       'Execute quits if mIsconstant is true.
        Me.ExecuteByRef mvarConstantValue
         mIsConstant = True
        'die.
        Set mFirstItem = Nothing
    ElseIf GetFormitemCount(mFirstItem) > ParserSettings.CacheOptimizeBoundary Then
        'otherwise,-
        'if we have more then FORMITEM_OPTIMIZE_MIN values.
        
            'ok, since we had a heck of a time parsing this expression,
            'we'll put it in the OptimizedStacks Dictionary, then any other
            'Parser object in this address space that encounters the same expression
            'will be able to skip the parse step (oftentimes the most lengthy part of the process)
            CDebug.Post "ParseInfix: Whew, that formula '(" & StrParse & ")' was a toughie- I'll put it on the stack, so my friends can have an easier time."
            mFirstItem.Extra = StrParse
            On Error Resume Next
            OptimizedStacks.Push mFirstItem, StrParse
        
        
        
    End If
        mEvalListener.ParseComplete Me
        RaiseEvent ParseComplete
        mState = ParseState_Idle
End Sub

Friend Sub RaiseParseError(ParserError As CParserError, RecoveryConst As ParserErrorRecoveryConstants)
'default to ignore it, and return the error info as the return value.
    RecoveryConst = PERR_RETURN
    RaiseEvent Error(ParserError, RecoveryConst)
    If Not mParentParser Is Nothing Then
        'HA! so our "error" event was PROBABLY ignored.
        'well, force the error to be recognized, erm,- somewhere.
        With ParserError
            Err.Raise .Code, .Source, .Description
        End With
    
    End If
End Sub
Public Function GetVersionInformation() As ParserVersionInformation
    Static Ret As ParserVersionInformation
    If Ret.Major = 0 Then
        Ret.Major = App.Major
        Ret.Minor = App.Minor
        Ret.Build = App.Revision
    
    End If
    GetVersionInformation = Ret

End Function
Public Sub AddTypelibConstants(Flib As Variant)
Attribute AddTypelibConstants.VB_Description = "Adds The Constants and Enumeration s From the given typelib to the current Variable namespace."
    'adds as variable of constant type.
    'Input can be:
    'TypelibInfo object
    'Object (will get typelib from that)
    'CLSID (gross)
    'ProgID (will use to create object and load typelib)
    'Filename.
    Dim ConstFilter As tli.CustomFilter
    Dim fromlib As TypeLibInfo
    Dim results As SearchResults, Currloop As Object
    Dim currTI As Long, loopItem As SearchItem
    Dim CurrTypeInfo As TypeInfo, EnumConstant As MemberInfo
    Dim EnumType As TypeInfo, newvariable As CVariable
    Static TLIApp As tli.TLIApplication
    If TLIApp Is Nothing Then Set TLIApp = New tli.TLIApplication
    Set fromlib = Flib

        
        'retrieve all constants.
        Set results = fromlib.GetTypes(, 4)
        'now, iterate through each enum...
        
        
        
    
        Dim valNames() As String, Values() As Variant, newEnumVar As CObject, mvar As CVariable
        Dim valcount As Long
        For Each loopItem In results
            'use the typeinfonumber to get the proper TypeInfo object for this Enumeration...
            Set EnumType = fromlib.GetTypeInfo(loopItem.TypeInfoNumber)
            ReDim valNames(1 To EnumType.Members.Count)
            ReDim Values(1 To EnumType.Members.Count)
            valcount = 0
            Set newEnumVar = New CObject

            For Each EnumConstant In EnumType.Members
               
                    valcount = valcount + 1
                    'valNames(valcount) = EnumConstant.Name
                    'Values(valcount) = EnumConstant.Value
                    newEnumVar.Variables.Add EnumConstant.Name, EnumConstant.Value
                    
                    
            Next EnumConstant
            
            
            
            
            
            
            'now, finally, add a variable with the name of the enumeration....
            
            Set mvar = mVariables.Add(loopItem.Name, newEnumVar)
            Set mvar.Value = newEnumVar
            mvar.Constant = True
            
            
            
        Next loopItem
        
        
        
    
    
  


End Sub
Friend Function IsDecendentOf(ParserObj As CParser)
Dim CurrParser As CParser
Set CurrParser = Me
Do
    If CurrParser.ParentParser Is ParserObj Then
        IsDecendentOf = True
        Exit Do
    ElseIf CurrParser.ParentParser Is Nothing Then
        IsDecendentOf = False
        Exit Do
    End If
    Set CurrParser = CurrParser.ParentParser
Loop





End Function
