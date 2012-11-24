VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BPCoreOpFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : BPCoreOpFunc
' DateTime  : 1/4/2007 03:37
' Author    : Michael Burgwin
' Purpose   :contains the logic that implements BASeParser's Core operators and functions.
'              the list is quite large, and, in most cases, should make the creation
'           of new plugins unnecessary.
'---------------------------------------------------------------------------------------
'NOTICE:

'If I don't release the source code for the library, I will still release this class. Then you
'can use it as an example of how to make a plugin for the parser- this "plugin" contains
'the core implementation of operators and functions.- as well as Object-Like access to numbers and
'arrays.

'of course, it is somewhat difficult to create another plugin based in this one
'without the parser code, since it uses methods and properties that are only
'available to those privileged classes inside the library.




Option Explicit
Implements IEvalEvents
Implements IPluginUIData
Implements IOpFuncInfo
Implements ICorePlugin
Const MODULE_NAME = "BPCoreopFunc"
Const MODULE_FILE_NAME = "BPCoreopFunc.Cls"
Private Const Flowchange_private = -256


'Operators:
'I summarize the operators here. Followed by the actual constant that
'defines them, and is returned by the interface method back to
'the parser engine.
'As far as operators go, I try to implement those from popular languages.
'note, however, the absence of =. I want to at some point make it an assignment operator.
'********
'*NOTE*- Yep- that comment above is old. Anyway, I made ":=" the assignment operator. "=" is
'the comparision operator. Wait- no it isn't...
'********
' +     Addition
' -     Subtraction, and Unary negation.
' *     Multiplication
' \     Division
' / and DIV    Integer Division
'^ and **   Exponentiation
'Comparision operators:
'==     logical equality
'<>,!=     logical inequality.
'<=,=<  less then or equal to.
'>= =>  Greater then or equal to.
' <     less than
' >     greater than
'<=>    "spaceship" sort comparator. Taken from Perl. I feel no guilt.

'stuff I want but are having difficulty with Via VB:
'Bit shifts.
'<<<,>>>
'<<
'>>



'Logical operators:

'AND
'OR
'XOR
'EQV
'IMP

'Unary operators:
'prefix:
'NOT, - (negation), SIZEOF (len function, basically)
'Suffix:
'!

'OBJECT ACCESS: (OK, this is actually handled within the
'Parser itself, but-
'the @ sign is used between an object expression and a property/Method call.



'All handled operators, in one clump.
'this makes testing if we can handle a operator as simple as a single boolean conversion of
'the intrinsic Instr() Function.
'update- flipped precedence of + and - to make it consistent with BASIC. Didn't flip their meanings though. :)
'added a bunch of operators- Repeat, ternary helper operators ? and :, which
'can be used separately- REPEAT, which repeats a string,$$, which performs RegEx pattern
'matching and returns the VBScript RegExp object,etc.
'Modulo arithmetic uses MOD and "//" (from rexx). I want to use %, but I think I'll use that
'as a item delimiter to indicate environment variables.
'(this is implemented within the Core plugin, CPlugEnvString
Private Const HandledOpConst = " << >> UPTO + - = == =~ MOD // * / \ DIV ^ ** != <> <=> <= =< >= >= < > AND && OR ||  XOR EQV IMP REPEAT LIKE IS IN NCR NPR PICK SORT $$ ? : "



'PICKMENU: not implemented yet.
'the idea will be to pop-up a context menu with the specified array elements-
'array items within that array will become sub-menus.
'very important- functions should be listed from longest to shortest, or, at the very least,
'functions who are the same as another function but with something on the end should come after
'the other- for example, InstrRev should come BEFORE Instr, because the parser will look in our GetHandledFunctionString
'and see that the Instr part of InstrRev matches our "InStr" function, then it will find a variable
'named rev.


'Kludge: needed to add the intrinsic array methods here so that
'the parser will accept them as functions, rather then, say, finding the Or in Sort and treating it as an operator
'between two variables, s and t
'of course, now WE need to implement it for non-objaccess calls, but I should probably do that anyway)
Const AllVarsIntercepted As String = "TRUE FALSE BPEMPTY VBEMPTY BPNULLCHAR VBNULLCHAR BPMONDAY VBMONDAY BPTUESDAY " & "VBTUESDAY BPWEDNESDAY VBWEDNESDAY BPTHURSDAY VBTHURSDAY BPFRIDAY VBFRIDAY " & "BPSATURDAY VBSATURDAY BPSUNDAY VBSUNDAY CURDIR CURDIR$ PARSER DIR DIR$ LISTVARS BPTEXTCOMPARE BPBINARYCOMPARE BPUNICODE I "

                
  '"TRUE","FALSE","BPEMPTY", "VBEMPTY","BPNULLCHAR", "VBNULLCHAR",
      '"BPMONDAY", "VBMONDAY","BPTUESDAY", "VBTUESDAY"
      '"BPWEDNESDAY", "VBWEDNESDAY"
       ' "BPTHURSDAY", "VBTHURSDAY"
         '"BPFRIDAY", "VBFRIDAY"
       '"BPSATURDAY", "VBSATURDAY"
       '"BPSUNDAY", "VBSUNDAY"
       '"CURDIR", "CURDIR$"
     '"PARSER"
      '"DIR", "DIR$"
         '"LISTVARS", "ALLVARS"
'ArayOperations: holds the operators which
'EXPECT to have an array operand.

Private FuncInfo As MSXML2.DOMDocument30       'Function information XML file.

Const ArrayOperations = " IN PICK ? : SORT "
Private Const UnaryOpConst = " << -- ++ - NOT SIZEOF SORT "
Private Const UnaryOpSuffix = " ! -- ++ "
Private mvarDisabledFunctions As String
Private mvarDisabledOperators As String 'disabled functions, and Operators.
Private Const NonConstantOp = ""


Private Const IT_ASSIGNMENT = -23485324
Private Const IT_ABSPIPE = -23485323
'If these appear in an expression that expression is not constant- that is,
'the parser is not allowed to optimize them.
'Private Const NonConstantOp = " := += -= *= /= \= ^= "
'BASeParser Core Operators and Functions.
'This was originally inside the actual interface, but I decided to move it here.
Public Function CompleteParsedFunctions() As String
Static Ret As String, Flinit As Boolean
If Not Flinit Then
    Ret = " DUMPEXPRESSION INSTALLPLUGIN LOADPLUGIN LISTPLUGINS FOREACH LISTVARIABLES IF " & _
    "FRAC FLATTEN MATRIX COMPLEX EVALARR COUNTIF PICKMENU RANDOMARRAY RANDOM STORE WHILE SOLVE ABS " & _
    "ARRAY ASC ORD ATN CHOOSE CHR$ CHR CREATEOBJECT CVERR DATEADD DATEDIFF DATEPART DATESERIAL "
    Ret = Ret & "DATEVALUE YEAR MONTH DAY MINUTE SECOND DDB DIR$ DIR ENVIRON ERROR EXP FILEDATETIME EXEC " & _
    "FILELEN FILTER FORMAT$ FORMATCURRENCY FORMATDATETIME FORMATNUMBER FORMATPERCENT FV FORMAT LISTMEMBERS "
    Ret = Ret & "GETATTR GETOBJECT HEX$ HEX IIF INPUTBOX$ INPUTBOX INSTRREV INSTR INT FIX IPMT IRR ISARRAY " & "ISDATE ISEMPTY ISERROR ISNULL ISNUMERIC ISOBJECT JOIN LBOUND LCASE$ LCASE LEFT$ LEFT LEN LOG " & _
    "LTRIM$ LTRIM RTRIM$ RTRIM TRIM$ TRIM MID$ MID MIRR MONTHNAME MSGBOX NPER NPV OCT$ OCT PARTITION " & _
    "PMT PPMT PV POS SUBSTR QBCOLOR RATE REPLACE$ REPLACE RGB RIGHT$ RIGHT ROUND " & _
    "SGN SHELL SLN SPACE$ SPACE SPLIT SQRT SQR STR$ STR STRCOMP STRCONV STRING$ STRING STRREVERSE SWITCH " & "SYD TIMESERIAL TIMEVALUE TYPENAME UBOUND" & " UCASE UCASE$ VAL VARTYPE"
    Ret = Ret & " WEEKDAYNAME WEEKDAY MINUS MEAN AVERAGE SORT MIN MAX " & "SEC SIN COS TAN ASEC ASIN ACOS ATN HASEC HSIN HCOS HTAN HASIN HACOS HATAN SEQEX " & "SEQ CEIL FLOOR SELECT COBJECT CINT CLNG CSNG CDBL CCUR CBYTE CSTR VERSION WITH "
    
Flinit = True
End If
CompleteParsedFunctions = Ret
End Function



Private Sub Class_Initialize()
    On Error Resume Next
    Dim fileXML As String
    CDebug.Post "BPCoreOpFunc Class Initialized"
    'load the XML data for Function information.
    'NOTE: change this code to load the file from a resource file instead.
    
    fileXML = App.Path & "\infodata\bpcorehelp.xml"
    '!!This needs to be changed to a resource file once the XML file is done.
    
    
    
        Set FuncInfo = New DOMDocument
        
        If Not FuncInfo.Load(fileXML) Then
            Debug.Print "failed to Load XML data:" & vbCrLf & vbCrLf & FuncInfo.ParseError.reason & " on line """ & FuncInfo.ParseError.Line & """" & vbCrLf & _
                "position """ & FuncInfo.ParseError.linepos & """."
            'Debug.Assert False
        End If
    
    Err.Clear
End Sub

Private Sub Class_Terminate()
On Error Resume Next
CDebug.Post "BPCoreOpFunc Class Terminated"

End Sub

Private Function ICorePlugin_CanHandleItem(withparser As CParser, Itemcheck As CFormItem) As Boolean
'
If Itemcheck.ItemType = IT_ASSIGNMENT Or Itemcheck.ItemType = IT_ABSPIPE Then
    ICorePlugin_CanHandleItem = True
End If
End Function

Private Function ICorePlugin_GetPluginUIData() As IPluginUIData
Set ICorePlugin_GetPluginUIData = IEvalEvents_GetPluginUIData
End Function

Private Sub ICorePlugin_HandleItem(withparser As CParser, ItemHandle As CFormItem, retval As Variant)
'TODO: add more robustness for the assignment operators to work with arrays.
Dim byrefret As Variant, newvalue As Variant
Debug.Print "handling op:" & ItemHandle.op
Dim NewVar As CVariable
If ItemHandle.ItemType = IT_ABSPIPE Then
     ItemHandle.Value.ExecuteByRef retval
     Exit Sub
End If

If ItemHandle.ItemType = IT_ASSIGNMENT Then
If withparser.Variables.Exists(ItemHandle.Value(1)) Then
    Set NewVar = withparser.Variables.Item(ItemHandle.Value(1))
Else

    Set NewVar = withparser.Variables.Add(ItemHandle.Value(1), 0)
End If


    If ItemHandle.op = "+=" Then
        ItemHandle.Value(0).ExecuteByRef byrefret
    
        Call withparser.PerformOperation("+", NewVar.Value, byrefret, retval)
        NewVar.Value = retval
    ElseIf ItemHandle.op = ":=" Then


        ItemHandle.Value(0).ExecuteByRef byrefret
        If Not IsObject(byrefret) Then
            'retval = withparser.Variables.Add(ItemHandle.Value(1), byrefret)
            Set retval = withparser.Variables.Add(ItemHandle.Value(1), byrefret)
        Else
             Set retval = withparser.Variables.Add(ItemHandle.Value(1), byrefret)
        End If
    End If
End If
End Sub

Private Function ICorePlugin_Init(withparser As CParser) As IPluginInitCode
    CDebug.Post "BPCoreOpFunc Initialized as CorePlugin"
End Function

Private Function ICorePlugin_ParseLocation(withparser As CParser, ByVal StrParse As String, Position As Long, currItem As CFormItem) As ParseLocationReturnConstants
'
'Specialty Code for somewhat kludged handling of &H and &O and so forth

'what happens here, you ask? well, the core Parser shouldn't have to deal with such arcane things as hexadecimal
'values and so forth, so rather, I made the Core Operators/Functions evaluator plugin also implement the ICorePlugin interface.
'then, when it finds &H or another set base identifier, it will create a new IT_VALUE type on the passed item, which the core parser WILL understand.
'this way, the actual core string manip routines can be optimized to NOT deal with ANY separate number system.
Const HexChars As String = "1234567890ABCDEF."
Const OctalChars As String = "12345670."
Const BinChars As String = "10"
Dim CharView As String
Dim Charsets() As String
Dim Prefixes() As Variant
ReDim Charsets(1 To 3)
ReDim Prefixes(1 To 3)
Charsets(1) = HexChars
Charsets(2) = OctalChars
Charsets(3) = BinChars
Prefixes(1) = Array("&H", "0x")
Prefixes(2) = "&O"
Prefixes(3) = "&B"
Const IDX_HEX = 1
Const IDX_OCT = 2
Const IDX_BIN = 3
Dim ValidChars As String
Dim LoopPos As Long
Dim Loopprefix As Long, InnerPref As Long
Dim FoundFix As Long
Dim PrefixFound As Long, PrefixFoundStr As String
'yes, a period is a valid token in a hex literal.
'never mind the connotations- I think for now I will ignore it.

'notes on parsing:
'For hex, we will find either a 0x or an &H at the specified location. to parse, we simply
'find the next space,open bracket, or Non hex character.
'Hex= &H or 0x
'Octal &O
'Binary &B
Dim inext As Long, lblname As String
'Step one: detect DEFLABEL lines.
'If Mid$(StrParse, Position, Len("DEFLABEL")) = "DEFLABEL" Then
'    'find the next close bracket. the contents are the label name.
'    inext = InStr(Position + Len("DEFLABEL") + 1, StrParse, ";", vbTextCompare)
'    lblname = Mid$(StrParse, Position + Len("DEFLABEL"), inext - Position - Len("DEFLABEL"))
'
'    'create the formula item.
'    currItem.ItemType = IT_NULL
'    currItem.Value = lblname
'
'    On Error Resume Next
'
'    Labelled.Add currItem, lblname
'
'    Position = inext + 1
'End If
Dim makeparsers() As Variant


Dim foundassignment As Boolean
Dim bracketcount As Long, intick As Boolean, inQuote As Boolean
Dim Currpos As Long
Dim assignmentOps() As String, absstring As String
ReDim assignmentOps(1 To 5)
assignmentOps(1) = ":="
assignmentOps(2) = "+="
assignmentOps(3) = "-="
assignmentOps(4) = "*="
assignmentOps(5) = "/="
Const BracketStarts = "[{(<"
Const BracketEnds = ">)}]"
Dim currassign As Long
'If Mid$(StrParse, Position, 2) = ":=" Then

  If Mid$(StrParse, Position, 1) = "|" Then
        Currpos = Position + 1
        absstring = ""
        Do
            If Mid$(StrParse, Currpos, 1) = "|" Then
               Exit Do
            Else
               absstring = absstring & Mid$(StrParse, Currpos, 1)
            End If
            Currpos = Currpos + 1
        Loop
    
        currItem.ItemType = IT_ABSPIPE
        currItem.Extra = absstring
        Set currItem.CoreOwner = Me
        currItem.op = "|"
        Set currItem.Value = withparser.Clone
        currItem.Value.Expression = "ABS(" & absstring & ")"
        Position = Position + Len(absstring) + 2
        ICorePlugin_ParseLocation = ParseLocation_recognized
        Exit Function
    
    End If


If Position = 1 Then
'only do this ONCE for an expression...
'what are we doing? iterating through the string. if we encounter a := that is NOT within any brackets or quotes, we split the string in half right there.

    For Currpos = 1 To Len(StrParse)
        If InStr(1, BracketStarts, Mid$(StrParse, Currpos, 1), vbBinaryCompare) > 0 Then
            bracketcount = bracketcount + 1
        ElseIf InStr(1, BracketEnds, Mid$(StrParse, Currpos, 1), vbBinaryCompare) > 0 Then
            bracketcount = bracketcount - 1
        ElseIf Mid$(StrParse, Currpos, 1) = """" Then
            inQuote = Not inQuote
        ElseIf Mid$(StrParse, Currpos, 1) = "`" Then
            intick = Not intick
        ElseIf bracketcount = 0 And Not inQuote And Not intick Then
            'well- we aren't in anything...
            'is it a assignment op, though?
            Dim currop As String
            For currassign = 1 To UBound(assignmentOps)
            currop = assignmentOps(currassign)
            
              If Mid$(StrParse, Currpos, Len(currop)) = currop Then
                  'ha HA!
                  'split in half...
                  Dim VarName As String
                  Dim Expression As String
                  VarName = Mid$(StrParse, 1, Currpos - 1)
                  Expression = Mid$(StrParse, Currpos + 2)
                      currItem.ItemType = IT_ASSIGNMENT
                      currItem.op = currop
                  
                  currItem.Tag = VarName
                  ReDim makeparsers(1)
                  Set makeparsers(0) = New CParser
                  Set makeparsers(0) = withparser.Clone
                  makeparsers(0).Expression = Expression
                  makeparsers(1) = VarName
                  currItem.Value = makeparsers
                  
                  ICorePlugin_ParseLocation = ParseLocation_recognized
                  
                      
    
                  Position = Len(StrParse) * 2 'end of string...
                  Exit Function
              End If
            Next currassign
        End If
    
    Next
    

    'get var name. assume everything UP to the :=.
    'change this- perhaps to find whitespace.
    currItem.Tag = Mid$(StrParse, 1, Position - 1)

    
    
    Exit Function
End If

For Loopprefix = 1 To UBound(Prefixes)
    If IsArray(Prefixes(Loopprefix)) Then
        For InnerPref = 0 To UBound(Prefixes(Loopprefix))
            If Mid$(StrParse, Position, Len(Prefixes(Loopprefix)(InnerPref))) = Prefixes(Loopprefix)(InnerPref) Then
                FoundFix = True
                PrefixFound = Loopprefix
                PrefixFoundStr = Prefixes(Loopprefix)(InnerPref)
                Exit For
            End If
            
        
        Next InnerPref
    Else
        If Mid$(StrParse, Position, Len(Prefixes(Loopprefix))) = Prefixes(Loopprefix) Then
                FoundFix = True
                PrefixFound = Loopprefix
                PrefixFoundStr = Prefixes(Loopprefix)
                Exit For
        End If
    End If

Next
If PrefixFound Then
    ValidChars = Charsets(PrefixFound)
    Dim I As Long
    I = Position + Len(PrefixFoundStr)
    Do Until InStr(1, ValidChars, Mid$(StrParse, I, 1), vbBinaryCompare) = 0 Or Mid$(StrParse, I, 1) = ""
        I = I + 1
    Loop
    'so, we go from Position to I
    Dim StrExtract As String
    StrExtract = Mid$(StrParse, Position, I - Position)
    currItem.ItemType = IT_VALUE
    currItem.Value = ValEx(StrExtract)
    Position = I
    ICorePlugin_ParseLocation = ParseLocation_recognized
    Exit Function
Else
     
End If

'last, examine for special contructs, such as IF and so forth.

'the formula item will be set to flowchange_private , OP will be the type of flowchange construct (IF, FOR, etc.

'then we parse out the contents of said flowchange. For example- an IF will be like so:
'IF(expression);
'<Statements separated by ;>
'ELSE
'<more statements- separated by ;>
'END IF





End Function

Private Property Let ICorePlugin_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get ICorePlugin_Settings(withparser As CParser, ByVal PropName As String) As Variant
''
End Property

Private Function ICorePlugin_Term(FromParser As CParser) As IPluginTermCode
'
End Function

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As CParser, ByVal VarName As String, VarValue As Variant)
'
'change the value as appropriate.

End Sub



Private Sub AddOurVars(Toparser As CParser)
Dim v As String, I As Long
Dim AllVarNames() As String
v = "<Variable Placeholder>"
    AllVarNames = Split(AllVarsIntercepted, " ")
    For I = 0 To UBound(AllVarNames)
        AllVarNames(I) = Trim$(AllVarNames(I))
        If AllVarNames(I) <> "" Then
            Toparser.Variables.Add AllVarNames(I), v
        End If
    Next I
    Toparser.Variables.Add "i", Complex(0, 1)
                    
      '"TRUE","FALSE","BPEMPTY", "VBEMPTY","BPNULLCHAR", "VBNULLCHAR",
      '"BPMONDAY", "VBMONDAY","BPTUESDAY", "VBTUESDAY"
      '"BPWEDNESDAY", "VBWEDNESDAY"
       ' "BPTHURSDAY", "VBTHURSDAY"
         '"BPFRIDAY", "VBFRIDAY"
       '"BPSATURDAY", "VBSATURDAY"
       '"BPSUNDAY", "VBSUNDAY"
       '"CURDIR", "CURDIR$"
     '"PARSER"
      '"DIR", "DIR$"
         '"LISTVARS", "ALLVARS"
End Sub

Private Function IEvalEvents_AssignSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, AssignThis As Variant) As Boolean
'
If IsArray(OnValue) Then
    Call Assign(OnValue(subscripts(0).Execute), AssignThis)
    IEvalEvents_AssignSubScript = True
End If


End Function

Private Function IEvalEvents_BeforeFunctionCall(withparser As CParser, FuncCalling As String, WithHandler As IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_GetOperatorFlags(withparser As CParser, ByVal OpCheck As String, Unary As IEvalUnaryOpFlags) As OperatorFlags
    Dim retThis As OperatorFlags
    If InStr(NonConstantOp, " " & OpCheck & " ") <> 0 Then
        retThis = Operator_Implemented
    Else
        retThis = Operator_Implemented + Operator_Optimizable
    End If
    IEvalEvents_GetOperatorFlags = retThis
End Function

Private Function IEvalEvents_GetPluginUIData() As IPluginUIData

    Set IEvalEvents_GetPluginUIData = Me
End Function

Private Function IEvalEvents_HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'handle strings.
'Subscripts is a Array of CParser objects.
Dim strcast As String, I As Long
Dim words As Variant
Dim strsplit() As String
Select Case VarType(OnValue)
    Case vbString
                    'UPDATE: allow this to support multiple arguments (with results similiar to mid....)
                    'for example:
                    '"HELLO"[2,2] would be "EL"
                    
                    strcast = OnValue
                    'ok, grab that digit.
                    On Error Resume Next
                    If UBound(subscripts) = 0 Then
                        If Err <> 0 Then
                            'If an error occured, we have no subscripts at all.
                            'return the string array as an array of characters.
                            strsplit = Split(strcast, " ")
                            ReDim words(1 To UBound(strsplit) + 1)
                            For I = 0 To UBound(strsplit)
                                words(I + 1) = strsplit(I)
                            Next I
                            retval = words
                            IEvalEvents_HandleSubScript = True
                            Exit Function
                            
                            
                            
                            
                        Else
                            On Error GoTo 0
                            Dim returnValue As Variant, strcoerced As String
                            subscripts(0).ExecuteByRef returnValue
                            If IsArray(returnValue) Then
                                If IsArray(returnValue(0)) Then returnValue(0) = Flatten(returnValue(0))(0)
                                If IsArray(returnValue(1)) Then returnValue(1) = Flatten(returnValue(1))(1)
                                strcast = Mid$(strcast, returnValue(0), returnValue(1))
                            Else
                                strcoerced = withparser.ResultToString(returnValue)
                                strcast = Mid$(strcast, strcoerced, 1)
                            End If
                        End If
                    ElseIf UBound(subscripts) > 0 Then
                        strcast = Mid$(strcast, subscripts(0).Execute, subscripts(1).Execute)
                    
                    End If
                    'Strings only support a single "Subscript"
                    retval = strcast
                    IEvalEvents_HandleSubScript = True
    Case vbInteger  '4[2]=0
        'copies either byte One or two from the integer, and returns it.
        'remember, this is in big endian.
        Select Case subscripts(0).Execute
            Case 0
                'return an error describing that 1 should be used.
                'ParserError(WithParser, 9, "Subscripts used with Integer value are base 1- 1 or 2 to return the corresponding byte.", "BPCoreOpFunc::HandleSubScript").Throw
                Err.Raise 9, "BPCoreOpFunc.HandleSubscript", "subscripts used with Integer Value are base 1. 1 or 2 to return corresponding byte."
            Case 1
                'copy first byte.
                retval = LoByte(CInt(OnValue))
            Case 2
            
                retval = HiByte(CInt(OnValue))
            Case Else
                Err.Raise 9, "BPCoreOpFunc.HandleSubScript", "Illegal Subscript"
                'ParserError(WithParser, 9, "Illegal Subscript", "BPCoreOpFunc.HandleSubScript").Throw
        End Select
  '  Case vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal
  
        IEvalEvents_HandleSubScript = True
    
    
    
End Select
End Function

Private Function IEvalEvents_OnInvokeObject(withparser As CParser, ObjInvoke As Variant, currItem As CFormItem, Result As Variant) As Boolean
'Since the parser, intrinsically, only supports the @  object access with, umm- Objects, BPCoreOpFunc
'adds Array "methods", such as MEAN().
Dim rCount As Long, Evalled() As Variant
Dim looper As Long

    

If IsArray(ObjInvoke) Then
                    'array calls.
                    
                   Call performArrayOp(withparser, ObjInvoke, currItem, Result)
                   IEvalEvents_OnInvokeObject = True
ElseIf VarType(ObjInvoke) = vbString Then
    'strings have methods too. They get translated directly into calls to
    'the PARSER's IEvalEvents handler. This way, if another handler adds a string manipulation routine, we can use it
    'as a method of the string object- the only requirement is that the first argument be a string.
    
    Evalled = ParserArrToResultArr(currItem.Value, rCount)
    'OK- remember- the first argument will be the string.
    'Of course, we might get a type mismatch.....
    On Error Resume Next
    ReDim Preserve Evalled(UBound(Evalled) + 1)
    If Err = 0 Then
        'move all elements up one.
        For looper = 0 To UBound(Evalled) - 1
            Call Assign(Evalled(looper + 1), Evalled(looper))
        Next
    Else
        ReDim Evalled(0)
    End If
    Evalled(0) = ObjInvoke
    Err.Clear
    
    'let errors propagate.
    On Error GoTo 0
    Call withparser.EvalListener.Self.HandleFunction(withparser, currItem.op, Evalled(), Result)
    IEvalEvents_OnInvokeObject = True
ElseIf VarType(ObjInvoke) = vbDate Then
    Evalled = ParserArrToResultArr(currItem.Value, rCount)
    Select Case LCase$(currItem.op)
        Case "year"
        Result = Year(ObjInvoke)
        Case "month"
        Result = Month(ObjInvoke)
        Case "day"
        Result = Day(ObjInvoke)
        Case "hour"
        Result = Hour(ObjInvoke)
        Case "minute"
        Result = Minute(ObjInvoke)
        Case "second"
        Result = Second(ObjInvoke)
        Case "add"
        If rCount = 1 Then
            Result = DateAdd("d", Evalled(LBound(Evalled)), ObjInvoke)
            
        End If
            
        
    
    End Select
    IEvalEvents_OnInvokeObject = Not IsEmpty(Result)

End If
    'IEvalEvents_OnInvokeObject = True

End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As CParser, ByVal VarName As String, ByRef cancel As Boolean) As Variant
'[Constant replacement code]

Dim Ret As Variant
Set Ret = Me  'ha! now THAT is a flag variable,
'there is NO WAY that the constant/variable can return THAT :)
'this is where functions that can be called without brackets would be implemented.

'--5/26/2007 @ 05:25--
'
'commented out environment variable handling to test the 30-minute old ICorePlugin interface
'implementation of such. Works! Great :)
'If left$(VarName, 1) = "%" And right$(VarName, 1) = "%" Then
'environment variable.
'    Ret = Environ$(Mid$(VarName, 2, Len(VarName) - 2))

'Else
    Select Case UCase$(VarName)
        Case "I"
        'create a new variable I that is a complex
        'number with realpart 0 and imaginary part 1.
        Set Ret = Complex(0, 1)
        Case "TIMER"
        
        
            Ret = Timer
        Case "RANDOMIZE"
            Randomize
        Case "NOW"
            Ret = Now
          Case "PI"
            Ret = Atn(1) * 4
        Case "E"
            Ret = Exp(1)
        Case "NULL"
            Ret = Null
        Case "TRUE"
            Ret = True
        Case "FALSE"
            Ret = False
        Case "BPEMPTY", "VBEMPTY"
            Ret = vbEmpty
        Case "BPNULLCHAR", "VBNULLCHAR"
            Ret = vbNullChar
        Case "BPMONDAY", "VBMONDAY"
            Ret = vbMonday
        Case "BPTUESDAY", "VBTUESDAY"
            Ret = vbTuesday
        Case "BPWEDNESDAY", "VBWEDNESDAY"
            Ret = vbWednesday
        Case "BPTHURSDAY", "VBTHURSDAY"
            Ret = vbThursday
        Case "BPFRIDAY", "VBFRIDAY"
            Ret = vbFriday
        Case "BPSATURDAY", "VBSATURDAY"
            Ret = vbSaturday
        Case "BPSUNDAY", "VBSUNDAY"
          
            Ret = vbSunday
        Case "BPTEXTCOMPARE"
            Ret = vbTextCompare
        Case "BPBINARYCOMPARE"
            Ret = vbBinaryCompare
        Case "BPUNICODE"
            Ret = vbUnicode
        Case "CURDIR", "CURDIR$"
            Ret = CurDir$
        Case "PARSER"
            'returns object. Parser.
            Set Ret = ParserUse
        Case "DIR", "DIR$"
           ' On Error Resume Next    'hey, it isn't OUR fault!
            Ret = Dir$
            If Err <> 0 Then
                Ret = ""
   
            End If
        Case "LISTVARS"
            Dim LoopVar As CVariable, I As Long
            Dim Subparser As CParser
            'return a listing of all variables.
            'of course, it only includes those
            'variables that have been added either
            'implicitly or explicitly through STORE()
            With ParserUse.Variables
                If .Count = 0 Then
                    'no variables! return that message.
                    Ret = "No Variables Defined."
                Else
                    Ret = ""
                    Ret = "--Variables Listing--" & vbCrLf & _
                          "NAME" & vbTab & "VALUE" & vbCrLf
                          Set Subparser = ParserUse.Clone
                          Subparser.Create ParserUse.configset
                    'For Each LoopVar In ParserUse.Variables
                    For I = ParserUse.Variables.Count To 1 Step -1
                    Set LoopVar = ParserUse.Variables.Item(I)
                    'DON'T retrieve the VALUE of the variable-
                    'instead, (given the name isn't LISTVARS ( :))
                    'create a subParser of the CParser given, parse the variable name,
                    'and return THAT.
                    'this way, all the other pseudo constants will be listed.
                        If StrComp(UCase$(LoopVar.Name), "LISTVARS", vbBinaryCompare) <> 0 And LoopVar.Name <> "" Then
                            
                            Subparser.Expression = LoopVar.Name
                            
                            Ret = Ret & LoopVar.Name & "=" & Subparser.ResultAsString & vbCrLf
                        
                        
                        End If
                    Next
                    Ret = Ret & "------------------------" & vbCrLf
                    
                    Ret = Ret & "Total Variables Listed:" & .Count
                    
                End If
                
                    
                    
            End With
                
                
            
            
          
    End Select
'End If
If Not IsObject(Ret) Then


        IEvalEvents_BeforeResolveVariable = Ret
        cancel = True
    Else
        If Not Ret Is Me Then
            cancel = True
       
        Set IEvalEvents_BeforeResolveVariable = Ret
         End If
End If
       


End Function


Private Function IEvalEvents_CanHandleFunction(withparser As CParser, ByVal funcName As String, NoParsedArgs() As Boolean) As Boolean
'CompleteParsedFunctions:
'this constant contains all the functions that accept completely parsed expressions.
'This takes care of pretty much all the built-in functions.


'-whew. that took a while.





'redimension the NoParsedArgs Array to prevent an error.
'256 is a good limit. but only if it is not higher then that already.
On Error Resume Next
'If UBound(NoParsedArgs) < 255 Then
'    ReDim NoParsedArgs(LBound(NoParsedArgs), 256)
'End If
'On Error GoTo 0
If InStr(1, mvarDisabledFunctions, UCase$(funcName), vbBinaryCompare) <> 0 Then
    'disabled.
    IEvalEvents_CanHandleFunction = False
     CDebug.Post "ignored query for handling of function, " & """" & funcName & """."
     Exit Function
End If
Select Case True
       
'        Case left$(FuncName, 13) = "BPCOREOPFUNC."
'            'force recognition. might not work though.
'
'            IEvalEvents_CanHandleFunction = True
'
'        Case UCase$(funcName) = "GOTO"
        Case UCase$(funcName) = "WHILE"
'        "WHILE(Expression,<statements>)"
        NoParsedArgs(0) = True
        NoParsedArgs(1) = True
        Case UCase$(funcName) = "FILTER"
            'Filter(Array,Variable,Expression)
            NoParsedArgs(1) = True
            NoParsedArgs(2) = True
        
        
        Case UCase$(funcName) = "WITH"
        NoParsedArgs(0) = False
        NoParsedArgs(1) = True
        
        Case UCase$(funcName) = "STORE"           'STORE(VarName,Value)
            'stores a variable value.
            'the first argument is the name.
            IEvalEvents_CanHandleFunction = True
            NoParsedArgs(0) = True
            'NoParsedArgs(2) = True
        'String manip functions.
        'for most of these, all the values are parsed.
        Case UCase$(funcName) = "FOREACH"
        'FOREACH({1,2,3},varname,Expression)
        NoParsedArgs(1) = True
        NoParsedArgs(2) = True
        Case UCase$(funcName) = "IIF" Or UCase$(funcName) = "IF"
            'IIF: second and
            NoParsedArgs(1) = True
            NoParsedArgs(2) = True
        
        
        Case UCase$(funcName) = "SEQ"
            'first one (expression) is not evaluated:
            NoParsedArgs(0) = True
            NoParsedArgs(1) = True
            IEvalEvents_CanHandleFunction = True
        Exit Function
        Case UCase$(funcName) = "EVALARR"
        'first two NOT parsed (Expression,VarName)
            NoParsedArgs(0) = True
            NoParsedArgs(1) = True
            
        Case UCase$(funcName) = "SEQEX"
        'Expression , Varname, InitialValue, IncrementExpression, TerminateExpression
        'expression,incrementexpression, and terminateexpression are not parsed.
        'so, 0,3,4, and 5.
        'Withparser, Expression, Varname, InitialValue, IncrementExpression, TerminateExpression
        NoParsedArgs(0) = True
        NoParsedArgs(1) = True
        NoParsedArgs(3) = True
        
        NoParsedArgs(4) = True
        'there we are.
        Case UCase$(funcName) = "DUMPEXPRESSION"
            NoParsedArgs(0) = True
    End Select
        If InStr(CompleteParsedFunctions, " " & UCase$(funcName) & " ") <> 0 Then
            IEvalEvents_CanHandleFunction = True
        End If
End Function



Private Function IEvalEvents_CanHandleOperator(withparser As CParser, ByVal OpCheck As String, ByRef Unary As IEvalUnaryOpFlags) As Boolean
 'return wether this parser plugin can handle this operator.
    'this helps the parser decide if something is an operator or not.
    OpCheck = UCase$(OpCheck)
    If InStr(1, mvarDisabledOperators, OpCheck, vbBinaryCompare) <> 0 Then
        'disabled!
        IEvalEvents_CanHandleOperator = False
        CDebug.Post "ignored query for handling of operator, " & """" & OpCheck & """."
        Exit Function
    End If
        
        IEvalEvents_CanHandleOperator = InStr(1, HandledOpConst, " " & Trim$(OpCheck) & " ", vbBinaryCompare) <> 0
        'if it is a -, it can also be unary.
        If Unary And UnaryOp_Prefix Then
            IEvalEvents_CanHandleOperator = False
            'unless it is NOT, or -
            IEvalEvents_CanHandleOperator = InStr(1, UnaryOpConst, " " & OpCheck & " ", vbBinaryCompare) <> 0
        ElseIf Unary And UnaryOp_Suffix Then
            IEvalEvents_CanHandleOperator = False
            IEvalEvents_CanHandleOperator = InStr(1, UnaryOpSuffix, " " & OpCheck & " ", vbBinaryCompare) <> 0
        End If
    
   ' Case "@"
        '
        'NoParseArgs = True  'do NOT parse the arguments.
        'the arguments, my friend, are the Object <- and the method/property.
        
    
End Function



Private Property Get IEvalEvents_Description() As String
IEvalEvents_Description = "BASeParser Core Operator and Function Plugin"
End Property


Private Function IEvalEvents_GetFunctionFlags(withparser As CParser, ByVal FuncString As String) As FunctionFlags
  Select Case UCase$(FuncString)
        Case "RANDOM", "SHELL", "CREATEOBJECT", "GETOBJECT", "COUNTIF", "SOLVE", "STORE", "MSGBOX", "INPUTBOX", "INPUTBOX$", "FILTER"
         IEvalEvents_GetFunctionFlags = Function_Implemented
        Case Else
         IEvalEvents_GetFunctionFlags = Function_Optimizable + Function_Implemented
    End Select
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As CParser) As String
IEvalEvents_GetHandledFunctionString = CompleteParsedFunctions
End Function

Private Function IEvalEvents_GetHandledOperatorsString(withparser As CParser) As String
IEvalEvents_GetHandledOperatorsString = HandledOpConst & " " & UnaryOpConst & UnaryOpSuffix

End Function


Private Function IEvalEvents_GetInfoObject() As IOpFuncInfo
Set IEvalEvents_GetInfoObject = Me
End Function
Private Sub RecurseArrayOp(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval)
    'performs a scalar operation on the array.
    'IE, for each element in opa, adds to the elements in opb, etc...
    
    'BUGFIX: this procedure was calling the recursive call to GetOperation
    'with the array element ALWAYS first. This worked most of the time- until I tried
    'to use operators whose behaviour was different when the order was reversed, or, worse, when they
    'weren't the same type- (can you say "REPEAT operator")?
    
    Dim CurrScalar As Variant
    Dim I As Long, flipargs As Boolean
    Dim arrCopy, OpWith
    Dim A As Long, B As Long
    If IsArray(OpA) Xor IsArray(OpB) Then
        If IsArray(OpA) Then
        'only OpA is an array.
        'ReDim retval(UBound(OpA))
            arrCopy = OpA
            OpWith = OpB
        Else
            'logically it must be OpB that is an array.
            arrCopy = OpB
            OpWith = OpA
            flipargs = True
            'We must flip the arguments when we call "PerformOperation".
            
        End If
        ReDim retval(LBound(arrCopy) To UBound(arrCopy))
        For I = LBound(arrCopy) To UBound(arrCopy)
            'now, perform the same operation between the array element and opwith.
    '                Call IEvalEvents_GetOperation(WithParser, OpUse, ArrCopy(I), OpWith, currItem, retval(I))
    '02/12/07: changed "IEvalEvents_GetOperation" call to a more appropriate and flexible
        '"WithParser.PerformFunction" method call. This will allow more flexibility. I suppose.
            If Not flipargs Then
                'don't flip the arguments.
                withparser.PerformOperation OpUse, arrCopy(I), OpWith, retval(I)
                'tada.
            Else
                'err- DO flip the arguments.
                withparser.PerformOperation OpUse, OpWith, arrCopy(I), retval(I)
            End If
        Next I
    Else
        'Otherwise, they must BOTH be arrays.
        
            'Woop. there. now we need to loop through each element, assigning the
            'items in this new array as appropriate.
            'ha ha. bug here. should be creating an array of arrays.
            ReDim retval(LBound(OpA) To UBound(OpA))
            For A = LBound(OpA) To UBound(OpA)
                ReDim CurrScalar(LBound(OpB) To UBound(OpB))
                For B = LBound(OpB) To UBound(OpB)
                    'CurrScalar = retval(A)
                    
                    
                    Call IEvalEvents_GetOperation(withparser, OpUse, OpA(A), OpB(B), currItem, CurrScalar(B))
                    
                    
                Next B
                'there. assign this array to the array element.
                retval(A) = CurrScalar
            Next A
    
    End If
    
    
    







End Sub

Private Sub IEvalEvents_GetOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, ByVal OpB As Variant, currItem As CFormItem, retval As Variant)
  'performs the operation.
    'if the operation cannot be dealt with, return CvErr(GPRV_NOTHANDLED)
    '
    Dim isArrAcc As Boolean
    Dim Ret As Variant, Temp As Variant, I As Long
    Dim castcomplex As CComplex
    Dim subexecute As String
    Dim AssignMe As CVariable   'for assignment ops.
'CDebug.Post "was asked to """ & OpUse & " values.," & OpA & " And " & OpB

 
'ARRAY handling- array handling is implemented for all operators.

If CBool((VarType(OpA) And vbArray) Or CBool(VarType(OpB) And vbArray)) Then
    'array-
    CDebug.Post "Array Operand Detected for operator """ & OpUse & """"
    'some operators can work with arrays. A prime example is the assignment operator, := (thanks, Pascal :-) )
    If InStr(1, ArrayOperations, " " & OpUse & " ", vbTextCompare) = 0 Then
        CDebug.Post "Operator " & OpUse & " Does not support array operands. disambiguating Array reference into elements..."
        RecurseArrayOp withparser, OpUse, OpA, OpB, currItem, retval
        Exit Sub
    Else
        CDebug.Post "Operator " & OpUse & " supports Array operators directly."
    End If
End If

'Check for assignment operators- since they
'don't really care WHAT type the rval is, as long as
'the lval is a variable.
'Oh, and it checks for that.
If Not currItem.Prev Is Nothing Then
    If Not currItem.Prev.Prev Is Nothing Then
    'we need to look back TWO values, because the parser will have
    'rearranged the original Infix notation to RPN, so our arguments *SHOULD*
    'be the last two values.
        'oh- there IS a previous item, but is it a variable?
        
            
        
        
        
        
        If currItem.Prev.Prev.ItemType = IT_VARIABLE Or _
        currItem.Prev.Prev.ItemType = IT_ARRAYACCESS Then

        'Then    'the assignment operators can only operate
            'on IT_VARIABLE types.
            'TODO: extend support
            'to Arrays,
            'by checking for IT_ARRACCESS, and, if so,
            'grabbing those subscripts and looking
            'backwards on the list for
            'the first IT_VALUE, making it an array, and assigning to the
            'specified item (that's a mouthful).
            '--4/26/2007 @ 6:18--
            'I looked over the above comment, and actually COMPLETELY forgot that I had already implemented
            'the array assignment! I was just thinking what I would need to change, then noticed
            'the "ElseIf currItem.Prev.Prev.ItemType=IT_ARRAYACCESS Then" line.
            
            
            If currItem.Prev.Prev.ItemType = IT_VARIABLE Then
                Set AssignMe = withparser.Variables.Item(currItem.Prev.Prev.op)
            ElseIf currItem.Prev.Prev.ItemType = IT_ARRAYACCESS Then
                'create a dummy CVariable here, which we assign to assignme.
                'later, we check to see again wether this was a dummy, and if so do the
                'actual work of assignment.
                Set AssignMe = New CVariable
                AssignMe.Name = "SOMEDUMMY"
                isArrAcc = True
            End If
            Select Case True

              
                Case UCase$(OpUse) = "GOTO"
                Dim getlabel As CFormItem
                Set getlabel = Labelled.Item(OpA)
                Set currItem = getlabel
                 
                
                Case UCase$(OpUse) = ":="
               
                    'Assignment.
                    'Recursed upon by the +=,-=, etc.
                    'the OpB is the value we want. OpA is just the variable we are assigning to.
                    If IsObject(OpB) Then Set AssignMe.Value = OpB Else AssignMe.Value = OpB
                      If IsObject(AssignMe.Value) Then Set retval = AssignMe.Value Else retval = AssignMe.Value
                    Exit Sub
                    'Friday July 4th 2008 8:39
                    'Not sure why exit sub was commented out here, seemed to prevent it from working properly.
                    
                'more cases, for +=,-=,*=,^=,etc.
                'Case "+=", "-=", "*=", "/=", "\=", "^="
                

                
                'if the last character is a "=", then it is one of +=,-=, etc..
                Case right$(OpUse, 1) = "="
                    subexecute = Mid$(OpUse, 1, InStrRev(OpUse, "="))
                    'this and -= are REALLY clever.
                    'why do the work when we don't have to-
                    'performOperation +, and then :=.
                    withparser.PerformOperation subexecute, AssignMe.Value, OpB, retval
                    'now, do the assignment recursively.
                    'WithParser.PerformOperation ":=", AssignMe.Value, retval, retval
                     If IsObject(OpB) Then Set AssignMe.Value = OpB Else AssignMe.Value = OpB
                    'retval is the new value-\
                    'Set AssignMe = WithParser.Variables.Item(currItem.Prev.Prev.op)
                    If IsObject(retval) Then Set AssignMe.Value = retval Else AssignMe.Value = retval
                
                
                  If IsObject(AssignMe.Value) Then Set retval = AssignMe.Value Else retval = AssignMe.Value
                    Exit Sub
                
            End Select
            'return the value of assignme.
            If isArrAcc Then
                'not so fast!
                'since the variable we assigned was simply a dummy, we need to
                'look back, not TWO, but three.
                If currItem.Prev.Prev.Prev.ItemType = IT_VARIABLE Then
                    With withparser.Variables.Item(currItem.Prev.Prev.Prev.op)
                    'PRB discovered: when the variables value is an array, we cannot
                    'directly assign to it because it will only assign the temporary return value.
                        .SetElement currItem.Prev.Prev.Value(0).Execute, AssignMe.Value
                    End With
                End If
                
                
            Else
                
                '    Err.Raise 5, "BPCoreOpFunc::GetOperation", "Failed to set value of literal."
                Assign withparser.Variables.Item(currItem.Prev.Prev.op).Value, OpB
                
            End If
                
            
            
            
        End If
    End If
End If







'internal Handling, for internal Ops.
'" - + * \ / ^ ** == != <> <= =< >= >= "
'

    'don't drop to the Case yet-
    'here we have one last thing we should try.
    'If either item is still an object, we should try to Cast the object to a
    '"IOperable" interface (I should define it) If this test succeeds, we have an object that
    'supports BASeParser. We can then call the PerformOp() Member of that interface, giving the parser,
    'the other operand, and other basic information, and using the ByRef return argument as our result.
    
    'If Either or both of the items are a Object, we need to do the following:
    If IsObject(OpA) Or IsObject(OpB) Then
        'if either or both are objects, call the sub...
        Call ObjectOperation(withparser, OpUse, OpA, OpB, retval)
        Exit Sub
    End If
    
    
    'convert string as necessary.
    If (VarType(OpA) = vbString) Xor (VarType(OpB) = vbString) Then
        If VarType(OpA) = vbString Then
            If Val(OpA) = 0 And OpA <> "0" Then
                OpB = CStr(OpB)
            Else
                OpA = CStr(OpA)
            End If
        ElseIf VarType(OpB) = vbString Then
            If Val(OpB) = 0 And OpB <> "0" Then
                OpA = CStr(OpA)
            Else
                OpB = CStr(OpB)
            End If
        End If
    
    End If
    
    
    
    
        Select Case UCase$(OpUse)
        Case "=~"
                'match opA using regular Expression OpB.
                Dim matchGot As MatchCollection, castRe As RegExp
                Dim useRe As String
                If VarType(OpA) <> vbString Then OpA = withparser.ResultToString(OpA)
                If TypeOf OpB Is VBScript_RegExp_55.RegExp Then
                    Set castRe = OpB
                    
                    Set matchGot = castRe.Execute(OpA)
                    Set Ret = matchGot
                ElseIf VarType(OpB) <> vbObject Then
                    Set castRe = New RegExp
                    castRe.Pattern = OpB
                    castRe.MultiLine = True
                    castRe.Global = True
                    Set matchGot = castRe.Execute(OpA)
                    Set retval = matchGot
                
                
                    
                
                
                End If
                
           Case ":="
            'assignment operator.
            'hmm... this WAS working.
            
            'prev three times. huh... WTF?
            Debug.Assert False
            If currItem.Prev.Prev.Prev.ItemType <> IT_VARIABLE Then
                'extend to support assignment to properties of objects. Sounds complicated though.
            
            
            Else
                withparser.Variables.Add currItem.Prev.Prev.Prev.op, OpB
                retval = OpB
            
            End If
            
            
                
                '+,-,* and \ are defined for Complex numbers.
                'oh, and ==.
            Case "$$"
                'Regular expression- just a test, right now.
                Static Matcher As RegExp
                Dim matched As MatchCollection
                
                'First operand is string, second is what to operate on.
                'the $$ operator has a precedence higher even then @, or, that is what
                'I want it to have.
                Set Matcher = New RegExp
                Matcher.Global = True
                Matcher.IgnoreCase = True
                Matcher.MultiLine = True
                Matcher.Pattern = OpB
                Set matched = Matcher.Execute(OpA)
                Set retval = matched
                'convert this Collection into an array. I don't know why.\
'                If matched.Count > 0 Then
'                ReDim retval(0 To matched.Count - 1)
'                For I = 0 To matched.Count - 1
'                    Set retval(I) = matched.Item(I)
'                Next I
'                Else
'                    ReDim retval(0)
'                    Erase retval
'                End If
            Case "+"
                
              
                retval = OpA + OpB
            Case "-"
                retval = OpA - OpB
            Case "/"
                retval = OpA / OpB
            Case "\", "DIV"
                retval = OpA \ OpB
            Case "MOD", "//"
'var MOD power of 2 constant
'can be var AND constant-1
                        
                'retval = OpA Mod OpB
                If OpA < -32767 Or OpA > 32768 Or _
                    OpB < -32767 Or OpB > 32768 Then
                retval = Modulus(OpA, OpB)
                Else
                    retval = OpA Mod OpB
                End If
                
            Case "*"
            'if the first operand is a string that cannot be converted to a number, then we will repeat that string the given number of times.
            'note that oddly enough I'll allow for fractional components here.
            
            
                retval = OpA * OpB
            Case "UPTO"
                'return an array from OpA to OpB...
                ReDim retval(OpB - OpA)
                For I = OpA To OpB
                    retval(I - OpA) = I
                
                Next I
            
            
            Case "^", "**"
                'special handling for Complex numbers (ie, taking to the -0.5 power) is handled
                'in CComplex.
                'I was going to add it here. Then I realized that that is stupid. HA!
            
            
            
                retval = OpA ^ OpB
            Case "==", "="
                retval = (OpA = OpB)
            Case "<>", "!="
                retval = OpA <> OpB
            Case "<=>"
                'Spaceship operator- thanks perl!.
                retval = compare(withparser, OpA, OpB)
            Case "<=", "=<"
                retval = (OpA <= OpB)
            Case ">=", "=>"
                retval = OpA >= OpB
            Case ">"
                retval = OpA > OpB
            Case "<"
                retval = OpA < OpB
             Case "AND", "&&"
                retval = OpA And OpB
             Case "OR", "||"
                retval = OpA Or OpB
             Case "XOR"
                retval = OpA Xor OpB
             Case "EQV", "="
                retval = OpA Eqv OpB
             Case "IMP"
                retval = OpA Imp OpB
                '(NOT a) OR b.

             'Case "NOT"
             Case "IS"
                retval = OpA Is OpB
            Case "REPEAT"
                'repeats a given string a specified number of times.
                'Yes I DO have to many operators of questionable usefulness.
                retval = Repeat(OpA, OpB)
             Case "LIKE"
                retval = OpA Like OpB
             Case "IN"      'take that, Delphi/Pascal.
                retval = Among(withparser, OpA, OpB)
            Case UCase$(OpUse) = ">>"
                    'write opA into OpB. (binary)
                    If TypeName(OpB) = "FileStream" Then
                        OpB.WriteString OpA
                    
                    End If
            Case ":"
                'This is my own operator. I'll be honest- this is somewhat of a cheat.
                'to support the Ternary conditional Expression.
                'That is why the : has a higher precedence then ?, so that when the evaluation engine
                'reaches the ?, this will have been evaluated and created a (shh- it's a secret) Array.
                'an Array? Yes, it is a fairly kludgy type of workaround simply to support the ?, which
                'nobody seems fond of anyway. Oh well.
                'I'll explain further in the devlog.
                
                
                
                ReDim retval(0 To 1)
                    'wow. this is probably the simplest operator ever.
                    Call Assign(retval(0), OpA)
                    Call Assign(retval(1), OpB)
                    'wow.
                    
                
                
                
            Case "?"
                'SHIVER! the ternary operator
                'has snuck out of it's C++ cave, and
                'attacked!
                'it's a secret- BASeParser doesn't support ternary operators- so what do I do?
                'I implement a relatively gross hack, involving the : converting it's operands to an array.
                
                'OK, first, we kind of need to ensure that the second operand is an array with at least two elements.
                'if not, raise error 5.
                Static TernBool As Boolean
                If Not IsArray(OpB) Then
                    
                    
                    
                    
                    
                    Err.Raise 5, "BPCoreOpFunc::GetOperation", """?"" operator needs Array operand to it's right."
                Else
                   ' On Error Resume Next
                    TernBool = CBool(OpA)
                    If OpA Then
                        'return first Array element.
                        Call Assign(retval, OpB(LBound(OpB)))
                    Else
                        Call Assign(retval, OpB(LBound(OpB) + 1))
                    End If
                    
                
                
                End If
                
                
                
                
                
                
                'Ahh, and who could forget?
                'A couple interesting operators. I could have just as easily made them functions,
                'but, I mean, 7 PICK 3 looks so natural, doesn't it?
            Case "NPR"
                'Number of possible combinations.
                retval = Permut(OpA, OpB)
            
            Case "NCR"
                retval = Comb(OpA, OpB)
            Case "PICK"
                'I just thought of this operator- sometimes, you'll have an array and you'll want
                'to choose a certain number of those array values and make another array.
                'for example, the expression:
                '{1,2,3,4,5,6,7} SELECT 2
                'will return a new array of size 2 with two random items chosen from the first array.
                'IMPORTANT!
                'only the first item can be an array.
                'if the second operator is an array, then we should
                'create an array. for example,
                '"{1,2,3,4,5,6} SELECT {1,3}"
                'should return an array, with the first element being another array which is the SELECT 1, and
                'another that is "Select 3"
                'Yes, I know it is a toughie.
                'the trick is to use the same logic, so the default operation is actually
                'to use an array :).
            
                retval = SelectItems(withparser, OpA, OpB)
            
            
        End Select
    
    'End If  '(ComplexCalc)


End Sub
Private Sub ObjectOperation(withparser As CParser, Operation As String, OpA As Variant, OpB As Variant, ByRef retval)
    'procedure called by GetOperation to handle when an operator is applied to an object.
     Dim TestCast As IOperable, TestMe As Object
     Dim Casted(0 To 1) As IOperable
     Dim OtherOp As Variant
    On Error Resume Next
            
    If IsObject(OpA) Xor IsObject(OpB) Then
        'Only one can be true.
        'OK- which one?
        If IsObject(OpA) Then
            Set TestMe = OpA
            OtherOp = OpB
        Else
            Set TestMe = OpB
            OtherOp = OpA
        End If
        'Here we go- attempt to cast the variable "Testme" to an IOperable interface.
        Err.Clear
        Set TestCast = TestMe
        If Err <> 0 Then
            'Since we now know that the object in question does NOT support the Ioperable interface, we should
            'raise an error.
            Err.Raise 5, "BPCoreOpFunc::ObjectOperation", "Object Expression could not be evaluated. The object type """ & TypeName(TestMe) & """. does not implement the Ioperable interface."
            
        Else
            'whew.
            CDebug.Post "Succeeded casting Object Operand to Ioperable interface."
            'we succeeded. That's a load off.
            'now we simply take the casted item and call it's performOp Method.
            'Of course, we aren't quite off the hook- Since <Gulp>- the object might
            'not recognize the particular operand we are giving it, or may not support it
            '(For example, CComplex doesn't support strings)
           ' On Error GoTo 0     'if the interface does raise an error, it is beyond our control.
            
            
            
            
            
            
        End If
        
        
        
        
        
    ElseIf IsObject(OpA) And IsObject(OpB) Then
        'This one could be a bit trickier.
        'If only one of the objects supports the IOperable interface, we are scot-free.
        On Error Resume Next
        Set Casted(0) = OpA
        Set Casted(1) = OpB
        If Casted(0) Is Nothing Xor Casted(1) Is Nothing Then
            'if one is nothing and the other isn't.
            'this is easier. simply get the one that IS,put it in
            'testcast, and throw the other in OtherOp.
            If Casted(0) Is Nothing Then
                Set TestCast = Casted(1)
                Set OtherOp = OpA     'use Set- we KNOW they are both object expressions.
            Else
                Set TestCast = Casted(0)
                Set OtherOp = OpB
            End If
            
        ElseIf (Not Casted(0) Is Nothing) And Not (Casted(1) Is Nothing) Then
        
            'This one herrm- we are getting really deep here.
            'BOTH of the provided objects support the "IOperable" interface.
            'what does this mean? Well, we have an ambiguous operation, really. Since they both could possibly
            'support the respective operation. The Solution? Set an error trap, try to call the first one/
            'if it succeeds, break. OtherWise, disable the error handler and call the second one.
            On Error Resume Next
            Err.Clear
            Call Casted(0).PerformOp(withparser, Operation, OpB, UnaryOp_None, retval)
            If Err = 0 Then
                'the first call succeeded, and as such we can continue our merry way.
                Exit Sub
            Else
                'No such luck. disable the error handler and try to call the second one.
                'if that one raised an error, it will go right past this procedure, and probably all the
                'way up to CollapseStack.
                On Error GoTo 0 'On purpose.
                Call Casted(1).PerformOp(withparser, Operation, OpA, UnaryOp_None, retval)
                Exit Sub
            End If
        
        
        
        
        End If
    
    
    
    
    
    End If
            
            
    
    'Check if one of the objects implements the IOperable interface.
        'If neither one does, raise an error.
        'If only one does:
            'Call that interfaces PerformOp Method, and return any result.
        
        'If they both do- we have our work cut out for us.
    On Error GoTo 0
    TestCast.PerformOp withparser, Operation, OtherOp, UnaryOp_None, retval

End Sub
Private Function SupportsIOperable(ObjTest As Variant) As Boolean

    If TypeOf ObjTest Is IOperable Then
        SupportsIOperable = True
    Else
        SupportsIOperable = False
    End If
End Function


Private Sub IEvalEvents_GetUnaryOperation(withparser As CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As CFormItem, retval As Variant)
'
Dim NextItem As CFormItem
Dim beforeIncrement
Const IncrementAmount As Integer = 1
Dim varval As Variant, workitem As CFormItem
If currItem.Tag = "SUFFIX" Then
    Set workitem = currItem.Prev
    If workitem Is Nothing Then
        Err.Raise 5, "BPCoreOpFunc::GetUnaryOperation", "suffix-type Unary operator not valid without an operand."
    End If
Else
    Set workitem = currItem.Next_
    If workitem Is Nothing Then
        Err.Raise 5, "BPCoreOpFunc::GetUnaryOperation", "Prefix-type Unary operator not valid without an operand."
    End If
End If
Set NextItem = currItem.Next_
Select Case UCase$(OpUse)
    Case "<<"
        If TypeName(OpA) = "FileStream" Then
            retval = OpA.ReadLine()
        End If
    
    Case "GOTO"
    
       Dim getlabel As CFormItem
       Set getlabel = Labelled.Item(OpA)
       Set currItem = getlabel
                
    Case "NOT"   ', "!"
        retval = Not OpA
    Case "-"
        If IsObject(OpA) Then
            If TypeOf OpA Is CComplex Then
                Set retval = OpA.Negate
            End If
        Else
            retval = -OpA
        End If
    Case "SIZEOF"
        retval = Len(OpA)
    Case "!"
       
        retval = Factorial(OpA)
    
    
    Case "SORT"
        'Unary Operation.
        'sorts it's operand, which MUST be an array.
        'of course- we check for it.
        ForceArray OpA
        SortArray OpA
        retval = OpA
    Case "--", "++"  'it worked?
                'prefix:
      
                
                'this is the prefix decrement operator- as such,
                'err- decrement the value. duh.
                'also- make sure it is actually a Variable-
                If workitem.ItemType <> IT_VARIABLE Then
                    'they might have meant double negative.
                    Call Assign(retval, OpA)
                Else
                
                
                   
                    Call Assign(varval, withparser.Variables.Item(workitem.op).Value)
                    'Whoa!
                    
                    Call Assign(beforeIncrement, varval)
                    
                   
                    
                    
                    'increment the value:
                    'clever boolean expression here.
                    
                      'make it negative if the op is "--", otherwise, keep it positive.
                        IncrementValue withparser, varval, IncrementAmount * Sgn(((OpUse = "--") * 2) + 1)
                    'assign back to variable.
                    withparser.Variables.Item(workitem.op).Value = varval
                    'ok, if we are a suffix-type, we get added then returned.
                    'otherwise we get returned and added.\
                    If currItem.Tag = "SUFFIX" Then
                        retval = varval
                        
                    Else
                        'assign to value BEFORE increment.
                        retval = beforeIncrement
                    End If
                    
                    
                End If
        
        

    
End Select
End Sub
Private Sub IncrementValue(withparser As CParser, ByRef Value As Variant, ByVal ByAmount)
'delegate to parser.
    Call withparser.PerformOperation("+", Value, ByAmount, Value)

'If IsNumeric(Value) Then
'    Value = Value + ByAmount
'ElseIf VarType(Value) = vbString Then
'
'
'End If
End Sub
Private Function ListMembers(OfObject As Object) As Variant

    Dim tliuse As TLIApplication
    Dim RetArray As Variant, retcount As Long
    Dim LoopMember As MemberInfo
    Set tliuse = New TLIApplication
    
    Dim IInfo As InterfaceInfo
    Set IInfo = tliuse.InterfaceInfoFromObject(OfObject)
    
    
    For Each LoopMember In IInfo.Members
        retcount = retcount + 1
        ReDim Preserve RetArray(1 To retcount)
        RetArray(retcount) = LoopMember.Name
    
    
    Next
    
    ListMembers = RetArray





End Function
Private Sub IEvalEvents_HandleFunction(withparser As CParser, ByVal funcName As String, FuncParams() As Variant, ReturnVal As Variant)

    'performs the function.

    Dim ProgIDUse As String
    Dim Ret As Variant, paramcount As Long, I As Long
    Dim P() As Variant, varGrab As CVariable
    Dim parsetemp As CParser
    Dim arrtemp As Variant
    P = FuncParams      'yes, I know it sort of wastes memory.
                        'but I doubt the parameter
                        'list will get long enough to
                        'make a difference. I mean, seriously. Memory isn't really
                        'an issue, even on my old (1999) machine, which has 512MB.
                        'My '96 laptop has 144MB. Yes my laptops have year models.
                        'I also have a '82 Mouse that had a ball upgrade to a '84.
                        'Now it blows blue smoke whenever I click.
                        
                        
    

    
 
    On Error Resume Next
    paramcount = UBound(FuncParams) + 1
    If Err <> 0 Then paramcount = 0
    Err.Clear
    On Error GoTo FunctionError
    'error 450: wrong number of arguments or invalid property assignment
    'this is only used internally so we can break to the FunctionError label.
    'Once we get there, we create a new "ParserError" object and throw it.
    If InStr(App.CompanyName, "BASeCamp") = 0 Then
    'ahh, now WHY does our company Name String not contain the word "BASeCamp"?
            Ret = "This code belongs to BASeCamp Corporation! THIEF!THIEF!THIEF!"
            Exit Sub
    'ElseIf InStr(1, App.CompanyName, "not BASeCamp", vbTextCompare) Then
    
    End If
    Dim ObjCast As Object
    
    Const ArrayFunctions = " ABS SIN COS TAN ATN "
    
    '"ArrayFunctions" space delimited list of functions that are compatible with this method of "expanded" parameter handling. This includes essentially those functions
    'that take a single argument.
        Dim buildarr As Variant, passarr()
    If InStr(1, ArrayFunctions, UCase$(funcName), vbTextCompare) > 0 And paramcount > 1 Then
     If paramcount = 0 Then GoTo NotOptional
            
            If paramcount > 1 Then
            ReDim buildarr(0 To paramcount - 1)
            ReDim passarr(0)
            On Error Resume Next
            For I = 0 To paramcount - 1
                passarr(0) = P(I)
                
                Call withparser.EvalListener.Self.HandleFunction(withparser, funcName, passarr(), buildarr(I))
                If Err.Number <> 0 Then
                   Set buildarr(I) = ParserError(withparser, Err.Number, Err.Description, funcName, 0)
                
                End If
            Next I
             Ret = buildarr
            Else
                Ret = Abs(P(0))
            
            End If
    
    Else
    
    
       Select Case UCase$(funcName)
            'trigonometric functions.
            '" ABS ARRAY ASC ATN " & _
        "CHOOSE CHR CHR$ COS CREATEOBJECT CVERR " & _
        "DATEADD DATEDIFF DATEPART DATESERIAL DATEVALUE DAY DDB DIR DOEVENTS " &
            Case "GOTO"
            Debug.Print "GOTO"
            Case "VERSION"
              Ret = "BASeParser Expression Evaluation Library Version " & Trim$(App.Major) & "." & Trim$(App.Minor) & " Revision " & Trim$(App.Revision)
           Case "DUMPEXPRESSION"
            'DumpExpression: dumps the token list created for a specified expression as a String.
            Set parsetemp = withparser.Clone
            parsetemp.ParseInfix FuncParams(0)
            Ret = parsetemp.FirstItem.GetString(withparser)
            Case "WITH"
            If paramcount < 2 Then GoTo NotOptional
             Set parsetemp = withparser.Clone
             Set parsetemp.ImpliedObject = P(0)
             parsetemp.ParseInfix FuncParams(1)
             parsetemp.ExecuteByRef Ret
            Case "LISTMEMBERS" '
                Dim castobj As Object
                Set castobj = P(0)
                If paramcount < 1 Then GoTo NotOptional
                Assign Ret, ListMembers(castobj)
            Case "LISTVARIABLES"
            Ret = ""
            Dim LoopVar As CVariable
            For Each LoopVar In withparser.Variables
                Ret = Ret & LoopVar.Name & " = " & withparser.ResultToString(LoopVar.Value) & vbCrLf
            
            Next LoopVar
            
            
            
            
            Case "INSTALLPLUGIN"
                
                
                    'Install a plugin.]
                    'Parameters:
                    'INSTALLPLUGIN("Plugin.ProgID"|Object)
                    'installs the given plugin. If an object is passed
                    'we retrieve the progID and work with that.
                    If IsObject(P(0)) Then
                        Set ObjCast = P(0)
                        ProgIDUse = GetObjectProgID(ObjCast)
                    Else
                        ProgIDUse = CStr(P(0))
                    End If
                    'simply delegate to god ol' ParserSettings.
                    'install it into the current parser's configset.
                    ParserSettings.InstallPlugin ProgIDUse, withparser.configset
                    
        '            If paramcount = 2 Then
        '                If CBool(P(1)) Then
        '                    withparser.EvalListener.AddImplementor createObject(ProgIDUse)
        '                End If
        '            End If
                    Ret = True
                Case "LOADPLUGIN"
                  'load a plugin
                  'parameter is progID or actual object.... use same ol boilerplate code
                     If IsObject(P(0)) Then
                        Set ObjCast = P(0)
                        ProgIDUse = GetObjectProgID(ObjCast)
                    Else
                        ProgIDUse = CStr(P(0))
                    End If
                    
                    withparser.EvalListener.AddImplementor createObject(ProgIDUse)
                    CDebug.Post "loaded plugin dynamically:" & ProgIDUse
                Case "UNINSTALLPLUGIN"
                    
                    
                    If IsObject(P(0)) Then
                        Set ObjCast = P(0)
                        ProgIDUse = GetObjectProgID(ObjCast)
                    Else
                        ProgIDUse = CStr(P(0))
                    End If
                
                
                ParserSettings.UnInstallPlugin ProgIDUse, withparser.configset
                
                Case "UNLOADPLUGIN"
                    'remove the plugin. Again, could be progID or Actual object.
                    '"unload" in this context meaning to remove it from the collection of plugins.
                    
                Case "LISTPLUGINS"
                'list all loaded plugins. Returns a string.
                Dim BuildIt As cStringBuilder
                Set BuildIt = New cStringBuilder
                Dim tempInfo As IPluginUIData
                Dim LoopEval As IEvalEvents, LoopCore As ICorePlugin
                BuildIt.Append "Evaluation Plugins:" & vbCrLf & "--------------" & vbCrLf
                For Each LoopEval In withparser.EvalListener.GetCollection
                    BuildIt.Append LoopEval.Name & "("
                    On Error Resume Next
                    Set tempInfo = LoopEval.GetPluginUIData
                    If tempInfo Is Nothing Then
                        BuildIt.Append "No Information Provided)"
                    Else
                        BuildIt.Append tempInfo.Description & ")"
                    
                    End If
                    BuildIt.Append vbCrLf
                Next
                BuildIt.Append "Core Plugins" & vbCrLf & "-------------" & vbCrLf
                For Each LoopCore In withparser.CorePlugins.Col
                    Set tempInfo = LoopCore.GetPluginUIData
                    
                
                Next
                Ret = BuildIt.ToString
                Case "COBJECT"
                'simply return a new Cobject.
                'The CObject class is quite, odd- check out CObject.cls for details.
                    Set Ret = New CObject
                Case "CINT", "CLNG", "CSNG", "CDBL", "CCUR", "CVAR", "CDATE", "CBYTE"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Switch(UCase$(funcName), CInt(P(0)), UCase$(funcName) = "CINT", _
                                         CLng(P(0)), UCase$(funcName) = "CLNG", _
                                        CSng(P(0)), UCase$(funcName) = "CSNG", _
                                         CDbl(P(0)), UCase$(funcName) = "CDBL", _
                                         CCur(P(0)), UCase$(funcName) = "CCUR", _
                                        CVar(P(0)), UCase$(funcName) = "CVAR", _
                                         CDate(P(0)), UCase$(funcName) = "CDATE", _
                                         CByte(P(0)), UCase$(funcName) = "CBYTE")
                                        
                                        
                
                Case "FRAC"
                'new- Fractions are representable by objects.
                'kind of silly.
                    'create a fraction.
                    'if a single argument was given, use setvalue.
                    Dim newfrac As CFraction
                    Set newfrac = New CFraction
                    With newfrac
                    If paramcount = 0 Then GoTo NotOptional
                        If paramcount = 1 Then
                            newfrac.SetValue P(0)
                        ElseIf paramcount = 2 Then
                            newfrac.SetValues P(0), P(1)
                        
                        
                        
                        
                        End If
                    End With
                    Set Ret = newfrac
                Case "ABS"
                    If paramcount = 0 Then GoTo NotOptional
                
                    Ret = Abs(P(0))
                    
       
                    
                Case "ARRAY"
                    'ha HA HA! this is a heck of a lot easier and
                    'faster then it was in my FIRST baseParser.
                    'If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 0 Then
                        'handle "Empty list"...
                        'make it an EMPTY array.
                        'hee hee! just return the array of parameters.
                        'couldn't be simpler.
                        
                        ReDim Ret(0)
                        Erase Ret
                    
                    Else
                        Ret = P
                    End If
                    
                    
                Case "ASC", "ORD"
                    'no-brainer
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Asc(P(0))
                Case "ATN"
                    'doesn't require intelligence.
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Atn(P(0))
                Case "CHOOSE"
                    'a little bit more intricate, but generally easy nonetheless.
                    'Choose, index, Paramarray choice.
                    'OK, this one is actually pretty easy.
                    'but, we want error handling GRRR
                   
                    'the index shouldn't be zero or they are in for a surprise- they'll end up returning- 0. that was their zero-th parameter.
                    'tada.
                    If paramcount < 2 Then GoTo NotOptional 'obviously if we have fewer
                    'this one looks cryptic.
                    'to some.
                    'the first argument is the index, and the first argument is P(0). As such,
                    'take element number "index". Index=P(0), so:
                    
                    Ret = P(P(0))
                Case "EXEC"
                    'This is a pretty cool function- it will execute the
                    'given command, and redirect all Standard output to
                    'the returned value.
                    'well- it will. Not yet though ;P
                    'I have to re-create my console capture component.
                    'Ret = Redirect(CStr(P(0)))
                    Dim UseCon As CMExec
                    Set UseCon = New CMExec
                    Ret = UseCon.Exec(P(0))
                    
                    
                    
                    '
                Case "WHILE"
                '"WHILE(Expression,<Statements>)"
                'Step one: make the two clones...
                Dim ParserExpression As CParser, ParserStatements As CParser
                Set ParserExpression = withparser.Clone
                Set ParserStatements = withparser.Clone
                ParserExpression.Expression = P(0)
                ParserStatements.Expression = P(1)
                Do While ParserExpression.Execute
                   ParserStatements.Execute
                Loop
                
                Case "FLATTEN"
                'flatten an array.
                'flattening an array coerces any elements in the array which are themselves arrays
                'to their respective values. For example:
                
                '{1,2,3,{4,5},1,{5,7,{7,4},2,3},4}
                'would be:
                '{1,2,3,4,5,1,5,7,7,4,2,3,4}
                'because the individual elements which are themselves arrays
                'are flattened as well.
                'just call the flatten() function. It'll do the work.
                    Ret = Flatten(P(0))
                    'although the above example makes it look simple- a simple removal of all but
                    'the first and last list delimiters, it is actually a recursive procedure.
                Case "MATRIX"
                'perhaps matrices should be affected by the aforementioned Flatten() function.
                'oh well.
                    Dim MakeMat As CMatrix
                    Dim Loopparam As Long
                    Set MakeMat = New CMatrix
                    Set MakeMat.Parser = withparser
                    'The given values are arrays- each parameter
                    'is another ROW to add to the matrix.
                    Call MakeMat.Resize(UBound(P(0)) + 1, UBound(P) + 1)
                    'manually assign each array item,
                    For Loopparam = 0 To UBound(P)
                        MakeMat.SetRow Loopparam + 1, P(Loopparam)
                    Next
                    
                    
                    Set Ret = MakeMat
                
                
                
                Case "COMPLEX"
                    'complex number.
                    'kind of silly, because i, the Sqr(-1), is a valid token. (well, it's a variable set to the
                    'value of COMPLEX(0,1), so we still use this code here.
                    'If complex number support is disabled, disallow any reference to this function.
                    'this means any reference to i will also cause an error (Complex number support disabled)
                    If withparser.Settings.GetPluginSetting(Me, "CplxDisable", withparser.configset) <> 0 Then
                        'raise the appropriate error.
                        'ParserError(WithParser, 17, "Complex Operations are Disabled.", "BPCoreOpFunc.HandleOperation").Throw
                    Err.Raise 17, "BPCoreOpFunc.HandleFunction", "Complex operations are disabled."
                    
                    Else
                        If paramcount = 1 Then
                            'if numeric, real component. otherwise, it's the STRINGFORM.
                            If IsNumeric(P(0)) Then
                                Set Ret = Complex(P(0))
                            Else
                                Set Ret = New CComplex
                                Ret.StringForm = P(0)
                            End If
                        ElseIf paramcount >= 2 Then
                            Set Ret = Complex(P(0), P(1))
                        End If
                    End If
                Case "CHR", "CHR$"
                'the Chr function will operate faster
                'when using and returning a variant.
                'rather then onverting a string to a variant. returning a string, and
        '        then converting that string to a variant...
                If paramcount = 0 Then GoTo NotOptional
                'Chr$ can't operate on an array, so..
                'Or- can it?
                    Dim tempret As String
                    P = Flatten(P)
                    tempret = Chr(P)
                    ReDim Ret(0 To Len(tempret) - 1)
                    For I = 1 To Len(tempret)
                        Ret(I - 1) = Mid$(tempret, I, 1)
                    Next I
                    If UBound(Ret) = LBound(Ret) Then
                        Ret = Ret(UBound(Ret))
                    End If
                Case "COS"
                    'Proof Egg.
                    
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Cos(P(0))
                Case "SIN"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Sin(P(0))
                Case "TAN"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Tan(P(0))
                Case "SEC"
                     If paramcount = 0 Then GoTo NotOptional
                     Ret = Sec(P(0))
                Case "ASEC"
                     If paramcount = 0 Then GoTo NotOptional
                     Ret = ArcSec(P(0))
                Case "HSEC"
                     If paramcount = 0 Then GoTo NotOptional
                     Ret = HSec(P(0))
                Case "HASEC"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = HArcSec(P(0))
                    'ASIN ACOS ATN HSIN HCOS HTAN HASIN HACOS HATAN
                Case "ASIN"
                     If paramcount = 0 Then GoTo NotOptional
                    Ret = ArcSin(P(0))
                Case "ACOS"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = ArcCos(P(0))
                Case "ATN"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = Atn(P(0))
                Case "HSIN"
                If paramcount = 0 Then GoTo NotOptional
                    Ret = hSin(P(0))
                Case "HCOS"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = hCos(P(0))
                Case "HTAN"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = hTan(P(0))
                Case "HASIN"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = HArcSin(P(0))
                Case "HACOS"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = HArcCos(P(0))
                Case "HATAN"
                 If paramcount = 0 Then GoTo NotOptional
                    Ret = HArcTan(P(0))
                Case "EVALARR"
                    
                    If paramcount < 2 Then GoTo NotOptional
                    Set parsetemp = withparser.Clone()
                    'EvalArr:
                    'EvalArr(Expression,VarName,Array)
                    Set varGrab = parsetemp.Variables.Add(P(1), 0)
                    'Evaluates the given expression for each item in the array,
                    'cycling the value of "varname" through each element of the array.
                    parsetemp.Expression = P(0)
                    arrtemp = P(2)
                    For I = 0 To UBound(arrtemp)
                        varGrab.Value = arrtemp(I)
                        parsetemp.ExecuteByRef arrtemp(I)
                    Next I
                    Ret = arrtemp
                Case "SELECT"
                    'SELECT:
                    'SELECT(Array,Numitems)
                    'same as PICK operator.
                    If paramcount < 2 Then GoTo NotOptional
                    Ret = SelectItems(withparser, P(0), P(1))
                
                
                
                Case "SEQ"  'This is the "Simpler" version
                    If paramcount = 4 Then
                        Ret = Seq(withparser, P(0), P(1), P(2), P(3))
                    Else
                        Ret = Seq(withparser, P(0), P(1), P(2), P(3), P(4))
                    End If
             
                    
                Case "SEQEX"
                    'SeqEx Withparser, Expression, Varname, InitialValue, IncrementExpression, TerminateExpression,IterationLimit
                    'TODO:// add hang-resistant code.
                    'if the WithParser variable
                    'is Constant AND evaluates to False, then
                    'raise an error, "Termination Expression is Constant False" or something.
                    If paramcount < 5 Then GoTo NotOptional
                    
                    'Check for the constant false termination.
                    If paramcount = 5 Then
                    
                    Ret = SeqEx(withparser, P(0), P(1), P(2), P(3), P(4))
                    
                    ElseIf paramcount >= 6 Then
                    
                        Ret = SeqEx(withparser, P(0), P(1), P(2), P(3), P(4), P(5))
                    End If
                    
                    
                Case "SOLVE"
                    'SOLVE:
                    'SOLVE(Expression,varname,start,end,[tolerance])
                    
                    
                    ' attempts to solve the given Expression for 0, using root-bisection on varname.
                    'uses the root-bisection method. Obviously of little use if the given function contains a call to
                    'a dynamic function or operator.
                    If paramcount < 4 Then GoTo NotOptional
                    If paramcount = 4 Then
                    Ret = Solve(withparser, P(0), P(1), P(2), P(3))
                    ElseIf paramcount = 5 Then
                    
                    Ret = Solve(withparser, P(0), P(1), P(2), P(3), P(4))
                    End If
                
                Case "COUNTIF"
                    Dim IfCount As Long
                    Dim CountIfParser As CParser
                    Dim varmanip As CVariable
        
                    'we need at least 2 arguments.
                    If paramcount < 2 Then GoTo NotOptional
                        
                    Set CountIfParser = New CParser
                    Set CountIfParser = withparser.Clone
                    
                    
                    
                    'like the Excel function.
                    'since we kind of don't have ranges,
                    'we have the following parameter list:
                    'CountIf(Expression,VarName,Value1,Value2,Value3)
                    
                    'hey- RANGES! ;)
                    'Expression: expression to evaluate for each one.
                    'a variable named Varname is added and assigned to each value given.
                    'if evaluating the expression results in True, then the return value will
                    'be incremented.
                    CountIfParser.Expression = P(0)
                    'add a variable. it will replace any existing variable with the same name,
                    'naturally raising the event.
                    'Of course, if the current set of Plugins have defined constants...
                    Set varmanip = CountIfParser.Variables.Add(P(1), 0)
                    'now, loop through the remaining elements, parsing the expression.
                    'Errors in this function are taken as False values.
                    On Error Resume Next
                    IfCount = 0
                    For I = 2 To UBound(P)
                        'If P(I) = 6 Then Stop
                        If Err <> 0 Then
                            'False.
                            Err.Clear
                        Else
                            'assign the current value to the variable.
                            'CountIfParser.Variables.Add P(1), P(I)
                            varmanip.Value = P(I)
                            CountIfParser.ExecuteByRef ReturnVal
                            ReturnVal = CBool(ReturnVal)
                            If Err.Number <> 0 Then
                                'no increment.
                            ElseIf ReturnVal Then
                                'true- increment.
                                
                                IfCount = IfCount + 1
                            End If
                        
                        End If
                    Next I
                    
                    
                    
                    Ret = CVar(IfCount)
                        
                    
                    
                Case "PICKMENU"
                    'Pickmenu.
                    'not handled.
                    'well, not yet. maybe never.
                    'isn't this friendly, we return true for Handled in the CanHandleFunction Interface routine,
                    'but then, when the Parser calls us, we just bitch and return an error. :)
                    Ret = CVErr(GPRV_NOTHANDLED)
                
                
                    
                Case "YEAR"
                
                    If paramcount = 0 Then Ret = Year(Now) Else Ret = Year(P(0))
                Case "MONTH"
                    If paramcount = 0 Then Ret = Month(Now) Else Ret = Month(P(0))
                Case "DAY"
                    If paramcount = 0 Then Ret = Day(Now) Else Ret = Day(P(0))
                Case "HOUR"
                    If paramcount = 0 Then Ret = Hour(Now) Else Ret = Hour(P(0))
                Case "MINUTE"
                    If paramcount = 0 Then Ret = Minute(Now) Else Ret = Minute(P(0))
                Case "SECOND"
                    If paramcount = 0 Then Ret = Second(Now) Else Ret = Second(P(0))
                Case "CREATEOBJECT"
                    If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 1 Then
                        Set Ret = VBA.createObject(P(0))
                    Else
                        Set Ret = VBA.createObject(P(0), P(1))
                    
                    End If
                Case "CVERR"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = CVErr(P(0))
                Case "DATEADD"
                If paramcount <= 3 Then GoTo NotOptional
                
                    Ret = DateAdd(P(0), P(1), P(2))
                Case "DATEDIFF"
                If paramcount < 3 Then GoTo NotOptional
                    If paramcount = 3 Then
                        Ret = DateDiff(P(0), P(1), P(2))
                 
                    ElseIf paramcount >= 4 Then
                        Ret = DateDiff(P(0), P(1), P(2), P(3))
                    
                    
                    
                    End If
                    
                    
                Case "DATEPART"
                    If paramcount < 2 Then GoTo NotOptional
                    If paramcount = 2 Then
                        Ret = DatePart(P(0), P(1))
                    ElseIf paramcount >= 3 Then
                        Ret = DatePart(P(0), P(1), P(3))
                    End If
        
                Case "DATESERIAL"
                    If paramcount < 3 Then GoTo NotOptional
                    Ret = DateSerial(P(0), P(1), P(2))
                Case "DATEVALUE"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = DateValue(CStr(P(0)))
                    
                Case "DDB"
                    'DDB(cost,salvage,life,period,factor)
                    If paramcount < 4 Then GoTo NotOptional
                    If paramcount = 4 Then
                        Ret = DDB(P(0), P(1), P(2), P(3))
                    ElseIf paramcount >= 5 Then
                        Ret = DDB(P(0), P(1), P(2), P(3), P(4))
                    
                    Else
                    
                    End If
                    
                Case "DIR", "DIR$"
                    'The dir$ function acts like a Variable and a function, which is why you find one here and in beforeresolveVariable.
                    If paramcount = 0 Then
                        Ret = Dir$
                    ElseIf paramcount = 1 Then
                        Ret = Dir$(P(0))
                    ElseIf paramcount >= 2 Then
                        Ret = Dir$(P(0), P(1))
                    End If
                    
                        
                    
                    
                    
                Case "DOEVENTS"
                    'ignore params.
                    Ret = DoEvents
                Case "ENVIRON"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Environ$(P(0))
                Case "ERROR"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Error$(P(0))
                Case "EXP"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Exp(P(0))
                Case "FILEDATETIME"
                    Ret = FileDateTime(P(0))
                Case "FILELEN"
                    Ret = FileLen(P(0))
                Case "FILTER"
                
                'Improved syntax
                'Filter(array,variable,expression)
                'very similar to foreach; except for each array element <variable> is set and if <expression> returns true, that item will be in the result set.
                'Variable and Expression are unparsed.
                If paramcount < 3 Then GoTo NotOptional
                'Call Assign(Ret, FilterEx(Withparser, P(0), P(1), P(2)))
                Call Assign(Ret, FilterArray(withparser, P(0), P(1), P(2)))
                
                
'                    If paramcount < 2 Then GoTo NotOptional
'                    If paramcount = 2 Then
'                    Ret = Filter(P(0), P(1))
'                    ElseIf paramcount = 3 Then
'                        Ret = Filter(P(0), P(1), P(2))
'                    ElseIf paramcount = 4 Then
'                        Ret = Filter(P(0), P(1), P(2), P(3))
'
'                    End If
                Case "FORMAT", "FORMAT$"
                    If paramcount = 0 Then GoTo NotOptional
                    Select Case paramcount
                        Case 1
                            'hey, who are we to judge? they want
                            'to format a string with no format, be my guest.
                            'in fact, we may as well have just skipped the call to format.
                            'they could have relied on internal conversion. But then they might be curious why varType was still giving a
                            'number when they use it as the only argument to format.
                            Ret = Format(P(0))
                        Case 2
                            Ret = Format(P(0), P(1))
                        Case 3
                            Ret = Format(P(0), P(1), P(2))
                        
                        Case Is <= 4
                            Ret = Format(P(0), P(1), P(2), P(3))
                    End Select
                    
                    
                    
                    
                    
                Case "FORMATCURRENCY"
                'there are 7 possible params.
                    If paramcount = 0 Then GoTo NotOptional
                    Select Case paramcount
                        Case 1
                            Ret = FormatCurrency(P(0))
                        Case 2
                            Ret = FormatCurrency(P(0), P(1))
                        Case 3
                            Ret = FormatCurrency(P(0), P(1), P(2))
                        Case 4
                            Ret = FormatCurrency(P(0), P(1), P(2), P(3))
                        Case Is >= 5
                            Ret = FormatCurrency(P(0), P(1), P(2), P(3), P(4))
                    End Select
                Case "FORMATDATETIME"
                    If paramcount = 0 Then GoTo NotOptional
                        If paramcount = 1 Then
                            Ret = FormatDateTime(P(0))
                        ElseIf paramcount > 1 Then
                            Ret = FormatDateTime(P(0), P(1))
                        End If
                Case "FORMATNUMBER"
                    '5 parameters.
                    If paramcount = 0 Then GoTo NotOptional
                    Select Case paramcount
                        Case 1
                            Ret = FormatNumber(P(0))
                        Case 2
                            Ret = FormatNumber(P(0), P(1))
                        Case 3
                            Ret = FormatNumber(P(0), P(1), P(3))
                        Case 4
                            Ret = FormatNumber(P(0), P(1), P(3), P(4))
                        Case Is >= 5
                            Ret = FormatNumber(P(0), P(1), P(3), P(4), P(5))
                    End Select
                Case "FORMATPERCENT"
                    If paramcount = 0 Then GoTo NotOptional
                    Select Case paramcount
                        Case 1
                            Ret = FormatPercent(P(0))
                        Case 2
                            Ret = FormatPercent(P(0), P(1))
                        Case 3
                            Ret = FormatPercent(P(0), P(1), P(3))
                        Case 4
                            Ret = FormatPercent(P(0), P(1), P(3), P(4))
                        Case Is >= 5
                            Ret = FormatPercent(P(0), P(1), P(3), P(4), P(5))
                      
                    End Select
                Case "FV"
                If paramcount < 5 Then GoTo NotOptional
                    '5 parameters. last two optional.
                    If paramcount = 3 Then
                        Ret = FV(P(0), P(1), P(2))
                    ElseIf paramcount = 4 Then
                        Ret = FV(P(0), P(1), P(2), P(3))
                    ElseIf paramcount = 5 Then
                        Ret = FV(P(0), P(1), P(2), P(3))
                    End If
                Case "GETATTR"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = GetAttr(P(0))
                Case "GETOBJECT"
                    If paramcount = 0 Then GoTo NotOptional
                    If IsEmpty(P(0)) Then
                        Set Ret = GetObject(, P(1))
                    ElseIf paramcount = 1 Then
                        Set Ret = GetObject(P(1))
                    Else
                        Set Ret = GetObject(P(0), P(1))
                    End If
                Case "GETSETTING"
                    'getsetting(
                    
                    If paramcount < 3 Then GoTo NotOptional
                    If paramcount = 3 Then
                        Ret = GetSetting(P(0), P(1), P(2))
                    ElseIf paramcount = 3 Then
                        Ret = GetSetting(P(0), P(1), P(2), P(3))
                    End If
                Case "HEX", "HEX$"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Hex$(P(0))
                Case "IIF", "IF"
                    
                    If paramcount < 2 Then GoTo NotOptional
                    Dim PtempParse As CParser
                    Set PtempParse = New CParser
                    Set PtempParse = withparser.Clone
                    'Ret = IIf(P(0), P(1), P(2))
                    If CBool(P(0)) Then
                        PtempParse.Expression = P(1)
                    ElseIf paramcount > 2 Then
                        PtempParse.Expression = P(2)
                    
                    End If
                    If PtempParse.Expression <> "" Then
                        PtempParse.ExecuteByRef Ret
                    End If
                    
                    
                    
                    
                    
                    
                Case "INPUTBOX", "INPUTBOX$"
                    '7 params. all but first optional.
                    If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 1 Then
                        Ret = InputBox(P(0))
                    ElseIf paramcount = 2 Then
                        Ret = InputBox(P(0), P(1))
                    ElseIf paramcount = 3 Then
                        Ret = InputBox(P(0), P(1), P(2))
                    ElseIf paramcount = 4 Then
                        Ret = InputBox(P(0), P(1), P(2), P(3))
                    ElseIf paramcount = 5 Then
                        Ret = InputBox(P(0), P(1), P(2), P(3), P(4))
                    ElseIf paramcount = 6 Then
                        Ret = InputBox(P(0), P(1), P(2), P(3), P(4), P(5))
                    ElseIf paramcount = 7 Then
                        Ret = InputBox(P(0), P(1), P(2), P(3), P(4), P(5), P(6))
                    End If
                Case "INSTR"
            'first arg optional. The rest are not.
            'if p(0) is empty, change to 1. Just out of kindness.
            
                If IsEmpty(P(0)) Then P(0) = 1
                If VarType(P(0)) = vbString Then
                    If paramcount = 2 Then
                        Ret = InStr(1, P(0), P(1))
                    ElseIf paramcount = 3 Then
                        Ret = InStr(1, P(0), P(1), P(2))
                    End If
                ElseIf IsNumeric(P(0)) Then
                    Ret = InStr(P(0), P(1), P(2))
                End If
                
            
                Case "INSTRREV"
                    '
                    '2,3,4,5
                    If paramcount < 2 Then GoTo NotOptional
                    If paramcount = 2 Then
                        Ret = InStrRev(P(0), P(1))
                    ElseIf paramcount = 3 Then
                        Ret = InStrRev(P(0), P(1), P(2))
                    ElseIf paramcount = 4 Then
                        Ret = InStrRev(P(0), P(1), P(2), P(3))
                        
                    
                    End If
                    
                Case "INT", "FLOOR"     'int and the C Floor function.
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Int(P(0))
                Case "CEIL"
                    If paramcount = 0 Then GoTo NotOptional
                    'ceil. returns smallest integer greater then or equal to the parameter.
                    Ret = Ceil(P(0))    'Custom func.
                Case "FIX"
                    Ret = Fix(P(0))
                Case "IPMT"
                    If paramcount = 4 Then
                        Ret = IPmt(P(0), P(1), P(2), P(3))
                    ElseIf paramcount = 4 Then
                        Ret = IPmt(P(0), P(1), P(2), P(3), P(4))
                    ElseIf paramcount >= 5 Then
                        Ret = IPmt(P(0), P(1), P(2), P(3), P(4))
                    End If
                Case "IRR"
                'we need to create a double array from the obvious variant array
                'that should be the first parameter.
                Dim dblMake() As Double
                ReDim dblMake(UBound(P(0)))
                For I = 0 To UBound(dblMake)
                    dblMake(I) = CDbl(P(0)(I))
                Next I
                    If paramcount = 1 Then
                        Ret = IRR(dblMake)
                    ElseIf paramcount >= 2 Then
                        Ret = IRR(dblMake, P(1))
                    End If
                   ' If paramcount = 1 Then
                        
                Case "ISARRAY"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = IsArray(P(0))
                Case "ISDATE"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = IsDate(P(0))
                Case "ISEMPTY"
                If paramcount = 0 Then GoTo NotOptional
                    Ret = IsEmpty(P(0))
                Case "ISERROR"
                If paramcount = 0 Then GoTo NotOptional
                    Ret = IsError(P(0))
                Case "ISNULL"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = IsNull(P(0))
                Case "ISNUMERIC"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = IsNumeric(P(0))
                Case "ISOBJECT"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = IsObject(P(0))
                Case "JOIN"
                
                    'ahh sheesh.
                    If paramcount = 0 Then GoTo NotOptional
                    Dim JoinEm As Variant
                    ReDim JoinEm(1 To UBound(P) - 1)
                    
                    
                    
                    For I = 1 To UBound(P)
                    
                        Assign JoinEm(I + 1), P(I)
                    Next
                Case "LBOUND"
                    'uggh. And, with my new support for matrices(well, SORT of), I get the fun task of handling the optional argument.
                    If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 1 Then
                        Ret = LBound(P(0))
                    ElseIf paramcount >= 2 Then
                        Ret = LBound(P(0), P(1))
                    
                    End If
                Case "LCASE", "LCASE$"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = LCase(P(0))
                Case "LEFT", "LEFT$"
                    If paramcount <= 1 Then GoTo NotOptional
                    Ret = left(P(0), P(1))
                Case "LEN"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Len(P(0))
                Case "LOG"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Log(P(0))
                Case "LTRIM", "LTRIM$"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = LTrim(P(0))
                Case "RTRIM", "RTRIM$"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = RTrim(P(0))
                Case "MID", "MID$"
                      If paramcount <= 1 Then GoTo NotOptional
                      If paramcount = 2 Then
                        Ret = Mid(P(0), P(1))
                      ElseIf paramcount > 2 Then
                        Ret = Mid$(P(0), P(1), P(2))
                      End If
                Case "MIRR"
                    'ret = MIRR(P(0), P(1), P(2))
                Case "MONTHNAME"
                    If UBound(FuncParams) = 0 Then
                        Ret = MonthName(FuncParams(0))
                    Else
                        Ret = MonthName(FuncParams(0), FuncParams(1))
                    End If
                Case "MSGBOX"
                    'msgbox(1,[2],[3],[4],[5])
                    If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 1 Then
                        Ret = MsgBox(P(0))
                    ElseIf paramcount = 2 Then
                        Ret = MsgBox(P(0), P(1))
                    ElseIf paramcount = 3 Then
                        Ret = MsgBox(P(0), P(1), P(2))
                    ElseIf paramcount = 4 Then
                        Ret = MsgBox(P(0), P(1), P(2), P(3))
                    ElseIf paramcount >= 5 Then
                        Ret = MsgBox(P(0), P(1), P(2), P(3), P(4))
                    End If
                Case "NPER"
                
                
                
                'nper rate,pmt,pv,[fv],[due]
                '4 required, then 2 optional.
                Select Case paramcount
                Case Is <= 3
                    GoTo NotOptional
                Case 4
                    Ret = NPer(P(0), P(1), P(2))
                Case 5
                    Ret = NPer(P(0), P(1), P(2), P(3))
                Case Is >= 6
                    Ret = NPer(P(0), P(1), P(2), P(3), P(4))
                End Select
                Case "NPV"
                'npv(rate,valueArray())
                'convert P(1) to a DOUBLE array.
                
                Erase dblMake
                P(1) = Flatten(P(1))
                ReDim dblMake(LBound(P(1)) To UBound(P(1)))
                For I = LBound(dblMake) To UBound(dblMake)
                    On Error Resume Next
                    dblMake(I) = P(1)(I)
                    If Err <> 0 Then
                        dblMake(I) = 0
                    End If
                Next I
                Ret = NPV(P(0), dblMake())
                Case "OCT", "OCT$"
                    Ret = Oct(P(0))
                Case "PARTITION"
                    Ret = Partition(P(0), P(1), P(2), P(3))
                Case "PMT"
                    '
                    '3,4,5
                    'Pmt Rate, NPer, PV, [FV], [due]
                    If paramcount < 3 Then GoTo NotOptional
                    If paramcount = 3 Then
                        Ret = Pmt(P(0), P(1), P(2))
                    ElseIf paramcount = 4 Then
                        Ret = Pmt(P(0), P(1), P(2), P(3))
                    ElseIf paramcount >= 5 Then
                        Ret = Pmt(P(0), P(1), P(2), P(3), P(4))
                        
                    End If
                    
                Case "PPMT"
                    '
                    '4,5,6
                    'ppmt rate,per,nper,pv,[fv],[due]
                    If paramcount = 4 Then
                        Ret = PPmt(P(0), P(1), P(2), P(3))
                    ElseIf paramcount = 5 Then
                        Ret = PPmt(P(0), P(1), P(2), P(3), P(4))
                    ElseIf paramcount >= 6 Then
                        Ret = PPmt(P(0), P(1), P(2), P(3), P(4), P(5))
                    
                    End If
                Case "PV"
                '
                '3,4,5
                'pv rate,nper,pmt,fv,due
                Case "QBCOLOR"
                    If paramcount = 0 Then GoTo NotOptional
                    
                    Ret = QBColor(P(0))
                Case "RATE"
                '3,4,5,6
                    If paramcount < 3 Then GoTo NotOptional
                    Select Case paramcount
                        Case 3
                            Ret = Rate(P(0), P(1), P(2))
                        Case 4
                            Ret = Rate(P(0), P(1), P(2), P(3))
                        Case 5
                            Ret = Rate(P(0), P(1), P(2), P(3), P(4))
                        Case Is >= 6
                            Ret = Rate(P(0), P(1), P(2), P(3), P(4), P(5))
                    End Select
                    
                    
                Case "REPLACE", "REPLACE$"
                    If paramcount < 3 Then GoTo NotOptional
                    If paramcount = 3 Then
                        Ret = Replace(P(0), P(1), P(2))
                    ElseIf paramcount = 4 Then
                        Ret = Replace(P(0), P(1), P(2), P(3))
                    ElseIf paramcount = 5 Then
                        Ret = Replace(P(0), P(1), P(2), P(3), P(4))
                    ElseIf paramcount >= 6 Then
                        Ret = Replace(P(0), P(1), P(2), P(3), P(4), P(5))
                    End If
                Case "RGB"
                    Dim retcolour As Colour
                    'TODO: make this point to a new instance of my ClsColor class instead of a Long.
                    Set retcolour = New Colour
                    If paramcount < 3 Then GoTo NotOptional
                    retcolour.Red = P(0)
                    retcolour.Green = P(1)
                    retcolour.Blue = P(2)
                    Set Ret = retcolour
                    'ret = rgb(P(0), P(1), P(2))
                Case "RIGHT", "RIGHT$"
                    If paramcount < 2 Then GoTo NotOptional
                    Ret = right(P(0), P(1))
                Case "ROUND"
                    If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 1 Then
                        Ret = Round(P(0))
                    ElseIf paramcount >= 2 Then
                        Ret = Round(P(0), P(1))
                    End If
                Case "SGN"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Sgn(P(0))
                Case "SHELL"
                    If paramcount = 0 Then GoTo NotOptional
                    If paramcount = 1 Then
                        Ret = Shell(P(0))
                    ElseIf paramcount = 2 Then
                        Ret = Shell(P(1))
                    End If
                Case "SLN"
                    'sln(cost,salvage,life)
                    If paramcount < 3 Then GoTo NotOptional
                    
                        Ret = SLN(P(0), P(1), P(2))
                Case "SPACE", "SPACE$"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Space(P(0))
                Case "SPLIT"    'Works only on strings, oh, and arrays, too.
                'when the first argument is not an array:
                'split(Value,[delimiter],[limit],[compare])
                'when the first item is an array:
                
                
                
                'split(Array,FirstItem,[Chunksize],[numchunks])
                
                
                
                If paramcount < 1 Then GoTo NotOptional
                    'darn split- we need to traverse the returned array and convert it into a variant
                    'array.
                    Dim SplitStrings() As String
                    'split(Expression,delimiter,limit,compare)
                    If Not IsArray(P(0)) Then
                        Select Case paramcount
                            Case 1  'expression
                                SplitStrings = Split(P(0))
                            Case 2  'expression,delimiter
                                SplitStrings = Split(P(0), P(1))
                            Case 3  'expression,delimiter,limit
                                SplitStrings = Split(P(0), P(1), P(2))
                            Case 4 'expression,delimiter,limit,compare
                                SplitStrings = Split(P(0), P(1), P(2), P(3))
                        End Select
                            'now, we convert splitstrings into a variant array.
                            ReDim Ret(0 To UBound(SplitStrings))
                            For I = 0 To UBound(SplitStrings)
                                Ret(I) = SplitStrings(I)
                            Next I
                            
                            'there, that should do it...
                    Else
                        'invalid procedure call or argument.
                        ParserError(Me, 5, "Split can't split an array!").Throw
                        
                        
                    
                    End If
                    
                    
                Case "SQR", "SQRT"
                    'NOTE! Square root of positive numbers returns an array of the positive and negative square roots.
                    'HOWEVER! we only return the array if the second parameter was specified and is true.
                    'the second parameter determines wether to return both.
                    'Oh, and since Complex number support is included, it also supports taking the square root
                    'of negative numbers.
                If (P(0) < 0) Then 'less than zero. Complex number.
                    Set Ret = Complex(0, Sqr(Abs(P(0))))
                Else
                    If paramcount > 1 Then
                       If CBool(P(1)) = True Then
                        Ret = Array(Sqr(Abs(P(0))), -Sqr(Abs(P(0))))
                       
                       Else
                       
                            Ret = Sqr(P(0))
                       End If
                    Else
                        Ret = Sqr(P(0))
                    End If
                End If
                Case "STR", "STR$"
                    If VarType(P(0)) <> vbString Then
                    Ret = str(P(0))
                    Else
                     Ret = P(0)
                    End If
                Case "STRCOMP"
                'SOAB! I thought I finished this part....
                    'strcomp(string1,string2,comparemode)
                    If paramcount < 2 Then GoTo NotOptional
                    If paramcount = 2 Then
                        Ret = StrComp(P(0), P(1))
                    
                    
                    ElseIf paramcount >= 3 Then
                        Ret = StrComp(P(0), P(1), P(2))
                    
                    
                    End If
                Case "STRCONV"
                    'strconv(string,conversion,localeid)
                    If paramcount < 2 Then GoTo NotOptional
                    If paramcount = 2 Then
                        Ret = StrConv(P(0), P(1))
                    
                    
                    ElseIf paramcount >= 3 Then
                        Ret = StrConv(P(0), P(1), P(2))
                    
                    
                    End If
                Case "STRING"
                    'string(number,character)
                    If paramcount < 2 Then GoTo NotOptional
                    Ret = String(P(0), P(1))
                Case "STRREVERSE"
                    'strreverse(expression)
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = StrReverse(P(0))
                Case "SWITCH"
                    If paramcount < 3 Then GoTo NotOptional
                    'switch(varexpr())
                    
                    For I = 1 To paramcount Step 2
                        If P(I) = P(0) Then
                            If I = UBound(P) Then
                                Err.Raise 5, "BPCoreOpFunc::HandleFunction", "True Expression contained no following Value to return."
                            Else
                                Ret = P(I + 1)
                            End If
                        End If
                        
                    Next I
                    
                    
                Case "SYD"
                    'SYD(cost,salvage,life,period)
                    If paramcount < 4 Then
                        Ret = SYD(P(0), P(1), P(2), P(3))
                    End If
                Case "TIMESERIAL"
                    'timeserial hour,minute,second
                    If paramcount < 3 Then GoTo NotOptional
                        Ret = TimeSerial(P(0), P(1), P(2))
                Case "TIMEVALUE"
                    'timevalue time
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = TimeValue(P(0))
                Case "TYPENAME"
                    'typename varname
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = TypeName(P(0))
                Case "UBOUND"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = UBound(P(0))
                Case "UCASE", "UCASE$"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = UCase$(P(0))
                Case "VAL"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Val(P(0))
                Case "VARTYPE"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = VarType(P(0))
                Case "WEEKDAY"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Weekday(P(0))
                Case "WEEKDAYNAME"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = WeekdayName(P(0))
                Case "RANDOMARRAY"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = RandomArray(P(0), P(1), P(2))
                Case "RANDOM"
                    If paramcount = 0 Then GoTo NotOptional
                    Ret = Random(P(0), P(1))
                
        '"MID MID$ MIRR MONTHNAME MSGBOX " & _
        "NPER NPV OCT OCT$ PARTITION PMT PPMT PV " & _
        "QBCOLOR " & _
        "RATE REPLACE REPLACE$ RGB RIGHT RIGHT$ ROUND " & _
        "SGN SHELL SIN SLN SPACE SPLIT SQR STR STRCOMP STRCONV STRING STRREVERSE SWITCH SYD " & _
        "TAN TIMESERIAL TIMEVALUE TYPENAME UBOUND UCASE UCASE$ VAL VARTYPE WEEKDAY WEEKDAYNAME "
                Case "FOREACH"
                'FOREACH({1,2,3},varname,Expression)
                
                Dim arrayuse As Variant
                Dim VarName As String
                Dim Expression As String
                'Debug.Assert False
                If paramcount < 3 Then GoTo NotOptional
                arrayuse = P(0)
                VarName = P(1)
                Expression = P(2)
                ForEach withparser, arrayuse, VarName, Expression
                
                
                
                
                
                Case "STORE"
                
                    'STORE: parameters:
                    'STORE(VarName,VarValue,Scope)
                    'STORE(Varname,VarValue,IMMUTABLE|CONSTANT|FINAL)
                    
                
                    'this may look wrong, but funcparams(0) is NOT a parsed value,
                    'and thus should be a variable name.
                    Set varGrab = withparser.Variables.Item(FuncParams(0))
                   If IsObject(FuncParams(1)) Then
                    Set varGrab.Value = FuncParams(1)        'if it is constant already, this will
                   Else
                    varGrab.Value = FuncParams(1)
                   End If
                   'raise an error.
                   'the functions return value is the
                   'assigned value.
                   'check if it is supposed to be constant.
                   If paramcount > 2 Then
                    Select Case UCase$(FuncParams(2))
                        Case "IMMUTABLE", "CONSTANT", "FINAL", "CONST"
                            varGrab.Constant = True
                        
                    End Select
                   
                   End If
                   If IsObject(varGrab.Value) Then
                   Set Ret = varGrab.Value
                    Else
                        Ret = varGrab.Value
                    End If
              
                
                Case "MINUS"
                    Ret = -P(0)
                'MORE FUNCTIONS :) EvalArr, Comb and so forth.
                'Of course I should probably put the following into place with the current functions:
                'Make sure that I use Case Statements for any of the Paramcount checks that are larger then two entries.
                'with is ANY with optional arguments (the case for not optional, the case for those non-optional arguments, and then the check for that one additional argument)
                
                
                
                
            End Select
            'copy back.
    End If
    If IsObject(Ret) Then
        Set ReturnVal = Ret
    Else
        ReturnVal = Ret
    End If
    
    Exit Sub
FunctionError:
    'Function error.
    'TODO:// add some sort of IDebugger interface, that can be called from here.
    'I imagine it'll be a Property of the CParser, like the IEvalEventSinks.
    'For now, raise back to the CParser object. The
    'CollapseStack will catch the error and suffix on Line and other such information, raising an event
    'to the client before attempting to use Err.Raise.
    
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
    
    
   On Error GoTo 0
   Exit Sub
   
   'NotOptional-
   'we Goto (SHUT UP!) here
   'when we find that paramcount is less
   'then the minimum number of arguments for
   'the function.
   
   
NotOptional:
'Interesting note: guess what happens when we raise the error! it'll go UP to FunctionError! Cool, heh?
    Err.Raise 45, "BPCoreOpFunc::HandleFunction", "Argument not Optional in function " & funcName & "()"
    
    
    Exit Sub
End Sub
Private Function FilterEx(withparser As CParser, ArrayFilter, FilterVariable, FilterExpression) As Variant
'Returns a filtered array. Each item that satisfies "FilterExpression" is returned.

    Dim OurParser As CParser
    Dim RetArray(), retcount As Long
    Dim currindex As Long
    retcount = -1
    ArrayFilter = Flatten(ArrayFilter)
    Set OurParser = withparser.Clone
    OurParser.Variables.Add FilterVariable, 0
    OurParser.Expression = FilterExpression
    
    For currindex = LBound(ArrayFilter) To UBound(ArrayFilter)
        OurParser.Variables.Add FilterVariable, ArrayFilter(currindex)
        If CBool(OurParser.Execute) Then
            retcount = retcount + 1
            ReDim Preserve RetArray(retcount)
            Call Assign(RetArray(retcount), ArrayFilter(currindex))
        
        End If


    Next
    
    FilterEx = RetArray









End Function
Private Property Get IEvalEvents_Name() As String
IEvalEvents_Name = "BPCoreOpFunc"
End Property

Private Sub IEvalEvents_PreExecute(withparser As CParser, FirstItem As CFormItem)
'
End Sub

Public Sub IEvalEvents_PreParse(Parser As CParser, ByRef Expression As String)
    'fired immediately when Parsing begins.
    'to define special extensions, change those extensions
    'into their equivalent Visual Basic code.
    'COMPLEX NUMBER support:
    'since this event (SHOULD) be fired by child parsers when they are given an expression,
'    'we can check for valid Complex number syntax.
'    CDebug.Post "BPCoreOpFunc Class was asked to Preparse the Expression """ & Expression & """."
'    Dim retcomplex As CComplex
'    Dim isComplex As Boolean
'    'ADDENDUM:
'    'I want to support more then just constants within these Complex number expressions.
'    'I want to be able to do stuff like:
'    '"(5^2 - Sqr(2)i)/(7-Ai)"
'    'No, this will NOT be easy.
'    'for one thing, it introduces a new ambiguity- what if a variables name ends with i?
'    'and it it the last token in a sub-expression? In this case:
'    '"(5+2/PHI)+2"
'    'the parser will think (or more precisely, this "pre-parser" will think
'    'that the i used is to indicate that the expression is a complex one. In fact, this is not the case.
'    'the case.

    'the above comment is redundant- I simply added a new variable i with a complex value.

'    isComplex = ExpressionisComplex(Expression, Parser)
'
'        If isComplex Then
'            'it is a complex number- at least as far as we can tell.
'           ' Set retcomplex = New CComplex
'           ' retcomplex.StringForm = Expression
'           ' Expression = "COMPLEX(" & retcomplex.Realpart & "," & retcomplex.ImagPart & ")"
'           ' CDebug.Post "BPCoreOpFunc replaced a Complex number Expression to a Complex Number Class Instance."
'            'tada.
'            Expression = GetComplexReplace(Expression)
'            CDebug.Post "replaced a Complex Expression with " & Expression
'        End If
    
   
End Sub
'-commented out due to redundancy.
'Private Function ExpressionisComplex(ByVal Expression As String, ByVal withparser As CParser) As Boolean
'    'Returns: wether the given expression is a Complex number expression.
'    'the desired output is illustrated:
'    'for the expression:
'    '"4*2/3+i"
'    'returns false.
'    'whereas:
'    '"Sqr(4+3)-17i"
'    'should return True. However:
'    '"3^A-5*Phi" should NOT return true. This is discovered when
'    'a variable named "phi" is found.
'    Dim VarsI As Collection
'    Dim LoopVar As CVariable
'    'Step one: we know a expression isn't complex if it doesn't end with i.
'    If right$(Expression, 1) <> "i" Then
'        ExpressionisComplex = False
'    Else
'        'OK.
'        Set VarsI = New Collection
'        For Each LoopVar In withparser.Variables
'            If right$(LoopVar.Name, 1) = "i" Then
'                VarsI.Add LoopVar
'            End If
'        Next
'        'if the count is zero, we KNOW we have a expression that can be called Complex.
'        If VarsI.Count = 0 Then
'            ExpressionisComplex = True
'        Else
'            'more checking.
'            For Each LoopVar In VarsI
'                If StrComp(LoopVar.Name, right$(Expression, Len(LoopVar.Name)), vbTextCompare) = 0 Then
'                    'equal.
'                    'these are ignored.
'                    ExpressionisComplex = False
'                    Exit Function
'                End If
'            Next
'            ExpressionisComplex = True
'
'        End If
'
'
'
'
'
'
'    End If
'End Function

Public Sub IEvalEvents_ParseComplete(ParserDone As CParser)
    'fired when a CParser object completes the parsing portion.
End Sub
Public Sub IEvalEvents_ExecuteComplete(ParserDone As CParser)
    'fired when a CParser object completes execution.
End Sub

Public Function IEvalEvents_Init(withparser As CParser) As IPluginInitCode
'
CDebug.Post "Core Function/Operator facility Created"
AddOurVars withparser
mvarDisabledFunctions = withparser.Settings.GetPluginSetting(Me, "Disabled Functions", withparser.configset, "")
mvarDisabledOperators = withparser.Settings.GetPluginSetting(Me, "Disabled Operators", withparser.configset, "")
If Not mvarDisabledFunctions = "" Then
    CDebug.Post "Disabled Functions found in registry-" & mvarDisabledFunctions, Severity_Information

End If
If Not mvarDisabledOperators = "" Then
    CDebug.Post "Disabled Operators found in registry-" & mvarDisabledOperators, Severity_Information
End If

'Now if this was any other plugin (especially a external one) this would definitely be questionable behaviour.)
CDebug.Post "forcing use of BPCoreOpFunc Class as ICorePlugin as well..."
If withparser.CorePlugins.Parser Is Nothing Then Set withparser.CorePlugins.Parser = withparser
withparser.CorePlugins.Add Me
CDebug.Post "BPCoreOpFunc Class added to CorePlugins collection "



'lastly, add a BCFile object of BCFile is installed on the current machine.
Dim testCreate As Object
On Error Resume Next
Set testCreate = createObject("BCFile.BCFSObject")
If Not testCreate Is Nothing Then
    withparser.Variables.Add "BCFile", testCreate
End If


End Function


Private Property Get IEvalEvents_Self() As Object
    Set IEvalEvents_Self = Me
End Property




Private Property Let IEvalEvents_Settings(withparser As CParser, ByVal PropName As String, ByVal RHS As Variant)
'
  
'End Property
    Static mreguse As cRegistry, mkey As String
    If mreguse Is Nothing Then Set mreguse = New cRegistry
    mreguse.ClassKey = HHKEY_CURRENT_USER

    mkey = ParserSettings.PluginKey(withparser.configset, "BaseParserXP.BPCoreOpFunc")
    mreguse.ValueEx(HHKEY_CURRENT_USER, mkey, PropName, RREG_SZ, "") = RHS
End Property

Private Property Get IEvalEvents_Settings(withparser As CParser, ByVal PropName As String) As Variant
'
'we have a small table- given a property name, we wish to look up a registry location,key,type, and default...

    Static mreguse As cRegistry, mkey As String
    If mreguse Is Nothing Then Set mreguse = New cRegistry
    mreguse.ClassKey = HHKEY_CURRENT_USER

mkey = ParserSettings.PluginKey(withparser.configset, "BaseParserXP.BPCoreOpFunc")

IEvalEvents_Settings = mreguse.ValueEx(HHKEY_CURRENT_USER, mkey, PropName, RREG_SZ, "")

    
    
End Property

Public Function IEvalEvents_Term(withparser As CParser) As IPluginTermCode
'terminates from a parser.
End Function





Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function




Private Sub IOpFuncInfo_GetFunctionInformation(withparser As CParser, ByVal funcName As String, finfo As FUNCTIONINFORMATION)

   finfo = GetFunctionInfoFromXML(withparser, funcName, FuncInfo)
   finfo.StrFunctionName = "BPCoreOpFunc." & finfo.StrFunctionName

End Sub

Private Property Get IPluginUIData_Description() As String
    IPluginUIData_Description = "BASeParser XP Core Plugin."
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As stdole.Picture
'NOT YET--
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String

'return BPControls.UCCoreSettings
IPluginUIData_GetSettingsPageProgID = "BPXPGUI.UCCoreSettings"


End Function
