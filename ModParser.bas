Attribute VB_Name = "ModParser"
Option Explicit
'Project Dependencies:


'Typelib Information
'Microsoft Script Control 1.0
'Microsoft XML, v3.0
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions
'
'Edanmo 's OLE interfaces & functions v1.81
'Edanmo 's OLE interfaces for implements v1.51

'MAY replace some settings code with calls to BCSettings library, on account of that being written to, you know- be used.


'Windows COmmon controls 6.0 SP6
'MS richtextbox control 6.0 SP6


'BASeParser
'View development log for notes, as well as a sad description of a fairly long chain of bad luck I
'had with my computers, through no fault of them. Although, maybe from me.


'TODO: tuples; like:

'{A,B}=Function()


'THE CODE

'Hopefully, you have all the modules and classes BASeParser needs. I do use quite a lot of classes
'from my other projects, so forgive me if I forget one or two the first time I release the code.



'ModCipher
'HEE HEE! I only use this for a certain easter egg.
'Anyways, the ModCipher module simply contains a few Cryptography functions, they aren't exactly advanced,
'but they seem to do the job well. And since I only include the hashed version to test, there is no evidence
'of the original password I passed to the cipher function. I assume it isn't impossible to reverse-engineer it, since I
'give the code that creates the hash. But- just look at it!


'The idea here is to prevent anybody stealing my code and claiming it as their own.
'Thatr is probably one of my greater fears as far as this project goes.


'ModDebug
'one-liner that I include with CDebug. No comment. Except that. That was one comment.

'ModImprove
'a few improvements on the intrinsic VB functions, such as ASC and CHR.
'for example, CHR can take multiple arguments.

'ModSort
'Implements the sorting and binary search procedures used by the BPCoreOpFunc Core Plugin.

'ModTiming
'Profiling stuff. QueryPerformanceTimer, a Timer function replacement, etc.
'Probably just as well to put it in ModImprove...

'BPCoreOpFunc
    'the Core Parser Plugin.
    


'CComplex
    'implements Complex numbers. I must apologize, as this Class isn't exactly portable to any other project, on account of it's dependency on CParser's PerformOperation.
    
'CDataList

    'planned improvement on current Array manipulation. probably won't be done for a while, since
    'the current method works fine.
    
'CDebug

    'I think this class is pretty cool. It is in several of my projects. When it is instantiated it creates two debugging files(given the conditional constant, "NODEBUGLOG", isn't defined.)
    'one, a Debug log, is written with all the output from the Post method. Another, a speed profiling CSV file, will contain timing data on ALL functions which call the CDebug
    'object's PushP and PopP procedures. It will have such things as the total number of times called, total time, and average time. of all of these,
    'average time is probably the most useful. I find myself deleting megabytes of these log and profiling files
    'after a few debug sessions.

'CFormItem

    'represents a single token within a expression. For example, each basic token (string,number, etc..) is a IT_VALUE. every function is a
    'IT_FUNCTION whose "Value" property (on the formitem) is a array of CParser objects whose expressions are the function arguments.
    'Operators are IT_OPERATOR types. a Object access operator, @, is IT_OBJACCESS, and array access, IT_ARRACCESS, etc.
    
'CFraction
    
    'used to represent Fractions. As I write this, it doesn't support IOperable's PerformOperation interface method, but does support
    'the toString() method.
    
'CFunction

    'represents a Function. These objects are handled by the intrinsic FunctionHandler Class.
    'found within the CFunctions collection Class
    
'CFunctions

    'collection of CFunction objects. Enough said.
    
'Colour

    'represents a Colour. very versatile. I wrote it quite some time ago.
    
'CMatrix
    
    'Self-explanatory.
    




'OK. I just compiled to Native Code, optimizing for Small code, and I got a DLL that was 1.2 Megabytes. Not good.
'I guess I'll be compiling to P-Code until I find out why. probably should remove dead-code.

#Const BASEPARSER = 1
'ANYWAY...

'This project is my Expression evaluation library. My original version was lost through an accidental
'drive repartition. Yes I had a backup. Well, it WAS my backup. The original was so chock full of features,
'I didn't think I'd be able to re-create them all in my new version. Boy, was I wrong. In fact, this version not
'only has more features, but it is a lot easier to add more, as well as virtually painless to create new functions,
'since it supports external plugins. (it is a pain in the neck creating large libraries of functions, however, due to
'all the work that must be done validating the function parameters.)




'a note about Parse Stack Caching:

'After having the Debug-logging feature active for some time, I have reviewed the output from sessions in which I wasn't breaking into the development environment (which would become part of the code timing)
'and realized that, even though it was called the fewest number of times out of the functions I profile, the RipFormula() routine had the highest TOTAL time. This was deeply concerning.
'the problem was not necessarily  badly written code, but rather the large number of things that need to be checked for. In order to minimize the number of times RipFormula needs to be executed, I have devised the
'Parse Cache. The parse Cache is activated for all expressions containing more than 6 Formula items. the time it takes to parse such an expression seems to average out to about a tenth of a second, whereas executing that built stack takes less than a hundredth of that.
'what happens is the Parse Cache will contain a queue of Parsed Stacks. For example, evaluation of the expression:
'"Sin(X)/Sqr(Y)/(2*3)+4" contain 7 formula items, the Sin and Sqr functions and the division operator between them (three)
'as well as the additional division and the parenthesized expression (the parenthesized expression contains three items, but these don't count) and finally the addition and the literal 4.
'once this is parsed into the internal data structure optimized for execution, the parser will realize that it contained more than 6 items(the value 6 is somewhat arbitrary here, since it can be changed)
'and push it onto the optimized expression queue. The interesting thing is, since each CParser will create additional "child" parser objects for such things, and each of these parsers will itself check the global
'parse cache, the code will automatically allow for future cache hits on expressions
'such as "Sqr(Sin(X)/Sqr(Y)/(2*3)+4)"
'unfortunately, the entire expression must appear as it does in the parse cache- that is, as a separate entity. For example, this expression
'would not have ANY cache hits (given that the parse cache contains a stack for the aforementioned example expression:

'Sin(X)/Sqr(Y)/(2*3)+4/2

'Adding such a feature is slightly more difficult than easy. the Parse routine could inspect each cached item to see if
'it's expression is a substring of the current expression. if so, it could then replace said subexpression with a bracketed form.
'The problem, of course, comes about when we have an optimized version of something like "3-4+3", and try to execute:
'5*3-4+3+2  (which equals 16)
'the parser will think, hey, that there 3-4+3 has already been parsed! lets do this:
'5*(3-4+3)+2 which equals 12)
'it was due to this problem and trying to iron out such inconsistencies that forced me to
'abandon any such scheme. The checks to ensure parenthesizing the portion of the expression in question
'could highly likely surpass any prescribed gain. Of course, it could always appear in a later version...

'as Parenthesized expressions and function arguments (not to mention list elements)





'The core operators and functions are documented within the BPCoreOpFunc class. Technically, the CParser object is clueless about any operator-
'it simply asks the plugin if a given location within the string contains the operator. Take note that unlike my old version, a operator called ++ could exist.
'in fact, one called ** is defined as the same as ^)

'(the following is done)
'TODO:\\ I think a very useful feature would be to extend unary support to allow for them AFTER
'the thing they work with. The only difficulty is I would once again
'need to change the IEvalEvents interface, by changing the isUnary argument from a boolean to
'a custom Enumeration (UnaryTypeConstants) It should be much more painless.


'The syntax is fully comprehended by the parser however.

'VARIABLES:

'variables are fully handled by the execution code. a variable is defined as a non-numeric value that has been parsed.
'for example, the following are variables:

'A, VARIABLE,VARCODE

'if the supposed variable name is immediately preceded by a bracket, however, it is then parsed out as a Function with that name.
'if a variable is present at any level in an expression, it prevents that expression from being fully optimized.
'LISTS:
'a special token is the list, defined by a group of arguments delimited by "{" and "}". These are replaced by a Function call
'to a function called "ARRAY". obviously, the core  operator/Function plugin must be loaded.

'UNARY OPERATOR:

'unary operators are much better handled in this version of BASeParser. This version uses the same logic to parse out unary operators,
'but it treats that operator as a unary operator if it either followed another operator or is the first item in the expression.


'OPTIMIZATION:

'again, a new feature in this BASeParser, is the intrinsic ability for it to optimize any given expression.
'optimization can occur at any level of an expression. For example, the following expression:

'"Sin(Cos(Tan(5+2)))+(5/3*2/5)+Sin(Cos(T))"
'will optimize the Sin() call, the numeric expression, but not the last Sin call. The last Sin call cannot be
'optimized because it contains a Cos Function call that is not constant, which is because of the T.
'(Variables, by definition, are never constant).

'IMPORTANT NOTE:
'before you go criticizing my use of Variants so very often
'in this library, remember that in order to handle all the different
'types of data the parser will encounter, it needs to store it. (duh)
'Variants actually aren't as slow as I thought, but unlike a lot
'of variant advocates say, it causes a heck of a lot of bugs.

'Also of importance is that this library goes relatively slow when
'Debugging is on. This is primarily due to the fact that the library makes excessive calls
'to CDebug.Post, and that method writes to a log file. Often, after a short 5 or 10 minute test, the log can bloat to
'several megabytes. The Debugging Constant is "NODEBUG". If this is defined, then every procedure in
'the CDebug Class becomes a stub.

'Overview of the engine.

'OK, first you send it the Expression.
'Eventually (after a few gyrations with ensuring there are spaces around operators)
'we start ripping the formula-
'now, ripping the formula does involve some communication with our
'set of plugins (which are installed(added to the collection) when the CParser object's
'"Create()" method is called), in the logic to determine what tokens are operators and so forth,
'not to mention interfacing with Core plugins to find custom itemtypes.
'As an example, what happens with this expression:

'5+Sin(A)

'Since Sin will be found to be a function in the Core Operator/Function Plugin, we will end up
'with this InFix:

'IT_VALUE=5,IT_OPERATOR=+,IT_FUNCTION=SIN
'note the absence of any reference to the actual arguments of the SIN function, this is because the
'parser creates ANOTHER parser object for each function argument.

'Anyway, now, BuildStack_Infix builds the RPN stack from the essentially unintelligible Infix stack, and we end of with:

'IT_OPERATOR=+;IT_VALUE=5;IT_FUNCTION=SIN


'and thus the parse is completed. This stack will be stored in the global collection of Stacks, and cached
'so any other parses will happen much faster.


'And then the fun begins with the execution....








'Once ripformula is completed, BuildStack_Infix changes the build stack from Infix notation
'to the "native" Reverse Polish Notation (RPN) that we can execute.
'(little todo for myself, possibly add other parse abilities, such as Postfix and such.)



Public Declare Function GetAsyncKeyState Lib "USER32.DLL" (ByVal vKey As Long) As Integer

Public Const E_NOTIMPL As Long = &H80004001
Public MemState As New CMemState
Public CParserCount As Long
Public Const DefaultBracketStart = " ({["
Public Const DefaultBracketEnd = " ]})"
'Hey, who put this COM crap in here- :)

Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As IShellFolderEx_TLB.Guid)

Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As IShellFolderEx_TLB.Guid) As Long

      Private Declare Function progIDfromCLSID Lib "ole32.dll" Alias "ProgIDFromCLSID" (pCLSID As IShellFolderEx_TLB.Guid, lpszProgID As Long) As Long

      Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As IShellFolderEx_TLB.Guid, lpszProgID As Long) As Long
'I thought this was hidden in the library like VarPtr. Nope.
'gotta alias then...
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

Private Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public ParserSettings As New ParserSettings
Private Declare Function wsprintf Lib "USER32.DLL" (ByVal lpStr As String, ByVal lpcstr As String, ByRef OptionalArguments As Any) As Long
Public Declare Sub CopyMemoryLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal nBytes As Long)
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Const Code_ForceRemain = &H700   'secret return code that forces a
                                        'plugin to remain loaded when an attempt is made to unload it.
Public Const ParserErrorBase = vbObjectError + &H45234
Public Declare Function KillTimer Lib "USER32.DLL" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SetTimer Lib "USER32.DLL" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function AppendMenu Lib "USER32.DLL" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "USER32.DLL" () As Long
Private Declare Function DeleteMenu Lib "USER32.DLL" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function TrackPopupMenu Lib "USER32.DLL" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByRef lprc As RECT) As Long
Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Declare Function GetCursorPos Lib "USER32.DLL" (ByRef lpPoint As PointAPI) As Long
Private Declare Function DestroyMenu Lib "USER32.DLL" (ByVal hMenu As Long) As Long

'Public Declare Function SafeArrayGetElement Lib "oleaut32.dll" (ByRef psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As Long


'HRESULT SafeArrayGetElement(
'  SAFEARRAY *  psa,
'  long *  rgIndices,
'  void *  pv
');
Public Declare Function SafeArrayGetElement Lib "oleaut32.dll" (ByRef psa As Any, ByVal rgIndices As Long, ByRef pv As Any) As Long


'HRESULT SafeArrayGetVartype(
'  SAFEARRAY *  psa,
'  VARTYPE *  pvt
');
Public Declare Function SafeArrayGetVartype Lib "oleaut32.dll" (ByRef psa As Any, ByRef VarTypeRet As Long) As Long

Private mAsyncCalls As DataStack    'holds Cparser objects.
'NOT USED YET :P

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
'Public OptimizedStacks As New Scripting.Dictionary
Public OptimizedStacks As DataStack
Public Labelled As New Collection


'Public OptimizedStacks As New DataStack
'Public FrmConfig As FrmConfig
'I don't like VB implicit form variables.
Public Enum ObjectMemberTypes
    Obj_Method = 2
    Obj_Prop_Let = 4
    Obj_Prop_Get = 8
    Obj_Prop_Set = 16
    obj_Prop_Any = Obj_Method + Obj_Prop_Let + Obj_Prop_Get + Obj_Prop_Set

End Enum
Private AsyncProcessing As Boolean  'flag set during asynchronous execution within
'the timer procedure.
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public mPluginDirContents() As String 'initialized in Main
Public mPluginDirContentCount As Long
'populated with the progIDs of all valid IEvalEvents and ICorePlugin implementors found in DLL and OCX files
'in our plugin folder.


Public Type SAFEARRAYBOUND
    cElements               As Long             ' # of elements in the array dimension
    lLbound                 As Long             ' lower bounds of the array dimension
End Type

Public Type SAFEARRAY
    cDims                   As Integer          ' // Count of dimensions in this array.
    fFeatures               As Integer          ' // Flags used by the SafeArray
                                                ' // routines documented below.
    cbElements              As Long             ' // Size of an element of the array.
                                                ' // Does not include size of
                                                ' // pointed-to data.
    cLocks                  As Long             ' // Number of times the array has been
                                                ' // locked without corresponding unlock.
    pvData                  As Long             ' // Pointer to the data.
    ' Should be sized to cDims:
    rgsabound()             As SAFEARRAYBOUND   ' // One bound for each dimension.
End Type

Private Declare Function SafeArrayGetDim Lib "oleaut32" (psa As Long) As Long
Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long


Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Function RegisterComponent(ByVal StrFilename As String)
    Dim TLIApp As TLIApplication
    Set TLIApp = New TLIApplication
    TLIApp.TypeLibInfoFromFile(StrFilename).Register



End Function



' FOR Access 97/VBE.dll clients like Word 97 and Excel 97
' Declare Function EbExecuteLine Lib "vba332.dll"  (ByVal pStringToExec As Long, ByVal Unknownn1 As Long,  ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long
'not used for the parser (Duh)
Public Function ExecuteLine(sCode As String, Optional fCheckOnly As Boolean) As Boolean
   ExecuteLine = EbExecuteLine(StrPtr(sCode), 0&, 0&, Abs(fCheckOnly)) = 0
End Function

'Public Function SecretFunction(a As Long, b As Long) As String
'  SecretFunction = "Secret calculation: " & a & " * " & b & " = " & a * b
'End Function

Public Function GetSafeArrayInfo(TheArray As Variant, ArrayInfo As SAFEARRAY) As Boolean
'===========================================================================
'   GetSafeArrayInfo - Fills a SAFEARRAY structure for the supplied array. The
'   information contained in the SAFEARRAY structure allows the caller to
'   identify the number of dimensions and the number of elements for each
'   dimension (among other things). Element information for each dimension is
'   stored in a one-based sub-array of SAFEARRAYBOUND structures (rgsabound).
'
'   TheArray        The array to get information on.
'   ArrayInfo       The output SAFEARRAY structure.
'
'   RETURNS         True if the array is instantiated.
'===========================================================================

    Dim lpData      As Long         ' Pointer to the variants data item
    Dim VType       As Integer      ' the VARTYPE member of the VARIANT structure

    ' Exit if no array supplied
    If Not IsArray(TheArray) Then Exit Function
    
    With ArrayInfo
    
        ' Get the VARTYPE value from the first 2 bytes of the VARIANT structure
        CopyMemoryLong ByVal VarPtr(VType), ByVal VarPtr(TheArray), 2
        
        ' Get the pointer to the array descriptor (SAFEARRAY structure)
        ' NOTE: A Variant's descriptor, padding & union take up 8 bytes.
        CopyMemoryLong ByVal VarPtr(lpData), ByVal (VarPtr(TheArray) + 8), 4

        ' Test if lpData is a pointer or a pointer to a pointer.
        If (VType And VT_BYREF) <> 0 Then

            ' Get real pointer to the array descriptor (SAFEARRAY structure)
            CopyMemoryLong ByVal VarPtr(lpData), ByVal lpData, 4
            
            ' This will be zero if array not dimensioned yet
            If lpData = 0 Then Exit Function
            
        End If

        ' Fill the SAFEARRAY structure with the array info
        ' NOTE: The fixed part of the SAFEARRAY structure is 16 bytes.
        CopyMemoryLong ByVal VarPtr(ArrayInfo.cDims), ByVal lpData, 16

        ' Ensure the array has been dimensioned before getting SAFEARRAYBOUND information
        If ArrayInfo.cDims > 0 Then

            ' Size the array to fit the # of bounds
            ReDim .rgsabound(1 To .cDims)

            ' Fill the SAFEARRAYBOUND structure with the array info
            CopyMemoryLong ByVal VarPtr(.rgsabound(1)), ByVal lpData + 16, ArrayInfo.cDims * Len(.rgsabound(1))

            ' So caller knows there is information available for the array in output SAFEARRAY
            GetSafeArrayInfo = True
            
        End If

    End With

End Function




'Public Function IsOptimized(ByVal Expression As String) As Boolean
'
'
'End Function
'Public Function GetPStack(ByVal ForExpression As String) As cFormitem
'
'
'
'End Function
Public Sub ExecuteAsync(ParseObject As CParser)
    'push onto the stack.
    mAsyncCalls.Push ParseObject
    If Not AsyncProcessing Then
    
    
    End If




End Sub
Public Function GetArrElement(ByRef ArrayGet As Variant, indices() As Long) As Variant
    Dim elementPtr As Variant, retval As Long
    retval = SafeArrayGetElement(ByVal ArrayGet, indices(0), elementPtr)
    GetArrElement = elementPtr
    

End Function
Public Sub Testgetelement()
    Dim testarr, indices() As Long
    Dim gotval As Variant
    ReDim indices(0)
    indices(0) = 5
    testarr = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    gotval = GetArrElement(testarr, indices())


End Sub

Public Function Random(ByVal iLo As Long, ByVal iHi As Long, Optional ByVal AllowFloat As Boolean = True) As Variant

If Not AllowFloat Then
    Random = Int(iLo + (Rnd * (iHi - iLo + 1)))
Else
    Random = CDec(iLo + (Rnd * (iHi - iLo + 1)))
End If

End Function


Public Function GetPluginsFromDir(ByVal Dirsearch As String, Optional ByRef retcount As Long) As String()
    'returns a list of Plugins, loaded from DLLs and OCX files in the given directory.
    'the ProgIDs are verified before being returned.
    Dim FileNames() As String
    Dim Currfilename As String
    Dim fcount As Long, SearchSpecs(1 To 2) As String
    Dim currspec As Long
    Dim CurrLib As Long, buildProgID As String
    Dim Ret() As String
    Dim tlInfo As TypeLibInfo
    fcount = 0
    If right$(Dirsearch, 1) <> "\" Then Dirsearch = Dirsearch & "\"
    SearchSpecs(1) = Dirsearch & "*.dll"
    SearchSpecs(2) = Dirsearch & "*.ocx"
    
    For currspec = 1 To 2
        Currfilename = Dir$(SearchSpecs(currspec))
        Do Until Currfilename = ""
            fcount = fcount + 1
            ReDim Preserve FileNames(1 To fcount)
                FileNames(fcount) = Dirsearch & Currfilename
                
            
            Currfilename = Dir$
            
        Loop
    Next currspec
    Dim loopclass As CoClassInfo, loopinterfaces As InterfaceInfo, getparent As TypeLibInfo
    For CurrLib = 1 To fcount
        On Error Resume Next
        
        Set tlInfo = TLIApplication.TypeLibInfoFromFile(FileNames(fcount))
        If Err = 0 Then
            For Each loopclass In tlInfo.CoClasses
            'loop on Coclasses...
            For Each loopinterfaces In loopclass.Interfaces
                Debug.Print loopinterfaces.Name, tlInfo.Name
                'the loopinterfaces.name should be "_IEvalEvents"
                If StrComp(loopinterfaces.Name, "_IEvalEvents", vbTextCompare) = 0 Then
                    Debug.Print loopinterfaces.Parent.Name
                    Set getparent = loopinterfaces.Parent
                    If getparent.Name = "BASeParserXP" Then
                        'lastly- make sure it is the right version!
                        If getparent.MajorVersion < App.Major Then
                            CDebug.Post "impending load issue:" & loopclass.Name & " Version of Implemented interface differs from this library."
                        End If
                        
                        'ElseIf getparent.MajorVersion = App.Major Then 'getparent.MajorVersion >= app.Major
                        'hmmm....
                            If getparent.MinorVersion < App.Minor Then
                                CDebug.Post "Version error:  " & loopclass.Name & " - Minor version mismatch."
                            
                            Else
                                'Alright- it all checks out.
                                buildProgID = loopclass.Parent.Name & "." & loopclass.Name
                                retcount = retcount + 1
                                ReDim Preserve Ret(1 To retcount)
                                Ret(retcount) = buildProgID
                            
                            End If
                        
                        
                        'End If
                    
                    End If
                End If
            
            Next loopinterfaces
            
            Next
        
        
        End If
    
    
    Next CurrLib
    
    
    

    GetPluginsFromDir = Ret
End Function
Public Sub Main()
    
   ' we need to call InitCommonControls before we
   ' can use XP visual styles.  Here I'm using
   ' InitCommonControlsEx, which is the extended
   ' version provided in v4.72 upwards (you need
   ' v6.00 or higher to get XP styles)
   
   'Init the debug object.
  
'    Call mSysTray.CreateIcon(0, LoadResPicture("TRAYICO", vbResIcon).Handle, "BASeParser")
    mPluginDirContents = GetPluginsFromDir(App.Path & "Plugins", mPluginDirContentCount)
   On Error Resume Next
   ' this will fail if Comctl not available
   '  - unlikely now though!
   'initialize the Resources.
   'note that we still include a resource file in
   'this project so that we
   'can properly have XP themes.
   'initialize the resource grabber.
   'InitResGrabber "BPROOS1033.CResGrabber"
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
 Set mAsyncCalls = New DataStack
    Set ParserSettings = New ParserSettings
    
    Set OptimizedStacks = New DataStack
   'create the Optimized Stacks object, set to FIFO, and
   'assign the maxitems.
    With OptimizedStacks
        .stackmode = FirstInFirstOut
        .MaxItems = 48
        .AutoPurge = True
    End With
    'no code yet.

   'Set FrmDebug.DebugObj = CDebug
   'FrmDebug.RefreshList
   'FrmDebug.Show
End Sub
Public Function ParserArrToResultArr(Parsers As Variant, Optional ByRef ResultCount As Long) As Variant()
'converts an array of parsers packed into a variant into a variant array of the results from those
'parsers.
Dim results() As Variant
Dim I As Long
    ReDim results(0)
    If IsEmpty(Parsers) Then Exit Function
If Not IsArray(Parsers) Then
    If IsObject(Parsers) Then
        If TypeOf Parsers Is CParser Then
        
            results(0) = Parsers.Execute
        Else
            'not of type CParser
        End If
        
    Else
        'not even an object.
        'ok- just put that value as the result, I suppose.
        'not much else we COULD do-
        results(0) = Parsers
        Exit Function
    End If
End If
ReDim results(0 To (UBound(Parsers) - LBound(Parsers) + 1))
For I = LBound(Parsers) To UBound(Parsers)
    results(I - LBound(Parsers)) = Parsers(I).Execute

Next I
    ParserArrToResultArr = results
    ResultCount = UBound(results) + 1
End Function

Public Function workaround(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
    ByVal dwTime As Long) As Long
    workaround = AsyncCParser_Timer(hwnd, uMsg, idEvent, dwTime)
End Function
'Feb 23 2007- I have yet to actually implement the asynchronous code. It shouldn't be to bad, just push the Parser object onto
'the AsyncCalls Stack, fire up the timer, and away we go- the stack is just in case another parser start's asynchronous
'execution. It'll push itself onto the stack. I don't think that
'I need to worry about synchronization, since I am not actually dealing with threads, but more like Fibers.
'since execution (generally) is an operation whose result is needed for the client to continue,
'asynchronous calls are made only for Parsing, which generally takes longer, and oftentimes can be completed before the
'corresponding execute is called.
Public Sub InsertImplicit(mStartItem As CFormItem)
'looks through the specified Linked List and insert a IT_OPERATOR (*) between any two of IT_VALUE,IT_FUNCTION,IT_SUBEXPRESSION, and so on.
Exit Sub
Dim CreatedItem As CFormItem
Dim loopitem As CFormItem
Dim TempItem As CFormItem
Dim ret1 As Boolean, ret2 As Boolean
Set loopitem = mStartItem
Do Until loopitem.Next_ Is Nothing
    ret1 = False: ret2 = False
    With loopitem
    'check the current item and the next one.
        ret1 = (.ItemType = IT_FUNCTION Or .ItemType = IT_VALUE Or .ItemType = IT_VARIABLE)
    End With
    With loopitem.Next_
        ret2 = (.ItemType = IT_FUNCTION Or .ItemType = IT_VALUE Or .ItemType = IT_VARIABLE)
    End With
    If ret1 And ret2 Then
        CDebug.Post "Tested Positive for Implicit multiplication."
        'now the tricky part. the insertion of an item between loopitem and it's next.
        Set TempItem = loopitem.Next_
        Set CreatedItem = New CFormItem
        CreatedItem.ItemType = IT_OPERATOR
        CreatedItem.op = "*"
        CreatedItem.Value = "*"
        CreatedItem.ExprPos = loopitem.Next_.ExprPos
        
        Set CreatedItem.Prev = loopitem
        Set CreatedItem.Next_ = TempItem
        Set loopitem.Next_ = CreatedItem
        Set TempItem.Prev = CreatedItem
        'there, safely inserted (theoretically)
    End If
    Set loopitem = loopitem.Next_
Loop




End Sub
Public Function AsyncCParser_Timer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
    ByVal dwTime As Long) As Long
    Dim Popped As CParser
    AsyncProcessing = True
    
    'always kill the timer first, at least, for VB.
    KillTimer hwnd, idEvent
    'ASynchronous timer.
    'this is a pretty common way to force VB to create a new STA.
    'The thing is, What about multiple CParser objects? what happens then?
    'well, there are two possible things to do.
    'First, I could only keep track of the last one Set.
    'that is a dumb idea, stuff could really get screwed.
    'Or, I use a DataStack to store them, popping during each call to the function, which
    'eventually kills itself when the stack empties.
    
    If mAsyncCalls.Count > 0 Then
        'if there are Parsers waiting to be executed, don't just sit here! Pop one off and start
        'it's execute method. Of course, somebody will need to sink that specific Parser's Events.
        Set Popped = mAsyncCalls.Pop
        're-set the timer during the execute.
        'this way, we'll re-enter this procedure and pop the next parser and call it's Execute
        'method asynchronously with all the others.
        idEvent = SetTimer(hwnd, idEvent, 1, AddressOf workaround)
        Popped.ParseInfix Popped.Expression

        'reset the timer, since we just deleted it.
        
    Else
        'break out.
        'and don't reset the timer.
    End If
    AsyncProcessing = False
'   End of Old Async Code
End Function


Public Sub InvokeArrayMethod(withparser As CParser, ArrInvoke, ByVal Member As String, Arguments As Variant, ByRef retval As Variant)
    'invokes an array method. Array methods are built-into the parser-
    'for example, this makes the following valid:
    
    'Array(1,2,3)@Item(2)
    'or
    '{1,2,3}@Length()
    Dim paramcount As Long
    
    On Error Resume Next
    paramcount = UBound(Arguments) + 1
    If Err <> 0 Then paramcount = 0
    Err.Clear
    Member = UCase$(Member)
    'Currently supported methods/properties:
    'Length & Count & Dim
    'Size (adds LEN of each element)
    'Item (returns the item)
    'StdDev (Standard Deviation)
    'Average (Average)
    'Minimum/Min
    'Maximum/Max
    'Median/Med
    'Mean
    'Intersection
    'Difference
    'Exclusion
    Select Case Member
        Case "LENGTH", "COUNT", "DIM"
            retval = UBound(ArrInvoke)
        Case "SIZE"
            'Why do MY functions have to do EVERYTHING!
            
            retval = Array_TotalSize(ArrInvoke)
        Case "ITEM"
            'lazy bastards.
            'damn, just use the friggin [] like everyone else!
            'well- I can't see how their lazy, I mean- they did type more...
            Call Assign(retval, ArrInvoke(Arguments(0).Execute))
        Case "MEAN", "AVERAGE"
            Call Assign(retval, Array_Average(withparser, ArrInvoke))
        Case "SORT"
            'sort has a bunch of args.
            Select Case paramcount
                Case 0
                    Call SortArray(ArrInvoke)
                Case 1
                    Call SortArray(ArrInvoke, Arguments(0).Execute)
                Case 2
                    Call SortArray(ArrInvoke, Arguments(0).Execute, Arguments(1).Execute)
                Case 3
                    Call SortArray(ArrInvoke, Arguments(0).Execute, Arguments(1).Execute, , Arguments(2).Execute)
            End Select
            retval = ArrInvoke
        Case "SELECT"
            'select arguments(0).execute values from Arrinvoke.
            Assign retval, SelectItems(withparser, ArrInvoke, Arguments(0).Execute)
        Case "SUB"
            'retrieve a "SubArray" which is my retarded name for a portion of the array.
            'takes two arguments.
            'the first item, and the number of items.
            If UBound(Arguments) = 0 Then
                retval = SubArray(ArrInvoke, Arguments(0).Execute, Arguments(1).Execute)
            ElseIf UBound(Arguments) > 0 Then
                retval = SubArray(ArrInvoke, Arguments(0).Execute, Arguments(1).Execute)
            End If
        Case "APPEND", "CONCAT"
            retval = Array_Append(ArrInvoke, Arguments(0).Execute)
        Case "ISECT", "INTERSECT", "INTERSECTION"
            Call Assign(retval, Array_Intersection(withparser, ArrInvoke, Arguments(0).Execute))
        Case "DIFF", "DIFFERENCE"
            Call Assign(retval, Array_Difference(withparser, ArrInvoke, Arguments(0).Execute))
        Case "EXCLUDE", "EXCLUSION"
            Call Assign(retval, Array_Exclusion(withparser, ArrInvoke, Arguments(0).Execute))
        Case "MAX"
            Call Assign(retval, Array_Max(withparser, ArrInvoke))
        Case "MIN"
            Call Assign(retval, Array_Min(withparser, ArrInvoke))
            Stop
        Case "UBOUND"
            retval = UBound(ArrInvoke)
        Case "LBOUND"
            retval = LBound(ArrInvoke)
        
        
    End Select
    
    
End Sub
'Array Functions. ONLY VALID ON ONE-DIMENSIONAL ARRAYS!
Public Function SubArray(ArrFrom, ByVal Start As Long, Optional ByVal Length As Long = -1) As Variant
    If Length = -1 Then Length = (UBound(ArrFrom) - Start) + 1
    Dim Ret
    Dim I As Long
    ReDim Ret(Length - 1)
    For I = Start To Start + Length - 1
        Ret(I - Start) = ArrFrom(I)
    
    
    Next I

    SubArray = Ret

End Function
Public Sub performArrayOp(withparser As CParser, OnArray As Variant, Formitem As CFormItem, ByRef retval As Variant)
    'perform the array operation.
    InvokeArrayMethod withparser, OnArray, Formitem.op, Formitem.Value, retval



End Sub
Function StdDev(Numbers As Variant) As Double
          'Calculate the Standard deviation of a set of numbers.
Dim Total As Double
Dim Sum As Double
Dim I As Long
For I = 0 To UBound(Numbers)
     Total = Total + (Numbers(I) ^ 2)
     Sum = Sum + Numbers(I)
Next I
Sum = Sum / UBound(Numbers) + 1
Total = Total / UBound(Numbers) + 1
StdDev = Sqr(Abs(Total - (Sum ^ 2)))
End Function
Public Function Repeat(ByVal StrString As String, ByVal Number As Double) As String
    Dim I As Long
    Dim Strret As String
    Dim numextra As Long
    If Number < 0 Then
        Err.Raise 5, "modParser::Repeat", "Cannot repeat a string a negative number of times."
    Else
        For I = 1 To Int(Number)
        
            Strret = Strret & StrString
        Next I
        'add any fractional component.
        If Int(Number) <> Number Then
                numextra = Len(StrString) * (Number - Int(Number))
                Strret = Strret & Mid$(StrString, 1, numextra)
        
        End If
    End If
    Repeat = Strret
    
End Function
Public Function Solve(withparser As CParser, ByVal Expression As String, ByVal VarName As String, ByVal LowerBound As Variant, ByVal UpperBound As Variant, Optional tolerance = 0.0001) As Variant
'solves the given expression, within the given bounds and within a tolerance of tolerance.


'IMPORTANT! withparser should probably be a numeric expression (duh)
'also, it tests for roots. if you want to find the "intersection" with another function, simple-
'make a new function (Function1 - function2) and find the root of that function.

Static Subparser As CParser
Static Initialized As Boolean
Static Varobj As CVariable
Dim lowReturn As Variant, HighReturn As Variant, midreturn As Variant
Dim MidLeft As Variant, MidRight As Variant
Dim diff
Dim middle
If Subparser Is Nothing Then Initialized = False

If Initialized Then
    'are we REALLY?
    If Not (Subparser Is Nothing) Then
        If (Not Subparser Is withparser) And Not (withparser Is Nothing) Then
            'we are initialized- but not to the right parent.
            Set Subparser = Nothing
            Initialized = False
        End If
    Else
        'is nothing.
        Initialized = False
    End If
End If
If Not Initialized Then
    'Note- if initialized is true, the first (non-recursive) would have
    'executed this code already.
    'thus they didn't even give us any arguments.
    Set Subparser = withparser.Clone()
    Subparser.Expression = Expression
    Set Varobj = Subparser.Variables.Item(VarName)
    Initialized = True
End If

'use root-bisection method :).
'assign the value to variable...
diff = UpperBound - LowerBound
middle = (diff / 2 + LowerBound)
MidRight = ((diff / 2) + LowerBound) + (diff / 4)
'middle right- one half is middle left.
Varobj.Value = middle
midreturn = Subparser.Execute
If Abs(midreturn) < tolerance Then
    Solve = Varobj.Value
    Exit Function
End If
MidLeft = MidRight - (diff / 2)
Varobj.Value = LowerBound
lowReturn = Subparser.Execute
Varobj.Value = UpperBound
HighReturn = Subparser.Execute

If Abs(lowReturn) < Abs(HighReturn) Then
    'follow the lowreturn side.
    Solve = Solve(Subparser, "", "", LowerBound, middle, tolerance)
    
    
    
Else
    'follow the highreturn side first.
    Solve = Solve(Subparser, "", "", middle, UpperBound, tolerance)
End If

End Function
Private Function ForceLBZero(OnArr As Variant)
    'Force the lowerbound to be zero.
    Dim newArray
    Dim I As Long
    ReDim newArray(0 To UBound(OnArr) - LBound(OnArr) + 1)
    For I = LBound(OnArr) To UBound(OnArr)
        newArray(I - LBound(OnArr)) = OnArr(I)
    
    Next I
    OnArr = newArray
End Function
Public Sub ForceArray(ValForce As Variant)
    'forces a variant into an array.
    Dim Temp As Variant
    If IsArray(ValForce) Then
        'it better have a lower bound of 0.
        If LBound(ValForce) <> 0 Then
            ForceLBZero ValForce
        End If
        Exit Sub
    Else
        Call Assign(Temp, ValForce)
        ReDim ValForce(0)
        Assign ValForce(0), Temp
    End If
    
        
End Sub
Public Sub Array_Push(ArrayOn As Variant, PushVal As Variant)
    ForceArray ArrayOn
    ReDim ArrayOn(UBound(ArrayOn) + 1)
    ArrayOn(UBound(ArrayOn)) = PushVal
    
End Sub
Public Function Array_Pop(ArrayPop As Variant) As Variant
    ForceArray ArrayPop
    Array_Pop = ArrayPop(UBound(ArrayPop))
    ReDim Preserve ArrayPop(UBound(ArrayPop) - 1)
    
End Function
Public Function Array_Shift(ArrayShift As Variant) As Variant
    'takes the first item off the array and
    'shifts all elements down.
    Dim RetItem As Variant, CurrMove As Long
        ForceArray ArrayShift
    Assign RetItem, ArrayShift(LBound(ArrayShift))
    For CurrMove = LBound(ArrayShift) + 1 To UBound(ArrayShift)
        Assign ArrayShift(CurrMove), ArrayShift(CurrMove - 1)
    Next
    'delete the last element.
    ReDim Preserve ArrayShift(LBound(ArrayShift) To UBound(ArrayShift) - 1)
    
    'tada.
    If IsObject(RetItem) Then Set Array_Shift = RetItem Else Array_Shift = RetItem
    
End Function
Public Function Array_Frequency(withparser As CParser, dataArray, BinsArray, Optional ByVal FlAllowEqual As Boolean = False) As Variant
    'returns the number of elements in dataarray that occur in the
    'intervals given by binsArray.
    'for example:
    'Array_Frequency({1,2,3,4,5,6,7,8,9},{{1,5}}) would give us {5}
    Dim retArr
    Dim Iterate As Long
    Dim LookBin As Long, checkbin As Variant
    Dim CompLess, CompMore
    Dim UseInterval As Variant
    Dim rettemp As Variant
    If withparser Is Nothing Then
        Set withparser = New CParser
        withparser.Create
    End If
    ForceArray dataArray
    ReDim retArr(UBound(BinsArray) + 1)
    'the last one is the number of elements larger then the largest
    'given interval.
    'for convenience, sort the array. This way, we'll be working with the data
    'from smallest to largest.
    SortArray dataArray
    'ta-da
    For Iterate = LBound(dataArray) To UBound(dataArray)
        'should be going from smallest to largest.
        'now, iterate through all the BinArray elements. If the given value falls between the
        'two values, add 1 to the corresponding element in RetArr.
        For LookBin = LBound(BinsArray) To UBound(BinsArray)
            Assign checkbin, BinsArray(LookBin)
            Assign CompLess, Min(withparser, checkbin(0), checkbin(1))
            Assign CompMore, Max(withparser, checkbin(0), checkbin(1))
            withparser.PerformOperation "<=>", dataArray(Iterate), CompLess, rettemp
            
            If rettemp = 1 Or (FlAllowEqual And rettemp = 0) Then
            'if is larger then the smaller bound.
                withparser.PerformOperation "<=>", dataArray(Iterate), CompMore, rettemp
                If rettemp = -1 Or (FlAllowEqual And rettemp = 0) Then
                    'if is smaller then the larger bound, as well.
                    'add one to the corresponding bin:
                    retArr(LookBin) = retArr(LookBin) + 1
                    
                
                End If
            
            End If
        Next LookBin
        
    
    Next Iterate
    Array_Frequency = retArr
    
    
    
    
'    For Iterate = 0 To UBound(BinsArray)
'
'        UseInterval = BinsArray(Iterate)
'            If VarType(UseInterval) = vbString Then
'                'split.
'                'of course, strings aren't parsed...
'                UseInterval = Split(UseInterval, ",")
'
'            End If


'    Next
    




End Function

'Public Function Object_Members(Withparser As CParser, ObjFrom As Object) As Variant
'    'like the foxpro AMEMBERS command.
'    'not implemented yet....
'    Dim IInfo As InterfaceInfo
'
'
'
'End Function
Public Function Array_GetElement(withparser As CParser, ArrayOn As Variant, indices As Variant) As Variant
    'Darn! I just CAN'T get the darn SafeArray Method to work. Oh well,
    'I had a workaround this whole time- nested array's.
    'Anyway, this is pretty simple- just loop through for each element in indices and save
    'the item retrieved from that spot from arrayOn.
    
    Dim CacheItem As Variant
    Dim currindex As Variant
    On Error Resume Next
    If Not IsArray(indices) Then
        CacheItem = indices
        ReDim indices(0)
        indices(0) = CacheItem
        'Tada.
        CacheItem = vbEmpty
    
    End If
    CacheItem = ArrayOn
    For currindex = LBound(indices) To UBound(indices)
        
        
        'if cacheItem is not an Array, raise an error of the sort, "Surplus Array Indices passed." or something.
        'Or, just ignore the error and return the item we do have.
        'OR!- have which behaviour to perform stored as a setting...
        
        If Not IsArray(CacheItem) Then
            If CBool(ParserSettings.ErrorOnSurplusArrayIndices(withparser.configset)) Then
                'raise error.
                'a PARSER error, no doubt.
                'ParserError(WithParser, 9, "To Many dimensions Specified.", "ModParser::Array_GetElement").Throw
                Err.Raise 9, "Array_GetElement", "Too Many Dimensions Specified."
            Else
                'don't raise an error, just return what we had so far.
                Array_GetElement = CacheItem
                Exit Function
            
            
            End If
        
        
        End If
        
        'Check bounds...
        'could just trap the actual error, but we can test it here, so What the heck.
        
        
        If LBound(CacheItem) > indices(currindex) Or _
            UBound(CacheItem) < indices(currindex) Then
            On Error GoTo 0
                'ParserError(WithParser, 9, "SubScript out of range in Dimension " & CurrIndex, "Array_GetElement").Throw
                Err.Raise 9, "Array_getelement", "subscript out of range in dimension " & currindex
        End If
        CacheItem = CacheItem(indices(currindex))
    
    
    Next
    Array_GetElement = CacheItem
End Function
'Public Function Array_GetElement(ArrayOn As Variant, Indices As Variant) As Variant
'
'    'Indices: array of indexes into the array.
'    Const DISP_E_BADINDEX As Long = &H8002000B
'    Const E_OUTOFMEMORY As Long = &H8007000E
'
'
'    Dim IndUse() As Long, I As Long
'    Dim vartypeCheck As VbVarType
'    Dim ElemGrab As Variant
'    Dim success As Long
'    Dim arrCopy() As Variant
'    arrCopy = ArrayOn
'    ReDim IndUse(0 To UBound(Indices) - LBound(Indices) + 1)
'    For I = 0 To UBound(IndUse) - 1
'        IndUse(I) = Indices(I + LBound(Indices))
'    Next I
''    F = SafeArrayGetElement(3, 3, 3)
'    success = SafeArrayGetVartype(ByVal VarPtrArray(arrCopy()), vartypeCheck)
'
'    success = SafeArrayGetElement(ByVal VarPtrArray(arrCopy()), IndUse(0), VarPtr(ElemGrab))
'
'
'
'
'End Function
Public Function Array_Append(ArrayOn As Variant, OtherArray As Variant) As Variant
    'Appends OtherArray to the end of arruse.
    
    'FORCE it into the array.
    Dim StartAppend As Long, NewUB As Long
    Dim I As Long, arruse As Variant
    arruse = ArrayOn
    ForceArray arruse
    ForceArray OtherArray
    'now, resize to allow for the elements in otherarray.
    StartAppend = UBound(arruse) + 1
    NewUB = StartAppend + (UBound(OtherArray) - LBound(OtherArray))
    ReDim Preserve arruse(NewUB)
    For I = StartAppend To NewUB
        Call Assign(arruse(I), OtherArray(I - StartAppend))
    
    Next I
    

    
    Array_Append = arruse

End Function
Public Function Array_Difference(withparser As CParser, Arr1 As Variant, Arr2 As Variant) As Variant
    'Any given value in the result is in A but NOT in B.
    ForceArray Arr1
    ForceArray Arr2
    Dim Result As Variant
    Dim NumHits As Long, I As Long
    
    Debug.Assert LBound(Arr1) = 0   'should be true.(forcearray should be making it so)
    
    
    'thie biggest possible return value
    'is the size of Arr1- given that
    'All items in A are not in B, and A has no dupes.
    
    ReDim Result(0 To UBound(Arr1) + 1)
    For I = 0 To UBound(Arr1)
        If Among(withparser, Arr1(I), Arr2) = False And Not Among(withparser, Arr1(I), Result) Then
            Result(NumHits) = Arr1(I)
            NumHits = NumHits + 1
        End If
    Next I
    ReDim Preserve Result(NumHits - 1)
    Array_Difference = Result
End Function
Public Function Array_Exclusion(withparser As CParser, Arr1 As Variant, Arr2 As Variant) As Variant
   ForceArray Arr1
    ForceArray Arr2
    Dim NumHits As Long
    Dim CurrA As Long
    Dim Result As Variant
    Dim merged As Variant
    'Intially, redim the result to the size of the larger one. This is the maximum size.
    ReDim Result(Max(withparser, UBound(Arr1), UBound(Arr2)))
    merged = Array_Append(Arr1, Arr2)
    For CurrA = 0 To UBound(merged)
        If Among(withparser, merged(CurrA), Arr1) Xor Among(withparser, merged(CurrA), Arr2) And Not Among(withparser, merged(CurrA), Result) Then
            Assign Result(NumHits), merged(CurrA)
            NumHits = NumHits + 1
        End If
    Next CurrA
    ReDim Preserve Result(NumHits - 1)
    
    Array_Exclusion = Result
End Function
Public Function Array_Intersection(withparser As CParser, Arr1 As Variant, Arr2 As Variant) As Variant
    'create a new array containing only those
    'items in both Arr1 and Arr2.
    ForceArray Arr1
    ForceArray Arr2
    Dim NumHits As Long
    Dim CurrA As Long
    Dim Result As Variant
    'Intially, redim the result to the size of the smaller one. This is the maximum size.
    ReDim Result(Min(withparser, UBound(Arr1), UBound(Arr2)))
    For CurrA = LBound(Arr1) To UBound(Arr1)
        If Among(withparser, Arr1(CurrA), Arr2) Then
            If Not Among(withparser, Arr1(CurrA), Result) Then
                Call Assign(Result(NumHits), Arr1(CurrA))
                NumHits = NumHits + 1
            End If
        End If
    Next
    If NumHits > 0 Then
        ReDim Preserve Result(NumHits - 1)
    Else
        Erase Result
    End If
    Array_Intersection = Result
    
    
    
    
End Function
Public Function Array_AveDev(withparser As CParser, arruse As Variant) As Variant
'
'Returns average deviation from the mean.
Dim avg As Variant, I As Long
Dim Diffs As Variant
avg = Array_Average(withparser, arruse)
ReDim Diffs(UBound(arruse))
For I = 0 To UBound(arruse)
    Diffs(I) = Abs(arruse(I) - avg)

Next I
Array_AveDev = Array_Average(withparser, Diffs)
Erase Diffs

End Function
Public Function GetTempFile(ByVal Prefix As String, Optional ByVal Extension As String = "TMP") As String
'GetTempFile: returns a non-existent filename in the users TEMP directory.
    Dim tPath As String
    Dim tfile As String
    tPath = Space$(255)
    tfile = Space$(255)
    GetTempPath 255, tPath
    tPath = Replace$(Trim(tPath), vbNullChar, "")
    GetTempFileName tPath, Prefix, 0, tfile
    tfile = Replace$(Trim(tfile), vbNullChar, "")
    tfile = Mid$(tfile, 1, InStrRev(tfile, "."))
    tfile = tfile & Extension
    If right$(tPath, 1) <> "\" Then tPath = tPath & "\"
    GetTempFile = tfile





End Function
Public Function Among(mParser As CParser, ByVal TestFor As Variant, TestIn As Variant) As Boolean
    Dim I As Long
    Among = True
    For I = LBound(TestIn) To UBound(TestIn)
        If IsObject(TestIn(I)) And IsObject(TestFor) Then
            If TestIn(I) Is TestFor(I) Then Exit Function
        ElseIf IsObject(TestIn(I)) Xor IsObject(TestFor) Then
            'err-
                If compare(mParser, TestIn(I), TestFor) = 0 Then Exit Function
        Else
            'neither.
            'If TestIn(I) = "fifty" And TestFor(I) = "fifty" Then Stop
            If compare(mParser, TestIn(I), TestFor) = 0 Then Exit Function
        End If
    Next I
    Among = False




End Function
Public Function Array_TotalSize(arruse As Variant) As Long
    Dim AccumSize As Variant, I As Long
    'irregardless of the type of values within the array-
    'all Types have a Length- Objects jet get a Long pointer 4.
    For I = LBound(arruse) To UBound(arruse)
        AccumSize = AccumSize & arruse
    Next I


    Array_TotalSize = AccumSize

End Function
Public Function CreateArray(ByVal lb As Long, ByVal ub As Long, Optional DefaultValue As Variant = vbEmpty) As Variant
    Dim createArr, I As Long
    ReDim createArr(lb To ub)
    If Not IsEmpty(DefaultValue) Then
        For I = LBound(createArr) To UBound(createArr)
            createArr(I) = DefaultValue
        Next I
    End If
    CreateArray = createArr

End Function
Public Function Array_Length(arruse As Variant) As Variant
    Array_Length = (UBound(arruse) - LBound(arruse) + 1)
End Function
Public Function Array_Average(mParser As CParser, arruse As Variant) As Variant

    'Average-
    'finds the average of the values.
    'like the other array values, it ignores non-numeric entries.
    'uses CDec to try to get a nice, accurate value.
    'Dim CurrAvg As Variant
    Dim CurrAccum As Variant
    Dim I As Long
    For I = LBound(arruse) To UBound(arruse)
        'If IsNumeric(arruse(I)) Then
           ' CurrAccum = CurrAccum + arruse(I)
           If TypeOf arruse(I) Is CComplex Then
            'Debug.Assert False
        End If
           mParser.PerformOperation "+", CurrAccum, arruse(I), CurrAccum
           
        'End If
    Next I
   ' Array_Average = CDec(CurrAccum) / Array_Length(arruse)
   'lastly, divide by array length.
   Dim retme As Variant
   mParser.PerformOperation "/", CurrAccum, Array_Length(arruse), retme
    If Not IsObject(retme) Then Array_Average = retme Else Set Array_Average = retme
End Function
Public Function Array_Max(mParser As CParser, arruse As Variant, Optional ByRef GetIndex As Long) As Variant
    'returns the minimum value of the array.
    'first off, it assumes that all the values are numeric.
    'if it encounters a non-numeric value,
    'it is ignored (unless it is a numeric string)
    Dim I As Long, CurrMax, HasMax As Boolean
    Dim compareresult As Variant
    GetIndex = LBound(arruse) - 1
    
    For I = LBound(arruse) To UBound(arruse)
        If IsNumeric(arruse(I)) Then
        'if it is bigger then our current Maximum, or if we don't even
        'have a current Maximum, assign it.
            mParser.PerformOperation ">", arruse(I), CurrMax, compareresult
            If compareresult Or Not HasMax Then
            
                Assign CurrMax, arruse(I)
                'save the index too.
                GetIndex = I
                HasMax = True
        
            End If
        
        End If
    
    
    
    
    Next I
    
    
    If Not IsObject(CurrMax) Then Array_Max = CurrMax Else Set Array_Max = CurrMax
    
    
    
    
    
    
End Function
Public Function Min(withparser As CParser, ParamArray Items() As Variant) As Variant
    Dim coerce
    coerce = Items
    Min = Array_Min(withparser, coerce)
End Function
Public Function Max(withparser As CParser, ParamArray Items() As Variant) As Variant
    Dim coerce
    coerce = Items
    Max = Array_Max(withparser, coerce)
End Function

Public Function Array_Min(withparser As CParser, arruse As Variant, Optional ByRef GetIndex As Long) As Variant
    'returns the minimum value of the array.
    'first off, it assumes that all the values are numeric.
    'if it encounters a non-numeric value,
    'it is ignored (unless it is a numeric string)
    Dim I As Long, CurrMin, HasMin As Boolean
    Dim evaltmp As Variant
    
    For I = LBound(arruse) To UBound(arruse)
        If IsNumeric(arruse(I)) Then
        'if it is smaller then our current minimum, or if we don't even
        'have a current minimum, assign it.
            'If arruse(I) < CurrMin Or Not HasMin Then
            If HasMin Then
                Call withparser.PerformOperation("<", arruse(I), CurrMin, evaltmp)
            End If
            If CBool(evaltmp) Or Not HasMin Then
            
            
            
                Call Assign(CurrMin, arruse(I))
                'save the index too.
                GetIndex = I
                HasMin = True
        
            End If
        
        End If
    
    
    
    
    Next I
    
    
    If Not IsObject(CurrMin) Then Array_Min = CurrMin Else Set Array_Min = CurrMin
    
    
    
    
    
    
End Function

Public Function RandomArray(ByVal NumElements As Long, ByVal LowerBound As Double, ByVal UpperBound As Double, Optional ByVal AllowFloats As Boolean = True) As Variant
    Dim retThis As Variant, I As Long
    ReDim retThis(0 To NumElements - 1)
    For I = 0 To NumElements - 1
        retThis(I) = Random(LowerBound, UpperBound, AllowFloats)
    
    Next I




End Function


Public Sub InvokeDynamic(ByVal onObj As Object, ByVal Membername As String, Arguments As Variant, retval As Variant)
    Dim ArgsPass() As Variant
    Dim AccessMode As InvokeKinds
    Dim I As Long
    Dim Current As Long
    Dim TryEm(10) As InvokeKinds, TryCount As Long
    Dim MemberInf As tli.SearchItem
    Dim FuncFlags As InvokeKinds
    Dim IntInfo As tli.InterfaceInfo, Getfirst As Boolean
    On Error Resume Next
    Set IntInfo = tli.InterfaceInfoFromObject(onObj)
    'Set MemberInf=IntInfo.Members.GetFilteredMembers(False).Item(
     Set MemberInf = FindSearchItem(IntInfo.Members.GetFilteredMembers(False), Membername)
    If MemberInf Is Nothing Then
    'Execution Error: Method or Data member not found.
 
            On Error GoTo 0
            Err.Raise ExecuteErrors.Exec_UnsupportedOperation, "ModParser::InvokeDynamic", "Interface member name " & Membername & " Not present in class name " & TypeName(onObj)

    Else
        'The method DOES exist. good.
        'if we have an array of arguments, put them in reverse order for  the InvokeHook...
         If IsArray(Arguments) Then
            ReDim ArgsPass(UBound(Arguments))

            For I = UBound(Arguments) To 0 Step -1
                
                If IsObject(Arguments(I)) Then
                    
                    Set ArgsPass(Current) = Arguments(I)
                
                Else
                    
                    ArgsPass(Current) = Arguments(I)
                
                End If
                
                Current = Current + 1
            Next I
        End If
    'OK, if we have no parameters, it is POSITIVE that, if it is a property, it will be the retrieval of it.
    'with parameters, however, is a different story. we'll need to try a Put first, then a get if an error occurs.
    
   ' If IsArray(Arguments) Then
        'An Array/
        'possibly Get/Put/PutRef/Function.
        'INVOKE_PROPERTYPUT and INVOKE_PROPERTYPUTREF require arguments.
        If MemberInf.InvokeKinds And INVOKE_PROPERTYPUT And IsArray(Arguments) Then
            
        'try Good ol' put first...
            TryEm(TryCount) = INVOKE_PROPERTYPUT
            TryCount = TryCount + 1
        
        End If
        If MemberInf.InvokeKinds And INVOKE_PROPERTYPUTREF And IsArray(Arguments) Then
            TryEm(TryCount) = INVOKE_PROPERTYPUTREF
            TryCount = TryCount + 1
        End If
        If MemberInf.InvokeKinds And INVOKE_PROPERTYGET Then
            TryEm(TryCount) = INVOKE_PROPERTYGET
            TryCount = TryCount + 1
        End If
        If MemberInf.InvokeKinds And INVOKE_FUNC Then
            TryEm(TryCount) = INVOKE_FUNC
            TryCount = TryCount + 1
        End If
 
    
    
    
    
    End If
    Dim CurrTry As Long
    
    'Final Stage- Iterate through each item in TryEm from 0 to trycount and attempt an InvokeHook.
    For CurrTry = 0 To TryCount
        On Error Resume Next
        If UBound(ArgsPass) < 0 Then
           Err.Clear
           Call Assign(retval, CallByName(onObj, Membername, TryEm(CurrTry)))
           
        Else
            Call Assign(retval, tli.InvokeHookArray(onObj, Membername, TryEm(CurrTry), ArgsPass))
            
        End If
        If Err = 0 Then Exit For
        Err.Clear
    Next
    
End Sub
Public Function FindSearchItem(SearchRes As SearchResults, ByVal StrName As String) As SearchItem
    Dim looper As SearchItem
    For Each looper In SearchRes
        If StrComp(looper.Name, StrName, vbTextCompare) = 0 Then
            Set FindSearchItem = looper
            Exit Function
        End If
    
    Next
    Set FindSearchItem = Nothing
End Function

Public Sub InvokeDynamic1(ByVal onObj As Object, ByVal Membername As String, Arguments As Variant, retval As Variant)
    'Dynamically invokes a method.
    'use InvokeHookArray of TLIApplication
    'first, we'll need to flip around the arguments.
    Dim ArgsPass() As Variant
    Dim AccessMode As InvokeKinds
    Dim Numerrs As Long
    Dim Errors(0 To 5)
    Dim I As Long, Current As Long
    Dim MemberInf As tli.MemberInfo
    On Error Resume Next
    'retrieve the memberinf object for future reference.

    Set MemberInf = tli.InterfaceInfoFromObject(onObj).GetMember(Membername)
    If Err <> 0 Then
        Err.Raise ExecuteErrors.Exec_UnsupportedOperation, "ModParser::InvokeDynamic", "Interface member name " & Membername & " Not present in class name " & TypeName(onObj)

    End If
    If IsArray(Arguments) Then
            ReDim ArgsPass(UBound(Arguments))

        For I = UBound(Arguments) To 0 Step -1
            If IsObject(Arguments(I)) Then
                Set ArgsPass(Current) = Arguments(I)
            Else
                ArgsPass(Current) = Arguments(I)
            End If
            Current = Current + 1
        Next I
    End If
    'arguments must be in reverse order.
    'OK, here goes the call.
    AccessMode = INVOKE_FUNC
        'default to a Function.
        On Error GoTo HandleBadMode
    Dim temphold As Variant
    If IsArray(Arguments) Then
'
'        If IsObject(TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())) Then
'            Set retval = TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())
'        Else
'            retval = TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())
'        End If

'the above was replaced with an assign call to fix the problem that we called the member twice.
      Assign retval, tli.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())

    Else
        'use CallByName- the constants are the same, IE- INVOKE_FUNC will be the same
        'as vbMethod, so I imagine the constants are interchangable.
        'Anyway- no parameters.
           ' Assign retval, CallByName(onObj, Membername, AccessMode)
           'fix-
            Assign retval, tli.InvokeHook(onObj, Membername, AccessMode)
    End If
    Err.Clear
Exit Sub
HandleBadMode:
'handlebad modes.
Errors(Numerrs) = Err.Number
Numerrs = Numerrs + 1
Select Case Err.Number
    'Don 't raise the error back if it is an OLE error, fired from TLI- or CallByName()
    Const ERR_NOAUTOMATION = 430
    Const ERR_PROPNOTFOUND = 422


    Case ERR_PROPNOTFOUND, 5, ERR_NOAUTOMATION, 435 To 447, Is < Round(TliErrors.tliErrArrayBoundsNotAvailable), -500

    Select Case Numerrs
        Case 1
            'TLI check here- is it read-only?
            If MemberInf.InvokeKind And INVOKE_PROPERTYPUT Then


            AccessMode = INVOKE_PROPERTYPUT
        Else
            Numerrs = 2
            AccessMode = INVOKE_PROPERTYGET
        End If
        Case 2
            AccessMode = INVOKE_PROPERTYGET
        Case 3
            AccessMode = INVOKE_PROPERTYPUTREF
        Case 4
            AccessMode = INVOKE_CONST
        Case 5
            AccessMode = INVOKE_UNKNOWN
        Case Else
            'raise the error.
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Err.Clear
        Resume
Case Else
    'real error.
    Err.Raise Err.Number
End Select

'whew.
End Sub
Public Function GetArrayElementVar(ArrGet As Variant, ElemSubscript As Variant) As Variant

Dim longsubscript() As Long, I As Long, Temp As Variant
If Not IsArray(ElemSubscript) Then
    
    ReDim Temp(0)
    Temp(0) = ElemSubscript
Else
    Temp = ElemSubscript
End If
ReDim longsubscript(UBound(Temp))
For I = 0 To UBound(Temp)
    longsubscript(I) = Temp(I)
Next
GetArrayElementVar = GetArrayElement(ArrGet, longsubscript())

End Function
Public Function GetArrayElement(ArrGet As Variant, ElemSubscript() As Long) As Variant
    'returns the item at the location given the subscripts.
    'I suppose that, right now, I'll need a large Select Case statement. sigh.
    'take the CHEAP way out.
    
    Dim createCode As String, I As Long
    Dim subscriptStr As String
    Dim ArrElem As Variant
    Dim ScriptLoop As Long
    'OK,- right now, BASeParser implements a Array as a Ragged variant array,
    'that is, each item can be an array, and each element of that array can be another, etc.
    
    'GetArrayElement = ArrGet(ElemSubScript(0))
    'initialize the Loop array-
    ArrElem = ArrGet
    For ScriptLoop = LBound(ElemSubscript) To UBound(ElemSubscript)
        'assign the loop array to the next successive element.
        'assume errors are handled elsewhere.
        ArrElem = ArrElem(ElemSubscript(ScriptLoop))
        'there we go.
    Next
    GetArrayElement = ArrElem
End Function
Public Function SetArrayElement(ArrSet As Variant, ElemSubscript() As Long, NewValue As Variant) As Variant
  '
  Call Assign(ArrSet(ElemSubscript(0)), NewValue)
End Function
    
    
    

Public Function compare(withparser As CParser, ByVal OpA As Variant, ByVal OpB As Variant, Optional ByVal FlDescending As Boolean = True) As Integer
    'Compares two items.
    'returns 1 if OpA should come before OpB
    'returns -1 if OpA should come after OpB.
    Dim Ret As Integer
'    If VarType(OpA) = vbString Or VarType(OpB) = vbString Then
        
            'only one is a string
            'If VarType(OpA) = vbString Then
'                On Error Resume Next
'                OpA = Val(OpA)
'                If Err <> 0 Then
'                    OpB = CStr(OpB)
'                    Err.Clear
'                End If
'            Else
'                On Error Resume Next
'                OpB = Val(OpB)
'                If Err <> 0 Then
'                    OpA = CStr(OpA)
'                    Err.Clear
'                End If
'            End If
            
'    End If
    
    
    
   ' If VarType(OpA) = VarType(OpB) Then
        Dim CompareRet As Variant
        'On Error Resume Next
        'Stop
        Call withparser.PerformOperation("=", OpA, OpB, CompareRet)
        If CompareRet Then
            Ret = 0
        Else
            Call withparser.PerformOperation("<", OpA, OpB, CompareRet)
            If CompareRet Then
                Ret = -1
            Else
                Ret = 1
            End If
        End If
        
        
        
'        If OpA < OpB Then
'            ret = -1
'        ElseIf OpA = OpB Then
'            ret = 0
'        Else
'            ret = 1
'        End If
   ' Else
        
   ' End If
        If Not FlDescending Then
            Ret = Ret * -1
        End If

    compare = Ret


End Function



Public Function GetArrayDimensionCount(ForArr As Variant) As Long
    'returns the number of dimensions.
    Dim I As Long
    Dim Curruse As Long
    Dim tempGrab As Long
    Curruse = 1
    On Error Resume Next
    Do
        
        tempGrab = UBound(ForArr, Curruse)
        If Err <> 0 Then Exit Do
        Curruse = Curruse + 1
        I = I + 1
    Loop
    GetArrayDimensionCount = I
End Function


Function ArcSin(ByVal radial As Variant) As Variant
On Error Resume Next
ArcSin = Atn(CDec(radial) / Sqr(-radial * radial + 1))
End Function
Function ArcCos(ByVal radial As Variant) As Variant
ArcCos = Atn(CDec(radial) / Sqr(-radial * radial + 1)) + 1.5708
End Function
Function Sec(ByVal radial As Variant) As Variant
Sec = 1 / Cos(CDec(radial))
End Function
Function Cosec(ByVal radial As Variant) As Variant
Cosec = 1 / Sin(radial)
End Function
Function Cotan(ByVal radial As Variant) As Variant
Cotan = 1 / Tan(radial)
End Function
Function ArcSec(ByVal radial As Variant) As Variant
ArcSec = Atn(radial / Sqr(-radial * radial + 1)) + Sgn(Sgn(radial) - 1) * 1.5708
End Function
Function ArcCoSec(ByVal radial As Variant) As Variant
ArcCoSec = Atn(radial / Sqr(radial * radial - 1)) + (Sgn(radial) - 1) * 1.5708
End Function
Function ArcCoTan(ByVal radial As Variant) As Variant
ArcCoTan = Atn(radial) * 1.5708
End Function
Function hSin(ByVal radial As Variant) As Variant
hSin = (Exp(radial) - Exp(-radial)) / 2
End Function
Function hCos(ByVal radial As Variant) As Variant
hCos = (Exp(radial) + Exp(-radial)) / 2
End Function
Function hTan(ByVal radial As Variant) As Variant
hTan = (Exp(radial) - Exp(-radial)) / (Exp(radial) + Exp(-radial))
End Function
Function HSec(ByVal radial As Variant) As Variant
HSec = 2 / (Exp(radial) + Exp(-radial))
End Function
Function HCoSec(ByVal radial As Variant) As Variant
HCoSec = 2 / (Exp(radial) - Exp(-radial))
End Function
Function HCoTan(ByVal radial As Variant) As Variant
HCoTan = (Exp(radial) + Exp(-radial)) / (Exp(radial) - Exp(-radial))
End Function
Function HArcSin(ByVal radial As Variant) As Variant
HArcSin = Log(radial + Sqr(radial * radial + 1))
End Function

Function HArcCos(ByVal radial As Variant) As Variant
HArcCos = Log(radial + Sqr(radial * radial - 1))
End Function
Function HArcTan(ByVal radial As Variant) As Variant
HArcTan = Log(1 + radial) / (1 - radial) / 2
End Function
Function HArcSec(ByVal radial As Variant) As Variant
HArcSec = Log((Sqr(-radial * radial + 1) = 1) / radial)
End Function
Function HArcCoSec(ByVal radial As Variant) As Variant
HArcCoSec = Log((Sgn(radial) * Sqr(radial * radial + 1) + 1) / radial)
End Function
Function HArcCoTan(ByVal radial As Variant) As Variant
HArcCoTan = Log(radial + 1) / (radial - 1) / 2
End Function
Function Sieve(ai As Variant) As Long
    Dim iLast As Integer, cPrime As Integer, iCur As Integer, I As Integer
    Dim af() As Boolean
    ' Parameter should have dynamic array for maximum number of primes
    If LBound(ai) <> 0 Then Exit Function
    iLast = UBound(ai)
    ' Create array large enough for maximum prime (initializing to zero)
    ReDim af(0 To iLast + 1) As Boolean
    For iCur = 2 To iLast
        ' Anything still zero is a prime
        If Not af(iCur) Then
            ' Cancel its multiples because they can't be prime
            For I = iCur + iCur To iLast Step iCur
                af(I) = True
            Next
            ' Count this prime
            ai(cPrime) = iCur
            cPrime = cPrime + 1
        End If
    Next
    ' Resize array to the number of primes found
    ReDim Preserve ai(0 To cPrime - 1) As Variant
    Sieve = cPrime
End Function
Public Function Factorial(ByVal n As Double) As Variant

    
    Dim Accum As Variant, CurrNum As Double
    '0! = 1. negative numbers are invalid.
    If n < 0 Then
        'bad.
        Err.Raise 5, "Factorial", "Cannot take factorial of " & n & " the operation is undefined."
    
    End If
    If n = 0 Then
        Factorial = 1
        Exit Function
    ElseIf n < 0 Then
       'return "Null"
        Factorial = Null
    End If
    'could use recursive algorithm, but- why waste stack space?
    Accum = CDec(n)
    For CurrNum = n - 1 To 1 Step -1
        Accum = Accum * CDec(CurrNum)
    Next
    Factorial = Accum
    



End Function
Public Function GetFormitemCount(OfItem As CFormItem) As Long
    Dim I As Long, useit As CFormItem
    Set useit = OfItem
    Do Until useit Is Nothing
        I = I + 1
        Set useit = useit.Next_

    Loop
    GetFormitemCount = I
End Function
Public Function Comb(ByVal Number As Variant, ByVal number_chosen As Variant) As Variant

 'Number of combinations-
 'The number of combinations is as follows, where number = n and number_chosen = k:
        '7 nCr 2 = 21
        'number of unique combinations
        'of 7 items into 2 positions.
        'result = (n!)/(k!*(n-k)!)
    Dim K, n, Result
    K = number_chosen
    n = Number
    Result = Factorial(n) / (Factorial(K) * Factorial(n - K))
    Comb = Result

End Function
Public Function ShiftleftStr(ByVal str As String, Numchars As Long) As String
Dim Ret As String
Ret = str
Str_Shl Ret, Numchars
ShiftleftStr = Ret
End Function
Public Function ShiftRightStr(ByVal str As String, Numchars As Long) As String
Dim Ret As String
Ret = str
Str_Shr Ret, Numchars
ShiftRightStr = Ret
End Function
Public Sub Str_Shl(ByRef StrShift As String, Optional ByVal Numchars As Long = 1)
    'String Shift Left.
    
    
    Dim pivot As Long
    Dim beforepivot As String, afterpivot As String
    If Numchars > Len(StrShift) Then Numchars = (Numchars Mod Len(StrShift))
    
    pivot = Numchars
    beforepivot = Mid$(StrShift, 1, pivot)
    afterpivot = Mid$(StrShift, pivot)
    StrShift = afterpivot & beforepivot






End Sub
Public Sub Str_Shr(ByRef StrShift As String, Optional ByVal Numchars As Long = 1)
'shift right.

    Dim pivot As Long
    Dim beforepivot As String, afterpivot As String
    If Numchars > Len(StrShift) Then Numchars = (Numchars Mod Len(StrShift))
    pivot = Len(StrShift) - Numchars
    beforepivot = Mid$(StrShift, 1, pivot)
    afterpivot = Mid$(StrShift, pivot)
    StrShift = afterpivot & beforepivot
End Sub

Public Function Permut(ByVal Number As Variant, ByVal number_chosen As Variant) As Variant

Dim K, n, Result
K = number_chosen
n = Number
'n!/(n-k)!
Result = Factorial(n) / Factorial(n - K)
Permut = Result
End Function
Public Function PopupMenuEx(MenuItems As Variant, Optional ByVal ReturnSelArray As Boolean = True) As Variant
    'Pops up a menu. the given Array should be a Variant Array of
    'Strings, or Other arrays, for example:
    
    
    'Array(Array("File","New","Open","Print"),"Help")
    'would create a File menu popup with New,Open and print, and a single item, Help.

    'right now I use a command-bar control from VBaccelerator.
    'If the need arises, though, I should use the CreatePopupMenu and InsertMenu stuff to create the menu.
    'here it is easy- I just generate the XML string and pass it on.
    


'
'    Dim hPopup As Long
'
'    Dim MousePos As POINTAPI
'    'create the menu
'    hPopup = CreatePopupMenu
'
'
'
'
'    DestroyMenu hPopup


End Function
'Public Function Dec2Bin(ByVal DecVal As Variant) As String
'    'create a binary representation of the given value.
'
'End Function
'Public Function Bin2Dec(ByVal BinVal As String) As Variant
'Dim ValCreate As Variant
'Dim lookStr As String
'Dim lookPos As Long
''reverse it. Not necessary- but makes it easier.
'BinVal = StrReverse(BinVal)
'For lookPos = 1 To Len(BinVal)
'    If Mid$(BinVal, lookPos, 1) = "1" Then
'        ValCreate = ValCreate + (2 ^ (lookPos - 1))
'    End If
'Next
'Bin2Dec = ValCreate
'
'
'
'
'
'
'
'
'End Function




'Functions for Parsing the XML data for Descriptions and such.

Public Function GetFunctionInfoFromXML(withparser As CParser, ByVal funcName As String, FuncInfo As DOMDocument) As FUNCTIONINFORMATION
Dim finfo As FUNCTIONINFORMATION

 Dim funcnode As MSXML2.IXMLDOMNode
    funcName = UCase$(funcName)
    finfo.StrFunctionName = StrConv(funcName, vbProperCase)
    'finfo.StrDescription
    'finfo.ParameterInfo
    'finfo.StrHelpHTML
  'return information on the requested function.
    'this function will be even bigger then the parsing one...
    'unless I store all the relevant data as resources....
    'The data for GetFunctionInformation will be stored
    'as a custom XML resource. the XML will be:
    '<FUNCTION NAME="Name" Description="Description" ParameterInfo="ParameterInfo" HELPRESOURCEID="Helpresourceid"></FUNCTION>
    'HELPRESOURCEID is used to load a resource of Type FuncHelp.
    '
    Set funcnode = FindFunctionNode(FuncInfo, funcName)
    With finfo
        On Error Resume Next
        .StrFunctionName = funcName
        .StrDescription = funcnode.Attributes.getNamedItem("DESCRIPTION").Text
        
        If Err <> 0 Then
            CDebug.Post "XML file has no DESCRIPTION attribute for function " & funcName & "!", Severity_Warning
        End If
      
        .ParameterInfo = CreateParameters(funcnode)
        On Error Resume Next
        '.ReturnType = funcnode.Attributes.getNamedItem("RETURNTYPE").Text
        '.ReturnDescription = funcnode.Attributes.getNamedItem("RETURNS").Text
        '<RETURNS> tag will contain the above info.
        .StrHelpHTML = funcnode.Attributes.getNamedItem("HELPRESOURCEID").Text
        Err.Clear
        
        
        .ParameterCount = UBound(.ParameterInfo) + 1
        If Err <> 0 Then .ParameterCount = 0
        
    
    End With

    GetFunctionInfoFromXML = finfo

End Function
Private Function CreateParameters(funcnode As IXMLDOMNode) As PARAMETERINFORMATION()
    Dim pinfo() As PARAMETERINFORMATION
    Dim CurrParm As IXMLDOMNode
    Dim I As Long
    'parameter structure-
    'has children, "PARAMETER"
    '
    '<PARAMETER NAME="NAME" TYPE="TYPE" ISOPTIONAL="[T|F]" DESCRIPTION="DESC" HANDLETYPE="[BYREF|BYVAL]"></PARAMETER>
    'Name As String
    'DataType As String
    'isOptional As Boolean
    'Description as String
    ReDim pinfo(0 To funcnode.childNodes.Length - 1)
    Set CurrParm = funcnode.firstChild
    Do Until CurrParm Is Nothing
        With pinfo(I)
        On Error Resume Next
            .Name = CurrParm.Attributes.getNamedItem("NAME").Text
            .DataType = CurrParm.Attributes.getNamedItem("TYPE").Text
            .isOptional = CBool(CurrParm.Attributes.getNamedItem("ISOPTIONAL").Text)
            .Description = CurrParm.Attributes.getNamedItem("DESCRIPTION").Text
            Select Case UCase$(CurrParm.Attributes.getNamedItem("HANDLETYPE"))
                Case "BYREF", "REFERENCE"
                    .HandleType = Parameter_ByRef
                Case "BYVAL", "VALUE"
                    .HandleType = Parameter_ByVal
            End Select
            
        End With
        I = I + 1
        Set CurrParm = CurrParm.nextSibling
    Loop
            
    
    
    
    
    
    
    
    
    
    CreateParameters = pinfo
End Function

Private Function FindFunctionNode(FuncInfo As MSXML2.DOMDocument30, ByVal vfuncname As String) As MSXML2.IXMLDOMNode
Dim CurrNode As MSXML2.IXMLDOMNode
Dim CurrFunc As MSXML2.IXMLDOMNode
    With FuncInfo
        Set CurrNode = .firstChild
        Do Until CurrNode Is Nothing
            If CurrNode.baseName = "FUNCTIONINFO" Then
                'Debug.Assert False
                Set CurrFunc = CurrNode.firstChild
                Do Until CurrFunc Is Nothing
                    If StrComp(CurrFunc.baseName, "FUNCTION", vbTextCompare) = 0 Then
                        'Debug.Print CurrFunc.Attributes.getNamedItem("NAME").Text
                        If StrComp(CurrFunc.Attributes.getNamedItem("NAME").Text, vfuncname, vbTextCompare) = 0 Then
                            'Debug.Assert False
                            'found it.
                            Set FindFunctionNode = CurrFunc
                            Exit Function
                        End If
                        
                    End If
                    Set CurrFunc = CurrFunc.nextSibling
                Loop
            End If
            Set CurrNode = CurrNode.nextSibling
        Loop
    
    End With





End Function

'Public Function CBool(Value as Variant) as Boolean
Public Function Seq(withparser As CParser, ByVal Expression As String, ByVal VarName As String, ByVal Start As Variant, ByVal EndVal As Variant, Optional ByVal StepValue As Variant = 1) As Variant
'finally made this a separate function.
Dim SeqStart, SeqEnd, SeqStep
Dim varGrab As CVariable
Dim parsetemp As CParser, I As Long


            Dim SeqLoop
            'If paramcount < 3 Then GoTo NotOptional
            
            SeqStart = Start
            SeqEnd = EndVal
            'Step is optional.-
            SeqStep = StepValue
            If SeqStep = 0 Then SeqStep = 1
            'sequence.
            'SEQ(Expression,Varname,Start,End,[Step])
            'Performs <expression> for values of <varname> from <start> to <end>
            'in increments of <Step>
            'NOTE: Step can not be zero (duh)
            'if step is zero, it is set to one.
            Set varGrab = withparser.Variables.Item(VarName)
            Set parsetemp = withparser.Clone
            
            'IMPORTANT! Always define variables before having the parser parse functions with those variables....
            'trust me, it is VERY IMPORTANT!
            parsetemp.Expression = Expression
            
            I = 0
            ReDim arrtemp(0)
            For SeqLoop = SeqStart To SeqEnd Step SeqStep
                'simple! assign the variable, evaluate, and create an array.
                varGrab.Value = SeqLoop
                ReDim Preserve arrtemp(I)
                parsetemp.ExecuteByRef arrtemp(I)
                I = I + 1
            Next
            Seq = arrtemp
End Function

Public Function FilterArray(withparser As CParser, ArrayFilter As Variant, ByVal VarName As String, ByVal ComparisonExpr As String) As Variant
    'iterates through all items in ArrayFilter, only returning thise items who, when substituted in the ComparisonExpr expression, return true.
    Dim FilterParser As CParser
    Dim currindex As Long
    Dim returnArray As Variant, IterateVar As CVariable
    Dim parserreturn As Variant, CurrCount As Long
    If Not IsArray(ArrayFilter) Then
        ParserError(withparser, ExecuteErrors.Exec_UnsupportedOperation, "FilterArray second argument must be an Array").Throw
    
    Else
        Set FilterParser = withparser.Clone
        If FilterParser.Variables.Exists(VarName) Then
            FilterParser.Variables.Remove VarName
        End If
        Set IterateVar = FilterParser.Variables.Add(VarName, 0)
        FilterParser.Expression = ComparisonExpr
        ReDim returnArray(0 To UBound(ArrayFilter) - LBound(ArrayFilter) + 1)
        For currindex = LBound(ArrayFilter) To UBound(ArrayFilter)
            If IsObject(ArrayFilter(currindex)) Then Set IterateVar.Value = ArrayFilter(currindex) Else IterateVar.Value = ArrayFilter(currindex)
            FilterParser.ExecuteByRef parserreturn
            If CBool(parserreturn) Then
                Assign returnArray(CurrCount), ArrayFilter(currindex)
                CurrCount = CurrCount + 1
                
                

            End If
        Next
        ReDim Preserve returnArray(CurrCount - 1)
        FilterArray = returnArray
    End If
    


End Function

'Public Function SeqEx(Byval Expression as String,)
Public Function SeqEx(withparser As CParser, ByVal Expression As String, ByVal VarName As String, _
    InitialValue As Variant, ByVal IncrementExpression As String, ByVal TerminateExpression As String, Optional ByVal Iterationlimit As Long = -1) As Variant
    
            'Sequence- Extended!'
            'Expression , Varname, varstart,InitialValue, IncrementExpression, TerminateExpression
            
            'Expression to be performed each iteration.
            'varName: Name of variable in Expression.
            'varStart: Initial Value of <VarName>
            'IncrementExpression:used each iteration to determine the next value of <Varname>
            'TerminateExpression:if this returns non-zero, then the loop will be terminated. tested every iteration.
            

    Dim IncrementParse As CParser
            Dim TerminateParse As CParser
            Dim ExprParse As CParser, Prev As Variant
            Dim TermTest As Boolean
            Dim tempvar As Variant
            Dim variter As CVariable
            Dim varGrab As CVariable
            Dim arrtemp As Variant, I As Long
            Dim PrevQueue As DataStack
            Const KEPT_PREV_ITEMS = 3
            ReDim Prev(0)
            Set IncrementParse = withparser.Clone
            Set TerminateParse = withparser.Clone
            Set ExprParse = withparser.Clone
            'assign the expressions...
            ExprParse.Expression = Expression
            IncrementParse.Expression = IncrementExpression
            TerminateParse.Expression = TerminateExpression
            
            'DIE:
            'if the terminateParse parser is constant.
            If TerminateParse.IsConstant And TerminateParse.Execute = False Then
                Err.Raise 5, "SeqEx", "Termination Expression Must not be Constant False."
            End If
            'also check that IncrementParse is not constant.
            If IncrementParse.IsConstant Then
                Err.Raise 5, "SeqEx", "Increment Expression must not be a constant value."
            
            End If
            'create <Varname>
            'eventually create a "prev()" array of previous items.
            'where 1 is the previous item, 2 is the next previous item, and so on.
            
            Set varGrab = ExprParse.Variables.Add(VarName, InitialValue)
            ReDim Prev(0)
            Prev(0) = varGrab.Value
            ReDim arrtemp(0)
            'OK, time to start the loop.
            'Prev(0) is always the current value of the variable.
            
            TermTest = False
            I = 0
            Set variter = ExprParse.Variables.Item("N")
            Do Until TermTest   'loop until the condition they gave is non-zero(true)
                'Redimension the array to hold the new item.
                variter.Value = I
                ReDim Preserve arrtemp(I)
                'now, get this value.
                ExprParse.ExecuteByRef arrtemp(I)
                'execute the increment code...
                IncrementParse.ExecuteByRef tempvar
                'tempvar is the new value of variter.
                
                I = I + 1
                'variter.Value = tempvar
                varGrab.Value = tempvar
                TerminateParse.ExecuteByRef TermTest
                If I > Iterationlimit And Iterationlimit > 0 Then
                    Exit Do
                End If
                
                
                'TODO:// Add PrevItem support. last item stored in prev(1), next last Prev(2), etc...
            Loop
            SeqEx = arrtemp
End Function



Public Function GenerateGUID() As String
    Dim GUIDuse As IShellFolderEx_TLB.Guid
    Dim retval As String, lngret As Long
    CoCreateGuid GUIDuse
        Call StringFromCLSID(GUIDuse, lngret)
        StringFromPointer lngret, retval

    GenerateGUID = retval
End Function
Public Sub StringFromPointer(pOLESTR As Long, strOut As String)
         Dim ByteArray(255) As Byte
         Dim intTemp As Integer
         Dim intCount As Integer
         Dim I As Integer

         intTemp = 1

         'Walk the string and retrieve the first byte of each WORD.
         While intTemp <> 0
            CopyMemory intTemp, ByVal pOLESTR + I, 2
            ByteArray(intCount) = intTemp

            intCount = intCount + 1
            I = I + 2
         Wend

         'Copy the byte array to our string.
         strOut = Space$(255)
         CopyMemory ByVal strOut, ByteArray(0), intCount
      End Sub
Public Function Ceil(ByVal n As Variant)
'returns the largest integer greater then or equal to N.
'we need to perform the reverse of the "Int" function. So first,
'we do a "clever" transform of the float part:


   Dim Temp As Double
     Temp = Int(n)
     Ceil = (Temp + IIf(n = Temp, 0, 1))
   




End Function
Public Function GetObjectProgID(ObjFrom As Object)
    Dim retstr As String
    Dim IInfo As InterfaceInfo
    CDebug.Post "retrieving object progID for " & TypeName(ObjFrom)
    
    Set IInfo = InterfaceInfoFromObject(ObjFrom)
    
    retstr = IInfo.Parent.Name & "."
    retstr = retstr & TypeName(ObjFrom)
    GetObjectProgID = retstr
    CDebug.Post "returned ProgID is " & retstr
    Set IInfo = Nothing
End Function
Public Function SelectItems(withparser As CParser, FromArray As Variant, Number As Variant) As Variant
'Prob: feb 02 2007; this procedure causing VB to crash with an IPF (?)
Dim OpA, OpB, Temp As Variant, I As Long
OpA = FromArray
OpB = Number
    If Not IsArray(OpB) Then
                'copy the value of OpB...
                Temp = OpB
                'now, redimension OPB as an array.
                ReDim OpB(0)
                'copy it back.
                OpB(0) = Temp
            End If
            'we are positive that OPB is an array.
            'however, what if opA isn't? Well, by all accounts, we should raise a type mismatch error:
            If Not IsArray(OpA) Then
                Err.Raise 13
                'why no description? Well, because above us in the call stack is a very versatile
                'error handler- in CParser.CollapseStack. It will raise the error AGAIN, but with descriptive text.
                'in this case, it will say a type mismatch occured in Op Select.
            End If
            'OK, OpA is an Array, and So Is OPB.
            ReDim Ret(UBound(OpB))
            
            For I = 0 To UBound(OpB)
                'take OpB(I) number of elements from OPA, and add
                'that result to the return array at I.
                Temp = Pick(withparser, OpA, OpB(I))
                'put this result in ret.
                Ret(I) = Temp
            Next I
                'if the "ret" array has one item, coerce it to the element.
                If UBound(Ret) = LBound(Ret) Then
                'Found the IPF. the IPF occured here.
                'it works now, but before these three statements where non-existent, I instead had
                'ret = ret(ret(0))
                'which caused VB to ipf.
                'shouldn't, but it kind of makes sense. assigning a value to a element of that value using an index also derived from that value?
                'gotta have a problem there...
                'UPDATE: weird. I decided to confirm this problem, and I cannot reproduce it in a different project.
                'I guess it had something to do with a class I was using that subclassed, even though it had nothing TO subclass.
                'Oh well- it works now, no worries, right?
                    Temp = LBound(Ret)
                    Temp = Ret(Temp)
                    Ret = Temp
                End If
                    
            SelectItems = Ret
End Function

Public Function Pick(withparser As CParser, FromArray As Variant, Number As Variant) As Variant
    'picks Number random items from "FromArray" and returns them.
    'important: don't allow duplicates.
    'as such, Number cannot be larger then the number of items in FromArray.
    Dim numitems As Long
    Dim currItem As Long
    Dim randIndex As Long
    Dim Ret As Variant
    Dim SelIndices As Variant   'variant, for use with "among"
    numitems = (UBound(FromArray) - LBound(FromArray)) + 1
    If Number > numitems Then
        Err.Raise 13, "ModParser::Pick", "Cannot select " & Number & " unique items from a set of " & numitems & " values."
    
    End If
    ReDim Ret(0 To Number - 1)
    ReDim SelIndices(0 To Number - 1)
    'initialize to a value that is unlikely to be  present in the FromArray-
    'vbnullchar.
    For currItem = 0 To Number - 1
        Ret(currItem) = vbNullChar
        SelIndices(currItem) = LBound(FromArray) - 1
    Next
    For currItem = 0 To Number - 1
        'select a random index.
        
        Do
        randIndex = Random(LBound(FromArray), UBound(FromArray), False)
            
        Loop Until Among(withparser, randIndex, SelIndices) = False 'keep looping until valid.
        Call Assign(Ret(currItem), FromArray(randIndex))
        Call Assign(SelIndices(currItem), randIndex)
    Next currItem
    
    
    
    Pick = Ret
    
    
    
    



End Function
Public Function e() As Variant
    e = Exp(1)
End Function
Public Function Complex(ByVal Realpart, Optional ByVal ImagPart = 0) As CComplex
    Dim retcomplex  As CComplex
    If IsObject(Realpart) Then
        If TypeOf Realpart Is CComplex Then
            Set Complex = Realpart
            Exit Function
        End If
    End If
    Set retcomplex = New CComplex
    retcomplex.Realpart = Realpart
    retcomplex.ImagPart = ImagPart
    Set Complex = retcomplex
End Function
Public Function SplitArray(ArrSplit, Optional ByVal FirstItem As Long = -1, Optional ByVal ChunkSize As Long = 1, Optional ByVal NumChunks As Long = -1) As Variant
    'SplitArray:
    'Splits an array into a bunch of smaller arrays.
    Dim retArr As Variant, AddItem As Variant
    
    If FirstItem = -1 Then FirstItem = LBound(ArrSplit)
    If ChunkSize <= 0 Then Err.Raise 5
    If NumChunks <= 0 Then
    NumChunks = (UBound(ArrSplit) - FirstItem) Mod ChunkSize
    
    End If




End Function
Public Function CheckEasterEgg(ByRef Vdata As Variant) As Boolean
    CheckEasterEgg = False
'    Exit Function
    'CheckEasterEgg.
    'checks if the given "Expression" is actually a secret easter egg entry.
    'returns true if so, and changes the Vdata argument into the resulting egg.
    '(this will end up parsed- so don't forget the quotes).
    'OK, so MAYBE- just, MAYBE
    If IsObject(Vdata) Then
        'this one requires programming skills.
        'so it might not be "found" for a VERRY long time, if I don't
        'release source code.
        'also, it requires that they actually assign a  Variant to the expression property.
        'I like it- it's sneaky, they need to make a class module named, of all things, "Turkey"
        'Not, however "CTurkey"
        If StrComp(TypeName(Vdata), "Turkey") = 0 Then
            Vdata = """Mmmm, Turkey :)"""

        End If

    End If
    Dim VCopy As String, original As Variant
    'first easter egg is kind of easy. well, not counting the password.
    If left$(Vdata, 12) = "SECRETPLACE" Then
        Dim pwlTest As String
        
        pwlTest = InputBox$("RESTRICTED ACCESS. ENTER PASSCODE:", "DevOptions", "")
        pwlTest = Trim$(pwlTest)
        If GetHash(pwlTest) = PwlHash Then
            MsgBox "Fatal Error 0xEE47: Nothing happened."
            MsgBox "Due to fatal error, uh 0xEE-something, nothing has occured. this is bad. But don't do anything, it'll make it worse."
            MsgBox "I said don't do anything, Not to click OK! this is terrible! Oh, wait, something happened. good, error 0x something was nothing to worry about."
            Dim strvar As String
REGETINPUT:
            strvar = InputBox$("Because An Error has occured, it is important that we get as much information about you so that we can pretend to do something about the problem. Please enter your name:", "Name Entry")
            Select Case UCase$(strvar)
                Case "FUCK YOU", "NO", ""
                    MsgBox "Sir, SIR! settle down! OK, we're sorry that error 0xEE47 Occured, but we need this information to help you! It would be in your best interest to cooperate."
                    GoTo REGETINPUT
                Case "I DON'T WANT TO HELP"
                    MsgBox "Why not! If you don't want help you may as well go screw a chicken in a darkroom closet, and then develop twelve photos of puppies wearing adult diapers."
                Case "I'M STUPID"
                    MsgBox "The first step is admitting you have a problem."
                Case "MY HEAD HURTS"
                    MsgBox "If you would stop using your forehead to type, it wouldn't hurt so much, you damn neanderthal."
                Case "I HAVE A PET WEASEL"
                    MsgBox "And I don't. That's wonderful."
                Case "AL GORE IS COOL"
                    MsgBox "Wow, I haven't seen someone this stupid for a long time."
                Case "WAFFLE-IRON"
                    MsgBox "Oh, Man! I LOVE waffles. I want some. shove some in my orifices. (I mean the drives, you sick perv.)"
            End Select
            MsgBox "This Application has violated the system sexually and will now be forced to marry it."
            MsgBox "Oh, man! I need to start dinner!"
            MsgBox "Sorry, user. I forgot about ya! Anywho, We here in the system, we apologize, we wanted to give you this earlier, but, golly, we're sorry. Click OK to see our present."
            MsgBox "Error 70: Permission denied."
            MsgBox "Woops! Sorry about that."
            MsgBox "I promise I'll make it up to ya ;)"
            MsgBox "A-BOOP-DITTY-DITTY-A-BOOP-DITTY-DOO"
            MsgBox "Thank you for your cooperation, " & strvar & " it is appreciated."
            Vdata = """GOURANGA!"""
        Else
            Vdata = """Access Denied."""
        End If
        
    'second easter egg-
    ElseIf Trim$(Vdata) = "BASeCamp©" Then
        'ha! how often do YOU randomly paste (r) into random apps?
        'this loads the resource file and displays it in whatever the default text file viewer is.
        'That is one reason I made it less than Notepad's limit.
        
    
    End If
    If VarType(Vdata) = vbString Then
        VCopy = Vdata
        Vdata = Trim$(UCase$(Vdata))
        original = Vdata
        Select Case True
            Case Vdata = "I LIKE TURKEY"
                Vdata = """Wow, we have SOOO much in common, since I like turkey TOO! :)"""
            Case Vdata = "I DON'T LIKE TURKEY"
                Vdata = """SURE, liar- everybody LOVES turkey."" Except those damn mexicans."
            Case Vdata = "FIREFOX ROX", Vdata = "FIREFOX ROCKS", Vdata = "FIREFOX IS GOOD", Vdata = "I LIKE FIREFOX"
                Vdata = """" & Replace$(Vdata, "Firefox", "Internet Explorer", , , vbTextCompare) & """"
            Case Vdata = "WHATS WRONG WITH FIREFOX?"
                Vdata = """I'll tell you what's wrong- the only reason people use it is because it isn't made by microsoft, regardless of the fact that is is a inferior piece of software. OK, thats a bit harsh- Oh, and it uses WAY more memory then Internet Explorer."
            Case Vdata = "I HATE THIS PARSER"
                Vdata = """What did I do to you? <sniff>."""
            Case Vdata = "WTF?"
                Vdata = """I see you are familiar with Abraham Lincolns favorite words."
            Case left$(Vdata, 17) = "THE MILLENIUM BUG"
                Vdata = """No, the CENTURY Bug- it would have occured, if computers existed during the rollover from 1800 to 1900."""
            'Date specific.
            Case Month(Now) = 4 And Day(Now) = 4        'My birthday. Gotta say somethin!
'                If Not ParserSettings.DisplayedDayEgg(Now) Then
'                    Vdata = """Happy birthday to me."""
'                End If
'            Case Month(Now) = 12 And Day(Now) = 25
'                If Not ParserSettings.DisplayedDayEgg(Now) Then
'                    Vdata = """Merry Christmas."""
'                End If
            Case Month(Now) = 3 And Day(Now) = 17
'                If Not ParserSettings.DisplayedDayEgg(Now) Then
'                    Vdata = """Happy St. Patricks day. Also- Happy birthday Davis."""
'
'                End If
            Case Year(Now) > 2050
                Vdata = """BASeParser is probably too old for you. You're probably laughing right now, at the ridiculous text and GUI, rather then the DBI (Direct Brain Interface). Well, I CAN'T CODE AGAINST SOMETHING THAT DOESN'T EXIST!"
            Case Year(Now) < 2000
                Vdata = """BEWARE OF Y2K! YOU WILL ALL DIE! PLANES WILL FALL FROM THE SKY! TOASTERS WILL GROW LEGS AND SHOOT LASER BEAMS! On the up side, the Toasters will actually work for once. Don't you hate when they undertoast it, and then you turn it up, and then it comes out black!? I mean, what the heck?"
        End Select
        If Vdata = original Then
            Vdata = VCopy
            CheckEasterEgg = False
        Else
            
            CheckEasterEgg = True
        End If
    End If
End Function
Private Function CompareDirect(VarA, VarB, Optional ByVal FlDescending As Boolean = False) As Integer
    Dim tempComplex As CComplex, Ret As Integer
    If IsObject(VarA) Or IsObject(VarB) Then
        If TypeOf VarA Is CComplex Then
            
                Set tempComplex = VarA
                Ret = tempComplex.compare(VarB, FlDescending)
        ElseIf TypeOf VarB Is CComplex Then
            Set tempComplex = VarB
                Ret = tempComplex.compare(VarB, FlDescending)
        Else
            Ret = VarA Is VarB * FlDescending
        
        
        End If
    Else
        If VarType(VarA) = vbString Xor VarType(VarB) = vbString Then
            Ret = StrComp(VarA, VarB)
        Else
            Ret = IIf(VarA > VarB, 1, -1)
            If VarA = VarB Then Ret = 0
        End If
        If FlDescending Then Ret = Ret * -1
            
    
    End If
    CompareDirect = Ret


End Function
Public Function CompareEx(VarA, VarB, Optional ByVal op As String = "<=>") As Integer
    'compares items with given op.
    'NOT a parser op-
    'can be:
    '=,==,<=,=>,>,<,<>,!=, or the Perl "spaceship" <=>(for ordering).
    Select Case op
        Case "=", "=="
            CompareEx = CompareDirect(VarA, VarB) = 0
        Case "<=", "=<"
            CompareEx = CompareDirect(VarA, VarB, True) >= 0
        Case ">=", "=>"
            CompareEx = CompareDirect(VarA, VarB, True) <= 0
        Case ">"
            CompareEx = CompareDirect(VarA, VarB, True) <= -1
        Case "<"
            CompareEx = CompareDirect(VarA, VarB, True) >= 1
        Case "<>", "!="
            CompareEx = CompareDirect(VarA, VarB, True) <> 0
        Case "<=>"
            CompareEx = CompareDirect(VarA, VarB, True)
    End Select
    
    
    




End Function
Public Sub PerformOperation(OpA As Variant, OpB As Variant, ByVal op As String, ByRef returnvalue As Variant, Optional withparser As CParser = Nothing)
Static ParserUse As CParser
Static PrevWith As CParser
If withparser Is Nothing Then
Set ParserUse = New CParser
ElseIf Not (PrevWith Is withparser) Then
Set ParserUse = withparser.Clone()
End If


'define two vars- OpitA, and OPITB.
'ParserUse.Variables.Add "OPITA", OpA
'ParserUse.Variables.Add "OPITB", OpB
''the expression uses the op.
'ParserUse.Expression = "OPITA " & op & " OPITB"
'ParserUse.ExecuteByRef returnValue
Call ParserUse.EvalListener.Self.GetOperation(ParserUse, op, OpA, OpB, Nothing, returnvalue)





End Sub
Public Function PwlHash() As String


'lowercase.
PwlHash = Chr$(&HA3) & Chr$(&HD7) & Chr$(&HD5) & Chr$(&H5D) & Chr$(&HE0) & Chr$(&H97) & Chr$(&H5D) & Chr$(&H51) & Chr$(&HDF) & Chr$(&HD7) & Chr$(&H91) & Chr$(&H97) & Chr$(&H5D) & Chr$(&H51) & Chr$(&HCB) & Chr$(&H70) & Chr$(&H91) & Chr$(&H97) & Chr$(&H8E) & Chr$(&H5D) & Chr$(&H51) & Chr$(&H8E) & Chr$(&H6F) '& Chr$(&HD) & Chr$(&HA)
End Function
Public Function ExecuteExpression(ByVal Expression As String) As String
    Static mParser As CParser, retval As Variant
    If mParser Is Nothing Then Set mParser = New CParser
    mParser.Create "Default"
    mParser.Expression = Expression
    mParser.ExecuteByRef retval
    ExecuteExpression = mParser.ResultToString(retval)

End Function
Public Function ExecExprStr(ByVal Expression As String, withparser As CParser, ByRef retval As Variant, ParamArray Varnamevalue() As Variant) As String
    Static mParser As CParser
    Dim LoopVar As Long
    If Not mParser Is Nothing Then
        'must be same parent.
        If Not mParser.Variables Is withparser.Variables Then
            'different parse trees, or whatever you want to call it.
            Set mParser = Nothing
    
        End If
    End If
    If withparser Is Nothing Then Set withparser = New CParser
    If mParser Is Nothing Then
        Set mParser = withparser.Clone
        
    End If
    'add the paramarray variables.
    For LoopVar = 0 To UBound(Varnamevalue) Step 2
        mParser.Variables.Add Varnamevalue(LoopVar), Varnamevalue(LoopVar + 1)
    Next
    'assign the expression-
    mParser.Expression = Expression
    'and EXECUTE!
    mParser.ExecuteByRef retval
    
        
        





End Function




'Public Function Sprintf(ByVal StrFormat As String, ParamArray Args() As _
'    Variant) As String
'    'sPrintf- from C to VB.
'
'
'    Dim StrBuild As cStringBuilder
'    Set StrBuild = New cStringBuilder
'    Dim CurrPos As Long, CurrChar As String
'    Dim Arguse As Long, appendit As String
'    Dim Width As Variant, Precision As Variant
'    Dim widthuse, precuse As Variant
'    Dim strtype As String   'type not used- retrieved for compat.
'    CurrPos = 1
'    Do Until CurrPos > Len(StrFormat)
'        CurrChar = Mid$(StrFormat, CurrPos, 1)
'        If CurrChar = "%" Then
'            'format specification- the next char
'            'needs to mean something, otherwise we'll simply
'            'ignore this percent sign, and append whatever the next character is.
'            '- can't be last char in string.
'            If Len(StrFormat) > CurrPos Then
'
'
'            'grab the parameters- or- attempt to, anyway :)
'            'ignore the very next char.
'            Width = Mid$(StrFormat, CurrPos + 2)
'            Precision = Mid$(StrFormat, CurrPos + 2 + Len(Width))
'            'get "Type" here.
'            If left$(Precision, 1) = "*" Then
'                Precision = "*"
'            ElseIf IsNumeric(Precision) Then
'                        precuse = CDbl(Precision)
'            Else
'                precuse = -1
'            End If
'
'
'            If IsNumeric(Width) Then
'                widthuse = Int(Width)
'            Else
'                widthuse = -1
'                Width = ""
'            End If
'            If Precision = "*" Then
'
'                precuse = Val(Args(Arguse))
'                Arguse = Arguse + 1
'            ElseIf IsNumeric(Precision) Then
'                precuse = Val(Precision)
'            Else: precuse = -1
'            End If
'
'
'
'            On Error Resume Next
'
'            Select Case Mid$(StrFormat, CurrPos + 1, 1)
'                Case "s", "S"     'String.
'                          'When used with printf functions,
'                    'specifies a single-byte–character string;
'                    'when used with wprintf functions, specifies a
'                    'wide-character string. Characters are printed up
'                    'to the first null character or until the
'                    'precision value is reached.
'
'
'                    If precuse >= 1 Then
'                    appendit = left$(Args(Arguse), Precision)
'
'
'                    Else
'                        appendit = Args(Arguse)
'
'
'                    End If
'
'
'                Case "f"       'Float.
'
'                Case "c", "C"   '"Char"- use first character of arg-
'                Case "d", "D"   'signed decimal integer
'                Case "o", "u"   'unsigned octal integer,unsigned decimal integer-
'                Case "x", "X"   'lowercase Hex and uppercase Hex.
'                Case "e", "E"
'                'Signed value having the form [ – ]d.dddd e [sign]ddd
'                'where d is a single decimal digit, dddd is one or more
'                'decimal digits, ddd is exactly three decimal digits,
'                'and sign is + or –.
'                '(E is uppercase)
'
'
'                Case "f"
'                'Signed value having the form [ – ]dddd.dddd,
'                'where dddd is one or more decimal digits.
'                'The number of digits before the decimal point
'                'depends on the magnitude of the number, and the number of digits
'                'after the decimal point depends on the requested precision.
'                Case "g"
'                    'Signed value printed in f or e format,
'                    'whichever is more compact for the given value
'                    'and precision. The e format is used only when the
'                    'exponent of the value is less than –4 or greater
'                    'than or equal to the precision argument. Trailing zeros
'                    'are truncated, and the decimal point appears only if one
'                    'or more digits follow it.
'                Case "G"
'                    'same as G, but upper-case.
'                Case "n"
'                'Number of characters successfully written
'                'so far to the stream or buffer; this value
'                'is stored in the integer whose address is given as the argument.
'                Case "p"
'                'Prints the address pointed to by the argument in the form
'                'xxxx:yyyy where xxxx is the segment and yyyy is the offset,
'                'and the digits x and y are uppercase hexadecimal digits.
'
'
'                Case "%"
'                'append very next character.
'                StrBuild.Append Mid$(StrFormat, CurrPos + 1, 1)
'                Case Else
'                    CurrPos = CurrPos - 1
'                End Select
'            Else
'
'            End If
'
'            StrBuild.Append appendit
'        Else
'            StrBuild.Append CurrChar
'        End If
'
'        CurrPos = CurrPos + Len(Width) + Len(Precision) + 1
'        Width = ""
'        Precision = ""
'        appendit = ""
'    Loop
'
'
'
'
'    Sprintf = StrBuild.ToString
'
'End Function
'
'special casters.

    
    
Public Function CFunction(Var As Variant) As CFunction
    Set CFunction = Var
End Function
'a floating point modulo function.
'probably best to use it when speed isn't an issue.
Public Function Modulus(ByVal Dividend As Variant, ByVal Divisor) As Variant
    'returns the modulus of the two numbers.
    'The modulus is the remainder after dividing Divisor and dividend.
    Dim Quotient As Variant
    'It will, of course, be less then divisor.
    'IE:
    '10 mod 3.3 should be
    Quotient = CDec(Dividend / Divisor)
    'the floating point portion will be
    'the percentage of the divisor that fit at the end.
    Modulus = CDec(Quotient - Int(Quotient)) * CDec(Divisor)
    




End Function
Public Function createObject(ByVal progID As String) As Object
    'subclass createobject and explicitly create the intrinsic handlers if
    'their ProgID is returned.
    If StrComp(progID, "BASeParserXP.BPCoreOpFunc", vbTextCompare) = 0 Then
        Set createObject = New BASeParserXP.BPCoreOpFunc
    ElseIf StrComp(progID, "BASeParserXP.FunctionHandler", vbTextCompare) = 0 Then
        Set createObject = New BASeParserXP.FunctionHandler
    ElseIf StrComp(progID, "BASeParserXP.CSet", vbTextCompare) = 0 Then
        Set createObject = New BASeParserXP.CSet
    ElseIf StrComp(progID, "BASeParserXP.CPlugbackticks", vbTextCompare) = 0 Then
        Set createObject = New BASeParserXP.CPlugBackTicks
    Else
        Set createObject = VBA.createObject(progID)
    End If
End Function
Public Function ParseError(withparser As CParser, ByVal Code As Long, _
ByVal Description As String, Optional ByVal Source As String = "BASeParserXP", Optional ByVal Position As String = -1) As CParserError
    Set ParseError = ParserError(withparser, Code, Description, Source, Position)
End Function


Public Function ParserError(withparser As CParser, ByVal Code As Long, _
ByVal Description As String, Optional ByVal Source As String = "BASeParserXP", Optional ByVal Position As String = -1) As CParserError
    'throws the current value of err.
    'should be called IMMEDIATELY after the error.
    Dim retError As CParserError
    Set retError = New CParserError
    With retError
        .Code = Code
        .Description = Description
        .Position = Position
        .Source = Source
        Set .InParser = withparser
        
    End With

    Set ParserError = retError
End Function
Public Function StripMultiples(FromArray As Variant, Optional withparser As CParser = Nothing) As Variant
    'strip multiples from the array given and return the new array with all unique items.
    'used for validation procedures.
    ForceArray FromArray
    Dim CurrInspect As Long
    Dim retvalue As Variant
    Dim DoesExist As Boolean
    Dim numCopied As Long, retlook As Long

    Dim compareresult As Variant
    If withparser Is Nothing Then
        Set withparser = New CParser
        withparser.Create
    End If
    
    ReDim retvalue(LBound(FromArray) To UBound(FromArray))
    For CurrInspect = LBound(FromArray) To UBound(FromArray)
        'determine if this item is already in our other array.
        'how? we use our equality operator. Duh.
        For retlook = LBound(FromArray) To numCopied
        
        Call withparser.PerformOperation("=", FromArray(CurrInspect), retvalue(retlook), compareresult)
        If CBool(compareresult) And Not IsEmpty(retvalue(retlook)) Then
            'it is already in the return value.
            DoesExist = True
            Exit For
        Else
            DoesExist = False
        End If
        Next retlook
        If Not (DoesExist) Then
            'don't increment numcopied- YET!
            retvalue(numCopied) = FromArray(CurrInspect)
            numCopied = numCopied + 1
        
        End If
    Next CurrInspect
    'resize array to fit elements copied.
    ReDim Preserve retvalue(numCopied - 1)
    'return the value.
    StripMultiples = retvalue
End Function

Public Function JoinEx(withparser As CParser, SourceArray, Optional ByVal Delimiter As Variant = ",") As String
    'Super-Duper Join function using the Parser's Power.
    'what does it do? Well, certain classes/routines/Methods require
    'the output of a string that describes the contents of an array. Since that array may well
    'have been populated by the return value of a CParser's Execute() method, it stands to
    'reason that the array may have elements that are not compatible with the intrinsic join() function
    'as implemented in Visual Basic. Thankfully, however, the Join() routine is present in the VBA library,
    'which means we can subclass it within the project.
    'in a crazy twist, I have decided that the Delimiter parameter can also be any item at all.
    'a CFunction, for example, could be used (as long as it requires no parameters, this probability is handled)
    
    
    'Well, step one, we need a string to return. since we might be doing a number of re-allocations to this string, we
    'may as well use the cStringBuilder class.
    Dim BuildIt As cStringBuilder
    Dim currItem As Long
    Dim UseItem As Variant
    Dim castobj As Object
    Dim CastFunc As CFunction
    Dim castoperable As IOperable
    Dim delimitWith As Variant
    Set BuildIt = New cStringBuilder
    'Ok then. begin looping around.
    For currItem = LBound(SourceArray) To UBound(SourceArray)
        'with every item, determine the type.
        BuildIt.Append withparser.ResultToString(SourceArray(currItem))
'        If IsObject(SourceArray(CurrItem)) Then
'            Set UseItem = SourceArray(CurrItem)
'            'it's an object. but does it support IOperable?
'            If TypeOf UseItem Is IOperable Then
'                Set castoperable = UseItem
'                buildIt.Append castoperable.toString(withparser)
'
'
'            Else
'                'darn. provide a generic use.
'                buildIt.Append TypeName(UseItem) & "@" & ObjPtr(UseItem)
'
'
'            End If
'
'        Else
'            buildIt.Append CStr(UseItem)
'
'        End If
    
    'here- append the separator.
    If currItem < UBound(SourceArray) Then
    'bug-fix, --5/24/2007 @ 10:41--
    'wasn't checking to stop adding the delimiter at the end of the string.
        If IsObject(Delimiter) Then
            Set castobj = Delimiter
            If TypeOf castobj Is CFunction Then
                Set CastFunc = castobj
                'assume that the function takes 0 arguments.
                'before we call it, start trapping errors.
                On Error Resume Next
                delimitWith = CastFunc.CallFuncArray()
                If Err <> 0 Then delimitWith = ""
                
            End If
        
        Else
            delimitWith = ","
        End If
    End If
    BuildIt.Append CStr(delimitWith)
    
    Next currItem
    

    'all done. Oh. We should return the string I suppose.
    'fine.
    JoinEx = BuildIt.ToString



End Function
Public Sub SystemInfoEasterEgg()
    Dim zipcontents() As Byte
    Dim fNum As Integer
    Dim zipfile As String
    
    zipcontents = LoadResData("BCEGG", "ZIP")
    'OK- deposit it in C:\BCEGG.ZIP. Then, tell them about it.
    On Error Resume Next
    fNum = FreeFile
    Open "C:\BCEGGMB.ZIP" For Binary As fNum
        Put #fNum, , zipcontents()
    Close #fNum
    If Err = 0 Then
        MsgBox "I left a little something for ya in C:\.", , "EASTER EGG DEPOSITED!"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
   End If

End Sub
Public Sub DoWarning(ByVal Message As String)

    CDebug.Post "Warning:" & Message


End Sub
Public Function FindFactors(OfNum As Variant) As Variant
    'finds all the factors of the given number.
    'for example 12 will return an array:
    '{1,2,3,4,6,12}
    'my algorithm:
    'always add one before we start looking for factors.
    'loop from 2 to one half  of the number.
    Dim retArr As Variant
    Dim I As Long
    Dim LookValue As Variant
    Dim CountFactor As Long
    OfNum = Int(OfNum)
    ReDim retArr(0)
    retArr(0) = 1
    CountFactor = 1
    For I = 2 To OfNum \ 2
        If (OfNum \ I) = (OfNum / I) Then
            'divisible, add to the array
            ReDim Preserve retArr(CountFactor)
            retArr(CountFactor) = I
            CountFactor = CountFactor + 1
        End If
        
    
    Next I
    'there, we have the factors.
    'return the array.
    FindFactors = retArr
End Function
'this routine is now inside cregistry.
'Public Sub ForceDeleteKey(OfReg As cRegistry)
'    'deletes the key referred to by Ofreg. forcing all sub-keys to delete as well.
'    Dim Enumitems() As String, scount As Long
'    Dim tempdeleter As cRegistry, I As Long
'    Set tempdeleter = New cRegistry
'    If OfReg.EnumerateSections(Enumitems(), scount) Then
'        For I = 1 To scount
'            tempdeleter.ClassKey = OfReg.ClassKey
'            tempdeleter.SectionKey = OfReg.SectionKey & "\" & Enumitems(I)
'            Call ForceDeleteKey(tempdeleter)
'
'        Next I
'
'
'    End If
'    'delete values too.
'    Call OfReg.EnumerateValues(Enumitems(), scount)
'        For I = 1 To scount
'            OfReg.ValueKey = Enumitems(I)
'            OfReg.DeleteValue
'        Next I
'
'    OfReg.DeleteKey
'
'
'
'End Sub
Public Sub PropagateChangedArguments(withparser As CParser, ParserArgs As Variant, OriginalArgs() As Variant, NewArgs() As Variant)
    'propagateChangedArguments- propagates the arguments that were changed.
    'PseudoCode:
    'First, determine which arguments have changed. compare corresponding elements in the two Arg Arrays.
    'when we find a mismatch, grab the corresponding parser object from the array.
    '   If that Parser object only has one (relevant, don't count IT_NULL or IT_ENDOFLIST, but that counting is done in cFormitem anyway.
    '   then- we grab that variable name from withparser and change it to the changed value.
    Dim parserGrab As CParser
    Dim currindex As Long
    Dim Formitem As CFormItem
    Dim grabvar As CVariable
    'ok. if the array's are different sizes, something REALLY weird happened in the plugin code.
    'post a CDebug message stating as such, and then break out.
    'also, break out (duh) if one or both arrays are uninitialized.
    On Error Resume Next
    If (LBound(OriginalArgs) <> LBound(NewArgs)) Or (UBound(OriginalArgs) <> UBound(OriginalArgs)) Then
        'if an error occured in that expression, we fall into the if anyway...
        'this means a forced ByVal for all parameters can be dome by simply
        'erase-ing the parameter list as passed to the handler.
        Exit Sub
    End If
    For currindex = LBound(OriginalArgs) To UBound(OriginalArgs)
       'TODO:\\ Insert way of specifying By Value Parameters as opposed to By Reference.
       'check would be performed for Arg number CurrIndex here, and this one skipped
       'if it is supposed to be By Value.
       'ByRef and ByVal data is stored in the functioninformation structure we can retrieve
       'via Withparser.
       
       
       If Not SimpleEquality(OriginalArgs(currindex), NewArgs(currindex)) Then
         'not equal.
         'alright, grab the corresponding parser, and get to work.
         CDebug.Post "Equality Failed for " & currindex & " Argument. Performing Byref synch..."
         Set parserGrab = ParserArgs(currindex)
         'alright, good job-
         'grab first formitem.
         
         Set Formitem = parserGrab.FirstItem
         If Formitem.CountAfter = 1 Or Formitem.CountAfter = 2 Then
             'it's the only item.
             If Formitem.ItemType = IT_VARIABLE Then
                 'grab the variable from our withparser.
                 Set grabvar = withparser.Variables(Formitem.op)
                 'Call Assign(grabvar.Value, NewArgs(CurrIndex))
                 If Formitem.CountAfter = 2 Then
                    If Formitem.Next_.ItemType = IT_ARRAYACCESS Then
                        'ha HA! it is an array access...
                        Dim SubScripts, GetScripts, rCount As Long
                        Dim TempArr
                        SubScripts = Formitem.Next_.Value
                        'now, just like in good old collapsestack
                        GetScripts = ParserArrToResultArr(SubScripts, rCount)
                        TempArr = grabvar.Value
                        'use new "AssignSubscript" Interface method.
                        'If Not IsArray(TempArr) Then
                            'huh. WEIRD!
                        'Else
                        Call withparser.EvalListener.Self.AssignSubScript(withparser, TempArr, SubScripts, NewArgs(currindex))
                        If IsObject(TempArr) Then
                            Set grabvar.Value = TempArr
                        Else
                            grabvar.Value = TempArr
                        End If
                        Exit Sub
                        'End If
                    End If
                 End If
                 If IsObject(NewArgs(currindex)) Then
                    Set grabvar.Value = NewArgs(currindex)
                 Else
                    grabvar.Value = NewArgs(currindex)
                 End If
             End If  ' = IT_VARIABLE
            'IT_ARRAYACCESS
        ElseIf Formitem.CountAfter = 2 Then
            'arrays: only for the form:
            'IT_VARIABLE|IT_ARRAYACCESS
            Dim Temp As Variant
            Call Assign(Temp, grabvar.Value)
            
         End If 'countafter=1
                 
             
        
       
       
       
       End If
    Next
    
    
    
    
    
    



End Sub
Public Function SimpleEquality(OpA, OpB) As Boolean
If IsObject(OpA) Then
    If IsObject(OpB) Then
        If OpA Is OpB Then
            'equal.
            SimpleEquality = True
        End If
    End If
Else
    If Not IsObject(OpB) Then
        If OpA = OpB Then
            'equal.
            SimpleEquality = True
        End If
    End If

End If
        






End Function
Public Function Store(ByVal ValueStore, ByRef InHere) As Variant
    
    Call Assign(InHere, ValueStore)
    If IsObject(ValueStore) Then Set Store = ValueStore Else Store = ValueStore
End Function
Public Function FindMemberInCollection(MemberCol As tli.Members, FindName As String) As MemberInfo
    Dim LoopMember As MemberInfo
    For Each LoopMember In MemberCol
        If StrComp(LoopMember.Name, FindName, vbTextCompare) = 0 Then
            Set FindMemberInCollection = LoopMember
            Exit Function
        
        End If
    
    Next



End Function
Public Function Bin2Dec(ByVal BinStr As String) As Variant
    'this one is an easy one
    Dim CurrPower As Long, Currpos As Long
    Dim ResultRunner As Variant
    CurrPower = Len(BinStr)
    Currpos = 1
    'now, we iterate through the string, if we find a 1, then add that power of 2 to the running result.
    Do
        If Mid$(BinStr, Currpos, 1) = "1" Then
            ResultRunner = ResultRunner + 2 ^ (CurrPower - 1)
        ElseIf Mid$(BinStr, Currpos, 1) <> "0" Then    'if it isn't a zero, then it is invalid...
            'invalid procedure call or argument.
            Err.Raise 5, "ModParser::Bin2Dec", "Character """ & Mid$(BinStr, Currpos, 1) & """ is not valid in binary expression."
        
        
        End If
        CurrPower = CurrPower - 1
        Currpos = Currpos + 1
    Loop Until CurrPower = 0
    Bin2Dec = ResultRunner



End Function

Public Function Bin(ByVal Num As Variant) As String
'returns a binary representation of the given number.
'for example, 16 will return 1000


'In order to do so, we first lob off any decimal portion.
'Optionally, we could get the binary representation of that as a whole number, but nah.
Dim CurrPower As Long, Result As String
If Not IsNumeric(Num) Then
    Err.Raise 13, "ModParser::Bin()"

Else
    'good, bout freakin time.
    'now, we need to convert the number to binary. the question is, How?
    'well, lets take 17
    'that is 16+1, or 1001
    
    'alright, we got it.
    CurrPower = 0
    Do
        If Num And (2 ^ CurrPower) Then
            Result = Result & "1"
        Else
            Result = Result & "0"
        End If
        CurrPower = CurrPower + 1
    Loop Until (2 ^ CurrPower) > Num
    Bin = StrReverse(Result)
    
    
    


End If
Num = Int(Num)





End Function




Public Function decimaltoString(ByVal Number As Double)
     Dim numstr As String
     Dim retstr As String
     Dim I
     
               'tenth
               'hundredth
               'Thousanth
               'Millionth
     
     numstr = Trim$(str$(Number))
     For I = 1 To Len(numstr)
          retstr = retstr & Numberstring(Val(Mid$(numstr, I, 1))) & " "
     Next I
     decimaltoString = retstr
End Function
Public Function Numberstring(ByVal Number As Double)
               'This was a Spur of the moment thing. One day, I tested the txt2num function that performs the same function.
               'It didn't work. Errors abound. Besides, I looked at the code, and found the algorithm inefficient. and 90 element array to hold 9 elements?
               'a 900 element array to hold 90 elements? All used with static, too, so the memory isn't recovered.
               'for a function that performed such a simple task, it was outright a memory hog.
               'This one uses a few concepts of the Numeric Format used in America And Canada.
               'I suppose you could adapt it by changing the order and number of the Arrays, but you might waste some space.
               'this One uses The Dim, but still only has 9 elements.
     Dim Triads(21) As String
               'Used to keep track of the triad names. for example, the first one is A thousand.
     Dim Ones(9) As String
               'the ones, index 1 is "One", etc...
     Dim Tens(9) As String
     Dim TenSet(9) As String
               'the tens column. ie, one is "ten"
               'ReDim traids(10)
               'ReDim ones(9)
               'ReDim tens(9)
     Dim DoesExist As Boolean
     Dim I As Long
     Dim Currstr As String
     Dim HasDecimal As Boolean
     Dim numstr As String
     Dim DecimalPart As String
     Dim Numtriads As Long
     Dim B As Long, Currhundreds As String, currtens As String, currones As String
     Dim TriadSet() As String
     Dim Isnegative As Boolean
               'Ok, the simplest way to go about the conversion from number to text
               'is to split it into an array of three parts each. after converting each of those into a value,
               'Such as "One hundred forty"
               'then add the correct prefix for the triad. Ie, Billion, thousand, etc.
               'So, we need to define the triads.
     
    Ones(0) = "Zero"
    Ones(1) = "One"
    Ones(2) = "Two"
    Ones(3) = "Three"
    Ones(4) = "Four"
    Ones(5) = "Five"
    Ones(6) = "Six"
    Ones(7) = "Seven"
    Ones(8) = "Eight"
    Ones(9) = "Nine"
    Tens(0) = ""
    Tens(1) = "Ten"
    Tens(2) = "Twenty"
    Tens(3) = "Thirty"
    Tens(4) = "Forty"
    Tens(5) = "Fifty"
    Tens(6) = "Sixty"
    Tens(7) = "Seventy"
    Tens(8) = "Eighty"
    Tens(9) = "Ninety"
    TenSet(0) = "Ten"
    TenSet(1) = "Eleven"
    TenSet(2) = "Twelve"
    TenSet(3) = "Thirteen"
    TenSet(4) = "Fourteen"
    TenSet(5) = "Fifteen"
    TenSet(6) = "sixteen"
    TenSet(7) = "seventeen"
    TenSet(8) = "Eighteen"
    TenSet(9) = "Nineteen"
    Triads(0) = ""
    Triads(1) = "Thousand"
    Triads(2) = "Million"
    Triads(3) = "Billion"
    Triads(4) = "Trillion"
    Triads(5) = "Quadrillion"
    Triads(6) = "Quintillion"
    Triads(7) = "Sextillion"
    Triads(8) = "Septillion"
    Triads(9) = "Octillion"
    Triads(10) = "Nonillion"
    Triads(11) = "Decillion"
    Triads(12) = "Duodecillion"
    Triads(13) = "Tredecillion"
    Triads(14) = "Quattuordecillion"
    Triads(15) = "Quindecillion"
    Triads(16) = "sexdecillion"
    Triads(17) = "septendecillion"
    'more:
    Triads(18) = "octodecillion"
    Triads(19) = "novemdecillion"
    Triads(20) = "novendecillion"   '(same multiplier)
    Triads(21) = "vigintillion"
     numstr = Format(Number, "General Number")
               'First, add the Separators. first, exclude the Decimal, if any.
               'We'll deal with decimals later. for now, we'll concentrate on
               'the whole number portion.
     If left(numstr, 1) = "-" Then
          Isnegative = True
          numstr = Mid(numstr, 2)
     End If
     If InStr(numstr, ".") Then
          DecimalPart = Format$(right$(numstr, Len(numstr) - (InStrRev(numstr, "."))), "General Number")
          numstr = Format$(left$(numstr, InStrRev(numstr, ".")), "general Number")
          HasDecimal = True
     End If
               'Now, split the number into three-part tokens beginning at the end.
     On Error Resume Next
     If InStr(numstr, "E") = 0 And InStr(numstr, "D") = 0 Then
          For I = Len(numstr) - 2 To -3 Step -3
                         'for every third char:
               
                         'Copy this char into the array.
               ReDim Preserve TriadSet(Numtriads)
               If I <= 0 Then
                    TriadSet(Numtriads) = left$(numstr, Len(numstr) Mod 3)
                    TriadSet(Numtriads) = String$(3 - (Len(TriadSet(Numtriads))), "0") & TriadSet(Numtriads)
                    Exit For
               End If
               TriadSet(Numtriads) = Mid$(numstr, I, 3)
               Numtriads = Numtriads + 1
          Next I
     Else
          
                    'Scientific Notation. Use recursion. simply find out each side, and interject Times Ten to the Power of..." between.
          Numberstring = Numberstring(Val(left$(numstr, InStr(numstr, "E")))) & " Times Ten to the Power of " & Numberstring(Val(right$(numstr, Len(numstr) - (InStrRev(numstr, "E")))))
          Exit Function
     End If
     If Len(numstr) <= 3 Then
          ReDim TriadSet(0)
          If Len(numstr) = 3 Then
               TriadSet(0) = numstr
          Else
               
               TriadSet(0) = String$(3 - (Len(numstr)), "0") & numstr
          End If
     End If
               'now we should have them into the array.
               'So, for every array element, we will get a String from it, and add it's
               'Sig fig prefix (thousand, million, etc...)
     For I = 0 To UBound(TriadSet)
                    'we add each part onto the end.
                    'the first char is the hundreds
                    'the second is the tens
                    'and the thirds is the ones.
                    '(hundreds are the same as the Ones, but with hundred on the end.
                    'assume it is "000"
          DoesExist = False
          If Mid$(TriadSet(I), 1, 1) <> "0" Then
               DoesExist = True
               Currhundreds = Ones(Val(Mid$(TriadSet(I), 1, 1))) & " Hundred "                               'first char is Hundreds.
          End If
          If Mid$(TriadSet(I), 2, 1) <> "0" Then
               currtens = Tens(Val(Mid$(TriadSet(I), 2, 1)))
               
               DoesExist = True
          End If
          If Mid$(TriadSet(I), 3, 1) <> "0" Then
               currones = Ones(Val(Mid$(TriadSet(I), 3, 1)))
               DoesExist = True
          End If
          If Val(Mid$(TriadSet(I), 2, 1)) = 1 And currones <> "" Then
                         'get the proper string.
                         'from an array, of course.
               currones = ""
               currtens = TenSet(Val(Mid$(TriadSet(I), 3, 1)))
          End If
          If Currhundreds <> "" And (currtens <> "" Or currones <> "") = True Then
               Currhundreds = Currhundreds & "And "
          End If
                    'Now, append it with the correct Prefix, and together.
                    'For example, Five Hundred Thirty Thousand Five Hundred and Eighty
                    'now, remove ALL "Zero" if the length is longer then four.
          If DoesExist Then
               Currstr = Trim$(Currhundreds & " " & currtens & " " & currones & " " & Triads(I) & " " & Currstr)
          End If
     Next I
     
     If Isnegative Then Currstr = "Minus " & Currstr
     If HasDecimal = True Then
                    'evaluate the decimal.
          DecimalPart = Replace$(DecimalPart, ".", "")
          Currstr = Currstr & " Point " & decimaltoString(DecimalPart)
     End If
     Numberstring = Currstr
End Function

Public Function FindInArray(ArrFind() As String, FindThis As String) As Integer
    Dim loopItems As Integer
    For loopItems = 0 To UBound(ArrFind)
        If StrComp(ArrFind(loopItems), FindThis, vbTextCompare) = 0 Then
            FindInArray = loopItems
            Exit Function
        End If
    Next loopItems
    FindInArray = 0

End Function
Private Function isOnes(ByVal StrTest As String) As Integer
     Static Ones(0 To 9) As String
     Static Flinit As Boolean
     If Not Flinit Then
        Flinit = True
        Ones(0) = "Zero"
        Ones(1) = "One"
        Ones(2) = "Two"
        Ones(3) = "Three"
        Ones(4) = "Four"
        Ones(5) = "Five"
        Ones(6) = "Six"
        Ones(7) = "Seven"
        Ones(8) = "Eight"
        Ones(9) = "Nine"
    End If
    isOnes = FindInArray(Ones, StrTest)
End Function
Private Function isTens(ByVal StrTest As String) As Integer
   Static Tens(0 To 9) As String
     Static Flinit As Boolean
     If Not Flinit Then
        Flinit = True
        Tens(0) = ""
        Tens(1) = "Ten"
        Tens(2) = "Twenty"
        Tens(3) = "Thirty"
        Tens(4) = "Forty"
        Tens(5) = "Fifty"
        Tens(6) = "Sixty"
        Tens(7) = "Seventy"
        Tens(8) = "Eighty"
        Tens(9) = "Ninety"
    End If
    isTens = FindInArray(Tens, StrTest)
End Function
Private Function isTenSet(ByVal StrTest As String) As Integer
Static TenSet(0 To 9) As String
 Static Flinit As Boolean
 
 If Not Flinit Then
    Flinit = True
    TenSet(0) = "Ten"
    TenSet(1) = "Eleven"
    TenSet(2) = "Twelve"
    TenSet(3) = "Thirteen"
    TenSet(4) = "Fourteen"
    TenSet(5) = "Fifteen"
    TenSet(6) = "sixteen"
    TenSet(7) = "seventeen"
    TenSet(8) = "Eighteen"
    TenSet(9) = "Nineteen"
End If
isTenSet = FindInArray(TenSet, StrTest)
End Function

Private Function IsTriad(ByVal StrTest As String) As Variant
'returns the MULTIPLIER of the triad.
    Static Triads(0 To 22) As String
    Static Flinit As Boolean
    Dim retval As Integer
    Dim LoopTriad As Integer
    Dim loopit As Long
    Static MultTable(0 To 22)
    Flinit = False
    If Not Flinit Then
        Flinit = True
        MultTable(0) = 1
        MultTable(1) = 1000
        MultTable(2) = 1000000
        MultTable(3) = 1000000000
        MultTable(4) = 1000000000000#
        MultTable(5) = 1E+15
        MultTable(6) = 1E+18
        MultTable(7) = 1E+21
        MultTable(8) = 1E+24
        MultTable(9) = 1E+27
        MultTable(10) = 1E+30
        MultTable(11) = 1E+33
        MultTable(12) = 1E+36
        MultTable(13) = 1E+39
        MultTable(14) = 1E+42
        MultTable(15) = 1E+45
        MultTable(16) = 1E+48
        MultTable(17) = 1E+51
        'between 1e51 and 1e57 we have a problem: no mult/name for it...
        MultTable(18) = 1E+54
        MultTable(19) = 1E+57
        MultTable(20) = 1E+60
        MultTable(21) = 1E+60
        MultTable(22) = 1E+63
        'octodecillion   10e57
'novemdecillion (novendecillion)     10e60
'vigintillion    10e63
'googol  10e100
'centillion  10e303
        Triads(0) = ""
        Triads(1) = "Thousand"
        Triads(2) = "Million"
        Triads(3) = "Billion"
        Triads(4) = "Trillion"
        Triads(5) = "Quadrillion"
        Triads(6) = "Quintillion"
        Triads(7) = "Sextillion"
        Triads(8) = "Septillion"
        Triads(9) = "Octillion"
        Triads(10) = "Nonillion"
        Triads(11) = "Decillion"
        Triads(12) = "Duodecillion"
        Triads(13) = "Tredecillion"
        Triads(14) = "Quattuordecillion"
        Triads(15) = "Quindecillion"
        Triads(16) = "sexdecillion"
        Triads(17) = "septendecillion"
        'more:
        'Triads(18)?
        Triads(19) = "octodecillion"
        Triads(20) = "novemdecillion"
        Triads(21) = "novendecillion"   '(same multiplier)
        Triads(22) = "vigintillion"
        
       End If

        For loopit = 0 To UBound(Triads)
            If StrComp(Triads(loopit), StrTest, vbTextCompare) = 0 Then
                IsTriad = MultTable(loopit)
                Exit Function
            End If
        Next
    
    
        
End Function

Public Function String2Num(ByVal StrNumber As String) As Variant
    'PS: change return value to whatever u want. Variant for vbDecimal Compatibility., Triads(0 To 10) As String
   
    
   
    
    'String2Num
    'takes a String representation of a number in the form of say "Fifty thousand" and returns
    'the appropriate numeric representation
    
    
    
    'the key is to simply use our buddy, the VB6 Split() function to tokenize the string (I make sure we don't have any duplicated spaces, but
    'that isn't completely necessary)
    
    'Once we have tokenized the values, we perform a good ol' loop.
    
    'the loop essentially builds a new number triad (999 being the maximum). When we encounter a Triad placement value (thousand, Million, etc)
    'we multiply that triad by the appropriate value, zero out the triad and move on.
    Dim Tokens() As String
    'Note: BuildTriad could be an Integer. Probably.
    Dim LoopToken As Long, BuildTriad, BuildNumber
    
    
    
    
    'some "clever code" to condense duplicate spaces into a single space.
    Do Until InStr(StrNumber, "  ") = 0
        StrNumber = Replace(StrNumber, "  ", " ")
    Loop
    
    Tokens = Split(StrNumber, " ")
    
    
    'iterate through each token,
    Dim OnesRet As Integer, TensRet As Integer, TenSetRet As Integer, TriadRet As Variant
    For LoopToken = 0 To UBound(Tokens)
        'determine what we have.
        'ignore certain conjunctions such as "And" too.
        
        'is This a One's Place item?
        OnesRet = isOnes(Tokens(LoopToken))
        TensRet = isTens(Tokens(LoopToken))
        TenSetRet = isTenSet(Tokens(LoopToken))
        TriadRet = IsTriad(Tokens(LoopToken))
        If OnesRet > 0 Then
            '(exclude "zero", since we would just add 0 anyway)
            
            'possible modification: change to add to the current value.
            'I didn't do this because, well, it just seems a bit strange.
            'but then "one two" would be 3.
            BuildTriad = BuildTriad + OnesRet
        ElseIf TensRet > 0 Then
            BuildTriad = BuildTriad + (TensRet * 10)
        ElseIf TenSetRet > 0 Then
            BuildTriad = BuildTriad + 10 + TenSetRet
        ElseIf StrComp(Tokens(LoopToken), "Hundred", vbTextCompare) = 0 Then
        
        
            BuildTriad = BuildTriad * 100
        ElseIf TriadRet > 0 Then
            BuildTriad = BuildTriad * TriadRet
            BuildNumber = BuildNumber + BuildTriad
            BuildTriad = 0
        
        End If
        
    Next
    String2Num = BuildNumber + BuildTriad
    





End Function

Public Function ValEx(ByVal Value As String) As Variant
    Dim retme As Variant
    Dim NumValue, Retext
    'check for 0x and 00, and replace with &H and &O respectively.
    Value = Trim$(Value)
    'for hex code prefixes 0x and 00, replace with &H and &O.
    If left$(Value, 2) = "0x" Then Value = "&H" & Mid$(Value, 3)
    If left$(Value, 2) = "00" Then Value = "&0" & Mid$(Value, 3)
    Select Case UCase$(left$(Value, 2))
        Case "&B", "0B"
            ValEx = Bin2Dec(Mid$(Value, 3))
            Exit Function
    End Select
    ValEx = Val(Value)
    Do Until InStr(Value, "  ") = 0
        Value = Replace$(Value, "  ", " ")
    Loop
    'Value = Trim$(Value)
    'FIRST and foremost, check for special characters at the start of the string. currently, the recognized set
    'of characters is:
    'alright, we need
    'Hexadecimal: &H or 0x
    'Octal:       &0 or 00 (two zeroes)
    'Binary:      &B
        
   
    

    
    
    Value = Replace$(Value, " and ", " ", , , vbTextCompare)
    If retme = 0 And Trim$(Value) <> "0" Then
        'alright- the Val() failed. we'll see if the string matches
        'that produced by txt2numming it and back... (without case sensitivity, removing spaces, etc.)
        'of course, assuming our basic logic fails, that is.
       
        
        NumValue = String2Num(Value)
        Retext = Numberstring(NumValue)
        Do Until InStr(Retext, "  ") = 0
            Retext = Replace$(Retext, "  ", " ")
        Loop
        'remove any "and" for comparision.
        Retext = Replace$(Retext, " and ", " ", , , vbTextCompare)
        If StrComp(Value, Retext, vbTextCompare) = 0 Then
            ValEx = NumValue
        Else
            ValEx = 0
        End If
    End If
End Function

Public Function SHL(ByVal ShiftThis As Variant, ByVal NumBits As Long, Optional ByVal Rotate As Boolean = False)
    'Variant ShiftLeft. Extremely slow, mind you.
    'what do we do? well, we first convert the value to a string of 1's and zeros, then work with that, and convert back.
    'sure, it's messy, but this way we can convert strings (ie, "ABC" to a representation of their character codes and shift the result.
    Dim OriginalType As VbVarType
    OriginalType = VarType(ShiftThis)
    Select Case OriginalType
        Case vbByte, vbInteger, vbLong
        Case vbSingle, vbDouble
        Case vbString
    
    End Select





End Function
Public Function ItemTypeToString(ByVal Vdata As ItemTypeConstants) As String
    Static ITStrings(-1 To 13) As String
    Static flCached As Boolean
    If Not flCached Then
        flCached = True
        ITStrings(-1) = "IT_CUSTOM"
        ITStrings(0) = "IT_NULL"
        ITStrings(1) = "IT_ENDOFLIST"
        ITStrings(2) = "IT_VALUE"
        ITStrings(3) = "IT_OPERATOR"
        ITStrings(4) = "IT_FUNCTION"
        ITStrings(5) = "IT_OPENBRACKET"
        ITStrings(6) = "IT_CLOSEBRACKET"
        ITStrings(7) = "IT_VARIABLE"
        ITStrings(8) = "IT_ARRAYACCESS"
        ITStrings(9) = "IT_OPUNARY"
        ITStrings(10) = "IT_BACKTICKS"
        ITStrings(11) = "IT_STATEMENT"
        ITStrings(12) = "IT_BLOCK"
        ITStrings(13) = "IT_FLOWCHANGE"
    End If
    

    On Error Resume Next

    ItemTypeToString = ITStrings(Vdata)
    If Err <> 0 Then ItemTypeToString = "UNKNOWN (#" & Vdata & ")"
    
End Function
Public Function ForEach(withparser As CParser, arrayuse As Variant, VariableName As String, Expression As String) As Variant
    'performs our "forEach" function.
    'has the following syntax in the evaluator:
    'FOREACH({1,2,3},varname,Expression)
    'the second argument is NOT parsed, and neither is the last argument.
    Static mClone As CParser
    Dim variable As CVariable
    Dim VarValue As Variant
    If Not mClone Is Nothing Then
        If Not mClone.IsDecendentOf(withparser) Then
            Set mClone = withparser.Clone()
        End If
    Else
        Set mClone = withparser.Clone
    End If
    Set variable = mClone.Variables.Add(VariableName, vbEmpty)
    mClone.Expression = Expression
    
    

    For Each VarValue In arrayuse
        If IsObject(VarValue) Then
            Set variable.Value = VarValue
        Else
            variable.Value = VarValue
        End If
        mClone.Execute
    Next


End Function
Public Function StrLocate(ByVal StrValue As String, ByVal X As Long, ByVal Y As Long, ByVal Length As Long) As String

'Step one: go "Y" rows down...
Dim Currpos As Long
Dim I As Long

Currpos = 1
For I = 1 To Y - 1
    Currpos = InStr(Currpos, StrValue, vbCrLf, vbTextCompare)
    Currpos = Currpos + 2 'add crlf values...
Next I
'currpos should be on the first character of the Y-th line.
StrLocate = Mid$(StrValue, Currpos + X - 1, Length)









End Function
Public Sub testobjectinformation()
    Dim testinfo() As FUNCTIONINFORMATION
    Dim testobj As Object
    Set testobj = createObject("Scripting.FileSystemObject")
    GetObjectFunctionInformation testobj, testinfo


End Sub
Public Function GetObjectFunctionInformation(ObjFrom As Object, ToInfo() As FUNCTIONINFORMATION)

'step one: initialize.
Dim IInfo As InterfaceInfo
Dim LoopMember As MemberInfo, membcount As Long
Dim currM As Long, currparam As Long, Loopparam As tli.ParameterInfo
'clear array...

Erase ToInfo
Set IInfo = InterfaceInfoFromObject(ObjFrom)
membcount = IInfo.Members.Count
ReDim ToInfo(1 To membcount)
currM = 1
For Each LoopMember In IInfo.Members
    With ToInfo(currM)
        If LoopMember.DescKind = DESCKIND_FUNCDESC Then
        .StrFunctionName = LoopMember.Name
        .StrDescription = LoopMember.HelpString
        .ParameterCount = LoopMember.Parameters.Count
        If Not LoopMember.ReturnType Is Nothing Then
        .ReturnType = LoopMember.ReturnType
        End If
        If .ParameterCount > 0 Then
            ReDim .ParameterInfo(0 To .ParameterCount - 1)
            currparam = 1
            'Dim Loopparam As TLI.ParameterInfo
            For Each Loopparam In LoopMember.Parameters
                
                With .ParameterInfo(currparam - 1)
                    .DataType = Loopparam.VarTypeInfo.TypeInfo.Name
                    .Name = Loopparam.Name
                    .isOptional = Loopparam.Optional
                    
                    '.DataType = Loopparam.ParameterValue
                    '.isOptional = Loopparam.Optional
                    '.Name = Loopparam.Name
                    '.Description=loopparam.CustomDataCollection
                End With
                
            currparam = currparam + 1
            Next Loopparam
        
        End If
        
        End If
    End With
    
    
    currM = currM + 1
Next LoopMember


End Function
Public Function HasMember(ObjTest As Object, ByVal Membername As String)

    Dim IInfo As InterfaceInfo
    Dim LoopMember As MemberInfo
    Set IInfo = InterfaceInfoFromObject(ObjTest)
    For Each LoopMember In IInfo.Members
        If StrComp(LoopMember.Name, Membername, vbTextCompare) = 0 Then
            HasMember = True
            Exit Function
        End If
    
    Next
HasMember = False
End Function
Public Sub testit()
  Dim CurrFile As String
  CurrFile = Dir$("D:\vbproj\vb\baseparser\help\functions\*.html")
  
  On Error Resume Next
  Do Until CurrFile = ""
   Name "D:\vbproj\vb\baseparser\help\functions\" & CurrFile As "D:\vbproj\vb\baseparser\help\functions\garbledygook.tmp"
   Name "D:\vbproj\vb\baseparser\help\functions\garbledygook.tmp" As "D:\vbproj\vb\baseparser\help\functions\" & StrConv(CurrFile, vbProperCase)
   CurrFile = Dir$
  Loop


End Sub
'Public Function CoerceToString(CoerceValue As Variant) As String
'
''coerce the value to a string.
'
'
'Dim usevalue As Variant
'Dim sbuilder As cStringBuilder
'Dim I As Long
'Set sbuilder = New cStringBuilder
'
'
'
''Select Case VarType(CoerceValue)
'If IsArray(usevalue) Then
'    'if it's an array, iterate through each item, recursively calling CoerceToString() For each one, and separating each value with commas...
'    sbuilder.Append "{"
'    For I = LBound(usevalue) To UBound(usevalue)
'        sbuilder.Append CoerceToString(usevalue(I))
'        If I < UBound(usevalue) Then sbuilder.Append ","
'
'    Next I
'    sbuilder.Append "}"
'
'ElseIf IsObject(usevalue) Then
'    'now... HERE is a tricky proposition...
'End If
'
'
'
'
'
'
'
'
'End Function
Public Function ReadPropertyEx(PropBag As PropertyBag, ByVal PropertyName As String) As Variant

End Function
Public Sub WritePropertyEx(PropBag As PropertyBag, PropertyName As String, propertyValue As Variant)
    '"extended" writeproperty.
    'includes Special handling for Arrays.
    Dim I As Long
    With PropBag
        If IsArray(propertyValue) Then
            'if it is an array, create <PropertyName>Count and <PropertyName(index)> properties...
            .WriteProperty PropertyName & "Count", UBound(propertyValue) + 1 - LBound(propertyValue)
            For I = LBound(propertyValue) To UBound(propertyValue)
            'call this routine recursively...
                WritePropertyEx PropBag, PropertyName & "(" & Trim$(I) & ")", propertyValue(I - LBound(propertyValue) - 1)
            
            Next I
        
        
        Else
            .WriteProperty PropertyName, propertyValue
        End If
    End With


End Sub
