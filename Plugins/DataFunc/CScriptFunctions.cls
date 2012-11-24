VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScriptFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CScriptFunctions
'Implements the Function/Operator code that allows the parser to use Script files.
'NOTE:


'naturally- this isn't easy- for me.
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Private mRTState As Collection      'of ScriptControl objects.
Private Const mInternalCode As String = "Public Sub AddScript(Byval Namespace,Byval Module,Byval Filename)" & vbCrLf & "    ScriptHandler.LoadScriptIntoModule Namespace,module,Filename" & vbCrLf & "End Sub" & vbCrLf
Private mParser As CParser
'Private mSettings As SettingsData
Implements IEvalEvents
Implements IPluginUIData

Private Sub InitializeRuntime()
    'initialize the Run-time state.
    'this involves instantiating a Script control for each language.
    
End Sub
'Private Sub LoadSettingsData(ToDat As SettingsData, Optional ByVal EraseCurrent As Boolean = True)
'  'Settings Data- Found in "INIFILE"
'  'first step- open the Inifile (duh)
'  'iterate through the lines.
'  'ignore lines that start with ";"
'  'if we find a line start with "[", create
'  'and add a new ScriptControl- Language=the text within the "[".
'  'otherwise, use the currently assigned scriptcontrol, and assign
'  'the properties from the file.
'  Dim fNum As Long
'  Dim Lineparse As String
'  Dim langExtract As String
'  Dim CurrCtl As ScriptControl
'  Dim Numloaded As Long
'  On Error GoTo LoadError
'  fNum = FreeFile
'    If EraseCurrent Then
'        Set mRTState = New Collection
'    End If
'  'Step one: change to ini file directory.
'  'this won't support UNC.
'  ChDrive INIFILE
'  ChDir Mid$(INIFILE, 1, InStrRev(INIFILE, "\"))
'  Open INIFILE For Input As fNum
'  Do Until EOF(fNum)
'    Line Input #fNum, Lineparse
'    'parse.
'    Lineparse = Trim$(Lineparse)
'    Lineparse = Replace$(Lineparse, vbTab, "")
'
'    If Not (Left$(Lineparse, 1) = ";" Or Lineparse = "") Then
'        'if it isn't a comment...
'        If Left$(Lineparse, 1) = "[" Then
'            'a Section header.
'            langExtract = Mid$(Lineparse, 2, InStrRev(Lineparse, "]") - 2)
'            langExtract = Replace$(langExtract, "]", "")
'            ToDat.Numlanguages = ToDat.Numlanguages + 1
'            ReDim Preserve ToDat.mLangData(1 To ToDat.Numlanguages)
'            Set CurrCtl = New ScriptControl
'            CurrCtl.Language = langExtract
'            'add the parser as an object.
'            CurrCtl.AddObject "Parser", mParser, False
'            mRTState.Add CurrCtl, langExtract
'
'        Else
'
'            'set todat.mLangData(todat.Numlanguages) values.
'
'                'format of "lineparse"
'                'prop=Value
'                'ie:
'                'EXTENSIONS="VBS;BAS"
'            Dim propName As String
'            Dim propVal As String
'            propName = Mid$(Lineparse, 1, InStr(Lineparse, "=") - 1)
'            propVal = Mid$(Lineparse, Len(propName) + 2)
'            With ToDat.mLangData(ToDat.Numlanguages)
'                    .LanguageName = CurrCtl.Language
'                Select Case UCase$(propName)
'                    Case "EXTENSIONS", "EXT"
'                        .Extensions = propVal
'                    Case "LOAD"
'                        'load that script.
'                        .NumScripts = .NumScripts + 1
'                        ReDim Preserve .Scripts(1 To .NumScripts)
'                        .Scripts(.NumScripts) = propVal
'                        'OK- insert the file contents into
'                        'the current control.
'                        'this means that at this point, different
'                        'Script files specified via LOAD= cannot
'                        'have duplicate names and such (they
'                        'are all injected into the same ScriptControl)
'                        On Error Resume Next
'                        CurrCtl.AddCode LoadFileContents(propVal)
'                        On Error GoTo LoadError
'                End Select
'
'
'
'
'            End With
'
'
'
'
'        End If
'    End If
'    Loop
'  Close fNum
'    Exit Sub
'LoadError:
'    'oh dear.
'    'Error occured during initialization.
'
'End Sub

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
End Sub

Private Function IEvalEvents_AssignSubScript(withparser As BASeParserXP.CParser, OnValue As Variant, subscripts As Variant, assignthis As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeFunctionCall(withparser As CParser, FuncCalling As String, WithHandler As BASeParserXP.IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'
Dim LoopCtl As ScriptControl
Dim loopproc As MSScriptControl.Procedure

Set LoopCtl = FindProcedureControl(FuncName)
IEvalEvents_CanHandleFunction = (Not LoopCtl Is Nothing)
'For Each LoopCtl In mRTState
'    On Error Resume Next
'    For Each loopproc In LoopCtl.Procedures
'        If Err <> 0 Then Exit For
'        If StrComp(loopproc.Name, FuncName, vbTextCompare) = 0 Then
'            'yep.
'            IEvalEvents_CanHandleFunction = True
'            Exit Function
'        End If
'
'
'    Next
'Next



End Function

Private Function IEvalEvents_CanHandleOperator(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'
End Function

Private Property Get IEvalEvents_Description() As String
'
IEvalEvents_Description = "CScriptFunctions BASeParser Plugin."
End Property

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As BASeParserXP.CParser) As String
'
Dim LoopCtl As ScriptControl
Dim loopproc As MSScriptControl.Procedure
Dim strb As String
For Each LoopCtl In mRTState
    On Error Resume Next
    For Each loopproc In LoopCtl.Procedures
            If Err <> 0 Then Exit For
           
           strb = strb & " " & loopproc.Name & " "
           
        
            
    
    Next
Next
strb = Replace$(strb, "  ", " ")
IEvalEvents_GetHandledFunctionString = strb
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
Set IEvalEvents_GetPluginUIData = Me
End Function

Private Sub IEvalEvents_GetUnaryOperation(withparser As BASeParserXP.CParser, ByVal OpUse As String, ByVal OpA As Variant, currItem As BASeParserXP.CFormItem, retval As Variant)
'
End Sub
Private Function FindProcedureControl(ByRef ProcName As String) As ScriptControl
    Dim findCtl As ScriptControl
    Dim findproc As MSScriptControl.Procedure
    
    'Here is a great example of good use of bad practice (I think)
    'we will modify the passed procname procedure if a direct reference is made to
    'a particular item within our collection.
    'However, luckily, this will NOT be passed back to the BASeParser core, since all interface
    'methods pass function names as ByVal.
    
    
    'if we have a dot-notation access here, (check for a dot) see if we have
    'a Script Control with that key in the collection.
    If InStr(ProcName, ".") <> 0 Then
        'found a period in there. generally a dot is not allowed in function names, but
        'who knows the type of stuff allowed in Scripts, so we will check to
        'see if there is a ScriptControl with that key in our collection and if so we will return that control.
        On Error Resume Next
        Set findCtl = mRTState.Item(Left$(ProcName, InStr(ProcName, ".") - 1))
        If Not findCtl Is Nothing Then
            Set FindProcedureControl = findCtl
            ProcName = Mid$(ProcName, InStr(ProcName, ".") + 1)
        End If
        Exit Function
    End If
    For Each findCtl In mRTState
    
        On Error Resume Next
        For Each findproc In findCtl.Procedures
            If StrComp(findproc.Name, ProcName, vbTextCompare) = 0 Then
                Set FindProcedureControl = findCtl
                Exit Function
            End If
        Next
    Next

End Function
Private Sub IEvalEvents_HandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'

'FINALLY!
'a simple Eval should do it.
'- of course, first re-construct the value.
Dim Strbuild As String
Dim LoopParam As Long
Dim StopIPF As Variant
Dim LoopCtl As ScriptControl
Dim passparms As Variant
Dim PAramCount As Long
'On Error Resume Next
Set LoopCtl = FindProcedureControl(FuncName)
'InvokeDynamic the CodeObject.
'On Error Resume Next
passparms = FuncParams
'PAramCount = UBound(FuncParams) + 1
'If PAramCount = 0 Or Err <> 0 Then
'    'use Eval.
'    ReturnVal = loopctl.Eval(FuncName & "()")
'    Exit Sub
'
'End If
Set StopIPF = LoopCtl
Call InvokeDynamic(StopIPF.CodeObject, FuncName, passparms, ReturnVal)
'ReturnVal = CallByName(StopIPF.CodeObject, FuncName, VbMethod, FuncParams)
'all done.

'Stop

End Sub

Private Function IEvalEvents_HandleSubScript(withparser As CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
    '
    'TESTING: load Default.vbs from app.path.
    
    'RELEASE: load different script files from registry and INI file.
'    Dim VBScontrol As ScriptControl
'    Dim Strload As String
'    Set mRTState = New Collection
'    Set VBScontrol = New ScriptControl
'    With VBScontrol
'        .Language = "VBScript"
'        .AddCode LoadFileContents(App.Path & "\" & "Default.vbs")
'
'    End With
'    mRTState.Add VBScontrol
    Set mParser = withparser
    LoadSettingsData


End Function
Public Sub LoadSettingsData()
    'load from:
    Dim LoadFrom As String, mregistry As cRegistry
    Dim CurrCtl As ScriptControl
    Dim CurrSection As Long
    Dim CurrFile As Long
    Dim ScriptSpaces() As String, SSCount As Long
    Dim EnumFiles() As String, fCount As Long
    Dim scriptsection As String, filecontents As String
    LoadFrom = mParser.Settings.Pluginkey(mParser.ConfigSet, "BPCoreFunc.CScriptFunctions")
    'LoadFrom Section contains a section for each new script control to be created.
    'within each of those sections we will find values defining the Language, and other properties of the
    'scriptcontrol. each script file to be loaded within that scriptcontrol will be values within a "ScriptFiles" section here.
    Set mRTState = New Collection
    Set mregistry = New cRegistry
    mregistry.SectionKey = LoadFrom
    mregistry.ClassKey = HHKEY_CURRENT_USER
    CurrCtl.AddCode mInternalCode
    If Not mregistry.EnumerateSections(ScriptSpaces, SSCount) Then
        'GASP!
        Debug.Print "CScriptFunctions class could not find any Script NameSpace Sections within registry."
    Else
        'alright, we have a group of sections. Loop through each section, create a new ScriptControl
        'and populate it with the data from that section.
        For CurrSection = 1 To UBound(ScriptSpaces)
            'each Section key should have a "language" value.
            'create the script control, add it to the collection and set the language.
            'then look for scripts.
            Set CurrCtl = New ScriptControl
            'is there a language key....
            scriptsection = LoadFrom & "\" & ScriptSpaces(CurrSection)
            If mregistry.ValueEx(HHKEY_CURRENT_USER, scriptsection, "Language", RREG_SZ, "") <> "" Then
            'Language key exists? OK, then set it.
                CurrCtl.Language = mregistry.ValueEx(HHKEY_CURRENT_USER, scriptsection, "Language", RREG_SZ, "")
            Else
            
                CurrCtl.Language = "JScript"
            End If
            mRTState.Add CurrCtl, ScriptSpaces(CurrSection) 'the key is the section.
            'Add the variable for the Parser, aptly enough named "Parser".
            CurrCtl.AddObject "Parser", mParser, False
            CurrCtl.AddObject "ScriptHandler", Me, False
            
            'alright, now we load from the scriptfiles section underneath scriptsection.
            mregistry.SectionKey = scriptsection & "\scripts"
            mregistry.ClassKey = HHKEY_CURRENT_USER
            If Not mregistry.EnumerateValues(EnumFiles(), fCount) Then
                Debug.Print "No Scripts defined in key " & scriptsection
            
            
            Else
                'iterate through, add to the currctl control's code.
                'we probably should NOT crash if an error occurs, but rather abort that file.
                For CurrFile = 1 To fCount
                    'load the file contents....
                    'if the key's value is not set to "disabled...."
                    On Error Resume Next
                    filecontents = LoadFileContents(EnumFiles(CurrFile))
                    
                    CurrCtl.AddCode filecontents
                    If Err <> 0 Then
                        Debug.Print "Code Parse Error while processing file """ & EnumFiles(CurrFile) & """. " & Err.Description
                        Err.Clear
                    End If
                    filecontents = ""
                
                Next CurrFile
            
            End If
            
            
        Next CurrSection
    
    
    
    End If  'EnumerateSections


End Sub
Private Function LoadFileContents(ByVal Filename As String) As String
    Dim fNum As Long
    Dim Strbuild As String
    fNum = FreeFile
    Open Filename For Input As fNum
        Strbuild = Input$(LOF(fNum), fNum)
    Close #fNum
    LoadFileContents = Strbuild



End Function
Public Sub LoadScriptIntoModule(ByVal NameSpace As String, ByVal ModuleName As String, ByVal ScriptFile As String)
    Dim Objadd As ScriptControl
    Dim modadd As Module, loadedcontents As String
    On Error Resume Next
    loadedcontents = LoadFileContents(ScriptFile)
    If Err <> 0 Then
        Err.Raise 9, "CscriptFunctions.LoadScriptIntoModule", "Script file """ & ScriptFile & " not found."
    End If
    Set Objadd = mRTState.Item(ModuleName)
    If Err <> 0 Then
        Err.Raise 9, "CScriptFunctions.LoadScriptIntoModule", "Namespace """ & NameSpace & """ Not found."
    Else
        Set modadd = Objadd.Modules(ModuleName)
        If Err <> 0 Then
            'add it :)
            Set modadd = Objadd.Modules.Add(ModuleName)
        
        
        End If
    End If
    modadd.AddCode loadedcontents

End Sub
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

Private Property Let IEvalEvents_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get IEvalEvents_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String) As Variant
'
End Property

Private Function IEvalEvents_Term(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginTermCode
'
End Function

Private Function IEvalEvents_VariableChanged(ByVal VarName As String, ByVal OldValue As Variant, VarValue As Variant) As Boolean
'
End Function

Private Property Get IPluginUIData_Description() As String
    IPluginUIData_Description = "BASeParser XP Core WSH script delegate"
End Property

Private Function IPluginUIData_GetPluginPicture(ByVal PicSize As BASeParserXP.PluginPictureSizeConstants) As stdole.Picture
'
End Function

Private Function IPluginUIData_GetSettingsPageProgID() As String
IPluginUIData_GetSettingsPageProgID = "BPXPGUI.UCScript"
End Function
