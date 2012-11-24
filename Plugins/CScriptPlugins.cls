VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScriptPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEvalEvents
Implements ICorePlugin


'Registry Data Structure:
'The registry data structure used to hold the ScriptPlugin Information is somewhat complex and easy to forget (trust me) so
'I'll explain it here.

'First off, the root of all the data is the key:
'<SetKey(ConfigSet)>\BPCoreFunc.CScriptPlugins\
'underneath this key, each key represents a different Script Control to be created.

'+<SetKey(ConfigSet)>\BPCoreFunc.CScriptPlugins\
'+----<DescriptionN>
'       +----Modules
'              +---<ModuleNamen>
'                   +---FileKeyn
'                       <defaultvalue>=<Filename>

'the name of the key is insignificant, and rather used more for identification purposes. each one has a "Language" value that determines the, err- language to set on the scriptcontrol.
'Underneath we find a Modules key, containing none other then a key for each Module we are to add to that scriptcontrol. Each one contains a set of
'keys one per file, whose default value is a filename to load. whew.

'CScriptPlugins: Delegates plugin interface methods to installed Scripts,
'allowing Scripts to be used as actual plugins (string parsing and so forth)
Private mScriptCtls As Collection   'Collection of mScriptctls.
Private mreguse As cRegistry
Private mParser As CParser
'  One ScriptControl for each language.
'within each scriptcontrol there is a separate Module for each loaded Script file.

'also, each scriptcontrol gets an instance of this class added as a object named "Provider" which can be used to access
'settings and so forth.
Private Sub Class_Initialize()
'

End Sub

Private Sub Class_Terminate()
'
End Sub

Private Function ICorePlugin_CanHandleItem(withparser As BASeParserXP.CParser, Itemcheck As BASeParserXP.CFormItem) As Boolean
'
End Function

Private Function ICorePlugin_GetPluginUIData() As BASeParserXP.IPluginUIData
'
End Function

Private Sub ICorePlugin_HandleItem(withparser As BASeParserXP.CParser, ItemHandle As BASeParserXP.CFormItem, retval As Variant)
'
End Sub

Private Function ICorePlugin_Init(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
    '
    'initialization:
    Dim MLangSections() As String
    Dim langcount As Long, I As Long
    Dim NewCtl As ScriptControl
    Set mreguse = New cRegistry
    Set mParser = withparser
    Set mScriptCtls = New Collection
    'registry Structure:
    'HKEY_CURRENT_USER\Software\BASeCamp\BASeParser\Configuration Data\<configset>
    
    'structure underneath:(the important part)
    'A section for each language(with a language key for each)
    'Sections within each "language"- the name really doesn't matter. each one will have it's entire contents loaded into a single module.
    
    'for example:
    
    'VBScript
    '+---BASeExample
    '      +---Language="VBScript"
    '      +---FriendlyName="BASeCamp Example Script Plugin"
    '       +---Modules
    '           +----Global
    '                   +----Files
    '                          +---"D:\Programs\BPXP\Examples\Plugins\Script\BPPSCR.VBS" <this is a section, NOT a value...>
    mreguse.ClassKey = HHKEY_CURRENT_USER
    mreguse.SectionKey = "software\BASeCamp\BASeParser\Configuration Data\" & withparser.ConfigSet & "\Plugins\BPCoreFunc.CScriptPlugins"
    'enumerate the sections here
    Call mreguse.EnumerateSections(MLangSections(), langcount)
    'alright, load each section.
    For I = LBound(MLangSections) To UBound(MLangSections)
        'create a new ScriptControl....
        Set NewCtl = New ScriptControl
        'oh yeah- loadlangsection is really just loading a separate scriptcontrol for each item,
        'as such the languages could be duplicated in multiple sections without harm.
        
        If Not LoadLangSection(NewCtl, mreguse.SectionKey & "\" & MLangSections(I)) Then
            'darn, failed.
            withparser.Settings.DebugObject.Post "LoadLangSection For Section """ & MLangSections(I) & """ Failed."
            Set NewCtl = Nothing
        Else
            withparser.Settings.DebugObject.Post "LoadLangSection For Section """ & MLangSections(I) & """ Successful."
            mScriptCtls.Add NewCtl
        End If
    Next I
End Function


'Private Functions used during initialization to load the scripts and modules
'this is way more complicated then it NEEDS to be, but what the hell.


Private Function LoadLangSection(ToControl As ScriptControl, regSection As String) As Boolean
    Dim ModulesEnum() As String, modcount As Long, ModLoop As Long
     '
     'use mreguse to retrieve the data.
     'first, we need a language for the scriptcontrol.
     LoadLangSection = True
     'ToControl.Reset
     On Error Resume Next
     'ToControl.Language = Mid$(regSection, InStrRev(regSection, "\"))
     ToControl.Language = mreguse.ValueEx(HHKEY_CURRENT_USER, regSection, "Language", RREG_SZ, "")
     
     If ToControl.Language = "" Then
        
            'damn, not specified...
            mParser.Settings.DebugObject.Post "Section """ & regSection & """ does not contain a valid Language item, attempting to use section name as language...", Severity_Warning
            ToControl.Language = Mid$(regSection, InStrRev(regSection, "\"))
            If Err <> 0 Or ToControl.Language = "" Then
                'oh well, hey, we tried...
                'future versions may exhibit more exhaustive attempts to determine the language...
                mParser.Settings.DebugObject.Post "Section """ & regSection & """ being ignored. Language not specified or language specified was invalid. Script Control returns """ & ToControl.Error.Description & """."
                LoadLangSection = False
                Exit Function
            End If
         
     End If
     'at this point the language checks out. now we need to load the code into the scriptcontrol.
     'we'll find this much needed information in <regsection>\Modules
     '(later also add a Dirs key, but that is kind of silly, because it might try to load executables or something.
     
    'load the set of files that comprise each module.
    'kind of silly I think, because most sections underneath the plugin key will likely just be separate entities, rather
    'then being separated into several files.
    'Oh well, I guess it could be useful for including generic modules of non-plugin related script code, like public routines...
    
     
     mreguse.SectionKey = regSection & "\" & "Modules"
     If mreguse.EnumerateSections(ModulesEnum(), modcount) Then
        For ModLoop = LBound(ModulesEnum) To UBound(ModulesEnum)
                If LoadScriptModule(ToControl, regSection & "\Modules\", ModulesEnum(ModLoop)) Then
                '
                'yippee.
                Debug.Print "LoadScriptModule success!"
                End If
     
     
        Next ModLoop
     
     End If
     
End Function
Private Function LoadScriptModule(ToCtl As ScriptControl, ByVal RegBase As String, ByVal ModuleName As String) As Boolean
Dim FilesLoad() As String, NumFiles As Long
'
'LoadScriptModule: loads a Module into the given script control
'ModuleName is used to access the registry.
Dim AddFiles() As String, filecount As Long
Dim mdlAddCode As MSScriptControl.Module, I As Long
Dim CurrFilename As String
Dim sectionbase As String
'steps:
'the key regbase & modulename will contain a SECTION, yes, NOT a value, a section, for each Script file we are to load into the specified module of ToCtl.
'why a section? For later use as configuration I suppose.

'of course, we add the module to the object if necessary.
'Special cases:
'one special case exists so far: if the specified filename starts with "FILELIST="
'then the file after the equals sign is to be treated as a list of files we are to add.
'(this feature is actually being considered, but I think it would be better used on a
'higher level, that is, an INI file that encompasses all the different modules, layed
'out like a standard INI file or some such.
On Error Resume Next
LoadScriptModule = True
Set mdlAddCode = ToCtl.Modules.Item(ModuleName)
If Err <> 0 Then
    Set mdlAddCode = ToCtl.Modules.Add(ModuleName)
End If
mreguse.SectionKey = RegBase & ModuleName
'enumerate the files listed under this section.
'remember, they are sections, not values...
If mreguse.EnumerateSections(AddFiles(), filecount) Then
    'alrighty, iterate through and add the code within each file to the specified module.
    sectionbase = mreguse.SectionKey & "\"
    For I = 1 To UBound(AddFiles)
        'load each file into the module. Oh hey, look, theres a routine aptly named LoadFileIntoModule!
        'CurrFilename = mreguse.ValueEx(HHKEY_CURRENT_USER, mreguse.SectionKey & "\" & AddFiles(I), "", RREG_SZ, "")
        mreguse.SectionKey = sectionbase & AddFiles(I)
        CurrFilename = mreguse.Value
        If mreguse.Value <> "" Then
            If LoadFileIntoModule(ToCtl, CurrFilename, ModuleName) Then
                'success.
                Debug.Print "LoadFileIntoModule success. """ & CurrFilename & """ loaded into module """ & ModuleName & """"
            Else
                Debug.Print "Failed to load file """ & CurrFilename & """ into module """ & ModuleName & """"
                LoadScriptModule = False
            End If
        End If
    
    Next I

End If



End Function

Private Function LoadFileIntoModule(InCtl As ScriptControl, ByVal Filename As String, Optional ByVal ModuleName As String = "Global") As Boolean
    'Loads a Script file (ignoring language, assumes language has been set properly) into the given module of the passed scriptcontrol
    'we can't really load the code in chunks, so we have to grab the whole file all at once
    'shouldn't really be a problem though, I don't see anybody having a 2GB VBScript file, much less one for this plugin.
    Dim filecontents As String, fNum As Integer
    On Error GoTo FileError
    LoadFileIntoModule = True
    Open Filename For Input As fNum
        filecontents = Input$(fNum, LOF(fNum))
        'now, just addcode.
        On Error Resume Next
        InCtl.Modules.Item(ModuleName).AddCode filecontents
        If Err <> 0 Then
            Debug.Print "Error during AddCode()"
            Debug.Assert False
            LoadFileIntoModule = False
            
        End If
    Close fNum
    Exit Function
FileError:
    Debug.Assert False
    Resume
    LoadFileIntoModule = False
End Function
Private Function CallOnScripts(ByVal FuncName As String, ByVal UseReturn As Boolean, ParamArray Arguments())

'Generic wrapper used un all Interface routines. Prevents me from writing the same
'god damned loop code in all of them...
'iterate through each scriptcontrol and attempt to call the routine
'the return value will either be ignored or used to terminate the loop. which one happens
'depends on the UseReturn boolean parameter.
Dim LoopCtl As ScriptControl, loopmdl As Module, rval As Variant, Args() As Variant
For Each LoopCtl In mScriptCtls
    'also iterate through each module.
    Args = Arguments
    For Each loopmdl In LoopCtl
        'call the procedure.
        'any error in the call will simply cause us to move to the next iteration.
        On Error GoTo NextLoop
        'the "codeobject" contains a method for each routine. we shall use that paired with callbyname.
        InvokeDynamic loopmdl.CodeObject, FuncName, Args(), rval
        'if usereturn was passed, then the caller wants to terminate the loop at the first one that
        'returns a value. so do so...
        If UseReturn Then
            If Not IsObject(rval) Then
                If rval = 0 Then
                    
                Else
                    If IsObject(rval) Then
                        Set CallOnScripts = rval
                    Else
                        CallOnScripts = rval
                    End If
                    Exit Function
                End If
            End If
        End If
        
        
NextLoop:
    Next loopmdl
        'remove the error handler. don't want to jump inside the loop. freaky crap could happen.
        On Error GoTo 0

Next LoopCtl


'
End Function
'the script implementation adds a new parameter to the front of the list, a reference to This class.
'it doesn't have them yet, but I'll add "helper" routines.
Private Function ICorePlugin_ParseLocation(withparser As BASeParserXP.CParser, ByVal StrParse As String, Position As Long, currItem As BASeParserXP.CFormItem) As BASeParserXP.ParseLocationReturnConstants
'
Dim objpass As Object
Set objpass = withparser
'ICorePlugin_ParseLocation = CallOnScripts("ParseLocation", True, Me, objpass, StrParse, Position, currItem)
End Function

Private Property Let ICorePlugin_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String, ByVal RHS As Variant)
'
End Property

Private Property Get ICorePlugin_Settings(withparser As BASeParserXP.CParser, ByVal PropName As String) As Variant
'
End Property

Private Function ICorePlugin_Term(FromParser As BASeParserXP.CParser) As BASeParserXP.IPluginTermCode
'
'CallOnScripts "Term", True, Me, FromParser
End Function

Private Sub IEvalEvents_AfterResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, VarValue As Variant)
'
'CallOnScripts "AfterResolveVariable", False, Me, ParserUse, VarName, VarValue
End Sub

Private Function IEvalEvents_AssignSubScript(withparser As BASeParserXP.CParser, OnValue As Variant, subscripts As Variant, assignthis As Variant) As Boolean
'
'IEvalEvents_AssignSubScript = CallOnScripts("AssignSubscript", True, Me, Withparser, OnValue, subscripts, assignthis)
End Function

Private Function IEvalEvents_BeforeFunctionCall(withparser As CParser, FuncCalling As String, WithHandler As BASeParserXP.IEvalEvents, Argarray As Variant, ForceResult As Variant) As Boolean
'
End Function

Private Function IEvalEvents_BeforeResolveVariable(ParserUse As BASeParserXP.CParser, ByVal VarName As String, cancel As Boolean) As Variant
'
'IEvalEvents_BeforeResolveVariable = CallOnScripts("BeforeResolveVariable", True, Me, ParserUse, VarName, cancel)
End Function

Private Function IEvalEvents_CanHandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, NoParsedArgs() As Boolean) As Boolean
'this one might need boxing/unboxing code for the boolean array. sigh.


'IEvalEvents_CanHandleFunction = CallOnScripts("CanHandleFunction", True, Me, Withparser, FuncName, NoParsedArgs())

End Function

Private Function IEvalEvents_CanHandleOperator(withparser As BASeParserXP.CParser, ByVal OpCheck As String, Unary As BASeParserXP.IEvalUnaryOpFlags) As Boolean
'

End Function

Private Sub IEvalEvents_ExecuteComplete(ParserDone As BASeParserXP.CParser)
'
End Sub

Private Function IEvalEvents_GetFunctionFlags(withparser As BASeParserXP.CParser, ByVal FuncString As String) As BASeParserXP.FunctionFlags
'
End Function

Private Function IEvalEvents_GetHandledFunctionString(withparser As BASeParserXP.CParser) As String
'
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
'
End Sub

Private Sub IEvalEvents_HandleFunction(withparser As BASeParserXP.CParser, ByVal FuncName As String, FuncParams() As Variant, ReturnVal As Variant)
'
End Sub

Private Function IEvalEvents_HandleSubScript(withparser As BASeParserXP.CParser, OnValue As Variant, subscripts As Variant, retval As Variant) As Boolean
'
End Function

Private Function IEvalEvents_Init(withparser As BASeParserXP.CParser) As BASeParserXP.IPluginInitCode
''
End Function

Private Property Get IEvalEvents_Name() As String
IEvalEvents_Name = "BASeCamp CoreScript Plugin"
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
