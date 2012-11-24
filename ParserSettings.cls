VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009
'Will add a function to query data from there.
'simple parse, compared to the Evaluator.
Private Const AppKey As String = "Software\BASeCamp\BASeParser\"
Private mregistry As cRegistry
'Private Const INILocation = "%APPDATA%\BASeParser\<configset>.ini"

'this class was responsible for my creation of the BCSEttings Library.

'Registry Storage Structure:
'all... well, almost all, of BASeParser XP's Settings are stored in the registry.
'The first unit of study to focus on is the "Configset" a Configset is simply a subkey of
'the key
'HKEY_CURRENT_USER\Software\BASeCamp\BASeParser\Configuration Data
'the key's name is also the name of the configuration set.
'within the configuration set key, we find a few values that apply to that
'set as a whole:
'MinCacheSize,MaxQueueSizeBeforePurge, etc.


'Parser Plugin Manager.
'This adds support for different sets of plugin's for different applications,
'by loading the appropriate registry values, as well as saving them, to the specified location.
'This class will then add the appropriate objects to the passed CParser object.
Public Enum CleanRegistryKeyConstants
    Clean_user = 2
    clean_machine = 4
    CLEAN_BOTH = 6
End Enum
Public Enum PluginTypeConstants
    PluginType_Invalid = -50
    PluginType_Evaluation = 0
    PluginType_Core = 1
    PluginType_Both = 2

End Enum
Public Enum VerifyPluginErrorConstants
    Verify_CreateFailed = ParserErrorBase + 1024 + 1
    Verify_CastFailed = ParserErrorBase + 1024 + 2
End Enum
Public Type ParserConfigSet
    Setname As String
    Optimize As Boolean
    MinCacheSize As Long    'minimum length of the Stack before it is
                            'considered for placing onto the Queue of cached
                            'stacks. Ha!
    DisableCore As Boolean
    MaxQueueSizeBeforePurge As Long     'a long name to describe the maximum number of
                                        'previously built Parse Stacks that
                                        'are Cached. After this number is reached, the older ones are
                                        'discarded. HOWEVER- if they are accessed, they are moved to the front of the queue, so...
    'AllowImpliedDeclaration As Boolean    'Allow/Disallow Implicit Variable Creation
    DisplayInTray As Boolean    'should we display in the tray? Will use VBAccelerator CPopMenu DLL.
    
    mOtherEntries As Collection     'Collection of other entries.
                                    'ideally this should be by itself, with clients, for example,
                                    'retrieving the optimization status via a call to this collection.
                                    'that still works, but  we waste memory by also
                                    'storing it separately in the structure.
    PluginCount As Long
    progIDs() As String
   
End Type
Public Type ParserConfigSets
    ConfigCount As Long
    configsets() As ParserConfigSet
End Type
'Private mBaseINI As CINIData
'Private mINIObjects As Collection 'collection of cINIData Objects, referenced by their key, which is also their configset.

Public Enum PluginCreationErrors
    PCE_CREATEFAILED = 431
    PCE_INTERFACENOTFOUND = 432
End Enum
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Const CSIDL_APPDATA As Long = &H1A

'unused enumeration. Well, unused so far.
Public Enum UnSeparatedParenthesesBehaviour
    Paren_Insert_Mult
    Paren_Default_Object
End Enum
'Friend Function GetSetINIFile(ByVal StrSetName As String) As String
'    Dim StrAppdata As String
'    StrAppdata = GetAppData()
'    If right$(StrAppdata, 1) <> "\" Then StrAppdata = StrAppdata & "\"
'    StrAppdata = StrAppdata & "BASeParser XP\configsets\" & StrSetName & ".ini"
'
'    GetSetINIFile = StrAppdata
'
'
'
'
'End Function
'Friend Function GetSetINIObject(ByVal SetName As String) As CINIData
'    Dim NewObject As CINIData
'    Dim INIFile As String
'
'    On Error Resume Next
'    Set GetSetINIObject = mINIObjects.Item(SetName)
'    If Err <> 0 Then
'
'
'        INIFile = GetSetINIFile(SetName)
'        Set NewObject = New CINIData
'        NewObject.LoadINI INIFile
'        mINIObjects.Add NewObject, SetName
'    End If
'End Function
'Private Function GetAppData() As String
'    Dim Strret As String
'    Strret = Space$(255)
'    Call SHGetSpecialFolderPath(0, Strret, CSIDL_APPDATA, 1)
'    GetAppData = Trim$(Replace$(Strret, vbNullChar, ""))
'
'End Function
'Private Function GetCoreINIFileName() As String
'    Dim StrAppdata As String
'    StrAppdata = GetAppData()
'    If right$(StrAppdata, 1) <> "\" Then StrAppdata = StrAppdata & "\"
'    StrAppdata = StrAppdata & "BASeParser XP\core.ini"
'End Function

Public Property Get DefaultBracketBegin() As String
    DefaultBracketBegin = DefaultBracketStart
End Property
Public Property Get DefaultBracketStop() As String
    DefaultBracketStop = DefaultBracketEnd
End Property

Public Property Get RegObject() As Object
    Set RegObject = mregistry
    
End Property
Public Property Get AllowImpliedDeclaration(ByVal Setname As String) As Boolean
    AllowImpliedDeclaration = CBool(mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "AllowImpliedDeclaration", RREG_SZ, "1"))
    'AllowImpliedDeclaration = GetSetINIObject(SetName).ReadProfileSetting("Core", "AllowImpliedDeclaration", "True")
End Property
Public Property Get ErrorOnSurplusArrayIndices(ByVal Setname As String) As Long

    ErrorOnSurplusArrayIndices = mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "ErrorOnSurplusArrayIndices", RREG_DWORD, 1)
    'ErrorOnSurplusArrayIndices = GetSetINIObject(SetName).ReadProfileSetting("Core", "ErrorOnSurplusArrayIndices", 1)
    


End Property
Public Property Let ErrorOnSurplusArrayIndices(ByVal Setname As String, ByVal Vdata As Long)

    mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "ErrorOnSurplusArrayIndices", RREG_DWORD, 1) = Vdata
    'GetSetINIObject(SetName).WriteProfileSetting "Core", "ErrorOnSurplusArrayIndices", Vdata
    
End Property
Public Property Let CalloutputMethods(ByVal Setname As String, ByVal Vdata As Boolean)


    mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "CalloutputMethods", RREG_DWORD, 1) = Vdata


End Property
Public Property Get CalloutputMethods(ByVal Setname As String) As Boolean
    CalloutputMethods = mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "CalloutputMethods", RREG_DWORD, 1)
End Property

Public Property Get CorePluginsDisabled(ByVal Setname As String) As Boolean
    'returns wether Core Plugins are disabled.
    CorePluginsDisabled = mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Plugins\", "DisableCorePlugins", RREG_DWORD, 0)
    'CorePluginsDisabled = GetSetINIObject(SetName).ReadProfileSetting("Core", "DisableCorePlugins", "0")
End Property
Public Property Let CorePluginsDisabled(ByVal Setname As String, ByVal Vdata As Boolean)
    mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Plugins\", "DisableCorePlugins", RREG_DWORD, 0) = Vdata
    'Call GetSetINIObject(SetName).WriteProfileSetting("Core", "DisableCorePlugins", Vdata)
End Property
Public Function GetConfigSetNames(Optional ByRef retcount As Long) As String()
Dim configsets() As String
mregistry.ClassKey = HHKEY_CURRENT_USER
mregistry.SectionKey = AppKey & "Configuration Data"
Call mregistry.EnumerateSections(configsets(), retcount)
GetConfigSetNames = configsets

End Function



Public Function GetConfigSetsData() As ParserConfigSets
'
Dim configsets() As String, sectcount As Long, I As Long
Dim retSets As ParserConfigSets
mregistry.ClassKey = HHKEY_CURRENT_USER
mregistry.SectionKey = AppKey & "Configuration Data"
Call mregistry.EnumerateSections(configsets(), retSets.ConfigCount)
ReDim retSets.configsets(1 To retSets.ConfigCount)

For I = 1 To retSets.ConfigCount
    
    retSets.configsets(I) = GetConfigSetData(configsets(I))
    

Next I
    GetConfigSetsData = retSets

End Function
Public Sub SaveConfigSetsData(SetsSave As ParserConfigSets)
'this one is really evil, it totally wipes the entire
'Configuration key.
'OK. I lied. it doesn't.
Dim I As Long
With mregistry
    .ClassKey = HHKEY_CURRENT_USER
    .SectionKey = AppKey & "Configuration Data"
    '.DeleteKey
    'ha, I lied again! HA HA! It does delete it. WAAAH! DIE! DIE!
End With
For I = 1 To SetsSave.ConfigCount
    SaveConfigSetData SetsSave.configsets(I)
Next I
'all done.
End Sub
Public Function GetConfigSetData(ByVal configset As String) As ParserConfigSet
    Dim returnthis As ParserConfigSet
    Dim skey As String
    Dim I As Long
    Dim sValueNames() As String, ValueCount As Long
    '"Configuration Data\" & SetName & "\Plugins"
    'has the plugins as SECTIONS.
    skey = SetKey(configset)
    With returnthis
        .Setname = configset
        .progIDs = Me.GetPluginProgIDs(configset, .PluginCount, True)
        .Optimize = CBool(mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "AllowOptimize", RREG_DWORD, True))
        .MinCacheSize = mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "MinCacheSize", RREG_DWORD, 4)
        .MaxQueueSizeBeforePurge = mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "MaxQueueSizeBeforePurge", RREG_DWORD, 32)
        .DisplayInTray = CBool(mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "DisplayInTray", RREG_SZ, "False"))
        'LASTLY: enumerate all the values, creating a new item in the Collection with the key set to the Registry value Name, and, (Duh) the value
        'set to the value:
        mregistry.SectionKey = skey
        Call mregistry.EnumerateValues(sValueNames(), ValueCount)
        Set .mOtherEntries = New Collection
        For I = 1 To ValueCount
            'append "CFG:" to the start, so that we are sure that
            'the value doesn't start with a number.
            .mOtherEntries.Add mregistry.ValueEx(HHKEY_CURRENT_USER, skey, sValueNames(I), RREG_SZ, ""), "CFG:" & sValueNames(I)
            
        
        Next I
        
        
        
        
    End With
    
    GetConfigSetData = returnthis
    



End Function
Public Function GetBPSetting(FromSet As ParserConfigSet, ByVal SettingName As String) As String

    On Error GoTo RETURNEMPTY
    GetBPSetting = FromSet.mOtherEntries.Item("CFG:" & SettingName)
    Exit Function
RETURNEMPTY:
    GetBPSetting = ""



End Function
Public Sub SaveConfigSetData(ConfigSetData As ParserConfigSet)
    'first- DELETE the entire section. MWA HA HA! DIE!
    Dim skey As String
    With mregistry
    skey = SetKey(ConfigSetData.Setname)
    .ClassKey = HHKEY_CURRENT_USER
    '.SectionKey = AppKey & "\Configuration Data\" & ConfigSetData.Setname
    
    .SectionKey = skey
    '.DeleteKey
    'MWA HA HA! DIE! DIE! you WRETCHED KEYS!
    'wow, this is better then Quake.
    'now, save it back.
    With ConfigSetData
        mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "AllowOptimize", RREG_DWORD, True) = .Optimize
        SavePluginProgIds .progIDs(), .Setname
        mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "MaxQueueSizeBeforePurge", RREG_DWORD, 32) = .MaxQueueSizeBeforePurge
        mregistry.ValueEx(HHKEY_CURRENT_USER, skey, "MinCacheSize", RREG_DWORD, 4) = .MinCacheSize
        
    
    
    End With
    
    End With
End Sub
Public Sub DeinitializeParser(ParseObj As CParser, ByVal Setname As String)
    'deinitializes the parser.
    'this is a hell of a lot trickier then going the other way-
    'we want to save the progIDs listed.
    'we'll need our friend TLI for that.
    'Oh, and we don't save at all for "Default"
    If Setname = "Default" Then Exit Sub
    

End Sub
Public Function VerifyPlugin(ByVal progID As String) As Boolean
    'Verifies that the given ProgID is, indeed, a plugin.
    Dim testCreate As Object
    Dim TestCast As IEvalEvents
    On Error Resume Next
    CDebug.Post "verifying plugin " & progID & "..."
    Set testCreate = createObject(progID)
    If Err <> 0 Then
        VerifyPlugin = False
        Err.Raise (Verify_CreateFailed), "ParserSettings.VerifyPlugin", "Plugin Verification Failure: Could not create " & progID
        
    End If
    Set TestCast = testCreate
    If Err <> 0 Then
    VerifyPlugin = False
    Err.Raise (Verify_CastFailed), "ParserSettings>VerifyPlugin", "Plugin Verification Failed: Could not cast " & progID & " To IEvalEvents Interface."
    
    
    End If
    VerifyPlugin = True

    
End Function
Private Function CheckPatch(ByVal PluginProgID As String, Optional ByVal configset As String = "Default") As String
    'if the given ProgID has a patch, returns that progID, after verifying:
    '1. It exists
    '2. It supports IEvalEvents
    Dim PatchData As String
    'otherwise, simply returns the passed in ProgID.
    
    'Patches are stored in Appkey & "PluginPatches"
    PatchData = mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(configset) & "PluginPatches\", PluginProgID, RREG_SZ, "")
    On Error Resume Next
    If PatchData <> "" Then
        If VerifyPlugin(PatchData) Then
            CheckPatch = PatchData
            
        Else
            CheckPatch = PluginProgID
        End If
    Else
            CheckPatch = PluginProgID
    End If
    
    
End Function

Private Sub InitCore(InitThis As CParser, ByVal Setname As String)
    'Add the Core Plugins- but, first,
    'make sure they haven't been "Patched" in the registry.
    'this architecture means that if I, or someone else, finds a bug
    'in BASeParser that is ONLY the plugin, I can- err, somehow patch it
    'and make the exact functionality outside the dll. BPCoreOpFunc is pretty dependent on
    'ModParser, though- I imagine this
    'could be used if somebody wanted to write the operators/Functions that conform to
    'their native programming language...
    'Anyway-
    Dim newop As IEvalEvents, newcore As ICorePlugin
    Dim CoreProgID As String, FHID As String
    'CoreProgID = "BASeParserXP.BPCoreOpFunc"
    'FHID = "BASeParserXP.FunctionHandler"
    'CoreProgID = CheckPatch(CoreProgID)
    'FHID = CheckPatch(FHID)
    
    
    Set newop = New BPCoreOpFunc
    'Set newop = CreateObject(CoreProgID)
    InitThis.AddEventSink newop
    'newop.Init InitThis
    
    Set newop = New FunctionHandler
    'Set newop = CreateObject(FHID)
    InitThis.AddEventSink newop
    'newop.Init ParseObj
    
    'also, add the core Coreplugins
    'those that implement ICorePlugin.
    'right now, a hook that detects Environment variable references.
    Set newcore = New CPlugEnvString
    InitThis.CorePlugins.Add newcore
    newcore.Init InitThis
    Set newcore = New CSet
        'use this CSet as a plugin. otherwise it's data structures are unused anyway.
    InitThis.CorePlugins.Add newcore
    newcore.Init InitThis
End Sub
Public Sub InitializeParser(ParseObj As CParser, ByVal Setname As String, Optional ByVal IgnoreKeystateswitches As Boolean = False)
    'Initializes the parser with the settings from the given Set.
    'ALL sets add the Core parsers.
    'right now, there is no way around it.
    'I'll probably add a way, however.
    
    'New: allow holding of Shift/Ctrl/Alt keys to modify default behaviour when loading plugins.
    
    Dim mPlugins() As String, pcount As Long
    Dim assignIt As IEvalEventSinks, I As Long
    Dim newop As IEvalEvents
    'Dim OnlyDefault As Boolean
    Dim EvalError As Integer, CoreError As Integer
    Dim AddThem As Object, CastThem As IEvalEvents
    Dim castCore As ICorePlugin
    On Error GoTo InitFailure
    
    'If GetAsyncKeyState(vbKeyShift + vbKeyControl) <> 0 Then
    '    OnlyDefault = True
    'End If
    
    With ParseObj
    
    Set assignIt = New IEvalEventSinks
    Set .Variables = New CVariables
    Set .EvalListener = assignIt
    'initialize it's functions collection.
    Set .Functions = New CFunctions
    Set .Functions.Parser = ParseObj
    Set .CorePlugins = New ICorePlugins
    '.EvalListener. Me
    On Error Resume Next
    .EvalListener.Self.Init ParseObj
    On Error GoTo InitFailure
    'change: FunctionHandler will be AFTER the core operators/Functions.
    'why? speed. If the FunctionHandler were called for every function first, then it
    'would waste time first.
    .SetInitialized
    'I'm surprised this architecture is actually working the way I intended.
    'InitCore ParseObj, Setname
    'If CheckPatch("BASeParserXP.BPCoreOpFunc")
    
    End With
    'OK- Load the Plugins, specified for this set.
    mPlugins = GetPluginProgIDs(Setname, pcount, True)
    CDebug.Post "Loading and Initializing Plugins (found " & pcount & ")..."
    If pcount Then
        For I = 1 To UBound(mPlugins)
            If InStr(mPlugins(I), vbNullChar) = 0 Then  'if a nullchar exists, the verification went awry.
                   'good, load the object.
                   'First check if it is disabled, though.
                   If Not IsPluginDisabled(mPlugins(I), Setname) Then
                        If Not ParseObj.PluginLoaded(mPlugins(I)) Then
                            Set AddThem = createObject(mPlugins(I))
                            On Error Resume Next
                            Set CastThem = AddThem
                            EvalError = Err.Number
                            'If Err <> 0 Then
                                 On Error Resume Next
                                 Set castCore = AddThem
                                CoreError = Err.Number
                                If CoreError = 0 Then
                                    ParseObj.CorePlugins.Add castCore
                                    'castCore.Init ParseObj
                                    CDebug.Post "Successfully loaded " & mPlugins(I) & " as a Core Plugin."
                                Else
                                    'failed.
                                    CDebug.Post "Failed to load """ & mPlugins(I) & " as a Core Plugin.(Interface not supported)", Severity_Warning
                                
                                End If
                            
                                'Good, now add the implementor.
                                If EvalError = 0 Then
                                    ParseObj.EvalListener.AddImplementor CastThem
                                    'Also- Fire the "Init()" method on it too.
                                    CastThem.Init ParseObj
                                Else
                                    EvalError = 0
                                End If
                            
                        Else
                            'Already loaded.
                            CDebug.Post "Did not load Plugin, """ & mPlugins(I) & """. The plugin is Already loaded."
                        End If
                Else
                    If StrComp(mPlugins(I), "BASeParserXP.BPCoreOpFunc", vbTextCompare) = 0 Then
                        CDebug.Post "Core operator/Function Handler plugin is DISABLED!", Severity_Warning
                    
                    End If
                    CDebug.Post "Did not load Plugin, """ & mPlugins(I) & """. The plugin is disabled."
                End If
            Else
             'this plugin failed the verification.
                CDebug.Post left$(mPlugins(I), InStr(mPlugins(I), vbNullChar)) & " failed verification, Code (" & Mid$(mPlugins(I), InStr(mPlugins(I), vbNullChar) + 1) & ")"
            
            End If
        Next
    End If
    'mPlugins = GetParsePluginProgIDs(Setname, PCount, True)
    
    
    
    'lastly, we want to load the functions from the registry.
    'they will be enumerated as String values under the configuration set.
    '!!!moved into FunctionHandler.
    'InitFunctions ParseObj, Setname
    Exit Sub
InitFailure:
    'this is probably caused by some code change I made somewhere.
    'generally, it shouldn't happen without
    'a code change...
    CDebug.Post "Initialization error " & Error$
    Debug.Assert False
    Resume
End Sub

Public Function SetKey(Optional ByVal ForSet As String = "Default") As String
SetKey = AppKey & "Configuration Data\" & ForSet & "\"
End Function
Public Sub SavePluginProgIds(ProgIdsSave() As String, Optional ByVal Setname As String = "Default")
    Dim I As Long
    With mregistry
    .ClassKey = HHKEY_CURRENT_USER
    .SectionKey = SetKey(Setname) & "\Plugins"
    '.DeleteKey
    For I = LBound(ProgIdsSave) To UBound(ProgIdsSave)
    
        .ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "\Plugins\" & ProgIdsSave(I), "", RREG_SZ, "") = vbNullChar
        
    
    Next I
    'tada.
    



    End With
End Sub

Public Function GetPluginProgIDs(Optional ByVal Setname As String = "Default", Optional ByRef PluginCount As Long, Optional ByVal Verify As Boolean = True, Optional ByRef PluginTypes As Variant) As String()
    'loads plugin data, for a given configuration set.
    Dim PlugIDs() As String, pcount As Long
    Dim RetIDs() As String
    Dim I As Long, startingsize As Long
    Dim AttemptCreate As Object, AttemptCast As IEvalEvents, AttemptCoreCast As ICorePlugin
    ReDim PluginTypes(0)
    mregistry.ClassKey = HHKEY_CURRENT_USER
    mregistry.SectionKey = SetKey(Setname) & "Plugins\"
    
    '### Core Plugins automagically added here-
    'why? Just in case BASeParser XP has never run on this computer, and thus
    'the registry has no entries for us to load. Of course, these can be disabled, but it is assumed that
    'in such a case whoever caused it knows what they are doing (replacing the Core Operator/Function handler with their own,
    'probably)
    Const InitialCount = 5
        ReDim RetIDs(1 To 5)
        RetIDs(1) = "BASeParserXP.BPCoreOpFunc"
        RetIDs(2) = "BASeParserXP.FunctionHandler"
        RetIDs(3) = "BASeParserXP.CPlugEnvString"
        RetIDs(4) = "BASeParserXP.CSet"
        RetIDs(5) = "BASeParserXP.CPlugBackTicks"

        'add to this one, our plugins we got from the dir.
        If mPluginDirContentCount > 0 Then
            
            ReDim Preserve RetIDs(1 To InitialCount + mPluginDirContentCount)
            For I = 1 To mPluginDirContentCount
            CDebug.Post "adding Directory enumerated Plugin, """ & mPluginDirContents(I) & """"
                RetIDs(InitialCount + I) = mPluginDirContents(I)
            
            Next I
        
        
        End If
        
        
        Call mregistry.EnumerateSections(PlugIDs, pcount)
    'If mregistry.EnumerateSections(PlugIDs, pcount) Then
    startingsize = UBound(RetIDs)
    If pcount And Verify Then
        'make sure they exist.
        'two possible problems- invalid progIDs, or
        'the object doesn't support the interface.
        'in both, we still return that string, we just change it so that we can notify of errors.
        'With ProgID<null>ERRCODE
        ReDim PluginTypes(1 To pcount)
        On Error Resume Next
        
        For I = 1 To UBound(PlugIDs)
        'remove null chars.
        If Err <> 0 Then Exit For   'when no entries
        PluginTypes(I) = PluginType_Evaluation
        PlugIDs(I) = Replace$(PlugIDs(I), vbNullChar, "")
            'attempt to create it.
            On Error Resume Next
            Set AttemptCreate = createObject(PlugIDs(I))
            If Err <> 0 Then
                'error during create
                PlugIDs(I) = PlugIDs(I) & vbNullChar & Err.Number
            Else
                'successful creation.
                'So, we will now
                'try to cast it.
                Set AttemptCast = AttemptCreate
                If Err <> 0 Then
                    'Oh no, it doesn't support the interface.
                    'PCE_INTERFACENOTFOUND
                    'first, see if it is a "Core" plugin.
                    On Error Resume Next
                    
                    Set AttemptCoreCast = AttemptCreate
                    If Err = 0 Then
                        PluginTypes(I) = PluginType_Core
                    
                    Else
                        PlugIDs(I) = PlugIDs(I) & vbNullChar & Err.Number
                    End If
                End If
            
            End If
            
            ReDim Preserve RetIDs(1 To I + startingsize)
            RetIDs(I + startingsize) = PlugIDs(I)
        Next I
    
    Else
        On Error Resume Next
        For I = 1 To UBound(PlugIDs)
        If Err <> 0 Then Exit For
           ReDim Preserve RetIDs(1 To I + startingsize)
               RetIDs(I + startingsize) = PlugIDs(I)
        Next I
    
    End If
    
    
    
    
    PluginCount = pcount + startingsize
    GetPluginProgIDs = RetIDs
End Function
Public Property Get Optimize(Optional ByVal Setname As String = "Default") As Boolean
'returns wether optimization has been disabled (false) or is enabled (true)
Dim optret As String
optret = mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Optimizations\LocalCache", "Allow", RREG_SZ, "YES")
Select Case UCase$(optret)
    Case "YES", "YEP", "WHY NOT?", "TRUE", "POSITIVE", "AFFIRMATIVE"
        Optimize = True
        Exit Property
End Select

End Property

Public Property Get CacheOptimizeBoundary(Optional ByVal Setname As String = "Default") As Integer
    'returns wether optimization has been disabled (false) or is enabled (true)

    CacheOptimizeBoundary = mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Optimizations\LocalCache", "MINTOKENENABLE", RREG_DWORD, 4)

End Property
Private Function BuildKey(ObjPlugin As Variant, ByVal SettingName As String, ByVal configset As String) As String
    Dim PluginProgID As String, keyuse As String
    Dim cast As Object, appendsection As String, valueget
    If IsObject(ObjPlugin) Then Set cast = ObjPlugin
    If IsObject(ObjPlugin) Then PluginProgID = GetObjectProgID(cast) Else PluginProgID = ObjPlugin
    keyuse = PluginKey(configset, PluginProgID)
    SettingName = Replace$(SettingName, "/", "\")
    If InStr(SettingName, "\") <> 0 Then
        appendsection = Mid$(SettingName, 1, InStrRev(SettingName, "\"))
        valueget = Mid$(SettingName, InStrRev(SettingName, "\"))
        If left$(appendsection, 1) = "\" Then appendsection = Mid$(appendsection, 2)
    End If
    BuildKey = keyuse & appendsection
End Function
Public Function GetPluginSetting(ByVal ObjPlugin As Object, ByVal SettingName As String, Optional ByVal Setname As String = "Default", Optional ByVal DefaultValue As String = "") As Variant
    'retrieves a setting for a plugin.
    'First, we need to build the Section name and then the value name.
    'what? you say- shouldn't the section simply be
    'SetKey(SetName) & "Plugins\" & PluginProgID?
    'well, a basic implementation of this routine, maybe.
    'no- this one parses the passed Setting name for slashes, and tacks it onto the end of the key.
    
    Dim PluginProgID As String
    Dim keyuse As String, appendsection As String
    Dim valueget As String
    keyuse = BuildKey(ObjPlugin, SettingName, Setname)
    With mregistry
        GetPluginSetting = .ValueEx(HHKEY_CURRENT_USER, keyuse, SettingName, RREG_SZ, DefaultValue)
    
    End With
    


End Function
Public Sub SavePluginSetting(ByVal ObjPlugin As Object, ByVal SettingName As String, ByVal Vdata As String, Optional ByVal Setname As String = "Default")
'saves the plugin setting.
'generic Save routine, same as GetPluginSetting. That is, the actual plugin has no "voice" in the matter, and really will only see the
'settings if it uses the GetPluginSetting method to retrieve them...


Dim PluginProgID As String
    Dim keyuse As String, appendsection As String
    Dim valueget As String
    keyuse = BuildKey(ObjPlugin, SettingName, Setname)
    With mregistry
        'GetPluginSetting = .ValueEx(HHKEY_CURRENT_USER, keyuse, SettingName, RREG_SZ, DefaultValue)
        .ValueEx(HHKEY_CURRENT_USER, keyuse, SettingName, RREG_SZ, "") = Vdata
    
    End With
    

End Sub


Public Function GetPluginSettingEx(withparser As CParser, PluginName As String, ByVal SettingName As String) As Variant
    'PluginName must be the ProgID of the Plugin in question.
    'The plugin must be loadable as well.
    Dim GetPlugin As Object
    Dim EvalEvents As IEvalEvents, CoreObj As ICorePlugin
    On Error Resume Next
    Set GetPlugin = createObject(PluginName)
    If Err <> 0 Then Exit Function
    If TypeOf GetPlugin Is IEvalEvents Then
        Set EvalEvents = GetPlugin
        GetPluginSettingEx = EvalEvents.Settings(withparser, PluginName)
    ElseIf TypeOf GetPlugin Is ICorePlugin Then
        Set CoreObj = GetPlugin
        GetPluginSettingEx = CoreObj.Settings(withparser, PluginName)
    End If
            




End Function

Public Function SavePluginSettingEx(withparser As CParser, PluginName As String, SettingName As String, NewSetting As Variant)
    Dim GetPlugin As Object
    Dim EvalEvents As IEvalEvents, CoreObj As ICorePlugin
    On Error Resume Next
    Set GetPlugin = createObject(PluginName)
    If Err <> 0 Then Exit Function
    If TypeOf GetPlugin Is IEvalEvents Then
        Set EvalEvents = GetPlugin
        'GetPluginSettingEx = EvalEvents.Settings(withparser, PluginName)
        EvalEvents.Settings(withparser, PluginName) = NewSetting
    ElseIf TypeOf GetPlugin Is ICorePlugin Then
        Set CoreObj = GetPlugin
        'GetPluginSettingEx = CoreObj.Settings(withparser, PluginName)
        CoreObj.Settings(withparser, PluginName) = NewSetting
    End If
            




End Function

Public Function PluginKey(ByVal Setname As String, ByVal PluginProgID As String) As String
'note, NEVER has a slash on the end.
    PluginKey = SetKey(Setname) & "Plugins\" & PluginProgID
End Function
Public Sub InstallPlugin(ByVal PluginProgID As String, Optional ByVal Setname As String = "Default")
    'InstallPlugin- Installs a given Plugin ProgID under SetName.
    With mregistry
        'set the disabled value of it to 0. This will automatically create
        'all the relevant section keys.
        .ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Plugins\" & PluginProgID, "Disabled", RREG_DWORD, False) = False

    End With
    'wow- one significant line of code to install a plugin.
End Sub
Public Sub UnInstallPlugin(ByVal PluginProgID As String, Optional ByVal Setname As String = "Default")
    'Uninstalls a given plugin, by actually deleting it's entire key.
    
    With mregistry
        .ClassKey = HHKEY_CURRENT_USER
        .SectionKey = SetKey(Setname) & "Plugins\" & PluginProgID
        .DeleteKey True
        
    End With
    
    
End Sub
Public Sub DisablePlugin(ByVal PluginProgID As String, Optional ByVal Setname As String = "Default")
 With mregistry
        'set the disabled value of it to 1. This will disable it.
        
        .ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Plugins\" & PluginProgID, "Disabled", RREG_DWORD, False) = True
        
End With
End Sub
Public Function DebugObject() As CDebug
Attribute DebugObject.VB_Description = "Returns the debug/logged output object in use by the parser."
Attribute DebugObject.VB_MemberFlags = "40"
    Set DebugObject = CDebug
End Function


Public Sub SetPluginDisableState(ByVal PluginProgID As String, ByVal Enable As Boolean, Optional ByVal Setname As String = "Default")
    With mregistry
        If Enable Then
            .ClassKey = HHKEY_CURRENT_USER
            .SectionKey = SetKey(Setname) & "Plugins\" & PluginProgID
            .ValueKey = "Disabled"
            .Value = 0
            '.DeleteKey
        Else
            DisablePlugin PluginProgID, Setname
        End If
    End With


End Sub
Public Function GetPluginType(ValCheck As Object, Optional ByRef PopEval As IEvalEvents = Nothing, Optional ByRef PopCore As ICorePlugin = Nothing) As PluginTypeConstants
    Dim CastEval As IEvalEvents, castCore As ICorePlugin
    Dim CheckIt As Object
    
    'If IsObject(ValCheck) Then
            On Error Resume Next
            Set CastEval = ValCheck
            If Err <> 0 Then
                'an error has occured.
                'not EvalEvents...
                On Error Resume Next
                Set castCore = ValCheck
                If Err <> 0 Then
                    'not a plugin!
                Else
                    GetPluginType = PluginType_Core
                End If
            Else
                'Evaluation-Type plugin
                Set castCore = ValCheck
                If Err <> 0 Then
                    GetPluginType = PluginType_Evaluation
                Else
                    GetPluginType = PluginType_Both
                End If
            End If
    
    'End If


    Set PopEval = CastEval
    Set PopCore = castCore
End Function
Public Function IsPluginInstalled(ByVal progID As String, Optional ByVal Setname As String = "Default") As Boolean
    Dim progIDs() As String, ptypes
    Dim pcount As Long, I As Long
    progIDs = GetPluginProgIDs(Setname, pcount, False, ptypes)
    For I = LBound(progIDs) To UBound(progIDs)
        If progIDs(I) = progID Then
            IsPluginInstalled = True
            Exit Function
        End If
    Next I
    
        
            




End Function
Public Function IsPluginDisabled(ByVal PluginProgID As String, Optional ByVal Setname As String = "Default")
    IsPluginDisabled = CBool(mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "Plugins\" & PluginProgID, "Disabled", RREG_DWORD, False)) = True
End Function
'Public Function IsParsePluginDisabled(ByVal PluginProgID As String, Optional ByVal Setname As String = "Default")
'    IsParsePluginDisabled = CBool(mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname) & "ParsePlugins\" & PluginProgID, "Disabled", RREG_DWORD, False)) = True
'End Function
Public Function GetAppVersion() As String
    GetAppVersion = Replace$(CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision), " ", "")
End Function
Private Sub Class_Initialize()
    Set mregistry = New cRegistry
    
    'CreatePopups
End Sub
'Friend Sub CleanEggData()
'    'delete year sections that are for previous years.
'    'they are irrelevant.
'    Dim years() As String, Ycount As Long, I As Long
'    Const secbase = AppKey & "TOPSECRET" & "\"
'    With mregistry
'        .ClassKey = HHKEY_CURRENT_USER
'        .SectionKey = secbase
'        Call .EnumerateSections(years, Ycount)
'        For I = 1 To UBound(years)
'            If Val(years(I)) < Year(Now) Then
'                'delete this key.
'                .SectionKey = secbase & "\" & years(I)
'                .DeleteKey
'
'            End If
'        Next I
'    End With
'End Sub
Private Function SECRETS() As String
    SECRETS = AppKey & "TOPSECRET" & "\" & Trim$(str$(Year(Now))) & "\"
End Function
'Friend Function DisplayedDayEgg(ForDate As Date) As Boolean
'    'returns wether it has already displayed the Day Egg for the specified Day.
'    'the test does not look at the year portion.
'    Dim LookFor As String
'
'    LookFor = str$(Month(ForDate)) & "," & str$(Day(ForDate))
'    With mregistry
'        DisplayedDayEgg = .ValueEx(HHKEY_LOCAL_MACHINE, SECRETS, LookFor, RREG_DWORD, False) = True
'        .ValueEx(HHKEY_LOCAL_MACHINE, SECRETS, LookFor, RREG_DWORD, False) = True
'    End With
'End Function
'Friend Sub SavePrefetchData(DictPrefetch As Scripting.Dictionary)
'    'This has to do with Sub Main()
'    'when MAIN() executes, it will load the Prefetch dictionary, which is stored in a sort
'    'of table format.
'    Dim X As cFormitem
'
'
'End Sub
Public Function FileNameFromProgID(ByVal progID As String) As String

    Dim FName As String
    Dim Objuse As Object
    Dim IInfo As tli.InterfaceInfo
    Set Objuse = createObject(progID)
    Set IInfo = InterfaceInfoFromObject(Objuse)
    FileNameFromProgID = IInfo.Parent.ContainingFile
    Set IInfo = Nothing
    Set Objuse = Nothing
    




End Function
Public Property Get NumCustomFunctionParams(Optional ByVal Setname As String = "Default") As Long
    With mregistry
        NumCustomFunctionParams = .ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "NumCustomFunctionParams", RREG_DWORD, 26)
    End With


End Property
Public Property Let NumCustomFunctionParams(ByVal Setname As String, ByVal Vdata As Long)
    mregistry.ValueEx(HHKEY_CURRENT_USER, SetKey(Setname), "NumCustomFunctionParams", RREG_DWORD, 26) = Vdata



End Property
Public Sub ClearRegistryofBPData(Optional ByVal DeleteKeys As CleanRegistryKeyConstants = CLEAN_BOTH)
'


    With mregistry
        If DeleteKeys And clean_machine = clean_machine Then
            .ClassKey = HHKEY_LOCAL_MACHINE
            .SectionKey = AppKey
            .DeleteKey True
        End If

        If DeleteKeys And Clean_user = Clean_user Then
            .ClassKey = HHKEY_CURRENT_USER
            .SectionKey = AppKey
            .DeleteKey True
        End If
    End With

End Sub
Public Sub CleanRegistryOfSet(ByVal Setname As String, Optional ByVal DeleteKeys As CleanRegistryKeyConstants = CLEAN_BOTH)
    'deletes ALL entries in the registry of the given set... all settings when this set is loaded will revert to their defaults.

    With mregistry

        If DeleteKeys And clean_machine = clean_machine Then
            .ClassKey = HHKEY_LOCAL_MACHINE
            .SectionKey = SetKey(Setname)
            .DeleteKey True
        End If
        If DeleteKeys And Clean_user = Clean_user Then
            .ClassKey = HHKEY_LOCAL_MACHINE
            .SectionKey = SetKey(Setname)
            .DeleteKey True
        End If

    End With






End Sub

Private Sub Class_Terminate()
'    mBaseINI.SaveINI
'    Dim loopini As CINIData
'    For Each loopini In mINIObjects
'        loopini.SaveINI
'    Next
End Sub
