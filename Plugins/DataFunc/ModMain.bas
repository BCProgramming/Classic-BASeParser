Attribute VB_Name = "ModMain"
Option Explicit
Public mReg As New cRegistry


Private Type POINTAPI
    x As Long
    y As Long
End Type


Const Pluginkey = "Software\BASeCamp\BASeParser\Configuration Data\Plugins\"
Const Datafunckey = Pluginkey & "DataFunc.CDataFunc"
Const ScriptFuncKey = Pluginkey & "BPCoreFunc.CScriptFunctions"
Public Type ScriptLanguageData
    Name As String
    Masks As String    'Null-delimited string.
End Type

Public Type ScriptLangConfig
    NumTypes As Long
    ScriptData() As ScriptLanguageData
End Type
Sub Main()
    Debug.Print "DataFunc Plugin Sub Main()"
    



End Sub
Private Function DefaultDatabase() As String
Dim AppString As String
AppString = App.Path & "\datafunc.mdb"

DefaultDatabase = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppString & ";Persist Security Info=False"
End Function
Public Function GetConnection() As ADODB.Connection
    'GetConnection Function.
    'uses the registry to retrieve the appropriate Database connection string (defaulting to
    'a "DataFunc.mdb" JET database in the project directory), and opens the connection.
    'any errors are raised and should be handled.
    'TODO:\\
    'of course, a default installation will NOT include the database (well, - it might)
    'but, we still want to handle a non-existent database by creating another. Of course,
    'this new database will not have the forms or modules I may have added to the default database,
    'but- to bad.
    
    Dim mCon As Connection
    Dim ConString As String
    
    Set mCon = New ADODB.Connection
    On Error Resume Next
    'HKEY_CURRENT_USER\Software\BASeCamp\BASeParser\Configuration Data\Default\Plugins\DataFunc.CDataFunc
    
    ConString = mReg.ValueEx(HHKEY_CURRENT_USER, Datafunckey, "DBConnectionString", RREG_SZ, DefaultDatabase)
    If Err <> 0 Then
        Debug.Print "Registry access error-" & Err.Description
    End If
    On Error GoTo RollBackError
    mCon.Open ConString
    If mCon.State = adStateOpen Then
        Debug.Print "Connection Successfully opened to:" & ConString
    End If
    
    
    Set GetConnection = mCon
    
    Exit Function
RollBackError:
    'Rollback what we have done.
    'the only error possible here is the mCon.Open() String.
    If mCon.State <> adStateClosed Then
        'if the connection is open, close it.
        On Error Resume Next
        mCon.Close
        
    
    End If
    Debug.Print "Failed to open the Database- " & Error$
    Set mCon = Nothing
    Set GetConnection = Nothing



End Function

Public Sub GetScriptLangConfig(ToType As ScriptLangConfig)
    'GetScriptLangCopy: loads the initialization data from the registry
    'as to what language ID we need to give a scriptcontrol to use a specified script file.
    'we store this data, appropriately enough, ScriptFuncKey & "\LangData\<Script ID>\File masks"
    'so, we need to create a null-delimited list of the items within the scriptID key.
    '
    'So, enumerate the sections in langdata.
    Dim Senum() As String, scount As Long
    With mReg
    .ClassKey = HHKEY_CURRENT_USER
    .SectionKey = ScriptFuncKey & "\LangData\"
    'OK- enum the sections.
    If Not .EnumerateSections(Senum, scount) Then
        'sigh-
        'add VBScript and JavaScript.
    Else
        'whew.
        'redim the array to fit the stuff.
        ReDim ToType.ScriptData(0 To scount - 1)
    
    End If
    End With



End Sub

Public Sub InvokeDynamic(ByVal onObj As Object, ByVal Membername As String, Arguments As Variant, retval As Variant)
    'Dynamically invokes a method.
    'use InvokeHookArray of TLIApplication
    'first, we'll need to flip around the arguments.
    Dim ArgsPass() As Variant
    Dim AccessMode As InvokeKinds
    Dim Numerrs As Long
    
    Dim I As Long, current As Long
    If IsArray(Arguments) Then
    On Error Resume Next
            ReDim ArgsPass(UBound(Arguments))
            Err.Clear
        For I = UBound(Arguments) To 0 Step -1
            If Err <> 0 Then
                Arguments = ""
                Exit For
            End If
            If IsObject(Arguments(I)) Then
                Set ArgsPass(current) = Arguments(I)
            Else
                ArgsPass(current) = Arguments(I)
            End If
            current = current + 1
        Next I
    End If
    'arguments must be in reverse order.
    'OK, here goes the call.
    AccessMode = INVOKE_FUNC
        'default to a Function.
        On Error GoTo HandleBadMode
    Dim temphold As Variant
    If IsArray(Arguments) And Err.Number = 0 Then
'
'        If IsObject(TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())) Then
'            Set retval = TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())
'        Else
'            retval = TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())
'        End If

'the above was replaced with an assign call to fix the problem that we called the member twice.
      Assign retval, TLI.InvokeHookArray(onObj, Membername, AccessMode, ArgsPass())
    
    Else
        'use CallByName- the constants are the same, IE- INVOKE_FUNC will be the same
        'as vbMethod, so I imagine the constants are interchangable.
        'Anyway- no parameters.
            Assign retval, CallByName(onObj, Membername, AccessMode)
      
    End If
Exit Sub
HandleBadMode:
'handlebad modes.
Numerrs = Numerrs + 1
If Err.Number <> 0 Then

    'there are a few errors that could be raised by the Script host.
    'IE: wrong number of arguments or invalid property asignment.
    Select Case Err.Number
    
    
    
    End Select
    Select Case Numerrs
        Case 1
            AccessMode = INVOKE_PROPERTYPUT
            
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
        Resume
    Else
    Err.Raise Err.Number
End If

'whew.
End Sub

    Public Sub Assign(ByRef VarTo As Variant, ByRef assignthis As Variant)
        
        If IsObject(assignthis) Then
            Set VarTo = assignthis
        Else
            VarTo = assignthis
        End If
    
    
    
    
    
    
    
    End Sub

