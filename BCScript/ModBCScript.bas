Attribute VB_Name = "ModBCScript"
Option Explicit
'implements BASeParserXP.IEvalEvents
Private Declare Sub ExitProcessX Lib "kernel32" (ByVal uExitCode As Long)
Public Enum StringTableConstants
    STC_LOGO = 101
    STC_UNEXPECTEDERROR = 102
    STC_ERRTERMINATING = 103
End Enum
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer

Public mMainParser As Object
Public VerboseFlag As Boolean
Private mCCparser As Object
Private mScripting As CScripting

Private Function LoadResStringEx(ByVal ID As Variant, ParamArray params() As Variant) As String

Dim StrLoad As String, I As Long

StrLoad = LoadResString(ID)


'Perform replacements....
For I = 0 To UBound(params)
    StrLoad = Replace$(StrLoad, "%" & Trim$(I + 1), params(I))
Next I

StrLoad = Replace$(StrLoad, "%APPMAJOR%", Trim$(Str$(App.Major)))
StrLoad = Replace$(StrLoad, "%APPMINOR%", Trim$(Str$(App.Minor)))
StrLoad = Replace$(StrLoad, "%APPREVISION%", Trim$(Str$(App.Revision)))
StrLoad = Replace$(StrLoad, "/n", vbCrLf)
StrLoad = Replace$(StrLoad, "/t", vbTab)
LoadResStringEx = StrLoad



End Function
Private Sub ExitProcess(ByVal ExitCode As Long)
    If Compiled Then
        Set mMainParser = Nothing
        Set mCCparser = Nothing
        ExitProcessX ExitCode
    Else
        Debug.Print "exit code:" & ExitCode
        End
    End If
End Sub

Private Function AppLogo() As String


    AppLogo = "BASeCamp BCScript BASeParser Script Execution Engine" & vbCrLf & _
            "Version " & Trim$(App.Major) & "." & Trim$(App.Minor) & " Build " & App.Revision & vbCrLf
    


End Function
Public Sub ExecuteScript(ByVal StrScriptFile As String, ParserUse As Object)
    Dim StrCode As String
    Dim fNum As Long
    Dim mParser As Object
        'Use Late binding, so we can show a proper error if the COM objects fail to load.
        On Error Resume Next
        Set mParser = ParserUse.Clone
        If mParser Is Nothing Then
            WriteLine "ERROR: Failed to clone Parser Object Class While Processing script file """ & StrScriptFile & """."
            Err.Raise Err.Number, Err.Source, Err.Description
        Else
            mParser.Variables.Add "BCSH", mScripting
            
            
            
            
        
        End If
        fNum = FreeFile
        Open StrScriptFile For Input As #fNum
        If Err <> 0 Then
            WriteLine "ERROR: failed to open Script file, """ & StrScriptFile & """."
            Err.Raise Err.Number, Err.Source, Err.Description
        
        End If
        StrCode = Input$(LOF(fNum), fNum)
        If Err <> 0 Then
            WriteLine "ERROR: File I/O error reading script file, """ & StrScriptFile & """."
            WriteLine "#" & Err.Number & " Description:" & Err.Description
            
        
        End If
        Close #fNum
        On Error GoTo ParseError
        If VerboseFlag Then WriteLine "adding Script Object variable..."
        
          mParser.Variables.Add "BCSH", mScripting
        If VerboseFlag Then WriteLine "PARSER:Assigning Expression..."
          mParser.Expression = StrCode
          If VerboseFlag Then WriteLine "PARSER:Parsing..."
          mParser.ParseMulti StrCode
          If VerboseFlag Then WriteLine "PARSER:Parse Complete. Executing..."
          On Error GoTo ParseError
          mParser.Execute
          If VerboseFlag Then WriteLine "PARSER:Execute complete."
          If VerboseFlag Then WriteLine "PARSER: script """ & StrScriptFile & """ finished execution."
        Exit Sub
ParseError:
        Debug.Print "error."
        WriteLine "Error:" & Err.Description
End Sub

Private Sub ShowHelp()
    'write help text to the console.
    Con.WriteLine AppLogo
    Con.WriteLine "Syntax:"
    Con.WriteLine "BCScript <Filename.bcs> -STDIN -H|-? -V -NOLOGO "
    Con.WriteLine "Where:"
    Con.WriteLine ""
    Con.WriteLine "<Filename.bcs>" & vbTab & vbTab & "Script file to execute."
    Con.WriteLine "-STDIN" & vbTab & vbTab & "Specifies to read data from standard input." & vbCrLf & _
        vbTab & vbTab & "If a filename(s) is specified, the file will be loaded and run before the standard input is processed."
        
    Con.WriteLine "-V" & vbTab & vbTab & " Displays Verbose output."
    
    Con.WriteLine "-NOLOGO" & vbTab & vbTab & " Prevents the display of the startup text."
    
End Sub
Sub Main()
    Dim strinput As String, fNum As Long
    VerboseFlag = False
    
    '***change for final release for late-binding code.
    Dim mCCparser As Object
    Dim Fileinput As String
    Dim stdIn As Boolean
    
    
    
    
    '***
  '    Open "C:\templog.log" For Append As #1
  ' Print #1, "LOG:" & "opened."
  ' Close #1
    Con.Initialize
    
    If Not InStr(Command$, " /NOLOGO ") Then
        WriteLine AppLogo
        WriteLine "Copyright 2008-2009 BASeCamp Corporation. All rights reserved."
    End If
    
    If Not IsUserAnAdmin() Then
        WriteLine "warning: You are not running the interpreter as an administrator. this may cause issues."
    
    
    End If
    
    On Error Resume Next
    Set mCCparser = CreateObject("BCCParser.CommandLineParser")
    
    If Err <> 0 Then
        'report the error...
        'no commandlineparser.
        WriteLine "ERROR: BASeCamp CommandLine Parser object not available."
        WriteLine "Make sure BCCParser.dll is properly registered on your system."
        Con.ExitCode = 1
        Exit Sub
    Else
        mCCparser.parsearguments Command$
    End If
    Dim Switchloop As Object
    For Each Switchloop In mCCparser.Switches
        Select Case UCase$(Switchloop.SwitchUndec)
            Case "V"
                VerboseFlag = True
                Con.WriteLine "Verbose output enabled."
            Case "H", "?"
                ShowHelp
                Con.ExitCode = 0
                Exit Sub
            Case "STDIN"
                stdIn = True
        End Select
    
    
    Next Switchloop
    
    
    Err.Clear
    Set mMainParser = CreateObject("BASeParserXP.CParser")
    If Err <> 0 Then
        WriteLine "Fatal Error Creating Main CParser Object:"
        WriteLine Err.Description
        Con.ExitCode = 2
        Exit Sub
    ElseIf Err = 0 Then
        If VerboseFlag Then WriteLine "Main Parser object created successfully."
        
    End If
    On Error Resume Next
    If VerboseFlag Then WriteLine "Initializing Main parser object to Configset ""BCScript"""
    mMainParser.Create "BCScript"
    Set mScripting = New CScripting
    
    If Err <> 0 Then
        WriteLine "Error Initializing Parser object to Configuration set ""BCScript"" & vbcrlf & err.Description" _
        & "Trying default...."
        Err.Clear
        mMainParser.Create
        If Err <> 0 Then
            WriteLine "Failed to initialize Parser object:" & vbCrLf & Err.Description
            Exit Sub
        End If
    ElseIf Err = 0 Then
        If VerboseFlag Then WriteLine "Initialized Parser to configset ""BCScript"""
    
    End If
    If VerboseFlag Then WriteLine "Arguments passed:" & vbCrLf & Command$
    
    If VerboseFlag Then WriteLine "Parsing arguments..."
    Err.Clear
    mCCparser.parsearguments Command$
    If Err <> 0 Then
        WriteLine "error parsing arguments!" & vbCrLf & _
        Err.Description
    
    End If
  
    
    Dim argloop As Object
    'first, look at the switches.
    If mCCparser.Switches.Count > 0 Then
'        For Each SwitchLoop In mCCparser.Switches
'            If StrComp(SwitchLoop.SwitchUndec, "h", vbTextCompare) = 0 Then
'                ShowHelp
'                Exit Sub
'            ElseIf StrComp(SwitchLoop.SwitchUndec, "stdin", vbTextCompare) = 0 Then
'                'Standard Input...
'
'                stdIn = True
'
'
'            End If
'
'
'        Next


    End If
    If mCCparser.Arguments.Count > 0 Then
        For Each argloop In mCCparser.Arguments
           ExecuteScript argloop.ArgString, mMainParser
        
        
        Next argloop
    End If
    If stdIn Then
        If VerboseFlag Then WriteLine "reading Standard Input..."
        Dim fso As FileSystemObject
        Dim SScriptStream As TextStream
        Dim codestring As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set SScriptStream = fso.OpenTextFile("STDIN")
        codestring = SScriptStream.ReadAll
        SScriptStream.Close
        If VerboseFlag Then WriteLine "standard input was " & Len(codestring) & " bytes."
        If VerboseFlag Then WriteLine "standard input read. Executing."
        mMainParser.ParseMulti codestring
        mMainParser.Execute
        
        
        'mMainParser
    
    End If
    
    Exit Sub
ReportScriptError:
    'WriteLine "ERROR:" & Err.Description
    'WriteLine "Error Encountered during Script Execution. Terminating."
    

    Set mMainParser = Nothing
    
End Sub
Private Sub WriteLine(Text As String)
    If Con.Compiled Then
        Con.WriteLine Text
    Else
        Debug.Print Text
    End If
End Sub
