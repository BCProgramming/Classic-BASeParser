Attribute VB_Name = "ModEval"
Option Explicit
Private Function InIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    InIDE = Err.Number <> 0
End Function
Private Function WriteLineWrap(Optional ByVal Text As String = "", Optional ByVal CrLf As Boolean = True, Optional Destination As ConsoleOutputDestinations = conStandardOutput, Optional Alignment As ConsoleWriteAlignments = conAlignNone)
    If InIDE Then
        Debug.Print Text
        WriteLineWrap = Len(Text)
    Else
        WriteLineWrap = Con.WriteLine(Text, CrLf, Destination, Alignment)
    End If
End Function
Private Sub ShowHelp()
    WriteLineWrap "BASeCamp BASeParser advanced Command-Line evaluation utility"
    WriteLineWrap "Version " & Trim$(App.Major) & "." & Trim$(App.Minor) & " Revision " & App.Revision
    WriteLineWrap
    WriteLineWrap "Syntax:"
    WriteLineWrap App.EXEName & " ""[expression]"" [options]"
    



End Sub
Sub Main()
    'Dim mparser As CParser
    Dim mparser As Object
    Dim mScript As Object
    'Dim commandlineParser As Object
    Dim mExpression As String
    Dim mresult As String
    Dim BCFileObject As Object
    If Not InIDE Then Con.Initialize
    On Error GoTo ParserCreationError
    Set mparser = CreateObject("BASeParserXP.CParser")
    'Set mparser = New CParser
    On Error Resume Next
    'Set commandlineParser = CreateObject("BCCParser.CommandLineParser")
    Set BCFileObject = CreateObject("BCFile.BCFSObject")
    
    mExpression = Command$
    

    'MSScriptControl.ScriptControl
    'no error so far.
    mparser.Create "BCEval"
    If Left$(mExpression, 1) = """" And Right$(mExpression, 1) = """" Then
        'remove quotes.
        mExpression = Mid$(mExpression, 2, Len(mExpression) - 2)
    End If
    
    mparser.Variables.Add "BCFile", BCFileObject
    mparser.expression = mExpression
    mparser.execute
    mresult = mparser.resultAsString
    WriteLineWrap "parsed Expression " & mExpression
    WriteLineWrap "result:"
    WriteLineWrap mresult
    End
ParserCreationError:
    WriteLineWrap "error description:" + Error$
    WriteLineWrap "Error: Failed to create CParser Object."
    WriteLineWrap "Using Script Engine. This may cause errors."
    WriteLineWrap "to restore full Evaluation functionality, ensure BASeCamp Expression Evaluation library is properly registered on this machine."
    
    
End Sub
