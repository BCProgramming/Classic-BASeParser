VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmGUIevaluator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BASeParser XP GUI evaluator"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   1440
      TabIndex        =   14
      Top             =   5280
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   510
      Left            =   7920
      TabIndex        =   13
      Top             =   5340
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Configuration"
      Height          =   465
      Left            =   8160
      TabIndex        =   12
      Top             =   3225
      Width           =   1365
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   465
      Left            =   8160
      TabIndex        =   11
      Top             =   2685
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROFILE"
      Height          =   510
      Left            =   8160
      TabIndex        =   10
      Top             =   4200
      Width           =   1365
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "&Evaluate"
      Height          =   420
      Left            =   5880
      TabIndex        =   9
      Top             =   180
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display Mode"
      Height          =   1365
      Left            =   8160
      TabIndex        =   3
      Top             =   1200
      Width           =   1545
      Begin VB.OptionButton OptNotation 
         Caption         =   "&Standard"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton OptNotation 
         Caption         =   "&Scientific"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   540
         Width           =   1275
      End
      Begin VB.OptionButton OptNotation 
         Caption         =   "&Hexadecimal"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   765
         Width           =   1275
      End
      Begin VB.OptionButton OptNotation 
         Caption         =   "&Octal"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   990
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdFunctionHelp 
      Caption         =   "&Function Help..."
      Height          =   375
      Left            =   8250
      TabIndex        =   2
      Top             =   705
      Width           =   1275
   End
   Begin RichTextLib.RichTextBox RTBmessage 
      Height          =   4260
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   7514
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   90000
      TextRTF         =   $"frmGUIEvaluator.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbEvaluate 
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   582
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      RightMargin     =   90000
      TextRTF         =   $"frmGUIEvaluator.frx":0080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Output:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   585
      Width           =   525
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "context"
      Visible         =   0   'False
      Begin VB.Menu MnuFunction 
         Caption         =   "Function"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmGUIevaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents mParser As CParser
Attribute mParser.VB_VarHelpID = -1

Private mTimerStart As Double

'This form is FOR TESTING ONLY! IF used for other purposes I CANNOT
'PROVIDE SUPPORT- although I don't really provide support for anything else, either...
Private Sub CmdEvaluate_Click()
    Dim tempresult As Variant
    mTimerStart = Timer
    On Error GoTo ReportError
    Screen.MousePointer = vbArrowHourglass
    'In a production application, it
    'is probably a good Idea to
    'add handling of parser errors.
    'whenever a expression is parsed or executed, an error can result.
    mParser.Expression = rtbEvaluate.Text
    'this is important- otherwise, you may start wondering.
    'you need to force it to re-build the stack. This won't prevent
    'optimizations, however.
    
    'tempresult = mParser.Execute
    Screen.MousePointer = vbHourglass
    mParser.ExecuteByRef tempresult
    Screen.MousePointer = vbDefault
    Exit Sub
ReportError:
    Screen.MousePointer = vbDefault
    Showmessage "Error #" & Err.Number & " in " & Err.Source & vbCrLf & _
        "Description:" & Err.Description, vbRed
    'Debug.Print "TEMPRESULT:" & tempresult, "LASTRESULT:" & mParser.LastResult
End Sub

Private Sub cmdFunctionHelp_Click()
    'unload the current menu array, if any.
    Dim I As Long
    Dim allfuncs As String
    Dim Functionlist() As String
    On Error Resume Next
    For I = MnuFunction.UBound To MnuFunction.LBound Step -1
        Unload MnuFunction(I)
    
    Next I
    allfuncs = mParser.EvalListener.Self.GetHandledFunctionString(mParser)
    Functionlist = Split(allfuncs, " ")
    Err.Clear
    
    For I = 0 To UBound(Functionlist)
        Load MnuFunction(I)
        MnuFunction(I).Caption = Functionlist(I)
        MnuFunction(I).Visible = True
    Next I
    PopupMenu mnuFunctions
End Sub



Private Sub Command1_Click()
'purges buffer each time, for more accurate profiling output.
'(of course, this doesn't test any of the scenarios where
'optimizations occur. I think there is yet another bug with optimizations.
    mParser.Variables.Add "X", 0
    rtbEvaluate.Text = "SEQ(SQR(X+X)/X-X,X,3,300)"
    cmdEvaluate.Value = True
    'mParser.PurgeBuffer
    rtbEvaluate.Text = "SEQ(MID$(""TEST TEST TEST TEST"",X,1),X,2,10)"
    cmdEvaluate.Value = True
    'mParser.PurgeBuffer
    rtbEvaluate.Text = "STORE(X,50)"
    cmdEvaluate.Value = True
    'mParser.PurgeBuffer
    rtbEvaluate.Text = "Sqr(-X)*Sqr(-Abs(X))"
    cmdEvaluate.Value = True
    'mParser.PurgeBuffer
    rtbEvaluate.Text = "SEQ(SQR(X+Len(Str(X)))*3,X,1,300)"
    cmdEvaluate.Value = True
    
End Sub

Private Sub cmdAbout_Click()
mParser.ShowAbout Me
'mParser.Configure Me



'    Dim Testmatrix  As CMatrix
'    Dim Tester As CMatrix, TestInv As CMatrix
'    Dim testMult As CMatrix
'
'    Set Testmatrix = New CMatrix
'    Set Testmatrix = Testmatrix.CreateMatrix(2, 2, 2, 1, 5, 3)
'    Debug.Print Testmatrix.ToString
'    Set Testmatrix = Testmatrix.Invert
'    Debug.Print Testmatrix.ToString
'
'    Set Tester = Testmatrix.CreateMatrix(3, 3, 6, 3, 1, 2, 34, 2, 1, 7, 5)
'    Set TestInv = Tester.Invert
'    Set testMult = Tester.Multiply(TestInv)
'    MsgBox "original:" & vbCrLf & Tester.ToString
'    MsgBox "Inverse:" & vbCrLf & TestInv.ToString
'    MsgBox testMult.ToString
'    'msgbox "Multiply them:"
    
End Sub



Private Sub Command3_Click()
    Dim X As PropertyBag, testparser As CParser
    Set testparser = New CParser
    Set X = New PropertyBag
    testparser.Create
    testparser.Expression = """Hi"" + (5/Sqr(BPEMPTY+50))"
    
    X.WriteProperty "PARSER", testparser
    Set testparser = Nothing
    Set testparser = X.ReadProperty("PARSER")
    testparser.Create
    Stop
End Sub

Private Sub Command2_Click()
    mParser.Configure
End Sub

Private Sub Command4_Click()
FrmColored.Show
End Sub

Private Sub Form_Load()
    Set mParser = New CParser
    mParser.Create
    'Set mParser.CorePlugin = New CPlugEnvString
    'mParser.Functions.Add "P1+P2+P3", "TESTER"
    'add a test variable, FORM
    mParser.Variables.Add "FORM", Me
   ' Set VarView1.ParserWatch = mParser
   ' VarView1.Refresh
End Sub

Private Sub MnuFunction_Click(Index As Integer)
'used to be this one liner:
'     MsgBox mParser.FormatFunctionInformation(mParser.GetFunctionInformation(MnuFunction(Index).Caption), FH_EXTENSIVE), , "Help on " & MnuFunction(Index).Caption
Dim strFunction As String
Dim FuncInfo As FUNCTIONINFORMATION
Dim Strmsg As String
strFunction = MnuFunction(Index).Caption
FuncInfo = mParser.GetFunctionInformation(strFunction)
Strmsg = mParser.FormatFunctionInformation(FuncInfo, FH_EXTENSIVE)
MsgBox Strmsg, , "Help on " & strFunction
End Sub



Private Sub mParser_Error(ParserError As BASeParserXP.CParserError, RecoveryConst As BASeParserXP.ParserErrorRecoveryConstants)
    'Showmessage " Error:" & vbCrLf & ParserError.ToString, vbBlack
    RecoveryConst = PERR_RETURN
    Screen.MousePointer = MousePointerConstants.vbArrow
End Sub

Private Sub mParser_ExecuteComplete(valret As Variant)
Showmessage "Finished Executing expression:""" & mParser.Expression, vbRed
Showmessage "Execute Complete:" & Timer - mTimerStart, vbBlack
'Stop
Showmessage "Result:" & mParser.ResultToString(valret), vbBlue
End Sub

Private Sub mParser_ParseComplete()
    Showmessage "Parse Complete:" & Timer - mTimerStart, vbBlack
End Sub

Private Sub Showmessage(ByVal StrText As String, ByVal crColor As Long)
    With RTBmessage
        .SelStart = Len(.Text)
        .SelText = StrText & vbCrLf
        .SelColor = crColor
    
    
    End With




End Sub

Private Sub OptNotation_Click(Index As Integer)
mParser.Notation = Index
End Sub

Private Sub rtbEvaluate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CmdEvaluate_Click
    End If
End Sub
Public Function TESTAMETHOD(ByVal InputMe As String) As String

TESTAMETHOD = " Thanks, I love """ & InputMe & """"


    
End Function

