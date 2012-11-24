VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form GUIEvaluator 
   Caption         =   "BASeParser Evaluator Front-End"
   ClientHeight    =   4785
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   375
      Left            =   8700
      TabIndex        =   6
      Top             =   3600
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   315
      Left            =   8700
      TabIndex        =   5
      Top             =   3060
      Width           =   915
   End
   Begin VB.PictureBox Picgradient 
      AutoRedraw      =   -1  'True
      Height          =   1395
      Left            =   8580
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   4
      Top             =   1440
      Width           =   1395
   End
   Begin RichTextLib.RichTextBox RTBEvaluator 
      Height          =   3315
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   5847
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"GUIEvaluator.frx":0000
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "&Evaluate"
      Height          =   435
      Left            =   8700
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.TextBox txteval 
      Height          =   915
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8475
   End
   Begin VB.Label lbloutput 
      Caption         =   "Parser Output:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1140
      Width           =   1875
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "GUIEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements BASeParserXP.IParserOutput
Private mPropertyBag As PropertyBag
Private WithEvents mParser As cparser
Attribute mParser.VB_VarHelpID = -1
Public Sub ShowMessage(ByVal Message As String)
    RTBEvaluator.Text = RTBEvaluator.Text & vbCrLf & Message
    RTBEvaluator.SelStart = Len(RTBEvaluator.Text)
End Sub

Private Sub cmdEvaluate_Click()
'
mParser.Expression = txteval.Text
mParser.Execute
End Sub

Private Sub Command1_Click()
    Set mPropertyBag = New PropertyBag
    mPropertyBag.WriteProperty "testing", mParser.firstitem
End Sub

Private Sub Command2_Click()
    Set mParser.firstitem = mPropertyBag.ReadProperty("testing")
End Sub

Private Sub Form_Click()
    Dim tinfo As TypeLibInfo
    Set tinfo = TypeLibInfoFromFile("C:\windows\system32\scrrun.dll")
    mParser.AddTypelibConstants tinfo
End Sub

Private Sub Form_Load()
    Set mParser = New cparser
    mParser.Create
    mParser.Variables.Add "Color", New Colour
    mParser.Variables.Add "Form", Me
    mParser.Variables.Add "eventer", New CTemp



    mParser.AddEventSink New BPCoreFunc.CBPGUIKit
End Sub

Private Sub IParserOutput_Message(withparser As BASeParserXP.cparser, ByVal Message As String, Optional ByVal VerbosityLevel As Integer = 0)
'
ShowMessage Message
End Sub

Private Sub IParserOutput_ShowMessage(ByVal Message As String, Optional ByVal VerbosityLevel As Integer = 0)
'
    ShowMessage Message

End Sub

Private Sub mnuAbout_Click()
 MsgBox "BASeParser XP, BASeCamp Corporation, (c) 2006-2008. BCEvaluator is a very simple front end for testing the Library."
End Sub

Private Sub mParser_Error(ParserError As BASeParserXP.CParserError, RecoveryConst As BASeParserXP.ParserErrorRecoveryConstants)
    ShowMessage vbCrLf
    RecoveryConst = PERR_RETURN
End Sub

Private Sub mParser_ExecuteComplete(valret As Variant)
    'done!
    Dim strresult As String
    strresult = mParser.ResultToString(valret)
    ShowMessage "result=" & strresult & vbCrLf
End Sub

Private Sub Picgradient_Click()
    Dim mColourParser As cparser
    Dim currX  As Long, currY As Long
    Dim useX As Integer, useY As Integer
    
    Dim ParserStartTime As Double, ParserEndTime As Double
    Dim pureStartTime As Double, PureEndTime As Double
    Set mColourParser = New cparser
    mColourParser.Create
    mColourParser.Variables.Add "X", 0
    mColourParser.Variables.Add "Y", 0
    mColourParser.Expression = "RGB(X,Y,0)"
    ParserStartTime = Timer
    For currX = 0 To Picgradient.ScaleHeight
        For currY = 0 To Picgradient.ScaleWidth
        
        
            useX = (currX / Picgradient.ScaleWidth) * 255
            useY = (currY / Picgradient.ScaleHeight) * 255
            mColourParser.Variables("X").Value = useX
            mColourParser.Variables("Y").Value = useY
            Picgradient.PSet (currX, currY), CLng(mColourParser.Execute)
        Next currY
        DoEvents
        Picgradient.Refresh
    Next currX
    ParserEndTime = Timer
    Picgradient.Cls
    pureStartTime = Timer
        For currX = 0 To Picgradient.ScaleHeight
        For currY = 0 To Picgradient.ScaleWidth
            useX = (currX / Picgradient.ScaleWidth) * 255
            useY = (currY / Picgradient.ScaleHeight) * 255
            Picgradient.PSet (currX, currY), RGB(useX, useY, 0)
        Next currY
        DoEvents
        Picgradient.Refresh
    Next currX
    PureEndTime = Timer
    
    
    Debug.Print "Parser:" & ParserEndTime - ParserStartTime
    Debug.Print "Pure:" & PureEndTime - pureStartTime
    
    
    
End Sub
