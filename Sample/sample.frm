VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwVariables 
      Height          =   3075
      Left            =   5400
      TabIndex        =   4
      Top             =   540
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   5424
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox RTBMessage 
      Height          =   3030
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   5345
      _Version        =   393217
      RightMargin     =   90000
      TextRTF         =   $"sample.frx":0000
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      ToolTipText     =   "i fuck a monkey"
      Top             =   90
      Width           =   3660
   End
   Begin VB.CommandButton cmdeval 
      Caption         =   "&Evaluate"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   45
      Width           =   1140
   End
   Begin VB.Label lblVariables 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Variables:"
      Height          =   195
      Left            =   5445
      TabIndex        =   3
      Top             =   225
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub HandleOperation Lib "BPCOREOP2.dll" (ByVal Op As String, ByRef OpA As Variant, ByRef OpB As Variant, ByRef retval As Variant)
Private Declare Sub HandleFunction Lib "BPCOREOP2.dll" (ByVal FuncName As String, ByRef ArgList As Variant, retval As Variant)

Private Declare Function CanHandleOperation Lib "BPCOREOP2.dll" (ByVal Op As String) As Long
Private Declare Function CalHandleFunction Lib "BPCOREOP2.dll" (ByVal FuncName As String) As Long

    Dim WithEvents tester As CParser
Attribute tester.VB_VarHelpID = -1
    Dim WithEvents checkvariables As CVariables
Attribute checkvariables.VB_VarHelpID = -1
    Dim TimeStart As Double

Private Sub checkvariables_VarAdded(ByVal VarAdded As BASeParserXP.CVariable)
    lvwVariables.ListItems.Add(, VarAdded.Name, VarAdded.Name).SubItems(1) = VarAdded.Value
    
End Sub

Private Sub checkvariables_VarChanged(ByVal VarChanged As BASeParserXP.CVariable, ByVal OldValue As Variant)
    lvwVariables.ListItems(VarChanged.Name).SubItems(1) = VarChanged.Value
End Sub

Private Sub checkvariables_VarReplace(Replacing As BASeParserXP.CVariable, Allow As Boolean)
'
End Sub

Private Sub checkvariables_VarRetrieved(ByVal VarRetrieved As BASeParserXP.CVariable)
    '
End Sub

Private Sub cmdeval_Click()


    'tester.ParseInfix (txttest.Text)
    'tester.AddEventSink Me
    tester.Expression = txttest.Text
    TimeStart = Timer
    'lblVariables.Caption = tester.Execute
End Sub

Private Sub Form_Load()
    Set tester = New CParser
    Call tester.Create("Default")
    Set checkvariables = tester.Variables
    lvwVariables.ColumnHeaders.Add , , "Name"
    lvwVariables.ColumnHeaders.Add , , "Value"
End Sub


Private Sub WriteMessage(ByVal Strmessage As String, Optional ByVal Strcolor As Long = vbBlack)

    Dim active As Control
    Set active = ActiveControl
    RTBMessage.SelStart = Len(RTBMessage.Text)
    RTBMessage.SelText = vbCrLf & Strmessage
    RTBMessage.SelColor = Strcolor
    RTBMessage.SetFocus
    SendKeys "{Pgdn}"
    active.SetFocus
    




End Sub

Private Sub tester_ExecuteComplete(Valret As Variant)
    WriteMessage "ExecuteComplete " & Timer, vbGreen
    WriteMessage tester.ResultToString(Valret)
End Sub

Private Sub tester_ParseComplete()
    WriteMessage "ParseComplete " & Timer, vbRed
End Sub
