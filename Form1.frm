VERSION 5.00
Object = "{9AD34C4D-4B7C-410C-981E-6045605443D8}#1.0#0"; "vbaListView6.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   1740
      Width           =   1155
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwcontrol 
      Height          =   1215
      Left            =   660
      TabIndex        =   1
      Top             =   5520
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   3795
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim testctl As ScriptControl
Set testctl = New ScriptControl

testctl.Language = "VBScript"
testctl.AddObject "FormUse", Me, False
testctl.AddCode Text1.Text
testctl.CodeObject.TestProcedure Me

'Me.lvwcontrol.ListItems.Add , "TESTITEM" + Str$(Rnd), "test item"
End Sub

