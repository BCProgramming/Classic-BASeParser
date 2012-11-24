VERSION 5.00
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form FrmBCETray 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BASeParser TrayIcon Evaluator"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4635
   Icon            =   "BCETray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   420
      TabIndex        =   2
      Top             =   2220
      Width           =   2475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Top             =   1560
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   2895
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   720
      Top             =   1560
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin vbalCmdBar6.vbalCommandBar cmdbarpopups 
      Align           =   1  'Align Top
      Height          =   435
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MainMenu        =   -1  'True
      Style           =   3
   End
End
Attribute VB_Name = "FrmBCETray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mTrayIcon As cSysTray
Attribute mTrayIcon.VB_VarHelpID = -1


Private Sub cmdbarpopups_RequestNewInstance(index As Integer, ctl As Object)
Dim newindex As Long
newindex = cmdbarpopups.UBound + 1
Load cmdbarpopups(newindex)
cmdbarpopups(newindex).Align = vbAlignNone
Set ctl = cmdbarpopups(newindex)
End Sub

Private Sub Command1_Click()
mTrayIcon.Tip(0) = Text1.Text
End Sub

Private Sub Form_Load()
    LoadMenus
    Set cmdbarpopups(0).Toolbar = cmdbarpopups(0).CommandBars("POPUPMENU")
    Set mTrayIcon = New cSysTray
    Set mTrayIcon.OwnerForm = Me
    Call mTrayIcon.CreateIcon(0, Me.Icon, "BASeCamp Tray Expression Evaluator")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdbarpopups(0).ShowPopupMenu X, Y, cmdbarpopups(0).CommandBars("POPUPMENU")
    
End Sub

Private Sub mTrayIcon_Click(ByVal IconID As Long, ByVal Button As MouseButtonConstants)
''
'Me.Show
List1.AddItem "mtrayIcon_Click"
End Sub

Private Sub mTrayIcon_DblClick(ByVal IconID As Long)
''
List1.AddItem "mtrayIcon_DblClick"
End Sub

Private Sub mTrayIcon_MouseMove(ByVal IconID As Long)
'
List1.AddItem "mtrayIcon_MouseMove"
End Sub
Private Sub LoadMenus()

Call ModCmdBarXML.CreateCommandBarsFromXML(ModCmdBarXML.LoadFileText(App.Path & "\popupmenu.xml"), cmdbarpopups(0))
'
End Sub
