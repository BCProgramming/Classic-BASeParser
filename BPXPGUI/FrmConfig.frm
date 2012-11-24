VERSION 5.00
Begin VB.Form FrmConfig 
   Caption         =   "BASeParser Configuration (APP)"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleMode       =   0  'User
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add..."
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   90
      Width           =   1185
   End
   Begin VB.ComboBox CboConfigsets 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   135
      Width           =   1995
   End
   Begin VB.Frame FraConfiguration 
      Caption         =   "&Configuration For "
      Height          =   4515
      Left            =   45
      TabIndex        =   2
      Top             =   675
      Width           =   7170
      Begin BPXPGUI.UCConfigSet SetEditor 
         Height          =   4200
         Left            =   45
         TabIndex        =   4
         Top             =   270
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   7408
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5265
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6165
      TabIndex        =   0
      Top             =   5265
      Width           =   1050
   End
   Begin VB.Label lblSet 
      AutoSize        =   -1  'True
      Caption         =   "Configuration Set:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   225
      Width           =   1260
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer

Implements BASeParserXP.IConfigProvider
Private mCancel As Boolean
Private mshowstate As ConfigShowStates
Private Sub CboConfigsets_Change()
'change the current display.
    'SetEditor.SaveData
    SetEditor.LoadData CboConfigsets.List(CboConfigsets.ListIndex)
End Sub

Private Sub CmdCancel_Click()
mCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    'populate the combobox.
    'with all the configset names, enumerated from the registry.
    Dim Gotnames() As String, numnames As Long
    Dim currpop As Long
    Me.Caption = "BASeParser Configuration"
    'With seteditor
        Gotnames = parsersettings.GetConfigSetNames(numnames)
        CboConfigsets.Clear
        For currpop = 1 To UBound(Gotnames)
            CboConfigsets.AddItem Gotnames(currpop)
        Next
        CboConfigsets.ListIndex = 0
        CboConfigsets_Change
    'End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If CBool(GetAsyncKeyState(vbKeyShift)) Then
            mshowstate = Config_Intermediate
        
        End If
    
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetEditor.SaveData
End Sub

Private Function IConfigProvider_QueryShowState() As BASeParserXP.ConfigShowStates
    '
    IConfigProvider_QueryShowState = mshowstate
End Function

Private Property Get IConfigProvider_Setting(ByVal PluginName As String, ByVal SettingName As String) As Variant
'
End Property
