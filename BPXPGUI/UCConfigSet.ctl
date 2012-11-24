VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.UserControl UCConfigSet 
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ScaleHeight     =   4665
   ScaleWidth      =   7710
   Begin VB.PictureBox PicTabs 
      BackColor       =   &H000000FF&
      Height          =   2985
      Index           =   1
      Left            =   4725
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4560
      Begin BPXPGUI.UCPlugins UCPlugs 
         Height          =   2400
         Left            =   270
         TabIndex        =   11
         Top             =   225
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   4233
      End
   End
   Begin vbalDTab6.vbalDTabControl TabProperties 
      Height          =   3615
      Left            =   135
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox PicTabs 
         Height          =   2985
         Index           =   0
         Left            =   0
         ScaleHeight     =   2925
         ScaleWidth      =   4500
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4560
         Begin VB.CheckBox chkDisableCore 
            Caption         =   "Disable Core Plugins"
            Height          =   240
            Left            =   315
            TabIndex        =   10
            Top             =   1665
            Width           =   1860
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "&Plugins..."
            Height          =   465
            Left            =   3060
            TabIndex        =   9
            Top             =   2430
            Width           =   1410
         End
         Begin VB.TextBox txtMaxCached 
            Height          =   330
            Left            =   2700
            TabIndex        =   8
            Top             =   1035
            Width           =   555
         End
         Begin VB.TextBox txtMinExprSize 
            Height          =   330
            Left            =   2700
            TabIndex        =   5
            Top             =   585
            Width           =   555
         End
         Begin VB.CheckBox ChkOptimize 
            Caption         =   "&Optimize Expressions where possible"
            Height          =   285
            Left            =   270
            TabIndex        =   3
            Top             =   225
            Width           =   2940
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Maximum size of Cached Stacks:"
            Height          =   195
            Left            =   315
            TabIndex        =   7
            Top             =   1080
            Width           =   2340
         End
         Begin VB.Label LblFormItems 
            AutoSize        =   -1  'True
            Caption         =   "Formula Items."
            Height          =   195
            Left            =   3420
            TabIndex        =   6
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblminExpr 
            AutoSize        =   -1  'True
            Caption         =   "Minimum Expression size of "
            Height          =   195
            Left            =   315
            TabIndex        =   4
            Top             =   675
            Width           =   1965
         End
      End
   End
End
Attribute VB_Name = "UCConfigSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'UCConfigSet: modifies the properties of a single configuration set.
'manages everything, and exposes stuff via the Save and Load methods. (Load requires
'the config set to view)
Private IsDirty As Boolean
Private mConfigSet As String        'our configuration set, which we are viewing.
Private mSetData As BASeParserXP.ParserConfigSet
Private Sub Dirtied()
    IsDirty = True
End Sub

Public Sub LoadData(ByVal ConfigSet As String)
'load the data about this configuration set.
'thanks to the backend supplied by BASeParser XP itself (which is good, since we are supposed to be the FRONT end...
    mSetData = parsersettings.GetConfigSetData(ConfigSet)
    
    mConfigSet = ConfigSet
    With mSetData
        ChkOptimize.Value = Abs(.Optimize)
        txtMinExprSize.Text = CStr(.MinCacheSize)
        txtMaxCached.Text = CStr(.MaxQueueSizeBeforePurge)
        chkDisableCore.Value = Abs(.DisableCore)
    
    End With
    
    UCPlugs.LoadProgIDs mSetData.ProgIds, ConfigSet
    IsDirty = False
End Sub
Public Sub SaveData()
    mSetData.MaxQueueSizeBeforePurge = Val(txtMaxCached)
    mSetData.MinCacheSize = Val(txtMinExprSize.Text)
    mSetData.Optimize = (ChkOptimize.Value = vbChecked)
    'the parsersettings method call won't overwrite
    'this Save'd data because the Parser only knows about the stuff
    'it defines, such as wether a plugin is enabled.
    'the UCPlugs control instantiates the different
    'ISettingsPage implementors if they are defined for the control.
    'then, when told to save, it turns around and tells each one to
    'save to the correct registry key.
    UCPlugs.Save mConfigSet
    parsersettings.SaveConfigSetData mSetData
    
    IsDirty = False
End Sub

Private Sub ChkOptimize_Click()
    txtMaxCached.Enabled = (ChkOptimize.Value = vbChecked)
    txtMinExprSize.Enabled = (ChkOptimize.Value = vbChecked)
    Dirtied
End Sub

Private Sub cmdPlugins_Click()
FrmPlugins.RefreshList mConfigSet
FrmPlugins.Show vbModal, Me
End Sub

Private Sub PicTabs_Resize(Index As Integer)
'index one contains the userControl.
If Index = 1 Then
With PicTabs(Index)
    UCPlugs.Move 0, 0, .ScaleWidth, .ScaleHeight
End With
End If
End Sub

Private Sub txtMaxCached_Change()
Dirtied
End Sub

Private Sub txtMaxCached_Validate(Cancel As Boolean)
    ValCtl txtMaxCached
End Sub
Private Sub ValCtl(CtlVal As Control)
If Not IsNumeric(CtlVal.Text) Then
        MsgBox "You must enter a Number."
        CtlVal.BackColor = vbRed
        CtlVal.ForeColor = vbYellow
    Else
        CtlVal.BackColor = vbWindowBackground
    End If
End Sub
Private Sub txtMinExprSize_Change()
    Dirtied
End Sub

Private Sub txtMinExprSize_Validate(Cancel As Boolean)
    ValCtl txtMinExprSize
End Sub

Private Sub UserControl_Initialize()
'initialize the UI components.
    'the tab will have the following parts:
    'General, for the general parser stuff for that configuration set
    'and Plugins, for changing the properties of Plugins.
    With TabProperties
        Set .Tabs.Add("GENERAL", , "&General").Panel = PicTabs(0)
        
        
'        VB.Load PicTabs(1)
    
        'NOTE: the picTabs(1), Plugins tab will contain another UserControl,
        'UCPlugins, which takes care of configuring Plugins- since that
        'involves all sorts of ActiveX Voodoo.
        Set .Tabs.Add("PLUGINS", , "&Plugins").Panel = PicTabs(1)
    End With
    'whew, that was REALLY HARD! look! I'm sweating! Oh- never mind, I spilled my chocolate milk.
    'Wait a sec- I'm not drinking chocolate milk!
    
End Sub

Private Sub UserControl_Resize()
    Dim I As Long
    TabProperties.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    For I = PicTabs.LBound To PicTabs.UBound
        PicTabs(I).Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Next I
End Sub
