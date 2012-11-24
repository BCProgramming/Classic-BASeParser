VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmEditBrush 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Brush properties"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraOtherOptions 
      Caption         =   "BS_HATCHED"
      Height          =   1860
      Index           =   2
      Left            =   4770
      TabIndex        =   7
      Top             =   2160
      Width           =   4695
      Begin MSComctlLib.ImageCombo ICboHatch 
         Height          =   330
         Left            =   1170
         TabIndex        =   10
         Top             =   405
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Hatch Style:"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   405
         Width           =   870
      End
   End
   Begin VB.CommandButton CmdColour 
      Caption         =   "&Change..."
      Height          =   375
      Left            =   1980
      TabIndex        =   12
      Top             =   1440
      Width           =   1230
   End
   Begin VB.PictureBox PicColour 
      BackColor       =   &H000000FF&
      Height          =   645
      Left            =   1035
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   11
      Top             =   1305
      Width           =   825
   End
   Begin VB.Frame FraOtherOptions 
      Caption         =   "BS_SOLID"
      Height          =   1860
      Index           =   0
      Left            =   45
      TabIndex        =   6
      Top             =   2160
      Width           =   4695
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3735
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   8
      Top             =   810
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame FraOtherOptions 
      Caption         =   "BS_PATTERN"
      Height          =   1860
      Index           =   3
      Left            =   4770
      TabIndex        =   4
      Top             =   270
      Width           =   4695
      Begin VB.CommandButton CmdChange 
         Caption         =   "&Change..."
         Height          =   375
         Left            =   3465
         TabIndex        =   5
         Top             =   225
         Width           =   1050
      End
      Begin VB.Image ImgPattern 
         Height          =   1455
         Left            =   90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1860
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2610
      TabIndex        =   3
      Top             =   4140
      Width           =   1050
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3735
      TabIndex        =   2
      Top             =   4140
      Width           =   1050
   End
   Begin MSComctlLib.ImageCombo ICboStyle 
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Top             =   405
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "BS_SOLID"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colour:"
      Height          =   195
      Left            =   720
      TabIndex        =   13
      Top             =   990
      Width           =   495
   End
   Begin VB.Label LblStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Style:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   495
      Width           =   390
   End
End
Attribute VB_Name = "FrmEditBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mBrushEdit As CBrush

Private Sub CmdChange_Click()
    Dim fdlg As BCFile.CFileDialog
    Dim SelectedFile As CFile
    
    Set fdlg = New BCFile.CFileDialog
    With fdlg
        .Caption = "Select Brush Pattern:"
        .Filter = "All Picture files(*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|Compuserve Graphics Interchange(GIF)|*.gif|Bitmap(BMP)|*.bmp|Jpeg/Jiff|*.jpg;*.jpe"
        .flags = OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_ENABLESIZING
        .BufferSize = 32767
        .InitialDir = CurDir$
        
        Set SelectedFile = fdlg.SelectOpenFile(Me.hwnd)
        
        If Not SelectedFile Is Nothing Then
            ImgPattern.Picture = LoadPicture(SelectedFile.Filename)
        
        End If
        
    End With
    
End Sub

Private Sub Form_Load()
'resize the form, since I moved everything around at design time to make- well, designing easier.
Dim I As Long
'resize the form to be just past the OK button. Give it 5 pixels.
Me.ScaleMode = vbPixels
Me.Move Me.Left, Me.Top, CmdOK.Left + CmdOK.Width + 5, CmdOK.Top + CmdOK.Height + 5
'Tada!
With FraOtherOptions(0)
    For I = 0 To FraOtherOptions.UBound - 1
        FraOtherOptions(I).Move .Left, .Top, .Width, .Height
    Next I
    'move the first item to the top.
    .ZOrder vbBringToFront
    PopulateICombos
End With
End Sub

Private Sub ICboStyle_Change()
    Dim chopped As String
    With ICboStyle.SelectedItem
    chopped = Mid$(.Tag, InStr(.Tag, "|"))
    FraOtherOptions(Val(chopped)).ZOrder 0
    End With
End Sub

Private Sub ICboStyle_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub PopulateICombos()
    
    With ICboStyle.ComboItems
'       Public Enum BrushStyleConstants
' BS_NULL = 1
' BS_HATCHED = 2
' BS_HOLLOW = BS_NULL
' BS_PATTERN = 3
' BS_PATTERN8x8 = 7
' BS_SOLID = 0
'End Enum
'Public Enum HatchStyleConstants
' HS_BDIAGONAL = 3
' HS_CROSS = 4
' HS_DIAGCROSS = 5
' HS_FDIAGONAL = 2
' HS_HORIZONTAL = 0
' HS_VERTICAL = 1
'End Enum
        .Add(, , "BS_SOLID").Tag = "STYLE|" & BS_SOLID
        .Add(, , "BS_NULL").Tag = "STYLE|" & BS_NULL
        .Add(, , "BS_HATCHED").Tag = "STYLE|" & BS_HATCHED
        .Add(, , "BS_PATTERN").Tag = "STYLE|" & BS_PATTERN
        .Add(, , "BS_PATTERN8x8").Tag = "STYLE|" & BS_PATTERN
       
    End With



End Sub
Public Sub Edit(EditThis As CBrush)
    'Edit the given brush.
    Set mBrushEdit = EditThis
    
    'load the data from the brush.
    
    
    
    Me.Show
End Sub
