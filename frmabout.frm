VERSION 5.00
Begin VB.Form FrmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6600
   Icon            =   "frmabout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrRollback 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5940
      Top             =   3645
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5490
      TabIndex        =   0
      Top             =   4095
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit http://www.GlitchPC.com"
      Height          =   195
      Left            =   1740
      TabIndex        =   2
      Top             =   4260
      Width           =   2160
   End
   Begin VB.Image imgGlitch 
      Height          =   1020
      Left            =   4440
      Picture         =   "frmabout.frx":385A
      Stretch         =   -1  'True
      Top             =   3540
      Width           =   990
   End
   Begin VB.Image ImgBP 
      Height          =   3600
      Left            =   3465
      Top             =   0
      Visible         =   0   'False
      Width           =   6600
   End
   Begin VB.Image ImgLogo 
      Height          =   3615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   6585
   End
   Begin VB.Label LblVerCopy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   3645
      Width           =   480
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRolldirection As Integer
Private midling As Boolean
Private mEggSteps As Long
'Private ClickedItem As cButton

Private Property Let EggSteps(ByVal Vdata As Long)
    mEggSteps = Vdata
    'Debug.Print "EGGSTEPS = " & Vdata
End Property
Private Property Get EggSteps() As Long
    EggSteps = mEggSteps
End Property

'Private Sub PopUpper_ButtonClick(index As Integer, btn As vbalCmdBar6.cButton)
'    Set ClickedItem = btn
'End Sub
'
'Private Sub PopUpper_RequestNewInstance(index As Integer, ctl As Object)
'    Load PopUpper(PopUpper.UBound + 1)
'    Set ctl = PopUpper(PopUpper.UBound)
'End Sub



'Public Function ShowPopupMenu(MenuItems As Variant, Optional ByVal ReturnSelArray As Boolean = True) As Variant
''step one- generate the XML.
'Dim StrGenerate As String, Pos As POINTAPI
'    Dim posX As Long, PosY As Long
'
'    'we need to kind of duplicate the loop for PopupMenuEx.
'    'we need an outer menu, since it needs to actually all fit as a single pop-up.
'
'    StrGenerate = "<MENUSET NAME=""MAIN""><MENU NAME=""POPUP"" CAPTION=""TEST"">"
'    'Now- MenuItems should be an array- irregardless, we send the entire array off too
'    'PopupMenu GetXMLStr
'    StrGenerate = StrGenerate & GetXMLStr(MenuItems)
'    StrGenerate = StrGenerate & "</MENU></MENUSET>"
'    On Error Resume Next
'    'PopUpper(0).CommandBars.Remove "POPUP"
'    Err.Clear
'    GetCursorPos Pos
'    Set ClickedItem = Nothing
'    Call ModCmdBarXML.CreateCommandBarsFromXML(StrGenerate, PopUpper(0))
'    Call PopUpper(0).ShowPopupMenu(Pos.x, Pos.y, PopUpper(0).CommandBars("POPUP"))
'
'
'    'we havea form-level var named CLICKEDITEM.
'    ShowPopupMenu = ClickedItem.Caption
'
'End Function
Private Sub cmdOK_Click()
'save memory.
    Unload Me
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    'Call DoExplorerMenu(Me.hWnd, "C:\windows\system32\shell32.dll", cmdOK.left + ScaleX(Me.left, vbTwips, vbPixels), cmdOK.top + ScaleY(Me.top, vbTwips, vbPixels))


End If
End Sub

Private Sub Form_Click()
   mRolldirection = mRolldirection * -1
   midling = True
   idle
End Sub

Private Sub Form_Initialize()
   Form_Load
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyB And EggSteps = 2 Then
        EggSteps = 3
    ElseIf KeyCode = vbKeyA And EggSteps = 3 Then
        EggSteps = 4
    ElseIf KeyCode = vbKeyS And EggSteps = 4 Then
        EggSteps = 5
    ElseIf KeyCode = vbKeyE And EggSteps = 5 Then
        EggSteps = 6
    ElseIf KeyCode = vbKeyC And EggSteps = 6 Then
        EggSteps = 7
    ElseIf KeyCode = vbKeyA And EggSteps = 7 Then
        EggSteps = 8
    ElseIf KeyCode = vbKeyM And EggSteps = 8 Then
        EggSteps = 9
    ElseIf KeyCode = vbKeyP And EggSteps = 9 Then
        EggSteps = 10
    ElseIf KeyCode = vbKeyF7 And EggSteps = 10 Then
        On Error Resume Next
        SystemInfoEasterEgg
        If Err <> 0 Then
            'a SUPER DUPER SUPER SECRET EASTER EGG!
            'not sure wat though.
        End If
    End If
End Sub


Private Sub Form_Load()
    Set ImgBP.Picture = GetResPicEx("JPEG", "BPLOGO")
    Set ImgLogo.Picture = GetResPicEx("JPEG", "BCLOGO")
    LblVerCopy.Caption = "(VERSION:" & ParserSettings.GetAppVersion & " Copyright 2006-2007 BASeCamp Corporation" & vbCrLf & _
                        "All Rights Reserved."
                        mRolldirection = -1
                        Form_Click
 
End Sub

Private Sub idle()
Static nWidth As Long
Static toggledPic As Boolean
Do
    If Not midling Then Exit Sub
    Dim PicPaint As StdPicture, otherpaint As StdPicture
    If nWidth = 0 Then nWidth = Me.ScaleWidth
    Const mrollamount = 10
        'With ImgLogo
            If nWidth = 1 And mRolldirection = -1 Then
                midling = False
                Exit Sub
            End If
            nWidth = nWidth + (mRolldirection * mrollamount)
        If nWidth >= ScaleWidth Then
            nWidth = ScaleWidth
            midling = False
        ElseIf nWidth <= 0 Then
            nWidth = 1
            toggledPic = Not toggledPic
            midling = False
        End If
        'End With
        'Set PicPaint = IIf(toggledPic, ImgLogo.Picture, Me.Picture)
        Set PicPaint = ImgBP.Picture
    '    Set otherpaint = IIf(PicPaint Is ImgLogo.Picture, imgbp.Picture, ImgLogo.Picture)
    Set otherpaint = ImgLogo.Picture
        
        PaintPicture otherpaint, 0, 0, ScaleWidth, Me.ScaleHeight * (3 / 4)
        
        
        
        
        
        PaintPicture PicPaint, 0, 0, nWidth, Me.ScaleHeight * (3 / 4)
        Me.Refresh
        DoEvents
    Loop
End Sub



Private Sub lblWebpage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        EggSteps = 1
    ElseIf Button = vbRightButton And EggSteps = 1 Then
        EggSteps = 2
    End If
End Sub

Private Sub lblWebpage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set the mouse pointer, duh.
End Sub
