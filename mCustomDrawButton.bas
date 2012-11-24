Attribute VB_Name = "mCustomDrawButton"
Option Explicit
' rect
Private Type RECT
   left As Long
   tOp As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' Text functions:
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Private Const DT_BOTTOM = &H8
    Private Const DT_CENTER = &H1
    Private Const DT_LEFT = &H0
    Private Const DT_CALCRECT = &H400
    Private Const DT_WORDBREAK = &H10
    Private Const DT_VCENTER = &H4
    Private Const DT_TOP = &H0
    Private Const DT_TABSTOP = &H80
    Private Const DT_SINGLELINE = &H20
    Private Const DT_RIGHT = &H2
    Private Const DT_NOCLIP = &H100
    Private Const DT_INTERNAL = &H1000
    Private Const DT_EXTERNALLEADING = &H200
    Private Const DT_EXPANDTABS = &H40
    Private Const DT_CHARSTREAM = 4
    Private Const DT_NOPREFIX = &H800
    Private Const DT_WORD_ELLIPSIS = &H40000

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
   Private Const BF_LEFT = 1
   Private Const BF_TOP = 2
   Private Const BF_RIGHT = 4
   Private Const BF_BOTTOM = 8
   Private Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
   Private Const BDR_RAISEDOUTER = 1
   Private Const BDR_SUNKENOUTER = 2
   Private Const BDR_RAISEDINNER = 4
   Private Const BDR_SUNKENINNER = 8
   Private Const BDR_BUTTONPRESSED = BDR_SUNKENOUTER Or BDR_SUNKENINNER
   Private Const BDR_BUTTONNORMAL = BDR_RAISEDINNER Or BDR_RAISEDOUTER

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT = 1

Private m_ilsIcons As Object
Private m_lIconIndex As Long
Private m_lHDC As Long
Private m_sCaption As String

Public Sub InitCustomDrawButton( _
      ByVal ilsIcons As Object, ByVal lIndex As Long, _
      ByVal lBackHDC As Long, ByVal sCaption As String _
   )
   Set m_ilsIcons = ilsIcons
   m_lIconIndex = lIndex
   m_lHDC = lBackHDC
   m_sCaption = sCaption
   
End Sub
Public Sub DrawButton( _
      ByVal lHWnd As Long, _
      ByVal lHDC As Long, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long, _
      ByVal bPushed As Boolean, _
      ByVal bEnabled As Boolean, ByVal bInFocus As Boolean _
   )
Dim tR As RECT
Dim lY As Long
Dim tTR As RECT
Dim tWR As RECT
Dim tP As POINTAPI

   tR.left = lLeft: tR.tOp = lTop
   tR.Right = lRight: tR.Bottom = lBottom
   If (bPushed) Then
      lLeft = lLeft + 1
      lTop = lTop + 1
   End If
   ' Fill the background with a bitmap:
   GetWindowRect lHWnd, tWR
   tP.x = tWR.left: tP.y = tWR.tOp
   ScreenToClient GetParent(lHWnd), tP
   BitBlt lHDC, lLeft, lTop, lRight - lLeft, lBottom - lTop, m_lHDC, tP.x, tP.y, vbSrcCopy
   ' Draw the border:
   If (bPushed) Then
      DrawEdge lHDC, tR, BDR_SUNKENOUTER, BF_RECT
   Else
      DrawEdge lHDC, tR, BDR_RAISEDINNER, BF_RECT
   End If
   SetBkMode lHDC, TRANSPARENT
   ' Draw focus rectangle:
   If (bInFocus) Then
      LSet tWR = tR
      InflateRect tWR, -2, -2
      If bPushed Then
         OffsetRect tWR, 1, 1
      End If
      DrawFocusRect lHDC, tWR
   End If
   
   ' Draw the icon:
   If Not (m_ilsIcons Is Nothing) Then
      ' Assume 16x16 ils here.
      lY = (tR.Bottom - tR.tOp - 16) \ 2
      'ImageList_Draw m_hIml, m_lIconIndex, lHDC, tR.left + 4 + Abs(bPushed), lY + Abs(bPushed), ILD_TRANSPARENT
      m_ilsIcons.ListImages(m_lIconIndex + 1).Draw lHDC, (tR.left + 4 + Abs(bPushed)) * Screen.TwipsPerPixelX, (lY + Abs(bPushed)) * Screen.TwipsPerPixelY, ILD_TRANSPARENT
      tR.left = tR.left + 6 + 16
   End If
   ' Draw the text:
   InflateRect tR, -1, -1
   LSet tTR = tR
   DrawText lHDC, m_sCaption, -1, tTR, DT_LEFT Or DT_WORDBREAK Or DT_CALCRECT
   If (tTR.Bottom < tR.Bottom) Then
      tR.tOp = ((tR.Bottom - tR.tOp) - (tTR.Bottom - tTR.tOp)) \ 2
   End If
   OffsetRect tR, Abs(bPushed), Abs(bPushed)
   DrawText lHDC, m_sCaption, -1, tR, DT_LEFT Or DT_WORDBREAK Or DT_WORD_ELLIPSIS
   
End Sub

