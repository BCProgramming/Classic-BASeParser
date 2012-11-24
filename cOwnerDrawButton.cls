VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOwnerDrawButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_DRAWITEM = &H2B
Private Const ODT_BUTTON = 4

' Button messages:
Private Const BM_GETCHECK = &HF0&
Private Const BM_SETCHECK = &HF1&
Private Const BM_GETSTATE = &HF2&
Private Const BM_SETSTATE = &HF3&
Private Const BM_SETSTYLE = &HF4&
Private Const BM_CLICK = &HF5&
Private Const BM_GETIMAGE = &HF6&
Private Const BM_SETIMAGE = &HF7&

Private Const BST_UNCHECKED = &H0&
Private Const BST_CHECKED = &H1&
Private Const BST_INDETERMINATE = &H2&
Private Const BST_PUSHED = &H4&
Private Const BST_FOCUS = &H8&

' Button notifications:
Private Const BN_CLICKED = 0&
Private Const BN_PAINT = 1&
Private Const BN_HILITE = 2&
Private Const BN_UNHILITE = 3&
Private Const BN_DISABLE = 4&
Private Const BN_DOUBLECLICKED = 5&
Private Const BN_PUSHED = BN_HILITE
Private Const BN_UNPUSHED = BN_UNHILITE
Private Const BN_DBLCLK = BN_DOUBLECLICKED
Private Const BN_SETFOCUS = 6&
Private Const BN_KILLFOCUS = 7&

' Button Styles:
Private Const BS_3STATE = &H5&
Private Const BS_AUTO3STATE = &H6&
Private Const BS_AUTOCHECKBOX = &H3&
Private Const BS_AUTORADIOBUTTON = &H9&
Private Const BS_CHECKBOX = &H2&
Private Const BS_DEFPUSHBUTTON = &H1&
Private Const BS_GROUPBOX = &H7&
Private Const BS_LEFTTEXT = &H20&
Private Const BS_OWNERDRAW = &HB&
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_RADIOBUTTON = &H4&
Private Const BS_USERBUTTON = &H8&
Private Const BS_ICON = &H40&
Private Const BS_BITMAP = &H80&
Private Const BS_LEFT = &H100&
Private Const BS_RIGHT = &H200&
Private Const BS_CENTER = &H300&
Private Const BS_TOP = &H400&
Private Const BS_BOTTOM = &H800&
Private Const BS_VCENTER = &HC00&
Private Const BS_PUSHLIKE = &H1000&
Private Const BS_MULTILINE = &H2000&
Private Const BS_NOTIFY = &H4000&
Private Const BS_FLAT = &H8000&
Private Const BS_RIGHTBUTTON = BS_LEFTTEXT

' Windows general:
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
' This should return the id of the control in the parent's controls array:
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long


' GDI etc
Private Type RECT
   left As Long
   tOp As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal lHDC As Long, tR As RECT, ByVal eFlag As Long, ByVal eStyle As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type

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
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Const DFC_CAPTION = 1
Private Const DFC_MENU = 2
Private Const DFC_SCROLL = 3
Private Const DFC_BUTTON = 4
'#if(WINVER >= =&H0500)
Private Const DFC_POPUPMENU = 5
'#endif /* WINVER >= =&H0500 */

Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONHELP = &H4

Private Const DFCS_MENUARROW = &H0
Private Const DFCS_MENUCHECK = &H1
Private Const DFCS_MENUBULLET = &H2
Private Const DFCS_MENUARROWRIGHT = &H4
Private Const DFCS_SCROLLUP = &H0
Private Const DFCS_SCROLLDOWN = &H1
Private Const DFCS_SCROLLLEFT = &H2
Private Const DFCS_SCROLLRIGHT = &H3
Private Const DFCS_SCROLLCOMBOBOX = &H5
Private Const DFCS_SCROLLSIZEGRIP = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT = &H10

Private Const DFCS_BUTTONCHECK = &H0
Private Const DFCS_BUTTONRADIOIMAGE = &H1
Private Const DFCS_BUTTONRADIOMASK = &H2
Private Const DFCS_BUTTONRADIO = &H4
Private Const DFCS_BUTTON3STATE = &H8
Private Const DFCS_BUTTONPUSH = &H10

Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const DFCS_CHECKED = &H400

'#if(WINVER >= =&H0500)
Private Const DFCS_TRANSPARENT = &H800
Private Const DFCS_HOT = &H1000
'#endif /* WINVER >= =&H0500 */

Private Const DFCS_ADJUSTRECT = &H2000
Private Const DFCS_FLAT = &H4000
Private Const DFCS_MONO = &H8000

Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Enum EODBorderStyle
   eodBorderThick
   eodBorderThin
End Enum
Public Enum EODStandardButtonStyle
   eodNone = 0
   eodClose = 1
   eodHelp = 2
   eodMax = 3
   eodMin = 4
   eodRestore = 5
   eodDown = 6
   eodUp = 7
   eodLeft = 8
   eodRight = 9
   eodSizeBox = 10
End Enum

Private m_hWnd() As Long
Private m_lhWndCount As Long
Private m_lPtr As Long

Public Sub RedrawButton(ByRef ctlThis As Control)
   InvalidateRectAsNull ctlThis.hwnd, 0, 0
   UpdateWindow ctlThis.hwnd
End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oT As Object
   CopyMemory oT, lPtr, 4
   Set ObjectFromPtr = oT
   CopyMemory oT, 0&, 4
End Property

Public Sub Attach(ByRef IODButtonOwner As IOwnerDrawButton)
   Detach
   m_lPtr = ObjPtr(IODButtonOwner)
   AddhWnd IODButtonOwner.ButtonContainerhWnd
End Sub
Public Sub AddhWnd(ByVal hwnd As Long)
   m_lhWndCount = m_lhWndCount + 1
   ReDim Preserve m_hWnd(1 To m_lhWndCount) As Long
   m_hWnd(m_lhWndCount) = hwnd
   AttachMessage Me, m_hWnd(m_lhWndCount), WM_DRAWITEM
End Sub

Public Sub Detach()
Dim l As Long
   If m_lhWndCount <> 0 Then
      For l = 1 To m_lhWndCount
         DetachMessage Me, m_hWnd(l), WM_DRAWITEM
      Next l
      m_lhWndCount = 0
      Erase m_hWnd()
   End If
End Sub

Public Sub SetBorderStyle(ByRef ctlThis As Control, ByVal eStyle As EODBorderStyle)
   If eStyle = eodBorderThick Then
      RemoveProp ctlThis.hwnd, "vbalODBtn:Border"
   Else
      SetProp ctlThis.hwnd, "vbalODBtn:Border", 1
   End If
End Sub
Public Sub SetStandardButtonStyle(ByRef ctlThis As Control, ByVal eStyle As EODStandardButtonStyle)
   If (eStyle = eodNone) Then
      RemoveProp ctlThis.hwnd, "vbalODBtn:Style"
   Else
      SetProp ctlThis.hwnd, "vbalODBtn:Style", eStyle
   End If
End Sub

Private Sub pDrawButton(tDis As DRAWITEMSTRUCT)
Dim hBr As Long
Dim lState As Long
Dim bPushed As Boolean
Dim bEnabled As Boolean
Dim bChecked As Boolean
Dim bFocus As Boolean
Dim bDoDefault As Boolean
Dim lID As Long
Dim sBuf As String
Dim lLen As Long
Dim tTR As RECT
Dim lBorderPush As Long
Dim lBorderNoPush As Long
Dim eStyle As EODStandardButtonStyle
Dim lType As Long
Dim lStyle As Long
Dim iod As IOwnerDrawButton

   lState = SendMessageLong(tDis.hwndItem, BM_GETSTATE, 0, 0)
   bPushed = ((lState And BST_CHECKED) = BST_CHECKED) Or ((lState And BST_PUSHED) = BST_PUSHED)
   bChecked = (SendMessageLong(tDis.hwndItem, BM_GETCHECK, 0, 0) <> 0)
   bEnabled = IsWindowEnabled(tDis.hwndItem)
   bFocus = (GetFocus() = tDis.hwndItem)
   lID = GetDlgCtrlID(tDis.hwndItem)
   Debug.Print lID, bPushed, bChecked, bEnabled, bFocus
   
   hBr = GetSysColorBrush(vbButtonFace And &H1F&)
   FillRect tDis.hdc, tDis.rcItem, hBr
   DeleteObject hBr
   
   bDoDefault = True
   If (m_lPtr <> 0) Then
      Set iod = ObjectFromPtr(m_lPtr)
      bDoDefault = False
      iod.DrawItem _
         tDis.hwndItem, tDis.hdc, _
         tDis.rcItem.left, tDis.rcItem.tOp, tDis.rcItem.Right, tDis.rcItem.Bottom, _
         bPushed, bChecked, bEnabled, bFocus, bDoDefault
   End If
   
   ' Default draw:
   If bDoDefault Then
      ' Style:
      eStyle = GetProp(tDis.hwndItem, "vbalODBtn:Style")
      If (eStyle <> eodNone) Then
         ' Use draw frame control to render the button:
         Select Case eStyle
         Case eodUp
            lType = DFC_SCROLL
            lStyle = DFCS_SCROLLUP
         Case eodSizeBox
            lType = DFC_SCROLL
            lStyle = DFCS_SCROLLSIZEGRIP
         Case eodRight
            lType = DFC_SCROLL
            lStyle = DFCS_SCROLLRIGHT
         Case eodRestore
            lType = DFC_CAPTION
            lStyle = DFCS_CAPTIONRESTORE
         Case eodMin
            lType = DFC_CAPTION
            lStyle = DFCS_CAPTIONMIN
         Case eodMax
            lType = DFC_CAPTION
            lStyle = DFCS_CAPTIONMAX
         Case eodLeft
            lType = DFC_SCROLL
            lStyle = DFCS_SCROLLLEFT
         Case eodHelp
            lType = DFC_CAPTION
            lStyle = DFCS_CAPTIONHELP
         Case eodDown
            lType = DFC_SCROLL
            lStyle = DFCS_SCROLLDOWN
         Case eodClose
            lType = DFC_CAPTION
            lStyle = DFCS_CAPTIONCLOSE
         End Select
         If (bPushed) Or (bChecked) Then
            lStyle = lStyle Or DFCS_PUSHED
         End If
         If Not (bEnabled) Then
            lStyle = lStyle Or DFCS_INACTIVE
         End If
         DrawFrameControl tDis.hdc, tDis.rcItem, lType, lStyle
      Else
         ' Draw the border:
         If GetProp(tDis.hwndItem, "vbalODBtn:Border") = 0 Then
            ' default (thick)
            lBorderPush = BDR_SUNKENINNER
            lBorderNoPush = BDR_RAISEDOUTER
         End If
         If (bPushed Or bChecked) Then
            DrawEdge tDis.hdc, tDis.rcItem, BDR_SUNKENOUTER Or lBorderPush, BF_RECT
         Else
            DrawEdge tDis.hdc, tDis.rcItem, BDR_RAISEDINNER Or lBorderNoPush, BF_RECT
         End If
         
         ' Get the text (if any)
         sBuf = String$(1024, 0)
         lLen = SendMessageString(tDis.hwndItem, WM_GETTEXT, 1024, sBuf)
         If (lLen > 0) Then
            LSet tTR = tDis.rcItem
            InflateRect tTR, -2, -2
            If bPushed Then
               OffsetRect tTR, 1, 1
            End If
            sBuf = left$(sBuf, lLen)
            If bEnabled Then
               DrawText tDis.hdc, sBuf, -1, tTR, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            Else
               SetTextColor tDis.hdc, GetSysColor(vb3DHighlight And &H1F)
               OffsetRect tTR, 1, 1
               DrawText tDis.hdc, sBuf, -1, tTR, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
               SetTextColor tDis.hdc, GetSysColor(vbButtonShadow And &H1F)
               OffsetRect tTR, -1, -1
               DrawText tDis.hdc, sBuf, -1, tTR, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            End If
         End If
         
         If bFocus Then
            InflateRect tDis.rcItem, -2, -2
            If bPushed Then
               tDis.rcItem.Bottom = tDis.rcItem.Bottom - 1
               tDis.rcItem.Right = tDis.rcItem.Right - 1
               OffsetRect tDis.rcItem, 1, 1
            End If
            DrawFocusRect tDis.hdc, tDis.rcItem
         End If
      End If
   End If
   
End Sub

Private Sub Class_Terminate()
   Detach
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   If (CurrentMessage = WM_DRAWITEM) Then
      ISubclass_MsgResponse = emrConsume
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDis As DRAWITEMSTRUCT
Dim iod As IOwnerDrawButton

   Select Case iMsg
   Case WM_DRAWITEM
      CopyMemory tDis, ByVal lParam, Len(tDis)
      If tDis.CtlType = ODT_BUTTON Then
         If (m_lPtr <> 0) Then
            Set iod = ObjectFromPtr(m_lPtr)
            If iod.DoOwnerDraw(tDis.hwndItem) Then
               pDrawButton tDis
               ISubclass_WindowProc = 1
            Else
               ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
            End If
         Else
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If
      End If
   End Select
End Function
