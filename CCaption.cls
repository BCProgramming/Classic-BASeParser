VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CCaption: uses a LogFont structure,paths, and text output to create text effects
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private mvarFont As CFont
Private mvarBrush As CBrush
Private MvarPen As CPen
Private mEnabled As Boolean
Public Property Let Enabled(ByVal Vdata As Boolean)
    mEnabled = Vdata
End Property
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Get Brush() As CBrush
    Set Brush = mvarBrush
End Property
Public Property Set Brush(Vdata As CBrush)
    Set mvarBrush = Vdata
End Property
Public Property Get Pen() As CPen
    Set Pen = MvarPen
End Property
Public Property Set Pen(Vdata As CPen)
    MvarPen = Vdata
End Property
Public Property Set Font(Vdata As CFont)
    Set mvarFont = Vdata
End Property
Public Property Get Font() As CFont
    Set Font = mvarFont
End Property
Friend Sub OutputText(ByVal LpText As String, ByVal OnDC As Long, ByVal XPx As Long, ByVal YPx As Long)
    'output "lptext" on "OnDC" at XPx,YPx.
    Dim UsePath As CPath
    Dim OldSel As Long
    Set UsePath = New CPath
    UsePath.Begin OnDC
    'start the path.
    
        OldSel = SelectObject(OnDC, mvarFont.Handle)
        
        TextOut OnDC, XPx, YPx, LpText, Len(LpText)
        
        
        
        SelectObject OnDC, OldSel
        
    UsePath.End_
    UsePath.StrokeAndFill mvarBrush, MvarPen
    Set UsePath = Nothing
    
    
    
End Sub

