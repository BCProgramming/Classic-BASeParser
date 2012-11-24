VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOwnerDrawButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ButtonContainerhWnd() As Long

End Property
Public Property Get DoOwnerDraw(ByVal lhWnd As Long) As Boolean

End Property
Public Sub DrawItem( _
      ByVal lhWnd As Long, _
      ByVal lHDC As Long, _
      ByRef lLeft As Long, ByRef lTop As Long, _
      ByRef lRight As Long, ByRef lBottom As Long, _
      ByVal bPushed As Boolean, ByVal bChecked As Boolean, _
      ByVal bEnabled As Boolean, ByVal bInFocus As Boolean, _
      ByRef bDoDefault As Boolean _
   )
   
End Sub

