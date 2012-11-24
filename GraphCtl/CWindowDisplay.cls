VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWindowDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CWindowDisplay: Handles several things.
Public Enum WindowTypeConstants
    Windowtype_Cartesian = 0      'Standard Cartesian.
    'add more here.
End Enum
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long) As Long

Private WindowBGBrush As CBrush 'Background brush. defaults to "WINDOW_BACKGROUND"

Private mIndAxis As DisplayAxis 'independentAxis (X)
Private mDepAxis As DisplayAxis 'Dependent Axis. (Y)
Private mGraphCtl As FGraph
Private mUpperLeft As PointF
Private mLowerRight As PointF
Private mWindowType As WindowTypeConstants

'mIndAxis and mDepAxis are used when translating
'Grid coordinates to Pixel Coordinates.
'And back again.
Public Property Get IndependentAxis() As DisplayAxis
    Set IndependentAxis = mIndAxis
End Property
Public Property Get DependentAxis() As DisplayAxis
    Set DependentAxis = mDepAxis
End Property
Public Property Let WindowType(ByVal Vdata As WindowTypeConstants)
    mWindowType = Vdata
End Property
Public Property Get WindowType() As WindowTypeConstants
    WindowType = mWindowType
End Property
Friend Property Set GraphCtl(Vdata As FGraph)
    Set mGraphCtl = Vdata
    Set mIndAxis.Owner = Me
    Set mDepAxis.Owner = Me
End Property
Friend Property Get GraphCtl() As FGraph
    Set GraphCtl = mGraphCtl
End Property

Private Sub GetCartesianWindow(ByRef Upperleft As PointF, ByRef Lowerright As PointF)
If Upperleft Is Nothing Then Set Upperleft = New PointF
If Lowerright Is Nothing Then Set Lowerright = New PointF
    Select Case mWindowType
        Case Windowtype_Cartesian
            Set Upperleft = PointF(mIndAxis.Min, mDepAxis.Min)
            Set Lowerright = PointF(mIndAxis.Max, mDepAxis.Max)
        Case Else
            'More CASE's for different graph styles.
            
    End Select
    'there.
End Sub
Friend Function ToPixelCoord(GraphPoint As PointF) As PointF
    Dim retval As PointF
    Set retval = New PointF
    Dim CUpperleft As PointF, ClowerRight As PointF
    
    'translate the given "GraphPoint" into pixel coordinates.
    'first, GraphPoint should be a cartesian graph point.
    'Simple!
    'involves settings ScaleMode on the GraphCtl's PictureBox.
    GetCartesianWindow CUpperleft, ClowerRight
    With mGraphCtl.GraphBox
        'NOTE: rigt now it is assumed to default to pixel mode. I am sure this will be the case.
        .ScaleLeft = CUpperleft.X
        .ScaleTop = CUpperleft.Y
        .ScaleWidth = ClowerRight.X - CUpperleft.X
        .ScaleHeight = ClowerRight.Y - CUpperleft.Y
        
        'now, we simply call the Scale() method,
        'from VBUser to  vbpixels.
        retval.X = .ScaleX(GraphPoint.X - .ScaleLeft, vbUser, vbPixels)
        retval.Y = .ScaleY(GraphPoint.Y - .ScaleTop, vbUser, vbPixels)
        'hooray.
        
        .ScaleMode = vbPixels
    End With
    
    Set ToPixelCoord = retval
End Function
Private Sub DrawBG()
    
    Dim NullPen As CPen
    Dim PBox(1 To 2) As PointF
    Set NullPen = New CPen
    NullPen.PenStyle = PS_NULL
    GetCartesianWindow PBox(1), PBox(2)
    NullPen.SelectPen GraphCtl.GraphBox.hdc
    WindowBGBrush.SelectBrush GraphCtl.GraphBox.hdc
    Set PBox(1) = ToPixelCoord(PBox(1))
    Set PBox(2) = ToPixelCoord(PBox(2))
    Rectangle GraphCtl.GraphBox.hdc, PBox(1).X, PBox(1).Y, PBox(2).X, PBox(2).Y
    WindowBGBrush.UnSelectBrush GraphCtl.GraphBox.hdc
    NullPen.UnSelectPen GraphCtl.GraphBox.hdc
End Sub
Public Sub DrawWindow()
    'Draws the background (when I make it, probably just a brush :P)
    'then the Axe Lines.
    DrawBG
    
    mIndAxis.DrawAxis GraphCtl.GraphBox.hdc
    mDepAxis.DrawAxis GraphCtl.GraphBox.hdc
    
End Sub
Private Sub Class_Initialize()
    Set mDepAxis = New DisplayAxis
    Set mIndAxis = New DisplayAxis
    mDepAxis.Axe = Axis_Dependent
    mIndAxis.Axe = Axis_Independent
    Set WindowBGBrush = New CBrush
    'the dependent axis is reversed if we keep it as is.
    'so, flip it around.
    mDepAxis.Min = 10
    mDepAxis.Max = -10
    WindowBGBrush.SetStock vbWindowBackground
End Sub
