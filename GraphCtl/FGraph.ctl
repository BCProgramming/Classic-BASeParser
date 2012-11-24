VERSION 5.00
Begin VB.UserControl FGraph 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "FGraph.ctx":0000
   Begin VB.PictureBox PicGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   2475
      ScaleHeight     =   2925
      ScaleWidth      =   3870
      TabIndex        =   0
      Top             =   1800
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   30
      Left            =   1305
      Picture         =   "FGraph.ctx":0312
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "FGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'FGraph- a ActiveX control that wraps the functionality of
'my BASeParser ActiveX Component into a Graphical Function Graph Control.
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private mTestBrush As CBrush
Private mWindow As CWindowDisplay
Private mGraphFunctions As GraphFunctions
Friend Property Get GraphBox() As PictureBox
    Set GraphBox = PicGraph
End Property
Public Property Get Window() As CWindowDisplay
    Set Window = mWindow
End Property
Public Property Set GraphFunctions(Vdata As GraphFunctions)
    Set mGraphFunctions = Vdata
    Set mGraphFunctions.Owner = Me
End Property
Public Property Get GraphFunctions() As GraphFunctions
    Set GraphFunctions = mGraphFunctions
End Property

Private Sub UserControl_InitProperties()
Dim newfunc As CGraphFunction
Set newfunc = New CGraphFunction
newfunc.Expression = "Sin(X)*5"
    Set mGraphFunctions = New GraphFunctions
    Set mWindow = New CWindowDisplay
    Set mWindow.GraphCtl = Me
    
    Set mGraphFunctions.Owner = Me
    mGraphFunctions.Add newfunc
        
    
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo ErrorLoad
    With PropBag
        'Set mGraphFunctions = .ReadProperty("GraphFunctions", New GraphFunctions)
        
         Set mWindow.GraphCtl = Me
        mGraphFunctions.Owner = Me
    End With
ErrorLoad:
    Debug.Print "ERROR IN READPROPERTIES:" & Error$
    Resume Next
End Sub
Public Sub Refresh()

 mWindow.DrawWindow
    GraphFunctions.DrawAll PicGraph.hdc
    
    
 PicGraph.Refresh
End Sub
Private Sub UserControl_Resize()

   
   
    PicGraph.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Refresh
   
    
  
  
   
'    Dim mtestpen As CPen
'  Set mTestBrush = New CBrush
'  Set mtestpen = New CPen
'    mTestBrush.BrushStyle = BS_HATCHED
'    'Set mTestBrush.BrushPic = Image1.Picture
'    mtestpen.PenStyle = PS_DASH
'    Set mtestpen.Brush.BrushPic = Image1.Picture
'    mtestpen.Brush.BrushStyle = BS_PATTERN
'    mtestpen.Width = 5
'    mtestpen.PenType = PS_GEOMETRIC
'
'    mTestBrush.Colour = vbRed
'    mTestBrush.HatchStyle = HS_FDIAGONAL
'    mtestpen.SelectPen UserControl.hdc
'    mTestBrush.SelectBrush UserControl.hdc
'
'    Ellipse UserControl.hdc, 40, 40, 100, 100
'    UserControl.Refresh
'    'mTestBrush.UnSelectBrush UserControl.hdc
'    Set mTestBrush = Nothing
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "GraphFunctions", mGraphFunctions, Nothing
        If mGraphFunctions Is Nothing Then Set mGraphFunctions = New GraphFunctions
    
       
        
    End With
End Sub



Friend Sub GraphFunction(GraphIt As CGraphFunction)
    mWindow.DrawWindow
End Sub


