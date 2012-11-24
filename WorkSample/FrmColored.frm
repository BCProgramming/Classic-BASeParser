VERSION 5.00
Begin VB.Form FrmColored 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&GO"
      Height          =   495
      Left            =   5220
      TabIndex        =   7
      Top             =   1020
      Width           =   1155
   End
   Begin VB.TextBox TxtComponent 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Text            =   "X*Y"
      Top             =   960
      Width           =   3315
   End
   Begin VB.TextBox TxtComponent 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Text            =   "Y Mod 255"
      Top             =   600
      Width           =   3315
   End
   Begin VB.TextBox TxtComponent 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Text            =   "X mod 255"
      Top             =   240
      Width           =   3315
   End
   Begin VB.PictureBox PicColored 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   1260
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   0
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Blue:"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Green:"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   2
      Top             =   660
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Red:"
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   1
      Top             =   300
      Width           =   345
   End
End
Attribute VB_Name = "FrmColored"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private mRedFunc As CFunction, mGreenFunc As CFunction, mBlueFunc As CFunction
Private Sub Command1_Click()
Dim XVar As CVariable, YVar As CVariable

Dim Comps(1 To 3), I As Long
Static ParserUse As CParser
If ParserUse Is Nothing Then
    Set ParserUse = New CParser
    Call ParserUse.Create
End If
'Add variables before the functions to prevent parse errors.
Set XVar = ParserUse.Variables.Add("X", 0)
Set YVar = ParserUse.Variables.Add("Y", 0)
Set mRedFunc = ParserUse.Functions.Add(TxtComponent(0), "RED")
Set mGreenFunc = ParserUse.Functions.Add(TxtComponent(1), "GREEN")
Set mBlueFunc = ParserUse.Functions.Add(TxtComponent(2), "BLUE")

Dim CurrX As Long, CurrY As Long
For CurrX = 1 To PicColored.ScaleWidth
    For CurrY = 1 To PicColored.ScaleHeight
        XVar.Value = CurrX
        YVar.Value = CurrY
        Comps(1) = mRedFunc.CallFunc
        Comps(2) = mGreenFunc.CallFunc
        Comps(3) = mBlueFunc.CallFunc
        
        'PicColored.PSet (CurrX, CurrY), RGB(Comps(1), Comps(2), Comps(3))
        Call SetPixel(PicColored.hdc, CurrX, CurrY, RGB(Comps(1), Comps(2), Comps(3)))
        
    
    Next CurrY
    PicColored.Refresh
Next CurrX



End Sub
