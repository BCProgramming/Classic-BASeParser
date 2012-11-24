VERSION 5.00
Object = "*\AGraphControl.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin BPControls.UCConfigSet plugger 
      Height          =   4740
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   8361
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call plugger.LoadData("Default")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    plugger.SaveData
End Sub
