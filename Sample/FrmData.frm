VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtConsumer 
      Height          =   330
      Left            =   585
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   540
      Width           =   4110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objsource As CObject
Dim ObjBindingCollection As BindingCollection

Private Sub Form_Load()
    Set objsource = New CObject
    Set ObjBindingCollection = New BindingCollection
    Set ObjBindingCollection.DataSource = objsource
    ObjBindingCollection.Add TxtConsumer, "Text", "Name"
End Sub
