VERSION 5.00
Begin VB.Form FrmDenied 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NICE TRY"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1665
      TabIndex        =   1
      Top             =   1395
      Width           =   1230
   End
   Begin VB.Label lbldenied 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS DENIED."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   495
      TabIndex        =   0
      Top             =   135
      Width           =   4035
   End
End
Attribute VB_Name = "FrmDenied"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

