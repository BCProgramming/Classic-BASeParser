VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   495
      TabIndex        =   1
      Top             =   315
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Height          =   3390
      Left            =   585
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "permute.frx":0000
      Top             =   1395
      Width           =   4875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()




End Sub
Public Function GetPermute() As String
Dim Chars As String
Chars = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Dim currpermute As String
Dim outstr As String

Dim Char1 As Byte, char2 As Byte, char3 As Byte


For Char1 = 1 To Len(Chars) - 1
    For char2 = 1 To Len(Chars) - 1
        For char3 = 1 To Len(Chars) - 1
            outstr = outstr & Mid$(Chars, Char1, 1) & Mid$(Chars, char2, 1) & Mid$(Chars, char3, 1) & vbCrLf
            
    
        Next
    Next
    
Next
            
GetPermute = outstr

End Function
