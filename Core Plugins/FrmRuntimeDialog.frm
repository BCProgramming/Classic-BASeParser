VERSION 5.00
Begin VB.Form FrmRuntimeDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BCPRuntimeDialog"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmRuntimeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Click()
Public Function AddControl(ByVal ProgID As String, ByVal ControlName As String, Optional Parent As Object = Nothing) As Object
    Dim ctlreturn As Object
    If Parent Is Nothing Then
        Set ctlreturn = Controls.Add(ProgID, ControlName)
    Else
        Set ctlreturn = Controls.Add(ProgID, ControlName, Parent)
    End If
    Set AddControl = ctlreturn



End Function

Private Sub Form_Click()
    RaiseEvent Click
End Sub
Public Property Get Licenses() As Object
    Set Licenses = VB.Licenses
End Property
