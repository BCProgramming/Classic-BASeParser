VERSION 5.00
Begin VB.UserControl UCCoreSettings 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtDisabled 
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Tag             =   "REG:Disabled Operators"
      Top             =   2295
      Width           =   3975
   End
   Begin VB.TextBox txtDisabled 
      Height          =   285
      Index           =   0
      Left            =   315
      TabIndex        =   4
      Tag             =   "REG:Disabled Functions"
      Top             =   1620
      Width           =   3975
   End
   Begin VB.CheckBox chkDIsableComplex 
      Caption         =   "&Disable Complex Number support"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Tag             =   "REG:CplxDisable"
      Top             =   855
      Width           =   2805
   End
   Begin VB.Label lblDisabled 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disabled Operators:"
      Height          =   195
      Index           =   1
      Left            =   315
      TabIndex        =   3
      Top             =   2070
      Width           =   1395
   End
   Begin VB.Label lblDisabled 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disabled Functions:"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   2
      Top             =   1395
      Width           =   1395
   End
   Begin VB.Label LblCoreInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"UCCoreSettings.ctx":0000
      Height          =   585
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   4545
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "UCCoreSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements BASeParserXP.ISettingsPage

Private Sub ISettingsPage_DesistFromKey(ByVal RegKeyLoad As String)
Dim loopcontrol As Object
With parsersettings.RegObject

    Debug.Print "WAS ASKED TO DESIST FROM KEY:" & RegKeyLoad
    'txtDisabled.Text = .ValueEx(HHKEY_CURRENT_USER, RegKeyLoad, "TEMP", RREG_SZ, "texthere")
    'chkDIsableComplex.Value = (.ValueEx(HHKEY_CURRENT_USER, RegKeyLoad, "CplxDisable", RREG_SZ, 0) <> 0)
    'txtDisabled(0).Text = .ValueEx(HHKEY_CURRENT_USER, regkeyload, "Disabled Functions", RREG_SZ, "")
    'txtDisabled(1).Text = .ValueEx(HHKEY_CURRENT_USER, regleyload, "Disabled Operators", RREG_SZ, "")
     For Each loopcontrol In UserControl.Controls
            If Left$(loopcontrol.Tag, 4) = "REG:" Then
                 .ValueEx(HHKEY_CURRENT_USER, RegKeyLoad, Mid$(loopcontrol.Tag, 5), RREG_SZ, "") = GetControlData(loopcontrol)
            End If
        Next
End With
End Sub

Private Sub ISettingsPage_PersistToKey(ByVal RegKeySave As String)
Dim loopcontrol As Object
    With parsersettings.RegObject
    
        Debug.Print "WAS ASKED TO PERSIST TO KEY:" & RegKeySave
        For Each loopcontrol In UserControl.Controls
            If Left$(loopcontrol.Tag, 4) = "REG:" Then
                Call SetControlData(loopcontrol, .ValueEx(HHKEY_CURRENT_USER, RegKeySave, Mid$(loopcontrol.Tag, 4), RREG_SZ, ""))
            
            End If
        Next
'        .ValueEx(HHKEY_CURRENT_USER, RegKeySave, "TEMP", RREG_SZ, "texthere") = txtDisabled.Text
'       SavePage RegKeySave
'       .ValueEx(HHKEY_CURRENT_USER, RegKeySave, "CplxDisable", RREG_SZ, 0) = (chkDIsableComplex.Value <> 0)
'       .ValueEx(HHKEY_CURRENT_USER, regleySave, "Disabled Functions", RREG_SZ, "") = txtDisabled(1).Text
'       .ValueEx(HHKEY_CURRENT_USER, regkeySave, "Disabled Operators", RREG_SZ, "") = txtDisabled(2).Text
'
    End With
End Sub



Private Function ISettingsPage_QueryInformation(ByVal QueryType As BASeParserXP.EQueryInformationConstants) As Variant
    Dim r As Variant
    Select Case QueryType
    
        Case Query_UniqueName
            r = "BASeParserXP.BPCoreOpFunc"
        Case Query_DisplayName
            r = "Core"
        Case Query_DisplayIcon
            '
        Case Query_Hidden
            r = False
    End Select
    ISettingsPage_QueryInformation = r

End Function




Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    UserControl.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub SetControlData(ForControl As Object, DatSet As Variant)
    If TypeOf ForControl Is CheckBox Then
        ForControl = DatSet = 1
    Else
        ForControl = DatSet
    End If


End Sub

Private Function GetControlData(ForControl As Object) As Variant
    If TypeOf ForControl Is CheckBox Then
        GetControlData = ForControl = 1
    Else
        GetControlData = ForControl
    End If


End Function
