VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BPXPConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'BPXPAbout: Exposes the FrmConfig form to BASeParser.
'maybe one day I'll delegate the events of the form.
'probably not though.
Private mFrmConfig As FrmConfig

Private Sub Class_Initialize()
    Set mFrmConfig = New FrmConfig
    
End Sub

Private Sub Class_Terminate()
    Set mFrmConfig = Nothing
End Sub


'Displaying stuff-
Public Sub Show(Optional ByVal ShowMode As FormShowConstants = vbModeless, Optional Owner As Object = Nothing)
    If mFrmConfig.Visible Then Exit Sub
        'already shown.
        If Not Owner Is Nothing Then
            mFrmConfig.Show ShowMode, Owner
        Else
            mFrmConfig.Show ShowMode
        End If
    
    
    
End Sub
Public Sub Hide()
    mFrmConfig.Hide
End Sub
