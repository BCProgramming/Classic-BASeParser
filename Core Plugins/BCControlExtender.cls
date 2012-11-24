VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCControlExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type ParameterInfo
    ParameterName As String
    ParameterValue As Variant
End Type
Public Type EventInfo
    EventName As String
    ParameterCount As Long
    EventParameters() As ParameterInfo
End Type


Private WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mPicturebox As PictureBox
Attribute mPicturebox.VB_VarHelpID = -1
Private WithEvents mLabel As Label
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents lFrame As Frame
Attribute lFrame.VB_VarHelpID = -1
Private WithEvents mCommandButton As CommandButton
Attribute mCommandButton.VB_VarHelpID = -1
Private WithEvents mCheckBox As CheckBox
Attribute mCheckBox.VB_VarHelpID = -1
Private WithEvents mOptionButton As OptionButton
Attribute mOptionButton.VB_VarHelpID = -1
Private WithEvents mComboBox As ComboBox
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mListBox As ListBox
Attribute mListBox.VB_VarHelpID = -1
Private WithEvents mhScroll As HScrollBar
Attribute mhScroll.VB_VarHelpID = -1
Private WithEvents mvScroll As VScrollBar
Attribute mvScroll.VB_VarHelpID = -1
Private WithEvents mTimer As Timer
Attribute mTimer.VB_VarHelpID = -1
Private WithEvents mDriveList As DriveListBox
Attribute mDriveList.VB_VarHelpID = -1
Private WithEvents mDirList As DirListBox
Attribute mDirList.VB_VarHelpID = -1
Private WithEvents mFileList As FileListBox
Attribute mFileList.VB_VarHelpID = -1
Private WithEvents mImage As VB.Image
Attribute mImage.VB_VarHelpID = -1
Private WithEvents mData As VB.Data
Attribute mData.VB_VarHelpID = -1
Private WithEvents mOLE As VB.OLE
Attribute mOLE.VB_VarHelpID = -1
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private WithEvents mControlExtender As VBControlExtender
Attribute mControlExtender.VB_VarHelpID = -1



Private mOwner As Object
Public Event ObjectEvent(ObjectEvt As Object, pEventInfo As EventInfo)
Public Property Set Owner(vdata As Object)
    Set mOwner = vdata
End Property


Public Sub Init(ToObject As Object)
    If TypeOf ToObject Is TextBox Then
        Set mTextBox = ToObject
    ElseIf TypeOf ToObject Is PictureBox Then
        Set mPicturebox = ToObject
    ElseIf TypeOf ToObject Is Label Then
        Set mLabel = ToObject
    ElseIf TypeOf ToObject Is Frame Then
       ' Set mframe = ToObject
    ElseIf TypeOf ToObject Is CommandButton Then
        Set mCommandButton = ToObject
    ElseIf TypeOf ToObject Is CheckBox Then
        Set mCheckBox = ToObject
    ElseIf TypeOf ToObject Is OptionButton Then
        Set mOptionButton = ToObject
    ElseIf TypeOf ToObject Is ComboBox Then
        Set mComboBox = ToObject
    ElseIf TypeOf ToObject Is ListBox Then
        Set mListBox = ToObject
    ElseIf TypeOf ToObject Is HScrollBar Then
        Set mhScroll = ToObject
    ElseIf TypeOf ToObject Is VScrollBar Then
        Set mvScroll = ToObject
    ElseIf TypeOf ToObject Is Timer Then
        Set mTimer = ToObject
    ElseIf TypeOf ToObject Is DriveListBox Then
        Set mDriveList = ToObject
    ElseIf TypeOf ToObject Is DirListBox Then
        Set mDirList = ToObject
    ElseIf TypeOf ToObject Is FileListBox Then
        Set mFileList = ToObject
    ElseIf TypeOf ToObject Is Image Then
        Set mImage = ToObject
    ElseIf TypeOf ToObject Is Data Then
        Set mData = ToObject
    ElseIf TypeOf ToObject Is OLE Then
        Set mOLE = ToObject
    ElseIf TypeOf ToObject Is Form Then
        Set mForm = ToObject
    Else
        Set mControlExtender = ToObject
    
    End If
    
End Sub
Private Sub Evt(raisedby As Object, ByVal EventName As String, ParamArray Parameters())
    'creates the "package" and raises the event.
    'the paramarray is in the form "paramname,paramvalue"....
    Dim raisedStruct As EventInfo
    Dim I As Long, currindex As Long
    
    raisedStruct.EventName = EventName
    raisedStruct.ParameterCount = (UBound(Parameters) - 1) \ 2
    If raisedStruct.ParameterCount < 1 Then raisedStruct.ParameterCount = 0
    If raisedStruct.ParameterCount > 0 Then
    ReDim raisedStruct.EventParameters(1 To raisedStruct.ParameterCount)
    End If
    currindex = 1
    For I = 0 To UBound(Parameters) - 1 Step 2
        With raisedStruct.EventParameters(currindex)
            .ParameterName = Parameters(I)
            .ParameterValue = Parameters(I + 1)
        End With
    Next I
    
    
    
    If Not mOwner Is Nothing Then
        Call mOwner.DelegateEvent(Me, raisedby, raisedStruct)
    End If
    RaiseEvent ObjectEvent(raisedby, raisedStruct)



End Sub


'Checkbox Events....


Private Sub mCheckBox_Click()
    Evt mCheckBox, "Click"
End Sub

Private Sub mCheckBox_DragDrop(Source As Control, X As Single, Y As Single)
    Evt mCheckBox, "DragDrop", "Source", Source, "X", X, "Y", Y
End Sub

Private Sub mCheckBox_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Evt mCheckBox, "DragOver", "Source", Source, "X", X, "Y", Y, "State", State
End Sub

Private Sub mCheckBox_GotFocus()
    Evt mCheckBox, "GotFocus"
End Sub

Private Sub mCheckBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Evt mCheckBox, "KeyDown", "KeyCode", KeyCode, "Shift", Shift
End Sub

Private Sub mForm_Click()
Evt mForm, "Click"
End Sub

'
Private Sub mForm_Load()
    
End Sub
