VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPluginInstWiz 
   Caption         =   "Plugin Installation Wizard"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   5715
      TabIndex        =   3
      ToolTipText     =   "Cancel the wizard."
      Top             =   4095
      Width           =   1005
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "<< &Previous"
      Height          =   420
      Left            =   3555
      TabIndex        =   2
      ToolTipText     =   "Return to the previous step."
      Top             =   4095
      Width           =   1005
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next >>"
      Height          =   420
      Left            =   4590
      TabIndex        =   1
      ToolTipText     =   "&Move on to the next step."
      Top             =   4095
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Height          =   4425
      Left            =   45
      ScaleHeight     =   4365
      ScaleWidth      =   1890
      TabIndex        =   0
      Top             =   45
      Width           =   1950
   End
   Begin VB.Frame FraSteps 
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   3
      Left            =   2025
      TabIndex        =   18
      Top             =   -45
      Width           =   4695
      Begin MSComctlLib.ListView lvwclasses 
         Height          =   1500
         Left            =   135
         TabIndex        =   20
         Top             =   1125
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   2646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label LblTypelib 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Examining File..."
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   315
         Width           =   4515
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FraSteps 
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   2
      Left            =   2025
      TabIndex        =   16
      Top             =   -45
      Width           =   4695
      Begin VB.Label lblProgIDsuccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ProgID success message."
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   135
         Width           =   1845
      End
   End
   Begin VB.Frame FraSteps 
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   1
      Left            =   2025
      TabIndex        =   7
      Top             =   -45
      Width           =   4695
      Begin VB.CommandButton cmdBrowseTlb 
         Caption         =   "&Browse..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3690
         TabIndex        =   15
         Top             =   2430
         Width           =   1005
      End
      Begin VB.TextBox txtfilename 
         Enabled         =   0   'False
         Height          =   285
         Left            =   315
         TabIndex        =   14
         Top             =   2475
         Width           =   3300
      End
      Begin VB.OptionButton OptMethod 
         Caption         =   "By enumerating a Type Library"
         Height          =   285
         Index           =   1
         Left            =   45
         TabIndex        =   12
         Top             =   1755
         Value           =   -1  'True
         Width           =   2670
      End
      Begin VB.TextBox TxtProgID 
         Height          =   285
         Left            =   765
         TabIndex        =   10
         Top             =   1215
         Width           =   2940
      End
      Begin VB.OptionButton OptMethod 
         Caption         =   "With a specific ProgID"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   765
         Width           =   3165
      End
      Begin VB.Label lblfilename 
         AutoSize        =   -1  'True
         Caption         =   "&Filename:"
         Height          =   195
         Left            =   315
         TabIndex        =   13
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label lblProg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&ProgID:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How do you want to install the plugin?"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   2685
      End
   End
   Begin VB.Frame FraSteps 
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   0
      Left            =   2025
      TabIndex        =   4
      Top             =   -45
      Width           =   4695
      Begin VB.ComboBox cboSet 
         Height          =   315
         Left            =   900
         TabIndex        =   21
         Text            =   "Default"
         Top             =   1935
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   $"FrmPluginInstWiz.frx":0000
         Height          =   585
         Left            =   45
         TabIndex        =   6
         Top             =   855
         Width           =   4245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1665
      End
   End
   Begin VB.Frame FraSteps 
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   5
      Left            =   2025
      TabIndex        =   23
      Top             =   -45
      Width           =   4695
   End
   Begin VB.Frame FraSteps 
      BorderStyle     =   0  'None
      Height          =   3840
      Index           =   4
      Left            =   2025
      TabIndex        =   22
      Top             =   -45
      Width           =   4695
      Begin VB.Label lblSingleplugin 
         AutoSize        =   -1  'True
         Caption         =   "Plugin installer"
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.Line lnelight 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   1980
      X2              =   6705
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Line lnelight 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   0
      X1              =   1980
      X2              =   6660
      Y1              =   90
      Y2              =   90
   End
End
Attribute VB_Name = "FrmPluginInstWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCurrStep As Long
Private Enum StatusCodeConstants
    Status_Created
    Status_CreateFail
    Status_NoInterface


End Enum
Private Type GatherData
    ProgID As String
    ObjValue As Object
    EvalValue As IEvalEvents
    CoreValue As ICorePlugin    'optimally I'd use a Union (for the two interface pointers). but, hey, this ain't C.
    StatusCode As StatusCodeConstants
End Type
Private mGathered() As GatherData
'Private mGatheredPIDS() As String
'Private mGatheredObjs() As Object

Private Sub cmdprevious_Click()
Dim gohere As Long
CmdNext.Enabled = True
Select Case mCurrStep
    Case 0
        Beep
    Case 3
        gohere = 1
    Case 4
        gohere = 2
    Case 5
        gohere = 3
    Case Else
        gohere = mCurrStep - 1
End Select
ShowStep gohere
End Sub
'Index table (pages:)
'index 0: welcome screen
'Index 1: select between progID and TLB.
'index 2: instantiates progID in txtprogID. display appropriate message. On failure disable next button.
'           (next button should go to index 4).
'index 3: attempt to use TLI to enumerate the exposed classes in txtfilename. Display status message, and attempt
'to instantiate each progID we build (remember to remove those damn preceding underscores), and test
'to see if it exposes either of ICorePlugin or IEvalEvents. If so add it to the two arrays of found items, one of which is their ProgID, the other is the actual object.
'when done, if no valid items where found, display a message as such. Otherwise, display a listview, allowing user
'to toggle wether to install certain items.
'when complete, go on to step 5, which will install the items from step 3.
'index 4: this should install the progID in txtprogID. it has been validated and is a plugin.
'index 5: install ProgID's gathered from Step 3, display appropriate information.
Sub ShowStep(StepNum As Long)
'“”
Dim instObj As Object
Dim SupportsEval As Boolean
Dim SupportsCore As Boolean
Dim castEval As IEvalEvents, setcount As Long
Dim castcore As ICorePlugin
Dim csetnames() As String
mCurrStep = StepNum
FraSteps(StepNum).ZOrder vbBringToFront
'Now, perform relevant action with the given step.
Dim I As Long
        CmdNext.Caption = "&Next>>"
        CmdCancel.Enabled = True
Select Case StepNum
     Case 0
        'populate cbo steps with the configuration set.
        csetnames = parsersettings.GetConfigSetNames(setcount)
        cboSet.Clear
        For I = LBound(csetnames) To setcount
            cboSet.AddItem csetnames(I)
        
        Next I
    Case 2
        On Error Resume Next
        Set instObj = CreateObject(TxtProgID.Text)
        If Err.Number <> 0 Then
            lblProgIDsuccess.Caption = "The specified ProgID, """ & TxtProgID.Text & """ is not valid. press ""previous"" to re-enter."
            CmdNext.Enabled = False 'they cannot go forward, since
                                    'this step has failed.
         Else
            'success. OK- well, Object creation, anyway.
            'but can we cast it?
            On Error Resume Next
            Set castEval = instObj
            If Err <> 0 Then
                lblProgIDsuccess.Caption = """" & TxtProgID.Text & """ supports IEvalEvents."
                SupportsEval = True
            Else
                lblProgIDsuccess.Caption = """" & TxtProgID.Text & """ does NOT support IEvalEvents."
            End If
            Err.Clear
            Set castcore = instObj
            If Err <> 0 Then
                lblProgIDsuccess.Caption = lblProgIDsuccess.Caption & vbCrLf & """" & TxtProgID.Text & " supports ICorePlugin."
            Else
                'disable next, since it isn't a plugin.
                'CmdNext.Enabled = False
                If SupportsEval = False Then
                    lblProgIDsuccess.Caption = """" & TxtProgID.Text & """ Is Not a valid Plugin Object. Press Previous to re-enter."
                Else
                    lblProgIDsuccess.Caption = """" & TxtProgID.Text & """ does NOT support ICorePlugin."
                
                End If
            End If
        
        End If
    Case 3
        'enumerate using TLI, the library file pointed to by
        'txtfilename.text.
        On Error Resume Next
        Dim Tlbinfo As TypeLibInfo
        Dim loopclass As CoClassInfo, buildID As String
        Dim classcount As Long
        LblTypelib.Caption = "Loading typelib information from """ & txtfilename.Text & """."
        Set Tlbinfo = TLI.TypeLibInfoFromFile(txtfilename.Text)
        If Err <> 0 Then
            LblTypelib.Caption = "Failed to load Typelib information from the file, """ & txtfilename.Text & """."
            CmdNext.Enabled = True
        Else
            
            LblTypelib.Caption = "Typelib information loaded successfully. Beginning enumeration..."
            classcount = 0
            For Each loopclass In Tlbinfo.CoClasses
            
                'OK- we need to populate two variables,
                'Private mGatheredPIDS() As String
                'Private mGatheredObjs() As Object
                'while we are enumerating, we cannot really make SURE it supports IEvalEvents or ICoreplugin,
                'at least not without instantiating the object itself. (we could use InterfaceInfo to find out, but
                'it will simply be a string comparision. What if the plugin was made for a older, non-binary compatible version
                'of BASeParser?
                'build the progID.
                'create the object.
                'make sure it is a plugin.
                '
                'create the ProgID as appropriate...
                'buildID = Tlbinfo.Name & "."
                buildID = loopclass.Name
                If Left$(buildID, 1) = "_" Then buildID = Mid$(buildID, 2)
                If Right$(buildID, 1) = "_" Then buildID = Left$(buildID, Len(buildID) - 1)
                buildID = Tlbinfo.Name & "." & buildID
                LblTypelib.Caption = buildID
                'OK- now try to instantiate it.
                ReDim Preserve mGathered(classcount)
                On Error Resume Next
                Set instObj = CreateObject(buildID)
                With mGathered(classcount)
                If Err <> 0 Then
                    'shoot.
                    .StatusCode = Status_CreateFail
                    .ProgID = buildID
                    Set .ObjValue = Nothing
                Else
                    'creation succeeded. good.
                    If parsersettings.GetPluginType(instObj, .EvalValue, .CoreValue) <> PluginType_Invalid Then
                        'Yah, alright! Headlong! huh?
                        'It's a valid Plugin, so we can populate the UDT value to it's max.
                        Set .ObjValue = instObj
                        .ProgID = buildID
                        .StatusCode = Status_Created
                    Else
                        'isn't a plugin.
                        .StatusCode = Status_NoInterface
                        .ProgID = buildID
                        Set .ObjValue = Nothing
                    
                    
                    End If
                
                
                End If
                
                
                End With
                
                'ReDim Preserve mGatheredPIDS(classcount)
                'ReDim Preserve mGatheredObjs(classcount)
                classcount = classcount + 1
                'ReDim Preserve mGathered(classcount)
                
            Next
            'OK, now we re-iterate through the array from 0 to classcount-1.
  
            Dim makeitem As ListItem
            With Me.lvwclasses
            
                .ListItems.Clear
                .ColumnHeaders.Clear
           
                'columns:
                'Name (if applicable- the object needs to support the interface)
                'ProgID
                'Description
                'install status (Ready to Install,do not install,Already Installed)
               'when the item is already installed, Already Installed is the only value displayed
               'here.
                
                
                
                With .ColumnHeaders
                    .Add , "NAME", "Name"
                    .Add , "PROGID", "ProgID"
                    .Add , "DESCRIPTION", "Description"
                    .Add , "STATUS", "Install Status"
                    
                
                End With
                
                
                
                
            End With
            Dim castGUI As IPluginUIData
            For I = 0 To classcount - 1
                'create a item in the listview.
                If mGathered(I).StatusCode = Status_Created Then
                    Set makeitem = Me.lvwclasses.ListItems.Add
                    makeitem.ListSubItems.Add , "PROGID"
                    makeitem.ListSubItems.Add , "DESCRIPTION"
                    makeitem.ListSubItems.Add , "STATUS"
                    If Not parsersettings.IsPluginInstalled(buildID, cboSet.Text) Then
                        makeitem.Tag = 0
                        makeitem.ListSubItems("STATUS").Text = "Ready To Install"
                        
                    Else
                        makeitem.Tag = 1
                        makeitem.ListSubItems("STATUS").Text = "Already Installed"
                    
                    End If
                    With makeitem
                        makeitem.ListSubItems("PROGID") = mGathered(I).ProgID
                        If mGathered(I).StatusCode = Status_Created Then
                            If Not mGathered(I).EvalValue Is Nothing Then
                                makeitem.Text = mGathered(I).EvalValue.Name
                                Set castGUI = mGathered(I).EvalValue.GetPluginUIData
                            ElseIf Not mGathered(I).CoreValue Is Nothing Then
                                Set castGUI = mGathered(I).CoreValue.GetPluginUIData
                            
                            End If
                            If Not castGUI Is Nothing Then
                                makeitem.ListSubItems("DESCRIPTION").Text = castGUI.Description
                                
                            End If
                        End If
                    
                    End With
       
                End If
                
                
                
                
                
            Next I
            
        
        End If
    
    
    Case 4
    'install plugin in txtprogID.
    'then display message.
    Call parsersettings.InstallPlugin(TxtProgID.Text, cboSet.Text)
    lblSingleplugin.Caption = "The plugin """ & TxtProgID.Text & """ was successfully installed."
    
    Case 5
        'install all the plugins described in lvwplugins that have a tag that is 0.
        Dim loopitem As ListItem
        For Each loopitem In lvwclasses.ListItems
            'progID is in Loopitem.ListSubItems("PROGID").text
            If loopitem.Tag = -1 Then
                parsersettings.InstallPlugin (loopitem.ListSubItems("ProgID").Text)
            
            
            End If
        Next
        
    Case 6
        'final step.
        CmdNext.Caption = "&Finish"
        CmdCancel.Enabled = False
End Select



End Sub


Private Sub CmdNext_Click()
'the second step has two paths, one where they pick a progID and one where they select
'the Type library. the ProgID one uses page (index) number 2, whereas the typelib bersion is index 3.
'Index table (pages:)
'index 0: welcome screen
'Index 1: select between progID and TLB.
'index 2: instantiates progID in txtprogID. display appropriate message. On failure disable next button.
'           (next button should go to index 4).
'index 3: attempt to use TLI to enumerate the exposed classes in txtfilename. Display status message, and attempt
'to instantiate each progID we build (remember to remove those damn preceding underscores), and test
'to see if it exposes either of ICorePlugin or IEvalEvents. If so add it to the two arrays of found items, one of which is their ProgID, the other is the actual object.
'when done, if no valid items where found, display a message as such. Otherwise, display a listview, allowing user
'to toggle wether to install certain items.
'when complete, go on to step 5, which will install the items from step 3.
'index 4: this should install the progID in txtprogID. it has been validated and is a plugin.
'index 5: install ProgID's gathered from Step 3, display appropriate information.
Select Case mCurrStep
    Case 1
       If OptMethod(0).Value Then
        'step 3
        ShowStep 2
    ElseIf OptMethod(1).Value Then
        ShowStep 3
    
    End If
    Case 2
        ShowStep 4
    Case 3
        ShowStep 5
    Case 4
        ShowStep 6      '6 is the final page.
    Case 5
        ShowStep 6
    Case Else
        ShowStep mCurrStep + 1
        

End Select


End Sub



Private Sub Form_Load()
    ShowStep 0
End Sub

Private Sub Label4_Click()

End Sub

Private Sub lvwclasses_DblClick()
    Dim numTag As Long
    If lvwclasses.SelectedItem Is Nothing Then Exit Sub
    If lvwclasses.SelectedItem.Tag = 1 Then Exit Sub
    lvwclasses.SelectedItem.Tag = Not lvwclasses.SelectedItem.Tag
    numTag = CInt(lvwclasses.SelectedItem.Tag)
    lvwclasses.SelectedItem.ListSubItems("STATUS") = Switch(numTag = 1, "Already Installed", numTag = -1, "Ready To Install", numTag = 1, "Do Not Install")
End Sub

Private Sub OptMethod_Click(Index As Integer)
If Index = 1 Then
    txtfilename.Enabled = True
    cmdBrowseTlb.Enabled = True
    TxtProgID.Enabled = False
Else
    TxtProgID.Enabled = True
    cmdBrowseTlb.Enabled = False
    txtfilename.Enabled = False
End If
End Sub
