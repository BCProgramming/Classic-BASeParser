VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPlugins 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plugins for ConfigSet <Default>"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   420
      Left            =   4230
      TabIndex        =   3
      Top             =   3285
      Width           =   1095
   End
   Begin VB.CommandButton CmdDisEnable 
      Caption         =   "&Disable"
      Height          =   420
      Left            =   3060
      TabIndex        =   2
      Top             =   3285
      Width           =   1095
   End
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "&Uninstall"
      Height          =   420
      Left            =   1890
      TabIndex        =   1
      Top             =   3285
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwPlugins 
      Height          =   2625
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   4630
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "The following Plugins are currently installed for The configuration set named"
      Height          =   420
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   4875
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCurrSet As String      'currently displayed configuration set.
Private msel As ListItem
Public Sub RefreshList(Optional WithSet As String = "Default")
    'refreshes the display to show those plugins that are within the given set.
    Dim loadplug() As String, PTypes As Variant
    Dim pcount As Long
    Dim CurrPlug As Long
    Dim CurrID As String, typestring As String
    Dim castEval As IEvalEvents
    Dim castcore As ICorePlugin
    Dim TestCreate As Object
    Dim newItem As ListItem
    Me.Caption = "Plugins for Configset <" & WithSet & ">"
    Label1.Caption = "The following plugins are installed for the Configuration set, """ & WithSet & """."
    loadplug = parsersettings.GetPluginProgIDs(WithSet, pcount, False, PTypes)
    
    'the columns:
    'Name
    'Type
    'Status
    'Description
    'Class Name
    'FileName
    lvwPlugins.ListItems.Clear
    lvwPlugins.ColumnHeaders.Clear
    If pcount = 0 Then
        'hmm- THAT is DEFINITELY weird.
        Debug.Assert False
        Exit Sub
    
    End If
    With lvwPlugins.ColumnHeaders
        .Add , "NAME", "Name"
        .Add , "TYPE", "Type"
        .Add , "STATUS", "Status"
        
        .Add , "DESCRIPTION", "Description"
        .Add , "CLASS", "Class Name"
        .Add , "FILE", "File Name"
    
    End With
   
    For CurrPlug = 1 To pcount
     Set newItem = lvwPlugins.ListItems.Add
    With newItem
        .Tag = loadplug(CurrPlug)
        .ListSubItems.Add , "TYPE"
        .ListSubItems.Add , "STATUS"
        .ListSubItems.Add , "DESCRIPTION"
        .ListSubItems.Add , "CLASS"
        .ListSubItems.Add , "FILE"
        CurrID = loadplug(CurrPlug)
            'we told the parserSettings method not to verify, because WE are going to do it ourself.
            'first, see if the progID is even valid.
            On Error Resume Next
            Set TestCreate = CreateObject(loadplug(CurrPlug))
            If Err.Number <> 0 Then
                'failed to create object.
                'A lot of fields will be N/A....
                .Text = "N/A"
                .ForeColor = vbRed
                .ListSubItems("TYPE").Text = "N/A"
                .ListSubItems("DESCRIPTION").Text = "N/A"
                .ListSubItems("STATUS").Text = "Object Not found."
                .ListSubItems("CLASS").Text = loadplug(CurrPlug)
                .ListSubItems("FILE").Text = "N/A"
            Else
                .ListSubItems("FILE").Text = parsersettings.FileNameFromProgID(loadplug(CurrPlug))
                .ListSubItems("CLASS").Text = loadplug(CurrPlug)
                .ListSubItems("STATUS").Text = IIf(parsersettings.isplugindisabled(loadplug(CurrPlug), WithSet), "Disabled", "Enabled")
                Select Case parsersettings.GetPluginType(TestCreate)
                    Case PluginType_Evaluation
                        Set castEval = TestCreate
                        Set castcore = Nothing
                        typestring = "Evaluation"
                    Case PluginType_Core
                        Set castcore = TestCreate
                        Set castEval = Nothing
                        typestring = "Core"
                    Case PluginType_Both
                        Set castcore = TestCreate
                        Set castEval = TestCreate
                        typestring = "Eval/Core"
                End Select
                .ListSubItems("TYPE").Text = typestring
                
                'the description.
                typestring = vbNullString
                If Not castEval Is Nothing Then
                
                    typestring = castEval.GetPluginUIData.Description
                ElseIf Not castcore Is Nothing Then
                    
                    typestring = castcore.GetPluginUIData.Description
                End If
                .ListSubItems("DESCRIPTION").Text = typestring
            
            End If
    End With
    Next
 
End Sub
'IPluginUIData_Description

Private Sub CmdDisEnable_Click()
'determine wether the current item is enabled/or disabled.
Dim StrMsg As String
Dim setBool As Boolean
If parsersettings.isplugindisabled(lvwPlugins.SelectedItem.Tag) Then
    'it is disabled.
    'ask to enable it.
    StrMsg = """" & lvwPlugins.SelectedItem.Tag & """ is currently disabled. do you want to enable it?"
    setBool = True
Else
    'it is Enabled. Ask to disable it.
    StrMsg = "Disabling the plugin """ & lvwPlugins.SelectedItem.Tag & """ will cause it's functionality to be unavailable. Are you sure you want to Disable it?"
    setBool = False
End If
    If MsgBox(StrMsg, vbYesNo, "Confirm Action") = vbYes Then
        Call parsersettings.SetPluginDisableState(lvwPlugins.SelectedItem.Tag, setBool, mCurrSet)
    End If
    Me.RefreshList mCurrSet
    
End Sub

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    RefreshList
    Set lvwPlugins.SelectedItem = lvwPlugins.GetFirstVisible
End Sub

Private Sub lvwPlugins_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim b As Boolean

    b = Not lvwPlugins.SelectedItem Is Nothing
    CmdDisEnable.Enabled = b
    cmdUninstall.Enabled = b
    If b Then
        Set msel = lvwPlugins.SelectedItem
        CmdDisEnable.Caption = IIf(parsersettings.isplugindisabled(msel.Tag, mCurrSet), "&Enable", "&Disable")
        cmdUninstall.Enabled = True
    
    End If
End Sub
