VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl UCScript 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtLanguage 
      Height          =   285
      Left            =   945
      TabIndex        =   3
      Top             =   720
      Width           =   2220
   End
   Begin VB.Frame FraScriptprop 
      Caption         =   "Scripts Within <SCRIPTSET>"
      Height          =   2355
      Left            =   45
      TabIndex        =   4
      Top             =   1080
      Width           =   4605
      Begin VB.CommandButton CmdAddScript 
         Caption         =   "&Add..."
         Height          =   330
         Left            =   3285
         TabIndex        =   7
         ToolTipText     =   "Add a New Script"
         Top             =   720
         Width           =   1230
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   3285
         TabIndex        =   6
         ToolTipText     =   "Remove the selected Script or Script Section"
         Top             =   315
         Width           =   1230
      End
      Begin MSComctlLib.ListView LvwScriptSet 
         Height          =   1995
         Left            =   45
         TabIndex        =   5
         Top             =   270
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   3519
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.ComboBox cboScriptSet 
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Text            =   "cboScriptSet"
      Top             =   135
      Width           =   2130
   End
   Begin VB.Label lblLanguage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Language:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   765
   End
   Begin VB.Label lblconfig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Script Set:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
   Begin VB.Menu mnucontext 
      Caption         =   "&Context"
      Begin VB.Menu mnuAddSection 
         Caption         =   "&Section..."
      End
      Begin VB.Menu mnuaddscript 
         Caption         =   "S&cript..."
      End
   End
End
Attribute VB_Name = "UCScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements BASeParserXP.ISettingsPage
Private mregKey As String
Dim mregistry As New cRegistry
Private Sub CommitSetChanges()
    'commit changes to the current script set.
    'to do so, we construct the appropriate key to write to, delete ALL the values therein,
    'and write one new item for each ListItem in the listview.
    Dim regkeymodify As String
    Dim loopitem As ListItem
    
    regkeymodify = mregKey & "\" & cboScriptSet.Text & "\Scripts"
    'delete all values...
    mregistry.SectionKey = regkeymodify
    On Error Resume Next
    mregistry.DeleteKey
    'ok, now we re-create the file list.
    For Each loopitem In LvwScriptSet.ListItems
        mregistry.ValueEx(HHKEY_CURRENT_USER, mregistry.SectionKey, loopitem.Text, RREG_SZ, "") = ""
    Next
        mregistry.ValueEx(HHKEY_CURRENT_USER, mregKey & "\" & cboScriptSet.Text, "Language", RREG_SZ, "JavaScript") = txtLanguage.Text


End Sub
Private Sub cboScriptSet_Change()

'when the selection here changes, we want to do a few things:
'commit any changes.
'load the set of files and other information from the current set.
Static flCol As Boolean
Dim enumscripts() As String, ScriptCount As Long
Dim I As Long
'!--COMMIT CHANGES HERE--!
'clear the listview...
LvwScriptSet.ListItems.Clear
'LvwScriptSet.ColumnHeaders.Clear
'add the column headers...
If Not flCol Then
    flCol = True
    LvwScriptSet.ColumnHeaders.Clear
    With LvwScriptSet.ColumnHeaders
        'each one is a file.
        .Add , "NAME", "Filename"
        
    
    End With
End If
 mregistry.SectionKey = mregKey & "\" & cboScriptSet.Text & "\Scripts"
    If mregistry.EnumerateValues(enumscripts(), ScriptCount) Then
        'each value = one filename.
        For I = 1 To ScriptCount
           ' LvwScriptSet.ListItems.Add , , EnumScripts(i)
           AddScriptFile enumscripts(I)
        Next I
    Else
        CDebug.Post "No scripts found in set """ & cboScriptSet.Text & """"
    
    End If
End Sub

Private Sub cboScriptSet_Click()
    cboScriptSet_Change
End Sub

'It's a bit round-about, but the heirarchy WILL exist at the given registry keys...

Private Sub CmdAddScript_Click()
    Dim Filedialog As CFileDialog
    Dim FName As String, I As Long
    Dim CurrName As Long, fileub As Long
    Set Filedialog = New CFileDialog
    With Filedialog
        .Caption = "Select script file to add"
        .Filter = "All files(*.*)|*.*"
        .flags = OFN_EXPLORER + OFN_ENABLESIZING
    
    End With
    FName = Filedialog.SelectOpenFile(UserControl.hwnd)
    On Error Resume Next
    
        CDebug.Post "adding file, """ & FName & """"
        AddScriptFile FName
    
End Sub

Private Sub cmdRemove_Click()
'add prompt here probably....
If LvwScriptSet.SelectedItem Is Nothing Then Exit Sub
LvwScriptSet.ListItems.Remove (LvwScriptSet.SelectedItem.Index)
End Sub

'the heirarchy? Oh,
'a group of sections.
'each section contains the names of a group of like-language script files that
'will be loaded TOGETHER into the same code control (yes, multiple sections can define the same language)
'I haven't ironed out the details of how the plugin will allow for the explicit invokation of a function in
'a particular script control, right now it loops until it finds one.

Private Sub ISettingsPage_DesistFromKey(ByVal RegKeyLoad As String)
'
'the GUI:
'setting key (passed in)
'+---Section1
'|      +---Language    = "VBScript"
'|      +---Scripts
'|       +---"%MY DOCUMENTS%\test.vbs"   (I will add ExpandEnvironmentStrings API call to the correct location...
'|        +---"%PROGRAM FILES%\BASeCamp BASeParser Super Plus Pak\Scripts\Math.vbs"
'+---Section2
'       +---Language    = "JavaScript"
'|      +---Scripts
'         +---"jtest.js"
        '....etc....


'use a treeview to portray this data.

Dim EnumSets() As String, setcount As Long
Dim enumscripts() As String, ScriptCount As Long, loopset As Long
mregKey = RegKeyLoad

mregistry.ClassKey = HHKEY_CURRENT_USER
mregistry.SectionKey = RegKeyLoad
'enumerate the sections. there better be some Script Sets here....
txtLanguage.Text = mregistry.ValueEx(HHKEY_CURRENT_USER, mregKey, "Language", RREG_SZ, "JavaScript")
If mregistry.EnumerateSections(EnumSets(), setcount) Then
    'good. there are sets here....
    'for each section, add a new item to the combo box.
    

    'loop through the values within RegkeyLoad & "\Scripts"
    cboScriptSet.Clear
    For loopset = 1 To setcount
        'mregistry.SectionKey = RegKeyLoad & EnumSets(LoopSet) & "\Scripts"
        'If mregistry.EnumerateValues(EnumScripts(), ScriptCount) Then
        cboScriptSet.AddItem EnumSets(loopset)
        
    
    
    Next loopset
        
Else
    'GASP! no sections here!
    CDebug.Post "UCScript failed to find any script sections."
End If

End Sub
Private Sub AddScriptFile(ByVal Filename As String)
    LvwScriptSet.ListItems.Add , , Filename
End Sub
Private Sub ISettingsPage_PersistToKey(ByVal RegKeySave As String)
'
CommitSetChanges
End Sub

Private Function ISettingsPage_QueryInformation(ByVal QueryType As BASeParserXP.EQueryInformationConstants) As Variant
    Dim r As Variant
    Select Case QueryType
    
        Case Query_UniqueName
            r = "BPCoreFunc.CScriptFunctions"
        Case Query_DisplayName
            r = "Scripting"
        Case Query_DisplayIcon
            
        Case Query_Hidden
            r = False
    End Select
    ISettingsPage_QueryInformation = r
End Function



Private Sub UserControl_Initialize()
mnucontext.Visible = False
End Sub

Private Sub UserControl_InitProperties()
mnucontext.Visible = False
End Sub
