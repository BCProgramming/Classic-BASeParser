VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl UCGlobalObjects 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtModifier 
      Height          =   285
      Left            =   540
      TabIndex        =   2
      Top             =   3150
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame FraVariables 
      Caption         =   "&Global objects:"
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   4785
      Begin MSComctlLib.ListView lvwGlobals 
         Height          =   2580
         Left            =   45
         TabIndex        =   1
         Top             =   180
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   4551
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
End
Attribute VB_Name = "UCGlobalObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--1/23/2008 @ 07:18--
'This UserControl is currently dead code.
'a future revision however will use it to configure the Core Plugin for
'Global Objects. Oh yeah- that's dead code right now too.



'Storage structure within registry:
'<Section>\BPCoreFunc.GlobalObjects\<ProgID>\VariableName

'get it? sections within our configuration key (regkeyload/Save) are ProgIDs to objects that are created.
'the variableName key within that contains the Name of the variable created in the parser.

Implements ISettingsPage
Private mreguse As New cRegistry




Private Sub ISettingsPage_DesistFromKey(ByVal RegKeyLoad As String)
'
'load the data!
Dim NumItems As Long
Dim ProgIds() As String
Dim VarNames() As String, currloop As Long
mreguse.ClassKey = HHKEY_CURRENT_USER
mreguse.SectionKey = RegKeyLoad & "\Variables"
Call mreguse.EnumerateSections(ProgIds(), NumItems)
If NumItems > 0 Then
    'alrighty then, add an item for each one.
    With lvwGlobals
        .ListItems.Clear
        
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "ProgID"
        .ColumnHeaders.Add , , "Variable"
        
        For currloop = 1 To NumItems
            VarNames(currloop) = mreguse.ValueEx(HHKEY_CURRENT_USER, RegKeyLoad & "\Variables\" & ProgIds(currloop), "VariableName", RREG_SZ, "")
            If VarNames(currloop) = "" Then
                CDebug.Post "Warning: Object """ & ProgIds(currloop) & """ has no Variable name!", Severity_Warning
            End If
            .ListItems.Add(, , ProgIds(currloop)).SubItems(1) = VarNames(currloop)
        Next currloop
    
    End With

End If



End Sub

Private Sub ISettingsPage_PersistToKey(ByVal RegKeySave As String)
'
'CDebug.Post "Persistence not yet implemented in UCGlobalobjects! :(", Severity_Error
'save the items in the Listview.
Dim sections() As String, scount As Long
Dim LvwItem As ListItem, I As Long
    'alright, go through each section and delete the appropriate values.
    
    With mreguse
        .ClassKey = HHKEY_CURRENT_USER
        .SectionKey = RegKeySave & "\Variables"
        Call .EnumerateSections(sections(), scount)
        
    End With



End Sub

Private Function ISettingsPage_QueryInformation(ByVal QueryType As BASeParserXP.EQueryInformationConstants) As Variant
    
    Dim r As Variant
    Select Case QueryType
    
        Case Query_UniqueName
            r = "BPCoreFunc.GlobalObjects"
        Case Query_DisplayName
            r = "Globals"
        Case Query_DisplayIcon
            '
        Case Query_Hidden
            r = False
    End Select
    ISettingsPage_QueryInformation = r
End Function

Private Function ISettingsPage_UniqueName() As String
    ISettingsPage_UniqueName = "BPCoreFunc.GlobalObjects"
End Function
