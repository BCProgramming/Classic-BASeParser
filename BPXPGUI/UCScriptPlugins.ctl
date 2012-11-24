VERSION 5.00
Begin VB.UserControl UCScriptPlugins 
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   ScaleHeight     =   4290
   ScaleWidth      =   5805
   Begin VB.Frame fraFiles 
      Caption         =   "&Module Files"
      Height          =   3075
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   5595
   End
   Begin VB.Frame FraModule 
      Caption         =   "Module:"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3075
      Begin VB.ComboBox cboModules 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   2835
      End
   End
End
Attribute VB_Name = "UCScriptPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ISettingsPage
'First off, the root of all the data is the key:
'<SetKey(ConfigSet)>\BPCoreFunc.CScriptPlugins\
'underneath this key, each key represents a different Script Control to be created.

'+<SetKey(ConfigSet)>\BPCoreFunc.CScriptPlugins\
'+----<DescriptionN>
'       +----Modules
'              +---<ModuleNamen> has description value (optional)
'                   +---FileKeyn
'                       <defaultvalue>=<Filename>

'the name of the key is insignificant, and rather used more for identification purposes. each one has a "Language" value that determines the, err- language to set on the scriptcontrol.
'Underneath we find a Modules key, containing none other then a key for each Module we are to add to that scriptcontrol. Each one contains a set of
'keys one per file, whose default value is a filename to load. whew.

'display in a treeview control.
Private mreguse As New cRegistry
Private mBaseKey As String 'Base Registry key to use as root.

Private Sub InitTvw()
    'Initialize the treeview with data.
    'what data? Well, since the base key points to the Setkey(configset)\Plugins\BPCoreFunc.CScriptPlugins\ we will
    'have the root of the tree be the set of ScriptControl keys. of course, we'll populate those with ghosts and fix them up afterward.
    'basically as it stands we just enumerate the base key and add a new item for each section.

End Sub

Private Sub ISettingsPage_DesistFromKey(ByVal RegKeyLoad As String)
'
End Sub

Private Sub ISettingsPage_PersistToKey(ByVal RegKeySave As String)
'
End Sub

Private Function ISettingsPage_QueryInformation(ByVal QueryType As BASeParserXP.EQueryInformationConstants) As Variant
'
End Function

Private Sub UserControl_InitProperties()
    Set mreguse = New cRegistry
End Sub
