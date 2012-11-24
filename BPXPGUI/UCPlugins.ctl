VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.UserControl UCPlugins 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin vbalDTab6.vbalDTabControl TabPlugins 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   5106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
      Begin VB.PictureBox PicPlugins 
         Height          =   1905
         Index           =   0
         Left            =   45
         ScaleHeight     =   123
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   1
         Top             =   45
         Visible         =   0   'False
         Width           =   3075
      End
   End
End
Attribute VB_Name = "UCPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private mSettingsCtl As VBControlExtender
Private mSettingsCtls As Collection 'collection of VBControlExtenders.

Public Sub Save(ByVal SetFor As String)
    'Tells all of the members of mSettingsCtls to save.
    'it uses the defined method
    'parsersettings.SetKey()
    Dim keysave As String, mreguse As cRegistry
    Dim LoopCtl As ISettingsPage
    Dim loopy As Object
    Set mreguse = New cRegistry
    mreguse.ClassKey = HHKEY_CURRENT_USER
    keysave = parsersettings.SetKey(SetFor)
    For Each loopy In mSettingsCtls
        Set LoopCtl = loopy.object
        Call LoopCtl.PersistToKey(keysave & LoopCtl.QueryInformation(Query_UniqueName))
    Next
End Sub


Public Sub LoadProgIDs(ProgIds() As String, ByVal inset As String)
        'loads the IsettingsPage implementor for each control. If no control is found, simply
        'display the informative message, "There are no User configurable settings for this Plugin."
        'or something.
        
        Dim LoopID As Long
        Dim ObjPlugin As Object
        Dim IEEPlugin As IEvalEvents
        Dim IPlugin As IPluginUIData
        Dim StrctlID As String
        Dim addedtab As vbalDTab6.cTab
        Dim ObjCtl As VBControlExtender
        Dim PageLoad As ISettingsPage
        Dim CurrPic As Long
        Dim Failmessage As String
        Dim I As Long
        CurrPic = 0
        Set mSettingsCtls = New Collection
        TabPlugins.Tabs.Clear
        For I = PicPlugins.UBound To PicPlugins.LBound + 1 Step -1
        
            Unload PicPlugins(I)
        Next
        
    For LoopID = LBound(ProgIds) To UBound(ProgIds)
        On Error Resume Next
        Load PicPlugins(CurrPic)
        Err.Clear
    
        'a lot of tasks here.
        'FIRST:
        'instantiate the ProgID.
        Set ObjPlugin = CreateObject(ProgIds(LoopID))
        'if that fails, NEXT!
        If Err = 0 Then
            'try to cast it to IEvalEvents. If that fails
            'NEXT!
            Set IEEPlugin = ObjPlugin
            If Err = 0 Then
                
                'ask the IEvalEvents interface of the object for it's IPluginUIData pointer.
                Set IPlugin = IEEPlugin.GetPluginUIData()
                If Err = 0 And Not (IPlugin Is Nothing) Then
                    'if that fails, NEXT!
                    'then, ask the IPluginUIData interface for it's settings Page progID.
                    StrctlID = IPlugin.GetSettingsPageProgID()
                    'if that returns "", NEXT!
                    If StrctlID <> vbNullString Then
                        'if not, try to add it to the controls collection of a newly loaded picturebox
                        Set ObjCtl = Controls.Add(StrctlID, "ctl" & Trim$(CurrPic), PicPlugins(CurrPic))
    
                        'control array element. if that fails, NEXT!
                        If Err = 0 And Not (ObjCtl Is Nothing) Then
                            'if that succeeds, then try to cast this new VBControlExtender's Object to ISettingsPage.
                            Set PageLoad = ObjCtl.object
                            'IF that fails, NEXT! if it succeeds, we have done it!
                            If Err = 0 And Not ObjCtl Is Nothing Then
    
                                'add it to the collection, and make it visible.
                                mSettingsCtls.Add ObjCtl
                                ObjCtl.Visible = True
                                
                                'Load the PictureBox.
                                  'add a new tab, whose panel is that picturebox.
                                  Set addedtab = TabPlugins.Tabs.Add(ProgIds(LoopID), , IPlugin.Description)
                                  Set addedtab.Panel = PicPlugins(CurrPic)
                                  addedtab.Enabled = True
                                  
                                  addedtab.CanClose = False
                                  PicPlugins(CurrPic).Tag = mSettingsCtls.Count
                                  'tell "pageload" to desist it's data from the appropriate registry key.
                                  PageLoad.DesistFromKey parsersettings.SetKey(inset) & "PLUGINS\" & PageLoad.QueryInformation(Query_UniqueName)
                            Else
                                'Control isn't an ISettingsPage.
                                'So, Remove it .
                                Controls.Remove "ctl" & Trim$(CurrPic)
                                Failmessage = "The Settings Control for this Plugin does not expose" & vbCrLf & _
                                            "The proper interface."
                            End If
                        Else
                            'Failed to add. probably a bad ProgID.
                            Failmessage = "Failed to load the settingsPage progID," & StrctlID & vbCrLf & _
                                        "Error:" & vbCrLf & _
                                        Err.Description
                        End If      'add the settingspage to the controls collection.
                    Else
                        'failed to get the progID.
                        Failmessage = "This plugin does not have any settings."
                    End If  'get SettingsPage ProgID.
                    Else
                        'failed to get pluginUIdata
                        Failmessage = "This plugin does not have any settings."
                    End If    'Ask for IPluginUIData
    
                End If 'cast to IEvalEvents.
    
            End If  'Plugin object create failed.
    If Err <> 0 Or Failmessage <> vbNullString Then
        If Failmessage = vbNullString Then Failmessage = "This Plugin does not have any User-modifiable attributes."
        With Controls.Add("VB.Label", "Lbl" & Trim$(CurrPic), PicPlugins(CurrPic))
            .Move 5, 5
            .Caption = Failmessage
            .AutoSize = True
            .Visible = True
            .BackStyle = vbTransparent
            PicPlugins(CurrPic).Visible = True
            'add the tab, since it will not have been done before.
            'since we can't get a proper name for it, we'll need to use the progID on the tab.
            Set addedtab = TabPlugins.Tabs.Add(ProgIds(LoopID), , ProgIds(LoopID))
            addedtab.CanClose = True
            Set addedtab.Panel = PicPlugins(CurrPic)
            'Hide the tab itself.
            addedtab.Enabled = True 'so they can see the message.
            PicPlugins(CurrPic).Tag = "-1"
        End With
        Err.Clear
    End If
    CurrPic = CurrPic + 1
    Failmessage = ""
    
    Next LoopID
    TabPlugins.ShowTabs = True
    
End Sub

Private Sub PicPlugins_Resize(Index As Integer)
    'the tag is the index into the msettingsctls collection that
    'is our child.
    Debug.Print "PicPlugins.Resize"
    Dim getchild As VBControlExtender
    With PicPlugins(Index)
    If Val(.Tag) > 0 Then
        Set getchild = mSettingsCtls.Item(Val(.Tag))
        getchild.Move 0, 0, .ScaleWidth, .ScaleHeight
    
    
    End If
    End With
End Sub

Private Sub UserControl_Resize()
    Dim I As Long
    TabPlugins.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    For I = PicPlugins.LBound To PicPlugins.UBound
        PicPlugins_Resize (I)
    Next I
End Sub
