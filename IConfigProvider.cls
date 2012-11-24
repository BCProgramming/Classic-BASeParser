VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConfigProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IConfigProvider:
'Implemented by classes that can provide a UI for the display of configuration information.
'mainly a public interface for the Configuration form in the BPXPGUI project, for the enabling of any
'special mode-specific controls.
Public Enum ConfigShowStates
    Config_Basic = -1  'should only show basic stuff
    Config_Intermediate = 0 'More powerful configuration options should be shown in addition to all Config_Basic items.
    Config_Advanced = 1 'I think you get the picture from the above comments.
    Config_Debug = 2 'this is actually a secret mode in a way. should only contain debugging helper stuff. Although
            'I suppose one could create secrets within it....
End Enum

Public Function QueryShowState() As ConfigShowStates
    'return the current show state. the default should be Config_Intermediate.
End Function
Public Property Get Setting(ByVal PluginName As String, ByVal SettingName As String) As Variant
'
End Property
