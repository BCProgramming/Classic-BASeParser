VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISettingsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ISettingsPage interface
'Implemented by ActiveX Controls that are used to
'modify the configuration of a Plugin.
'I kind of stole this type of architecture (ActiveX controls used for configuration settings)
'from my old BASeEdit XP, God rest it's soul...

'Oh- it is fully possible to make an implementor of this interface deal with the settings from multiple
'Plugins. Although I wouldn't delve to deeply into such endeavours. At least not personally.

'the thing is, the host form in BASeEdit will have multiple tabs that go to the same page(control).

'I'd need to change this interface to get something like that to work....
Public Enum EQueryInformationConstants
    Query_Null = 0  'pointless.
    Query_UniqueName
    
    'UniqueName:
    'this is extremely important:
    'the value returned with this query is passed to the persist/desist methods.
    
    'as such, both the IsettingsPage implementor as well as the plugin need to be aware of
    'this value.
    
    Query_DisplayName
    Query_DisplayIcon
    Query_Hidden    'is this particular settings page to be displayed as a tab?
                    '(it doesn't prevent a load of the control, but since the tab will not be added,
                    'then the usercontrol will never be brought forward.
                    '(useful if you only need one page to configure multiple plugins)
End Enum


Public Function QueryInformation(ByVal QueryType As EQueryInformationConstants) As Variant
    'query information.
    
    
    
    
End Function


Public Sub PersistToKey(ByVal RegKeySave As String)
    'The control must persist it's state to the registry, using RegKeySave as the base.
    'This will generally be in the form:
    'Software\BASeCamp\BASeParser\Configuration Data\<Set Name>\<ProgID of plugin>
    'so that different configuration sets can have different data stored for the same plugin.
    'I neglect to mention that unless you have an easy way to query the registry, this can
    'be a pain in the ass.
    'this is one reason I make Cregistry Public From the BASeParser Project itself.
    'unfortunately, that makes it a pain in the ass for those that have cRegistry already.
    
End Sub

Public Sub DesistFromKey(ByVal RegKeyLoad As String)
    'Same as above, but it is loading it back FROM the registry.
    'also, Same as above, it is a pain in the ass if you don't have an easy way to query the registry.
    
End Sub
