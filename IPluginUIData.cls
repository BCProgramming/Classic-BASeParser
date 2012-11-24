VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPluginUIData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IPluginUIData: Implemented by plugins who care about their appearance :*|

Public Enum PluginPictureSizeConstants
    PPS_SMALL   'system small size. usually 16x16 will do...
    PPS_LARGE   'shell size. usually 32x32 will do...
End Enum
Public Function GetPluginPicture(ByVal PicSize As PluginPictureSizeConstants) As IPictureDisp
    'Should return a IPictureDisp Object that can be used to draw an icon representing this Plugin.
    'It can also return Nothing, in which case a default plugin icon will be used.



End Function

Public Function GetSettingsPageProgID() As String
    'If applicable, should return the ProgID of an ActiveX control which implements
    'BASeParserXP.ISettingsPage. This control will be used when configuring a Plugin via the User
    'Interface.
    'Can return an empty string, signalling that there is no settings page.
    
    'although it is kind of irrelevant, the Core plugins all have their
    'own ActiveX controls- err, they WILL have them, anyway.
    '
End Function
Public Property Get Description() As String
    'returns a description for this. the core is "BASeParser Core Operators & Functions Plug-in"
    
End Property
