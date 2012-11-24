VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICorePlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ParseLocationReturnConstants
    ParseLocation_Default   'Parser should continue with it's standard code.
    ParseLocation_recognized        'returned when the Core Plugin has recognized one of it's
    'tokens.

End Enum
'ICorePlugin Interface.
'--5/26/2007 @ 04:31--
'this is just a test...
'--5/28/2007 @ 15:47--
'seems to work as designed. No guarantees, of course.



'Public Function IsMovable(WithParser As CParser, Formitem As CFormItem) As Boolean
''returns wether buildstack_infix should relocate this formitem as if it
''were a literal, function, or variable.
'End Function
Public Function Init(WithParser As CParser) As IPluginInitCode
'
End Function
Public Function Term(FromParser As CParser) As IPluginTermCode
'
End Function
Public Function ParseLocation(WithParser As CParser, ByVal StrParse As String, ByRef Position As Long, ByRef currItem As CFormItem) As ParseLocationReturnConstants
'

'parses the given location in the string for a token that the particular Core plugin
'recognizes. Note that this is only called if BASeParser's Internal recognition of Functions,Operators and literals all fail to resolve the location. (this means Evaluator Plugins get the first crack at function resolution)



currItem.ItemType = -1      '-1 not allowed on any plugin, because the
'interface class definition uses it.
ParseLocation = ParseLocation_Default
End Function
Public Function GetPluginUIData() As IPluginUIData
    'Set GetPluginUIData = Me
End Function
Public Function CanHandleItem(WithParser As CParser, Itemcheck As CFormItem) As Boolean
    '



    'should return wether this plugin can deal with the given formitem (IE, it is a custom formItem)
     CanHandleItem = False
End Function
Public Sub HandleItem(WithParser As CParser, ItemHandle As CFormItem, retval As Variant)
    'handle the given item, returning a value as appropriate.
    
    
    
    'huh?
End Sub

Public Property Let Settings(WithParser As CParser, ByVal PropName As String, ByVal Vdata As Variant)
'
End Property

Public Property Get Settings(WithParser As CParser, ByVal PropName As String) As Variant
'
End Property
