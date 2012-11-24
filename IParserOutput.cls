VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IParserOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IParserOutput: Interface for the display of Parser messages.
'This isn't required and is more useful for a command-line type application.


'Message: display, log... whatever the message.
'NOTE: ParserOutput must be enabled in the Parser objects Config for it to call this method.
Public Sub Message(Withparser As CParser, ByVal Message As String, Optional ByVal VerbosityLevel As Integer = 0)
    'Show the message. Or log it, or WHATEVER the implementor wants to do with it.
    
End Sub
