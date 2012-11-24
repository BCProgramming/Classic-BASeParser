VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScripter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CScriptor
'Interprets "BASeScript control flow and decision structures and
'uses a CParser to perform evaluation.


Private mParser As CParser

Public Property Set Parser(Vdata As CParser)
    Set mParser = New CParser
End Property
Public Property Get Parser() As CParser
    Set Parser = mParser
End Property
