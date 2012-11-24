VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBPErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IBPErrorHandler. used to delegate errors back to the top-level CParser,
'which can then take an action as appropriate based on it's state
'(IE, Error modes).



Public Enum ErrorResolutionConstants
    Err_Resume      'execute a "Resume" statement.
    Err_ResumeNext  'execute a "Resume Next" Statement.
    Err_Ignore      'Ignore the error, and do whatever it takes to not cause that error. for example,
                    'the factorial Function will simply not take the factorial.
    Err_Fatal       'The error was fatal. Not used. Why is it here.
End Enum
