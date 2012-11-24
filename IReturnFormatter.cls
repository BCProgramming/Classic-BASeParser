VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReturnFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IReturnFormatter:
'the CParser object can have only a single IReturnFormatter function.
'the FormatResult() function will be called after a value is calculated. This is designed for
'such uses and rounding numeric values to a certain number of places. It is not recommended to
'do much of anything with object values that are passed in (such as replace the object return with a string
'representation), as that would make any Object access operators try to access a string value (understand?)



Public Sub FormatResult(ByRef Value As Variant)
    'format the result.
    'the CParser class itself will handle this.
    'it is recommended NOT to change the type of Value.
    'this may cause an expression to not work even though it looks fine.
    
End Sub
Public Function ResultToString(ByVal VarResult As Variant, ByRef Handled As Boolean) As String
    'formats a result to a string.
    'the recommended method is to implement
    'whatever special behaviours are desired, and assign True to handled . If a value
    'is not applicable for whatever special behaviours you wish to
    'implement, then simply assign False to Handled.
    
    'ResultToString of the Cparser implements behaviour relevant to
    'the properties assigned on the CParser object.
    'For example, formatting as hexadecimal, scientific, etc.
    
End Function
