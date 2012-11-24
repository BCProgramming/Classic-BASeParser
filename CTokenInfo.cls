VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTokenInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'COperatorInfo: used by implementors of IEvalEvents to
'return a value from GetOperatorInfo.


Private mOperator As String
Private mDescription As String
Private mIsUnary As Boolean

Public Property Let Operator(ByVal Vdata As String)
    mOperator = Vdata
End Property
Public Property Get Operator() As String
    Operator = mOperator
End Property

Public Property Let Description(ByVal Vdata As String)
    mDescription = Vdata
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
