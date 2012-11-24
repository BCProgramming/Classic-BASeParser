VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParserGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


'Globals.
Public Property Get ParserSettings() As ParserSettings
    Set ParserSettings = ModParser.ParserSettings
End Property

Public Sub EvaluateExpressionByRef(ByVal ExprEval As String, ByRef retval As Variant, Optional ByVal configset As String = "Default", Optional ByRef ParseObj As CParser)
    'This is for people that just want the parsing, and don't
    'want to have to create the class and such.
    'Of course, it WILL go slower, but I will try my best to avoid making
    'it go too slow.
    'Also, it will be much less extendable- IE: not able to retrieve variables (you can use
    'STORE() and reference the variables, though.
    'well, you can retrieve that type of stuff through code- you'll need to pass in a uninitialized CParser Object
    
    'Static ParseObj As CParser
    If Not ReturnParser Is Nothing Then
    'egads! It is almost as if they wan't to change it...
    'oh, if they reference the same object, there is no point in changing ours.
        If ReturnParser Is ParseObj Then
        
        Else
            'Oh, they are different. Huh.
            
                Set ParseObj = Nothing
            
                Set ParseObj = ReturnParser
                         
        End If
        
    End If
    If ParseObj Is Nothing Then
        'first time through.
        Set ParseObj = New CParser
        ParseObj.Create configset
    
    End If
    'assign the expression.
    ParseObj.Expression = ExprEval
    ParseObj.ExecuteByRef retval
    Set ReturnParser = ParseObj


End Sub
Public Function EvaluateExpression(ByVal ExprEval As String, Optional ByVal configset As String = "Default")



    Dim Ret
    EvaluateExpressionByRef ExprEval, Ret, configset
    If IsObject(Ret) Then
        Set EvaluateExpression = Ret
    Else
        EvaluateExpression = Ret
    End If
    



End Function
