VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Represents a Function.
'I wasn't sure wether to add this to the parser, but I figured it would be easier.
'There will be a "FunctionManager" class that implements IEvalEvents and calls these functions by name.

'Basically, a function, as far as this class is concerned- can only be implemented via a single expression.
'by definition, it will need to use variables in the expression. A function takes parameters-
'the parameters are defined within an array, P.

'my original implementation allowed for the provision to give the names to the parser itself.


'TODO:\\ implement multi-line expression support, in essence, define a simple script language.
'Woops- I just realized, I kind of have that, with the ";" separator.
'or, create a new class that "derives" from this one that implements it, I don't know...
Public Event FunctionCall(Params, ByRef returnValue As Variant, ByRef Handled As Boolean)

Private VNames() As String


Private mvarnames As String
Private mvarName As String
Private mFuncParser As CParser


'change to Object type for release build...
'Private mIsAPICall As Boolean
Private mLoader As BCDynaCall.Loader
Private mLib As BCDynaCall.BCDynaLib
Private mFunc As BCDynaCall.dynFunc
Public Property Get Parser() As CParser
    Set Parser = mFuncParser
End Property
Public Sub InitAPIFunction(ByVal ModuleName As String, ByVal ProcedureName As String, ByVal RealName As String, ByVal ParamString As String, ByVal ReturnType As String)
On Error Resume Next
Set mLoader = createObject("BCDynaCall.Loader")

Set mLib = mLoader.LoadLibrary(ModuleName)

Call mLib.DefineFunction(ProcedureName, RealName, ParamString, ReturnType)
Set mFunc = mLib.Func(ProcedureName)
If Not mFunc Is Nothing Then
    mvarName = ProcedureName
End If
' Case "i4" 'long
'      Ret = vbLong
'    Case "i2" 'integer
'      Ret = vbInteger
'    Case "u1" 'byte
'      Ret = vbByte
'    Case "str" 'string
'      Ret = vbString
'    Case "var" 'variant
'      Ret = vbVariant
'    Case "cy" 'currency
'      Ret = vbCurrency
'    Case "r8" 'double
'      Ret = vbDouble
'    Case "r4" 'single
'      Ret = vbSingle
'    Case "dt" 'date
'      Ret = vbDate
'    Case "obj" 'object
'      Ret = vbObject


End Sub


Public Property Let Name(ByVal Vdata As String)
    If mvarName = "TIT" Or mvarName = "TOT" Then Stop
    mvarName = Vdata
End Property
Public Property Get Name() As String
    Name = mvarName
End Property

Private Sub Class_Initialize()
     Dim I As Long, makeit As cStringBuilder
     Set makeit = New cStringBuilder
    Set mFuncParser = New CParser
    mFuncParser.Create
    'Defaults to "X,Y,Z", this is,
    'like a few other things in this library, merely because I am making a graphing program first.
    'there. 26 parameters better be enough, right? :)
    'EDIT: parameters are now in the form P,P2,P3,P4,P5,P6,P7,P8,P9, etc...
    'we also add an array, P. generally using the array within expressions is a good idea.
    'this makes it possible to create expressions that operate on the entire array-
    'SEQ(X,STORE(Y,P[X]+Y),Lbound(P),Ubound(P))
    'Me.VarNames = "X,Y,Z,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W"
    For I = 1 To ParserSettings.NumCustomFunctionParams
        makeit.Append "P"
        makeit.Append Trim$(I)
        If I < ParserSettings.NumCustomFunctionParams Then
            makeit.Append ","
        End If
    Next I
End Sub
Public Sub InitParser(withparser As CParser)
'I must alert all readers of this code to a interesting
'"feature" that is pretty neat-
'the Clone() method of the Parser copies a reference to it's Functions collection,
'so the parser we use will have access to the EXACT same functions as the parent, at ALL
'times!



    Set mFuncParser = withparser.Clone
End Sub
Public Property Let Expression(ByVal Vdata As String)
    'Expression string. The Expression string should contain
    mFuncParser.Expression = Vdata
End Property
Public Property Get Expression() As String
    Expression = mFuncParser.Expression
End Property
'Public Property Let VarNames(ByVal VarNames As String)
'    Dim I As Long
'    mvarnames = VarNames
'    'Split it at the comma's.
'    VNames = Split(mvarnames, ",")
'     For I = 0 To UBound(VNames)
'        Call mFuncParser.Variables.Add(VNames(I), 0)
'    Next I
'End Property
'Public Property Get VarNames() As String
'    VarNames = mvarnames
'End Property

Public Sub CallFuncByRef(ByRef ReturnParam, Parameters As Variant)
    'calls this function, returning the result in returnparam
        
    '  --1/15/2008 @ 17:41--
    
    Dim DLLParams() As Variant
    Dim I As Long, Handled As Boolean
    Dim paramcount As Long
    Dim Params() As Variant
    Dim retval As Variant
    On Error Resume Next
    paramcount = UBound(Parameters)
    Params = Parameters 'copy into array to prevent errors later.
    If Err <> 0 Or paramcount < 0 Then paramcount = 0

    
   
   
    

 On Error Resume Next
    If Not mFunc Is Nothing Then
        'we are an API call...
        
        'and look at that, we already HAVE a params array... talk about a cake-walk...
        ReturnParam = mFunc.DoCallRef(Params)
        Exit Sub
    
    End If
    
    
    
    'The idea is simple- simply assign each parameter value to the corresponding variable in
    'our parsing Class, and execute our expression! Quite simple. Like I said.
    
    'also, for better compat, create an array  as well, P[].
    'step one- set the variables in mFuncparser as appropriate.
    
    RaiseEvent FunctionCall(Parameters, ReturnParam, Handled)
    If Handled Then Exit Sub
    On Error GoTo CONTINUEIT
    'prevent Subscript errors within reason.
    
    If UBound(Parameters) < 255 Then    '255 is a ok amount, but it will waste about 4K per function instance-
                                        'on the other hand, it is only during this call.
        ReDim Preserve Parameters(LBound(Parameters) To 255)
    End If
    mFuncParser.Variables.Add "P", Parameters
    'EDIT: added code to set all un-passed parameters to 0.
    

    On Error Resume Next
CONTINUEIT:
Err.Clear
    'mFuncParser.Expression = mFuncParser.Expression
    mFuncParser.ExecuteByRef ReturnParam
End Sub
Public Function CallFuncArray(Optional Parameters As Variant)
    Dim retval As Variant
    If IsMissing(Parameters) Then
        ReDim Parameters(0)
        Erase Parameters
    End If
        
        
    CallFuncByRef retval, Parameters
    If Not IsObject(retval) Then CallFuncArray = retval Else Set CallFuncArray = retval
End Function

Public Function CallFunc(ParamArray Parameters() As Variant) As Variant
    Dim retval As Variant
    Dim copyparam() As Variant
    'Calls the function with the given parameters.
    
    copyparam = Parameters
    CallFuncByRef retval, copyparam()
    If IsObject(retval) Then
        Set CallFunc = retval
    Else
        CallFunc = retval
    End If
End Function
