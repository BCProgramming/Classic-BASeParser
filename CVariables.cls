VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CVariables collection, represents the variables posessed in a parser.

'fired during assignment and retrieval of variables.
'for example, VarSet can be hooked to check for Constants and refuse to allow
'them to be changed.


Public Event GetOwner(ByRef ParserGet As CParser)
Public Event VarAdded(ByVal VarAdded As CVariable)
Attribute VarAdded.VB_Description = "Fired when a variable is added."
Public Event VarChanged(ByVal VarChanged As CVariable, ByVal OldValue As Variant)
Attribute VarChanged.VB_Description = "Fired when a variable is changed."
Public Event VarRetrieved(ByVal VarRetrieved As CVariable)
Attribute VarRetrieved.VB_Description = "Fired when a variable is retrieved. allows per-reference subclassing of that value."
Public Event VarReplace(Replacing As CVariable, ByRef Allow As Boolean)
Attribute VarReplace.VB_Description = "Fired when a variable is to be replaced."
Public Event VarRemove(RemoveMe As CVariable)
Private mCol As Collection
Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of variable in the collection."
    Count = mCol.Count
End Function
Public Sub Clear(Optional ByVal Keepconstants As Boolean = True)
    'Clears all variables.
    'obviously very dangerous.
    'especially when your dealing with constants....
    'thus the parameter for keeping them.
    Dim currItem As Long
    Dim castVar As CVariable
    If Not Keepconstants Then
        Set mCol = New Collection
    Else
        For currItem = mCol.Count - 1 To 1 Step -1
            Set castVar = mCol.Item(currItem)
            If Not castVar.Constant Then
                mCol.Remove currItem
            
            End If
            Set castVar = Nothing
        
        Next
    End If
End Sub
Public Function GetNames() As String()
Attribute GetNames.VB_Description = "Returns a space-delimited list of all the variable names contained in the collection."
    Dim Ret() As String
    Dim LoopVar As CVariable
    Dim I As Long
    If mCol.Count = 0 Then
        Exit Function
    End If
    ReDim Ret(0 To mCol.Count - 1)
    For Each LoopVar In mCol
        Ret(I) = LoopVar.Name
        I = I + 1
    Next
    GetNames = Ret
End Function
Private Function CreateVariable(ByVal Varname As String, ByVal VarValue As Variant) As CVariable
    Dim NewVar As CVariable, Allow As Boolean
    Set NewVar = New CVariable
    'if already existent, use event.
    Dim gotvar As CVariable
    On Error Resume Next
    If InStr(Varname, "[") > 0 Then
        Set gotvar = Me.Item(Varname)
        Varname = gotvar.Name
        If IsObject(VarValue) Then Set gotvar.Value = VarValue Else gotvar.Value = VarValue
        Set CreateVariable = gotvar
        Exit Function
    Else
        Set gotvar = mCol.Item(Varname)
    End If
    If Not gotvar Is Nothing Then
        'we get here when there is an item already.
        If Err = 0 Then
            'if no error, it exists.
            Err.Clear
            'OK- check with the interface.
            'but first, raise the event to see if it is allowed.
            Set NewVar = gotvar
                
                RaiseEvent VarReplace(gotvar, Allow)
                If Allow Then
                    Set NewVar = gotvar
                    mCol.Remove Me.GetItemIndex(NewVar)
                Else
                    Err.Raise Exec_VariableAlreadyExists
                End If
          
        End If
    End If
    With NewVar
        .Name = Varname
        
        .Value = VarValue
        Set .Parent = Me
    End With
    
    mCol.Add NewVar, Varname
    RaiseEvent VarAdded(NewVar)
    Set CreateVariable = NewVar

End Function
Public Sub AddMany(ParamArray AddNameValue() As Variant)
Attribute AddMany.VB_Description = "Adds a number of new variables, given a variable list of pairs in the form Name,Value,name,Value..."
Dim I As Long
For I = 0 To UBound(AddNameValue) Step 2
    Me.Add AddNameValue(I), AddNameValue(I + 1)
Next I




End Sub
Private Sub Class_Initialize()
    Set mCol = New Collection
    
End Sub
Public Function Create(ByVal Varname As String, ByRef VarValue As Variant) As CVariable
    Dim gotvar As CVariable
    Set gotvar = mCol.Item(Varname)
    
End Function


Public Function Add(ByVal Varname As String, ByRef VarValue As Variant) As BASeParserXP.CVariable
Attribute Add.VB_Description = "Adds a new variable, given its name and value. If the variable name given already exists, a Event is fired in an attempt to resolve it."
 '
 If Varname = "" Then
    'nice try, but empty-string variables aren't allowed.
    'right now I ignore it, but should I raise an error?
    CDebug.Post "Ignored Attempt to add un-named variable."
    'Stop
    Exit Function
ElseIf Val(Varname) <> 0 Or IsNumeric(Varname) Then
   ' CDebug.Post "hey, F-off, no Numbers can be variable names!"
   Err.Raise 5, "CVariables::Add", "Variable Names may not start with numeric characters."
    
End If
    Set Add = CreateVariable(Varname, VarValue)
End Function
Public Function Remove(ByVal VIndex As Variant) As BASeParserXP.CVariable
Attribute Remove.VB_Description = "Removes a Variable from the collection. the parameter can be a collection ordinal, the variable name, or the CVariable object you wish to remove."
Dim remindex As Variant, I As Long
    If TypeOf VIndex Is CVariable Then
        For I = 1 To mCol.Count - 1
            If mCol(I) Is VIndex Then
                remindex = I
           
                Exit For
            
            End If
        Next I
    ElseIf VarType(VIndex) = vbString Or IsNumeric(VIndex) Then
        remindex = VIndex
        
        
    End If
    RaiseEvent VarRemove(mCol.Item(remindex))
    Set Remove = mCol.Item(remindex)
    mCol.Remove remindex
End Function
Public Function Exists(ByVal VariableName As String) As Boolean
    'returns wether a given variable exists in the collection.

    On Error Resume Next
    Exists = TypeOf GetItem(VariableName) Is CVariable
    If Err <> 0 Then Exists = False



End Function

Private Function GetItem(ByVal vdata As Variant)

'Trick: detect when a string is passed wether an array is being indexed and if so return that array variable with an element set.
'Dim x As CVariable

Dim firstbracket As Long, secondbracket As Long, Varname As String
Dim usevariable As CVariable, Mparser As CParser
firstbracket = InStr(vdata, "[") + 1
secondbracket = InStr(vdata, "]")
 On Error Resume Next
 
    If InStr(vdata, "[") < InStr(vdata, "]") And (InStr(vdata, "[") <> 0) Then
        'bit of a kludge atm... oh well.
        'Dim ElementIndex As Long
        Dim indices As Variant
        Dim strsubscripts As String, parsers() As CParser
        Dim subscripts() As String
        Dim I As Long, mparse As CParser
        Varname = Mid$(vdata, 1, InStr(vdata, "[") - 1)
        'later: split elementindex into an array of subscripts.
        strsubscripts = Mid$(vdata, firstbracket, secondbracket - firstbracket)
        RaiseEvent GetOwner(mparse)
        
        subscripts = mparse.ParseArguments("[" & strsubscripts & "]", "[", "]")
        ReDim parsers(UBound(subscripts))
        ReDim indices(UBound(subscripts))
        For I = 0 To UBound(subscripts)
            Set parsers(I) = mparse.Clone()
            parsers(I).Expression = Mid$(subscripts(I), 2, Len(subscripts(I)) - 2)
            indices(I) = parsers(I).Execute
        Next I
        
        
        If UBound(indices) = 0 Then indices = indices(0)
        'ElementIndex = Val()
        Set usevariable = New CVariable
        Set usevariable.Parent = Me
        
        usevariable.SetArrayAccess Varname, indices
        
        
        
        Set GetItem = usevariable
        Exit Function
    End If
 
    Set GetItem = mCol.Item(vdata)
    If Err <> 0 Then
        On Error GoTo 0 'allow the collection to raise'em.
        'add a new empty variable.
        Set GetItem = Add(vdata, vbEmpty)
    
    End If
End Function
Public Function GetItemIndex(ItemFind As CVariable) As Long
    Dim I As Long
    For I = 1 To mCol.Count
        If mCol.Item(I) Is ItemFind Then
            GetItemIndex = I
            Exit Function
        End If
    
    Next I
    


End Function
Public Function GetExisting(ByVal vdata As Variant) As BASeParserXP.CVariable
    Set GetExisting = mCol.Item(vdata)
End Function

Public Property Get Item(ByVal vdata As Variant) As BASeParserXP.CVariable
Attribute Item.VB_Description = "Returns a CVariable Object for the given Variable Name or index. If the variable does not exist, it will be created."
Attribute Item.VB_UserMemId = 0
Set Item = GetItem(vdata)
End Property
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Function


