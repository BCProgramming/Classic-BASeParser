VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CVariable Class- represents a variable (duh)


'All CVariable Object must be "posessed" by a
'CVariables Collection. Hey, I don't make the rules.

'What is the point of this class? well, first off, some uses for a variable involve
'incrementing a variable. If the CVariables Collection was continuously queried, we'd be wasting time.
'Instead, it is a good idea to retrieve the CVariable objects you will manipulate BEFORE the loop,
'and then change the "Value" property of those Variables as you see fit.

'Also, this class has been extended to support indexing into an array. This will allow for the direct storage of elements
'into an array like "Array[i]:=6".

'the trick is that any instance used in that manner is not actually a "member" of the CVariables collection. Instead,
'it is told that it is accessing a existing array variable in it's "mParent" (via IsArrVar)
'then, the variable name will be in mvarname, and the index in mvartag.


'NEW: password protected variable access
'basically this can restrict assignment to a particular
'variable by imposing a hash based restriction- that is, the

'client that wishes to exhibit exclusive control of any and all changes to a variable
'can create and save a random password hash and set it on the variable.
Public Enum EVarType
  Variable_Variant = VbVarType.vbVariant
  variable_array = VbVarType.vbArray
  Variable_Number = VbVarType.vbDecimal
  Variable_Object = VbVarType.vbObject
  Variable_String = VbVarType.vbString
End Enum
  

Public Event Change(ByVal OldValue, ByRef NewValue)      'raised when value changes.
Public Event NameChange(ByVal VOldName As String, ByRef newname As String)  'raised when name changes. (see a pattern here?)
Private mParent As CVariables
Private mSubScript As Variant
Private mvarName As String
Private mvarValue As Variant
Private mvarTag As Variant
Private misArrVar As Boolean
Private mvarConstant As Boolean     'are we constant (can we be changed?)
Private mDataType As EVarType
Private mBoundTo As CVariable
Attribute mBoundTo.VB_VarHelpID = -1

Private mvarReadHash    As String
Private mvarWriteHash As String
Public Enum VariableProtectType
    Protect_Read = 0
    Protect_Write = 1


End Enum

'NEW FEATURE for cvariable class:
'referenced binding.
'if mboundTo is assigned a variable, changes in that variable will be reflected in the value
'of this variable.
'Ideally, of course, whatever variable is being used to reference US should simply refer to that other variable, but this could come in handy at
'some point I'm sure.
'Public Sub HashProtect(ProtectIt As VariableProtectType, ByVal Passcode As String, Optional ByVal CurrPass As String = "")
'    'hash Protects the given method of access.
'    'if a hash has previously been set requires the original passcode as well.
'    Dim WorkHash As Long
'    If ProtectIt = Protect_Write Then WorkHash = mvarWriteHash Else WorkHash = mvarReadHash
'    If WorkHash = "" Then
'        WorkHash = GetHash(Passcode)
'    Else
'        'make sure the given current passcode is correct...
'        If GetHash(CurrPass) <> WorkHash Then
'
'            'HA! I think NOT!
'            'raise permission denied error.
'            Err.Raise 70, "Cvariable Access protection", "Given original Passcode does not generate correct Hash String."
'        Else
'            'ok, good.
'            WorkHash = GetHash(Passcode)
'            'basically if an empty string is passed it will remove the protection.
'        End If
'    End If
'
'    If ProtectIt = Protect_Write Then mvarWriteHash = WorkHash Else mvarReadHash = WorkHash
'
'End Sub
Public Property Let DataType(ByVal vdata As EVarType)
 mDataType = vdata

End Property
Public Property Get DataType() As EVarType
 DataType = mDataType
End Property
Public Sub BindValue(VarBind As CVariable)
'
    If Not mBoundTo Is Nothing Then
        '- perhaps notification to a CVariable when another CVariable Binds/Unbinds to/from it?
        
    
    End If
    Set mBoundTo = VarBind
End Sub
Public Sub UnBindValue()
    '
        Set mBoundTo = Nothing
End Sub
Friend Sub SetArrayAccess(ByVal ArrayName As String, ByVal SubScript)
    'sets this to an array-access type variable.
    misArrVar = True
    mSubScript = SubScript
    mvarName = ArrayName
End Sub



Public Property Let Constant(ByVal vdata As Boolean)
    mvarConstant = vdata
End Property
'Tag- shouldn't be used in CParser, anyway.
'I intend on using this for my upcoming Scripting Language,
'BASeScript, to hold the scripting specific Variable info, such as scope.
Public Property Get Constant() As Boolean
    Constant = mvarConstant
End Property
Public Property Let Tag(ByVal vdata As Variant)
    mvarTag = vdata
End Property
Public Property Get Tag() As Variant
    If IsObject(mvarTag) Then Set Tag = mvarTag Else Tag = mvarTag
End Property
Public Property Set Tag(ByVal vdata As Variant)
    Set mvarTag = vdata
End Property
Friend Property Set Parent(vdata As CVariables)
    Set mParent = vdata
End Property
Public Property Get Parent() As CVariables
Attribute Parent.VB_Description = "Returns the CVariables Collection that contains this Variable, if any."
    Set Parent = mParent
End Property

Public Property Let Value(ByVal vdata As Variant)
Attribute Value.VB_Description = "sets/Returns the value of this variable."
Attribute Value.VB_UserMemId = 0
    Dim arrvar As CVariable
    If mvarConstant Then
        'GASP! I think not!
        Err.Raise 5, "CVariable::Value", "Constants Cannot be re-assigned."
    
    End If
    RaiseChange mvarValue, vdata
        'access it :)
        If Not misArrVar Then
            If IsObject(vdata) Then
                Set mvarValue = vdata
            Else
                mvarValue = vdata
            End If
        Else
            'array access-type variable.
            'we only accept numeric indices right now.
            'get the actual array variable:
            If Not IsArray(mSubScript) Then
                mParent.Item(mvarName).SetElement mSubScript, vdata
            Else
                mParent.Item(mvarName).SetElementArray mSubScript, vdata
            End If
'            If Not IsArray(arrvar.Value) Then
'                Err.Raise 5, "CVariable::Value", "Failed to access " & mvarName & " as an array."
'            Else
''                If IsObject(vdata) Then
''                    arrvar(mSubScript) = vdata
''                Else
''                    Set arrvar(mSubScript) = vdata
''                End If
'
'            End If
        End If
End Property
Public Property Set Value(ByVal vdata As Variant)
    RaiseChange mvarValue, vdata
    'Set mvarValue = Vdata
    Let Me.Value = vdata
End Property
Public Sub SetElementArray(subscripts As Variant, ByVal Value As Long)
    SetElementEx mvarValue, subscripts, Value
End Sub
Friend Sub SetElementEx(ByRef ArraySet As Variant, subscripts As Variant, ByVal Value As Long)
Debug.Assert False

Dim GotArrays() As Variant
Dim I As Long


ReDim GotArrays(UBound(subscripts))
GotArrays(0) = ArraySet(subscripts(0))
For I = 1 To UBound(subscripts) - 1
        GotArrays(I) = GotArrays(I - 1)(subscripts(I))

Next I

'now, assign...
GotArrays(I)(subscripts(UBound(subscripts))) = Value

For I = UBound(subscripts) - 1 To 1 Step -1
    GotArrays(I - 1)(subscripts(I)) = GotArrays(I)
Next I
GotArrays(0)(subscripts(0)) = GotArrays(1)





End Sub
Public Function SetElement(ByVal subscripts As Long, ByVal Value) As Variant
    Dim GotArray
    Dim I As Long
    
    
    
    
    
    If IsObject(Value) Then
        Set mvarValue(subscripts) = Value
        SetElement = Value
    Else
        mvarValue(subscripts) = Value
        SetElement = Value
    End If
End Function
Public Property Get Value() As Variant
    Dim varaccess As Variant
    
    If Not misArrVar Then
        If IsObject(mvarValue) Then
            Set Value = mvarValue
        Else
            Value = mvarValue
        End If
    Else
        'we are being used to access an array element.
        '
        'Debug.Assert False
        varaccess = mParent.Item(mvarName).Value
        If Not IsArray(varaccess) Then
            Err.Raise 5, "CVariable::Value", "failed to access """ & mvarName & """ as an array."
        Else
        
        Value = GetArrayElementVar(varaccess, mSubScript)
            If IsObject(varaccess(mSubScript)) Then
                Set Value = varaccess(mSubScript)
            Else
                Value = varaccess(mSubScript)
            End If
            
        End If
    
    End If
End Property
Friend Sub Setname(ByVal vdata As String)
    mvarName = vdata
End Sub


'MODIFICATION: Feb 19 2007:
' if the "Value" property is a CFunction, we will return the name of the CFunction, not the name of the variable.
'trust me...

Public Property Let Name(ByVal vdata As String)
    'notify our parent.
    'Vdata = UCase$(Vdata)
'    If TypeOf mvarValue Is CFunction Then
'        'Special case- if our value is a CFunction object, we do not want
'        'to return or set the name of the variable itself, but rather of the underlying function.
'        With CFunction(mvarValue)
'            .Name = Vdata
'            'tada.
'
'        End With
'
'
'
'    'we still do the default stuff after this if.
'    End If
    
    
    RaiseEvent NameChange(mvarName, vdata)
    mvarName = vdata
End Property
Public Property Get Name() As String
    'retrieve. Again, if our value is
    'a CFunction, retrieve the name of the CFunction instead.
    'I explain the meaning of this goofiness in the devlog...
    If TypeOf mvarValue Is CFunction Then
        With CFunction(mvarValue)
            Name = .Name
        
        'all done.
        'the with block is in case I add some other crap here.
        
        End With
    
    
    
    End If
    Name = mvarName
End Property
Private Sub RaiseChange(ByVal OldValue, ByRef NewValue)
    RaiseEvent Change(OldValue, NewValue)
End Sub

Private Sub Class_Initialize()
    mDataType = Variable_Variant
End Sub
