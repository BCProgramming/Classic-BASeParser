VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGraphFunction3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private flDirty As Boolean
Private mvarTriPoints() As DxVBLibA.D3DTLVERTEX
    'transformed and lit vertex for us.
Private flDirty As Boolean 'wether MvarTriPoints reflects the values for the current function.
Private mParent As GraphFunctions
Private mvarFunction As BASeParserXP.CFunction   'the CFunction- we delegate
Private mvarGranularity As Double
Private mvarMinX As Double, MvarMinY As Double, MvarMaxX As Double, MvarMaxY As Double
Private Sub Changed()
    flDirty = True
    'Erase mvarTriPoints
End Sub
'MinX,MaxX,MinY,MaxY...
Public Property Let Granularity(ByVal Vdata As Double)
    If Vdata <= 0 Then Err.Raise 13
    mvarGranularity = Vdata
End Property
Public Property Get Granularity() As Double
    Granularity = mvarGranularity
End Property
Public Property Let MinX(ByVal Vdata As Double)
    
    mvarMinX = Vdata
End Property
Public Property Get MinX() As Double
    MinX = mvarMinX
End Property

Public Property Let MaxX(ByVal Vdata As Double)
    
    MvarMaxX = Vdata
End Property
Public Property Get MaxX() As Double
    MaxX = MvarMaxX
End Property

Public Property Let MinY(ByVal Vdata As Double)
   
    MvarMinY = Vdata
End Property
Public Property Get MinY() As Double
    MinY = MvarMinY
End Property

Public Property Let MaxY(ByVal Vdata As Double)
    
    MvarMaxY = Vdata
End Property
Public Property Get MaxY() As Double
    MaxY = MvarMaxY
End Property

Friend Property Set Parent(Vdata As GraphFunctions3D)
    Set mParent = Vdata
    
End Property
Public Property Get Pen() As CPen
    Set Pen = MvarPen
End Property
Public Property Let Expression(ByVal Vdata As String)
    mvarFunction.Expression = Vdata
End Property
Public Property Get Expression() As String
    Expression = mvarFunction.Expression
End Property


Public Property Let Name(ByVal Vdata As String)
    mvarFunction.Name = Vdata
End Property
Public Property Get Name() As String
    Name = mvarFunction.Name
End Property
Friend Property Set CFunction(Vdata As CFunction)
    Set mvarFunction = Vdata
End Property
Public Property Get CFunction() As CFunction
    Set CFunction = mvarFunction
    
End Property
Private Sub Class_Initialize()
On Error Resume Next
    Set mvarFunction = New CFunction
    MinX = -10
    MinY = -10
    MaxX = 10
    MaxY = 10
    mvarGranularity = 1
    'two variables- X and Y, which generate a given Z-coordinate.
    
    
 
    
End Sub

Private Sub Class_InitProperties()
    Set mvarFunction = New CFunction
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Name = .ReadProperty("Name", "")
   
        Me.Expression = .ReadProperty("Expression", Me.Expression)
    End With

End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Name", Me.Name, ""

        .WriteProperty "Expression", Me.Expression

        
    End With
    
 
End Sub
Public Sub GetValueForInputsByRef(ByRef retval As Variant, VarInputs() As Variant)
    'uses the CallFuncByRef on the Cfunction to get the values.
    
    Call mvarFunction.CallFuncByRef(retval, VarInputs())
    'wow. that was easy.



End Sub
Private Sub CreateTriPoints()

    're-creates the points within mvartrippoints.
    Erase mvarTriPoints
    Dim CurrX As Double, CurrY As Double
    'now, how do we order them?
    
    'the first point will be used to create a TRIANGLE FAN with 2,7 and 6
    
    
    
    ' 1-2-3-4-5
    ' | | | | |
    ' 6-7-8-9-10
    ' | | | | |
    ' +-+-+-+-+
    ' | | | | |
    ' +-+-+-+-+
    ' | | | | |
    ' +-+-+-+-+
    
    'run down CurrX from start to finish....
    For CurrX = MinX To MaxX Step mvarGranularity
    
    
    Next CurrX
End Sub
Public Function DrawFunctionD3D(D3DDev As Direct3DDevice8)
    'tasks:
    '1. re-create the points of the function if necessary.
    


    
End Function
