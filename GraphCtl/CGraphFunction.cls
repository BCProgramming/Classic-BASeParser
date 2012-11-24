VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGraphFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CGraphFunction- encapsulates the information that pertains to a Function
'in the GUI world, (Display components) but not to the Cparser's Functions
'collection. the "GraphFunctions" Collection implements the IEvalEvents
'interface to enable the functions to reference each other.
Private MvarPen As CPen     'draw with this pen.

Private mParent As GraphFunctions
Private mvarFunction As BASeParserXP.Cfunction   'the CFunction- we delegate
Private mvarGranularity As Double
                    'some methods/props.
Friend Property Set Parent(Vdata As GraphFunctions)
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
Friend Property Set Cfunction(Vdata As Cfunction)
    Set mvarFunction = Vdata
End Property
Public Property Get Cfunction() As Cfunction
    Set Cfunction = mvarFunction
    
End Property
Private Sub Class_Initialize()
On Error Resume Next
    Set mvarFunction = New Cfunction
    
    'only one variable= independent variable X.
    
    Set MvarPen = New CPen
    MvarPen.Brush.Colour = vbRed
    MvarPen.Width = 3
    mvarGranularity = 0.25
End Sub

Private Sub Class_InitProperties()
    Set mvarFunction = New Cfunction
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
        .WriteProperty "Pen", Me.Pen
        
    End With
    
 
End Sub
Public Sub GetValueForInputsByRef(ByRef retval As Variant, VarInputs() As Variant)
    'uses the CallFuncByRef on the Cfunction to get the values.
    
    Call mvarFunction.CallFuncByRef(retval, VarInputs())
    'wow. that was easy.



End Sub

Public Sub DrawFunction(ByVal OnDC As Long)
    Dim WndUse As CWindowDisplay
    Dim StartVal, EndVal
    Dim PrevPointsP() As PointF 'previous pixel coordinates. Number if elements depends on
                                'wether the returned value was an array.
    Dim PrevPointsG() As PointF 'same as above, but for Graph coordinates.
    Dim PrevCount As Long   'number of items in PrevPoints array.
    Dim temphold As Variant, valreturned As Variant
    Dim CurrVar As Long, LoopVal As Variant
    Dim Inputs() As Variant
    Dim GCoord As PointF
    Dim PCoord As PointF, Dummy As POINTAPI
    Set WndUse = mParent.Owner.Window
    'use that window.
    'what do we need it for? Well, we need to translate our Graph coordinates into pixel coordinates, for one.
    StartVal = WndUse.IndependentAxis.Min
    EndVal = WndUse.IndependentAxis.Max
    
    ReDim loopvalarr(0)
   ' prevreturned = Null
    'If Left$(Me.Expression, 1) = "{" Then Stop
    
    MvarPen.SelectPen OnDC
    
    'Algorithm:
    'initialize prevpoint arrays to a single element.
    'Iterate from the starting value to the ending value.
    
    ReDim PrevPointsP(0)
    ReDim PrevPointsG(0)
    Set PrevPointsP(0) = New PointF
    Set PrevPointsG(0) = New PointF
'    With PrevPointsG(0)
'        .X = WndUse.IndependentAxis.Min
'        .Y = 0
'    End With
    Set PrevPointsP(0) = Nothing
    Set PrevPointsG(0) = Nothing
    Dim varloop As CVariable
    Set varloop = mvarFunction.Parser.Variables.Add("X", 0)
    
    For LoopVal = StartVal To EndVal Step mvarGranularity
      'grab the current value:
      ReDim Inputs(0)
      Call Assign(Inputs(0), LoopVal)
      varloop.Value = LoopVal
      GetValueForInputsByRef valreturned, Inputs()
      'Debug.Print mvarFunction.Expression & " for X=" & Inputs(0) & " Is " & valreturned
      'OK.
      'now, if it isn't an array, force it to be one.
        If Not IsArray(valreturned) Then
            Assign temphold, valreturned
            ReDim valreturned(0)
            Assign valreturned(0), temphold
        End If
        'Now, iterate through each element of the array of returned values....
        For CurrVar = 0 To UBound(valreturned)
            'Steps:
            'Create a New Graph Coordinate.(X=LoopVal,Y=ValReturned(CurrVar))
            'map that graph coordinate to a pixel coordinate.
            Set GCoord = New PointF
            GCoord.X = LoopVal
            GCoord.Y = valreturned(CurrVar)
            Debug.Print "VALUE RETURNED:" & valreturned(CurrVar)
            Set PCoord = WndUse.ToPixelCoord(GCoord)
            
            'determine wether there is a corresponding element in the prevpoints() arrays.
            'do this by checking their UBound.
            If UBound(PrevPointsG) >= CurrVar Then
                'If so, draw from PrevPointsP(CurrVar) to the created Pixel Point.
                'move to the position.
                If Not PrevPointsP(CurrVar) Is Nothing Then
                    MoveToEx OnDC, PrevPointsP(CurrVar).X, PrevPointsP(CurrVar).Y, Dummy
                    'Draw the line:
                    LineTo OnDC, PCoord.X, PCoord.Y
                    'TADA!
                    CDebug.Post "Drew a Line from " & PrevPointsP(CurrVar).ToString & " To " & PCoord.ToString
                Else
                    ReDim Preserve PrevPointsP(CurrVar)
                    ReDim Preserve PrevPointsG(CurrVar)
                    Set PrevPointsP(CurrVar) = PCoord
                    Set PrevPointsG(CurrVar) = GCoord
                End If
            
            Else
            
                'If Not, do NOT draw anything- Instead, we simply resize the array, and make the calculated
                'coordinates the previous coordinates at that location.
                'redimension.
                ReDim Preserve PrevPointsP(CurrVar)
                ReDim Preserve PrevPointsG(CurrVar)
                Set PrevPointsP(CurrVar) = PCoord
                Set PrevPointsG(CurrVar) = GCoord
            End If
        
        
        
        
        Next CurrVar
        
        
        
            
        
        
        
    Next
    MvarPen.UnSelectPen OnDC


    



End Sub
