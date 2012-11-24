VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisplayAxis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'"DisplayAxis" class- holds information relating to a axis.
'Also handles drawing of the axes, once it nows wether it is the X or Y.

Private mLabelling As CCaption
Private mPen As CPen    'the pen we use.
Private mOriginPen As CPen  'The pen we use when we draw the origins. This is not part of the step.

        'There is no brush for an Axis (except within the pen...)
'For maximum flexibility:
'instead of requiring a straight-up step value, we use a BASeParserXP Expression. This is also so I can make sure that
'that library works properly with non-test data.





                                              
Public Enum DisplayAxisConstants
    Axis_Independent        'the X axis.
    Axis_Dependent           'the Y axis
End Enum
'I used variant for maximum scalability.
Private mDirty As Boolean   'Do we need to recreate the array?
Private mValueArray As Variant
Private mowner As CWindowDisplay
Private mvarStep As Variant
Attribute mvarStep.VB_VarHelpID = -1
Private mvarMin As Variant
Attribute mvarMin.VB_VarHelpID = -1
Private mvarMax As Variant
Attribute mvarMax.VB_VarHelpID = -1
Private mvarAxis As DisplayAxisConstants
Public Property Get Labelling() As CCaption
    Set Labelling = mLabelling
End Property
Public Property Get OriginPen() As CPen
    Set OriginPen = mOriginPen
End Property
Public Property Get Pen() As CPen
    Set Pen = mPen
End Property
Friend Property Set Owner(Vdata As CWindowDisplay)
    Set mowner = Vdata
End Property
Friend Property Get Owner() As CWindowDisplay
    Set Owner = mowner
End Property
Public Property Let Axe(ByVal Vdata As DisplayAxisConstants)
    mvarAxis = Vdata
End Property
Public Property Get Axe() As DisplayAxisConstants
    Axe = mvarAxis
End Property
Public Property Let StepVal(ByVal Vdata As Variant)
    If Vdata = 0 Or Sgn(mvarMax - mvarMin) <> Sgn(Vdata) Then
        Err.Raise 5, , """StepVal"" value cannot be zero and must cause min to go to max."
    End If
    mvarStep = Vdata
    
End Property
Public Property Get StepVal() As Variant
    StepVal = mvarStep
End Property
Public Property Let Min(ByVal Vdata As Variant)
    
    
    mvarMin = Vdata
    mDirty = True
End Property
Public Property Get Min() As Variant
    Min = mvarMin
    
End Property

Public Property Let Max(ByVal Vdata As Variant)
    mvarMax = Vdata
    mDirty = True
End Property
Public Property Get Max() As Variant
    Max = mvarMax
End Property
Private Function CreateSeqExpression()
CreateSeqExpression = UCase$("SEQ(X,X," & mvarMin.Name & "," & mvarMax.Name & "," & mvarStep.Name & ")")
End Function

Private Sub Class_Initialize()
'whatever the axis we will be, default to minimum of -10 to 10.
'later- add initialization via registry.
    
  
    'Set mLabelFont = New CFont
    Set mLabelling = New CCaption
    'init font.
    'mLabelFont.LogFont.Name = "Tahoma"
    'mLabelFont.LogFont.Size = 10
    Set mPen = New CPen
    mPen.width = 1
    mPen.Brush.Colour.Blue = 255
    Set mOriginPen = New CPen
    mOriginPen.Brush.Colour.LongVal = vbBlack
    
    mOriginPen.width = 2
    
   
    
  
    mvarMin = -10
    mvarMax = 10
    mvarStep = 1
    
   
    
    'there.
    mDirty = True
End Sub
Private Function RecreateArray() As Boolean
    'recreates the private Array mValueArray
    'HA! but it is a simple function call.
    On Error Resume Next
    Dim I As Long
    Dim runner As Variant
    mvarStep = Abs(mvarStep) * Sgn(mvarMax - mvarMin)
    ReDim mValueArray(0 To Abs((mvarMax - mvarMin) \ mvarStep))
    runner = mvarMin
    For I = 0 To Abs((mvarMax - mvarMin) \ mvarStep)
        mValueArray(I) = runner
        runner = runner + mvarStep
    
    
    
    
    
    Next I
    
    
    
    
    
    
'        mValueArray = m_Parser.Execute  'we don't want objects, so use the function version.
'        If Err <> 0 Or (UBound(mValueArray) <= 0) Then
'            'An error!
'            CDebug.Post "Not good- m_Parser returned an error."
'            Debug.Assert False
'            RecreateArray = False
'        End If
'        RecreateArray = True
'        mDirty = False
End Function

Private Sub mvarMax_Change(ByVal OldValue As Variant, NewValue As Variant)
    mDirty = True
End Sub

Private Sub mvarMin_Change(ByVal OldValue As Variant, NewValue As Variant)
    mDirty = True
End Sub

Private Sub mvarStep_Change(ByVal OldValue As Variant, NewValue As Variant)
    mDirty = True
End Sub


Public Sub DrawAxis(OnDC As Long)
    'Draws the Axis on OnDC.
    
    'First, if mDirty is true, recreate the array.
    '   if the recreateArray function returns false- assert.
    
    'if the Axe is X, then the return values are the Y values of horizontal lines we must draw.
    'if the Axe is Y, then the return values are the X values of the Vertical lines we must draw.
    'There will be other modes, for Radial plotting. as well as properties for
    'drawing Axe Labels.
    Dim ArrRet As Boolean
    Dim CurrIndex As Long
    Dim cVal As Variant
    Dim drawline(1 To 2) As PointF
    Dim ValConvert As PointF, Converted As PointF
    Dim OriginPoints(1 To 2)   As PointF
    Dim oldpoint As POINTAPI
    If mDirty Then
        ArrRet = RecreateArray()
        'Debug.Assert ArrRet
        
    End If
    'SPEED OPTIMIZATION:
    'select our pen.
    Debug.Print "Selecting Pen for DisplayAxis."
    mPen.PenStyle = PS_SOLID
    mPen.SelectPen OnDC
    For CurrIndex = LBound(mValueArray) To UBound(mValueArray)
        cVal = mValueArray(CurrIndex)
        'this is the Independent.
        'the dependent is the entire width, because we will draw a line.
        If mowner.WindowType = WindowTypeConstants.Windowtype_Cartesian Then
            'cartesian axes are a single line.
            'others, for example- Polar, will have circles and lines and other mixes,
            'so they will require different points.
            Set drawline(1) = PointF(cVal, Min)
            Set drawline(2) = PointF(cVal, Max)
            'CHECK: are we the X axis or Y axis.
            'if we are Y(dependent) axis, Invert the points
            If mvarAxis = Axis_Dependent Then
                Debug.Print "UnSwapped:" & drawline(1).ToString & " to " & drawline(2).ToString
                drawline(1).SwapXY
                drawline(2).SwapXY
                Debug.Print "Swapped:" & drawline(1).ToString & " to " & drawline(2).ToString
            End If
            
            
            
            'now convert these two Points to the appropriate coordinates.
            Set drawline(1) = mowner.ToPixelCoord(drawline(1))
            Set drawline(2) = mowner.ToPixelCoord(drawline(2))
            'since it is now a pixel coordinate, we shall:
            'draw a line.
            MoveToEx OnDC, drawline(1).X, drawline(1).Y, oldpoint
            LineTo OnDC, drawline(2).X, drawline(2).Y
            
            
            
        End If
    
    Next
    mPen.UnSelectPen OnDC
    
    'Now, we draw the origin lines, using, duh, the origin pen settings.
    Set OriginPoints(1) = PointF(0, Min)
    Set OriginPoints(2) = PointF(0, Max)
       If mvarAxis = Axis_Dependent Then
                OriginPoints(1).SwapXY
                OriginPoints(2).SwapXY
            
        End If
    
        'Okay, select the pen, draw the lines.
        Set OriginPoints(1) = mowner.ToPixelCoord(OriginPoints(1))
        Set OriginPoints(2) = mowner.ToPixelCoord(OriginPoints(2))
        mOriginPen.SelectPen OnDC
        
        MoveToEx OnDC, OriginPoints(1).X, OriginPoints(1).Y, oldpoint
        LineTo OnDC, OriginPoints(2).X, OriginPoints(2).Y
        'And Unselect the pen...
        
        mOriginPen.UnSelectPen OnDC
    
    
End Sub

