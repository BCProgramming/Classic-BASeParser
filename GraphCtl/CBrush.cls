VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CBrush- Wraps Windows Brush API calls into a nice clean object.
'Replacing my older CBrush which was lost in a hard Drive crash. sigh.

'Author: Michael Burgwin, BASeCamp Corporation.
'you may use this code as you see fit. I would appreciate any comments you can give me.
'I can sometimes be reached at BC_Programming@hotmail.com. (when I don't use it for a month, all my mail dissapears.
'Nice, Microsoft. (of course, it is understandable..)

Public Event Change()

Public Enum BrushStyleConstants
 BS_NULL = 1
 BS_HATCHED = 2
 BS_HOLLOW = BS_NULL
 BS_PATTERN = 3
 BS_PATTERN8x8 = 7
 BS_SOLID = 0
End Enum
Public Enum HatchStyleConstants
 HS_BDIAGONAL = 3
 HS_CROSS = 4
 HS_DIAGCROSS = 5
 HS_FDIAGONAL = 2
 HS_HORIZONTAL = 0
 HS_VERTICAL = 1
End Enum
Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type PTAPI
    x As Long
    y As Long
End Type
Public Type SELECTED_DCS
    mDC As Long
    mOld As Long   'old brush handle.
End Type

Private mDCSelected() As SELECTED_DCS
Private Const OBJ_BRUSH As Long = 2
Private mInitializing As Boolean
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Declare Function CreateHatchBrush Lib "gdi32.dll" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lppt As PTAPI) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (ByRef lpLogBrush As LOGBRUSH) As Long
Private mvarHatchStyle As HatchStyleConstants
Private mvarBrushStyle As BrushStyleConstants
Private mvarBrushPic As StdPicture

Private mvarhBrush As Long
Private WithEvents mvarColour As ClsColour
Attribute mvarColour.VB_VarHelpID = -1
Private mdirty As Boolean
Private Sub HasChanged()
    If mInitializing Then Exit Sub
    mdirty = True
    RaiseEvent Change
End Sub
Public Property Set BrushPic(Vdata As StdPicture)
    Set mvarBrushPic = Vdata
    HasChanged
End Property
Public Property Get BrushPic() As StdPicture
    Set BrushPic = mvarBrushPic
End Property
Public Property Let BrushStyle(ByVal Vdata As BrushStyleConstants)
    mvarBrushStyle = Vdata
    HasChanged
End Property
Public Property Get BrushStyle() As BrushStyleConstants
    mvarBrushStyle = mvarBrushStyle
End Property

Public Property Let HatchStyle(ByVal Vdata As HatchStyleConstants)
   mvarHatchStyle = Vdata
   HasChanged
    
End Property
Public Property Get HatchStyle() As HatchStyleConstants
    HatchStyle = mvarHatchStyle
End Property

'Specifies a hatch style. The meaning depends on the brush style defined by lbStyle.
'If lbStyle is BS_DIBPATTERN, the lbHatch member contains a handle to a packed DIB. To obtain this handle, an application calls the GlobalAlloc function with GMEM_MOVEABLE (or LocalAlloc with LMEM_MOVEABLE) to allocate a block of memory and then fills the memory with the packed DIB. A packed DIB consists of a BITMAPINFO structure immediately followed by the array of bytes that define the pixels of the bitmap.
'
'If lbStyle is BS_DIBPATTERNPT, the lbHatch member contains a pointer to a packed DIB. The pointer derives from the memory block created by LocalAlloc with LMEM_FIXED set or by GlobalAlloc with GMEM_FIXED set, or it is the pointer returned by a call like LocalLock (handle_to_the_dib). A packed DIB consists of a BITMAPINFO structure immediately followed by the array of bytes that define the pixels of the bitmap.
'
'If lbStyle is BS_HATCHED, the lbHatch member specifies the orientation of the lines used to create the hatch. It can be one of the following values. Value Meaning
'HS_BDIAGONAL A 45-degree upward, left-to-right hatch
'HS_CROSS Horizontal and vertical cross-hatch
'HS_DIAGCROSS 45-degree crosshatch
'HS_FDIAGONAL A 45-degree downward, left-to-right hatch
'HS_HORIZONTAL Horizontal hatch
'HS_VERTICAL Vertical hatch
'
'
'
'If lbStyle is BS_PATTERN, lbHatch is a handle to the bitmap that defines the pattern. The bitmap cannot be a DIB section bitmap, which is created by the CreateDIBSection function.

'If lbStyle is BS_SOLID or BS_HOLLOW, lbHatch is ignored.

'Remarks
'Private Sub ValidateDCs()
'
'    Dim CurrDC As Long, I As Long
'    Dim BrushPtr As LOGBRUSH
'    Dim BrushCheck As Long
'    Dim TempBrush As Long
'    Dim NewArray() As SELECTED_DCS
'    Dim DCCount As Long
'    'validates the DCs in the selectedDCs array. makes Sure they point to this
'    'Brush as the OBJ_BRUSH.
'    On Error Resume Next
'    TempBrush = CreateSolidBrush(0)
'    For I = 0 To UBound(mDCSelected)
'        If Err <> 0 Then Exit For
'        CurrDC = mDCSelected(I).mDC
'        'in order to retrieve the current brush in the DC, we need to selectObject
'        'a new one, then swap the return value back in.
'
'        BrushCheck = SelectObject(CurrDC, TempBrush)
'        'brush check is the current brush.
'        'now select it back into the DC.
'        TempBrush = SelectObject(CurrDC, BrushCheck)
'        If BrushCheck = mvarhBrush Then
'            'it DOES point to us.
'            ReDim Preserve NewArray(DCCount)
'            NewArray(DCCount) = mDCSelected(I)
'            DCCount = DCCount + 1
'        Else
'            'it doesn't.
'            'don't add it- we have nothing to
'            'do with it now.
'        End If
'    Next I
'
'End Sub

Private Function MakeBrush() As Long
'Re-Create the brush.
    'NOTE: the current hBrush will be deleted, so the other procs in this class
    'are programmed to ensure they unselected all DCs that use this
    'brush, then they will reset it to the new brush afterwards.
    Dim deleted As Long
    Dim brushlog As LOGBRUSH
    If mvarhBrush <> 0 Then 'if we have a brush.
        deleted = DeleteObject(mvarhBrush)
        If deleted = 0 Then
            'SHAME!
            Debug.Print "failed to delete old brush."
        End If
        'current brush deleted.
        mvarhBrush = 0
    End If
    
    brushlog = Me.GetLogBrush
    'use CreateBrushInDirect to create the new brush.
  
        mvarhBrush = CreateBrushIndirect(brushlog)
        'whew.
        If mvarhBrush = 0 Then
            'OH DEAR! this cannot be good.
            'add call to "RaiseAPIError" here, once I add it.
            Debug.Assert False
        End If
        'return this value.
        
    
    
        MakeBrush = mvarhBrush






End Function
Public Function GetLogBrush() As LOGBRUSH
    Dim brushlog As LOGBRUSH
  With brushlog
        .lbColor = mvarColour.LongVal
        .lbStyle = mvarBrushStyle
        Select Case mvarBrushStyle
            Case BrushStyleConstants.BS_HATCHED
                'wow, it is actually a hatched brush.
                .lbHatch = mvarHatchStyle
            Case BrushStyleConstants.BS_HOLLOW, BS_NULL, BS_SOLID
                'Doesn't matter- the API will ignore this
                'field anyway.
            Case BrushStyleConstants.BS_PATTERN, BS_PATTERN8x8
                'a pattern- requires a HBITMAP.
                'surprise- we'll find that in the Handle field of our StdPicture.
                If mvarBrushPic Is Nothing Then
                    'Oh, that is just great.
                    'hmmm.
                    'to avoid a GPF (I don't know what'll happen if I pass 0 as the hbmp)
                    'i'll force a solid brush.
                    .lbStyle = BS_SOLID
                    .lbColor = mvarColour.LongVal
                Else
                    'good, at least they gave us a pattern.
                    .lbHatch = mvarBrushPic.Handle
                End If
        
        
        End Select
        'OK, now we can create the brush. Hold your breath....
    End With
GetLogBrush = brushlog
End Function
Private Sub UnSelectDCs()
    'Unselect this brush from All the DCs.
    Dim I As Long
    Dim OldBrush As Long
    Debug.Print "CBrush: Unselecting DCs..."
    On Error Resume Next
    For I = 0 To UBound(mDCSelected)
        If Err <> 0 Then
        Debug.Print "No DCs to unselect."
        Exit For
        End If
        With mDCSelected(I)
        If .mDC <> 0 Then       'if we have a 0, then we JUST redimmed it in a different stack-frame.
            'retrieve the current brush of this DC.
            'if it is NOT (!) this brush, then we delete it. (this should foil VB when it
            'automatically sets the brush styles and such)
            Debug.Print "Unselecting DC,""" & .mDC & """ Giving oldbrush, """ & .mOld & """."
            OldBrush = SelectObject(.mDC, .mOld)
            If OldBrush <> mvarhBrush And mvarhBrush <> 0 Then
                'if we have a brush, and that is NOT it,
                'well, then.
                DeleteObject OldBrush
                'what are they doing adding their own brushes. WE OWN ALL BRUSHES! ROAR!
            
                
            
            End If
        End If
        End With
        'good, unselected in all counts. now all the DCs have their old Brush selected.
        
    Next I
    
End Sub
Private Sub ReSelectDCs()
    Dim I As Long
    Dim OldBrush As Long
    On Error Resume Next
    For I = 0 To UBound(mDCSelected)
        If Err <> 0 Then Exit For
        With mDCSelected(I)
            If .mDC <> 0 Then
            OldBrush = SelectObject(.mDC, mvarhBrush)
            
            If OldBrush <> mDCSelected(I).mOld Then
                'hmm.
                'Not sure which to do here-
                'delete OldBrush and forget about it,
                'or erase ours and replace it with the one we found. (?)
                DeleteObject OldBrush
                
            End If  'otherwise, the brush is proper- the old one is the one we have.
            End If
        End With
    Next I
End Sub
Public Property Get BrushHandle() As Long
    If mvarhBrush = 0 Or mdirty Then
        
        UnSelectDCs
            'the makebrush function will delete this
            'brush if necessary.
        mvarhBrush = MakeBrush
        ReSelectDCs
    End If
    BrushHandle = mvarhBrush


End Property
Public Sub SelectBrush(ByVal inDC As Long)
    'Selects the brush into the DC.
    'also, adds that DC and it's old brush into the private array.
    Dim OldBrush As Long
    Dim NewOldDC As SELECTED_DCS
    On Error Resume Next
        If (UBound(mDCSelected) + 1) Then
            If Err <> 0 Then
                'an error.
                ReDim mDCSelected(0)
            Else
                'no error.
                ReDim Preserve mDCSelected(UBound(mDCSelected) + 1)
                
            End If
        Else
            'huh?
            Debug.Assert False
        
        
        End If
        OldBrush = SelectObject(inDC, Me.BrushHandle)
        With mDCSelected(UBound(mDCSelected))
            .mDC = inDC
            .mOld = OldBrush
        End With
        
        
        
        
End Sub
Public Sub UnSelectBrush(ByVal FromDC As Long)
    'UnSelect the brush from the given DC.
    'NOTE: we ensure that this brush is in fact already selected, by looking through the SelectedDCs array.
    
    Dim I As Long, MoveIt As Long
    For I = 0 To UBound(mDCSelected)
        If Err <> 0 Then
            'no selected brushes, and thus we conclude that
            'the DC given could not possibly have us as a brush.
            Exit For
        Else
            With mDCSelected(I)
                If .mDC = FromDC Then
                    'found it!
                    SelectObject .mDC, .mOld
                    
                    'remove the item from the array.
                    For MoveIt = I To UBound(mDCSelected)
                        mDCSelected(I) = mDCSelected(I + 1)
                    Next MoveIt
                    ReDim Preserve mDCSelected(UBound(mDCSelected) - 1)
                    
                    
                    
                    
                    'whew.
                    
                    
                End If
            End With
        
        
        End If
        
    
    
    
    
    
End Sub


Public Property Get Colour() As ClsColour
    Set Colour = mvarColour
End Property





Private Sub Class_Initialize()
mInitializing = True
Set mvarColour = New ClsColour
mvarBrushStyle = BS_SOLID
mvarHatchStyle = HS_HORIZONTAL
mInitializing = False
End Sub

Private Sub Class_InitProperties()
    Class_Initialize
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mvarBrushStyle = .ReadProperty("Style", mvarBrushStyle)
        Set mvarColour = New ClsColour
        mvarColour.LongVal = .ReadProperty("Colour", vbBlue)
        Set mvarBrushPic = .ReadProperty("Picture", Nothing)
        mvarHatchStyle = .ReadProperty("HatchStyle", HS_CROSS)
    End With
End Sub

Private Sub Class_Terminate()
    Debug.Print "CBrush Terminated."
    UnSelectDCs
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    'write the brush properties to the property bag.
    With PropBag
        'good thing StdPicture is a persistable property, Right?
        .WriteProperty "Style", mvarBrushStyle, BS_SOLID
        .WriteProperty "Colour", mvarColour.LongVal, vbBlue
        .WriteProperty "Picture", mvarBrushPic, Nothing
        .WriteProperty "HatchStyle", mvarHatchStyle, HS_CROSS
        
    
    End With
End Sub

Private Sub mvarColour_Change()
    HasChanged
End Sub
