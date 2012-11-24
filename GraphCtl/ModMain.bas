Attribute VB_Name = "ModMain"
Option Explicit

Public CDebug As New CDebug
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, ByRef lppt As POINTAPI, ByVal cPoints As Long) As Long

Sub Main()
    Set CDebug = New CDebug
End Sub

Public Sub Assign(VarTo, VarAssign)
    If IsObject(VarAssign) Then
        Set VarTo = VarAssign
    Else
        VarTo = VarAssign
    End If
    
End Sub

Public Function PointF(ByVal X As Double, ByVal Y As Double) As PointF
    Dim makepoint As PointF
    Set makepoint = New PointF
    
    makepoint.X = X
    makepoint.Y = Y
    Set PointF = makepoint
End Function
Public Function createObject(ByVal ProgID As String) As Object
    'subclass createobject and explicitly create the intrinsic handlers if
    'their ProgID is returned.
    If StrComp(ProgID, "BASeParserXP.BPCoreOpFunc", vbTextCompare) = 0 Then
        Set createObject = New BASeParserXP.BPCoreOpFunc
    ElseIf StrComp(ProgID, "BASeParserXP.FunctionHandler", vbTextCompare) = 0 Then
        Set createObject = New BASeParserXP.FunctionHandler
    Else
        Set createObject = VBA.createObject(ProgID)
    End If
End Function

