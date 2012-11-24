VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Vector Object- used to encapsulate Vector expressions within the parser.
'technically, a vector can be in two or three dimensions (well, at least in this app)
'they are denoted with angle brackets, as in <1,2,3>*<3,4,2> which would
'be the scalar product of those two vectors.
'This class uses my other class, PointF.
'Of course, it keeps the z portion to itself, (if applicable) and uses it as necessary.
'Oh, and in the spirit of things,
Public Enum VectorCoordinateConstants
    Vector_X = 1
    Vector_Y = 2
    Vector_Z = 3


End Enum
Private mvarVecPart(1 To 3) As Variant      'As Decimal :)


Public Sub InitVector(ParamArray Entries() As Variant)
    Dim I As Long
    If UBound(Entries) < 4 Then
        For I = 1 To UBound(Entries) + 1
            mvarVecPart(I) = Entries(I + 1)
        Next I
    End If
            


End Sub
Public Property Get X() As Variant
    mvarX = mvarVecPart(Vector_X)
End Property

