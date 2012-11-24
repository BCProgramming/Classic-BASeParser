VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Represents an Array within the Parser.
'Why do I use this class instead of an array? Good question.
'Well, the parser initially used arrays- it worked, until I decided to add multi-dimensional array support. Which, of course,
'failed miserably, since I had no way to convert an array of indexes into a paramarray without a giant Select Case statement.
'ANYWAY-
'this class contains data. surprise. What we want
'it to be able to do is allow un-managed indexing, so we can get a matrix.
'this could be implemented via Sub-lists:

'[1] [2] [3]
'[4] [5] [6]
'[7] [8] [9]

'this matrix could be split into three separate DataList classes.
'but what about three dimensions? Well, a three dimensional object
'is simply a certain number of two dimensional objects, we'd just have a number of
'sub lists that each contain matrices, which contain the base-case lists.

'I'll see if this algorithm works. I THINK it is unique...

Private mvarSubLists() As CDataList
Private mvarData() As Variant


Public Sub Resize(ByVal KeepData As Boolean, ParamArray newdimensions() As Variant)
    'resize the datalist.
    'if we were given One parameter, we're OK, since we only need to redim the variables.
    If UBound(newdimensions) = 0 Then
        If KeepData Then
            ReDim Preserve mvarData(newdimensions(0))
        End If
    Else
        'darn, we don't get away with a single dimension.
   End If
End Sub

Public Property Get Value(ByVal Indices As Variant) As Variant


'
'returns the value at the appropriate index. If Indices is an array, the elements of the array will be used in drilling down
'through our sub-lists.
Dim MultiIndex As Boolean
If IsObject(Indices) Then
    If TypeOf Indices Is CDataList Then
    
    
    End If
End If
If IsArray(Indices) Then


End If



End Property
Public Property Let Value(ByVal Indices As Variant, ByVal vData As Variant)


'




End Property
