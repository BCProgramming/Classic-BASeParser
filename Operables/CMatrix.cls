VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'The CMatrix Class
'This class doesn't represent a matrix as a matrix, but as an array of an array of Variants.
'this is because it is designed to be used as part of my Expression Evaluation library, which
'can only access multi-dimensioned array in that type of "simulated" format.
'the advantages are more- we can resize both dimensions as we please.

Implements IOperable

'UPDATE:
'So far, this class works for plain ol' numbers. But, I want it to be really cool, and work with anything the
'parser can handle- for example, Complex numbers.
'changes that must occur:
'a lot- all of the different functions that would normally have plain old arithmetic need to instead perform that
'arithmetic through the parser's PerformOperation() Function.


Private mRows As Long, MCols As Long
Private mMatrix As Variant
Private mParser As CParser   'assigned after creation
Private withparser As CParser
Public Sub SetElement(ByVal Col As Long, ByVal Row As Long, Vdata As Variant)
    Call Assign(mMatrix(Col)(Row), Vdata)
End Sub
Public Property Set Parser(Vdata As CParser)
    Set mParser = Vdata
    Set withparser = mParser.Clone
End Property
Public Property Get Parser() As CParser
    Set Parser = mParser
End Property
Friend Sub SetDirect(Vset As Variant)
    mMatrix = Vset
    mRows = UBound(mMatrix(1))
    MCols = UBound(mMatrix)
End Sub
Public Property Get Rows() As Long
   
    Rows = mRows
    
End Property
Public Property Get Cols() As Long

    Cols = MCols

End Property

Public Sub Resize(ByVal Cols As Long, ByVal Rows As Long, Optional ByVal FillValue As Variant = vbEmpty, Optional ByVal PreserveData As Boolean = False)
'X dim is first array- Y dim is the size of each array that is in X.
    Dim I As Long
    Dim LoopFill As Long
    Dim DimMe As Variant
    If PreserveData Then
        ReDim Preserve mMatrix(1 To Cols)
    Else
        ReDim mMatrix(1 To Cols)
    End If
    For I = 1 To Cols
        'DimMe = mMatrix(Cols)
        If PreserveData Then
            ReDim Preserve DimMe(Rows)
        Else
            
            ReDim DimMe(1 To Rows)
            If FillValue <> vbEmpty Then
                For LoopFill = 1 To Rows
                    Call Assign(DimMe(LoopFill), FillValue)
                Next LoopFill
            End If
        End If
        mMatrix(I) = DimMe
    Next I
    'remember variable shadowing :)
    mRows = Rows: MCols = Cols
    'tada. :)



End Sub
Public Function Identity(ByVal size As Long) As CMatrix
    Dim Ret As CMatrix
    Dim I As Long
    Set Ret = New CMatrix
    Set Ret.Parser = mParser
    Ret.Resize size, size, 0
    For I = 1 To size
        Ret.Item(I, I) = 1
    Next I
    Set Identity = Ret
End Function
Public Function Transpose() As CMatrix
    'Transpose this matrix into a new matrix.
    'the transpose of a matrix swaps the rows for columns- in a way, it has been rotated so
    'that what was row 1, Column 3 would become row 3, column 1.
    Dim RetMat As CMatrix
    Dim CCol As Long, CRow As Long
    Set RetMat = New CMatrix
    Set RetMat.Parser = mParser
    'resize to fit the elements.
    RetMat.Resize Me.Rows, Me.Cols
    For CCol = 1 To Me.Cols
        For CRow = 1 To Me.Rows
            
            'Call Assign(RetMat.Item(CRow, CCol), Me.Item(CCol, CRow))
            RetMat.SetElement CRow, CCol, Me.Item(CCol, CRow)
        
        Next CRow
    Next CCol
    'Debug.Print "Transposed:"
    Me.ToString
    'Debug.Print "result was:"
    RetMat.ToString
    Set Transpose = RetMat
End Function
Public Function ScalarOp(Otherval As Variant, ByVal op As String) As CMatrix
'This Class is designed ONLY for BASeParser XP.
'Changed this procedure so that is also supports Complex numbers, and anything else
'that BASeParser can throw at it.
    Dim RetMat As CMatrix
    Dim retval As Variant
    Dim CCol As Long, CRow As Long
    Dim WithThis As Variant
    
    Dim AMAT As CVariable, BMAT As CVariable
    
    
    'Define two variables- AMAT and BMAT, which are
    'the individual sides of the operation.
    'Set AMAT = WithParser.Variables.Add("AMAT", 0)
    'Set BMAT = WithParser.Variables.Add("BMAT", 0)
    'WithParser.Expression = "AMAT " & op & " BMAT"
    Set RetMat = New CMatrix
    Set RetMat.Parser = mParser
    'resize to fit the elements.
    RetMat.Resize Me.Rows, Me.Cols, False
    
    For CCol = 1 To Me.Cols
        For CRow = 1 To Me.Rows
            If TypeOf Otherval Is CMatrix Then
                WithThis = Otherval.Item(CCol, CRow)
            Else
                WithThis = Otherval
            End If
            'AMAT.Value = Item(CRow, CCol)
            'BMAT.Value = WithThis
            'WithParser.ExecuteByRef RetMat.Item(CRow, CCol)
            'CHANGE: uses the parser's PerformOperation() routine.
            mParser.PerformOperation op, Item(CRow, CCol), WithThis, retval
            RetMat.SetElement CCol, CRow, retval
'            Select Case Op
'            Case "+"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) + OtherVal
'            Case "-"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) - OtherVal
'            Case "*"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) * OtherVal
'            Case "\"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) \ OtherVal
'            Case "/"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) / OtherVal
'            Case "=", "=="
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) = OtherVal
'            Case "<>", "!="
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) <> OtherVal
'            Case "<"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) < OtherVal
'            Case "<=", "=<"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) <= OtherVal
'            Case ">=", "=>"
'                RetMat.Item(CRow, CCol) = Me.Item(CCol, CRow) >= OtherVal
'            End Select
        Next CRow
    Next CCol




    Set ScalarOp = RetMat


End Function
Public Property Get Element(ByVal Col As Long, ByVal Row As Long) As Variant
    Element = Item(Col, Row)
End Property
Public Property Let Element(ByVal Col As Long, ByVal Row As Long, ByVal Vdata As Variant)
    Item(Col, Row) = Vdata
End Property

Public Property Let Item(ByVal Col As Long, ByVal Row As Long, ByVal Vdata As Variant)
    mMatrix(Col)(Row) = Vdata
End Property
Public Property Get Item(ByVal Col As Long, ByVal Row As Long) As Variant
    Item = mMatrix(Col)(Row)
End Property
Public Function ToString() As String
    Dim LoopCol As Long, LoopRow As Long
    Dim strbuffer As String, GotItem
    strbuffer = "{" & vbCrLf
    For LoopRow = 1 To Rows
      strbuffer = strbuffer & "{"
        For LoopCol = 1 To Cols
      
            GotItem = Me.Item(LoopCol, LoopRow)
            If IsEmpty(GotItem) Then
                GotItem = "0"
            End If
            strbuffer = strbuffer & GotItem
            If LoopCol < Cols Then
                strbuffer = strbuffer & ","
            End If
        
        
    
        
        Next LoopCol
            strbuffer = strbuffer & "}"
            strbuffer = strbuffer & vbCrLf
    Next LoopRow
    strbuffer = strbuffer & "}"
    ToString = strbuffer
End Function


Public Function Multiply(OtherMatrix As CMatrix) As CMatrix
    'Perform a Matrix Multiplication
    Dim Ret As CMatrix
    Dim K As Long, I As Long, j As Long
    Dim Tot As Variant
    Dim resHolder As Variant
    Set Ret = New CMatrix
    Ret.Resize Me.Cols, OtherMatrix.Rows
    With Me
        For K = 1 To OtherMatrix.Cols
            For I = 1 To .Rows
                Tot = 0     'reset total count.
                For j = 1 To .Cols
                '     ' Multiply the corresponding elements and add them together
                'Tot = Tot + .Item(I, j) * OtherMatrix.Item(j, K)
                Call mParser.PerformOperation("*", .Item(I, j), OtherMatrix.Item(j, K), resHolder)
                'Call PerformOperation(Tot, resHolder, "+", Tot, mParser)
                mParser.PerformOperation "+", Tot, resHolder, Tot
                'there we are.
                
                
                
                
                Next j
                'store element total.
                'Call Assign(ret.Item(I, K), Tot)
                Ret.SetElement I, K, Tot
            Next I
        Next K
        
                
    End With
    Set Multiply = Ret
End Function
' Temp = CreateMatrix(dMat1.M, dMat2.n)
'm=rows,n=cols
' With dMat1
'  For K = 1 To dMat2.n  ' For each row in Matrix B
'   For I = 1 To .M      ' For each row in Matrix A
'    Tot = 0             ' Reset total
'    For j = 1 To .n     ' For each column in matrix A
'     ' Multiply the corresponding elements and add them together
'     Tot = Tot + .Element(I, j) * dMat2.Element(j, K)
'    Next j
'    ' Store the total
'    Temp.Element(I, K) = Tot
'   Next I
'  Next K
' End With
'
' MultiplyMatrix = Temp
Public Function Clone() As CMatrix
    Dim Ret As CMatrix
    Set Ret = New CMatrix
    Ret.SetDirect mMatrix
    Set Clone = Ret
End Function '
Public Sub SetRandom()
    Dim X As Long, Y As Long
    For X = 1 To Cols
    For Y = 1 To Rows
        Me.Element(X, Y) = Rnd * 100 + 50
    Next Y
    Next X
End Sub
Public Function Determinant() As Variant
   'm=rows,n=cols
   'you know whats annoying?  People who decide that more then one character variable names are some sort of
   'blaspheme, and use names like these.
   'I also dislike people that think their error messages are some kind of god's gift, and insist
   'on SHOUTING them, or worse, using !. I never got error 5, "INVALID PROCEDURE CALL OR ARGUMENT!" before.
Dim I As Integer, j As Integer
Dim m As Integer, n As Integer

Dim A1 As Variant, B1 As Variant
Dim A2 As Variant, B2 As Variant

Dim Mult As Variant
Dim Det As Variant
Dim Temp As CMatrix

Dim X As Integer, Y As Integer

 With Me
  I = .Rows
  j = .Cols
 
 'Check input
  Select Case True
   Case I <> j
    Err.Raise 5, "CMatrix - Determinant", "Cannot calculate Determinant of non-square matrix."
   Case I = 0
    Err.Raise 5, "CMatrix - Determinant", "Matrix too small to calculate Determinant."
  End Select
 
  Select Case j
   Case 1
   'A matrix of only 1 element...
   'determinant is only element.
   Call Assign(Det, .Element(1, 1))
   Case 2
   '2x2 matrices are very common... hardcoding the math is an easy way to boost performance!
   Dim tempresult, tempresult2
   mParser.PerformOperation "*", .Element(1, 1), .Element(2, 2), tempresult
   mParser.PerformOperation "*", .Element(1, 2), .Element(2, 1), tempresult2
    'Det = (.Element(1, 1) * .Element(2, 2)) - (.Element(1, 2) * .Element(2, 1))
    mParser.PerformOperation "-", tempresult, tempresult2, Det
    
   Case Else
   '3x3 or larger...
    Set Temp = Me.Clone
    Det = 1
   
   'Do Gaussian Elimination on matrix
    For I = 1 To j
     For m = I + 1 To j
      If Temp.Element(I, I) = 0 Then
       Determinant = 0 'Shortcut!!
       Exit Function
      Else
      'Modify the matrix if Temp.Element(i,i) would be 0
      'after the substitution (fixes bug)
      Dim restemp As Variant, restemp2, restemp3
      
     
     'ExecExpr "(ELEMA-(ELEMB/ELEMC)*ELEMD)", mParser, resTemp, "ELEMA", Temp.Element(m, m), _
     "ELEMB", Temp.Element(m, I), "ELEMC", Temp.Element(I, I), "ELEMD", Temp.Element(I, m)
     
     
      'need to use the performOperation sub to get the right values here.
      Call mParser.PerformOperation("/", Temp.Element(m, I), Temp.Element(I, I), restemp)
      'now we need:
      ' Temp.Element(m,m) - restemp * temp.Element(I,m) = 0
      'Times is higher precedence
      mParser.PerformOperation "*", restemp, Temp.Element(I, m), restemp
      'now, just temp.element(m,m) - restemp
      mParser.PerformOperation "-", Temp.Element(m, m), restemp, restemp
      'last but not least- the actual comparision. we do this, again, because complex
      'numbers aren't exactly what we want to fiddle with here.
      mParser.PerformOperation "==", restemp, 0, restemp
      'OK.
      
      
      
       'If Temp.Element(m, m) - Temp.Element(m, I) / Temp.Element(I, I) * Temp.Element(I, m) = 0 Then
       
      If CBool(restemp) Then
       'Find a row where Temp.Element(y,i) is not zero and add it to row n
        For Y = 1 To j
         If Not Temp.Element(Y, I) = 0 Then
          For X = 1 To j
          
          
           'Call ExecExpr("ELEMA+ELEMB", mParser, resTemp, "ELEMA", Temp.Element(m, x), "ELEMB", Temp.Element(m + 1, x))
           
           'Temp.Element(m, x) = resTemp
           'UGHH! darn it, I hate having to convert all these arithmetic operations.
           'and This is only the beginning- I still have inverse to do.
           Temp.Element(m, X) = Temp.Element(m, X) + Temp.Element(m + 1, X)
           mParser.PerformOperation "+", Temp.Element(m, X), Temp.Element(m + 1, X), restemp
           Temp.SetElement m, X, restemp
           'there.
          Next X
         End If
        Next Y
       End If
      
       'Mult = Temp.Element(m, I) / Temp.Element(I, I)
       'another arithmetic conversion.
       Call mParser.PerformOperation("/", Temp.Element(m, I), Temp.Element(I, I), Mult)
       
       
       'Call ExecExpr("ELEMA*ELEMB", mParser, Mult, "ELEMA", Temp.Element(m, I), "ELEMB", Temp.Element(I, I))
       
      End If
      
      For n = I To j   'Adjust column of elements
      ' Temp.Element(m, n) = Temp.Element(m, n) - (Mult * Temp.Element(I, n))
        'Another modification.
        'First, Mult*...
        'then tem..ment(m,n) - that result:
        Call mParser.PerformOperation("*", Mult, Temp.Element(I, n), restemp)
        'there we go.
        'now the subtraction...
        Call mParser.PerformOperation("-", Temp.Element(m, n), restemp, restemp)
        'whew. Now we assign it.
        Temp.Element(m, n) = restemp
        'whew.
        
      Next n
     Next m
    
    'Determinant is the product of all the major diagonal
    'elements! Keep a running total...
     'Det = Det * Temp.Element(I, I)
     mParser.PerformOperation "*", Det, Temp.Element(I, I), Det
     
     
     
     
    Next I
  End Select
 End With
If IsObject(Det) Then
    
    Set Determinant = Det
Else
    Determinant = Det
End If
End Function

Public Function Invert() As CMatrix
Dim I As Long, j As Long
Dim X As Long, Y As Long
Dim size As Long, Sine As Variant
Dim Temp As CMatrix, sMat As CMatrix
Dim Det As Variant
Dim restemp As Variant
Dim dMat As CMatrix

With Me

 I = .Cols
 j = .Rows
 size = j
 
 Select Case True
  Case I <> j
   Err.Raise 5, "CMatrix - InvertMatrix", "Cannot invert a non-square matrix!"
  Case I = 1
    Call mParser.PerformOperation("/", 1, .Element(1, 1), restemp)
    
    
    
   Set Invert = CreateMatrix(1, 1, restemp)
   'Exit Function
 End Select
 
 Call Assign(Det, Determinant)

 Select Case True
  Case Det = 0
   Err.Raise 5, "MatrixLib - InvertMatrix", "Cannot invert matrix with Determinant 0."
   'again, removed for brevity.
'  Case I = 2
'   ' Calculate the inverse of a 2x2 matrix directly
'   Set Temp = CreateMatrix(2, 2, .Element(2, 2), -.Element(1, 2), -.Element(2, 1), .Element(1, 1))
'   Set Invert = Temp.ScalarOp(1 / Det, "*")
'
'
'   Exit Function
 End Select

 Set Temp = CreateMatrix(size, size)
 
 For I = 1 To size
  For j = 1 To size
  
'     SubMat = CreateMatrix(size - 1, size - 1)
   
'   'Copy out the SubMatrix from from the original matrix
'   oldrow = 0
'   For x = 1 To size - 1
'    If x = i Then oldrow = oldrow + 2 Else oldrow = oldrow + 1
'    oldcol = 0
'    For y = 1 To size - 1
'     If y = j Then oldcol = oldcol + 2 Else oldcol = oldcol + 1
'     SubMat.Element(x, y) = .Element(oldrow, oldcol)
'    Next y
'   Next x

   Set sMat = SubMat(I, j)
   
   ' Temp is our C matrix
   Sine = (-1) ^ (I + j)
   'Temp.Element(I, j) = sMat.Determinant * Sine
   mParser.PerformOperation "*", sMat.Determinant, Sine, restemp
   Call Temp.SetElement(I, j, restemp)

  Next j
 Next I


 End With
 
 Set Temp = Temp.Transpose() 'Transpose C matrix
'put 1 / Det into restemp, here.
Call mParser.PerformOperation("/", 1, Det, restemp)
 Set Invert = Temp.ScalarOp(restemp, "*")  'ScaleMatrix(Temp, 1 / Det) 'Divide C matrix by Det
 
End Function
Public Function CreateMatrix(ByVal Cols As Long, ByVal Rows As Long, ParamArray Elements() As Variant) As CMatrix
    Dim Ret As CMatrix
    Dim CurrElem As Long
    Dim CRow As Long, CCol As Long
    Set Ret = New CMatrix
    Call Ret.Resize(Cols, Rows)
    On Error GoTo AssignError:
    
        For CRow = 1 To Rows
        For CCol = 1 To Cols
            Ret.Element(CCol, CRow) = Elements(CurrElem)
            CurrElem = CurrElem + 1
        
    Next CCol
Next CRow


  
AssignError:
  Set CreateMatrix = Ret
End Function

Public Function SubMat(ByVal eRow As Integer, ByVal eCol As Integer) As CMatrix
' Returns the submatrix to an element. The submatrix of
' an element is found by crossing out the row and column
' the element is in and copying the remaining elements into
' a new matrix;
'
' Example: Find the submatrix of element f:
'
'      *
'  [ a b c d ]       [ a * c d ]     [ a c d ]
' *[ e f g h ]*  --> [ * * * * ] --> [ i k l ]
'  [ i j k l ]       [ i * k l ]     [ m o p ]
'  [ m n o p ]       [ m * o p ]
'      *
'
Dim Temp As CMatrix
Dim oldrow As Integer, oldcol As Integer
Dim X As Integer, Y As Integer
Dim m As Integer, n As Integer


 With Me
  m = .Rows
  n = .Cols
  Set Temp = CreateMatrix(m - 1, n - 1)
   
 'Copy out the SubMatrix from from the original matrix
  oldrow = 0
  For X = 1 To m - 1
   If X = eRow Then oldrow = oldrow + 2 Else oldrow = oldrow + 1
   oldcol = 0

   For Y = 1 To n - 1
    If Y = eCol Then oldcol = oldcol + 2 Else oldcol = oldcol + 1
    Call Temp.SetElement(X, Y, .Element(oldcol, oldrow))
   Next Y
  Next X
 
 End With
 
 Set SubMat = Temp
End Function

Public Sub SetColumn(ByVal Colset As Long, Vdata As Variant)
    Dim makeCopy As Variant, loopit As Long
    If LBound(Vdata) <> 1 Then
        ReDim makeCopy(1 To (UBound(Vdata) - LBound(Vdata)) + 1)
        For loopit = LBound(Vdata) To UBound(Vdata)
            makeCopy((loopit - LBound(Vdata)) + 1) = Vdata(loopit)
        Next
    Else
        
        makeCopy = Vdata
    End If
    
    


    mMatrix(Colset) = makeCopy
End Sub
Public Sub SetRow(ByVal RowSet As Long, Vdata As Variant)
        Dim makeCopy As Variant, loopit As Long
    If LBound(Vdata) <> 1 Then
        ReDim makeCopy(1 To (UBound(Vdata) - LBound(Vdata)) + 1)
        For loopit = LBound(Vdata) To UBound(Vdata)
            Call Assign(makeCopy((loopit - LBound(Vdata)) + 1), Vdata(loopit))
        Next
    Else
        
        makeCopy = Vdata
    End If
    
    'with makecopy, we must iterate.
    For loopit = 1 To Me.Cols
        Call Me.SetElement(loopit, RowSet, makeCopy(loopit))
    Next loopit
    
    
End Sub
Private Sub Class_Initialize()
    'Set mParser = New CParser
    'mParser.Create
End Sub

Private Function IOperable_PerformFunc(withparser As CParser, ByVal FuncName As String, AllArguments As Variant, retval As Variant) As PerformFunctionReturnConstants
'
End Function

Private Sub IOperable_PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, retval As Variant)
Err.Raise E_NOTIMPL, "CMatrix::IOperable.PerformOp", "CMatrix Class does not yet implement operators"
End Sub

Private Function IOperable_toString(withparser As CParser) As String
    Dim CurrRow As Long, strbuffer As String
    Dim CurrCol As Long, castMat As CMatrix
    'Even though the Cmatrix Class has a ToString() method, we'll need to
    'create our own, since it will raise an error if there is an
    'object (IE, CComplex) within the matrix.
    Set castMat = Me
    strbuffer = ""
    For CurrRow = 1 To castMat.Rows
        strbuffer = strbuffer & "{"
        For CurrCol = 1 To castMat.Cols
                'recursively call result to string, just like for arrays.
                strbuffer = strbuffer & withparser.ResultToString(castMat.Item(CurrCol, CurrRow))
                If CurrCol < castMat.Cols Then strbuffer = strbuffer & ","
            
        Next CurrCol
        strbuffer = strbuffer & "}" & vbCrLf
    Next CurrRow
    IOperable_toString = strbuffer
End Function
