VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'represents a Formula Item
'The Parser Rips an expression and forms a linked list of Formula Items, known
'affectionately as the expression stack, formula stack, or whatever.
'Each Formula Item corresponds to a token within the original expression.
'Interestingly enough, IT_FUNCTION,IT_OBJACCESS contain "Sub"-parsers, essentially
'making the Linked list into a er- strange data structure. Because our "Value" property in those instances
'refers to a Array of CParser objects, corresponding to the argument list to the function (for IT_FUNCTION) or method (for IT_OBJACCESS)
'even MORE interesting, is how these sub-expressions themselves are optimized. For example, if "5^X/Y+N" was run through the parser, and then used (in the same form) as an argument
'to a function or Method call, the Parser would discover the already cached result of parsing the expression and use the
'cached Linked List. Novel, isn't it?

'I am REALLY proud of this project so far. On my machine, a 350Mhz AMD K6-2, it often performs a full parse and execute of some long expressions in milliseconds. The reason
'is because it was taking several seconds for the computer to do so earlier in development, and so I spent a good amount of time adding a plethora of
'parsing and execution optimizations that I probably would not have added if my computer was, say, a 1.8Ghz Pentium 4.
'Slower computers are better anyway. Especially as a development platform. I cannot count the number of
'Programs I have downloaded and tried to run on this computer something that could have easily run on my computer if
'the original developer had ever heard of optimization. But, guess what! It was developed on a computer with the latest and greatest CPU
'and video card. So what's the point? well, it runs like crap on my machine.

'I would have stuck with my ATI 3d Rage pro if it gave me acceptable gaming, but using the software renderer with my CPU was
'not exactly a speedy option.


Public Enum ItemTypeConstants
    IT_PLUGIN = -1     'handled by a ICorePlugin Implementor.
        'actually, This IT_PLUGIN is just a comparator- any value less than 0 will be assumed to
        'be a plugin-handled formula item.
        'each Core Plugin should use a unique ID number and assign it to the cFormItem objects
        'that it can handle. If two ID's conflict, then the first one encountered during the loop is used.
        'I know it isn't very robust with conflicts- I plan to fix that later.
    IT_NULL = 0
    IT_ENDOFLIST
    IT_VALUE
    IT_OPERATOR
    IT_FUNCTION
    IT_OPENBRACKET
    IT_CLOSEBRACKET
    IT_VARIABLE
    IT_ARRAYACCESS
    IT_ObjAccess
   
    IT_OPUNARY
    'IT_ENVIRON
    IT_BACKTICKS        'BackTick ` ` enclosed stuff is parsed as a parser Metacommand. none yet.
    IT_STATEMENT
    
    IT_BLOCK        'a block of statements.
    'how is THIS parsed, you ask?
    'well, an IF THEN, for example, could be a block.
    'Our OP would probably be IF, and our Tag will be a CParser. Why only one? Well,
    'since the Parser recognizes the semicolon as a statement separator, we simply
    'grab the block and replace all carriage return/Line-feed characters with
    'a semicolon, and put that whole block into the parser.
    IT_FLOWCHANGE
    
    
    'look in labelled collection for the value, which should be a labelname:
    

    
        'FLOWCHANGE: New!
        'the value property will be a cformItem. all we know is that rather then using the
        'current items next_ property, we should instead continue execution at the cformitem instance
        'within the value. Of course, this will completely screw up code that relies on the collapse to be done in the proper order.
        'also, it's probably a good idea to prevent infinite loops...
        'NEW(again) implement a "condition" of sorts, so that the flow only changes when certain conditions (IE, expressions) are true.
        'this will enable the parsing of a "Do...Loop", for example:
        
        'Do While (X<5);F[x]:=x^2;Loop
        
        'the parsing code will notice the Do While loop starting, and populate a formitem as appropriate (placing it on a stack, affectionately known as
        'the "BlockStack", the populating will include storing the condition and such...
        
        'Any other loops found within this one will simply push themselves onto both the parsestack and the "BlockStack"
        
        'When we encounter the appropriate ending of the loop, we will be able to evaluate the condition, and if true, we will move on to the Flowchanged item rather then our current nextitem.
        
        
    IT_METACOMMAND  'issues a metacommand to the parser, such as changing precision or something.
    IT_SUBEXPRESSION = IT_OPENBRACKET
    
End Enum
Private mvarobjCoreOwner As Object
Private mvaritemtype As ItemTypeConstants
Private mvarValue As Variant, MvarOp As Variant
Private mvarTag As Variant      'Extra data. For operators, holds the previous item.
Private mvarExpressionPosition As Long
Private mvarnext As CFormItem
Private mvarPrev As CFormItem
Private mExtra As Variant
Public Function Clone() As CFormItem
    Dim mNewItem As CFormItem
    Set mNewItem = New CFormItem
    With mNewItem
        Set .CoreOwner = Me.CoreOwner
         .ExprPos = Me.ExprPos
         If IsObject(Me.Extra) Then Set .Extra = Me.Extra Else .Extra = Me.Extra
         .ItemType = Me.ItemType
         Set .Next_ = Me.Next_
         Set .Prev = Me.Prev
         If IsObject(Me.Tag) Then Set .Tag = Me.Tag Else .Tag = Me.Tag
         If IsObject(Me.Value) Then Set .Value = Me.Value Else .Value = Me.Value
    End With
    Set Clone = mNewItem


End Function
Public Property Let Extra(ByVal Vdata As Variant)
    mExtra = Vdata
End Property
Public Property Get Extra() As Variant
    Extra = mExtra
End Property
Public Sub InsertAfter(cFormItemAfter As CFormItem)
    'inserts a linked list into this one at this location.
    
    'A-B-C-D-E
    
    'A-B-E-F-G-C-D-E
    'we need to link our next item to the first item of the given item.
    'we need to link the last item of that list to our current next item.
    Dim CurrNext As CFormItem, currPrev As CFormItem
    Dim lastinserted As CFormItem
    If Me.ItemType = IT_ENDOFLIST Then
        Me.ItemType = IT_NULL
    End If
    Set CurrNext = Me.Next_
    Set currPrev = Me.Prev
    Set Me.Next_ = cFormItemAfter
    Set cFormItemAfter.Prev = Me
    Set lastinserted = Me.GetLast
    Set lastinserted.Next_ = CurrNext
    If Not CurrNext Is Nothing Then
        Set CurrNext.Prev = lastinserted
    End If
    'that SHOULD do it.
    

End Sub
Friend Function GetLast() As CFormItem
    Dim lookCurr As CFormItem
    Dim lookPrev As CFormItem
    Set lookCurr = Me
    Do Until lookCurr Is Nothing
        Set lookPrev = lookCurr
        Set lookCurr = lookCurr.Next_
    Loop
    Set GetLast = lookPrev


End Function
Friend Property Set Prev(Vdata As CFormItem)
    Set mvarPrev = Vdata
End Property
Public Property Get Prev() As CFormItem
    Set Prev = mvarPrev
End Property
Friend Property Let ExprPos(ByVal Vdata As Long)
    mvarExpressionPosition = Vdata
End Property
Public Property Get ExprPos() As Long
    ExprPos = mvarExpressionPosition
End Property
Public Property Let Tag(ByVal Vdata As Variant)
'If Vdata <> "SUFFIX" And mvarTag = "SUFFIX" Then Stop
    mvarTag = Vdata
End Property
Public Property Set Tag(Vdata As Variant)
    Set mvarTag = Vdata
End Property
Public Property Get Tag() As Variant
    If IsObject(mvarTag) Then
        Set Tag = mvarTag
    Else
        Tag = mvarTag
    End If
End Property
Public Function GetString(withparser As CParser, Optional ByVal IncludeNext As Boolean = True) As String

    Dim Ret As String, cast As CParser
    Ret = "[ItemType """ & ItemTypeToString(mvaritemtype) & ""","
    Ret = Ret & "OP= """ & MvarOp & ""","
    
    Ret = Ret & "Value=" & withparser.ResultToString(mvarValue)
    Ret = Ret & "]"
    
    
    
   ' Ret = "OP:" & MvarOp & " mvaritemtype=" & mvaritemtype & ";"
    If Not mvarnext Is Nothing Then
        Ret = Ret & mvarnext.GetString(withparser)
    End If
    GetString = Ret

End Function
'Private Function ItemTypeToString(ByVal IType As ItemTypeConstants) As String
'    Dim RetVal As String
'    'If IType = IT_ARRAYACCESS Then
'    '    RetVal = "IT_ARRAYACCESS"
'    'elseif itype=IT_BACKTICKS
'
'
'
'   ' End If
'End Function


Public Property Get ItemType() As ItemTypeConstants

    ItemType = mvaritemtype
End Property
Public Property Let ItemType(Vdata As ItemTypeConstants)
    
    mvaritemtype = Vdata
End Property
Public Property Set Next_(Vdata As CFormItem)
'automatically set the next items' previtem to This.
    Set mvarnext = Vdata
    If Not mvarnext Is Nothing Then
        Set mvarnext.Prev = Me
    End If
End Property
Public Property Get Next_() As CFormItem
    Set Next_ = mvarnext
End Property
Public Property Let Value(ByVal Vdata As Variant)
    mvarValue = Vdata
End Property
Public Property Get Value() As Variant
    If IsObject(mvarValue) Then
        Set Value = mvarValue
    Else
    
    Value = mvarValue
    End If
End Property
Public Property Set Value(Vdata As Variant)
On Error Resume Next
    Set mvarValue = Vdata
End Property
Public Property Let op(ByVal Vdata As Variant)

    MvarOp = Vdata
    'If MvarOp = "*" Then Stop
End Property
Public Property Get op() As Variant
    op = MvarOp
 
End Property
'Friend Function Serialize(Optional propBag As PropertyBag) As Byte()
''use a property Bag :).
'If propBag Is Nothing Then
'    Set propBag = New VBRUN.PropertyBag
'End If
'
'
'
'
'
'
'
'
'
'
'End Function
Private Sub Class_Initialize()

End Sub



Private Sub Class_InitProperties()
'
End Sub





Private Sub Class_Terminate()
On Error Resume Next
    Debug.Print "CFormItem Terminating."

Set mvarnext = Nothing
Set mvarPrev = Nothing
    Set mvarobjCoreOwner = Nothing
End Sub


Public Sub Destroy()
    'Destroys the ENTIRE thing. This is because
    'I implemented it as a linked list,
    'which means that the terminate events won't fire anymore.
    If Not mvarnext Is Nothing Then
        Set mvarnext.Prev = Nothing
        'Assign the Previous item of our next item to nothing to prevent
        'a mutual recursion scenario.
        mvarnext.Destroy
        Set mvarnext = Nothing
    End If
    If Not mvarPrev Is Nothing Then
        Set mvarPrev.Next_ = Nothing
        mvarPrev.Destroy
        Set mvarPrev = Nothing
    End If
End Sub

Public Property Get CoreOwner() As Object

    Set CoreOwner = mvarobjCoreOwner

End Property
Friend Function CountAfter() As Long
    Dim inc As Integer
    If Not Me.ItemType = IT_NULL Or Me.ItemType = IT_ENDOFLIST Then
        inc = 1
    Else
        inc = 0
    End If
    If mvarnext Is Nothing Then
        CountAfter = inc
    Else
        CountAfter = inc + mvarnext.CountAfter
    End If
    



End Function
Public Property Set CoreOwner(objCoreOwner As Object)

    Set mvarobjCoreOwner = objCoreOwner

End Property
'Private Sub Class_ReadProperties(PropBag As PropertyBag)
''
'End Sub
'Private Sub Class_WriteProperties(PropBag As PropertyBag)
''
'
''me.ExprPos
''me.Extra
''me.ItemType
''Me.op
''me.Tag
''me.value
'Dim CurrSave As Long
'With PropBag
'    .WriteProperty "ItemType", mvaritemtype
'    .WriteProperty "Op", MvarOp
'    .WriteProperty "Tag", mvarTag
'
'
'    If IsArray(mvarValue) Then
'        WriteProperty "ValueCount", UBound(mvarValue) + 1
'        For CurrSave = 0 To UBound(mvarValue)
'            .WriteProperty "Value" & Trim$(CurrSave), mvarValue
'        Next
'    Else
'        .WriteProperty "ValueCount", 0
'        .WriteProperty "Value", mvarValue
'    End If
'
'
'    .WriteProperty "_Next", mvarnext
'End With
'
'
'
'End Sub



