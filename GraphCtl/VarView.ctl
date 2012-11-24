VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl VarView 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ListView LvwVariables 
      Height          =   2220
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   3916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "VarView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mParser As BASeParserXP.CParser
Attribute mParser.VB_VarHelpID = -1
Private WithEvents mVariables As BASeParserXP.CVariables
Attribute mVariables.VB_VarHelpID = -1

Public Sub Refresh()
    Dim CvarLoop As CVariable
    LvwVariables.ListItems.Clear
    LvwVariables.ColumnHeaders.Clear
    With LvwVariables.ColumnHeaders
        .Add , , "Name"
        .Add , , "Value"
        
    
    End With
    
    
    
    If mVariables Is Nothing Then Exit Sub
    
    For Each CvarLoop In mVariables
        AddItem CvarLoop
    Next
End Sub
Private Sub AddItem(VarAdd As CVariable)
        Dim listadd As ListItem
        Set listadd = LvwVariables.ListItems.Add(, , VarAdd.Name)
        With listadd
            Set .Tag = VarAdd
            .SubItems(1) = mParser.ResultToString(VarAdd.Value)
        End With
        
End Sub

Public Property Set ParserWatch(Vdata As Object)
    Set mParser = Vdata
    Set mVariables = mParser.Variables
End Property
Public Property Get ParserWatch() As Object
    Set ParserWatch = mParser
End Property

Private Sub mVariables_VarAdded(ByVal VarAdded As BASeParserXP.CVariable)
    AddItem VarAdded
End Sub
Private Function ReverseLookupItem(Varlookup As CVariable)
    Dim LvwLoop As ListItem
    For Each LvwLoop In LvwVariables.ListItems
        If LvwLoop.Tag Is Varlookup Then
            Set ReverseLookupItem = LvwLoop
            Exit Function
        End If
    
    Next


End Function
Private Sub mVariables_VarChanged(ByVal VarChanged As BASeParserXP.CVariable, ByVal OldValue As Variant)
Dim finditem As ListItem
    If Not TypeOf VarChanged.Tag Is ListItem Then
        Set finditem = ReverseLookupItem(VarChanged)
        
    Else
        Set finditem = VarChanged.Tag
    End If
    If Not finditem Is Nothing Then
        
        finditem.SubItems(1) = mParser.ResultToString(VarChanged.Value)
        
        
    End If
End Sub

Private Sub mVariables_VarRemove(RemoveMe As BASeParserXP.CVariable)
    'remove that variable from the collection.
    Dim finditem As ListItem
    Set finditem = ReverseLookupItem(RemoveMe)
    LvwVariables.ListItems.Remove finditem.Index
End Sub

Private Sub mVariables_VarReplace(Replacing As BASeParserXP.CVariable, Allow As Boolean)
'
    Dim finditem As ListItem
    Set finditem = ReverseLookupItem(Replacing)
    LvwVariables.ListItems.Remove finditem.Index



End Sub

Private Sub UserControl_Resize()
LvwVariables.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
