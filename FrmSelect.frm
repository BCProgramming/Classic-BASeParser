VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select An Item"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LvwItemPick 
      Height          =   3840
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   6773
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2970
      TabIndex        =   0
      Top             =   4005
      Width           =   1005
   End
End
Attribute VB_Name = "FrmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'to be honest- the function that uses this form is kind of stupid. But- oh well.


Private mItemPicked As ListItem
Public Function ChooseItem(ArrData As Variant) As String
    Dim I As Long
    With LvwItemPick
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , , "Choose Item:"
        For I = LBound(ArrData) To UBound(ArrData)
            .ListItems.Add , , ArrData(I)
        Next I
    End With
    If Not Visible Then Show vbModal
    ChooseItem = mItemPicked.Text
End Function
Public Function ChooseItemEx(ArrData As Variant) As String
    Populate ArrData
    If Not Visible Then Show vbModal
    ChooseItemEx = JoinItem(mItemPicked)
End Function
Private Function JoinItem(LItem As ListItem)
    Dim arrMake As Variant, I As Long
    ReDim arrMake(1)
    arrMake(0) = LItem.Text
    I = 1
    On Error GoTo endLoop
    Do
        ReDim Preserve arrMake(I)
        arrMake(I) = LItem.SubItems(I)
        I = I + 1
        
    Loop
    ReDim Preserve arrMake(UBound(arrMake) - 1)
endLoop:
    JoinItem = Join(arrMake, ",")
End Function
Public Sub Populate(ArrData As Variant)
    'Format of ArrData:
    'An Array. the array should contain OTHER arrays- in this format:
    
    'Array(Array(Columnheader1,columnheader2...),(ItemArrays)...)
    'The ItemArray is similiar to the columnheader array, except
    'that it contains the sub-item text of the item.
    Dim ArrColumns, ArrItem As Variant, I As Long
    Dim sitems As Long
    If Not IsArray(ArrData) Then
        Err.Raise 13, , "Must give Array to InputItem"
    End If
    ArrColumns = ArrData(0)
    Me.LvwItemPick.ListItems.Clear
    LvwItemPick.ColumnHeaders.Clear
    For I = LBound(ArrColumns) To UBound(ArrColumns)
        LvwItemPick.ColumnHeaders.Add , , ArrColumns(I)
    Next I
    
    For I = 1 To UBound(ArrData)
    
        ArrItem = ArrData(I)
        With LvwItemPick.ListItems.Add
            .Text = ArrItem(0)
            For sitems = 1 To UBound(ArrItem)
                .SubItems(sitems) = ArrItem(sitems)
            
            Next
            
        End With
    Next I
End Sub

Private Sub CmdOK_Click()
Set mItemPicked = LvwItemPick.SelectedItem
Me.Hide
End Sub

Private Sub LvwItemPick_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmdOK.Enabled = True
End Sub
