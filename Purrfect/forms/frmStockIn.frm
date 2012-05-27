VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockIn 
   BackColor       =   &H00404040&
   Caption         =   "Manage Stock-in"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   11670
   Icon            =   "frmStockIn.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7320
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   60
      ScaleHeight     =   7185
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   60
      Width           =   11535
      Begin MSComctlLib.ListView lsvItemList 
         Height          =   3435
         Left            =   240
         TabIndex        =   9
         Top             =   1860
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6059
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   8880
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6360
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ + ] ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   8880
         MaskColor       =   &H00000080&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5340
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00325641&
         Height          =   675
         Left            =   7080
         TabIndex        =   5
         Top             =   5340
         Width           =   1575
      End
      Begin VB.TextBox txtItemCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00325641&
         Height          =   675
         Left            =   240
         TabIndex        =   4
         Top             =   5340
         Width           =   6735
      End
      Begin MSComctlLib.ListView lsvStockIn 
         Height          =   3975
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   2837822
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7080
         TabIndex        =   6
         Top             =   4980
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   300
         TabIndex        =   3
         Top             =   4920
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   240
         X2              =   11280
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "MANAGE STOCK - IN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   3375
      End
   End
   Begin VB.Menu mnu_stock_in_file 
      Caption         =   "File"
      Begin VB.Menu mnu_remove_item 
         Caption         =   "Remove Item"
      End
   End
End
Attribute VB_Name = "frmStockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemIdToStockIn As Integer
Dim lsvItemListIsClicked As Boolean

Private Sub cmdAdd_Click()
Dim lst As ListItem

    Set lst = lsvStockIn.ListItems.Add(, , itemIdToStockIn)
    lst.SubItems(1) = txtItemCode.Text
    lst.SubItems(2) = Val(txtQty.Text)
    
'        item_to_stock.items.load_item (itemIdToStockIn)
'        item_to_stock.QtyToBeAdd = Val(txtQty.Text)
'        Call activeStockInList.Add(item_to_stock)
'
    itemIdToStockIn = 0
       
    clearTextField
    txtItemCode.SetFocus
End Sub

Private Sub cmdNext_Click()
Dim list As ListItem
'Dim lst As ListItem
'If lsvStockIn.ListItems.Count > 0 Then
'    For Each lst In lsvStockIn.ListItems
'        Dim item_stockin As New items
'        item_stockin.load_item (Val(lst.Text))
'        item_stockin.addStock (Val(lst.SubItems(2)))
'        MsgBox "added to " & item_stockin.item_name & " : " & Val(lst.SubItems(2)) & vbCrLf & "Total: "
'    Next
'End If
Dim item_to_stock As New StockIn

For Each list In lsvStockIn.ListItems
        itemIdToStockIn = Val(list.Text)
        item_to_stock.items.load_item (itemIdToStockIn)
        item_to_stock.QtyToBeAdd = Val(list.SubItems(2))
        Call activeStockInList.Add(item_to_stock)
Next
    
    If lsvItemList.ListItems.Count > 0 Then
        Unload Me
        frmStockInPreview.Show 1
          
    Else
        MsgBox "Please add item to stock...", vbInformation, "Item is empty"
    End If

End Sub


Private Sub Form_Load()
Set activeStockInList = New StockInCollection

lsvItemListIsClicked = False
setStockInListview lsvStockIn
lsvStockIn.ColumnHeaders(1).width = 0
lsvStockIn.ColumnHeaders(2).width = 8000
lsvStockIn.ColumnHeaders(3).width = 2000
Call setItemsDescriptionColumns(lsvItemList)



Call hideAllColumnsExept("Item Code", lsvItemList)
lsvItemList.ColumnHeaders(2).width = 4000
lsvItemList.ColumnHeaders(3).width = 4000
Call loadAllItemsToListview(lsvItemList, "item_code")
End Sub

Private Sub lsvItemList_Click()
    If lsvItemList.ListItems.Count > 0 Then
        lsvItemListIsClicked = True
        itemIdToStockIn = Val(lsvItemList.SelectedItem.Text)
        txtItemCode.Text = lsvItemList.SelectedItem.SubItems(2)
        lsvItemList.Visible = False
        txtQty.SetFocus
    End If
     lsvItemListIsClicked = False
End Sub

Private Sub lsvStockIn_Click()
    lsvItemList.Visible = False
    txtItemCode.Text = ""
End Sub

Private Sub lsvStockIn_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnu_stock_in_file
End If
End Sub

Private Sub mnu_remove_item_Click()
If lsvStockIn.ListItems.Count > 0 Then
    lsvStockIn.ListItems.Remove (lsvStockIn.SelectedItem.Index)
    
End If
End Sub

Private Sub Picture1_Click()
    lsvItemList.Visible = False
    txtItemCode.Text = ""
End Sub

Private Sub txtItemCode_Change()
  Dim rs As New ADODB.Recordset
    If lsvItemListIsClicked = False Then
        lsvItemList.Visible = True
        Call loadItemRSToListCiew(lsvItemList, searchItemsByItemCode(txtItemCode.Text))
            If lsvItemList.ListItems.Count > 0 Then
              lsvItemList.DropHighlight = lsvItemList.ListItems.Item(1)
            End If
    End If
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'MsgBox lsvItemList.DropHighlight.SubItems(1)
    If lsvStockIn.ListItems.Count > 0 Then
        itemIdToStockIn = Val(lsvItemList.DropHighlight.Text)
        txtItemCode.Text = lsvItemList.DropHighlight.SubItems(2)
        lsvItemList.Visible = False
        txtQty.SetFocus
    End If
End If
End Sub

Private Sub txtItemCode_LostFocus()
'If lsvItemListIsClicked = False Then
'    Call toogleListView(lsvItemList)
'End If
End Sub

Sub clearTextField()
    txtItemCode.Text = ""
    lsvItemList.Visible = False
    txtQty.Text = ""
End Sub

Private Sub txtQty_GotFocus()
    lsvItemList.Visible = False
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
   Case 8, 48 To 57  ' BS, 0 - 9
   Case 13
    cmdAdd.SetFocus
   Case Else
     KeyAscii = 0
 End Select
 
End Sub
