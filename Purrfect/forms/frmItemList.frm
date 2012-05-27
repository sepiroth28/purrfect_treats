VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item List"
   ClientHeight    =   6405
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   11775
   Icon            =   "frmItemList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   732
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   11565
      TabIndex        =   5
      Top             =   4860
      Width           =   11595
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   8760
         ScaleHeight     =   525
         ScaleWidth      =   2685
         TabIndex        =   7
         Top             =   60
         Width           =   2715
         Begin VB.Label lblAvailability 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C8761C&
            Height          =   315
            Left            =   180
            TabIndex        =   8
            Top             =   120
            Width           =   2355
         End
      End
      Begin VB.Label lblSelectedItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   6
         Top             =   180
         Width           =   5055
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   8340
      TabIndex        =   1
      Text            =   "1"
      Top             =   5700
      Width           =   1632
   End
   Begin VB.TextBox txtSearchItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   60
      TabIndex        =   0
      Top             =   5700
      Width           =   4332
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "SELECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   10080
      TabIndex        =   2
      Top             =   5700
      Width           =   1572
   End
   Begin MSComctlLib.ListView lsvItemList 
      Height          =   4695
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7500
      TabIndex        =   4
      Top             =   5820
      Width           =   810
   End
End
Attribute VB_Name = "frmItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelect_Click()
    amount_to_be_debt = activeSales.get_total_amount
    Dim items As New cart_items
    
    If activeSales.isSoldToWalkIn Then
             Call items.Item.load_item(Val(lsvItemList.SelectedItem.Text))
            items.qty_purchased = Val(txtQty.Text)
            items.tracking_price = getTrackingPriceOfCurrentCustomer(activeSales.sold_to.customers_id)
            If items.Item.item_qty >= items.qty_purchased Then
                Call activeSales.items_sold.Add(items)
                Call loadActiveCartItems(frmMenu.lsvItemsInCart)
                Call updateTotalAmount
                
                Unload Me
            Else
                MsgBox "Cannot transact, insufficient stock remain", vbInformation, "Insufficient stock"
            End If
    Else
            If isInLimit(activeSales.sold_to.customers_id) Then
                MsgBox "Customers reach his/her credit limit...Please refer to the SO history of this customer", vbInformation, "Credit Limit reached"
            Else
                
                     Call items.Item.load_item(Val(lsvItemList.SelectedItem.Text))
                         items.item_price = getPriceToBeUsed(items)
                        amount_to_be_debt = activeSales.get_total_amount + items.item_price
                    If isInLimit(activeSales.sold_to.customers_id) Then
                        MsgBox "Customers reach his/her credit limit...Please refer to the SO history of this customer", vbInformation, "Credit Limit reached"
                    Else
                        items.qty_purchased = Val(txtQty.Text)
                        items.tracking_price = getTrackingPriceOfCurrentCustomer(activeSales.sold_to.customers_id)
                        If items.Item.item_qty >= items.qty_purchased Then
                            Call activeSales.items_sold.Add(items)
                            Call loadActiveCartItems(frmMenu.lsvItemsInCart)
                            Call updateTotalAmount
                            
                            Unload Me
                        Else
                            MsgBox "Cannot transact, insufficient stock remain", vbInformation, "Insufficient stock"
                        End If
                    End If
                    
            End If
    End If
    
End Sub


Private Sub Form_Load()
Call setItemsDescriptionColumns(lsvItemList)
lsvItemList.ColumnHeaders(1).width = 0
lsvItemList.ColumnHeaders(2).width = 4000
lsvItemList.ColumnHeaders(3).width = 0
lsvItemList.ColumnHeaders(4).width = 0
lsvItemList.ColumnHeaders(5).width = 2000
lsvItemList.ColumnHeaders(6).width = 2000
lsvItemList.ColumnHeaders(7).width = 2000
lsvItemList.ColumnHeaders(8).width = 0
lsvItemList.ColumnHeaders(9).width = 0
Call loadAllItemsToListview(lsvItemList, "item_code")
cmdSelect.Enabled = False
End Sub

Private Sub lsvItemList_Click()
If lsvItemList.ListItems.Count > 0 Then
    Dim Item As New items
    lblSelectedItem.Caption = lsvItemList.SelectedItem.SubItems(2)
    
    Item.load_item (Val(lsvItemList.SelectedItem.Text))
    lblAvailability.Caption = Item.displayAvailability
        If Item.checkStockQty Then
            cmdSelect.Enabled = True
        Else
            cmdSelect.Enabled = False
        End If
    Set Item = Nothing
End If
End Sub

Private Sub lsvItemList_DblClick()

    Dim items As New cart_items
    Call items.Item.load_item(Val(lsvItemList.SelectedItem.Text))
    items.item_price = items.Item.item_price
    items.qty_purchased = Val(txtQty.Text)
    
    Call activeSales.items_sold.Add(items)
    Call loadActiveCartItems(frmMenu.lsvItemsInCart)
    Call updateTotalAmount
    Unload Me
    
End Sub

Private Sub lsvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub txtSearchItem_Change()
Dim rs As New ADODB.Recordset
Set rs = searchItemsByItemCode(txtSearchItem.Text)
Call loadItemRSToListCiew(lsvItemList, rs)
End Sub

Private Sub txtSearchItem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

