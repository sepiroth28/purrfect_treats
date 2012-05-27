VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjustSaleTransaction 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjust Sales Transaction"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmAdjustSaleTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   7635
      Left            =   60
      ScaleHeight     =   7605
      ScaleWidth      =   14925
      TabIndex        =   0
      Top             =   60
      Width           =   14955
      Begin MSComctlLib.ListView lsvCODList 
         Height          =   2172
         Left            =   480
         TabIndex        =   30
         Top             =   3780
         Visible         =   0   'False
         Width           =   4092
         _ExtentX        =   7223
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000018&
         Caption         =   "Affected Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2112
         Left            =   8520
         TabIndex        =   28
         Top             =   5340
         Width           =   6192
         Begin MSComctlLib.ListView lsvAffectedItems 
            Height          =   1632
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   5952
            _ExtentX        =   10504
            _ExtentY        =   2884
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Quantity"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Unit Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10080
         TabIndex        =   27
         Top             =   4500
         Width           =   2235
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12420
         TabIndex        =   25
         Top             =   4500
         Width           =   2235
      End
      Begin VB.CheckBox chkNetTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "Net Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9120
         TabIndex        =   24
         Top             =   3600
         Width           =   1272
      End
      Begin VB.TextBox txtNetTotal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10860
         TabIndex        =   23
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CheckBox chkDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9120
         TabIndex        =   22
         Top             =   2940
         Width           =   1452
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10860
         TabIndex        =   21
         Top             =   2880
         Width           =   2655
      End
      Begin VB.CheckBox chkGrandTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "Grand Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9120
         TabIndex        =   20
         Top             =   2340
         Width           =   1572
      End
      Begin VB.TextBox txtGrandTotal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10860
         TabIndex        =   19
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox chkSoldTo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "Sold To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9120
         TabIndex        =   18
         Top             =   1260
         Width           =   1632
      End
      Begin VB.CommandButton cmdBrowseCustomer 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   435
         Left            =   13560
         TabIndex        =   17
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtSoldTo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10860
         TabIndex        =   15
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000018&
         Caption         =   "COD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2115
         Left            =   180
         TabIndex        =   8
         Top             =   5280
         Width           =   8235
         Begin VB.CommandButton cmdSOList 
            Caption         =   "..."
            Height          =   492
            Left            =   7260
            TabIndex        =   26
            Top             =   720
            Width           =   672
         End
         Begin VB.CommandButton cmdLoadSalesOrder 
            Caption         =   "Load"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6060
            TabIndex        =   11
            Top             =   1320
            Width           =   1875
         End
         Begin VB.TextBox txtSalesOrder 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   300
            TabIndex        =   9
            Top             =   720
            Width           =   6855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Order No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   300
            TabIndex        =   10
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000018&
         Caption         =   "ACCOUNT RECEIVABLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   4515
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   8235
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   495
            Left            =   3960
            TabIndex        =   4
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox txtCustomer 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   3
            Top             =   780
            Width           =   3675
         End
         Begin MSComctlLib.ListView lsvCustomerList 
            Height          =   2175
            Left            =   240
            TabIndex        =   5
            Top             =   1260
            Visible         =   0   'False
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   3836
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "customer_id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "customer_name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lsvSales 
            Height          =   2715
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4789
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sales Order No."
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Net total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "remarks"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "date"
               Object.Width           =   4410
            EndProperty
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   7
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   8580
         X2              =   14640
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label lblAgent 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10860
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9120
         TabIndex        =   14
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lblSO 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10860
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order No : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9120
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   12240
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Sales Transaction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmAdjustSaleTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_sales_order As Sales
Dim update_values(3) As String

Private Sub chkDiscount_Click()
If chkDiscount.Value Then
    txtDiscount.Enabled = True
Else
    txtDiscount.Enabled = False
End If
End Sub

Private Sub chkGrandTotal_Click()
If chkGrandTotal.Value Then
    txtGrandTotal.Enabled = True
Else
    txtGrandTotal.Enabled = False
End If
End Sub

Private Sub chkNetTotal_Click()
If chkNetTotal.Value Then
    txtNetTotal.Enabled = True
Else
    txtNetTotal.Enabled = False
End If
End Sub

Private Sub chkSoldTo_Click()
If chkSoldTo.Value Then
    txtSoldTo.Enabled = True
    cmdBrowseCustomer.Enabled = True
Else
    txtSoldTo.Enabled = False
    cmdBrowseCustomer.Enabled = False
End If
End Sub

Private Sub chkTotal_Click()
If chkTotal.Value Then
    txtNetTotal.Enabled = True
Else
    txtNetTotal.Enabled = False
End If
End Sub

Private Sub cmdBrowse_Click()
Call toogleListView(lsvCustomerList)
End Sub

Private Sub cmdDelete_Click()
Dim ans As Byte

ans = MsgBox("Are you sure you want to delete this SALES ORDER?", vbYesNo, "Delete SO")
If ans = vbYes Then
    Dim sql As String
    sql = "INSERT INTO deleted_so (SELECT * from stock_out_transaction WHERE sales_order_no = '" & edit_sales_order.transaction_id & "')"
    db.execute sql
    edit_sales_order.delete
    Call clearData
    txtCustomer.Text = lsvCustomerList.SelectedItem.SubItems(1)
    Call loadSalesOrderOfCustomerToListview(Val(lsvCustomerList.SelectedItem.Text), lsvSales)
End If
End Sub

Private Sub cmdLoadSalesOrder_Click()
Set edit_sales_order = New Sales
edit_sales_order.loadSalesOrder (txtSalesOrder.Text)
End Sub

Private Sub cmdSOList_Click()
Call toogleListView(lsvCODList)
End Sub

Private Sub cmdUpdate_Click()
Dim x As Integer
Dim update As String

x = 0
If chkSoldTo.Value Then
    update_values(x) = "sold_to = '" & txtSoldTo.Text & "'"
    x = x + 1
End If

If chkGrandTotal.Value Then
    update_values(x) = "grand_total = " & Val(txtGrandTotal.Text)
    x = x + 1
End If

If chkDiscount.Value Then
    update_values(x) = "discount = " & Val(txtDiscount.Text)
    x = x + 1
End If

If chkNetTotal.Value Then
    update_values(x) = "net_total =" & Val(txtNetTotal.Text)
    x = x + 1
End If
Dim update_values2() As String
ReDim update_values2(x - 1)
Dim y As Integer

y = 0
For y = 0 To x - 1
    update_values2(y) = update_values(y)
Next y

    
update = "UDPATE stock_out_transaction SET " & Join(update_values2, ",") & " WHERE sales_order_no = '" & edit_sales_order.transaction_id & "'"
MsgBox update
End Sub

Private Sub Form_Load()
lsvCustomerList.ColumnHeaders(1).width = 0
lsvCustomerList.ColumnHeaders(2).width = 4000
lsvCustomerList.ColumnHeaders(3).width = 0
lsvCustomerList.ColumnHeaders(4).width = 0
lsvCustomerList.ColumnHeaders(5).width = 0

Call loadAllCustomersToListview(lsvCustomerList)
Call loadAllCODToListview(lsvCODList)
End Sub

Private Sub lsvCODList_Click()
On Error Resume Next
    txtSalesOrder.Text = lsvCODList.SelectedItem.Text
    Call toogleListView(lsvCODList)

End Sub

Private Sub lsvCustomerList_Click()
Call toogleListView(lsvCustomerList)
txtCustomer.Text = lsvCustomerList.SelectedItem.SubItems(1)
Call loadSalesOrderOfCustomerToListview(Val(lsvCustomerList.SelectedItem.Text), lsvSales)
End Sub

Private Sub lsvSales_DblClick()
Set edit_sales_order = New Sales
edit_sales_order.loadSalesOrder (lsvSales.SelectedItem.Text)
Call loadAffectedItems
Call renderSalesOrderData
End Sub
Sub loadAffectedItems()
    Dim cart As New cart
    Dim items As New cart_items
    Dim list As ListItem
    
    Set cart = edit_sales_order.items_sold
    lsvAffectedItems.ListItems.Clear
    For Each items In cart
        Set list = lsvAffectedItems.ListItems.Add(, , items.qty_purchased)
        list.SubItems(1) = items.Item.unit_of_measure
        list.SubItems(2) = items.Item.item_description
        If edit_sales_order.sold_to.dealers_type = DEALER Then
            list.SubItems(3) = items.Item.dealers_price
            list.SubItems(4) = FormatCurrency((items.Item.dealers_price - items.discount) + items.tracking_price, 2)
        Else
            list.SubItems(3) = items.Item.item_price
            list.SubItems(4) = FormatCurrency((items.Item.item_price - items.discount) + items.tracking_price, 2)
        End If
        
    Next

End Sub
Private Sub txtCustomer_Change()
Dim rs As New ADODB.Recordset
Set rs = searchCustomersByName(txtCustomer)
Call loadCustomerRSToListView(lsvCustomerList, rs)
End Sub

Sub renderSalesOrderData()
    With edit_sales_order
        lblSo.Caption = .transaction_id
        txtSoldTo.Text = .sold_to.customers_name
        lblAgent.Caption = .sold_to.mvaragent.agent_name
        txtGrandTotal.Text = .info_grand_total
        txtNetTotal.Text = .info_net_total
        txtDiscount.Text = .info_discount
    End With
End Sub
Sub clearData()
txtSoldTo.Text = ""
txtDiscount.Text = ""
txtGrandTotal.Text = ""
txtNetTotal.Text = ""
lblAgent.Caption = ""
lblSo.Caption = ""
End Sub

Sub disableTextboxes()
txtSoldTo.Enabled = False
txtDiscount.Enabled = False
txtGrandTotal.Enabled = False
txtNetTotal.Enabled = False
End Sub

Private Sub txtSalesOrder_Click()
Call toogleListView(lsvCODList)
End Sub
