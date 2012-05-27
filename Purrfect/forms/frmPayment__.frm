VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayment 
   Appearance      =   0  'Flat
   BackColor       =   &H00C8761C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment"
   ClientHeight    =   8340
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13092
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13092
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   8235
      Left            =   60
      ScaleHeight     =   8208
      ScaleWidth      =   12948
      TabIndex        =   0
      Top             =   60
      Width           =   12975
      Begin MSComctlLib.ListView lsvSales 
         Height          =   1752
         Left            =   2220
         TabIndex        =   4
         Top             =   1260
         Visible         =   0   'False
         Width           =   4332
         _ExtentX        =   7641
         _ExtentY        =   3090
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
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sales Order No."
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4680
         TabIndex        =   25
         Top             =   7440
         Width           =   3075
      End
      Begin VB.CommandButton cmdProcess 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PROCESS PAYMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   7320
         Width           =   4875
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5955
         Left            =   180
         ScaleHeight     =   5928
         ScaleWidth      =   12588
         TabIndex        =   6
         Top             =   1320
         Width           =   12615
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remarks"
            ForeColor       =   &H80000008&
            Height          =   2115
            Left            =   240
            TabIndex        =   20
            Top             =   3660
            Width           =   6615
            Begin MSComctlLib.ListView lsvRemarks 
               Height          =   1755
               Left            =   180
               TabIndex        =   21
               Top             =   240
               Width           =   6255
               _ExtentX        =   11028
               _ExtentY        =   3090
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
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "amount paid"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "balance"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "date paid"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin MSComctlLib.ListView lsvItemsPurchased 
            Height          =   2175
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   12075
            _ExtentX        =   21294
            _ExtentY        =   3831
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
               Size            =   9.6
               Charset         =   0
               Weight          =   700
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
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   9060
            TabIndex        =   23
            Top             =   5400
            Width           =   3255
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "BALANCE :"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7080
            TabIndex        =   22
            Top             =   5460
            Width           =   1695
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Discount :"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   19
            Top             =   4380
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NET TOTAL :"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8040
            TabIndex        =   18
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label lblNetTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10500
            TabIndex        =   17
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Visible         =   0   'False
            X1              =   12300
            X2              =   9540
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label lblDiscount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10500
            TabIndex        =   16
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label lblGrandTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10500
            TabIndex        =   15
            Top             =   4020
            Width           =   1695
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8340
            TabIndex        =   13
            Top             =   120
            Width           =   4095
         End
         Begin VB.Label lblAddress 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   12
            Top             =   1140
            Width           =   6195
         End
         Begin VB.Label lblCustomerName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   11
            Top             =   840
            Width           =   60
         End
         Begin VB.Label lblAgent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1920
            TabIndex        =   10
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Name: "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   9
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblSalesOrderNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   8
            Top             =   60
            Width           =   60
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Order No :"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   7
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   495
         Left            =   5940
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtSalesOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2220
         TabIndex        =   3
         Top             =   720
         Width           =   3675
      End
      Begin VB.Label lblFullyPaid 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FULLY PAID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   288
         Left            =   300
         TabIndex        =   27
         Top             =   7560
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Shape paidBorder 
         BorderColor     =   &H000000FF&
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   552
         Left            =   240
         Top             =   7440
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT PAID :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Top             =   7560
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order No :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   12720
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim active_sales_for_payment As Sales
Dim new_payment As Payment

Private Sub cmdBrowse_Click()
Call toogleListView(lsvSales)
End Sub

Private Sub cmdProcess_Click()
    'id, sales_order_no, amount, balance, payment_date, remarks
    
    new_payment.amount = Val(txtAmountPaid.Text)
    new_payment.balance = new_payment.active_sales.get_total_amount - new_payment.amount
    new_payment.savePayment
    new_payment.printPaymentInfoAndNewBalance
    Set new_payment = Nothing
    Set active_sales_for_payment = Nothing
    Call loadSalesInfo
    txtAmountPaid.Text = FormatNumber(0, 2)
End Sub

Private Sub Form_Load()
Call setSalesListview(lsvSales)
With lsvSales
    .ColumnHeaders(2).width = 0
    .ColumnHeaders(3).width = 0
    .ColumnHeaders(4).width = 0
    .ColumnHeaders(5).width = 0
    .ColumnHeaders(6).width = 0
    .ColumnHeaders(7).width = 0
    .ColumnHeaders(8).width = 0
    .ColumnHeaders(9).width = 0
    .ColumnHeaders(10).width = 0
End With
Call loadAllSalesToListview(lsvSales, False, PAYMENT_ACCOUNT_RECEIVABLE)
End Sub

Private Sub lsvSales_Click()

    lsvRemarks.ListItems.Clear
    
'Call toogleListView(lsvSales)
 Call loadSalesInfo
 
End Sub

Private Sub txtAmountPaid_Change()
    validateInputToBalance
    If Val(txtAmountPaid.Text) > new_payment.getActualBalance Then
        cmdProcess.Enabled = False
    Else
        cmdProcess.Enabled = True
    End If
End Sub
Sub validateInputToBalance()
    If Val(new_payment.getActualBalance) - Val(txtAmountPaid.Text) = 0 Then
        lblFullyPaid.Visible = True
        paidBorder.Visible = True
        new_payment.remarks = "fully paid"
    Else
        lblFullyPaid.Visible = False
        paidBorder.Visible = False
        new_payment.remarks = ""
    End If
End Sub
Private Sub txtSalesOrder_Click()
 Call toogleListView(lsvSales)
End Sub

Sub loadSalesInfo()
    Dim rs As New ADODB.Recordset
    Dim cart As New cart
    Dim items As New cart_items
    Dim list As ListItem
    Set active_sales_for_payment = New Sales
    Set new_payment = New Payment
    
    new_payment.sales_order_no = lsvSales.SelectedItem.Text
    active_sales_for_payment.loadSalesOrder (lsvSales.SelectedItem.Text)
    Set new_payment.active_sales = active_sales_for_payment
    txtSalesOrder.Text = lsvSales.SelectedItem.Text
    Call toogleListView(lsvSales)
    
    With active_sales_for_payment
        lblSalesOrderNo.Caption = .transaction_id
        lblCustomerName.Caption = .sold_to.customers_name
        lblAgent.Caption = .sold_to.mvaragent.agent_name
        lblAddress.Caption = .sold_to.customers_add
        lblDate.Caption = .date_transact
        'lblDiscount.Caption = FormatNumber(.get_discount_total(), 2)
        'lblGrandTotal.Caption = FormatNumber((.get_total_amount + .get_discount_total()) - .get_tracking_total(), 2)
        lblNetTotal.Caption = FormatNumber((.get_total_amount), 2)
        
        Set cart = .items_sold
        
        lsvItemsPurchased.ListItems.Clear
        
        For Each items In cart
            Set list = lsvItemsPurchased.ListItems.Add(, , items.qty_purchased)
            list.SubItems(1) = items.Item.unit_of_measure
            list.SubItems(2) = items.Item.item_description
            If .sold_to.dealers_type = DEALER Then
                list.SubItems(3) = items.Item.dealers_price
                list.SubItems(4) = FormatCurrency((items.Item.dealers_price - items.discount) + items.tracking_price, 2)
            Else
                list.SubItems(3) = items.Item.item_price
                list.SubItems(4) = FormatCurrency((items.Item.item_price - items.discount) + items.tracking_price, 2)
            End If
            
        Next
        
    End With
    Call new_payment.loadRemarksToListview(lsvRemarks)
    lblBalance.Caption = FormatNumber(new_payment.getActualBalance, 2)
    
    If Val(lblBalance.Caption) = 0 Then
        lblFullyPaid.Visible = True
        paidBorder.Visible = True
        new_payment.remarks = "fully paid"
    Else
        lblFullyPaid.Visible = False
        paidBorder.Visible = False
        new_payment.remarks = ""
    End If
End Sub
