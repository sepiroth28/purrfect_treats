VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewSalesInDetails 
   BackColor       =   &H00C8761C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail Items"
   ClientHeight    =   8820
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   9750
   Icon            =   "frmViewSalesInDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8595
      Left            =   120
      ScaleHeight     =   8565
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   120
      Width           =   9555
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6840
         TabIndex        =   3
         Top             =   7800
         Width           =   2532
      End
      Begin MSComctlLib.ListView lsvDetailSales 
         Height          =   5415
         Left            =   120
         TabIndex        =   1
         Top             =   1740
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9551
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Quantity"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   4410
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Label lblPreparedBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1800
         TabIndex        =   13
         Top             =   7320
         Width           =   45
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2100
         TabIndex        =   12
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   10
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   9300
         TabIndex        =   9
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblTotals 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   5160
         TabIndex        =   8
         Top             =   7260
         Width           =   3915
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2100
         TabIndex        =   7
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer name: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lblSo 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S.O. : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10200
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Deatailed Per Items"
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
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   3732
      End
   End
End
Attribute VB_Name = "frmViewSalesInDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    loadTolistInDetails
End Sub

'####################################################################################################
'### NOTE THIS IS USED FOR RETRIEVING RECORDS WHICH DOES NOT AFFECT THE CHANGING OF PRICE
'#####################################################################################################
Sub loadTolistInDetails()

Dim s As New Sales
    Dim cart As New cart
    Dim items As New cart_items
    Dim list As ListItem
     
    With s
        .loadSalesOrder (activeSalesOrderForViewSalesDetails)
        lblSo.Caption = activeSalesOrderForViewSalesDetails
        lblCustomerName.Caption = s.sold_to.customers_name
        lblAddress.Caption = s.sold_to.customers_add
        
            lblDate.Caption = FormatDateTime(s.date_transact, vbLongDate)
        Set cart = s.items_sold
        On Error Resume Next
         Dim amount_of_items As Double
         Dim grand_total As Double
         Dim total_trucking_amount As Double
         Dim total_to_discounted As Double
         Dim original_price As Double
         
            For Each items In cart
                Set list = lsvDetailSales.ListItems.Add(, , items.qty_purchased)
                list.SubItems(1) = items.Item.unit_of_measure
                list.SubItems(2) = items.Item.item_description
                If .sold_to.dealers_type = DEALER Then
                    'list.SubItems(3) = FormatNumber(items.Item.dealers_price, 2)
                    'modified, should not be based on items price...
                    list.SubItems(3) = FormatNumber(items.item_price, 2)
                    
                    
                    total_trucking_amount = items.tracking_price * items.qty_purchased
                    total_to_discounted = items.discount * items.qty_purchased
                    original_price = ((items.total_price + total_to_discounted) - total_trucking_amount) / items.qty_purchased
                    amount_of_items = (original_price * items.qty_purchased) + total_trucking_amount    '((items.item_price - items.discount) + items.tracking_price) * items.qty_purchased
                    list.SubItems(4) = FormatNumber(amount_of_items, 2)
                    grand_total = (grand_total + amount_of_items) - total_to_discounted
                    'FormatNumber(getTotalAmountOfAccountReceivableOfThisCustomer(customer_id), 2
                Else
                    list.SubItems(3) = FormatNumber(items.item_price, 2)
                    
                    total_trucking_amount = items.tracking_price * items.qty_purchased
                    total_to_discounted = items.discount * items.qty_purchased
                    original_price = ((items.total_price + total_to_discounted) - total_trucking_amount) / items.qty_purchased
                   amount_of_items = (original_price * items.qty_purchased) + total_trucking_amount
                    
                    list.SubItems(4) = FormatNumber(amount_of_items, 2)
                    grand_total = (grand_total + amount_of_items) - total_to_discounted
                End If
        
            Next
        'lblTotals.Caption = FormatNumber(s.get_total_amount, 2)
        lblTotals.Caption = FormatNumber(grand_total, 2)
        lblPreparedBy.Caption = s.prepared_by
    End With
End Sub

