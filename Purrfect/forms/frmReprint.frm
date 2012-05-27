VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReprint 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint Sales Order"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   Icon            =   "frmReprint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   8352
      Left            =   60
      ScaleHeight     =   8325
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   60
      Width           =   5775
      Begin MSComctlLib.ListView lsvCODList 
         Height          =   2172
         Left            =   480
         TabIndex        =   15
         Top             =   3660
         Visible         =   0   'False
         Width           =   4272
         _ExtentX        =   7541
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
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   792
         Left            =   3420
         TabIndex        =   12
         Top             =   7380
         Width           =   2172
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
         TabIndex        =   5
         Top             =   600
         Width           =   5415
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
            TabIndex        =   7
            Top             =   780
            Width           =   3675
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   495
            Left            =   3960
            TabIndex        =   6
            Top             =   780
            Width           =   615
         End
         Begin MSComctlLib.ListView lsvCustomerList 
            Height          =   2175
            Left            =   240
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   1560
            Width           =   4935
            _ExtentX        =   8705
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
            NumItems        =   3
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
            TabIndex        =   10
            Top             =   420
            Width           =   1815
         End
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
         TabIndex        =   1
         Top             =   5160
         Width           =   5415
         Begin VB.CommandButton cmdBrowseCODList 
            Caption         =   "..."
            Height          =   492
            Left            =   4620
            TabIndex        =   14
            Top             =   720
            Width           =   492
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
            TabIndex        =   3
            Top             =   720
            Width           =   4272
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
            Left            =   3240
            TabIndex        =   2
            Top             =   1320
            Width           =   1875
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
            TabIndex        =   4
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.Label lblSOFound 
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
         Height          =   732
         Left            =   240
         TabIndex        =   13
         Top             =   7440
         Width           =   3072
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRINT SALES ORDER"
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
         TabIndex        =   11
         Top             =   120
         Width           =   2955
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   12240
         Y1              =   420
         Y2              =   420
      End
   End
End
Attribute VB_Name = "frmReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim print_sale As Sales
Private Sub cmdBrowse_Click()
Call toogleListView(lsvCustomerList)
End Sub

Private Sub cmdBrowseCODList_Click()
Call toogleListView(lsvCODList)
End Sub

Private Sub cmdLoadSalesOrder_Click()
If txtSalesOrder.Text <> "" Then
    print_sale.loadSalesOrder (txtSalesOrder.Text)
    Set activeSales = print_sale
    lblSOFound.Caption = print_sale.transaction_id & " is ready to print..."
End If
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
    Call activeSales.printDeliveryReceipt

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
Call toogleListView(lsvCODList)
If lsvCODList.ListItems.Count Then
    txtSalesOrder.Text = lsvCODList.SelectedItem.Text
End If

End Sub

Private Sub lsvCustomerList_Click()
Call toogleListView(lsvCustomerList)
txtCustomer.Text = lsvCustomerList.SelectedItem.SubItems(1)
Call loadSalesOrderOfCustomerToListview(Val(lsvCustomerList.SelectedItem.Text), lsvSales)

End Sub

Private Sub lsvSales_DblClick()
Set print_sale = New Sales
print_sale.loadSalesOrder (lsvSales.SelectedItem.Text)
print_sale.payment_type = PAYMENT_ACCOUNT_RECEIVABLE
Set activeSales = print_sale
lblSOFound.Caption = print_sale.transaction_id & " is ready to print..."
End Sub

Private Sub txtCustomer_Change()
Dim rs As New ADODB.Recordset
Set rs = searchCustomersByName(txtCustomer)
Call loadCustomerRSToListView(lsvCustomerList, rs)
End Sub

Private Sub txtSalesOrder_Click()
Call toogleListView(lsvCODList)
End Sub
