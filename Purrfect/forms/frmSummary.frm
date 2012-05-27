VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummary 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Summary"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10665
   Icon            =   "frmSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   9075
      Left            =   60
      ScaleHeight     =   9045
      ScaleWidth      =   10545
      TabIndex        =   1
      Top             =   60
      Width           =   10572
      Begin VB.CommandButton cmdDone 
         Caption         =   "DONE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6780
         TabIndex        =   9
         Top             =   8100
         Width           =   3255
      End
      Begin VB.TextBox txtTenderedAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5940
         TabIndex        =   0
         Top             =   6480
         Width           =   4035
      End
      Begin MSComctlLib.ListView lsvItems 
         Height          =   3312
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   10212
         _ExtentX        =   18018
         _ExtentY        =   5847
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Unit Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tracking price"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblTrackingPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   5940
         TabIndex        =   16
         Top             =   5400
         Width           =   4032
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tracking price total : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   2160
         TabIndex        =   15
         Top             =   5400
         Width           =   3672
      End
      Begin VB.Label lblReferenceNo 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   4980
         TabIndex        =   14
         Top             =   180
         Width           =   3675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NET TOTAL : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   2160
         TabIndex        =   13
         Top             =   5940
         Width           =   3672
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT TOTAL : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   2160
         TabIndex        =   12
         Top             =   4920
         Width           =   3672
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GRAND TOTAL : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   2160
         TabIndex        =   11
         Top             =   4440
         Width           =   3672
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00325641&
         Height          =   552
         Left            =   6300
         TabIndex        =   8
         Top             =   7380
         Width           =   3672
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   2160
         TabIndex        =   7
         Top             =   7440
         Width           =   3672
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TENDERED AMOUNT : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C8761C&
         Height          =   552
         Left            =   1620
         TabIndex        =   6
         Top             =   6600
         Width           =   4212
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   5940
         TabIndex        =   5
         Top             =   5940
         Width           =   4032
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   552
         Left            =   5940
         TabIndex        =   4
         Top             =   4860
         Width           =   4032
      End
      Begin VB.Label lblGrandTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   5940
         TabIndex        =   3
         Top             =   4320
         Width           =   4032
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   9960
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   8580
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUMMARY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public done As Boolean
Public tenderedAmount As Double
Private Sub cmdDone_Click()
If done Then
    Dim cart As New cart
    Dim items As New cart_items
    Dim list As ListItem
    Dim ctr As Integer
    Set cart = activeSales.items_sold
    Dim walk_in_id As Integer
    'name,qty,price, total
    
   
    For Each items In cart
       With items
            .Item.stockOut (.qty_purchased)
       End With
    Next
    activeSales.tendered_amount = Val(txtTenderedAmount.Text)
    activeSales.change = activeSales.tendered_amount - activeSales.get_total_amount
    If activeSales.isSoldToWalkIn Then
        walk_in_id = activeSales.sold_to.insert
        Set activeSales.sold_to = New Customers
        activeSales.sold_to.load_customers (walk_in_id)
    End If
    activeSales.prepared_by = activeUser.username
    activeSales.save_sales
    activeSales.updateReferenceNo
    activeSales.printDeliveryReceipt
    Call prepareNewTransaction
    Unload Me
End If
End Sub

Private Sub Form_Load()
done = False
tenderedAmount = 0
lblReferenceNo.Caption = activeSales.transaction_id
Call loadActiveCartItems(lsvItems)
lblGrandTotal.Caption = FormatCurrency(activeSales.get_total_amount + activeSales.get_discount_total, 2)

If activeSales.hasDiscount Then
    lblNetTotal.Caption = FormatCurrency(activeSales.get_total_amount, 2)
Else
    lblDiscount.Caption = FormatCurrency(activeSales.get_discount_total, 2)
    lblTrackingPrice.Caption = FormatCurrency(activeSales.get_tracking_total, 2)
    lblNetTotal.Caption = FormatCurrency(activeSales.get_total_amount, 2)
End If

If activeSales.payment_type = PAYMENT_COD Then
    txtTenderedAmount.Enabled = True
Else
    done = True
    txtTenderedAmount.Enabled = False
End If
End Sub

Private Sub txtTenderedAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tenderedAmount = Val(txtTenderedAmount.Text)
    If activeSales.payment_type = PAYMENT_COD Then
        If Val(tenderedAmount) >= Val(activeSales.get_total_amount) Then
            Dim change As Double
            change = Val(tenderedAmount) - Val(activeSales.get_total_amount())
            lblChange.Caption = FormatCurrency(change, 2)
            done = True
            txtTenderedAmount.BackColor = &HFFFFFF
            cmdDone.SetFocus
        Else
            txtTenderedAmount.BackColor = &H80FF&
        End If
    End If
End If
End Sub
