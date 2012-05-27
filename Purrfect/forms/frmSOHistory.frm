VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSOHistory 
   Appearance      =   0  'Flat
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SO Hostory"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14610
   Icon            =   "frmSOHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C7FEF3&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   60
      ScaleHeight     =   8385
      ScaleWidth      =   14445
      TabIndex        =   0
      Top             =   60
      Width           =   14475
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
         Height          =   675
         Left            =   12000
         TabIndex        =   13
         Top             =   7620
         Width           =   2355
      End
      Begin MSComctlLib.ListView lsvHistory 
         Height          =   6075
         Left            =   120
         TabIndex        =   2
         Top             =   1500
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   10716
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sales Order"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Net Total"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Remarks"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date purchased"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date paid"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "prepared by"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Technician:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10920
         TabIndex        =   12
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label lblTechnician 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   12240
         TabIndex        =   11
         Top             =   1140
         Width           =   2085
      End
      Begin VB.Label lblContact 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   12240
         TabIndex        =   10
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer contact #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10320
         TabIndex        =   9
         Top             =   720
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbladdress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label lblAmountCustomer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   7800
         Width           =   1875
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   7800
         Width           =   855
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   2085
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   14340
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order History"
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
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   3795
      End
   End
End
Attribute VB_Name = "frmSOHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SOHistory As New Sales

Private Sub cmdPrint_Click()
Dim rs As New ADODB.Recordset
Set rs = getSalesOrderOfThisCustomer(SOHistory.sold_to.customers_id)

Set dtaSoHistory.DataSource = rs
dtaSoHistory.Sections(1).Controls("lblCustomerName").Caption = SOHistory.sold_to.customers_name
dtaSoHistory.Sections(1).Controls("lblAddress").Caption = SOHistory.sold_to.customers_add
dtaSoHistory.Sections(1).Controls("lblDate").Caption = Date
dtaSoHistory.Show 1
End Sub

Private Sub Form_Load()
SOHistory.loadSalesOrder selectedSOForHistory
lblCustomerName.Caption = SOHistory.sold_to.customers_name
lblAddress.Caption = SOHistory.sold_to.customers_add
lblContact.Caption = SOHistory.sold_to.customers_number
lblTechnician.Caption = SOHistory.sold_to.mvaragent.agent_name
Call loadSalesOrderOfCustomerToListview(SOHistory.sold_to.customers_id, lsvHistory)
lblAmountCustomer.Caption = FormatNumber(getTotalAmountOfAccountReceivableOfThisCustomer(SOHistory.sold_to.customers_id), 2)
End Sub

Sub loadSOOthisCustomer()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim list As ListItem

sql = "SELECT sales_order_no FROM `stock_out_transaction` where responsible_customer = " & SOHistory.sold_to.customers_id
Set rs = db.execute(sql)
If rs.RecordCount > 0 Then
    Dim temp As New Sales
    temp.loadSalesOrder rs.Fields(0).Value
    Set list = lsvHistory.ListItems.Add(, , rs.Fields(0).Value)
       
End If
End Sub

Private Sub lsvHistory_Click()
If lsvHistory.ListItems.Count > 0 Then
    activeSalesOrderForViewSalesDetails = lsvHistory.SelectedItem.Text
    frmViewSalesInDetails.Show 1
End If
End Sub
