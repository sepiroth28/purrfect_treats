VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCreditLimit 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Credit Limit"
   ClientHeight    =   7260
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5055
   Icon            =   "frmCreditLimit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   7152
      Left            =   60
      ScaleHeight     =   7125
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   60
      Width           =   4932
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   3180
         TabIndex        =   7
         Top             =   6360
         Width           =   1512
      End
      Begin VB.TextBox txtLimit 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   1740
         TabIndex        =   6
         Top             =   5460
         Width           =   1572
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   1740
         TabIndex        =   4
         Top             =   4320
         Width           =   2952
      End
      Begin MSComctlLib.ListView lsvCustomer 
         Height          =   3612
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   4512
         _ExtentX        =   7964
         _ExtentY        =   6376
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   372
         Left            =   1740
         TabIndex        =   9
         Top             =   4980
         Width           =   2952
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   180
         TabIndex        =   8
         Top             =   4980
         Width           =   1512
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   240
         X2              =   4680
         Y1              =   6180
         Y2              =   6180
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   180
         TabIndex        =   5
         Top             =   5580
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Customer"
         Height          =   432
         Left            =   180
         TabIndex        =   3
         Top             =   4440
         Width           =   1272
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   180
         X2              =   4752
         Y1              =   480
         Y2              =   492
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Credit Limit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   2112
      End
   End
End
Attribute VB_Name = "frmCreditLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim customer_credit As Credit

Private Sub cmdUpdate_Click()
    customer_credit.limit = Val(txtLimit.Text)
    customer_credit.update
    Set customer_credit = Nothing
    txtLimit.Text = ""
    lblName.Caption = ""
    MsgBox "Successfully updated...", vbOKOnly, "Update"
End Sub

Private Sub Form_Load()
Call loadAllCustomersToListview(lsvCustomer)
End Sub

Private Sub lsvCustomer_DblClick()
    Set customer_credit = New Credit
    Call customer_credit.customer_info.load_customers(Val(lsvCustomer.SelectedItem.Text))
    lblName.Caption = customer_credit.customer_info.customers_name
    txtLimit.Text = customer_credit.customer_info.credit_limit
End Sub

Private Sub txtCustomer_Change()
Dim rs As New ADODB.Recordset
Set rs = searchCustomersByName(txtCustomer)
Call loadCustomerRSToListView(lsvCustomer, rs)
End Sub
