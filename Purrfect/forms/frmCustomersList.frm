VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomersList 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer List"
   ClientHeight    =   6405
   ClientLeft      =   30
   ClientTop       =   660
   ClientWidth     =   5910
   Icon            =   "frmCustomersList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4260
      TabIndex        =   2
      Top             =   5700
      Width           =   1572
   End
   Begin VB.TextBox txtSearchCustomer 
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
      TabIndex        =   1
      Top             =   5700
      Width           =   4092
   End
   Begin MSComctlLib.ListView lsvCustomerList 
      Height          =   5532
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5772
      _ExtentX        =   10186
      _ExtentY        =   9763
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   8819
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
   Begin VB.Menu mnu_file 
      Caption         =   "file"
      Begin VB.Menu mnu_sohistory 
         Caption         =   "Show SO History"
      End
   End
End
Attribute VB_Name = "frmCustomersList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSelect_Click()
 If activeSales.payment_type = PAYMENT_ACCOUNT_RECEIVABLE Then
    If isInLimit(Val(lsvCustomerList.SelectedItem.Text)) Then
        MsgBox "Customers reach his/her credit limit...Please refer to the SO history of this customer", vbInformation, "Credit Limit reached"
    Else
        Call activeSales.sold_to.load_customers(Val(lsvCustomerList.SelectedItem.Text))
        frmMenu.txtCustomers.Text = activeSales.sold_to.customers_name
        frmMenu.lblAgent.Caption = activeSales.sold_to.mvaragent.agent_name
        frmMenu.lblDealerType.Caption = activeSales.sold_to.dealers_type
        'checkProcessButton r
        Unload Me
    End If
ElseIf activeSales.payment_type = PAYMENT_COD Then
    Call activeSales.sold_to.load_customers(Val(lsvCustomerList.SelectedItem.Text))
        frmMenu.txtCustomers.Text = activeSales.sold_to.customers_name
        frmMenu.lblAgent.Caption = activeSales.sold_to.mvaragent.agent_name
        frmMenu.lblDealerType.Caption = activeSales.sold_to.dealers_type
        'checkProcessButton
        Unload Me
End If
 
' If activeSales.payment_type = PAYMENT_ACCOUNT_RECEIVABLE Then
'    If isInLimit(Val(lsvCustomerList.SelectedItem.Text)) Then
'        MsgBox "Customers reach his/her credit limit...", vbInformation, "Credit Limit reached"
'    Else
'        Call activeSales.sold_to.load_customers(Val(lsvCustomerList.SelectedItem.Text))
'        frmMenu.txtCustomers.Text = activeSales.sold_to.customers_name
'        frmMenu.lblAgent.Caption = activeSales.sold_to.mvaragent.agent_name
'        frmMenu.lblDealerType.Caption = activeSales.sold_to.dealers_type
'        'checkProcessButton r
'        Unload Me
'    End If
'ElseIf activeSales.payment_type = PAYMENT_COD Then
'    Call activeSales.sold_to.load_customers(Val(lsvCustomerList.SelectedItem.Text))
'        frmMenu.txtCustomers.Text = activeSales.sold_to.customers_name
'        frmMenu.lblAgent.Caption = activeSales.sold_to.mvaragent.agent_name
'        frmMenu.lblDealerType.Caption = activeSales.sold_to.dealers_type
'        'checkProcessButton
'        Unload Me
'End If
End Sub

Private Sub Form_Load()
Call loadAllCustomersToListview(lsvCustomerList)
End Sub

Private Sub mnu_sohistory_Click()
customer_id_for_list_of_account_receivable = Val(lsvCustomerList.SelectedItem.Text)
frmCustomerAccountReceivable.Show 1
End Sub

Private Sub txtSearchCustomer_Change()
Dim rs As New ADODB.Recordset
Set rs = searchCustomersByName(txtSearchCustomer)
Call loadCustomerRSToListView(lsvCustomerList, rs)
End Sub
