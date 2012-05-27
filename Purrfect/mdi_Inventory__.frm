VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_Inventory 
   BackColor       =   &H8000000C&
   Caption         =   "Nutrimart Enterpries Inventory & Cashering System"
   ClientHeight    =   8352
   ClientLeft      =   5556
   ClientTop       =   756
   ClientWidth     =   13380
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbNutrimart 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   868
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "UserName"
            TextSave        =   "UserName"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_manage_manufacturers 
         Caption         =   "Manage Manufacturers"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuManageDiscount 
         Caption         =   "Manage Discount"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepUserAccount 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu_setting 
      Caption         =   "Settings"
      Begin VB.Menu mnuSOResponsible 
         Caption         =   "Sales Order Responsible"
      End
      Begin VB.Menu mnuSepSOResponsible 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserAccount 
         Caption         =   "Manage User &Account"
      End
   End
   Begin VB.Menu mnu_sales 
      Caption         =   "Sales"
      Begin VB.Menu mnu_view_sales 
         Caption         =   "View Sales"
      End
      Begin VB.Menu mnu_view_sales_details 
         Caption         =   "Print Sales details"
         Begin VB.Menu mnu_detail_sales_all 
            Caption         =   "All Sales"
         End
         Begin VB.Menu mnu_detail_sales_cod 
            Caption         =   "COD Sales"
         End
         Begin VB.Menu mnu_detail_sales_acr 
            Caption         =   "Account Receivable Sales"
         End
      End
   End
   Begin VB.Menu mnu_sub 
      Caption         =   "sub_menu"
      Visible         =   0   'False
      Begin VB.Menu mnu_delete_item 
         Caption         =   "Delete item"
      End
   End
End
Attribute VB_Name = "mdi_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Call prepareNewTransaction
    frmMenu.Show
   ' Form1.Show
End Sub

Private Sub mnu_delete_item_Click()
Dim items As New cart_items
        Dim x As Integer
        x = 1
        For Each items In activeSales.items_sold
            If items.Item.item_name = frmMenu.lsvItemsInCart.SelectedItem.SubItems(1) Then
                Exit For
            End If
        x = x + 1
        Next
        activeSales.items_sold.Remove (x)
        Call loadActiveCartItems(frmMenu.lsvItemsInCart)
        Call updateTotalAmount
End Sub

Private Sub mnu_detail_sales_acr_Click()
Dim rs As New ADODB.Recordset
    Dim sql  As String
    sql = "SELECT acr.sales_order_no," & _
          " c.customers_name as `sold_to`," & _
          " a.name as `agent`," & _
          " i.item_code," & _
          " so.qty_out," & _
          " IF(c.dealers_type = 'dealer',i.dealers_price,i.item_price) as `price`," & _
          " so.amount," & _
          " so.discount," & _
          " so.tracking_price," & _
          " (so.amount-IF(so.discount IS NULL,0,so.discount))+IF(so.tracking_price IS NULL,0,so.tracking_price) as `net_total`," & _
          " sot.`delivery_date` as `date`" & _
          " FROM account_receivable acr " & _
          "  LEFT JOIN `stock_out_transaction_stock_out_items` sot_sot " & _
          "  ON acr.`sales_order_no` = sot_sot.`sales_order_no` " & _
          "  LEFT JOIN stock_out_transaction sot " & _
          "  ON sot_sot.sales_order_no = sot.sales_order_no " & _
        " LEFT JOIN customers c " & _
        " ON c.customers_id = sot.responsible_customer " & _
        " LEFT JOIN agent a " & _
        " ON a.agent_id = sot.responsible_agent " & _
        " LEFT JOIN stock_out so " & _
        " ON sot_sot.stockout_id = so.stockout_id " & _
        " LEFT JOIN items i " & _
        " ON so.item_id = i.item_id " & _
        " WHERE DATE_FORMAT(acr.`date`,'%Y-%m-%d') = CURDATE() "
    Set rs = db.execute(sql)
    Set dtaDetailedSalesReport.DataSource = rs
    dtaDetailedSalesReport.Sections(1).Controls("lblDate").Caption = FormatDateTime(Date, vbLongDate)
    
    dtaDetailedSalesReport.Show 1
End Sub

Private Sub mnu_detail_sales_all_Click()
    Dim rs As New ADODB.Recordset
    Dim sql  As String
    sql = "SELECT " & _
          " sot_sot.sales_order_no," & _
          " c.customers_name as `sold_to`," & _
          " a.name as `agent`," & _
          " i.item_code," & _
          " so.qty_out," & _
          " IF(c.dealers_type = 'dealer',i.dealers_price,i.item_price) as `price`," & _
          " so.amount," & _
          " so.discount," & _
          " so.tracking_price," & _
          " (so.amount-IF(so.discount IS NULL,0,so.discount))+IF(so.tracking_price IS NULL,0,so.tracking_price) as `net_total`," & _
          " sot.`delivery_date` as `date`" & _
        " FROM `stock_out_transaction_stock_out_items` sot_sot " & _
        " LEFT JOIN stock_out_transaction sot " & _
        " ON sot_sot.sales_order_no = sot.sales_order_no " & _
        " LEFT JOIN customers c " & _
        " ON c.customers_id = sot.responsible_customer " & _
        " LEFT JOIN agent a " & _
        " ON a.agent_id = sot.responsible_agent " & _
        " LEFT JOIN stock_out so " & _
        " ON sot_sot.stockout_id = so.stockout_id " & _
        " LEFT JOIN items i " & _
        " ON so.item_id = i.item_id " & _
        " WHERE DATE_FORMAT(sot.`delivery_date`,'%Y-%m-%d') = CURDATE() "
    Set rs = db.execute(sql)
    Set dtaDetailedSalesReport.DataSource = rs
    dtaDetailedSalesReport.Sections(1).Controls("lblDate").Caption = FormatDateTime(Date, vbLongDate)
    
    dtaDetailedSalesReport.Show 1
    
End Sub

Private Sub mnu_detail_sales_cod_Click()
Dim rs As New ADODB.Recordset
    Dim sql  As String
    sql = "SELECT cod.sales_order_no," & _
          " c.customers_name as `sold_to`," & _
          " a.name as `agent`," & _
          " i.item_code," & _
          " so.qty_out," & _
          " IF(c.dealers_type = 'dealer',i.dealers_price,i.item_price) as `price`," & _
          " so.amount," & _
          " so.discount," & _
          " so.tracking_price," & _
          " (so.amount-IF(so.discount IS NULL,0,so.discount))+IF(so.tracking_price IS NULL,0,so.tracking_price) as `net_total`," & _
          " sot.`delivery_date` as `date`" & _
          " FROM cod cod " & _
          "  LEFT JOIN `stock_out_transaction_stock_out_items` sot_sot " & _
          "  ON cod.`sales_order_no` = sot_sot.`sales_order_no` " & _
          "  LEFT JOIN stock_out_transaction sot " & _
          "  ON sot_sot.sales_order_no = sot.sales_order_no " & _
        " LEFT JOIN customers c " & _
        " ON c.customers_id = sot.responsible_customer " & _
        " LEFT JOIN agent a " & _
        " ON a.agent_id = sot.responsible_agent " & _
        " LEFT JOIN stock_out so " & _
        " ON sot_sot.stockout_id = so.stockout_id " & _
        " LEFT JOIN items i " & _
        " ON so.item_id = i.item_id " & _
        " WHERE DATE_FORMAT(sot.`delivery_date`,'%Y-%m-%d') = CURDATE() "
    Set rs = db.execute(sql)
    Set dtaDetailedSalesReport.DataSource = rs
    dtaDetailedSalesReport.Sections(1).Controls("lblDate").Caption = FormatDateTime(Date, vbLongDate)
    
    dtaDetailedSalesReport.Show 1
End Sub

Private Sub mnu_exit_Click()
    End
End Sub

Private Sub mnu_manage_manufacturers_Click()
frmManageManufacturers.Show 1
End Sub

Private Sub mnu_view_sales_Click()
frmViewSales.Show 1
End Sub

Private Sub mnuManageDiscount_Click()
    frmManageDiscount.Show 1
End Sub

Private Sub mnuMunicipality_Click()
    editmode = False
    frmMunicipality.txtMunicipality.Enabled = False
    frmMunicipality.Show
End Sub

Private Sub mnuSOResponsible_Click()
    frmSalesOrder_Responsible.Show
End Sub

Private Sub mnuUserAccount_Click()
    frmManageUserAccount.Show 1
End Sub
