VERSION 5.00
Begin VB.MDIForm mdi_Inventory 
   BackColor       =   &H8000000C&
   Caption         =   "PURRFECT TREATS PETSHOP"
   ClientHeight    =   8355
   ClientLeft      =   5550
   ClientTop       =   750
   ClientWidth     =   13380
   Icon            =   "mdi_Inventory.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_manage_agent 
         Caption         =   "Manage Agent"
      End
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
      Begin VB.Menu mnu_manage_municipal 
         Caption         =   "Manage Municipalities"
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
      Begin VB.Menu mnu_view_aging_account 
         Caption         =   "View AGING ACCOUNT"
      End
   End
   Begin VB.Menu mnu_sub 
      Caption         =   "sub_menu"
      Visible         =   0   'False
      Begin VB.Menu mnu_delete_item 
         Caption         =   "Delete item"
      End
   End
   Begin VB.Menu mnu_admin 
      Caption         =   "Admin"
      Begin VB.Menu mnuCreditLimit 
         Caption         =   "Credit Limit"
      End
      Begin VB.Menu mnuViewStockIn 
         Caption         =   "view StockIn"
      End
      Begin VB.Menu mnu_sales_adjustment 
         Caption         =   "Sales Adjustment"
      End
      Begin VB.Menu mnu_admin_customer_visible 
         Caption         =   "Customer Visibility"
      End
   End
   Begin VB.Menu mnu_print_receipt 
      Caption         =   "Print Receipt"
   End
   Begin VB.Menu mnu_rebates 
      Caption         =   "Rebates"
      Begin VB.Menu mnu_rebate_item 
         Caption         =   "Rebate Items"
      End
      Begin VB.Menu mnu_rebate_price_table 
         Caption         =   "Rebate Price Table"
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

Private Sub mnu_admin_customer_visible_Click()
frmCustomerVisibility.Show 1
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

Private Sub mnu_manage_agent_Click()
frmAgentForm.Show
End Sub

Private Sub mnu_manage_manufacturers_Click()
frmManageManufacturers.Show 1
End Sub

Private Sub mnu_manage_municipal_Click()
frmMunicipalities.Show 1
End Sub

Private Sub mnu_print_receipt_Click()
frmReprint.Show 1
End Sub

Private Sub mnu_rebate_item_Click()
frmManageItemRebates.Show
End Sub

Private Sub mnu_rebate_price_table_Click()
frmRebatesPriceTable.Show
End Sub

Private Sub mnu_sales_adjustment_Click()
frmAdjustSaleTransaction.Show
End Sub

Private Sub mnu_view_aging_account_Click()
frmAgingAccounts.Show
End Sub

Private Sub mnu_view_sales_Click()
frmViewSales.Show
End Sub

Private Sub mnuCreditLimit_Click()
frmCreditLimit.Show
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

Private Sub mnuViewStockIn_Click()
    frmViewStockIn.Show
End Sub
