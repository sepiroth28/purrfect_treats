Attribute VB_Name = "Helper_UserPreviliges"
Sub grantAdminPreviligesToActiveUser()
   ' Call activeUser.previliges.clearPreviliges
    'activeUser.previliges.grantAll
End Sub
Sub grantUserPreviliges(active_username As String)
   Dim rs As New ADODB.Recordset
   Dim sql As String
   sql = "SELECT up.*,p.previleges as previleges_description FROM user_previleges up INNER JOIN previleges p ON up.previleges = p.id WHERE up.username = '" & active_username & "'"
   Set rs = db.execute(sql)
   If rs.RecordCount > 0 Then
    Do Until rs.EOF
       Call setPrevileges(rs.Fields("previleges_description").Value, rs.Fields("status").Value)
    rs.MoveNext
    Loop
   End If
End Sub
Sub setPrevileges(previleges, action As Boolean)
    Select Case previleges
        Case "payment":
            activeUser.previliges.canProcessPayment = action
        Case "customer_add":
            activeUser.previliges.canAddCustomer = action
        Case "stockin":
            activeUser.previliges.canStockIn = action
        Case "inventory":
            activeUser.previliges.canInventory = action
        Case "stockout":
            activeUser.previliges.canStockOut = action
        Case "technician":
            activeUser.previliges.canManagetechnician = action
        Case "manage_manufacturer":
            activeUser.previliges.canManageManufacturer = action
        Case "sales_order_responsible":
            activeUser.previliges.canManageSORep = action
        Case "manage_useraccount":
            activeUser.previliges.canManageUserAccount = action
        Case "view_sales":
            activeUser.previliges.canViewSales = action
        Case "print_sales_details":
            activeUser.previliges.canPrintSalesDetails = action
        Case "credit_limit":
            activeUser.previliges.canManageCreditLimit = action
        Case "view_stock_in":
            activeUser.previliges.canViewStockIn = action
        Case "sales_adjustment":
            activeUser.previliges.canSalesAdjustment = action
        Case "print_receipt":
            activeUser.previliges.canPrintReceipt = action
        Case "delete_customer":
            activeUser.previliges.canDeleteCustomer = action
        Case "manage_item":
            activeUser.previliges.canManageItem = action
        Case "delete_item":
            activeUser.previliges.canDeleteItem = action
        Case "customer_visibility":
            activeUser.previliges.canManageCustomerVisibility = action
        Case "can_accept_remit_payments":
            activeUser.previliges.can_accept_remit_payments = action
        Case "can_issue_rebate":
            activeUser.previliges.can_issue_rebate = action
    End Select
End Sub




Sub renderButtonBasedOnUserPreviliges()
    With activeUser.previliges
       frmMenu.cmdManageItem.Enabled = .canManageItem
       'toolbar menu
       frmMenu.Toolbar1.Buttons(1).Visible = .canManageItem
       
       mdi_Inventory.mnu_manage_manufacturers.Enabled = .canManageManufacturer
       mdi_Inventory.mnu_view_sales.Enabled = .canViewSales
       
       mdi_Inventory.mnuSOResponsible.Enabled = .canManageSORep
       mdi_Inventory.mnuUserAccount.Enabled = .canManageUserAccount
       mdi_Inventory.mnu_print_receipt = .canPrintReceipt
       mdi_Inventory.mnuCreditLimit = .canManageCreditLimit
       mdi_Inventory.mnuViewStockIn = .canViewStockIn
       mdi_Inventory.mnu_sales_adjustment = .canSalesAdjustment
       mdi_Inventory.mnu_view_sales_details = .canPrintSalesDetails
       mdi_Inventory.mnu_admin_customer_visible.Enabled = .canManageCustomerVisibility
    
       frmMenu.cmdViewSales.Visible = .canViewSales
       
       frmMenu.cmdCustomer.Enabled = .canAddCustomer
       'toolbar menu customer
       frmMenu.Toolbar1.Buttons(2).Visible = .canAddCustomer
       
       frmMenu.cmdAgent.Enabled = .canManagetechnician
       frmMenu.cmdStockIn.Enabled = .canStockIn
        'toolbar menu stock in
       frmMenu.Toolbar1.Buttons(3).Visible = .canStockIn
       
       
       frmMenu.cmdView.Enabled = True
       
       frmMenu.cmdPayment.Enabled = .canProcessPayment
        'toolbar menu payment
       frmMenu.Toolbar1.Buttons(4).Visible = .canProcessPayment
       
       
       frmMenu.cmdInventory.Enabled = .canInventory
        'toolbar menu inventory
       frmMenu.Toolbar1.Buttons(5).Visible = .canInventory
       
       frmMenu.cmdStockout.Enabled = .canStockOut
       frmMenu.cmdNewTransaction.Enabled = True
       frmMenu.cmdNewAccountReceivable.Enabled = True
       
       
    End With
End Sub



