Attribute VB_Name = "Helper_Credit"
Function isInLimit(customer_id As Integer) As Boolean
    Dim sql As String
    Dim customer_limit As New Customers
    Dim debt As Double
    
    Call customer_limit.load_customers(customer_id)
    debt = getTotalDebtOfThisCustomer(customer_id) + amount_to_be_debt
    If debt >= customer_limit.credit_limit And customer_limit.credit_limit > 0 Then
        isInLimit = True
    Else
        isInLimit = False
    End If
    frmMenu.lblCreditLimit.Caption = "0.00"
    frmMenu.lblCreditLimit.Caption = FormatNumber(debt, 2)
End Function

Function getTotalDebtOfThisCustomer(customer_id) As Double
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT SUM(net_total) as total FROM `stock_out_transaction` sot " & _
      "  inner join account_receivable acr " & _
      "  ON acr.sales_order_no = sot.sales_order_no " & _
      "  where sot.responsible_customer = " & customer_id & " AND acr.remarks = 'unsettled' " & _
      "  GROUP BY responsible_customer"
Set rs = db.execute(sql)
If rs.RecordCount > 0 Then
    getTotalDebtOfThisCustomer = rs.Fields(0).Value
End If
End Function
