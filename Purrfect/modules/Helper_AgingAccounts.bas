Attribute VB_Name = "Helper_AgingAccounts"
Sub loadAgingAccounts(lsv As ListView, diff_month As Integer)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim list As ListItem
'sales_order_no, responsible_customer, responsible_agent, discount, grand_total, net_total, tendered_amount, change, delivery_date, prepared_by, DATE_FORMAT(delivery_date,'%M')

sql = "SELECT " & _
        " sot.responsible_customer ," & _
        " c.customers_name," & _
        " count(sot.sales_order_no) as remarks " & _
        " FROM `stock_out_transaction` sot " & _
        " INNER JOIN account_receivable acr ON sot.sales_order_no = acr.sales_order_no " & _
        " INNER JOIN customers c ON sot.responsible_customer = c.customers_id " & _
        " where DATE_FORMAT(delivery_date,'%Y-%m-%d') " & _
        " < DATE_SUB(curdate(), INTERVAL " & diff_month & " MONTH) " & _
        " AND acr.remarks = 'unsettled' " & _
        " group by sot.responsible_customer "
       
Set rs = db.execute(sql)

lsv.ListItems.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        Set list = lsv.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields("customers_name").Value
        list.SubItems(2) = rs.Fields("remarks").Value
    rs.MoveNext
    Loop
End If

Set rs = Nothing
End Sub
