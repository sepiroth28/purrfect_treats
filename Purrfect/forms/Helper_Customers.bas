Attribute VB_Name = "Helper_Customers"
Sub loadAllCustomersToListview(lsv As ListView)
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    Dim customer As New Customers
    lsv.ListItems.Clear
    Set Collection = getAllCustomersCollection
        For Each customer In Collection
            Set list = lsv.ListItems.Add(, , customer.customers_id)
            list.SubItems(1) = customer.customers_name
            list.SubItems(2) = customer.customers_add
            list.SubItems(3) = customer.customers_number
            list.SubItems(4) = customer.dealers_type
        Next
End Sub
Sub loadAllCustomersToListviewHidden(lsv As ListView)
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    Dim customer As New Customers
    lsv.ListItems.Clear
    Set Collection = getAllCustomersCollectionHidden
        For Each customer In Collection
            Set list = lsv.ListItems.Add(, , customer.customers_id)
            list.SubItems(1) = customer.customers_name
            list.SubItems(2) = customer.customers_add
            list.SubItems(3) = customer.customers_number
            list.SubItems(4) = customer.dealers_type
        Next
End Sub
Function getAllCustomersCollectionHidden() As CustomersCollection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim customers_col As New CustomersCollection
    Dim temp_customers As New Customers
    
    sql = "SELECT * FROM customers WHERE visible = 0 ORDER BY customers_name ASC"
    Set data = db.execute(sql)
    On Error Resume Next
    Do Until data.EOF
        With temp_customers
            .customers_id = data.Fields("customers_id").Value
            .customers_name = data.Fields("customers_name").Value
            .customers_add = data.Fields("customers_add").Value
            .customers_number = data.Fields("customers_number").Value
            .dealers_type = data.Fields("dealers_type").Value
        End With
         customers_col.Add temp_customers, data.Fields("customers_id").Value
         data.MoveNext
    Loop
   
    Set getAllCustomersCollectionHidden = customers_col
End Function
Function getAllCustomersCollection() As CustomersCollection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim customers_col As New CustomersCollection
    Dim temp_customers As New Customers
    
    sql = "SELECT * FROM customers WHERE visible = 1 ORDER BY customers_name ASC"
    Set data = db.execute(sql)
    On Error Resume Next
    Do Until data.EOF
        With temp_customers
            .customers_id = data.Fields("customers_id").Value
            .customers_name = data.Fields("customers_name").Value
            .customers_add = data.Fields("customers_add").Value
            .customers_number = data.Fields("customers_number").Value
            .dealers_type = data.Fields("dealers_type").Value
        End With
         customers_col.Add temp_customers, data.Fields("customers_id").Value
         data.MoveNext
    Loop
   
    Set getAllCustomersCollection = customers_col
End Function

Sub deleteCustomer(customer_id As Integer)
    
        db.execute "DELETE FROM customers WHERE customers_id = " & customer_id
   
End Sub
Function searchCustomersByName(customers_name As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "SELECT * FROM `customers` WHERE customers_name like '" & customers_name & "%' AND visible = 1"
            
    Set rs = db.execute(sql)
    Set searchCustomersByName = rs
End Function
Sub loadCustomerRSToListView(lsv As ListView, rs As ADODB.Recordset)

Dim list As ListItem
lsv.ListItems.Clear
    Do Until rs.EOF
        'customers_id, customers_name, customers_add, customers_number
        Set list = lsv.ListItems.Add(, , rs.Fields("customers_id").Value)
            list.SubItems(1) = rs.Fields("customers_name").Value
            list.SubItems(2) = rs.Fields("customers_add").Value
            list.SubItems(3) = rs.Fields("customers_number").Value
            list.SubItems(4) = rs.Fields("dealers_type").Value
        rs.MoveNext
    Loop
End Sub

