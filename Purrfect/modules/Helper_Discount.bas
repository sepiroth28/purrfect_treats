Attribute VB_Name = "Helper_Discount"
Sub loadAllDiscountToListview(lsv As ListView)
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    Dim discount As New Customers_Discount
    lsv.ListItems.Clear
    Set Collection = getAllDiscountCollection
        For Each discount In Collection
            Set list = lsv.ListItems.Add(, , discount.discount_id)
            list.SubItems(1) = discount.discount_code
            list.SubItems(2) = discount.discount_name
            list.SubItems(3) = discount.discount_amount
        Next
End Sub

Function getAllDiscountCollection() As Discount_Collection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim discount_col As New Discount_Collection
    Dim temp_discount As New Customers_Discount
    
    sql = "SELECT * FROM discount"
    Set data = db.execute(sql)
    
    Do Until data.EOF
        With temp_discount
            .discount_id = data.Fields("discount_id").Value
            .discount_code = data.Fields("discount_code").Value
            .discount_name = data.Fields("discount_name").Value
            .discount_amount = data.Fields("amount").Value
        End With
         discount_col.Add temp_discount, data.Fields("discount_id").Value
         data.MoveNext
    Loop
   
    Set getAllDiscountCollection = discount_col
End Function
Sub Delete_Discount(discount_id As Integer)
    
        db.execute "DELETE FROM discount WHERE discount_id = " & discount_id
   
End Sub
