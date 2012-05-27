Attribute VB_Name = "Helper_RebatePriceTable"
Sub loadRebatePriceTable(lsv As ListView)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim list As ListItem

sql = "SELECT * FROM rebate_price_table"

Set rs = db.execute(sql)
lsv.ListItems.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        Set list = lsv.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields(1).Value
        list.SubItems(2) = rs.Fields(2).Value
        list.SubItems(3) = rs.Fields(3).Value
        
    rs.MoveNext
    Loop
End If
End Sub

Sub renderRebateTableRates(lsv As ListView)
Dim list As ListItem
Dim qty_bought As Double
Dim rebate_price As Double
Dim total_amount As Double

rebate_grand_total = 0
rebate_grand_total_qty = 0

For Each list In lsv.ListItems
    qty_bought = Val(list.SubItems(3))
    rebate_price = getRebateRate(qty_bought)
    total_amount = qty_bought * rebate_price
    list.SubItems(5) = rebate_price
    list.SubItems(6) = total_amount
    
    rebate_grand_total = rebate_grand_total + total_amount
    rebate_grand_total_qty = rebate_grand_total_qty + qty_bought
Next

End Sub

Function getRebateRate(qty As Double) As Double
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT * FROM rebate_price_table WHERE qty_from <= " & qty & " AND qty_to >= " & qty

Set rs = db.execute(sql)

If rs.RecordCount > 0 Then
    getRebateRate = Val(rs.Fields("applied_price").Value)
Else
    getRebateRate = 0
End If


End Function
