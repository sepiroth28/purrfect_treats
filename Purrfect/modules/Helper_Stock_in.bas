Attribute VB_Name = "Helper_Stock_in"
Function getStockInRecordsByDate(as_of As String) As ADODB.Recordset
    Dim sql As String
    sql = "SELECT " & _
          "  s.`stock_in_transaction_id`," & _
          "  s.`reference_no`," & _
          "  s.`stocked_in_to`, " & _
          "  m.`manufacturers_name`," & _
          "  s.`remarks`," & _
          "  s.`total_number_of_items` " & _
          "  FROM stock_in_transaction s " & _
          "  LEFT JOIN manufacturers m " & _
          "  ON s.`from_supplier` = m.`manufacturers_id` " & _
          " WHERE stock_in_date = '" & as_of & "'"
    Set getStockInRecordsByDate = db.execute(sql)
End Function
Sub loadStockInListByDate(as_of As String, lsv As ListView)
    Dim rs As New ADODB.Recordset
    Dim list As ListItem
    Set rs = getStockInRecordsByDate(as_of)
    
    lsv.ListItems.Clear
    If rs.RecordCount > 0 Then
    On Error Resume Next
      Do Until rs.EOF
        Set list = lsv.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields(1).Value
        list.SubItems(2) = rs.Fields(2).Value
        list.SubItems(3) = rs.Fields(3).Value
        list.SubItems(4) = rs.Fields(4).Value
         
      rs.MoveNext
      Loop
    End If
End Sub

Sub loadAlStockInList(lsv As ListView)
    Dim rs As New ADODB.Recordset
    Dim list As ListItem
    Set rs = getAllStockInRecords()
    
    lsv.ListItems.Clear
    If rs.RecordCount > 0 Then
    On Error Resume Next
      Do Until rs.EOF
        Set list = lsv.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields(1).Value
        list.SubItems(2) = rs.Fields(2).Value
        list.SubItems(3) = rs.Fields(3).Value
        list.SubItems(4) = rs.Fields(4).Value
        
      rs.MoveNext
      Loop
    End If
End Sub

Function getAllStockInRecords() As ADODB.Recordset
    Dim sql As String
    sql = "SELECT " & _
          "  s.`stock_in_transaction_id`," & _
          "  s.`reference_no`," & _
          "  s.`stocked_in_to`, " & _
          "  m.`manufacturers_name`," & _
          "  s.`remarks`," & _
          "  s.`total_number_of_items` " & _
          "  FROM stock_in_transaction s " & _
          "  LEFT JOIN manufacturers m " & _
          "  ON s.`from_supplier` = m.`manufacturers_id` "
    Set getAllStockInRecords = db.execute(sql)
End Function
Sub loadStockInItemsToListView(stock_in_no As Integer, lsv As ListView)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim list As ListItem
    sql = "SELECT i.item_code,si.qty_in FROM stock_in_transaction_to_stock_in_items s " & _
          "  LEFT JOIN stock_in si " & _
          "  ON s.stock_id = si.stockin_id " & _
          "  LEFT JOIN items i " & _
          "  ON si.item_id = i.item_id " & _
          "  Where s.stock_in_transaction_id = " & stock_in_no
    Set rs = db.execute(sql)
    
    lsv.ListItems.Clear
    If rs.RecordCount > 0 Then
    On Error Resume Next
    x = 1
       Do Until rs.EOF
            Set list = lsv.ListItems.Add(, , x)
            list.SubItems(1) = rs.Fields(0).Value
            list.SubItems(2) = rs.Fields(1).Value
       x = x + 1
       rs.MoveNext
       Loop
    End If
End Sub
