Attribute VB_Name = "Helper_manufacturers"
'getting all manufacturers returns recordset format
Function getAllManufacturers(Optional sortBy As String = "") As adodb.Recordset
    Dim sql As String
    Dim data As New adodb.Recordset
    
    If sortBy <> "" Then
        sql = "SELECT * from manufacturers ORDER BY " & sortBy
    Else
        sql = "SELECT * from manufacturers"
    End If
    
        sql = "SELECT * from manufacturers"
        Set data = db.execute(sql)
        Set getAllManufacturers = data
End Function


'getting an item with specified item_code
Function getManufacturer(manufacturers_id As Integer) As manufacturers
    Dim sql As String
    Dim data As adodb.Recordset
    Dim temp_manufacturers As New manufacturers
    
    'manufacturers_id, manufacturers_name, manufacturers_add, manufacturers_number
    
        sql = "SELECT * from manufacturers WHERE manufacturers_id = " & manufacturers_id
  
        Set data = db.execute(sql)
        
        If data.RecordCount > 0 Then
            temp_manufacturers.load_manufacturers (data.Fields("manufacturers_id").Value)
        End If
    Set getManufacturer = temp_manufacturers
End Function

'deleting an item with specified itemCode
Function deleteManufacturers(manufacturers_id As Integer)
    Dim delete As String
    
    delete = "DELETE FROM manufacturers WHERE manufacturers_id = " & manufacturers_id
    db.execute (delete)
    
End Function

Function loadAllmanufacturersToListview(lsv As ListView) As ListView
Dim list As ListItem
Dim rs As New adodb.Recordset

'manufacturers_id, manufacturers_name, manufacturers_add, manufacturers_number
Set rs = getAllManufacturers
   lsv.ListItems.Clear
   Do Until rs.EOF
            Set list = lsv.ListItems.Add(, , rs.Fields("manufacturers_id").Value)
            list.SubItems(1) = rs.Fields("manufacturers_name").Value
            list.SubItems(2) = rs.Fields("manufacturers_add").Value
            list.SubItems(3) = rs.Fields("manufacturers_number").Value

   rs.MoveNext
   Loop
    
End Function

