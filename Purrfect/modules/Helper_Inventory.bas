Attribute VB_Name = "Helper_Inventory"
Public Function getBeginningBalance() As ADODB.Recordset
Dim sql As String
Dim i As New Inventory

    'If i.hasLastInventory Then
    '    sql = "SELECT li.`item_id`, li.`item_code`, li.`ending_balance` as beginning_balance,i.item_qty as ending_balance FROM `last_inventory` li LEFT JOIN items i ON li.item_id = i.item_id"
   ' Else
        sql = "SELECT i.`item_id`, i.`item_code`, i.item_qty as beginning_balance,i.item_qty as ending_balance FROM `items` i ORDER BY i.item_code"
   ' End If
    Set getBeginningBalance = db.execute(sql)
    
End Function

Public Sub loadTodaysInventoryToListView(lsv As ListView)
    Dim lst As ListItem
    Dim rs As New ADODB.Recordset
    
    Set rs = getBeginningBalance
    On Error Resume Next
    Do Until rs.EOF
        Set lst = lsv.ListItems.Add(, , rs.Fields("item_code").Value)
        'lst.SubItems(1) = rs.Fields("beginning_balance").Value
        lst.SubItems(2) = rs.Fields("ending_balance").Value
    rs.MoveNext
    Loop
End Sub
Sub loadInventoryDateSelection(cbo As ComboBox)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT `date` FROM `inventory` group by `date`"
    Set rs = db.execute(sql)
    cbo.Clear
    cbo.Text = "Today"
    Do Until rs.EOF
        cbo.AddItem rs.Fields(0).Value
    rs.MoveNext
    Loop
    
End Sub

Sub loadCategoryToListview(icategory As String, lsv As ListView)
    
   Dim lst As ListItem
    Dim rs As New ADODB.Recordset
    
    Set rs = getAllCategory(icategory)
    On Error Resume Next
    lsv.ListItems.Clear
    
    If rs.RecordCount > 0 Then
     Do Until rs.EOF
         
         Set lst = lsv.ListItems.Add(, , rs.Fields("item_code").Value)
         lst.SubItems(1) = rs.Fields("beginning_balance").Value
         lst.SubItems(2) = rs.Fields("ending_balance").Value
     rs.MoveNext
     Loop
     
    End If
       Set rs = Nothing
      
End Sub

Function getAllCategory(ByVal icat As String) As Recordset
    Dim qry As String
        
    If icat = "All" Then
'    qry = "SELECT lst.item_id,lst.item_code,lst.ending_balance " & _
'          "FROM last_inventory lst left join item_category icat on " & _
'          "lst.item_code = icat.item_code"
    qry = "SELECT lst.item_id,lst.item_code,lst.item_qty as ending_balance " & _
          "FROM items lst left join item_category icat on " & _
          "lst.item_code = icat.item_code order by lst.item_code"
    Else
        qry = "SELECT lst.item_id,lst.item_code,lst.item_qty as ending_balance " & _
          "FROM items lst left join item_category icat on " & _
          "lst.item_code = icat.item_code WHERE icat.category = '" & icat & "' order by lst.item_code"
    End If
    Set getAllCategory = db.execute(qry)
    
End Function

Sub load_to_category_combo(cbo As ComboBox)
    Dim qry As String
    Dim rs As New ADODB.Recordset
    
    qry = "SELECT distinct(category) FROM item_category order by category asc"
    
    Set rs = db.execute(qry)
    cbo.Clear
    cbo.AddItem "All"
    Do Until rs.EOF
    
        cbo.AddItem rs.Fields(0).Value
        rs.MoveNext
    Loop
    
End Sub
