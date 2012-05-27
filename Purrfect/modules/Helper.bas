Attribute VB_Name = "Helper_items"
'getting all items returns recordset format
Function getAllItems(Optional sortBy As String = "") As ADODB.Recordset
    Dim sql As String
    Dim data As New ADODB.Recordset
    
    If sortBy <> "" Then
        sql = "SELECT * from items ORDER BY " & sortBy
    Else
        sql = "SELECT * from items"
    End If
    
        sql = "SELECT * from items"
        Set data = db.execute(sql)
        Set getAllItems = data
End Function
'getting all items returns ItemCollection
Function getSearchItemsCollection(item_code As String) As ItemCollection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim item_coll As New ItemCollection
    Dim temp_item As New items
    Dim man As New manufacturers
    Dim sort_by As String
    'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    
    sortBy = "item_code"
    
    If sortBy <> "" Then
        sql = "SELECT * from items i INNER JOIN items_description id on i.item_code = id.item_code where i.item_code like '" & item_code & "%' ORDER BY i." & sortBy
    Else
        sql = "SELECT * from items i INNER JOIN items_description id on i.item_code = id.item_code"
    End If
        Set data = db.execute(sql)
        On Error Resume Next
        Do Until data.EOF
            With temp_item
                .item_id = data.Fields("item_id").Value
                .item_code = data.Fields("item_code").Value
                .item_name = data.Fields("item_name").Value
                .item_description = data.Fields("item_description").Value
                .item_qty = data.Fields("item_qty").Value
                .item_price = data.Fields("item_price").Value
                .dealers_price = data.Fields("dealers_price").Value
                .date_added = data.Fields("date_added").Value
                .date_modified = data.Fields("date_modified").Value
                .manufacturers_id = data.Fields("manufacturers_id").Value
                .reorder_point = data.Fields("reorder_point").Value
                .unit_of_measure = data.Fields("unit_of_measure").Value
                'add here additional field from items_description
                 
                'load records manufacturer of this item
                .manufacturer.load_manufacturers (.manufacturers_id)
            End With
            
            item_coll.Add temp_item, data.Fields("item_id").Value
        data.MoveNext
        Loop
    Set getSearchItemsCollection = item_coll
End Function
'getting all items returns ItemCollection
Function getAllItemsCollection(Optional sortBy As String = "") As ItemCollection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim item_coll As New ItemCollection
    Dim temp_item As New items
    Dim man As New manufacturers
    'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    
    If sortBy <> "" Then
        sql = "SELECT * from items i INNER JOIN items_description id on i.item_code = id.item_code ORDER BY i." & sortBy
    Else
        sql = "SELECT * from items i INNER JOIN items_description id on i.item_code = id.item_code"
    End If
        Set data = db.execute(sql)
        On Error Resume Next
        Do Until data.EOF
            With temp_item
                .item_id = data.Fields("item_id").Value
                .item_code = data.Fields("item_code").Value
                .item_name = data.Fields("item_name").Value
                .item_description = data.Fields("item_description").Value
                .item_qty = data.Fields("item_qty").Value
                .item_price = data.Fields("item_price").Value
                .dealers_price = data.Fields("dealers_price").Value
                .date_added = data.Fields("date_added").Value
                .date_modified = data.Fields("date_modified").Value
                .manufacturers_id = data.Fields("manufacturers_id").Value
                .reorder_point = data.Fields("reorder_point").Value
                .unit_of_measure = data.Fields("unit_of_measure").Value
                
                .include_in_rebate = data.Fields("include_in_rebate").Value
                'add here additional field from items_description
                 
                'load records manufacturer of this item
                .manufacturer.load_manufacturers (.manufacturers_id)
            End With
            
            item_coll.Add temp_item, data.Fields("item_id").Value
        data.MoveNext
        Loop
    Set getAllItemsCollection = item_coll
End Function

'getting an item with specified item_code
Function getItem(itemCode As String) As items
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim temp_item As New items
    
    'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    
        sql = "SELECT * from items WHERE item_code = '" & itemCode & "'"
  
        Set data = db.execute(sql)
        
        If data.RecordCount > 0 Then
            temp_item.load_item (data.Fields("items_id").Value)
        End If
    Set getItems = temp_item
End Function

'deleting an item with specified itemCode
Function deleteItem(itemCode As String)
    Dim delete As String
    
    delete = "DELETE FROM items WHERE item_code = '" & itemCode & "'"
    db.execute (delete)
    
    delete = "DELETE FROM items_description WHERE item_code = '" & itemCode & "'"
    db.execute (delete)
    
End Function

Function loadAllItemsToListview(lsv As ListView, sort_by As String) As ListView
Dim list As ListItem
Dim rs As New ADODB.Recordset
Dim Item As New items
'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
lsv.ListItems.Clear
Set Collection = getAllItemsCollection(sort_by)

    For Each Item In Collection
            Set list = lsv.ListItems.Add(, , Item.item_id)
            list.SubItems(1) = Item.item_code
            list.SubItems(2) = Item.item_name
            list.SubItems(3) = Item.item_description
            list.SubItems(4) = Item.item_qty
            list.SubItems(5) = Item.item_price
            list.SubItems(6) = Item.dealers_price
            list.SubItems(7) = Item.unit_of_measure
            
            If Item.manufacturers_id > 0 Then
                Item.manufacturer.load_manufacturers (Item.manufacturers_id)
                list.SubItems(8) = Item.manufacturer.manufacturers_name
            Else
                list.SubItems(8) = ""
            End If
'            list.SubItems(6) = item.item_status
    Next
    
End Function
Function loadAllItemsToListviewForRebates(lsv As ListView, sort_by As String) As ListView
Dim list As ListItem
Dim rs As New ADODB.Recordset
Dim Item As New items
Dim sql As String

'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
lsv.ListItems.Clear
Set Collection = getAllItemsCollection(sort_by)

    For Each Item In Collection
            Set list = lsv.ListItems.Add(, , Item.item_id)
            
            list.Checked = Item.include_in_rebate
            
            list.SubItems(1) = Item.item_code
            list.SubItems(2) = Item.item_name
            list.SubItems(3) = Item.item_description
            list.SubItems(4) = Item.item_qty
            list.SubItems(5) = Item.item_price
            list.SubItems(6) = Item.dealers_price
            list.SubItems(7) = Item.unit_of_measure
            
            If Item.manufacturers_id > 0 Then
                Item.manufacturer.load_manufacturers (Item.manufacturers_id)
                list.SubItems(8) = Item.manufacturer.manufacturers_name
            Else
                list.SubItems(8) = ""
            End If
'            list.SubItems(6) = item.item_status
    Next
    
End Function
Function loadSearchItemsToListview(lsv As ListView, item_code As String) As ListView
Dim list As ListItem
Dim rs As New ADODB.Recordset
Dim Item As New items
'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
lsv.ListItems.Clear
Set Collection = getSearchItemsCollection(item_code)

    For Each Item In Collection
            Set list = lsv.ListItems.Add(, , Item.item_id)
            list.SubItems(1) = Item.item_code
            list.SubItems(2) = Item.item_name
            list.SubItems(3) = Item.item_description
            list.SubItems(4) = Item.item_qty
            list.SubItems(5) = Item.item_price
            list.SubItems(6) = Item.dealers_price
            list.SubItems(7) = Item.unit_of_measure
            
            If Item.manufacturers_id > 0 Then
                Item.manufacturer.load_manufacturers (Item.manufacturers_id)
                list.SubItems(8) = Item.manufacturer.manufacturers_name
            Else
                list.SubItems(8) = ""
            End If
'            list.SubItems(6) = item.item_status
    Next
    
End Function

Function searchItemsByItemCode(itemCode As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "SELECT * from items i INNER JOIN items_description id ON i.item_code = id.item_code " & _
            " WHERE i.item_code like '" & itemCode & "%'"
    Set rs = db.execute(sql)
    Set searchItemsByItemCode = rs
End Function
Sub loadItemRSToListCiew(lsv As ListView, rs As ADODB.Recordset)

Dim list As ListItem
lsv.ListItems.Clear
    Do Until rs.EOF
    
        Set list = lsv.ListItems.Add(, , rs.Fields("item_id").Value)
            list.SubItems(1) = rs.Fields("item_code").Value
            list.SubItems(2) = rs.Fields("item_name").Value
            list.SubItems(3) = rs.Fields("item_description").Value
            list.SubItems(4) = rs.Fields("item_qty").Value
            list.SubItems(5) = rs.Fields("item_price").Value
            list.SubItems(6) = rs.Fields("unit_of_measure").Value
            list.SubItems(7) = rs.Fields("manufacturers_id").Value
           
    rs.MoveNext
    Loop
End Sub
Sub addThisItemToLastInventory(item_id)

End Sub
Function isInLastInventory(item_id) As Boolean
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * from last_inventory WHERE item_id = " & item_id
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        isInLastInventory = True
    Else
        isInLastInventory = False
    End If
End Function

Sub loadItemsByCategory(icat As String, lsv As ListView)
    Dim rs As New ADODB.Recordset
    Dim list As ListItem
    Dim sql As String
    Dim where As String
    
    If icat <> "All" Then
        where = " where ic.category = '" & icat & "'"
    Else
        where = ""
    End If
    
    sql = "SELECT i.item_id,i.item_code,id.item_name " & _
            " FROM `item_category` ic " & _
            " inner join items i on ic.item_code = i.item_code " & _
            " inner join items_description id on i.item_code = id.item_code " & where

    
    Set rs = db.execute(sql)
    lsv.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
            Set list = lsv.ListItems.Add(, , rs.Fields("item_id").Value)
            list.SubItems(1) = rs.Fields("item_code").Value
            list.SubItems(2) = rs.Fields("item_name").Value
        rs.MoveNext
        Loop
    End If
End Sub

Sub updateItemsRebate(item_id As Integer, is_include As Boolean)
    Dim insert As String
    insert = "UPDATE items SET include_in_rebate = " & is_include & " WHERE item_id = " & item_id
    db.execute insert
End Sub
