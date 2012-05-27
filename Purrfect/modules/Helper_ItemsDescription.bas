Attribute VB_Name = "Helper_ItemsDescription"
'======
Sub loadAllItemsDescriptionToListview(lsv As ListView)
Dim list As ListItem
Dim rs As New ADODB.Recordset
Dim Item As New itemdescription
'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point

Set collection = getAllItemsDescriptionCollection

    For Each Item In collection
            Set list = lsv.ListItems.Add(, , Item.item_code)
            list.SubItems(1) = Item.item_name
            list.SubItems(2) = Item.item_description
            list.SubItems(3) = Item.unit_of_measure
    Next
End Sub

'==========================
Function getAllItemsDescriptionCollection(Optional sortBy As String = "") As ItemDescriptionCollection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim item_coll As New ItemDescriptionCollection
    Dim temp_item As New itemdescription
    'items_id, item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    
    If sortBy <> "" Then
        sql = "SELECT * from items_description ORDER BY " & sortBy
    Else
        sql = "SELECT item_code,item_name,item_description,unit_of_measure from items_description"
    End If
        Set data = db.execute(sql)
        
        Do Until data.EOF
            With temp_item
               
                .item_code = data.Fields("item_code").Value
                .item_name = data.Fields("item_name").Value
                .item_description = data.Fields("item_description").Value
                .unit_of_measure = data.Fields("unit_of_measure").Value
                                
            End With
            
            item_coll.Add temp_item, data.Fields("item_code").Value
        data.MoveNext
        Loop
    Set getAllItemsDescriptionCollection = item_coll
End Function
