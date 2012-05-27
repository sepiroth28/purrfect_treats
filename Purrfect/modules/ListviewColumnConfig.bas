Attribute VB_Name = "ListviewColumnConfig"
'Set here the column settings for all listview used in program
Sub setViewStockInListview(lsv As ListView)
    Dim stockIn_column As New Collection
    'stock_in_transaction_id, reference_no, stocked_in_to, from_supplier, remarks, stock_in_date, total_number_of_items, total_qty, prepared_by, approved_by, received_by
    With stockIn_column
        .Add "No"
        .Add "Reference No"
        .Add "Stock in to"
        .Add "from Supplier"
        .Add "remarks"
        .Add "Total Number of Items"
       
    End With
     setListviewColumn lsv, stockIn_column
    Set payment_column = Nothing
End Sub
Sub setPaymentReceivedListview(lsv As ListView)
    Dim payment_column As New Collection
    'pr.`id`, pr.`sales_order_no`,c.customers_name, pr.`amount`, pr.`balance`, pr.`payment_date`, pr.`remarks`,pr.received_by
    With payment_column
        .Add "id"
        .Add "Sales Order No"
        .Add "payment from"
        .Add "amount"
        .Add "balance"
        .Add "payment date"
        .Add "remarks"
        .Add "received_by"
       
    End With
     setListviewColumn lsv, payment_column
    Set payment_column = Nothing
End Sub
Sub setSalesListview(lsv As ListView)
    Dim sales_column As New Collection
    'sales_order_no, customer_name, Name, discount, grand_total, net_total, tendered_amount, change, delivery_date
    With sales_column
        .Add "Sales Order No"
        .Add "Customer Name"
        .Add "Agent Name"
        .Add "Discount"
        .Add "Grand Total"
        .Add "Net Total"
        .Add "Tendered Amount"
        .Add "Change"
        .Add "Delivery Date"
        .Add "Prepared by"
    End With
     setListviewColumn lsv, sales_column
    Set stockInColumns = Nothing
End Sub
Sub setStockInPreviewListview(lsv As ListView)
    Dim stockInColumns As New Collection
    
    With stockInColumns
        .Add "Item Id"
        .Add "Item Code"
        .Add "Description"
        .Add "UM"
        .Add "Quantity"
    End With
    setListviewColumn lsv, stockInColumns
    Set stockInColumns = Nothing
End Sub

Sub setUserAccountColumn(lsv As ListView)
    Dim useraccount_column As New Collection
    
    With useraccount_column
        .Add "Username"
        .Add "Password"
        .Add "User Type"
    End With
    
    setListviewColumn lsv, useraccount_column
    
    Set useraccount_column = Nothing
End Sub

Sub setStockInListview(lsv As ListView)
    Dim stockInColumns As New Collection
    
    With stockInColumns
        .Add "Item Id"
        .Add "Item Code"
        .Add "Qty"
    End With
    setListviewColumn lsv, stockInColumns
    Set stockInColumns = Nothing
End Sub

Sub setDiscountColumns(lsv As ListView)
    Dim discount As New Collection
    
    With discount
        .Add "Discount Id"
        .Add "Discount Code"
        .Add "Discount Name"
        .Add "Amount"
    End With
    setListviewColumn lsv, discount
    Set discount = Nothing

End Sub
Sub setMunicipalColumns(lsv As ListView)
    Dim municipal_col As New Collection
    
    With municipal_col
        .Add "Municipal ID"
        .Add "Municipality Name"
    End With
    setListviewColumn lsv, municipal_col
    
    Set municipal_col = Nothing
End Sub

'Items Description
Sub setItemsDescriptionColumns(lsv As ListView)
    Dim ItemsDescriptionColumns As New Collection
    
    With ItemsDescriptionColumns
        .Add "Item id"
        .Add "Item Code"
        .Add "Item Name"
        .Add "Item Description"
        .Add "No. of stocks"
        .Add "Price"
        .Add "Dealers Price"
        .Add "Unit of Measure"
        .Add "Manufacturer"
    End With
    setListviewColumn lsv, ItemsDescriptionColumns
    Set ItemsDescriptionColumns = Nothing
End Sub


'manufactures listview
Sub setManufacturersColumns(lsv As ListView)

    Dim manufacturers_column As New Collection
    
    manufacturers_column.Add "id"
    manufacturers_column.Add "Name"
    manufacturers_column.Add "Address"
    manufacturers_column.Add "Phone No."
    'manufacturers_column.Add "email"
    
    'this function assign all the values in the collection to column header of the listview
    setListviewColumn lsv, manufacturers_column
    
    Set manufacturers_column = Nothing
End Sub

Sub setCustomersColumns(lsv As ListView)
    Dim customer_column As New Collection
    
    With customer_column
        .Add "CustomerID"
        .Add "Customer name"
        .Add "Address"
        .Add "Conctact Number"
        .Add "Dealers type"
    End With
    setListviewColumn lsv, customer_column
    
   Set setCustomersColumn = Nothing
End Sub

Sub setAgentColumns(lsv As ListView)
    Dim agent_column As New Collection
    
    With agent_column
        .Add "AgentID"
        .Add "Agent name"
        .Add "Addres"
        .Add "Conctact Number"
    End With
    setListviewColumn lsv, agent_column
    
   Set agent_column = Nothing
End Sub

'this function assign all the values in the collection to column header of the listview
Public Function setListviewColumn(ByVal lsv As ListView, column As Collection, Optional width As Integer = 1500)
    
    For Each col In column
        lsv.ColumnHeaders.Add , , col, width
    Next

End Function

Public Function setListviewColumnWidth(ByVal lsv As ListView, columnWidth As Collection, Optional width As Integer = 1500)
    lsv.ColumnHeaders.Clear
    For Each col In columnWidth
        Dim x As Integer
        x = 0
        lsv.ColumnHeaders(x).width = width
    Next

End Function
Public Sub hideAllColumnsExept(columns As String, lsv As ListView)
    Dim x As Integer
    
    For x = 1 To lsv.ColumnHeaders.Count
        If lsv.ColumnHeaders(x).Text <> columns Then
           lsv.ColumnHeaders(x).width = 0
        End If
    Next x
End Sub

