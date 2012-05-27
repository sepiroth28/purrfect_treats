VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "item class"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"item_category"
'local variable(s) to hold property value(s)
Private mvaritem_code As String 'local copy
Private mvaritem_qty As Double 'local copy
Private mvardate_added As String 'local copy
Private mvarmanufacturers_id As Integer 'local copy
Private mvarreorder_point As Double 'local copy
Private mvaritem_name As String 'local copy
Private mvaritem_description As String 'local copy
Private mvarimage As String 'local copy
Private mvarstatus As Integer 'local copy
Private mvarunit_of_measure As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvaritem_price As Double 'local copy
'local variable(s) to hold property value(s)
Private mvardate_modified As String 'local copy
'local variable(s) to hold property value(s)
Private mvarmanufacturer As New manufacturers 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvaritem_id As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvardealers_price As Double 'local copy
'local variable(s) to hold property value(s)
Private mvaritem_category As New item_category 'local copy
'local variable(s) to hold property value(s)
Private mvartemp_item_code As String 'local copy
'local variable(s) to hold property value(s)
Private mvarinclude_in_rebate As Boolean 'local copy
Public Property Let include_in_rebate(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.include_in_rebate = 5
    mvarinclude_in_rebate = vData
End Property


Public Property Get include_in_rebate() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.include_in_rebate
    include_in_rebate = mvarinclude_in_rebate
End Property



Public Property Let temp_item_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.temp_item_code = 5
    mvartemp_item_code = vData
End Property


Public Property Get temp_item_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.temp_item_code
    temp_item_code = mvartemp_item_code
End Property



Public Property Set item_category(ByVal vData As item_category)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.item_category = Form1
    Set mvaritem_category = vData
End Property


Public Property Get item_category() As item_category
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_category
    Set item_category = mvaritem_category
End Property



Public Property Let dealers_price(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dealers_price = 5
    mvardealers_price = vData
End Property


Public Property Get dealers_price() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dealers_price
     dealers_price = mvardealers_price
End Property



Public Sub stockOut(qty As Double)
    db.execute "UPDATE items SET item_qty = item_qty - " & qty & " WHERE item_id = " & Me.item_id
End Sub

Public Sub addStock(qty As Double)
    db.execute "UPDATE items SET item_qty = item_qty + " & qty & " WHERE item_id = " & Me.item_id
End Sub
Function displayAvailability() As String
    If checkStockQty Then
        displayAvailability = "IN STOCK"
    Else
        displayAvailability = "OUT OF STOCK"
    End If
End Function
Function checkStockQty() As Integer
    If Me.item_qty > 0 Then
        checkStockQty = ITEM_IN_STOCK
    Else
        checkStockQty = ITEM_OUT_OF_STOCK
    End If
End Function
Public Property Let item_id(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.date_modified = 5
    mvaritem_id = vData
End Property

Public Property Get item_id() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.date_modified
    item_id = mvaritem_id
End Property


Public Property Get manufacturer() As manufacturers
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.manufacturer
    
    Set manufacturer = mvarmanufacturer
End Property

Public Property Let date_modified(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.date_modified = 5
    mvardate_modified = vData
End Property


Public Property Get date_modified() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.date_modified
    date_modified = mvardate_modified
End Property



Public Property Let item_price(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_price = 5
    mvaritem_price = vData
End Property

Public Property Get item_price() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_price
    item_price = mvaritem_price
End Property

Public Function load_item(item_id As Integer) As items
    'loads specific records of items
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim temp As New items
    
    'item_id,item_code, item_qty, item_price,dealers_price date_added, date_modified, manufacturers_id, reorder_point, item_code, item_name, item_description, image, status, unit_of_measure
    sql = "SELECT * FROM `items` LEFT JOIN `items_description` ON `items`.item_code = `items_description`.item_code WHERE `items`.item_id = " & item_id
    Set rs = db.execute(sql)
    
'    Me.setItemId (rs.Fields("item_id").Value)
On Error Resume Next
    With Me
        .item_id = rs.Fields("item_id").Value
        .item_code = rs.Fields("item_code").Value
        .temp_item_code = rs.Fields("item_code").Value
        .item_qty = rs.Fields("item_qty").Value
        .date_added = rs.Fields("date_added").Value
        .date_modified = rs.Fields("date_modified").Value
        .manufacturers_id = rs.Fields("manufacturers_id").Value
        .reorder_point = rs.Fields("reorder_point").Value
        .include_in_rebate = rs.Fields("include_in_rebate").Value
        .item_name = rs.Fields("item_name").Value
        .item_description = rs.Fields("item_description").Value
        '.image = rs.Fields("image").Value
        .item_status = rs.Fields("status").Value
        .unit_of_measure = rs.Fields("unit_of_measure").Value
        .item_price = rs.Fields("item_price").Value
        .dealers_price = rs.Fields("dealers_price").Value
        .manufacturer.load_manufacturers (.manufacturers_id)
        .item_category.load_item_category (.item_code)
    End With
    Set rs = Nothing
    'Set load_item = temp
End Function

Public Sub update()
    Dim update As String
    'item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    'table items_description : item_code, item_name, item_description, image, status, unit_of_measure
    
    update = "UPDATE items SET " & _
                "item_code = '" & Me.item_code & "'," & _
                "item_qty = " & Me.item_qty & "," & _
                "item_price = " & Me.item_price & "," & _
                "dealers_price = " & Me.dealers_price & "," & _
                "date_modified = CURDATE()," & _
                "manufacturers_id = " & Me.manufacturers_id & "," & _
                "reorder_point = " & Me.reorder_point & "," & _
                "include_in_rebate = " & Me.include_in_rebate & _
                " WHERE item_id = " & Me.item_id
    
    db.execute update
    
    update = "UPDATE items_description SET " & _
                "item_code = '" & Me.item_code & "'," & _
                "item_name = '" & Me.item_name & "'," & _
                "item_description = '" & Me.item_description & "'," & _
                "unit_of_measure = '" & Me.unit_of_measure & "'," & _
                "status = " & Me.item_status & _
                " WHERE item_code = '" & Me.temp_item_code & "'"
    db.execute update
    
    update = "UPDATE inventory set item_code = '" & Me.item_code & "' WHERE item_id = " & Me.item_id
    db.execute update
    
    'MsgBox update
End Sub

Public Sub insert()
    Dim tbl_item_insert As String
    Dim tbl_item_desc As String
    'table items : item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    'table items_description : item_code, item_name, item_description, image, status, unit_of_measure
    
    tbl_item_insert = "INSERT INTO items " & _
            "VALUES (null," & _
            "'" & Me.item_code & "'," & _
            "" & Me.item_qty & "," & _
            "" & Me.item_price & "," & _
            "" & Me.dealers_price & "," & _
            "" & Me.date_added & "," & _
            "" & Me.date_modified & "," & _
            "" & Me.manufacturers_id & "," & _
            "" & Me.reorder_point & "," & _
            "" & Me.include_in_rebate & ")"
    
    tbl_item_desc = "INSERT INTO items_description " & _
            "VALUES (" & _
            "'" & Me.item_code & "'," & _
            "'" & Me.item_name & "'," & _
            "'" & Me.item_description & "'," & _
            "'" & Me.image & "'," & _
            "" & Me.item_status & "," & _
            "'" & Me.unit_of_measure & "')"
    
    db.execute tbl_item_insert
    Dim last_id As Integer
    last_id = Val(db.execute("select last_insert_id()").Fields(0).Value)
    
    db.execute tbl_item_desc
    
    'item_id, item_code, beginning_balance, ending_balance, date
    Dim insert_to_last_inventory As String
    
    insert_to_last_inventory = "INSERT INTO last_inventory VALUES(" & last_id & ",'" & Me.item_code & "'," & Me.item_qty & "," & Me.item_qty & ",CURDATE())"
     
    
    'MsgBox tbl_item_desc
End Sub
Public Sub checkSql()
    Dim tbl_item_insert As String
    Dim tbl_item_desc As String
    'table items : item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point
    'table items_description : item_code, item_name, item_description, image, status, unit_of_measure
    
     tbl_item_insert = "INSERT INTO items " & _
            "VALUES (null," & _
            "'" & Me.item_code & "'," & _
            "" & Me.item_qty & "," & _
            "" & Me.item_price & "," & _
            "" & Me.date_added & "," & _
            "" & Me.date_modified & "," & _
            "" & Me.manufacturers_id & "," & _
            "" & Me.reorder_point & ")"
    
    tbl_item_desc = "INSERT INTO items_description " & _
            "VALUES (" & _
            "'" & Me.item_code & "'," & _
            "'" & Me.item_name & "'," & _
            "'" & Me.item_description & "'," & _
            "'" & Me.image & "'," & _
            "" & Me.item_status & "," & _
            "'" & Me.unit_of_measure & "')"
    
    'db.execute tbl_item_insert
    'db.execute tbl_item_desc
    MsgBox tbl_item_insert & vbCrLf & tbl_item_desc
End Sub
Public Property Let unit_of_measure(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.unit_of_measure = 5
    mvarunit_of_measure = vData
End Property

Public Property Set unit_of_measure(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.unit_of_measure = Form1
    Set mvarunit_of_measure = vData
End Property


Public Property Get unit_of_measure() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.unit_of_measure
    If IsObject(mvarunit_of_measure) Then
        Set unit_of_measure = mvarunit_of_measure
    Else
        unit_of_measure = mvarunit_of_measure
    End If
End Property

Public Property Let item_status(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.status = 5
    mvarstatus = vData
End Property

Public Property Get item_status() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.status
    item_status = mvarstatus
End Property

Public Property Let image(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.image = 5
    mvarimage = vData
End Property

Public Property Get image() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.image
    image = mvarimage
End Property

Public Property Let item_description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_description = 5
    mvaritem_description = vData
End Property

Public Property Get item_description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_description
    item_description = mvaritem_description
End Property

Public Property Let item_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_name = 5
    mvaritem_name = vData
End Property

Public Property Get item_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_name
    item_name = mvaritem_name
End Property

Public Property Let reorder_point(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.reorder_point = 5
    mvarreorder_point = vData
End Property


Public Property Get reorder_point() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.reorder_point
    reorder_point = mvarreorder_point
End Property

Public Property Let manufacturers_id(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.manufacturers_id = 5
    mvarmanufacturers_id = vData
End Property

Public Property Get manufacturers_id() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.manufacturers_id
    manufacturers_id = mvarmanufacturers_id
End Property

Public Property Let date_added(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.date_added = 5
    mvardate_added = vData
End Property

Public Property Get date_added() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.date_added
    date_added = mvardate_added
End Property

Public Property Let item_qty(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_qty = 5
    mvaritem_qty = vData
End Property

Public Property Get item_qty() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_qty
    item_qty = mvaritem_qty
End Property

Public Property Let item_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_code = 5
    mvaritem_code = vData
End Property

Public Property Get item_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_code
    item_code = mvaritem_code
End Property

Private Sub Class_Initialize()
    Me.date_added = "CURDATE()"
    Me.date_modified = "CURDATE()"
    Me.image = ""
    Me.reorder_point = 0
    Me.item_price = 0
    Me.dealers_price = 0
    Me.item_qty = 0
    Me.item_status = 1
End Sub
