VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaras_of As String 'local copy
Public Property Let as_of(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.as_of = 5
    mvaras_of = vData
End Property


Public Property Get as_of() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.as_of
    as_of = mvaras_of
End Property

Private Sub clearLastInventory()
    db.execute "truncate last_inventory"
End Sub

Private Sub prepareNewInventory()
    Dim cmd As String
    Dim sql As String
    Dim new_inventory As String
    Dim updateDate As String
    
    If hasLastInventory Then
        sql = "SELECT li.`item_id`, li.`item_code`, li.`ending_balance`,i.item_qty FROM `last_inventory` li LEFT JOIN items i ON li.item_id = i.item_id;"
    Else
        sql = "SELECT i.`item_id`, i.`item_code`, i.item_qty as beginning_balance,i.item_qty as ending_balance FROM `items` i;"
    End If
    cmd = "truncate `dbinventory`.`temp`;"
    db.execute cmd
    cmd = "INSERT INTO temp " & sql
    db.execute cmd
    
    Call clearLastInventory
    new_inventory = "INSERT INTO last_inventory (item_id,item_code,beginning_balance,ending_balance)(SELECT * from temp)"
    db.execute new_inventory
    
    updateDate = "UPDATE last_inventory SET `date` = NOW()"
    db.execute updateDate
End Sub

Sub destroy_temp_table()
    db.execute "DROP TABLE temp"
End Sub



Public Sub setNewEndingBalance()
    Dim insert As String
    Dim setDateToday As String
    Dim sql As String
    'sql = "SELECT li.`item_id`, li.`item_code`, li.`ending_balance`,i.item_qty FROM `last_inventory` li LEFT JOIN items i ON li.item_id = i.item_id"
    sql = "SELECT item_id,item_code,beginning_balance,ending_balance,date from last_inventory"
        Call prepareNewInventory
        
        insert = "INSERT INTO inventory (item_id,item_code,beginning_balance,ending_balance,date) " & _
                 "(" & sql & ")"
        db.execute insert
    
End Sub

Public Function hasLastInventory() As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT * from last_inventory"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        hasLastInventory = True
    Else
        hasLastInventory = False
    End If
    
    Set rs = Nothing
End Function
Public Function getTodaysInventoryForPrinting(item_category As String) As ADODB.Recordset
    Dim sql As String
    
    If item_category = "All" Then
        sql = "SELECT inv.*,id.item_description from last_inventory inv LEFT JOIN items_description id ON inv.item_code = id.item_code order by inv.item_code"
    Else
        sql = "SELECT inv.*,id.item_description from " & _
              "  last_inventory inv " & _
              "  LEFT JOIN item_category itc " & _
              "  ON inv.item_code = itc.item_code " & _
              "  LEFT JOIN items_description id " & _
              "  ON inv.item_code = id.item_code " & _
              "  WHERE itc.category = '" & item_category & "'"
    End If
    
    Set getTodaysInventoryForPrinting = db.execute(sql)
    
End Function


