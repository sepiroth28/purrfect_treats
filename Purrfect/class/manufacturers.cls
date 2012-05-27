VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "manufacturers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "manufacturer class"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarmanufacturers_id As Integer 'local copy
Private mvarmanufacturers_name As String 'local copy
Private mvarmanufacturers_add As String 'local copy
Private mvarmanufacturers_number As String 'local copy

'manufacturers_id, manufacturers_name, manufacturers_add, manufacturers_number
Public Function load_manufacturers(id As Integer) As manufacturers
    'loads specific records on the manufacturers class
    Dim man As New manufacturers
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT * from manufacturers WHERE  manufacturers_id = " & id
    Set rs = db.execute(sql)
    
    If Not rs.EOF Then
        With Me
            .manufacturers_id = rs.Fields("manufacturers_id").Value
            .manufacturers_name = rs.Fields("manufacturers_name").Value
            .manufacturers_add = rs.Fields("manufacturers_add").Value
            .manufacturers_number = rs.Fields("manufacturers_number").Value
        End With
    End If
   ' Set load_manufacturers = man
    Set rs = Nothing
    
End Function

Public Sub update()
Attribute update.VB_Description = "Updates records"
    Dim update As String
    'manufacturers_id, manufacturers_name, manufacturers_add, manufacturers_number
    update = "UPDATE manufacturers SET " & _
                "manufacturers_name = '" & Me.manufacturers_name & "'," & _
                "manufacturers_add = '" & Me.manufacturers_add & "'," & _
                "manufacturers_number = '" & Me.manufacturers_number & "' " & _
                "WHERE " & _
                "manufacturers_id = " & Me.manufacturers_id
    
    db.execute update
    'MsgBox update
End Sub

Public Sub insert()
Attribute insert.VB_Description = "insert record"
    Dim tbl_item_insert As String
    Dim tbl_item_desc As String
   
    'manufacturers_id, manufacturers_name, manufacturers_add, manufacturers_number
    tbl_item_insert = "INSERT INTO manufacturers " & _
            "VALUES (" & _
            "null," & _
            "'" & Me.manufacturers_name & "'," & _
            "'" & Me.manufacturers_add & "'," & _
            "'" & Me.manufacturers_number & "'" & _
            ")"
    
    db.execute tbl_item_insert
    'MsgBox tbl_item_insert
End Sub

Public Property Let manufacturers_number(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.manufacturers_number = 5
    mvarmanufacturers_number = vData
End Property


Public Property Get manufacturers_number() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.manufacturers_number
    manufacturers_number = mvarmanufacturers_number
End Property



Public Property Let manufacturers_add(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.manufacturers_add = 5
    mvarmanufacturers_add = vData
End Property


Public Property Get manufacturers_add() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.manufacturers_add
    manufacturers_add = mvarmanufacturers_add
End Property



Public Property Let manufacturers_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.manufacturers_name = 5
    mvarmanufacturers_name = vData
End Property


Public Property Get manufacturers_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.manufacturers_name
    manufacturers_name = mvarmanufacturers_name
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



