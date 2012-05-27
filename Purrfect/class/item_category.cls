VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "item_category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvaritem_code As String 'local copy
Private mvarcategory As String 'local copy
Public Sub load_item_category(item_code As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT * FROM item_category WHERE item_code = '" & item_code & "'"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        Me.item_code = rs.Fields(0).Value
        Me.category = rs.Fields(1).Value
    End If
End Sub
Public Sub update()
Dim sql As String
Dim delete As String

delete = "DELETE FROM item_category WHERE item_code = '" & Me.item_code & "'"
db.execute delete

sql = "INSERT INTO item_category VALUES ('" & Me.item_code & "','" & Me.category & "')"
db.execute sql
End Sub
Public Sub save()
    Dim insert As String
    Dim last_id As String
    
    insert = "INSERT INTO item_category VALUES ('" & Me.item_code & "','" & Me.category & "')"
    db.execute insert
    
    
    
End Sub

Public Property Let category(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.category = 5
    mvarcategory = vData
End Property


Public Property Get category() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.category
    category = mvarcategory
End Property



Public Property Let item_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_id = 5
    mvaritem_code = vData
End Property


Public Property Get item_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_id
     item_code = mvaritem_code
End Property



