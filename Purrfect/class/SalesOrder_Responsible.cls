VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SalesOrder_Responsible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarprepared_by As String 'local copy
Private mvarchecked_by As String 'local copy
Private mvarposted_by As String 'local copy
Private mvardelivered_by As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSalesOrder_Responsible_id As String 'local copy
Public Property Let SalesOrder_Responsible_id(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SalesOrder_Responsible_id = 5
    mvarSalesOrder_Responsible_id = vData
End Property


Public Property Get SalesOrder_Responsible_id() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SalesOrder_Responsible_id
    SalesOrder_Responsible_id = mvarSalesOrder_Responsible_id
End Property



Public Property Let delivered_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.delivered_by = 5
    mvardelivered_by = vData
End Property


Public Property Get delivered_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.delivered_by
    delivered_by = mvardelivered_by
End Property



Public Property Let posted_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.posted_by = 5
    mvarposted_by = vData
End Property


Public Property Get posted_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.posted_by
    posted_by = mvarposted_by
End Property



Public Property Let checked_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.checked_by = 5
    mvarchecked_by = vData
End Property


Public Property Get checked_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.checked_by
    checked_by = mvarchecked_by
End Property



Public Property Let prepared_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.prepared_by = 5
    mvarprepared_by = vData
End Property


Public Property Get prepared_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.prepared_by
    prepared_by = mvarprepared_by
End Property

Sub Save_SalesOrder_Responsible()
    Dim tbl_save_SalesOrder_Responsible As String
    
    db.execute "DELETE FROM salesorder_responsible"
    
    tbl_save_SalesOrder_Responsible = "INSERT INTO salesorder_responsible " & _
                                      " VALUES(null,'" & Me.prepared_by & "'," & _
                                      " '" & Me.checked_by & "'," & _
                                      " '" & Me.posted_by & "'," & _
                                      " '" & Me.delivered_by & "')"
    
    db.execute tbl_save_SalesOrder_Responsible
    
End Sub

Sub Update_SalesOrder_Responsible()

End Sub

Sub loadToSalesOrder_Responsible()
    
    Dim rs As New ADODB.Recordset
    Dim str As String
    
    'SaleOrder_Responsible_Id, Prepared_by, Checked_by, Posted_by, Delivered_by
    str = "SELECT * FROM salesorder_responsible"
    
    Set rs = db.execute(str)
    If rs.RecordCount > 0 Then
        Me.prepared_by = rs.Fields(1).Value
        Me.checked_by = rs.Fields(2).Value
        Me.posted_by = rs.Fields(3).Value
        Me.delivered_by = rs.Fields(4).Value
    End If
    Set rs = Nothing
    
    
End Sub


Private Sub Class_Initialize()
    Me.SalesOrder_Responsible_id = "NULL"
    
    
End Sub
