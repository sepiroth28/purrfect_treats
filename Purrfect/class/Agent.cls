VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaragent_name As String 'local copy
Private mvaragent_contact_number As String 'local copy
'local variable(s) to hold property value(s)
Private mvaragent_id As String 'local copy
'local variable(s) to hold property value(s)
Private mvaragent_address As String 'local copy
Public Sub addCustomerToThisAgent(customer_id As Integer)
    Dim insert As String
    Dim delete As String
    
    delete = "DELETE FROM agent_customers WHERE agent_id = " & Me.agent_id & " AND customers_id = " & customer_id
    db.execute delete
    
    insert = "INSERT INTO agent_customers VALUES (" & Me.agent_id & "," & customer_id & ")"
    db.execute insert
    
End Sub

Public Sub removeCustomerOnThisAgent(customer_id As Integer)
    Dim Remove As String
    Remove = "DELETE FROM agent_customers WHERE agent_id = " & Me.agent_id & " and customers_id = " & customer_id
    db.execute Remove
End Sub


Public Sub removeAll()
     Dim lst As ListItem
    
    For Each lst In lsv
        Me.removeAssignedMunicipality (Val(lst.Text))
    Next
    
End Sub

Public Sub assignedAll(ByVal lsv As ListView)
    Dim lst As ListItem
    
    For Each lst In lsv.ListItems
        Me.assignMunicipal (Val(lst.Text))
    Next
    
End Sub

Public Sub removeAssignedMunicipality(municipal_id As Integer)
    Dim Remove As String
    Remove = "DELETE FROM municipal_agent WHERE agent_id = " & Me.agent_id & " AND municipal_id = " & municipal_id
    db.execute Remove
End Sub

Public Sub assignMunicipal(municipal_id As Integer)
    Dim insert As String
    Dim delete As String
    
    delete = "DELETE FROM municipal_agent WHERE agent_id =" & Me.agent_id & " AND municipal_id = " & municipal_id
    db.execute delete
    
    insert = "INSERT INTO municipal_agent VALUES(" & Me.agent_id & "," & municipal_id & ")"
    db.execute insert
End Sub

Public Function loadUnAssignedMunicipalities(lsv As ListView) As Boolean
    Dim rs As New ADODB.Recordset
    Dim lst As ListItem
    Dim sql As String
    loadUnAssignedMunicipalities = False
    'municipal_id, municipal_name FROM
    sql = "SELECT * FROM `dbinventory`.`municipalities` WHERE municipal_id NOT IN (SELECT GROUP_CONCAT(municipal_id SEPARATOR ',') FROM municipal_agent WHERE agent_id = " & Me.agent_id & ")"
    Set rs = db.execute(sql)
    lsv.ListItems.Clear
    If rs.RecordCount > 0 Then
        loadUnAssignedMunicipalities = True
        Do Until rs.EOF
           Set lst = lsv.ListItems.Add(, , rs.Fields(0).Value)
           lst.SubItems(1) = rs.Fields(1).Value
           
        rs.MoveNext
        Loop
    Else
        loadUnAssignedMunicipalities = False
    End If
    Set rs = Nothing
End Function

Public Sub loadAssignedMunicipalities(lsv As ListView)
    Dim rs As New ADODB.Recordset
    Dim lst As ListItem
    Dim sql As String
    'municipal_id, municipal_name FROM
    sql = "SELECT ma.municipal_id,m.municipal_name FROM municipal_agent ma INNER JOIN municipalities m ON ma.municipal_id = m.municipal_id WHERE ma.agent_id = " & Me.agent_id
    lsv.ListItems.Clear
    Set rs = db.execute(sql)
    
        Do Until rs.EOF
           Set lst = lsv.ListItems.Add(, , rs.Fields(0).Value)
           lst.SubItems(1) = rs.Fields(1).Value
        rs.MoveNext
        Loop
    
    Set rs = Nothing
   
End Sub

Public Property Let agent_address(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_address = 5
    mvaragent_address = vData
End Property


Public Property Get agent_address() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_address
    agent_address = mvaragent_address
End Property



Public Property Let agent_id(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_id = 5
    mvaragent_id = vData
End Property


Public Property Get agent_id() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_id
    agent_id = mvaragent_id
End Property



Public Sub UpdateAgent()
    Dim tbl_Update_Agent As String
    'agent_id, Name, Mobile, address
        tbl_Update_Agent = "UPDATE agent SET " & _
            "`Name` = '" & Me.agent_name & "'," & _
            "`Mobile` = '" & Me.agent_contact_number & "'," & _
            "`address` = '" & Me.agent_address & "'" & _
            " WHERE agent_id = " & Me.agent_id
        
        db.execute tbl_Update_Agent
End Sub

Public Function InsertAgent() As Integer
    Dim tbl_Insert_Agent As String
    
    tbl_Insert_Agent = "INSERT INTO agent " & _
                       " VALUES(" & Me.agent_id & "," & _
                       "'" & Me.agent_name & "'," & _
                       "'" & Me.agent_contact_number & "'," & _
                       "'" & Me.agent_address & "')"
    
    db.execute tbl_Insert_Agent
  Me.agent_id = db.execute("select last_insert_id()").Fields(0).Value
End Function


Public Function load_agent(agent_id As Integer) As agent
    'loads specific records of items
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim temp As New agent
    
    'item_id,item_code, item_qty, item_price, date_added, date_modified, manufacturers_id, reorder_point, item_code, item_name, item_description, image, status, unit_of_measure
'    sql = "SELECT * FROM `items` LEFT JOIN `items_description` ON `items`.item_code = `items_description`.item_code WHERE `items`.item_id = " & item_id
    sql = "SELECT agent_id,name,address,mobile FROM agent WHERE agent_id = " & agent_id
    Set rs = db.execute(sql)
    
'    Me.setItemId (rs.Fields("item_id").Value)
On Error Resume Next
    With Me
        .agent_id = rs.Fields("agent_id").Value
        .agent_name = rs.Fields("name").Value
        .agent_address = rs.Fields("address").Value
        .agent_contact_number = rs.Fields("mobile").Value
           
    End With
    Set rs = Nothing
    'Set load_item = temp
End Function




Public Property Let agent_contact_number(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_contact_number = 5
    mvaragent_contact_number = vData
End Property


Public Property Get agent_contact_number() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_contact_number
    agent_contact_number = mvaragent_contact_number
End Property

Public Property Let agent_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_name = 5
    mvaragent_name = vData
End Property


Public Property Get agent_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_name
    agent_name = mvaragent_name
End Property

Private Sub Class_Initialize()
    Me.agent_id = "NULL"
End Sub
