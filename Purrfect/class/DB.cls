VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarserver As String 'local copy
Private mvardatabase_name As String 'local copy
Private mvarusername As String 'local copy
Private mvarpassword As String 'local copy
Dim conn As New ADODB.Connection
Public Function execute(sql As String) As ADODB.Recordset
    Set execute = conn.execute(sql)
End Function

Public Sub die()
    conn.Close
End Sub

Public Function connect() As Boolean
conn.CursorLocation = adUseClient
    conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=" & Me.server & ";DATABASE=" & Me.database_name & ";UID=" & Me.username & ";PWD=" & Me.Password & ";"
    On Error GoTo Error
        conn.Open
        connect = True
Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Database connection error"
        connect = False
End Function

Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.password = 5
    mvarpassword = vData
End Property


Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.password
    Password = mvarpassword
End Property



Public Property Let username(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.username = 5
    mvarusername = vData
End Property


Public Property Get username() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.username
    username = mvarusername
End Property



Public Property Let database_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.database_name = 5
    mvardatabase_name = vData
End Property


Public Property Get database_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.database_name
    database_name = mvardatabase_name
End Property



Public Property Let server(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.server = 5
    mvarserver = vData
End Property


Public Property Get server() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.server
    server = mvarserver
End Property



