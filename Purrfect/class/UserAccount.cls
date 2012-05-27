VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User_Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"UserAccountCollection"
'local variable(s) to hold property value(s)
Private mvarusername As String 'local copy
Private mvarpassword As String 'local copy
'local variable(s) to hold property value(s)
Private mvaruser_type As String 'local copy
Private mvarUserAccountCollection As UserAccountCollection
'local variable(s) to hold property value(s)
Private mvarpreviliges As New User_Previleges  'local copy
Public Property Set previliges(ByVal vData As User_Previleges)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.previliges = Form1
    Set mvarpreviliges = vData
End Property


Public Property Get previliges() As User_Previleges
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.previliges
    Set previliges = mvarpreviliges
End Property





Public Property Get UserAccountCollection() As UserAccountCollection
    If mvarUserAccountCollection Is Nothing Then
        Set mvarUserAccountCollection = New UserAccountCollection
    End If


    Set UserAccountCollection = mvarUserAccountCollection
End Property


Public Property Set UserAccountCollection(vData As UserAccountCollection)
    Set mvarUserAccountCollection = vData
End Property
Private Sub Class_Terminate()
    Set mvarUserAccountCollection = Nothing
End Sub

Sub SaveUserAccount()
    Dim tbl_save_useraccount As String
    
    tbl_save_useraccount = "INSERT INTO useraccount " & _
                           "VALUES( " & _
                           "'" & Me.username & "'," & _
                           "MD5('" & Me.Password & "')," & _
                           "'" & Me.user_type & "')"
     db.execute (tbl_save_useraccount)
     MsgBox "User Account has been successfully saved.", vbInformation, "Nutrimart Enterprise"
    
End Sub
Sub changePassword(new_password As String)
    Dim tbl_update_useraccount As String
    
    tbl_update_useraccount = "UPDATE useraccount SET " & _
                             "password =md5('" & new_password & "') " & _
                             " WHERE username ='" & Me.username & "'"
                             
    db.execute (tbl_update_useraccount)
    MsgBox "User Account has been successfully change password.", vbInformation, "Nutrimart Enterprise"
                             
End Sub
Sub UpdateUserAccount()
    Dim tbl_update_useraccount As String
    
    tbl_update_useraccount = "UPDATE useraccount SET " & _
                             "username ='" & Me.username & "'," & _
                             "user_type ='" & Me.user_type & "'" & _
                             " WHERE username ='" & Me.username & "'"
                             
    db.execute (tbl_update_useraccount)
    MsgBox "User Account has been successfully updated.", vbInformation, "Nutrimart Enterprise"
                             
End Sub

Sub DeleteUserAccount()
    Dim tbl_delete_useraccount As String
    
    tbl_delete_useraccount = "DELETE FROM useraccount WHERE username = '" & Me.username & "'"
    db.execute (tbl_delete_useraccount)
                                
End Sub

Public Function Check_UserAcount(ByVal username As String, ByVal Password As String) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT * FROM useraccount WHERE username ='" & username & "' AND password = md5('" & Password & "')"
    
    Set rs = db.execute(sql)
    
    If rs.RecordCount > 0 Then
        Check_UserAcount = True
    Else
        Check_UserAcount = False
    End If
    Set rs = Nothing
End Function

Function Check_UserType(ByVal username As String) As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    
    sql = "SELECT user_type FROM useraccount WHERE  username = '" & username & "'"
    
    Set rs = db.execute(sql)
    
    If rs.RecordCount > 0 Then
         Check_UserType = rs.Fields("user_type").Value
    End If
   
    
    Set rs = Nothing
    
End Function

Function loadUserAccount(ByVal username As String) As User_Account
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = "SELECT * FROM useraccount WHERE username = '" & username & "'"
    
    Set rs = db.execute(sql)
    
    Me.username = rs.Fields(0).Value
    Me.Password = rs.Fields(1).Value
    Me.user_type = rs.Fields(2).Value
    
End Function


Public Property Let user_type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.user_type = 5
    mvaruser_type = vData
End Property


Public Property Get user_type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.user_type
    user_type = mvaruser_type
End Property

Public Property Let Password(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarpassword = vData
End Property




Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarpassword
End Property



Public Property Let username(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarusername = vData
End Property


Public Property Get username() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    username = mvarusername
End Property

