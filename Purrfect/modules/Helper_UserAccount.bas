Attribute VB_Name = "Helper_UserAccount"
Sub loadAllUserAccountToListview(lsv As ListView)
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    Dim useraccount As New user_account
    lsv.ListItems.Clear
    Set Collection = getAllUserAccountCollection
        For Each useraccount In Collection
            Set list = lsv.ListItems.Add(, , useraccount.username)
            list.SubItems(1) = useraccount.Password
            list.SubItems(2) = useraccount.user_type
        Next
End Sub

Function getAllUserAccountCollection() As UserAccountCollection
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim useraccount_col As New UserAccountCollection
    Dim temp_useraccount As New user_account
    
    
    sql = "SELECT * FROM useraccount"
    Set data = db.execute(sql)
    On Error Resume Next
    Do Until data.EOF
        With temp_useraccount
            .username = data.Fields("username").Value
            .Password = data.Fields("password").Value
            .user_type = data.Fields("user_type").Value
        End With
         useraccount_col.Add temp_useraccount, data.Fields("username").Value
         data.MoveNext
    Loop
   
    Set getAllUserAccountCollection = useraccount_col
End Function

Sub DeleteUserAccount(ByVal username As String)
    Dim tbl_delete_useraccount As String
    
    tbl_delete_useraccount = "DELETE FROM useraccount WHERE username = '" & username & "'"
    db.execute (tbl_delete_useraccount)
                                
End Sub
