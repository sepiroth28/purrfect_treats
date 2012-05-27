Attribute VB_Name = "MainModule"
Sub main()
Call initializedConfig

db.server = DBSERVER
db.database_name = DB_NAME
db.username = DB_USERNAME
db.Password = DB_PASSWORD
If db.connect Then
    frmLogIn.Show
'    MsgBox "Successfuly connected to database...", vbInformation, "Nutrimart"
Else
    MsgBox "Failed to connect to database", vbInformation, "Nutrimart"
End If
  

municipal_list = "Albur,Baclayon,Tagbilaran"

End Sub
