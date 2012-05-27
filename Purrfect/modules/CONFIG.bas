Attribute VB_Name = "CONFIG"
Public DBSERVER As String
Public DB_NAME As String
Public DB_USERNAME As String
Public DB_PASSWORD As String

Public Const PAYMENT_COD As Integer = 1
Public Const PAYMENT_ACCOUNT_RECEIVABLE As Integer = 2

Public Const ITEM_IN_STOCK As Integer = 1
Public Const ITEM_OUT_OF_STOCK As Integer = 0

Public Const DEALER As String = "dealer"
Public Const CONSUMER As String = "consumer"

Public Const ADMIN As String = "admin"
Public Const USER As String = "user"

Sub initializedConfig()
Dim file_name As String
Dim intEmpFileNbr As Integer
Dim server As String
Dim dba_name As String
Dim dba_username As String
Dim dba_pass As String

intEmpFileNbr = FreeFile
file_name = App.Path & "\config.dat"

Open file_name For Input As #intEmpFileNbr

Input #intEmpFileNbr, server, dba_name, dba_username, dba_pass

 DBSERVER = server
 DB_NAME = dba_name
 DB_USERNAME = dba_username
 DB_PASSWORD = dba_pass

Close #intEmpFileNbr

End Sub
