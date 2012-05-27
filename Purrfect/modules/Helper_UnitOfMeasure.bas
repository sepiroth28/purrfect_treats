Attribute VB_Name = "Helper_UnitOfMeasure"
Function getAllUnitOfMeasure() As adodb.Recordset
    Dim rs As New adodb.Recordset
    Dim sql As String
    sql = "SELECT DISTINCT unit_of_measure FROM `items_description`;"
    Set rs = db.execute(sql)
    Set getAllUnitOfMeasure = rs
    Set rs = Nothing
End Function

Sub loadUnitOfMeasureToCombo(cbo As ComboBox)
    Dim rs As New adodb.Recordset
    Set rs = getAllUnitOfMeasure
        Do Until rs.EOF
            cbo.AddItem rs.Fields(0).Value
        rs.MoveNext
        Loop
    Set rs = Nothing
End Sub
