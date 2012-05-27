Attribute VB_Name = "Helper_Municipal"
Public Sub loadAllMunicipalities(lsv As ListView)
    Dim rs As New ADODB.Recordset
    Dim lst As ListItem
    Dim sql As String
    'municipal_id, municipal_name FROM
    sql = "SELECT * FROM municipalities"
    Set rs = db.execute(sql)
    
        Do Until rs.EOF
           Set lst = lsv.ListItems.Add(, , rs.Fields(0).Value)
           lst.SubItems(1) = rs.Fields(1).Value
        rs.MoveNext
        Loop
    
    Set rs = Nothing
   
End Sub
Public Sub loadAllMunicipalitiesToCombo(cbo As ComboBox)
    Dim rs As New ADODB.Recordset
    Dim sql As String
    'municipal_id, municipal_name FROM
    sql = "SELECT * FROM municipalities"
    Set rs = db.execute(sql)
    
        Do Until rs.EOF
          cbo.AddItem rs.Fields("municipal_name").Value
        rs.MoveNext
        Loop
    
    Set rs = Nothing
   
End Sub
Sub loadMunicipalityToListView(lsv As ListView)
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    'Dim municipality As New Municipalities
    lsv.ListItems.Clear
    Set Collection = getAllMunicipalityCollection
        For Each municipality In Collection
            Set list = lsv.ListItems.Add(, , municipality.municipal_id)
            list.SubItems(1) = municipality.municipal_name
        Next
End Sub

'Function getAllMunicipalityCollection() As MunicipalityCollection
'    Dim sql As String
'    Dim data As ADODB.Recordset
'    Dim municipal_col As New MunicipalityCollection
'    Dim temp_municipal As New Municipalities
'
'    sql = "SELECT * FROM municipalities"
'    Set data = db.execute(sql)
'
'    Do Until data.EOF
'        With temp_municipal
'            .municipal_id = data.Fields("municipal_Id").Value
'            .municipal_name = data.Fields("municipal_name").Value
'
'        End With
'         municipal_col.Add temp_municipal, data.Fields("municipal_id").Value
'         data.MoveNext
'    Loop
'
'    Set getAllMunicipalityCollection = municipal_col
'End Function


