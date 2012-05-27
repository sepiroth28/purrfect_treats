Attribute VB_Name = "Helper_Agent"
Sub loadAgentToListview(lsv As ListView)
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
'    Dim agent As New agent
    lsv.ListItems.Clear
    Set rs = getAllAgent
      Do Until rs.EOF
            Set list = lsv.ListItems.Add(, , rs.Fields(0).Value)
            list.SubItems(1) = rs.Fields(1).Value
            list.SubItems(2) = rs.Fields(2).Value
            list.SubItems(3) = rs.Fields(3).Value
    rs.MoveNext
    Loop
    
End Sub

Function getAllAgent() As ADODB.Recordset
    Dim sql As String
    Dim data As ADODB.Recordset
    Dim newAgent As New AgentCollections
    Dim temp_agent As New agent

    sql = "SELECT agent_id,name,address,mobile FROM agent"
'    Set data = db.execute(sql)

'    Do Until data.EOF
'        With temp_agent
'            .agent_id = data.Fields("agent_id").Value
'            .agent_name = data.Fields("name").Value
'            .agent_contact_number = data.Fields("mobile").Value
'            .agent_address = data.Fields("address").Value
'        End With
'         newAgent.Add temp_agent, data.Fields("agent_id").Value
'         data.MoveNext
'    Loop
    

    Set getAllAgent = db.execute(sql)
End Function


Function findAgentOnThisMunicipalities(municipal_id) As agent
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim foundAgent As New agent
    
    sql = "SELECT * FROM `municipal_agent` WHERE municipal_id =" & municipal_id
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        foundAgent.load_agent Val(rs.Fields(0).Value)
    End If
    Set findAgentOnThisMunicipalities = foundAgent
End Function

Sub deleteAgent(agent_id As Integer)
    Dim sql As String
      sql = "DELETE FROM agent WHERE agent_id = " & agent_id
     db.execute sql
End Sub
