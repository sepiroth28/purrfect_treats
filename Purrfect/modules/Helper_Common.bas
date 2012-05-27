Attribute VB_Name = "Helper_Common"
Sub LoadToCombo(ByVal collectionToCombo As collection, cbo As ComboBox)
    'Dim valuesToCombo As String
    
    cbo.Clear
    For Each valuesToCombo In collectionToCombo
        cbo.AddItem valuesToCombo
    Next
End Sub

Sub openListview(lsv As ListView)
    If lsv.Visible = False Then
        lsv.Visible = True
    End If
End Sub

Sub closeListView(lsv As ListView)
    If lsv.Visible = True Then
        lsv.Visible = False
    End If
End Sub

Sub toogleListView(lsv As ListView)
    If lsv.Visible = True Then
        lsv.Visible = False
    ElseIf lsv.Visible = False Then
        lsv.Visible = True
    End If
End Sub
