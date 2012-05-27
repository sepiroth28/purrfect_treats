VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgentForm 
   BackColor       =   &H00E1A00B&
   Caption         =   "Agent Form"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "frmAgentForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   8355
      Left            =   60
      ScaleHeight     =   8325
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   60
      Width           =   7215
      Begin VB.CommandButton cmdRemovedAll 
         Caption         =   ">>"
         Height          =   495
         Left            =   3360
         TabIndex        =   14
         Top             =   6300
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdAssignedAll 
         Caption         =   "<<"
         Height          =   495
         Left            =   3360
         TabIndex        =   13
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView lsvUnAssigned 
         Height          =   2775
         Left            =   4080
         TabIndex        =   11
         Top             =   4680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "municipal_name"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   ">"
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   5220
         Width           =   615
      End
      Begin VB.CommandButton cmdAssigned 
         Caption         =   "<"
         Height          =   495
         Left            =   3360
         TabIndex        =   9
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdAddNewItem 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4440
         TabIndex        =   4
         Top             =   7560
         Width           =   2655
      End
      Begin VB.TextBox txtAgentName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox txtAgentAddress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2160
         Width           =   6735
      End
      Begin VB.TextBox txtAgentContactNumber 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   3720
         Width           =   3015
      End
      Begin MSComctlLib.ListView lsvAssigned 
         Height          =   2775
         Left            =   240
         TabIndex        =   12
         Top             =   4680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "municipal_name"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   3420
         Width           =   1605
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent's Name "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label lblRequiredMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Please fill up requireed fields..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmAgentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_agent As New agent
Private Sub cmdAddNewItem_Click()
saveData
 
End Sub

Sub saveData()
 Dim mvAgent_name As String
 Dim mvagent_address, mvagent_contact_number As String
 Dim newAgent As New agent
 'Dim newMunicipalities As New Municipalities
 mvAgent_name = txtAgentName.Text
 mvagent_address = txtAgentAddress.Text
 mvagent_contact_number = txtAgentContactNumber.Text
 

    
    If editmode = True Then
        edit_agent.agent_name = txtAgentName.Text
        edit_agent.agent_address = txtAgentAddress.Text
        edit_agent.agent_contact_number = txtAgentContactNumber.Text
        edit_agent.UpdateAgent
        Call edit_agent.assignedAll(lsvAssigned)
    Else
       With newAgent
        .agent_name = mvAgent_name
        .agent_contact_number = mvagent_contact_number
        .agent_address = mvagent_address
        .InsertAgent
         Call .assignedAll(lsvAssigned)
       End With
        MsgBox "Record has been successfully save.", vbInformation, "NutriMart"
    End If
    

Call loadAgentToListview(frmManageAgent.lsvAgent)
Unload Me
End Sub

Private Sub cmdAssigned_Click()
    Dim lst As ListItem
  If editmode = True Then
    Call edit_agent.assignMunicipal(Val(lsvUnAssigned.SelectedItem.Text))
    
    Call edit_agent.loadAssignedMunicipalities(lsvAssigned)
    
    noAssonedYet = edit_agent.loadUnAssignedMunicipalities(lsvUnAssigned)
        
        If noAssonedYet = False Then
            Call loadAllMunicipalities(lsvUnAssigned)
        End If
  Else
    Set lst = lsvAssigned.ListItems.Add(, , lsvUnAssigned.SelectedItem.Text)
        lst.SubItems(1) = lsvUnAssigned.SelectedItem.SubItems(1)
    lsvUnAssigned.ListItems.Remove (lsvUnAssigned.SelectedItem.Index)
  End If
End Sub

Private Sub cmdRemove_Click()

If lsvAssigned.ListItems.Count > 0 Then
    Dim lst As ListItem
    If editmode = True Then
       Call edit_agent.removeAssignedMunicipality(Val(lsvAssigned.SelectedItem.Text))
       Call edit_agent.loadAssignedMunicipalities(lsvAssigned)
       
       noAssonedYet = edit_agent.loadUnAssignedMunicipalities(lsvUnAssigned)
           
           If noAssonedYet = False Then
               Call loadAllMunicipalities(lsvUnAssigned)
           End If
    Else
       Set lst = lsvUnAssigned.ListItems.Add(, , lsvAssigned.SelectedItem.Text)
           lst.SubItems(1) = lsvAssigned.SelectedItem.SubItems(1)
       lsvAssigned.ListItems.Remove (lsvAssigned.SelectedItem.Index)
    End If
End If
End Sub

Private Sub Form_Load()
Dim noAssonedYet As Boolean
 If editmode = True Then

        With edit_agent
            .load_agent activeItemId
        txtAgentName.Text = .agent_name
        txtAgentAddress.Text = .agent_address
        txtAgentContactNumber = .agent_contact_number
        Call .loadAssignedMunicipalities(lsvAssigned)
        noAssonedYet = .loadUnAssignedMunicipalities(lsvUnAssigned)
        
        If noAssonedYet = False Then
            Call loadAllMunicipalities(lsvUnAssigned)
        End If
        
        End With
    
Else
    Call loadAllMunicipalities(lsvUnAssigned)
End If


End Sub

Private Sub lstMunicipal_Click()

End Sub
