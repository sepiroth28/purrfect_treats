VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00E1A00B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer"
   ClientHeight    =   6615
   ClientLeft      =   4470
   ClientTop       =   2565
   ClientWidth     =   7440
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelectAgent 
      Caption         =   "..."
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
      Left            =   3600
      TabIndex        =   11
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6345
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox cboMunicipalities 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1740
         TabIndex        =   15
         Top             =   2460
         Width           =   2895
      End
      Begin VB.TextBox txtCustomersAddress 
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
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   6300
         Width           =   6675
      End
      Begin VB.ComboBox cboDealersType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   3420
         TabIndex        =   13
         Top             =   3720
         Width           =   3492
      End
      Begin MSComctlLib.ListView lsvAgent 
         Height          =   1095
         Left            =   240
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1931
         View            =   3
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
         Top             =   4680
         Width           =   3255
      End
      Begin VB.TextBox txtContactNumber 
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
         TabIndex        =   2
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox txtCustomersName 
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
         Top             =   1200
         Width           =   6735
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
         Left            =   4320
         TabIndex        =   4
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ", Bohol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4740
         TabIndex        =   17
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Municipalities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   16
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealers type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   3420
         TabIndex        =   12
         Top             =   3420
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent's Name"
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
         TabIndex        =   9
         Top             =   4320
         Width           =   1365
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
         TabIndex        =   8
         Top             =   120
         Width           =   6735
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
         TabIndex        =   7
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer's Name "
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
         Top             =   840
         Width           =   1755
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
         TabIndex        =   5
         Top             =   3420
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_customer As New Customers

Private Sub cmdAddNewItem_Click()
    saveData
    editmode = False
    Unload Me
Call loadAllCustomersToListview(frmManageCustomer.lsvCustomer)

End Sub

Sub saveData()
Dim customer As New Customers
    Dim agent As New agent
Dim mvarcustomers_id As Integer
Dim mvarcustomers_name, mvarcustomers_add, mvarcustomers_number, mvardealers_type As String

'mvarcustomers_id = Val(txtCustomersID.Text)
mvarcustomers_name = txtCustomersName.Text
mvarcustomers_add = txtCustomersAddress.Text
mvarcustomers_number = txtContactNumber.Text
mvardealers_type = cboDealersType.Text

    
If editmode = True Then
    With edit_customer
  
        .customers_name = txtCustomersName.Text
        .customers_add = cboMunicipalities.Text & ",Bohol"
        .customers_number = txtContactNumber.Text
        .dealers_type = mvardealers_type
       If .mvaragent.agent_id <> "NULL" Then
              If .mvaragent.agent_id <> Val(lsvAgent.SelectedItem.Text) Then
               Call .mvaragent.removeCustomerOnThisAgent(.customers_id)
               agent.agent_id = Val(lsvAgent.SelectedItem.Text)
               
               Call agent.addCustomerToThisAgent(.customers_id)
                    .updateData
              Else
                    .updateData
              End If
       
       Else
            
            agent.agent_id = Val(lsvAgent.SelectedItem.Text)
               
            Call agent.addCustomerToThisAgent(.customers_id)
                .updateData
            
          
        End If
   End With

    'MsgBox "Record successfully updated."
Else
    With customer
        .customers_name = mvarcustomers_name
        .customers_add = cboMunicipalities.Text & ",Bohol"
        .customers_number = mvarcustomers_number
        .dealers_type = mvardealers_type
        Dim newCustomerID As Integer
        agent.agent_id = Val(lsvAgent.SelectedItem.Text)
        newCustomerID = .insert
        
        agent.addCustomerToThisAgent (newCustomerID)
        'MsgBox .customers_id & " " & .customers_name & " " & .customers_add & " " & .customers_number
    
    End With
End If
End Sub

Private Sub cmdSelectAgent_Click()
    Call toogleListView(lsvAgent)

End Sub

Private Sub Form_Load()
  txtAgentName.Locked = True
Call setAgentColumns(lsvAgent)
With lsvAgent
    .ColumnHeaders(1).width = 0
    .ColumnHeaders(2).width = 3500
    .ColumnHeaders(3).width = 0
    .ColumnHeaders(4).width = 0
End With

Call loadAgentToListview(lsvAgent)
    
    If editmode = True Then
        
       With edit_customer
            Call .load_customers(activecustomer)
            txtCustomersName.Text = .customers_name
            'txtCustomersAddress.Text = .customers_add
            Dim temp() As String
            temp = Split(.customers_add, ",")
            If UBound(temp) Then
                cboMunicipalities.Text = temp(0)
            End If
            txtContactNumber.Text = .customers_number
            txtAgentName.Text = .mvaragent.agent_name
            cboDealersType.Text = .dealers_type
       End With
       
       txtAgentName.Enabled = False
       cmdSelectAgent.Enabled = False
    End If
    
cboDealersType.AddItem DEALER
cboDealersType.AddItem CONSUMER

Call loadAllMunicipalitiesToCombo(cboMunicipalities)
End Sub

Private Sub lsvAgent_Click()
    txtAgentName.Text = lsvAgent.SelectedItem.SubItems(1)
    Call toogleListView(lsvAgent)
End Sub

Private Sub txtAgentName_Change()
  '  Call loadAgentToListview(lsvAgent)
End Sub
