VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChangeAgent 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Agent"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   Icon            =   "frmChangeAgent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   60
      ScaleHeight     =   4005
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      Begin MSComctlLib.ListView lsvAgent 
         Height          =   1515
         Left            =   180
         TabIndex        =   2
         Top             =   2340
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2672
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
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   5
         Top             =   3120
         Width           =   2055
      End
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
         Left            =   4860
         TabIndex        =   4
         Top             =   1860
         Width           =   615
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
         Left            =   180
         TabIndex        =   3
         Top             =   1860
         Width           =   4635
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Agent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
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
         Left            =   180
         TabIndex        =   6
         Top             =   1560
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   5460
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Agent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   2040
      End
   End
End
Attribute VB_Name = "frmChangeAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cus As New Customers

Private Sub cmdUpdate_Click()
If txtAgentName.Text <> "" Then
    Dim sql As String
    sql = "DELETE FROM agent_customers WHERE customers_id = " & activecustomer
    db.execute sql
    
    sql = "INSERT INTO agent_customers VALUES(" & Val(lsvAgent.SelectedItem.Text) & "," & activecustomer & ")"
    db.execute sql
    
    MsgBox "Update successfully...", vbOKOnly, "Update agent"
    Unload Me
End If
End Sub

Private Sub Form_Load()
cus.load_customers activecustomer
lblCustomerName.Caption = cus.customers_name

Call setAgentColumns(lsvAgent)
With lsvAgent
    .ColumnHeaders(1).width = 0
    .ColumnHeaders(2).width = 3500
    .ColumnHeaders(3).width = 0
    .ColumnHeaders(4).width = 0
End With

Call loadAgentToListview(lsvAgent)
End Sub

Private Sub lsvAgent_Click()
    txtAgentName.Text = lsvAgent.SelectedItem.SubItems(1)
    Call toogleListView(lsvAgent)
End Sub

Private Sub cmdSelectAgent_Click()
    Call toogleListView(lsvAgent)
End Sub
