VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerAgentIndex 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Agent Index Management"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12600
   Icon            =   "frmCustomerAgentIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   60
      ScaleHeight     =   6345
      ScaleWidth      =   12465
      TabIndex        =   0
      Top             =   60
      Width           =   12495
      Begin MSComctlLib.ListView lsvCustomerAgent 
         Height          =   5475
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   9657
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmCustomerAgentIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT c.customers_id,c.customers_name,ac.agent_id,a.agent_id,a.name FROM customers c " & _
        " left join agent_customers ac ON c.customers_id = ac.customers_id " & _
        " left join agent a ON ac.agent_id = a.agent_id"

Set rs = db.execute(sql)



End Sub

