VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlstCustomer 
   Caption         =   "List of Customer's"
   ClientHeight    =   8670
   ClientLeft      =   4740
   ClientTop       =   1440
   ClientWidth     =   13935
   LinkTopic       =   "Form3"
   ScaleHeight     =   8670
   ScaleWidth      =   13935
   Begin VB.CommandButton cboSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsvCustomers 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   13361
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   675
   End
End
Attribute VB_Name = "frmlstCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    editmode = False
    frmCustomer.Show
End Sub

Private Sub Form_Load()
Call setCustomersColumns(lsvCustomers)
Call loadAllCustomersToListview(lsvCustomers)
End Sub


Private Sub lsvCustomers_DblClick()
    editmode = True
    
    With frmCustomer
        .txtCustomerID.Text = lsvCustomers.SelectedItem.Text
        .txtCustomersName.Text = lsvCustomers.SelectedItem.ListSubItems(1).Text
        .txtAddress.Text = lsvCustomers.SelectedItem.ListSubItems(2).Text
        .txtTelNo.Text = lsvCustomers.SelectedItem.ListSubItems(3).Text
        .Show
    End With
End Sub

Sub loadToSearchBy()
    cboSearchBy.AddItem "Name"
End Sub
