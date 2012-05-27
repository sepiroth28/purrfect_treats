VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewList 
   BackColor       =   &H80000007&
   Caption         =   "View List"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13005
   Icon            =   "frmViewList.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8700
   ScaleWidth      =   13005
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8385
      ScaleWidth      =   12705
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.ComboBox cboViewType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   960
         Width           =   3255
      End
      Begin MSComctlLib.ListView lsv_ViewList 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   12091
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin VB.Label lblLstOfAll 
         BackStyle       =   0  'Transparent
         Caption         =   "List of all Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   12480
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   12720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select type of List :"
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
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmViewList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboViewType_Click()
    If cboViewType.Text = "Items" Then
            loadToItems
            lblLstOfAll.Caption = "List of all" & " " & cboViewType.Text
    ElseIf cboViewType.Text = "Customers" Then
           loadToCustomers
           lblLstOfAll.Caption = "List of all" & " " & cboViewType.Text
    End If
End Sub

Private Sub Form_Load()
    Dim view_value As New Collection
    
    view_value.Add "Items"
    view_value.Add "Customers"
    
    LoadToCombo view_value, cboViewType
    loadToItems
    
End Sub

Sub loadToItems()
            lsv_ViewList.ColumnHeaders.Clear
            lsv_ViewList.ListItems.Clear
            
            Call setItemsDescriptionColumns(lsv_ViewList)
            Call loadAllItemsToListview(lsv_ViewList, "item_code")
            
            lsv_ViewList.ColumnHeaders(1).width = 1000
            lsv_ViewList.ColumnHeaders(2).width = 1500
            lsv_ViewList.ColumnHeaders(3).width = 4000
            lsv_ViewList.ColumnHeaders(4).width = 4000
            lsv_ViewList.ColumnHeaders(5).width = 1000
            lsv_ViewList.ColumnHeaders(5).Alignment = lvwColumnCenter
            lsv_ViewList.ColumnHeaders(6).width = 1000
            lsv_ViewList.ColumnHeaders(6).Alignment = lvwColumnRight
            lsv_ViewList.ColumnHeaders(7).width = 1700
            lsv_ViewList.ColumnHeaders(7).Alignment = lvwColumnCenter
            lsv_ViewList.ColumnHeaders(8).width = 5000
End Sub
Sub loadToCustomers()
            lsv_ViewList.ColumnHeaders.Clear
            lsv_ViewList.ListItems.Clear
            
            Call setCustomersColumns(lsv_ViewList)
            Call loadAllCustomersToListview(lsv_ViewList)
            
            lsv_ViewList.ColumnHeaders(1).width = 1500
            lsv_ViewList.ColumnHeaders(2).width = 5000
            lsv_ViewList.ColumnHeaders(3).width = 5000
            lsv_ViewList.ColumnHeaders(4).width = 2000
End Sub

Private Sub Form_Resize()
Picture1.width = Me.ScaleWidth - 250
Picture1.Height = Me.ScaleHeight - 200

End Sub

Private Sub Picture1_Resize()
lsv_ViewList.width = Picture1.ScaleWidth - 300
lsv_ViewList.Height = Picture1.ScaleHeight - 2000
End Sub
