VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Description"
   ClientHeight    =   7335
   ClientLeft      =   4470
   ClientTop       =   1170
   ClientWidth     =   14040
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   14040
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
      Left            =   12000
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Text            =   "search here"
      Top             =   240
      Width           =   4095
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin MSComctlLib.ListView lsvItemsList 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Caption         =   "Search by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "frmItemsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    editmode = False
    frmItemsDescription.Show vbModal
    
End Sub

Private Sub lsvItemList_DblClick()
    editmode = True
    
    With frmItemsDescription
    
    
    End With
    
End Sub

Private Sub Form_Load()
    searchytpe
    
    Call setItemsDescriptionColumns(lsvItemsList)
    Call loadAllItemsDescriptionToListview(lsvItemsList)
        
End Sub


Sub searchytpe()
    With cboSearch
        .AddItem "Item Code"
        .AddItem "Item Name"
    End With
    
End Sub
