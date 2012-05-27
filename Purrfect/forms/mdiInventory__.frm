VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiInventory 
   BackColor       =   &H8000000C&
   Caption         =   "Inventory System version 1.0"
   ClientHeight    =   2655
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16290
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   16290
      _ExtentX        =   28734
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16290
      _ExtentX        =   28734
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Menu mnuDataMaintenance 
      Caption         =   "&Data Maintenance"
      Begin VB.Menu mnuCustomer 
         Caption         =   "&Customer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSepManage0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManufacturer 
         Caption         =   "&Manufacturer"
      End
      Begin VB.Menu mnuSepManage1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProduct 
         Caption         =   "&Product"
      End
      Begin VB.Menu mnuCategory 
         Caption         =   "Category"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuStockIn 
         Caption         =   "Stock-&In"
      End
      Begin VB.Menu mnuStockOut 
         Caption         =   "Stock-&Out"
      End
      Begin VB.Menu mnuSepTransaction0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItems 
         Caption         =   "Items"
      End
   End
   Begin VB.Menu mnuInquiry 
      Caption         =   "&Inquiry"
      Begin VB.Menu mnuCustomers 
         Caption         =   "List of &Customers"
      End
      Begin VB.Menu mnuSepInquiry0 
         Caption         =   "-"
      End
      Begin VB.Menu mnulstStocks 
         Caption         =   "List of &Stock"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuCustomersReport 
         Caption         =   "Customer's Report"
      End
   End
End
Attribute VB_Name = "mdiInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCustomers_Click()
    frmlstCustomer.Show vbModal
End Sub

