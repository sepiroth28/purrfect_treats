VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   5145
   ClientTop       =   1440
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   11535
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save item"
      Height          =   735
      Left            =   7800
      TabIndex        =   6
      Top             =   7200
      Width           =   2175
   End
   Begin MSComctlLib.ListView lsvManufacturers 
      Height          =   2175
      Left            =   600
      TabIndex        =   5
      Top             =   6720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "itrem_code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "item_name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "item_description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "item_qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "item_price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "item_status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   675
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

'================== SAMPLE RETRIEVING ITEMS USING COLLECTION ==============================================================================================
'declaring an itemcollection
Dim collection As ItemCollection

'getting all items in collection
Set collection = getAllItemsCollection
Dim Item As New items

'looping to all items in the collection
For Each Item In collection
    'item represents each item(class) in the collection
    
    MsgBox Item.item_code
Next
'================================================================================================================

End Sub

Private Sub Command2_Click()

'================== SAMPLE RETRIEVING ITEMS USING ADODB.RECORDSET ==============================================================================================
'declaring an adodb.recordset

Dim data As ADODB.Recordset

'getting all items in recrodset
Set data = getAllItems

'looping to all records in the recordset
Do Until data.EOF
    'data represents each record in the recordset
    MsgBox data.Fields("item_code").Value
    
'move to the next record
data.MoveNext
Loop
End Sub

Private Sub Command3_Click()
Dim Item As New items
Item.item_code = "1"
Item.date_added = "2011-9-5"
Item.date_modified = "2011-9-5"
Item.item_qty = 10
Item.manufacturers_id = 1
Item.reorder_point = 10

Item.item_name = "products 1"
Item.item_description = "product descriptions"
Item.item_status = 1
Item.unit_of_measure = "pcs"
Item.image = "image1.jpg"

Item.insert

End Sub

Private Sub Command4_Click()
frmCustomer.Show
Unload Me
End Sub

Private Sub Command5_Click()
Dim temp As New items
    
    With temp
        .item_code = "123"
        .item_name = "hog feeds"
        .item_description = "hog feeds for starter"
        .item_price = 100.25
        .item_qty = 100
        .item_status = 1
        .reorder_point = 50
        .unit_of_measure = "kilo"
        .manufacturers_id = 10
        .insert
    End With
End Sub

Private Sub Command6_Click()
    frmlstCustomer.Show
End Sub

Private Sub Form_Load()
Call loadAllItemsToListview(ListView1)
Call setManufacturersColumns(lsvManufacturers)

End Sub

Private Sub ListView1_DblClick()
MsgBox ListView1.SelectedItem.Index
End Sub
