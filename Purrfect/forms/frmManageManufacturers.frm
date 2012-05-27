VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageManufacturers 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Manufacturers"
   ClientHeight    =   7500
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   11655
   Icon            =   "frmManageManufacturers.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7395
      Left            =   60
      ScaleHeight     =   7365
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   60
      Width           =   11535
      Begin VB.TextBox txtSearchManufacturersName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CommandButton cmdAddNewManufacturer 
         Caption         =   "ADD NEW MANUFACTURER"
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
         Left            =   7920
         TabIndex        =   1
         Top             =   900
         Width           =   3375
      End
      Begin MSComctlLib.ListView lsvManufacturers 
         Height          =   5535
         Left            =   180
         TabIndex        =   3
         Top             =   1740
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   9763
         View            =   3
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Manufacturers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   11340
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Manufactures name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   5055
      End
   End
   Begin VB.Menu mnu_manufacturers_menu 
      Caption         =   "ManufacturersMenu"
      Begin VB.Menu mnu_delete_manufacturers 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmManageManufacturers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNewManufacturer_Click()
frmManufacturersForm.Show 1
End Sub

Private Sub Form_Load()

Call setManufacturersColumns(lsvManufacturers)
lsvManufacturers.ColumnHeaders(1).width = 0
lsvManufacturers.ColumnHeaders(2).width = 4000
lsvManufacturers.ColumnHeaders(3).width = 4000
Call loadAllmanufacturersToListview(lsvManufacturers)
End Sub

Private Sub lsvManufacturers_DblClick()
editManufacturer = True
edit_manufacturer_id = Val(lsvManufacturers.SelectedItem.Text)
frmManufacturersForm.Show 1
End Sub

Private Sub lsvManufacturers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Me.PopupMenu mnu_manufacturers_menu
End If
End Sub

Private Sub mnu_delete_manufacturers_Click()
 If MsgBox("Are you sure you want to delete?", vbYesNo, "Delete Manufacturers") = vbYes Then
   Call deleteManufacturers(Val(lsvManufacturers.SelectedItem.Text))
   Call loadAllmanufacturersToListview(lsvManufacturers)
End If
End Sub

