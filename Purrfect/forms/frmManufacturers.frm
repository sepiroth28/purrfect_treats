VERSION 5.00
Begin VB.Form frmManufacturersForm 
   BackColor       =   &H8000000D&
   Caption         =   "Manage Manufacturers"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "frmManufacturers.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4455
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   60
      ScaleHeight     =   4305
      ScaleWidth      =   7125
      TabIndex        =   0
      Top             =   60
      Width           =   7155
      Begin VB.TextBox txManufacturersName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Tag             =   "*Manufacturers name"
         Top             =   1020
         Width           =   6555
      End
      Begin VB.TextBox txtManufacturersAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Tag             =   "*Item name"
         Top             =   1860
         Width           =   6555
      End
      Begin VB.TextBox txtManufacturersNumber 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Tag             =   "*Item price"
         Top             =   2760
         Width           =   2895
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
         Left            =   4140
         TabIndex        =   4
         Top             =   3420
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturers name"
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturers address"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   4275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturers phone number"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2460
         Width           =   2955
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
         Top             =   180
         Visible         =   0   'False
         Width           =   6555
      End
   End
End
Attribute VB_Name = "frmManufacturersForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_manufacturer As New manufacturers
Dim edit_manufacturer As New manufacturers

Private Sub cmdAddNewItem_Click()

If editManufacturer = True Then
    With edit_manufacturer
       .manufacturers_name = txManufacturersName.Text
       .manufacturers_add = txtManufacturersAdd.Text
       .manufacturers_number = txtManufacturersNumber.Text
       .update
        Call loadAllmanufacturersToListview(frmManageManufacturers.lsvManufacturers)
        Unload Me
    End With
Else
    With new_manufacturer
        .manufacturers_name = txManufacturersName.Text
        .manufacturers_add = txtManufacturersAdd.Text
        .manufacturers_number = txtManufacturersNumber.Text
        .insert
        Call loadAllmanufacturersToListview(frmManageManufacturers.lsvManufacturers)
        Unload Me
    End With
End If
End Sub

Private Sub Form_Load()
If editManufacturer = True Then
    Set edit_manufacturer = New manufacturers
    edit_manufacturer.load_manufacturers (edit_manufacturer_id)
    With edit_manufacturer
        txManufacturersName.Text = .manufacturers_name
        txtManufacturersNumber.Text = .manufacturers_number
        txtManufacturersAdd.Text = .manufacturers_add
    End With
End If
End Sub
