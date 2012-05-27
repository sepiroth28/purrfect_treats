VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageUserAccount 
   BackColor       =   &H80000007&
   Caption         =   "Manage User Account"
   ClientHeight    =   8370
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   9405
   Icon            =   "frmManageUserAccount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8115
      Left            =   120
      ScaleHeight     =   8085
      ScaleWidth      =   9165
      TabIndex        =   0
      Top             =   120
      Width           =   9195
      Begin VB.CommandButton cmdAddNewItem 
         Caption         =   "ADD NEW USER"
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
         Left            =   6240
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtSearchItemCode 
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
         Left            =   180
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
      End
      Begin MSComctlLib.ListView lsvUserAccount 
         Height          =   6135
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   10821
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search User Name"
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
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage User Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   3555
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9000
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Menu mnuManageUser 
      Caption         =   "Manage User Account"
      Begin VB.Menu mnudelete 
         Caption         =   "delete"
      End
   End
End
Attribute VB_Name = "frmManageUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNewItem_Click()
    editmode = False
    frmUserAccount.Show 1
End Sub

Private Sub Form_Load()

    loaduserToListView

End Sub


Sub loaduserToListView()
    
    Call setUserAccountColumn(lsvUserAccount)
    
    With lsvUserAccount
        .ColumnHeaders(1).width = 3500
        .ColumnHeaders(2).width = 0
        .ColumnHeaders(3).width = 1800
    End With
    Call loadAllUserAccountToListview(lsvUserAccount)
End Sub

Private Sub lsvUserAccount_DblClick()
    editmode = True
       
    activeaseraccount_name = lsvUserAccount.SelectedItem.Text
    frmUserAccount.Show 1
End Sub

Private Sub lsvUserAccount_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuManageUser
    End If

End Sub

Private Sub mnuDelete_Click()
    If MsgBox("Are you sure you want to delete?", vbYesNo, "Delete user account") = vbYes Then
        DeleteUserAccount (lsvUserAccount.SelectedItem.Text)
        Call loadAllUserAccountToListview(lsvUserAccount)
    End If
End Sub
