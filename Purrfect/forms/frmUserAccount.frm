VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserAccount 
   BackColor       =   &H00E1A00B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Account Form"
   ClientHeight    =   5040
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11025
   Icon            =   "frmUserAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   60
      ScaleHeight     =   4785
      ScaleWidth      =   10845
      TabIndex        =   0
      Top             =   120
      Width           =   10875
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "Change Password"
         Height          =   495
         Left            =   2640
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin MSComctlLib.ListView lsvPrevileges 
         Height          =   4635
         Left            =   4680
         TabIndex        =   9
         Top             =   60
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8176
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.ComboBox cbouser_type 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "="
         TabIndex        =   5
         Top             =   2160
         Width           =   2475
      End
      Begin VB.TextBox txtusername 
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
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
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
         Left            =   1920
         TabIndex        =   1
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Type"
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
         TabIndex        =   7
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   900
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name "
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
         Top             =   840
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim useraccount As New User_Account

Sub loadAssignedPrevileges()
Dim rs As New ADODB.Recordset
Dim sql As String
Dim list As ListItem

sql = "SELECT p.id,p.previleges,up.`status` FROM user_previleges up inner join previleges p on up.previleges = p.id where username = '" & useraccount.username & "'"
Set rs = db.execute(sql)
lsvPrevileges.ListItems.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        Set list = lsvPrevileges.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields(1).Value
        If rs.Fields(2).Value = True Then
            list.Checked = True
        Else
            list.Checked = False
        End If
    rs.MoveNext
    Loop
End If
End Sub

Sub loadPrevileges()
Dim rs As New ADODB.Recordset
Dim sql As String
Dim list As ListItem

sql = "SELECT id,previleges FROM previleges"
Set rs = db.execute(sql)
lsvPrevileges.ListItems.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        Set list = lsvPrevileges.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields(1).Value
    rs.MoveNext
    Loop
End If
End Sub
Private Sub cmdAddNewItem_Click()
    
    saveData
    Unload Me
    Call loadAllUserAccountToListview(frmManageUserAccount.lsvUserAccount)
    
End Sub


Sub saveData()
    Dim useraccount As New User_Account
    Dim mvarusername, mvarpassword, mvaruser_type As String
    
    mvarusername = txtusername.Text
    mvarpassword = txtPassword.Text
    mvaruser_type = cbouser_type.Text
   
        If editmode = True Then
         With useraccount
            .username = activeaseraccount_name
            '.Password = mvarpassword
            .user_type = mvaruser_type
            .UpdateUserAccount
            Call savePrevileges(activeaseraccount_name)
         End With
         
        Else
            With useraccount
                .username = mvarusername
                .Password = mvarpassword
                .user_type = mvaruser_type
                .SaveUserAccount
              Call savePrevileges(mvarusername)
            End With
           
        End If
        
End Sub
Sub savePrevileges(username)
Dim list As ListItem
Dim prev_status As Integer
Dim prev_id As Integer

db.execute "DELETE FROM user_previleges WHERE username = '" & username & "'"

For Each list In lsvPrevileges.ListItems
    If list.Checked = True Then
        prev_status = 1
    Else
        prev_status = 0
    End If
    prev_id = Val(list.Text)
    db.execute "INSERT INTO user_previleges VALUES (null,'" & username & "'," & prev_id & "," & prev_status & ")"
Next
End Sub


Private Sub cmdChangePassword_Click()
useraccount.loadUserAccount (activeaseraccount_name)
useraccount.changePassword (txtPassword.Text)
End Sub

Private Sub Form_Load()

    With cbouser_type
        .AddItem ADMIN
        .AddItem USER
    End With
    
    With useraccount
        If editmode = True Then
                Call .loadUserAccount(activeaseraccount_name)
                txtusername.Text = activeaseraccount_name
                'txtpassword.Text = .Password
                cbouser_type.Text = .user_type
                Call loadAssignedPrevileges
        Else
                Call loadPrevileges
        End If
      
    End With
   
End Sub
