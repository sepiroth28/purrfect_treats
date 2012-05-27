VERSION 5.00
Begin VB.Form frmLogIn 
   BackColor       =   &H00E1A00B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   12
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogIn.frx":058A
   ScaleHeight     =   4500
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2940
      PasswordChar    =   "="
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2940
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   7620
      ScaleHeight     =   2505
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Image imgKey 
         Height          =   1695
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1905
         TabIndex        =   4
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Image imgLogin 
      Height          =   615
      Left            =   4920
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Image imgCancel 
      Height          =   615
      Left            =   2880
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   1875
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cntr As Integer
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOk_Click()
    checkAccount
End Sub
Sub checkAccount()
    
    Dim checkuser As New User_Account
    Dim mvarusername, mvarpassword As String
    Dim check_useraccount As Boolean
    Dim check_user_type As String
    
    mvarusername = txtusername.Text
    mvarpassword = txtPassword.Text
    
    On Error Resume Next
    check_useraccount = checkuser.Check_UserAcount(mvarusername, mvarpassword)
    
    If check_useraccount = True Then
    
        check_user_type = checkuser.Check_UserType(mvarusername)
            
            If check_user_type = "admin" Then
                MsgBox "Welcome " & mvarusername, vbInformation, "Welcome"
                activeUser.loadUserAccount mvarusername
                Call grantUserPreviliges(activeUser.username)
                
               ' mdi_Inventory.stbNutrimart.Panels(2).Text = mvarusername
                mdi_Inventory.Show
             
            ElseIf check_user_type = "user" Then
                MsgBox "Welcome " & mvarusername, vbInformation, "Login"
                activeUser.loadUserAccount mvarusername
                 Call grantUserPreviliges(activeUser.username)
                
               ' mdi_Inventory.stbNutrimart.Panels(1).Text = "UserName"
               ' mdi_Inventory.stbNutrimart.Panels(2).Text = mvarusername
                mdi_Inventory.Show
               
            End If
                Unload Me
    Else
        prompt
    End If

End Sub

Sub prompt()
For Each cnt In frmLogIn
    If TypeOf cnt Is TextBox Then
        cntr = cntr + 1
        If cntr = 1 Then
            MsgBox "Access Denied!..You only have 2 attempts remaining", vbInformation, "Warning!"
        ElseIf cntr = 2 Then
            MsgBox "Access Denied..!You only have 1 attempt remaining", vbInformation, "Warning!"
        Else
            MsgBox "Access failed in 3 attempts...System will now close!", vbExclamation, "Error Log-in"
            Unload Me
            End
        End If
    
        cnt.SetFocus
'        HLText cnt
        Exit Sub
    End If
Next cnt
End Sub

Private Sub Form_Activate()
     txtusername.SetFocus
End Sub

Private Sub Form_Load()
    'imgKey.Picture = LoadPicture(App.Path & "\images\keys1.jpg")
End Sub


Private Sub imgCancel_Click()
    Call cmdCancel_Click
End Sub

Private Sub imgLogin_Click()
    Call cmdOk_Click
End Sub
