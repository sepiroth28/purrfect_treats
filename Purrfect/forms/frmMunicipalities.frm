VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMunicipalities 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Municipalities"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5565
   Icon            =   "frmMunicipalities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   7635
      Left            =   60
      ScaleHeight     =   7605
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   60
      Width           =   5415
      Begin VB.CommandButton cmdClear 
         Caption         =   "CLEAR"
         Height          =   555
         Left            =   1920
         TabIndex        =   8
         Top             =   6900
         Width           =   1455
      End
      Begin VB.TextBox txtTrackingPrice 
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
         Left            =   3120
         TabIndex        =   5
         Top             =   5940
         Width           =   1995
      End
      Begin VB.TextBox txtMunicipalName 
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
         Left            =   300
         TabIndex        =   4
         Top             =   5940
         Width           =   2715
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   555
         Left            =   3480
         TabIndex        =   3
         Top             =   6900
         Width           =   1635
      End
      Begin MSComctlLib.ListView lsvMunicipalities 
         Height          =   5055
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   8916
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Municipal name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Tracking price"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tracking price"
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   6420
         Width           =   1875
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Municipal Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   6420
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MUNICIPALITIES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmMunicipalities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim municipality_id As Integer
Dim edit_municipal As Boolean

Private Sub cmdClear_Click()
municipality_id = 0
edit_municipal = False
txtMunicipalName.Text = ""
txtTrackingPrice.Text = ""
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSave_Click()
Dim sql As String

'municipal_id, municipal_name, tracking_price
If edit_municipal = True Then
    sql = "UPDATE municipalities SET municipal_name = '" & txtMunicipalName.Text & "',tracking_price=" & Val(txtTrackingPrice.Text) & " WHERE municipal_id = " & municipality_id
    db.execute sql
Else
    sql = "INSERT INTO municipal_name VALUES(null,'" & txtMunicipalName.Text & "'," & Val(txtTrackingPrice.Text) & ")"
    db.execute sql
End If

MsgBox "Successfully saved"
Call cmdClear_Click
Call loadMunicipal
End Sub

Private Sub Form_Load()
Call loadMunicipal
municipality_id = 0
edit_municipal = False

End Sub

Sub loadMunicipal()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim list As ListItem

sql = "SELECT * FROM `municipalities`;"
Set rs = db.execute(sql)
lsvMunicipalities.ListItems.Clear
On Error Resume Next
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        Set list = lsvMunicipalities.ListItems.Add(, , rs.Fields(0).Value)
        list.SubItems(1) = rs.Fields(1).Value
        list.SubItems(2) = rs.Fields(2).Value
    rs.MoveNext
    Loop
End If

End Sub

Private Sub lsvMunicipalities_DblClick()
edit_municipal = True
municipality_id = Val(lsvMunicipalities.SelectedItem.Text)
Call editMunicipal
End Sub

Sub editMunicipal()
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT * FROM `municipalities` WHERE municipal_id = " & municipality_id
Set rs = db.execute(sql)

On Error Resume Next
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        txtMunicipalName.Text = rs.Fields(1).Value
        txtTrackingPrice.Text = rs.Fields(2).Value
    rs.MoveNext
    Loop
End If


End Sub

Private Sub txtTrackingPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdSave_Click
End If
End Sub
