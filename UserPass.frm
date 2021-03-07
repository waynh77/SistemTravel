VERSION 5.00
Begin VB.Form UserPass 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALPUKAT (Aplikasi Penjualan & Akuntansi Travel)"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "UserPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "KELUAR"
      DownPicture     =   "UserPass.frx":08CA
      Height          =   855
      Left            =   2520
      MouseIcon       =   "UserPass.frx":1594
      MousePointer    =   99  'Custom
      Picture         =   "UserPass.frx":189E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "PROSES"
      DownPicture     =   "UserPass.frx":3220
      Height          =   855
      Left            =   120
      MouseIcon       =   "UserPass.frx":3EEA
      MousePointer    =   99  'Custom
      Picture         =   "UserPass.frx":41F4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "UserPass.frx":5B76
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "UserPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim id, jab, hak As String
Dim a As Boolean
With main_form
    .Text1(0) = ""
    .Text1(1) = ""
    .Text1(2) = ""
End With
With Data1.Recordset
Data1.Refresh
.MoveFirst
a = False
Do While Not .EOF
    If Text1 = !nama_user And Text2 = !Password Then
    id = !id_user
    jab = !jabatan
        With main_form
            .Text1(0) = id
            .Text1(1) = Text1
            .Text1(2) = jab
'            .Show
            WaynhSoft_frm.Show
            .Enabled = True
        End With
        Text1 = ""
        Text2 = ""
        Me.Hide
        a = True
        If !hak_pengguna = -1 Then
            main_form.admin_mnu.Visible = True
            stoktiket_frm.Command3.Visible = True
            stoktiket_frm.Command6.Visible = False
            client_frm.Command1(2).Visible = True
            dbpartner_frm.Command1(2).Visible = True
            client_frm.Command1(3).Visible = False
            dbpartner_frm.Command1(3).Visible = False
            main_form.Command2.Visible = True
        Else
            main_form.Command2.Visible = False
            main_form.admin_mnu.Visible = False
            stoktiket_frm.Command3.Visible = False
            stoktiket_frm.Command6.Visible = True
            client_frm.Command1(2).Visible = False
            dbpartner_frm.Command1(2).Visible = False
            client_frm.Command1(3).Visible = True
            dbpartner_frm.Command1(3).Visible = True
        End If
        If !dbtiket = -1 Then
            main_form.dbtiket_mnu.Visible = True
            main_form.Toolbar1.Buttons(1).Visible = True
        Else
            main_form.dbtiket_mnu.Visible = False
            main_form.Toolbar1.Buttons(1).Visible = False
        End If
        If !stok_tiket = -1 Then
            main_form.stok_mnu.Visible = True
            main_form.Toolbar1.Buttons(2).Visible = True
        Else
            main_form.stok_mnu.Visible = False
            main_form.Toolbar1.Buttons(2).Visible = False
        End If
        If !lg = -1 Then
            main_form.LG_mnu.Visible = True
            main_form.Toolbar1.Buttons(5).Visible = True
        Else
            main_form.LG_mnu.Visible = False
            main_form.Toolbar1.Buttons(5).Visible = False
        End If
        If !jual_tiket = -1 Then
            main_form.pesan_mnu.Visible = True
            main_form.Toolbar1.Buttons(6).Visible = True
        Else
            main_form.pesan_mnu.Visible = False
            main_form.Toolbar1.Buttons(6).Visible = False
        End If
        If !client = -1 Then
            main_form.dbClient_mnu.Visible = True
            main_form.Toolbar1.Buttons(3).Visible = True
        Else
            main_form.dbClient_mnu.Visible = False
            main_form.Toolbar1.Buttons(3).Visible = False
        End If
        If !suplier = -1 Then
            main_form.partner_mnu.Visible = True
            main_form.Toolbar1.Buttons(4).Visible = True
        Else
            main_form.partner_mnu.Visible = False
            main_form.Toolbar1.Buttons(4).Visible = False
        End If
        If !ntiket = -1 Then
            main_form.ntiket_mnu.Visible = True
        Else
            main_form.ntiket_mnu.Visible = False
        End If
        If !laporan = -1 Then
            main_form.lap_mnu.Visible = True
            InvNt_frm.lap_nt.Visible = True
        Else
            main_form.lap_mnu.Visible = False
            InvNt_frm.lap_nt.Visible = False
        End If
        If !akuntansi = -1 Then
            main_form.Akt_mnu.Visible = True
        Else
            main_form.Akt_mnu.Visible = False
        End If
        .MoveLast
    End If
    .MoveNext
Loop
If a = False Then
    x = MsgBox("User atau Password Salah", vbOKOnly, "Invalid User or Password")
    Text2 = ""
    Text2.SetFocus
End If
End With
End Sub

Private Sub Command2_Click()
x = MsgBox("Apakah anda yakin ingin keluar dari aplikasi?", vbYesNo, "EXIT")
If x = vbYes Then
    End
End If
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Text1 = ""
Text2 = ""
Text1.MaxLength = 30
Text2.MaxLength = 30
Call dbuser_awal
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub
