VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form dbuser_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATABASE USER"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   ClipControls    =   0   'False
   Icon            =   "dbuser_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hak Akses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   5055
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Program Akuntansi"
         Height          =   375
         Index           =   9
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Laporan"
         Height          =   375
         Index           =   8
         Left            =   3720
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Non Tiket"
         Height          =   375
         Index           =   7
         Left            =   3720
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Data Suplier"
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Data Client"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Penjualan Tiket"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Letter of Guarantee"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Stok Tiket"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Database Tiket"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hak Pengguna"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "HAPUS"
      DownPicture     =   "dbuser_frm.frx":1CCA
      Height          =   855
      Index           =   2
      Left            =   8760
      MouseIcon       =   "dbuser_frm.frx":2994
      MousePointer    =   99  'Custom
      Picture         =   "dbuser_frm.frx":2C9E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "EDIT"
      DownPicture     =   "dbuser_frm.frx":4620
      Height          =   855
      Index           =   1
      Left            =   6960
      MouseIcon       =   "dbuser_frm.frx":52EA
      MousePointer    =   99  'Custom
      Picture         =   "dbuser_frm.frx":55F4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "TAMBAH"
      DownPicture     =   "dbuser_frm.frx":6F76
      Height          =   855
      Index           =   0
      Left            =   5280
      MouseIcon       =   "dbuser_frm.frx":7C40
      MousePointer    =   99  'Custom
      Picture         =   "dbuser_frm.frx":7F4A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "dbuser_frm.frx":98CC
      Height          =   3855
      Left            =   5280
      OleObjectBlob   =   "dbuser_frm.frx":98E0
      TabIndex        =   10
      Top             =   120
      Width           =   5055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   " "
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D A T A B A S E   U S E R"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      Top             =   720
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   2400
      Picture         =   "dbuser_frm.frx":A2B3
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JABATAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   930
   End
End
Attribute VB_Name = "dbuser_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        Command1(0).Caption = "SIMPAN"
        Command1(1).Caption = "BATAL"
        Command1(2).Enabled = False
        buka
        kosong
        Call user_auto
        Text2.SetFocus
        Label2 = "t"
    Else
        simpan
    End If
Case 1
    If Text1.Text = main_form.Text1(0).Text Then
        x = MsgBox("Maaf transaksi tidak dapat diproses karena data user sedang dipakai...", vbOKOnly, "Validasi Data")
    Else
        If Command1(1).Caption = "BATAL" Then
            Data1.Refresh
            tutup
            isi
            Command1(0).Caption = "TAMBAH"
            Command1(1).Caption = "EDIT"
            Command1(2).Enabled = True
        Else
            If Not Data1.Recordset.BOF Then
                Command1(0).Caption = "SIMPAN"
                Command1(1).Caption = "BATAL"
                Command1(2).Enabled = False
                buka
                Text2.SetFocus
                Label2 = "e"
            Else
                x = MsgBox("data masih kosong, silahkan isi dahulu...", vbOKOnly, "blank data")
            End If
        End If
    End If
Case 2
    If Text1.Text = main_form.Text1(0).Text Then
        x = MsgBox("Maaf transaksi tidak dapat diproses karena data user sedang dipakai...", vbOKOnly, "Validasi Data")
    Else
        If Not Data1.Recordset.BOF Then
            x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
            If x = vbOK Then
                If Not Data1.Recordset.RecordCount = 1 Then
                    Data1.Recordset.Delete
                Else
                    x = MsgBox("Maaf proses tidak dapat dilaksanakan karena data tidak boleh kosong, terima kasih", vbOKOnly, "validasi data")
                End If
            End If
            Data1.Refresh
        End If
    End If
End Select
End Sub

Private Sub simpan()
Dim a As Boolean
a = False
If Text1 = "" Or Text2 = "" Or Combo1 = "" Or Text3 = "" Then
    x = MsgBox("data belum lengkap...", vbOKOnly, "Validasi Data")
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    ElseIf Combo1 = "" Then
        Combo1.SetFocus
    ElseIf Text3 = "" Then
        Text3.SetFocus
    End If
Else
    With Data1.Recordset
    If Label2 = "t" Then
        .MoveFirst
        Do While Not .EOF
            If !nama_user = Text2 Then
                x = MsgBox("data user sudah ada, silahkan isi user yg lain", vbOKOnly, "Validasi User")
                Text2.SetFocus
                .MoveLast
                a = True
            End If
            .MoveNext
        Loop
        If a = False Then
            .AddNew
            !id_user = Text1
            !nama_user = Text2
            !jabatan = Combo1
            !Password = Text3
            !hak_pengguna = Check1(0).Value
            !dbtiket = Check1(1).Value
            !stok_tiket = Check1(2).Value
            !lg = Check1(3).Value
            !jual_tiket = Check1(4).Value
            !client = Check1(5).Value
            !suplier = Check1(6).Value
            !ntiket = Check1(7).Value
            !laporan = Check1(8).Value
            !akuntansi = Check1(9).Value
            .Update
            Data1.Refresh
            isi
            tutup
            Command1(0).Caption = "TAMBAH"
            Command1(1).Caption = "EDIT"
            Command1(2).Enabled = True
        End If
    ElseIf Label2 = "e" Then
        .Edit
        !id_user = Text1
        !nama_user = Text2
        !jabatan = Combo1
        !Password = Text3
        !hak_pengguna = Check1(0).Value
        !dbtiket = Check1(1).Value
        !stok_tiket = Check1(2).Value
        !lg = Check1(3).Value
        !jual_tiket = Check1(4).Value
        !client = Check1(5).Value
        !suplier = Check1(6).Value
        !ntiket = Check1(7).Value
        !laporan = Check1(8).Value
        !akuntansi = Check1(9).Value
        .Update
        Data1.Refresh
        isi
        tutup
        Command1(0).Caption = "TAMBAH"
        Command1(1).Caption = "EDIT"
        Command1(2).Enabled = True
    End If
    End With
End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Command1(0).Caption = "SIMPAN" Then
    isi
End If
End Sub

Private Sub Form_Activate()
Command1(0).SetFocus
End Sub

Private Sub Form_Load()
Call dbuser
Data1.Refresh
Combo1.AddItem ("ADMINISTRATOR")
Combo1.AddItem ("OWNER")
Combo1.AddItem ("MARKETING")
Combo1.AddItem ("TIKETING")
Combo1.AddItem ("ACCOUNTING")
Combo1.AddItem ("OPERATOR")
Combo1.AddItem ("MESSANGER")
Combo1.AddItem ("USER")
Combo1.AddItem ("Lain-lain")
If Not Data1.Recordset.BOF Then
    isi
End If
tutup
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
main_form.Show
main_form.Enabled = True
End Sub

Private Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Text3.Enabled = False
Frame1.Enabled = False
End Sub

Private Sub buka()
'Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Text3.Enabled = True
Frame1.Enabled = True
End Sub

Private Sub kosong()
Text1 = ""
Text2 = ""
Combo1.ListIndex = 0
Text3 = ""
Check1(0).Value = 0
Check1(1).Value = 0
Check1(2).Value = 0
Check1(3).Value = 0
Check1(4).Value = 0
Check1(5).Value = 0
Check1(6).Value = 0
Check1(7).Value = 0
Check1(8).Value = 0
Check1(9).Value = 0
End Sub

Private Sub isi()
If Not Data1.Recordset.BOF Then
Text1 = Data1.Recordset!id_user
Text2 = Data1.Recordset!nama_user
Combo1 = Data1.Recordset!jabatan
Text3 = Data1.Recordset!Password
nilai_check1
End If
End Sub

Private Sub nilai_check1()
With Data1.Recordset
If !hak_pengguna = True Then
    Check1(0).Value = 1
Else
    Check1(0).Value = 0
End If
If !dbtiket = True Then
    Check1(1).Value = 1
Else
    Check1(1).Value = 0
End If
If !stok_tiket = True Then
    Check1(2).Value = 1
Else
    Check1(2).Value = 0
End If
If !lg = True Then
    Check1(3).Value = 1
Else
    Check1(3).Value = 0
End If
If !jual_tiket = True Then
    Check1(4).Value = 1
Else
    Check1(4).Value = 0
End If
If !client = True Then
    Check1(5).Value = 1
Else
    Check1(5).Value = 0
End If
If !suplier = True Then
    Check1(6).Value = 1
Else
    Check1(6).Value = 0
End If
If !ntiket = True Then
    Check1(7).Value = 1
Else
    Check1(7).Value = 0
End If
If !laporan = True Then
    Check1(8).Value = 1
Else
    Check1(8).Value = 0
End If
If !akuntansi = True Then
    Check1(9).Value = 1
Else
    Check1(9).Value = 0
End If
End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
