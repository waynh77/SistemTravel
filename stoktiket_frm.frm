VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form stoktiket_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERSEDIAAN TIKET"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   ClipControls    =   0   'False
   Icon            =   "stoktiket_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "S T O K   T I K E T"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Picture         =   "stoktiket_frm.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF8080&
      Caption         =   "KELUAR"
      DownPicture     =   "stoktiket_frm.frx":3994
      Height          =   975
      Left            =   4080
      MouseIcon       =   "stoktiket_frm.frx":465E
      MousePointer    =   99  'Custom
      Picture         =   "stoktiket_frm.frx":4968
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Keluar"
      DownPicture     =   "stoktiket_frm.frx":5632
      Height          =   855
      Left            =   9120
      MouseIcon       =   "stoktiket_frm.frx":62FC
      MousePointer    =   99  'Custom
      Picture         =   "stoktiket_frm.frx":6606
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Input No.Tiket Berurutan"
      DownPicture     =   "stoktiket_frm.frx":72D0
      Height          =   975
      Left            =   120
      MouseIcon       =   "stoktiket_frm.frx":7F9A
      MousePointer    =   99  'Custom
      Picture         =   "stoktiket_frm.frx":82A4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "stoktiket_frm.frx":9C26
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "stoktiket_frm.frx":9C3A
      TabIndex        =   1
      Top             =   4560
      Width           =   5295
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   7860
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   5055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "Combo5"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "HAPUS"
      DownPicture     =   "stoktiket_frm.frx":A60D
      Height          =   975
      Left            =   4080
      MouseIcon       =   "stoktiket_frm.frx":B2D7
      MousePointer    =   99  'Custom
      Picture         =   "stoktiket_frm.frx":B5E1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "EDIT"
      DownPicture     =   "stoktiket_frm.frx":CF63
      Height          =   975
      Left            =   2760
      MouseIcon       =   "stoktiket_frm.frx":DC2D
      MousePointer    =   99  'Custom
      Picture         =   "stoktiket_frm.frx":DF37
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "TAMBAH"
      DownPicture     =   "stoktiket_frm.frx":F8B9
      Height          =   975
      Left            =   1440
      MouseIcon       =   "stoktiket_frm.frx":10583
      MousePointer    =   99  'Custom
      Picture         =   "stoktiket_frm.frx":1088D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tangal Terima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   4
      Left            =   2160
      TabIndex        =   14
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Airlines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tiket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5760
      TabIndex        =   11
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tangal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Airlines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Tiket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   960
   End
End
Attribute VB_Name = "stoktiket_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Combo1_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
'End Sub

'Private Sub Combo2_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
'End Sub

'Private Sub Combo3_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
'End Sub

Private Sub Combo5_Change()
With dbtiket_frm.Data1.Recordset
.MoveFirst
Do While Not .EOF
    If !kode_maskapai = Combo5 Then
        Text2 = !nama_maskapai
        .MoveLast
    End If
    .MoveNext
Loop
End With
End Sub

Private Sub Combo5_Click()
With dbtiket_frm.Data1.Recordset
.MoveFirst
Do While Not .EOF
    If !kode_maskapai = Combo5 Then
        Text2 = !nama_maskapai
        .MoveLast
    End If
    .MoveNext
Loop
End With
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Command1_Click()
If Command1.Caption = "TAMBAH" Then
    tutup_cmd
    buka
    kosong
    isi_combo5
'    isi_combo
    Combo5.SetFocus
    Label2 = "t"
    Command4.Enabled = True
    DBGrid1.Enabled = False
Else
    simpan
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "EDIT" Then
    If Not Data1.Recordset.BOF Then
        buka
        Label2 = "e"
        'isi_combo5
        'isi_combo
        Combo5.SetFocus
        tutup_cmd
        Command4.Enabled = False
        DBGrid1.Enabled = False
    Else
        x = MsgBox("Data Kosong...", vbOKOnly, "Validasi Data")
    End If
Else
    tutup
    BUKA_CMD
    Data1.Refresh
    DBGrid1.Enabled = True
End If
End Sub

Private Sub Command3_Click()
If Data1.Recordset.BOF Then
    x = MsgBox("tidak ada data...", vbOKOnly, "Data Kosong")
Else
    x = MsgBox("apakah anda yakin?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
        Data1.Refresh
        Label3 = "Jumlah Stok Tiket = " & Data1.Recordset.RecordCount
        isi_list1
    End If
End If
End Sub

Private Sub Command4_Click()
Dim a, b, c As Single
Dim urutan As String
Dim hitung As Single
If Text1 = "" Or Combo5 = "" Then
    x = MsgBox("Data blm lengkap...", vbOKOnly, "Validasi Data")
    If Combo5 = "" Then
        Combo5.SetFocus
    ElseIf Text1 = "" Then
        Text1.SetFocus
    End If
Else
    a = InputBox("Masukan jumlah tiket", "Tiket Berurutan")
    If a <> "" Then
    If a > 1 Then
        b = Len(Text1)
        c = 0
        If b > 5 Then
        Do Until c = a
            With Data1.Recordset
                hitung = Val(Right(Text1, 5)) + c
                urutan = Mid(Text1, 1, b - 5) & hitung
                .AddNew
                !kode_maskapai = Combo5
                !no_tiket = urutan
                !tanggal_terima = Date
                !jam_terima = Time
                .Update
            End With
            c = c + 1
        Loop
            Data1.Refresh
            Label3 = "Jumlah Stok Tiket = " & Data1.Recordset.RecordCount
            isi
            tutup
            BUKA_CMD
            isi_list1
            DBGrid1.Enabled = True
        Else
            x = MsgBox("Jumlah kode tiket tidak boleh kurang dari 5 digit...", vbOKOnly, "Validasi Kode")
        End If
    Else
        x = MsgBox("jumlah tiket harus lebih dari 1", vbOKOnly, "validasi data")
    End If
    End If
End If
End Sub

Private Sub Command5_Click()
Me.Hide
main_form.Show
main_form.Enabled = True
Call isi_cekdat
End Sub

Private Sub Command6_Click()
Me.Hide
main_form.Show
main_form.Enabled = True
Call isi_cekdat
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command1.Caption <> "SIMPAN" And Not Data1.Recordset.BOF Then
    isi
End If
End Sub

Private Sub Form_Activate()
isi_combo5
If Data1.Recordset.BOF Then
    x = MsgBox("data msh kosong...", vbOKOnly, "Blank Data")
    buka
    kosong
    Combo5.SetFocus
    tutup_cmd
    Label2 = "t"
Else
    isi
    tutup
    BUKA_CMD
    Label3 = "Jumlah Stok Tiket = " & Data1.Recordset.RecordCount
    isi_list1
End If
End Sub

Private Sub Form_Load()
kosong
tutup
Label1(4) = Format(Date, "dd mmm yyyy")
Call stock_tiket
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
main_form.Show
main_form.Enabled = True
Call isi_cekdat
End Sub

Private Sub kosong()
Text1 = ""
Text2 = ""
Combo5 = ""
End Sub

Private Sub isi_combo5()
Dim a As String
Combo5.Clear
With dbtiket_frm
If Not .Data1.Recordset.BOF Then
.Data1.Recordset.MoveFirst
Do While Not .Data1.Recordset.EOF
    a = .Data1.Recordset!kode_maskapai
    Combo5.AddItem (a)
    .Data1.Recordset.MoveNext
Loop
End If
End With
End Sub

'Private Sub isi_combo()
'Dim tgl, bln As Byte
'Dim thn As Single
'tgl = 1
'bln = 1
'thn = 2007
'combo1.Clear
'Combo2.Clear
'Combo3.Clear
'Do Until tgl > 31
'    combo1.AddItem (tgl)
'    tgl = tgl + 1
'Loop
'Do Until bln > 12
'    Combo2.AddItem (bln)
'    bln = bln + 1
'Loop
'Do Until thn > 2050
'    Combo3.AddItem (thn)
'    thn = thn + 1
'Loop
'End Sub

Private Sub tutup()
Combo5.Enabled = False
Text1.Enabled = False
End Sub

Private Sub buka()
'combo1.Enabled = True
'Combo2.Enabled = True
'Combo3.Enabled = True
Combo5.Enabled = True
Text1.Enabled = True
End Sub

Private Sub isi()
With Data1.Recordset
If Not .BOF Then
    Combo5 = !kode_maskapai
    Text1 = !no_tiket
End If
End With
End Sub

Private Sub simpan()
'Dim tgl As Date
Dim a As Boolean
a = False
With Data1.Recordset
If Text1 = "" Or Combo5 = "" Then
    x = MsgBox("Data blm lengkap...", vbOKOnly, "Validasi Data")
    If Combo5 = "" Then
        Combo5.SetFocus
    ElseIf Text1 = "" Then
        Text1.SetFocus
    End If
Else
    If Label2 = "t" Then
        If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !kode_maskapai = Combo5 And !no_tiket = Text1 Then
                x = MsgBox("data tiket sudah ada, silahkan isi data tiket yg lain...", vbOKOnly, "Validasi Tiket")
                Combo5.SetFocus
                a = True
                .MoveLast
            End If
            .MoveNext
        Loop
        End If
        If a = False Then
            .AddNew
            !kode_maskapai = Combo5
            !no_tiket = Text1
            !tanggal_terima = Date
            !jam_terima = Time
            .Update
            Data1.Refresh
            Label3 = "Jumlah Stok Tiket = " & Data1.Recordset.RecordCount
            isi
            tutup
            BUKA_CMD
            isi_list1
            DBGrid1.Enabled = True
        End If
    Else
        .Edit
'        tgl = combo1 & "/" & Combo2 & "/" & Combo3
        !kode_maskapai = Combo5
        !no_tiket = Text1
        !tanggal_terima = Date
        !jam_terima = Time
        .Update
        Data1.Refresh
        isi_list1
        isi
        tutup
        BUKA_CMD
        DBGrid1.Enabled = True
    End If
End If
End With
End Sub

Private Sub tutup_cmd()
Command1.Caption = "SIMPAN"
Command2.Caption = "BATAL"
Command3.Enabled = False
Command4.Enabled = True
End Sub

Private Sub BUKA_CMD()
Command1.Caption = "TAMBAH"
Command2.Caption = "EDIT"
Command3.Enabled = True
Command4.Enabled = False
End Sub

Private Sub isi_list1()
Dim jml As Single
List1.Clear
If Not Data1.Recordset.BOF Then
    With dbtiket_frm.Data1.Recordset
    .MoveFirst
    Do While Not .EOF
        List1.AddItem ("Airlines :" & !nama_maskapai)
        Data1.Recordset.MoveFirst
        jml = 0
        Do While Not Data1.Recordset.EOF
            If Data1.Recordset!kode_maskapai = !kode_maskapai Then
                List1.AddItem ("   No.Tiket : " & Data1.Recordset!no_tiket & "  Tgl terima : " & Format(Data1.Recordset!tanggal_terima, "dd/mmm/yyyy"))
                jml = jml + 1
            End If
            Data1.Recordset.MoveNext
        Loop
        List1.AddItem ("")
        List1.AddItem ("               ===> Sub Total : " & jml & " Tiket")
        List1.AddItem ("----------------------------------------------------------------------------------")
        List1.AddItem ("")
        .MoveNext
    Loop
    End With
End If
Data1.Refresh
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
