VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Lapjual2_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5040
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Lapjual2_frm.frx":0000
      Height          =   4575
      Left            =   4560
      OleObjectBlob   =   "Lapjual2_frm.frx":0014
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   10575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Keluar"
      DownPicture     =   "Lapjual2_frm.frx":09FB
      Height          =   855
      Index           =   1
      Left            =   2280
      Picture         =   "Lapjual2_frm.frx":16C5
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Cetak Laporan"
      DownPicture     =   "Lapjual2_frm.frx":238F
      Height          =   855
      Index           =   0
      Left            =   240
      Picture         =   "Lapjual2_frm.frx":3059
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58064897
      CurrentDate     =   39500
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58064897
      CurrentDate     =   39500
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   2520
      Width           =   3855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL TIKET = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   7
      Left            =   1680
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL TIKET = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   120
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AIRLINES/CLIENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SORT BY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT DATA LAPORAN PENJUALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3900
   End
End
Attribute VB_Name = "Lapjual2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "Harian..."
Combo1.AddItem "Periodik..."
Combo1.ListIndex = 0
End Sub

Sub isi_cmb2()
Combo2.Clear
Combo2.AddItem "Client..."
Combo2.AddItem "Airlines..."
Combo2.ListIndex = 0
End Sub

Sub isi_cmb3()
Combo3.Clear
Combo3.AddItem "ALL..."
If Combo2 = "Client..." Then
    Label1(4).Caption = "CLIENT"
    With client_frm.Data1.Recordset
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                Combo3.AddItem !COMPANY
                .MoveNext
            Loop
            Combo3.ListIndex = 0
        End If
    End With
Else
    Label1(4).Caption = "AIRLINES"
    With dbtiket_frm.Data1.Recordset
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                Combo3.AddItem !nama_maskapai
                .MoveNext
            Loop
            Combo3.ListIndex = 0
        End If
    End With
End If
Combo3 = "ALL..."
End Sub

Sub isi_data()
Dim a, b, c As String
Dim d, e As Date
a = Combo1
b = Combo2
c = Combo3
d = Format(DTPicker1, "m/d/yyyy")
e = Format(DTPicker2, "m/d/yyyy")
If a = "Harian..." And b = "Client..." And c = "ALL..." Then
    Data1.RecordSource = "select tgl_inv,company,contact_person,nama_maskapai,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where cdate(tgl_inv) = '" & d & "' and db_invoice.kode_maskapai= db_maskapai.kode_maskapai order by currency, tgl_inv asc"
ElseIf a = "Harian..." And b = "Client..." And c <> "ALL..." Then
    Data1.RecordSource = "select tgl_inv,company,contact_person,nama_maskapai,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where cdate(tgl_inv) = '" & d & "' and db_invoice.kode_maskapai= db_maskapai.kode_maskapai and company = '" & c & "'order by currency, tgl_inv asc"
ElseIf a = "Periodik..." And b = "Client..." And c = "ALL..." Then
    Data1.RecordSource = "select tgl_inv,company,contact_person,nama_maskapai,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where  cdate(tgl_inv) >= '" & d & "' and cdate(tgl_inv)<= '" & e & "' and db_invoice.kode_maskapai= db_maskapai.kode_maskapai order by currency, tgl_inv asc"
ElseIf a = "Periodik..." And b = "Client..." And c <> "ALL..." Then
    Data1.RecordSource = "select tgl_inv,company,contact_person,nama_maskapai,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where  cdate(tgl_inv) >= '" & d & "' and cdate(tgl_inv)<= '" & e & "' and db_invoice.kode_maskapai= db_maskapai.kode_maskapai and company = '" & c & "'order by currency, tgl_inv asc"

ElseIf a = "Harian..." And b = "Airlines..." And c = "ALL..." Then
    Data1.RecordSource = "select tgl_inv,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where cdate(tgl_inv) = '" & d & "' and db_invoice.kode_maskapai= db_maskapai.kode_maskapai order by currency, tgl_inv asc"
ElseIf a = "Harian..." And b = "Airlines..." And c <> "ALL..." Then
    Data1.RecordSource = "select tgl_inv,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where cdate(tgl_inv) = '" & d & "' and db_invoice.kode_maskapai= db_maskapai.kode_maskapai and nama_maskapai = '" & c & "'order by currency, tgl_inv asc"
ElseIf a = "Periodik..." And b = "Airlines..." And c = "ALL..." Then
    Data1.RecordSource = "select tgl_inv,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where  db_invoice.kode_maskapai= db_maskapai.kode_maskapai and cdate(tgl_inv) >= '" & d & "' and cdate(tgl_inv)<= '" & e & "'order by currency, tgl_inv asc"
ElseIf a = "Periodik..." And b = "Airlines..." And c <> "ALL..." Then
    Data1.RecordSource = "select tgl_inv,no_invoice,no_lg,no_tiket,currency,hrg_tiket from db_invoice,db_maskapai where  db_invoice.kode_maskapai= db_maskapai.kode_maskapai and cdate(tgl_inv) >= '" & d & "' and cdate(tgl_inv)<= '" & e & "' and nama_maskapai = '" & c & "'order by currency, tgl_inv asc"
End If
Data1.Refresh
End Sub

Sub kosong()
Combo1.Clear
Combo2.Clear
Combo3.Clear
End Sub

Private Sub Combo1_Click()
If Combo1 = "Harian..." Then
    Label1(2).Visible = False
    DTPicker2.Visible = False
Else
    Label1(2).Visible = True
    DTPicker2.Visible = True
End If
isi_data
'Label1(7).Caption = Data1.Recordset.RecordCount
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
isi_cmb3
isi_data

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Combo3_Click()
isi_data
Label1(7).Caption = Data1.Recordset.RecordCount
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
'KeyAscii = 0

End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
Dim a, b, c As String
Dim d, e As Date
a = Combo1
b = Combo2
c = Combo3
d = DTPicker1
e = DTPicker2
If main_form.Text1(2) = "ADMINISTRATOR" Then
    If a = "Harian..." And b = "Client..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-client admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Harian..." And b = "Client..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-client admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ") And {db_invoice.company} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Client..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-client admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Client..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-client admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") And {db_invoice.company} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    
    ElseIf a = "Harian..." And b = "Airlines..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-airlines admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Harian..." And b = "Airlines..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-airlines admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ") And {db_maskapai.nama_maskapai} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Airlines..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-airlines admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Airlines..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-airlines admin.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") And {db_maskapai.nama_maskapai} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    End If
Else
    If a = "Harian..." And b = "Client..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-CLIENT4.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Harian..." And b = "Client..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-CLIENT4.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ") And {db_invoice.company} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Client..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-CLIENT4.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Client..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-CLIENT4.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") And {db_invoice.company} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    
    ElseIf a = "Harian..." And b = "Airlines..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-AIRLINES3.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Harian..." And b = "Airlines..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-AIRLINES3.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}= date(" & Format(DTPicker1, "yyyy,m,d") & ") And {db_maskapai.nama_maskapai} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Airlines..." And c = "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-AIRLINES3.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    ElseIf a = "Periodik..." And b = "Airlines..." And c <> "ALL..." Then
        CrystalReport1.ReportFileName = App.Path & "\laporan penjualan-AIRLINES3.rpt"
        CrystalReport1.SelectionFormula = "{db_invoice.tgl_inv}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {db_invoice.tgl_inv}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") And {db_maskapai.nama_maskapai} = '" & c & "'"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    End If
End If
Case 1
    Unload Me
    main_form.Enabled = True
    main_form.Show
End Select
End Sub



Private Sub DTPicker1_Change()
isi_data
Label1(7).Caption = Data1.Recordset.RecordCount
End Sub



Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub DTPicker2_Change()
If DTPicker2.Visible = True Then
    If DTPicker2 < DTPicker1 Then
        MsgBox "Tanggal Periode Tidak Valid...!!!", vbCritical, "Validasi Tanggal"
        DTPicker2 = DTPicker1
    End If
    isi_data
    Label1(7).Caption = Data1.Recordset.RecordCount
End If
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Form_Activate()
Label1(7).Caption = Data1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbtravel.mdb"
kosong
DTPicker1 = Date
DTPicker2 = Date
isi_cmb1
isi_cmb2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
main_form.Enabled = True
main_form.Show
End Sub
