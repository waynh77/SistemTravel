VERSION 5.00
Begin VB.Form CtkUlang_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK ULANG"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "CtkUlang_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "BATAL"
      Height          =   855
      Index           =   2
      Left            =   3720
      Picture         =   "CtkUlang_frm.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CETAK"
      Height          =   855
      Index           =   1
      Left            =   1920
      Picture         =   "CtkUlang_frm.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "CARI"
      Height          =   855
      Index           =   0
      Left            =   120
      Picture         =   "CtkUlang_frm.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   480
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR TRANSAKSI"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS TRANSAKSI"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "CtkUlang_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
isi_cmb2
Combo2.Enabled = True
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
Select Case Combo1.ListIndex
Case 0
    Data2.RecordSource = "select * from db_log where no_lg='" & Combo2 & "'"
    Data2.Refresh
Case 1
    Data2.RecordSource = "select * from db_invoice where no_invoice='" & Combo2 & "'"
    Data2.Refresh
Case 2
    Data2.RecordSource = "select * from inv_nt where no_inv='" & Combo2 & "'"
    Data2.Refresh
End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0

Case 1
    Select Case Combo1.ListIndex
    Case 0
        If Not Data2.Recordset.BOF Then
            cetak_lg
        End If
    Case 1
        If Not Data2.Recordset.BOF Then
            cetak_invTiket
        End If
    Case 2
        If Not Data2.Recordset.BOF Then
            cetak_invNT
        End If
    End Select
Case 2
    Unload Me
End Select
End Sub

Sub cetak_lg()
Dim urut_nama As Byte
Dim no_detil As Byte
Dim nm As String
Dim ttl As Byte
Dim flight As String
Dim turun As Byte
Dim cd As String
Dim ttl_hrg As Double
Dim cur As String
Dim tg As Date
Printer.CurrentX = 0
Printer.CurrentY = 0
turun = 0
Do Until turun = 12
    Printer.Print
    turun = turun + 1
Loop
With Data2.Recordset
Printer.Print Tab(25); !COMPANY
Printer.Print Tab(25); !contact_person
turun = 0
Do Until turun = 5
    Printer.Print
    turun = turun + 1
Loop
Data2.RecordSource = "select * from db_log where no_lg = '" & Combo1 & "' order by passanger_name asc"
Data2.Refresh
    .MoveFirst
    cd = !kode
    cur = !Currency
    tg = !tgl_lg
    urut_nama = 0
    turun = 0
    Do Until turun = 10
    Do While Not .EOF
        If !passanger_name <> nm Then
            urut_nama = urut_nama + 1
            Printer.Print Tab(20); Format(urut_nama, "00") & ". "; !passanger_name & "/" & !jenkel_penumpang
            turun = turun + 1
        End If
        nm = !passanger_name
        .MoveNext
    Loop
        turun = turun + 1
        Printer.Print
    Loop
    turun = 0
    Do Until turun = 12
        Printer.Print
        turun = turun + 1
    Loop
    Data2.RecordSource = "select * from db_log where no_lg='" & Combo1 & "' order by no_penerbangan asc"
    Data2.Refresh
    .MoveFirst
    ttl = 0
    no_detil = 0
    turun = 0
    Printer.Print Tab(20); "No."; Tab(35); "Flight"; Tab(50); "Class"; Tab(60); "Date"; Tab(75); "Route"; Tab(95); "Status"
    Do Until turun = 17
    Do While Not .EOF
        If !kode_maskapai & !no_penerbangan <> flight Then
            no_detil = no_detil + 1
            Printer.Print Tab(20); Format(no_detil, "00") & ". "; Tab(35); !kode_maskapai & !no_penerbangan; Tab(50); !Class; Tab(60); !tanggal_berangkat; Tab(75); !From & "-" & !To; Tab(95); !Status
            turun = turun + 1
        End If
        flight = !kode_maskapai & !no_penerbangan
        ttl_hrg = ttl_hrg + !harga
        ttl = ttl + 1
        .MoveNext
    Loop
    turun = turun + 1
    Printer.Print
    Loop
    Printer.Print Tab(20); "Total Pax = " & ttl
    Printer.Print Tab(20); "CODE : " & cd
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Data2.Refresh
    Printer.Print Tab(115); cur & " " & Format(ttl_hrg, "###,###,###,##")
    turun = 0
    Do Until turun = 5
        Printer.Print
        turun = turun + 1
    Loop
    Printer.Print Tab(20); "jakarta, " & Format(tg, "d mmmm yyyy")
    Printer.Print
    Printer.Print
    Printer.Print Tab(20); "( " & main_form.Text1(1) & ")"
End With
Printer.EndDoc
Data2.Refresh
End Sub

Sub cetak_invTiket()
Data2.Recordset.MoveFirst
With PrintInv_frm
    .Label1(4).Caption = Data2.Recordset!no_invoice 'nomor
    .Label2(3).Caption = Data2.Recordset!COMPANY 'company
    .Label2(4).Caption = Data2.Recordset!contact_person 'contact
    .Label2(5).Caption = Data2.Recordset!telp 'telp
    .Label2(6).Caption = Data2.Recordset!address 'address
    .Label2(23).Caption = Data2.Recordset!Currency 'curr
'    Label2(24).Caption = .Label1(41).Caption 'total nominal
    .Label2(13).Caption = "( " & main_form.Text1(1) & " )" 'user
    .Label2(25) = Format(Data2.Recordset!due_date, "d mmmm yyyy") 'tgl
    isi_listInv1
    .Show
End With
End Sub

Sub cetak_invNT()
With PrintInv2_frm
    .Label1(4).Caption = Data2.Recordset!no_inv 'NO INV
    .Label2(3).Caption = Data2.Recordset!COMPANY 'company
    .Label2(4).Caption = Data2.Recordset!contact_person 'contact person
    .Label2(5).Caption = Data2.Recordset!telp 'telp
    .Label2(6).Caption = Data2.Recordset!address 'address
    .Text1 = Data2.Recordset!detil_beli 'ket
    .Text2 = Data2.Recordset!Currency 'curr
    .Text3 = Format(Data2.Recordset!Nominal, "###,###.##") 'nominal trans
    .Label2(23).Caption = Data2.Recordset!Currency 'curr
    .Label2(24).Caption = Format(Data2.Recordset!Nominal, "###,###.##") 'total trans
    .Label2(25) = Format(Data2.Recordset!tgl, "d mmmm yyyy") 'tgl
    .Show
End With
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "Letter of Guarantee"
Combo1.AddItem "Invoice Tiket"
Combo1.AddItem "Invoice Lain-lain"
End Sub

Private Sub Form_Load()
Call db_ctkulang
isi_cmb1
Combo2.Clear
Combo2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
main_form.Enabled = True
main_form.Show
End Sub

Sub isi_cmb2()
Dim x As String
Combo2.Clear
Select Case Combo1.ListIndex
Case 0
    Data1.RecordSource = "db_log"
    Data1.Refresh
    With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !no_lg <> x Then
                Combo2.AddItem !no_lg
            End If
            x = !no_lg
            .MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
    End With
Case 1
    Data1.RecordSource = "db_invoice"
    Data1.Refresh
    With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !no_invoice <> x Then
                Combo2.AddItem !no_invoice
            End If
            x = !no_invoice
            .MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
    End With
Case 2
    Data1.RecordSource = "inv_nt"
    Data1.Refresh
    With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !no_inv <> x Then
                Combo2.AddItem !no_inv
            End If
            x = !no_inv
            .MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
    End With
End Select
End Sub

Sub isi_listInv1()
Dim no As Byte
Dim a, b As String
Dim tot As Double
PrintInv_frm.List1.Clear
no = 1
With Data2.Recordset
    .MoveFirst
    Do While Not .EOF
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If Data2.Recordset!no_penerbangan = !no_penerbangan And Data2.Recordset!kode_maskapai = !kode_maskapai Then
                    a = Format(!dep, "hh:mm")
                    b = Format(!arr, "hh:mm")
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        PrintInv_frm.List1.AddItem (Format(no, "#00") & ". " & Data2.Recordset!sex_psg & " " & Data2.Recordset!nama_psg & "  " & Data2.Recordset!kode_maskapai & Data2.Recordset!no_penerbangan & "  " & Data2.Recordset!Class & "  " & Format(Data2.Recordset!tgl_berangkat, "d/m/yyyy") & "  " & !From & "-" & !To & "  " & a & "  " & b & "  " & !Status)
        'list2.AddItem (!Currency)
        PrintInv_frm.list3.AddItem (!Currency & rkanan(!hrg_tiket, "###,###,###"))
        tot = tot + Data2.Recordset!hrg_tiket
        no = no + 1
        .MoveNext
    Loop
    PrintInv_frm.Label2(24).Caption = Format(tot, "###,###,###")
End With
End Sub

Private Function rkanan(ndata, cformat) As String
    rkanan = Format(ndata, cformat)
    rkanan = Space(Len(cformat) - Len(rkanan)) + rkanan
End Function

