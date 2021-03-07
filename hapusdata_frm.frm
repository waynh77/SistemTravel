VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form hapusdata_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MASTER PENJUALAN"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "hapusdata_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Delete"
      DownPicture     =   "hapusdata_frm.frx":66C0
      Height          =   735
      Index           =   2
      Left            =   5640
      Picture         =   "hapusdata_frm.frx":738A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Width           =   4020
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "hapusdata_frm.frx":840C
      Height          =   6735
      Left            =   120
      OleObjectBlob   =   "hapusdata_frm.frx":8420
      TabIndex        =   4
      Top             =   960
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Keluar"
      DownPicture     =   "hapusdata_frm.frx":8DF3
      Height          =   735
      Index           =   1
      Left            =   7080
      Picture         =   "hapusdata_frm.frx":9ABD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Delete All"
      DownPicture     =   "hapusdata_frm.frx":A227
      Height          =   735
      Index           =   0
      Left            =   4200
      Picture         =   "hapusdata_frm.frx":AEF1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "hapusdata_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Data1.RecordSource = Combo1
Data1.Refresh
Data1.Caption = Combo1
Command1(0).Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    x = MsgBox("apakah anda yakin ingin menghapus semua data...???", vbOKCancel, "Hapus semua data")
    If x = vbOK Then
        If Data1.Recordset.RecordCount <> 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.RecordCount = 0
                Data1.Recordset.Delete
                Data1.Recordset.MoveNext
            Loop
            Data1.Refresh
            x = MsgBox("data telah dihapus...", vbOKOnly, "Hapus Data")
            Command1(0).Enabled = False
        Else
            x = MsgBox("data kosong...", vbOKOnly, "Hapus Data")
        End If
    End If
Case 1
'    End
    Data1.Refresh
    Me.Hide
    main_form.Enabled = True
    main_form.Show
    Call isi_cekdat
Case 2
    x = MsgBox("apakah anda yakin ingin menghapus data...???", vbOKCancel, "Hapus data")
    If x = vbOK Then
        If Data1.Recordset.RecordCount <> 0 Then
            Data1.Recordset.Delete
            Data1.Refresh
            x = MsgBox("data telah dihapus...", vbOKOnly, "Hapus Data")
            Command1(0).Enabled = False
        Else
            x = MsgBox("data kosong...", vbOKOnly, "Hapus Data")
        End If
    End If
End Select
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem ("db_client")
Combo1.AddItem ("db_invoice")
Combo1.AddItem ("db_log")
Combo1.AddItem ("db_lokasi")
Combo1.AddItem ("db_maskapai")
Combo1.AddItem ("db_partner")
Combo1.AddItem ("db_penerbangan")
Combo1.AddItem ("db_user")
Combo1.AddItem ("stok_tiket")
Combo1.AddItem ("temp_inv")
Combo1.AddItem ("trans_lg")
Combo1.AddItem ("data_cetak")
Combo1.AddItem ("inv_nt")
Combo1.AddItem ("pembayaran")
Combo1.AddItem ("bayar_lg")
Combo1.AddItem ("byr_invtiket")
Combo1.AddItem ("byr_invnt")
Combo1.AddItem ("remainder")
Combo1.AddItem ("vcrhotel")
Combo1.ListIndex = 0
End Sub

Private Sub Form_Activate()
Command1(0).Enabled = False
isi_cmb1
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbtravel.mdb"
End Sub

