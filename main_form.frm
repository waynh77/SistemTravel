VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form main_form 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEM PENJUALAN & AKUNTANSI (Travel Version)"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ClipControls    =   0   'False
   Icon            =   "main_form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer9 
      Interval        =   500
      Left            =   11040
      Top             =   1320
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Hapus/Edit Data"
      DownPicture     =   "main_form.frx":1CCA
      Height          =   1095
      Left            =   7560
      Picture         =   "main_form.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Hanya untuk Programer"
      Top             =   7080
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":365E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":5338
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":7012
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":8CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":A9C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":C6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main_form.frx":E37A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer4 
      Index           =   1
      Interval        =   1000
      Left            =   5880
      Top             =   0
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4530
      Left            =   1440
      TabIndex        =   15
      Top             =   3960
      Width           =   5295
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   2011
      ButtonWidth     =   1852
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DB Tiket"
            Key             =   "DB_tiket"
            Object.ToolTipText     =   "Input Data Base Tiket"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stok Tiket"
            Key             =   "Stok"
            Object.ToolTipText     =   "Input Stok Tiket"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Client"
            Key             =   "client"
            Object.ToolTipText     =   "Input Data Client"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Suplier"
            Key             =   "suplier"
            Object.ToolTipText     =   "Input Data Suplier"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LG"
            Key             =   "LG"
            Object.ToolTipText     =   "Buat Letter of Guarantee"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Invoice"
            Key             =   "invoice"
            Object.ToolTipText     =   "Buat Invoice"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "keluar_mnu"
            Object.ToolTipText     =   "Keluar Aplikasi..."
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "main_form.frx":10054
      Begin VB.Timer Timer8 
         Index           =   1
         Interval        =   1000
         Left            =   840
         Top             =   480
      End
      Begin VB.Timer Timer8 
         Index           =   0
         Interval        =   1000
         Left            =   840
         Top             =   120
      End
      Begin VB.Timer Timer7 
         Index           =   1
         Interval        =   1000
         Left            =   3720
         Top             =   0
      End
      Begin VB.Timer Timer7 
         Index           =   0
         Interval        =   1000
         Left            =   3720
         Top             =   120
      End
      Begin VB.Timer Timer6 
         Index           =   1
         Interval        =   1000
         Left            =   2760
         Top             =   0
      End
      Begin VB.Timer Timer6 
         Index           =   0
         Interval        =   1000
         Left            =   2760
         Top             =   120
      End
      Begin VB.Timer Timer5 
         Index           =   1
         Interval        =   1000
         Left            =   1800
         Top             =   480
      End
      Begin VB.Timer Timer5 
         Index           =   0
         Interval        =   1000
         Left            =   1800
         Top             =   120
      End
      Begin VB.Timer Timer4 
         Index           =   0
         Interval        =   1000
         Left            =   5640
         Top             =   120
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   4680
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   4680
         Top             =   120
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   10485
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21749
            Text            =   "www.wnh-it.com"
            TextSave        =   "www.wnh-it.com"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "4/2/2017"
            Key             =   "Tanggal Sekarang "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "10:13 PM"
            Key             =   "Jam Sekarang"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "LOG OFF"
      DownPicture     =   "main_form.frx":1036E
      Height          =   1095
      Index           =   1
      Left            =   9600
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "main_form.frx":11038
      MousePointer    =   99  'Custom
      Picture         =   "main_form.frx":11342
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "GANTI PASSWORD"
      DownPicture     =   "main_form.frx":1300C
      Height          =   1095
      Index           =   0
      Left            =   11640
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "main_form.frx":13CD6
      MousePointer    =   99  'Custom
      Picture         =   "main_form.frx":13FE0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11520
      Top             =   3960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remainder"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11520
      MouseIcon       =   "main_form.frx":15CAA
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "CEK STATUS DATA"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Top             =   3600
      Width           =   5295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "ID USER"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   3
      Left            =   7755
      TabIndex        =   7
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "POSISI/JABATAN"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   2
      Left            =   7680
      TabIndex        =   6
      Top             =   6120
      Width           =   2100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   0
      Left            =   7680
      TabIndex        =   5
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label JAM 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hh:mm:ss"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   12000
      TabIndex        =   4
      Top             =   4080
      Width           =   1440
   End
   Begin VB.Label tanggal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   12240
      TabIndex        =   3
      Top             =   3720
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   6750
      TabIndex        =   2
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APLIKASI PENJUALAN DAN AKUNTANSI TRAVEL (Ver 1.0.0)"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Index           =   0
      Left            =   2220
      TabIndex        =   0
      Top             =   2520
      Width           =   10680
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   4
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   0
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   12255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   6360
      Picture         =   "main_form.frx":15FB4
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2160
   End
   Begin VB.Menu admin_mnu 
      Caption         =   "Administrator"
      Begin VB.Menu user_mnu 
         Caption         =   "&Hak Pengguna"
      End
      Begin VB.Menu dbjual_mnu 
         Caption         =   "&Data Penjualan"
      End
   End
   Begin VB.Menu DB_mnu 
      Caption         =   "&Database"
      Begin VB.Menu dbtiket_mnu 
         Caption         =   "&Tiket"
      End
      Begin VB.Menu dbClient_mnu 
         Caption         =   "&Client"
      End
      Begin VB.Menu partner_mnu 
         Caption         =   "&Suplier"
      End
   End
   Begin VB.Menu trx_mnu 
      Caption         =   "&Transaksi"
      Begin VB.Menu stok_mnu 
         Caption         =   "&Stok Tiket"
      End
      Begin VB.Menu LG_mnu 
         Caption         =   "L&G"
      End
      Begin VB.Menu pesan_mnu 
         Caption         =   "&Invoice Tiket"
      End
      Begin VB.Menu ntiket_mnu 
         Caption         =   "&Penjualan Non Tiket"
      End
      Begin VB.Menu hotel_mnu 
         Caption         =   "&Transaksi Hotel"
      End
      Begin VB.Menu CTKulang_mnu 
         Caption         =   "&Cetak Ulang"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu byr_mnu 
      Caption         =   "Pem&bayaran"
      Begin VB.Menu byrLG_mnu 
         Caption         =   "Letter of Guarantee"
      End
      Begin VB.Menu byrinvTiket_mnu 
         Caption         =   "Invoice Tiket"
      End
      Begin VB.Menu byrInvNT_mnu 
         Caption         =   "Invoice Non Tiket"
      End
   End
   Begin VB.Menu lap_mnu 
      Caption         =   "&Laporan"
      Begin VB.Menu lapjual_mnu 
         Caption         =   "Laporan &Penjualan Tiket"
      End
      Begin VB.Menu LPNT_mnu 
         Caption         =   "Laporan Penjualan Non Tiket"
      End
      Begin VB.Menu ARAP_mnu 
         Caption         =   "AR/AP"
      End
      Begin VB.Menu LVHotel_mnu 
         Caption         =   "Laporan Voucer Hotel"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Akt_mnu 
      Caption         =   "&Akuntansi"
   End
   Begin VB.Menu tool_mnu 
      Caption         =   "&Tools"
      Begin VB.Menu calk_mnu 
         Caption         =   "&Kalkulator"
      End
      Begin VB.Menu remain_mnu 
         Caption         =   "&Remainder"
      End
   End
   Begin VB.Menu bantu_mnu 
      Caption         =   "&Bantuan"
      Begin VB.Menu about_mnu 
         Caption         =   "About"
      End
      Begin VB.Menu Manual_mnu 
         Caption         =   "User Manual"
      End
   End
   Begin VB.Menu keluar_mnu 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Boolean

Private Sub about_mnu_Click()
WaynhSoft_frm.Show
End Sub

Private Sub Akt_mnu_Click()
    ShellExecute hwnd, "open", App.Path & "\akuntansi.exe", "", "C:\", 1
End Sub

Private Sub ARAP_mnu_Click()
Me.Enabled = False
CetakArAp_frm.Show
End Sub

Private Sub byrInvNT_mnu_Click()
Me.Enabled = False
byr_frm.Show
byr_frm.SSTab1.Tab = 2
End Sub

Private Sub byrinvTiket_mnu_Click()
Me.Enabled = False
byr_frm.Show
byr_frm.SSTab1.Tab = 1
End Sub

Private Sub byrLG_mnu_Click()
Me.Enabled = False
byr_frm.Show
byr_frm.SSTab1.Tab = 0
End Sub

Private Sub calk_mnu_Click()
    AppActivate Shell("calc.exe", 1)
End Sub

Private Sub Command1_Click(Index As Integer)
Dim x As String
Select Case Index
Case 0
    Me.Enabled = False
    Gantipass_frm.Show
    Gantipass_frm.Label1(0).Caption = Text1(0)
    Gantipass_frm.Label1(1).Caption = Text1(1)
Case 1
    x = MsgBox("Apakah anda yakin untuk Log Off?", vbYesNo, "LOG OFF")
    If x = vbYes Then
        Me.Enabled = False
        UserPass.Show
        UserPass.Text1 = ""
        UserPass.Text2 = ""
        Text1(0) = ""
        Text1(1) = ""
        Text1(2) = ""
        Me.Hide
    End If
End Select
End Sub

Private Sub Command2_Click()
Me.Enabled = False
hapusdata_frm.Show
End Sub

Private Sub CTKulang_mnu_Click()
Me.Enabled = False
CtkUlang_frm.Show
End Sub

Private Sub dbClient_mnu_Click()
Me.Enabled = False
client_frm.Show
Call hak_akses1
client_frm.Label3.Caption = "mm"
End Sub

Private Sub dbjual_mnu_Click()
Command2_Click
End Sub

Private Sub dbtiket_mnu_Click()
dbtiket_frm.Show
Me.Enabled = False
Call hak_akses1
End Sub

Private Sub Form_Activate()
Call isi_cekdat
Call hak_akses1
a = False
cek_remain
End Sub


Private Sub Form_Load()
Call db_main
tanggal = Format(Date, "dd-mmmm-yyyy")
JAM = Format(Time, "hh:mm:ss")
Text1(0).Enabled = False
Text1(2).Enabled = False
Text1(1).Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4(0).Enabled = False
Timer4(1).Enabled = False
End Sub

Private Sub Invoice_mnu_Click()
Call uncons
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hotel_mnu_Click()
Me.Enabled = False
Hotel_frm.Show
End Sub

Private Sub keluar_mnu_Click()
    x = MsgBox("Apakah anda yakin ingin keluar dari aplikasi...???", vbOKCancel, "Exit")
    If x = vbOK Then
        End
    End If
End Sub

Private Sub menu_Click(Index As Integer)
Me.Enabled = False
dbuser_frm.Show
End Sub

Private Sub Label1_Click()
REmain_frm.Show
End Sub

Private Sub lapjual_mnu_Click()
Me.Enabled = False
Lapjual2_frm.Show
End Sub

Private Sub LG_mnu_Click()
If dbtiket_frm.Data3.Recordset.RecordCount = 0 Or client_frm.Data1.Recordset.RecordCount = 0 Or dbpartner_frm.Data1.Recordset.RecordCount = 0 Then
    x = MsgBox("Harap periksa kembali database...", vbOKOnly, "Validasi Data")
Else
    Me.Enabled = False
    CetakLg_frm.Show
End If
End Sub

Private Sub LPNT_mnu_Click()
lapjualNT_frm.Show
End Sub

Private Sub Manual_mnu_Click()
Call uncons
End Sub

Private Sub ntiket_mnu_Click()
InvNt_frm.Show
Me.Enabled = False
End Sub

Private Sub partner_mnu_Click()
Me.Enabled = False
dbpartner_frm.Show
Call hak_akses1
End Sub

Private Sub pesan_mnu_Click()
Dim a As Boolean
a = False
Dim b As String
x = MsgBox("Cetak Invoice berdasarkan LG?", vbYesNo, "Cetak Invoice")
If x = vbNo Then
    With InvLG_frm
        .Data4.Top = 1680
        .Data1.Visible = False
        .DBGrid3.Top = 2160
        .DBGrid3.Height = 3495
        .Data4.Caption = "Data Detil Invoice"
        .Command1(0).Enabled = True
        .Frame3.Visible = True
        .Combo6.Visible = True
        .Label4.Caption = "inv"
        .Label2.Caption = "et"
        .Combo1.Enabled = True
        .DBGrid2.Enabled = False
        .DBGrid2.Visible = False
        .Frame1.Enabled = True
        .Label1(32).Caption = "Hrg Dasar"
        .Data2.RecordSource = "select * from db_client"
        .Data2.Refresh
        .Combo8.Clear
    End With
    InvLG_frm.Show
    Me.Enabled = False
Else
    With CetakLg_frm.Data1.Recordset
    If Not .BOF Then
        With InvLG_frm
        .DBGrid2.Visible = True
        .DBGrid2.Enabled = True
        .Command1(0).Enabled = False
        .Label2.Caption = "el"
        .Label4.Caption = "invlg"
        .Combo6.Visible = False
        .Combo8.Clear
        .Data4.Caption = "Data Tambah Tiket"
        .Data1.Visible = True
        .Data4.Top = 3720
        .DBGrid3.Top = 4200
        .Label1(32).Caption = "Harga LG"
        .DBGrid3.Height = 1455
        .Frame3.Visible = False
        .Combo1.Enabled = False
        .Frame1.Enabled = False
        End With
        .MoveFirst
        b = ""
        Do While Not .EOF
            If !status_lg = 0 Then
                a = True
                If !no_lg <> b Then
                    InvLG_frm.Combo8.AddItem (!no_lg)
                    b = !no_lg
                End If
            End If
            .MoveNext
        Loop
        If a = True Then
            Me.Enabled = False
            InvLG_frm.Show
            InvLG_frm.Combo8.ListIndex = 0
            InvLG_frm.Frame1.Enabled = False
        Else
            x = MsgBox("Data LG sudah dicetak invoice semua...", vbOKOnly, "Validasi Data")
        End If
    Else
        x = MsgBox("Tidak ada data LG...", vbOKOnly, "Validasi Data")
    End If
    End With
End If
End Sub

Private Sub remain_mnu_Click()
Me.Enabled = False
Remainder_frm.Show
End Sub

Private Sub stok_mnu_Click()
Me.Enabled = False
stoktiket_frm.Show
Call hak_akses1
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub timer1_Timer()
JAM = Format(Time, "hh:mm:ss")
End Sub

Private Sub Timer2_Timer()
Toolbar1.Buttons(5).Image = 0
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Toolbar1.Buttons(5).Image = 5
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer(Index As Integer)
Select Case Index
Case 0
    Toolbar1.Buttons(6).Image = 0
    Timer4(1).Enabled = True
    Timer4(0).Enabled = False
Case 1
    Toolbar1.Buttons(6).Image = 6
    Timer4(0).Enabled = True
    Timer4(1).Enabled = False
End Select
End Sub

Private Sub Timer5_Timer(Index As Integer)
Select Case Index
Case 0
    Toolbar1.Buttons(2).Image = 0
    Timer5(1).Enabled = True
    Timer5(0).Enabled = False
Case 1
    Toolbar1.Buttons(2).Image = 2
    Timer5(0).Enabled = True
    Timer5(1).Enabled = False
End Select
End Sub

Private Sub Timer6_Timer(Index As Integer)
Select Case Index
Case 0
    Toolbar1.Buttons(3).Image = 0
    Timer6(1).Enabled = True
    Timer6(0).Enabled = False
Case 1
    Toolbar1.Buttons(3).Image = 3
    Timer6(0).Enabled = True
    Timer6(1).Enabled = False
End Select
End Sub

Private Sub Timer7_Timer(Index As Integer)
Select Case Index
Case 0
    Toolbar1.Buttons(4).Image = 0
    Timer7(1).Enabled = True
    Timer7(0).Enabled = False
Case 1
    Toolbar1.Buttons(4).Image = 4
    Timer7(0).Enabled = True
    Timer7(1).Enabled = False
End Select
End Sub

Private Sub Timer8_Timer(Index As Integer)
Select Case Index
Case 0
    Toolbar1.Buttons(1).Image = 0
    Timer8(1).Enabled = True
    Timer8(0).Enabled = False
Case 1
    Toolbar1.Buttons(1).Image = 1
    Timer8(0).Enabled = True
    Timer8(1).Enabled = False
End Select
End Sub

Private Sub Timer9_Timer()
If a = False Then
'    Label1.Visible = True
    Label1.Caption = "REMAINDER"
    a = True
Else
'    Label1.Visible = False
    Label1.Caption = "Klik disini"
    a = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    dbtiket_mnu_Click
Case 2
    stok_mnu_Click
Case 3
    dbClient_mnu_Click
Case 4
    partner_mnu_Click
Case 5
    LG_mnu_Click
Case 6
    pesan_mnu_Click
Case 7
    keluar_mnu_Click
End Select
End Sub


Private Sub user_mnu_Click()
Me.Enabled = False
dbuser_frm.Show
End Sub

Sub cek_remain()
Dim cek As Boolean
cek = False
Call db_main
Data1.RecordSource = "select * from remainder where status=true"
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !tgl = Date Then
            If Val(Format(Time, "hh.mm")) >= Val(Format(!waktu, "hh.mm")) Then
                cek = True
                .MoveLast
            End If
        ElseIf !tgl < Date Then
            cek = True
            .MoveLast
        End If
        .MoveNext
    Loop
Else
    Timer9.Enabled = False
    Label1.Visible = False
End If
End With
If cek = True Then
    Timer9.Enabled = True
    Label1.Visible = True
Else
    Timer9.Enabled = False
    Label1.Visible = False
End If
Data1.Refresh
End Sub


