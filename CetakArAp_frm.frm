VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form CetakArAp_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Ar/Ap"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   ClipControls    =   0   'False
   Icon            =   "CetakArAp_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2160
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "BATAL"
      Height          =   855
      Index           =   1
      Left            =   4200
      Picture         =   "CetakArAp_frm.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "PREVIEW"
      Height          =   855
      Index           =   0
      Left            =   2760
      Picture         =   "CetakArAp_frm.frx":1D4C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   240
      Picture         =   "CetakArAp_frm.frx":2DCE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PILIH JENIS LAPORAN"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "CetakArAp_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub isi_cmb()
Combo1.Clear
Combo1.AddItem "AP/Hutang LG"
Combo1.AddItem "AR/Piutang Tiket"
Combo1.AddItem "AR/Piutang NonTiket"
Combo1.ListIndex = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Select Case Combo1.ListIndex
    Case 0
        CrystalReport1.ReportFileName = App.Path & "\AP LG Report.rpt"
    Case 1
        CrystalReport1.ReportFileName = App.Path & "\AR Tiket Report.rpt"
    Case 2
        CrystalReport1.ReportFileName = App.Path & "\AR NT Report.rpt"
    End Select
        CrystalReport1.SelectionFormula = "{Pembayaran.sisa}<> 0"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
isi_cmb
End Sub

Private Sub Form_Unload(Cancel As Integer)
main_form.Enabled = True
End Sub
