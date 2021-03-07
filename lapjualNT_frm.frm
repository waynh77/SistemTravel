VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form lapjualNT_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN PENJUALAN NON TIKET"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "lapjualNT_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin MSACAL.Calendar Calendar2 
      Height          =   2295
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   3495
      _Version        =   524288
      _ExtentX        =   6165
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2008
      Month           =   1
      Day             =   8
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      _Version        =   524288
      _ExtentX        =   6165
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2008
      Month           =   1
      Day             =   6
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Keluar"
      DownPicture     =   "lapjualNT_frm.frx":0CCA
      Height          =   855
      Index           =   1
      Left            =   2640
      Picture         =   "lapjualNT_frm.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Cetak Laporan"
      DownPicture     =   "lapjualNT_frm.frx":265E
      Height          =   855
      Index           =   0
      Left            =   120
      Picture         =   "lapjualNT_frm.frx":3328
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   360
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
      TabIndex        =   4
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   885
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
      Top             =   120
      Width           =   3900
   End
End
Attribute VB_Name = "lapjualNT_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As String

Private Sub Calendar1_Click()
Text1 = Format(Calendar1.Value, "d mmm yyyy")
Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
Text2 = Format(Calendar2.Value, "d mmm yyyy")
Calendar2.Visible = False
End Sub

Private Sub Combo1_Click()
If Combo1 = "Lain-lain..." Then
    Label1(2).Visible = True
    Text2.Visible = True
Else
    Label1(2).Visible = False
    Text2.Visible = False
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Combo1 = "Harian" And Text1 = "klik disini..." Then
        x = MsgBox("anda belum memasukan tanggal...", vbOKOnly, "Validasi Data")
        Calendar1.Visible = True
    ElseIf Combo1 = "Lain-lain..." And (Text1 = "klik disini..." Or Text2 = "klik disini...") Then
        x = MsgBox("anda belum memasukan tanggal...", vbOKOnly, "Validasi Data")
        Calendar1.Visible = True
    Else
        CrystalReport1.Reset
        If Combo1 = "Harian" Then
            CrystalReport1.ReportFileName = App.Path & "\laporan penjualan NON TIKET.rpt"
            CrystalReport1.SelectionFormula = "{inv_nt.tgl}= date(" & Format(Text1, "yyyy,m,d") & ")"
            CrystalReport1.RetrieveDataFiles
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Action = 1
        ElseIf Combo1 = "Lain-lain..." Then
            CrystalReport1.ReportFileName = App.Path & "\laporan penjualan NON TIKET.rpt"
            CrystalReport1.SelectionFormula = "{INV_nt.tgl}>= date(" & Format(Text1, "yyyy,m,d") & ") and {inv_nt.tgl} <= date(" & Format(Text2, "yyyy,m,d") & ")"
            CrystalReport1.RetrieveDataFiles
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Action = 1
        Else
            Call uncons
        End If
    End If
Case 1
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub Form_Activate()
isi_cmb1
Calendar1.Visible = False
Calendar2.Visible = False
Calendar1 = Date
Calendar2 = Date
kosong
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem ("Harian")
'Combo1.AddItem ("Bulanan")
'Combo1.AddItem ("Triwulanan")
'Combo1.AddItem ("Kwartalan")
'Combo1.AddItem ("Tahunan")
Combo1.AddItem ("Lain-lain...")
Combo1.ListIndex = 0
End Sub


Sub kosong()
Text1 = "klik disini..."
Text2 = "klik disini..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click (1)
End Sub

Private Sub Text1_Click()
Calendar1.Visible = True
End Sub

Private Sub Text2_Click()
Calendar2.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
