VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form InvNt_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penjualan Non Tiket"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   3840
      TabIndex        =   21
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58589185
      CurrentDate     =   39784
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   3840
      TabIndex        =   19
      Text            =   "Text6"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "InvNt_frm.frx":0000
      Height          =   735
      Index           =   1
      Left            =   2880
      MouseIcon       =   "InvNt_frm.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "InvNt_frm.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "InvNt_frm.frx":1C9E
      Height          =   735
      Index           =   0
      Left            =   240
      MouseIcon       =   "InvNt_frm.frx":2968
      MousePointer    =   99  'Custom
      Picture         =   "InvNt_frm.frx":2C72
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   3840
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "InvNt_frm.frx":393C
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "InvNt_frm.frx":3942
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   9
      Left            =   2400
      TabIndex        =   20
      Top             =   5160
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Jual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   8
      Left            =   2400
      TabIndex        =   18
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   7
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   240
      X2              =   5400
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Dasar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   6
      Left            =   2400
      TabIndex        =   12
      Top             =   4440
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detil Pembelian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telpon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pembeli"
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
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Perusahaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   23
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Menu lap_nt 
      Caption         =   "Laporan Penjualan Non Tiket"
   End
End
Attribute VB_Name = "InvNt_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
isi_pembeli2
End Sub

Private Sub Combo1_Click()
isi_pembeli2
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Text4 = "" Or Text5 = "" Then
        MsgBox "Data belum lengkap...", vbCritical, "Validasi Input"
        If Text4 = "" Then
            Text4.SetFocus
        ElseIf Text5 = "" Then
            Text5.SetFocus
        End If
    Else
        simpan
        PrintInv2_frm.Show
        main_form.Enabled = True
'        Unload Me
    End If
Case 1
    Unload Me
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Sub simpan()
Call db_invnt
With Data1.Recordset
    .AddNew
    !COMPANY = Combo1
    !no_inv = Label1(7).Caption
    !contact_person = Text1
    !telp = Text2
    !address = Text3
    !detil_beli = Text4
    !Currency = Combo2
    !Nominal = Text5
    !tgl = Date
    !JAM = Time
    !hrg_dasar = Text6
    !due_date = DTPicker1
    .Update
    Data1.Refresh
End With
End Sub

Private Sub Form_Activate()
Call invnt_auto
End Sub

Private Sub Form_Load()
Call db_invnt
isi_cmb1
isi_cmb2
isi_pembeli2
kosong
End Sub

Sub kosong()
Text4 = ""
Text5 = ""
Text6 = ""
DTPicker1 = Date
End Sub

Sub isi_cmb1()
Combo1.Clear
With client_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !COMPANY
        .MoveNext
    Loop
Combo1.ListIndex = 0
End If
End With
End Sub

Sub isi_pembeli2()
With client_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !COMPANY = Combo1 Then
            Text1 = !contact_person1
            Text2 = !telp1
            Text3 = !address1
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub isi_cmb2()
Combo2.Clear
Combo2.AddItem ("RP")
Combo2.AddItem ("USD")
Combo2.AddItem ("SGD")
Combo2.AddItem ("HKD")
Combo2.AddItem ("CNY")
Combo2.AddItem ("JPY")
Combo2.AddItem ("MYR")
Combo2.AddItem ("SAR")
Combo2.AddItem ("THB")
Combo2.AddItem ("NTD")
Combo2.AddItem ("EURO")
Combo2.AddItem ("GBP")
Combo2.AddItem ("CHF")
Combo2.AddItem ("AUD")
Combo2.AddItem ("NZD")
Combo2.AddItem ("CND")
Combo2.AddItem ("PHP")
Combo2.AddItem ("WON")
Combo2.AddItem ("IND")
Combo2.AddItem ("VND")
Combo2.AddItem ("AED")
Combo2.AddItem ("BND")
Combo2.AddItem ("OMR")
Combo2.AddItem ("EGP")
Combo2.AddItem ("SRI")
Combo2.AddItem ("QTR")
Combo2.AddItem ("ZAR")
Combo2.ListIndex = 0
End Sub

Private Sub lap_nt_Click()
    lapjualNT_frm.Show
    Unload Me
    main_form.Enabled = True
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If

End Sub
