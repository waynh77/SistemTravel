VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Hotel_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Hotel"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   ClipControls    =   0   'False
   Icon            =   "Hotel_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5640
      TabIndex        =   37
      Text            =   "Text9"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Check2"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "KELUAR"
      Height          =   735
      Index           =   3
      Left            =   7560
      MouseIcon       =   "Hotel_frm.frx":3482
      MousePointer    =   99  'Custom
      Picture         =   "Hotel_frm.frx":378C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "SIMPAN"
      Height          =   735
      Index           =   2
      Left            =   1920
      MouseIcon       =   "Hotel_frm.frx":3EF6
      MousePointer    =   99  'Custom
      Picture         =   "Hotel_frm.frx":4200
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "CETAK INVOICE"
      Height          =   735
      Index           =   1
      Left            =   2280
      MouseIcon       =   "Hotel_frm.frx":496A
      MousePointer    =   99  'Custom
      Picture         =   "Hotel_frm.frx":4C74
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "CETAK VOUCER"
      Height          =   735
      Index           =   0
      Left            =   120
      MouseIcon       =   "Hotel_frm.frx":593E
      MousePointer    =   99  'Custom
      Picture         =   "Hotel_frm.frx":5C48
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3360
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   285
      Left            =   5640
      TabIndex        =   18
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   39784
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "BY CHEQUE/GIRO"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   17
      Top             =   3480
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "ON ACCOUNT OF WINA TOURS AND TRAVEL"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   3120
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "PAID BY CILENTS DIRECTLY"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   15
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6120
      TabIndex        =   14
      Text            =   "Text8"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6120
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6120
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   840
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   39784
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   39784
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Hotel_frm.frx":6912
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   645
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Hotel_frm.frx":6918
      Top             =   840
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   39784
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issued By"
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
      Index           =   13
      Left            =   4560
      TabIndex        =   36
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      Index           =   12
      Left            =   120
      TabIndex        =   35
      Top             =   1560
      Width           =   990
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
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   11
      Left            =   4560
      TabIndex        =   30
      Top             =   3840
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No.In Party"
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
      Index           =   10
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   1155
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
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   9
      Left            =   4560
      TabIndex        =   28
      Top             =   1560
      Width           =   930
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
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   8
      Left            =   4560
      TabIndex        =   27
      Top             =   2280
      Width           =   1155
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
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   7
      Left            =   4560
      TabIndex        =   26
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conf. By"
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
      Index           =   6
      Left            =   4560
      TabIndex        =   25
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Index           =   5
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leaving"
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
      Index           =   4
      Left            =   4560
      TabIndex        =   23
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arriving"
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
      Left            =   4560
      TabIndex        =   22
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provide"
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
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For "
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
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Voucer"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "Hotel_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cek_input As Boolean
Dim cek_proses As Boolean

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Text3 = Combo2
Else
    Text3 = ""
    Text3.SetFocus
End If
End Sub

Private Sub Combo2_Click()
isi_pembeli
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
   cek_input = False
   cek
   If cek_input = True Then
    Me.Enabled = False
    With VHotel_frm
        .Text1(0) = UCase(Text3)
        .Text1(1) = UCase(Text4)
        .Text1(2) = Text6
        .Text1(3) = Format(DTPicker2, "d mmm yyyy")
        .Text1(4) = Format(DTPicker3, "d mmm yyyy")
        .Text1(5) = UCase(Text2)
        .Text1(9) = UCase(Text1)
        .Label1(0) = UCase(Text9)
        .Label1(1) = UCase(Text5)
        .Label1(2) = Format(DTPicker1, "d mmm yyyy")
        If Check1(0).Value = 0 Then
            .Image2(0).Visible = False
        Else
            .Image2(0).Visible = True
        End If
        If Check1(1).Value = 0 Then
            .Image2(1).Visible = False
        Else
            .Image2(1).Visible = True
        End If
        If Check1(2).Value = 0 Then
            .Image2(2).Visible = False
        Else
            .Image2(2).Visible = True
        End If
    End With
    VHotel_frm.Show
   End If
Case 1
    With InvNt_frm
        .Text1.Enabled = False
'        .Text1 = Text3
        .Text4.Enabled = False
        .Text4 = "Voucer Hotel " & Text2 & " No." & Text1 & ", tgl : " & Format(DTPicker1, "d/m/yy") & ", a/n " & Text3 & ", Arrival " & Format(DTPicker2, "d/m/yy") & ", Leaving " & Format(DTPicker3, "d/m/yy") & ", " & Text4 '& ", Confirm By " & Text5
        .DTPicker1 = DTPicker4
        .Combo2.Enabled = False
        .Combo2 = Combo1
        .Combo1 = Combo2
        .Combo1.Enabled = False
        .Text6.Enabled = False
        .Text5.Enabled = False
        .Text6 = Text7
        .Text5 = Text8
        .DTPicker1.Enabled = False
        .Text3 = ""
        .Text2 = ""
        .Text2.Enabled = False
        .Text3.Enabled = False
        .isi_pembeli2
        .Show
        simpan
        cek_proses = True
        Unload Me
    End With
Case 2

Case 3
    Unload Me
End Select
End Sub

Sub simpan()
Data1.Refresh
With Data1.Recordset
    .AddNew
    !no_voucer = Text1
    !tgl_vcr = DTPicker1
    !To = Text2
    !for = Text3
    !provide = Text4
    !nmr_party = Text6
    !arr = DTPicker2
    !leaving = DTPicker3
    !conf_by = Text5
    !curr = Combo1
    !hrg_dasar = Text7
    !hrg_jual = Text8
    !due_date = DTPicker4
    !payment = Check1(1).Caption
    .Update
End With
End Sub

Private Sub Form_Activate()
Auto_Vcr
Text1.MaxLength = 10
End Sub

Private Sub Form_Load()
Call db_hotel
kosong
'test
cek_proses = False
Command1(1).Enabled = False
isi_cmb2
isi_pembeli
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = "REZA"
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
DTPicker4 = Date
Check1(0).Value = 0
Check1(1).Value = 1
Check1(2).Value = 0
Combo1.Clear
Combo1.AddItem "Rp"
Combo1.AddItem ("USD")
Combo1.AddItem ("SGD")
Combo1.AddItem ("HKD")
Combo1.AddItem ("CNY")
Combo1.AddItem ("JPY")
Combo1.AddItem ("MYR")
Combo1.AddItem ("SAR")
Combo1.AddItem ("THB")
Combo1.AddItem ("NTD")
Combo1.AddItem ("EURO")
Combo1.AddItem ("GBP")
Combo1.AddItem ("CHF")
Combo1.AddItem ("AUD")
Combo1.AddItem ("NZD")
Combo1.AddItem ("CND")
Combo1.AddItem ("PHP")
Combo1.AddItem ("WON")
Combo1.AddItem ("IND")
Combo1.AddItem ("VND")
Combo1.AddItem ("AED")
Combo1.AddItem ("BND")
Combo1.AddItem ("OMR")
Combo1.AddItem ("EGP")
Combo1.AddItem ("SRI")
Combo1.AddItem ("QTR")
Combo1.AddItem ("ZAR")
Combo1.ListIndex = 0
End Sub

Sub cek()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text7 = "" Or Text8 = "" Then
    MsgBox "Input belum lengkap", vbInformation, "Validasi Input"
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    ElseIf Text3 = "" Then
        Text3.SetFocus
    ElseIf Text4 = "" Then
        Text4.SetFocus
    ElseIf Text5 = "" Then
        Text5.SetFocus
    ElseIf Text7 = "" Then
        Text7.SetFocus
    ElseIf Text8 = "" Then
        Text8.SetFocus
    End If
Else
    cek_input = True
End If
End Sub

Sub test()
Text1 = "12312"
Text2 = "dadadasfd"
Text3 = "dwsfsf"
Text4 = "fdsfs"
Text5 = "fdsfs"
Text6 = "2"
Text7 = "3121213"
Text8 = "4324242"
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
DTPicker4 = Date
Check1(0).Value = 0
Check1(1).Value = 1
Check1(2).Value = 0
Combo1.Clear
Combo1.AddItem ("Rp")
Combo1.AddItem ("USD")
Combo1.AddItem ("SGD")
Combo1.AddItem ("HKD")
Combo1.AddItem ("CNY")
Combo1.AddItem ("JPY")
Combo1.AddItem ("MYR")
Combo1.AddItem ("SAR")
Combo1.AddItem ("THB")
Combo1.AddItem ("NTD")
Combo1.AddItem ("EURO")
Combo1.AddItem ("GBP")
Combo1.AddItem ("CHF")
Combo1.AddItem ("AUD")
Combo1.AddItem ("NZD")
Combo1.AddItem ("CND")
Combo1.AddItem ("PHP")
Combo1.AddItem ("WON")
Combo1.AddItem ("IND")
Combo1.AddItem ("VND")
Combo1.AddItem ("AED")
Combo1.AddItem ("BND")
Combo1.AddItem ("OMR")
Combo1.AddItem ("EGP")
Combo1.AddItem ("SRI")
Combo1.AddItem ("QTR")
Combo1.AddItem ("ZAR")
Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Command1(1).Enabled = True And cek_proses = False Then
    MsgBox "Transaksi belum cetak invoice...", vbCritical, "Validasi Proses"
    Cancel = 1
ElseIf Command1(1).Enabled = False And cek_proses = False Then
    main_form.Enabled = True
    main_form.Show
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Sub isi_cmb2()
Combo2.Clear
With client_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo2.AddItem !COMPANY
        .MoveNext
    Loop
'Combo2.ListIndex = 0
End If
End With
End Sub

Sub isi_pembeli()
With client_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !COMPANY = Combo2 Then
            Text3 = !contact_person1
            'Text2 = !telp1
            'Text3 = !address1
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

