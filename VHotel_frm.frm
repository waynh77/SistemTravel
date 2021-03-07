VERSION 5.00
Begin VB.Form VHotel_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucer Hotel"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12255
   Icon            =   "VHotel_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   8640
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "VHotel_frm.frx":1CCA
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "VHotel_frm.frx":1CD0
      Top             =   5520
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   195
      Index           =   13
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   10200
      TabIndex        =   29
      Top             =   6360
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conf Name"
      Height          =   195
      Index           =   1
      Left            =   7440
      TabIndex        =   28
      Top             =   6360
      Width           =   795
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   12000
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO :"
      Height          =   195
      Left            =   8040
      TabIndex        =   27
      Top             =   960
      Width           =   330
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "VOUCHER HOTEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   7680
      TabIndex        =   26
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "BY"
      Height          =   195
      Index           =   16
      Left            =   9840
      TabIndex        =   25
      Top             =   7320
      Width           =   210
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Index           =   2
      Left            =   7100
      Top             =   4995
      Width           =   375
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Index           =   1
      Left            =   7100
      Top             =   4305
      Width           =   375
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Index           =   0
      Left            =   7100
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.wnh-it.com"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   15
      Left            =   465
      TabIndex        =   24
      Top             =   1440
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "WAYNH TOURS and TRAVEL"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   14
      Left            =   480
      TabIndex        =   23
      Top             =   1200
      Width           =   2730
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   720
      Picture         =   "VHotel_frm.frx":1CD6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fax 021-5367 3901, Email : Wisata.Wina@Gmail.com, Wisata_Wina@yahoo.com"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   22
      Top             =   8280
      Width           =   11775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gedung BERCA Lt.5 - 504, Jl Palmerah Utara No. 14 Slipi,Jakarta Barat Telp.021-5349 401, 5367 3901, "
      Height          =   315
      Index           =   11
      Left            =   255
      TabIndex        =   21
      Top             =   8040
      Width           =   11745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "VOUCER ISSUED"
      Height          =   195
      Index           =   10
      Left            =   10200
      TabIndex        =   20
      Top             =   6120
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONFIRMATION BY"
      Height          =   195
      Index           =   9
      Left            =   7440
      TabIndex        =   19
      Top             =   6120
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "No. In Party"
      Height          =   195
      Index           =   8
      Left            =   5160
      TabIndex        =   18
      Top             =   2640
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "BY CHEQUE/GIRO"
      Height          =   435
      Index           =   7
      Left            =   7680
      TabIndex        =   17
      Top             =   5160
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ON ACCOUNT OF WAYNHSOFT"
      Height          =   195
      Index           =   6
      Left            =   7680
      TabIndex        =   16
      Top             =   4440
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAID BY CLIENTS DIRECTLY"
      Height          =   315
      Index           =   5
      Left            =   7680
      TabIndex        =   15
      Top             =   3720
      Width           =   2190
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TO "
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   14
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROVIDE"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FOR"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LEAVING"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   11
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ARRIVING"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Value stated only as specified other charges                      should be paid by client"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "- Settlement can only be done by showing this                     voucer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   7080
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   360
      Index           =   2
      Left            =   7205
      Picture         =   "VHotel_frm.frx":18140
      Stretch         =   -1  'True
      Top             =   4995
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   360
      Index           =   1
      Left            =   7205
      Picture         =   "VHotel_frm.frx":19E0A
      Stretch         =   -1  'True
      Top             =   4305
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   360
      Index           =   0
      Left            =   7205
      Picture         =   "VHotel_frm.frx":1BAD4
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wahyu"
      Height          =   195
      Index           =   0
      Left            =   10200
      TabIndex        =   7
      Top             =   7320
      Width           =   510
   End
   Begin VB.Menu print_mnu 
      Caption         =   "Print"
   End
   Begin VB.Menu x_mnu 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "VHotel_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctk As Boolean

Private Sub Form_Load()
ctk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hotel_frm.Enabled = True
Hotel_frm.Show
End Sub

Private Sub print_mnu_Click()
x = MsgBox("Apakah cetak header?", vbYesNoCancel, "Cetak Voucer Hotel")
If x <> vbCancel Then
    If x = vbNo Then
        umpet
    End If
    Me.PrintForm
    muncul
    ctk = True
    Hotel_frm.Command1(1).Enabled = True
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub x_mnu_Click()
If ctk = True Then
    Unload Me
Else
    MsgBox "Voucer blm dicetak...", vbCritical, "Validasi Proses"
End If
End Sub

Sub umpet()
Dim x As Byte
x = 0
Do Until x = 16
    Label3(x).Visible = False
    x = x + 1
Loop
Image3.Visible = False
Label5(1).Visible = False
Label2(0).Visible = False
Label2(1).Visible = False
Line1.Visible = False
Shape2(0).Visible = False
Shape2(1).Visible = False
Shape2(2).Visible = False
Label6.Visible = False
End Sub

Sub muncul()
Dim x As Byte
x = 0
Do Until x = 16
    Label3(x).Visible = True
    x = x + 1
Loop
Image3.Visible = True
Label5(1).Visible = True
Label2(0).Visible = True
Label2(1).Visible = True
Line1.Visible = True
Shape2(0).Visible = True
Shape2(1).Visible = True
Shape2(2).Visible = True
Label6.Visible = True
End Sub
