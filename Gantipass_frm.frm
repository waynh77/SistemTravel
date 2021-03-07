VERSION 5.00
Begin VB.Form Gantipass_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GANTI PASSWORD"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "Gantipass_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   " "
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "text1"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "BATAL"
      DownPicture     =   "Gantipass_frm.frx":1CCA
      Height          =   975
      Left            =   2640
      MouseIcon       =   "Gantipass_frm.frx":2994
      MousePointer    =   99  'Custom
      Picture         =   "Gantipass_frm.frx":2C9E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PROSES"
      DownPicture     =   "Gantipass_frm.frx":6120
      Height          =   975
      Left            =   120
      MouseIcon       =   "Gantipass_frm.frx":6DEA
      MousePointer    =   99  'Custom
      Picture         =   "Gantipass_frm.frx":70F4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   2280
      Picture         =   "Gantipass_frm.frx":A576
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   795
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   240
      X2              =   4920
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   4920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Konfirmasi Password"
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
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password baru"
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
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password lama"
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
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "Gantipass_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b As Boolean
With Data1.Recordset
'validasi password lama
a = False
b = False
.MoveFirst
Do While Not .EOF
If Label1(0).Caption = !id_user And Text1 = !Password Then
    a = True
End If
.MoveNext
Loop
'validasi password baru
If Text2 = text3 Then
b = True
End If
If a = True And b = True Then
    .MoveFirst
    Do While Not .EOF
        If Label1(0).Caption = !id_user And Text1 = !Password Then
            .Edit
            !Password = Text2
            .Update
            .MoveLast
        End If
        .MoveNext
    Loop
    x = MsgBox("selamat password telah berhasil diganti", vbOKOnly, "validasi password")
    Data1.Refresh
    Text1 = ""
    Text2 = ""
    text3 = ""
    Me.Hide
    main_form.Enabled = True
    main_form.Show
Else
    x = MsgBox("password tidak valid", vbOKOnly, "validasi password")
    kosong
    Text1.SetFocus
End If
End With
End Sub

Private Sub Command2_Click()
Me.Hide
main_form.Enabled = True
main_form.Show
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Call gantipass
kosong
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
main_form.Enabled = True
main_form.Show
End Sub

Private Sub kosong()
Text1 = ""
Text2 = ""
text3 = ""
End Sub
