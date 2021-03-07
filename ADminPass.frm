VERSION 5.00
Begin VB.Form ADminPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Password"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "FoxPro 3.0;"
      DatabaseName    =   "C:\Documents and Settings\wahyu\My Documents\waynh project\vb project\travel project\DBtravel_fox"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "ADMIN"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "ADminPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Refresh
If Text1 <> Data1.Recordset!Password Then
    x = MsgBox("Password Salah...", vbOKOnly, "Invalid Password")
    Text1 = ""
    Text1.SetFocus
    Data1.Refresh
Else
    Text1 = ""
    Me.Hide
    main_form.Show
End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Text1 = ""
End Sub
