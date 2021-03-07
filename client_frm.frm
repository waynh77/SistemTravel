VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form client_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATABASE CLIENT"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "client_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "D A T A B A S E   C L I E N T"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Picture         =   "client_frm.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   240
      Width           =   4815
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text12"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text11"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "KELUAR"
      DownPicture     =   "client_frm.frx":3994
      Height          =   975
      Index           =   3
      Left            =   3360
      MouseIcon       =   "client_frm.frx":465E
      MousePointer    =   99  'Custom
      Picture         =   "client_frm.frx":4968
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "STATUS CLIENT"
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   5160
      TabIndex        =   27
      Top             =   7560
      Width           =   9975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "client_frm.frx":5632
      Top             =   7320
      Width           =   4575
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   25
      Text            =   "Text9"
      Top             =   1680
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "client_frm.frx":5639
      Height          =   6855
      Left            =   5160
      OleObjectBlob   =   "client_frm.frx":564D
      TabIndex        =   23
      Top             =   240
      Width           =   9975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "HAPUS"
      DownPicture     =   "client_frm.frx":6020
      Height          =   975
      Index           =   2
      Left            =   3360
      MouseIcon       =   "client_frm.frx":6CEA
      MousePointer    =   99  'Custom
      Picture         =   "client_frm.frx":6FF4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "EDIT"
      DownPicture     =   "client_frm.frx":8976
      Height          =   975
      Index           =   1
      Left            =   1800
      MouseIcon       =   "client_frm.frx":9640
      MousePointer    =   99  'Custom
      Picture         =   "client_frm.frx":994A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "TAMBAH"
      DownPicture     =   "client_frm.frx":B2CC
      Height          =   975
      Index           =   0
      Left            =   240
      MouseIcon       =   "client_frm.frx":BF96
      MousePointer    =   99  'Custom
      Picture         =   "client_frm.frx":C2A0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "client_frm.frx":DC22
      Top             =   6240
      Width           =   4575
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   5520
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2880
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax 2"
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
      Height          =   240
      Index           =   11
      Left            =   360
      TabIndex        =   30
      Top             =   4680
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax 1"
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
      Height          =   240
      Index           =   9
      Left            =   360
      TabIndex        =   29
      Top             =   4320
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   120
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address 2"
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
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   26
      Top             =   7080
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id. Client"
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
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email 2"
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
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   5520
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email 1"
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
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   21
      Top             =   5160
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person 2"
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person 1"
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
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telp. 1"
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
      Height          =   240
      Index           =   6
      Left            =   360
      TabIndex        =   18
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address 1"
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
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   6000
      Width           =   1065
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
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telp. 2"
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
      Height          =   240
      Index           =   10
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   735
   End
End
Attribute VB_Name = "client_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        cmd_simpan
        kosong
        buka
        Call client_auto
        Text1.SetFocus
        Label2 = "t"
    Else
        cek_simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
    If Not Data1.Recordset.RecordCount = 0 Then
        cmd_simpan
        buka
        Label2 = "e"
    Else
        x = MsgBox("Data Kosong", vbOKOnly, "Validasi Data")
    End If
    Else
        Data1.Refresh
        tutup
        cmd_awal
    End If
Case 2
    If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin data akan di hapus?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
    Else
        x = MsgBox("Data Kosong...", vbOKOnly, "Validasi Data")
    End If
Case 3
Me.Hide
If Label3.Caption = "lg" Then
    CetakLg_frm.Enabled = True
    CetakLg_frm.Show
ElseIf Label3.Caption = "inv" Then
    InvLG_frm.Enabled = True
    InvLG_frm.Show
ElseIf Label3 = "mm" Then
    main_form.Show
    main_form.Enabled = True
End If
Call isi_cekdat
End Select
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command1(0).Caption <> "SIMPAN" And Not Data1.Recordset.BOF Then
    isi
End If
End Sub

Private Sub Form_Activate()
Data1.Refresh
If Data1.Recordset.BOF Then
    x = MsgBox("Data masih kosong, silahkan diisi terlebih dahulu...", vbOKOnly, "Validasi Data")
    buka
    kosong
    Call client_auto
    Text1.SetFocus
    cmd_simpan
    Label2 = "t"
End If
Call hak_akses1
End Sub

Private Sub Form_Load()
kosong
tutup
Call dbclient
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click (3)
End Sub

Private Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text12 = ""
End Sub

Private Sub isi()
With Data1.Recordset
If Not .BOF Then
    Text1 = !COMPANY
    Text2 = !contact_person1
    Text3 = !contact_person2
    Text4 = !telp1
    Text5 = !telp2
    Text6 = !email1
    Text7 = !email2
    Text8 = !address1
    Text9 = !id_client
    Text10 = !address2
    Text11 = !fax1
    Text12 = !fax2
End If
End With
End Sub

Private Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
End Sub

Private Sub buka()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
End Sub

Private Sub cmd_awal()
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Enabled = True
Command1(3).Enabled = True
End Sub

Private Sub cmd_simpan()
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Enabled = False
Command1(3).Enabled = False
End Sub

Private Sub simpan_data()
With Data1.Recordset
    !COMPANY = Text1
    !contact_person1 = Text2
    !contact_person2 = Text3
    !telp1 = Text4
    !telp2 = Text5
    !email1 = Text6
    !email2 = Text7
    !address1 = Text8
    !address2 = Text10
    !id_client = Text9
    !fax1 = Text11
    !fax2 = Text12
    .Update
End With
End Sub

Private Sub cek_simpan()
Dim y As Boolean
y = False
If Text1 = "" Or Text2 = "" Or Text4 = "" Or Text6 = "" Or Text8 = "" Then
    x = MsgBox("Data belum lengkap...", vbOKOnly, "Validasi Data")
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    ElseIf Text4 = "" Then
        Text4.SetFocus
    ElseIf Text6 = "" Then
        Text6.SetFocus
    ElseIf Text8 = "" Then
        Text8.SetFocus
    End If
Else
    With Data1.Recordset
    If Label2 = "t" Then
            If Not .BOF Then
                .MoveFirst
                Do While Not .EOF
                If Text1 = !COMPANY And Text2 = !contact_person1 Then
                    .MoveLast
                    y = True
                End If
                .MoveNext
                Loop
            End If
            If y = True Then
                x = MsgBox("Data sudah ada, apakah tetap mau disimpan?", vbOKCancel, "Validasi Data")
                If x = vbOK Then
                    .AddNew
                    simpan_data
                    tutup
                    Data1.Refresh
                    cmd_awal
                Else
                    Text1.SetFocus
                End If
            Else
                .AddNew
                simpan_data
                tutup
                Data1.Refresh
                cmd_awal
            End If
    Else
        .Edit
        simpan_data
        tutup
        Data1.Refresh
        cmd_awal
    End If
    End With
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Text11_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub
