VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form dbtiket_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Tiket"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ClipControls    =   0   'False
   Icon            =   "dbtiket_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "D A T A B A S E   T I K E T"
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
      Height          =   1335
      Index           =   3
      Left            =   360
      Picture         =   "dbtiket_frm.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   240
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Penerbangan"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   10095
      Left            =   5760
      TabIndex        =   29
      Top             =   120
      Width           =   9375
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "dbtiket_frm.frx":3994
         Height          =   7335
         Left            =   240
         OleObjectBlob   =   "dbtiket_frm.frx":39A8
         TabIndex        =   36
         Top             =   2400
         Width           =   8895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Hapus"
         DownPicture     =   "dbtiket_frm.frx":437B
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   6480
         MouseIcon       =   "dbtiket_frm.frx":5045
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":534F
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Edit"
         DownPicture     =   "dbtiket_frm.frx":6CD1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   3360
         MouseIcon       =   "dbtiket_frm.frx":799B
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":7CA5
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Tambah"
         DownPicture     =   "dbtiket_frm.frx":9627
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   240
         MouseIcon       =   "dbtiket_frm.frx":A2F1
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":A5FB
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Width           =   2775
      End
      Begin MSDBCtls.DBCombo DBCombo3 
         Bindings        =   "dbtiket_frm.frx":BF7D
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "kode_maskapai"
         Text            =   "DBCombo3"
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5040
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   17
         Text            =   "Text6"
         Top             =   960
         Width           =   975
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "dbtiket_frm.frx":BF91
         Height          =   315
         Left            =   7440
         TabIndex        =   15
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   ""
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "dbtiket_frm.frx":BFA5
         Height          =   315
         Left            =   5040
         TabIndex        =   14
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "kode_lokasi"
         Text            =   "DBCombo1"
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   3960
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Airlines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   10
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flight No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   4
         Left            =   4200
         TabIndex        =   33
         Top             =   480
         Width           =   540
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
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   5
         Left            =   6720
         TabIndex        =   32
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ETD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   6
         Left            =   4200
         TabIndex        =   31
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ETA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   7
         Left            =   6720
         TabIndex        =   30
         Top             =   960
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Airlines"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4455
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   5535
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Tambah"
         DownPicture     =   "dbtiket_frm.frx":BFB9
         Height          =   735
         Index           =   0
         Left            =   120
         MaskColor       =   &H000000FF&
         MouseIcon       =   "dbtiket_frm.frx":CC83
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":CF8D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Edit"
         DownPicture     =   "dbtiket_frm.frx":E00F
         Height          =   735
         Index           =   1
         Left            =   1920
         MaskColor       =   &H000000FF&
         MouseIcon       =   "dbtiket_frm.frx":ECD9
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":EFE3
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Hapus"
         DownPicture     =   "dbtiket_frm.frx":10065
         Height          =   735
         Index           =   2
         Left            =   3720
         MaskColor       =   &H000000FF&
         MouseIcon       =   "dbtiket_frm.frx":10D2F
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":11039
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   465
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   840
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "dbtiket_frm.frx":120BB
         Height          =   2175
         Left            =   120
         OleObjectBlob   =   "dbtiket_frm.frx":120CF
         TabIndex        =   28
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Kode"
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
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Airlines"
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
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Airlines"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Lokasi"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   5535
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         TabIndex        =   38
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "dbtiket_frm.frx":12AA2
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "dbtiket_frm.frx":12AB6
         TabIndex        =   23
         Top             =   2280
         Width           =   5295
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Hapus"
         DownPicture     =   "dbtiket_frm.frx":13489
         Height          =   735
         Index           =   2
         Left            =   3720
         MouseIcon       =   "dbtiket_frm.frx":14153
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":1445D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Edit"
         DownPicture     =   "dbtiket_frm.frx":154DF
         Height          =   735
         Index           =   1
         Left            =   1920
         MouseIcon       =   "dbtiket_frm.frx":161A9
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":164B3
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Tambah"
         DownPicture     =   "dbtiket_frm.frx":17535
         Height          =   735
         Index           =   0
         Left            =   120
         MouseIcon       =   "dbtiket_frm.frx":181FF
         MousePointer    =   99  'Custom
         Picture         =   "dbtiket_frm.frx":18509
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text8"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text7"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Lokasi"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Daerah"
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
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Daerah"
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
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1380
      End
   End
End
Attribute VB_Name = "dbtiket_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub isi_combo1()
Combo1.Clear
Combo1.AddItem ("Domestik")
Combo1.AddItem ("Intenasional")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "Tambah" Then
        cmd_simpan_maskapai
        buka_maskapai
        kosong_maskapai
        Text2.SetFocus
        Label2 = "t"
        Call arl_auto
    Else
        simpan_maskapai
    End If
Case 1
    If Command1(1).Caption = "Edit" Then
        If Not Data1.Recordset.BOF Then
            cmd_simpan_maskapai
            buka_maskapai
            Data1.Recordset.Edit
            Text2.SetFocus
            Label2 = "e"
        Else
            x = MsgBox("data kosong...", vbOKOnly, "Blank Data")
        End If
    Else
        cmd_awal_maskapai
        tutup_maskapai
        Data1.Refresh
    End If
Case 2
    If Not Data1.Recordset.BOF Then
        x = MsgBox("apakah anda yakin data akan di hapus?", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data1.Recordset.Delete
            Data1.Refresh
        End If
    Else
        x = MsgBox("data kosong...", vbOKOnly, "Blank Data")
    End If
End Select
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
    If Command2(0).Caption = "Tambah" Then
        cmd_simpan_lokasi
        buka_lokasi
        kosong_lokasi
        isi_combo1
        Text7.SetFocus
        Label3 = "t"
    Else
        simpan_lokasi
    End If
Case 1
    If Command2(1).Caption = "Edit" Then
        If Data2.Recordset.BOF Then
            x = MsgBox("Data msh kosong...", vbOKOnly, "Validasi Data")
        Else
            cmd_simpan_lokasi
            buka_lokasi
            Label3 = "e"
            Data2.Recordset.Edit
            isi_combo1
            Text7.SetFocus
        End If
    Else
        cmd_awal_lokasi
        Data2.Refresh
        tutup_lokasi
    End If
Case 2
    If Not Data2.Recordset.BOF Then
        x = MsgBox("apakah anda yakin data akan di hapus?", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data2.Recordset.Delete
            Data2.Refresh
        End If
    Else
        x = MsgBox("data kosong...", vbOKOnly, "Blank Data")
    End If
End Select
End Sub

Private Sub Command3_Click(Index As Integer)
Select Case Index
Case 0
    If Command3(0).Caption = "Tambah" Then
        cmd_simpan_penerbangan
        buka_penerbangan
        kosong_penerbangan
        DBCombo3.SetFocus
        Label4 = "t"
    Else
        simpan_penerbangan
        Data3.Refresh
    End If
Case 1
    If Command3(1).Caption = "Edit" Then
        cmd_simpan_penerbangan
        buka_penerbangan
        Data3.Recordset.Edit
        DBCombo3.SetFocus
        Label4 = "e"
    Else
        cmd_awal_penerbangan
        Data3.Refresh
        tutup_penerbangan
    End If
Case 2
    If Not Data3.Recordset.BOF Then
        x = MsgBox("apakah anda yakin data akan di hapus?", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data3.Recordset.Delete
            Data3.Refresh
        End If
    Else
        x = MsgBox("data kosong...", vbOKOnly, "Blank Data")
    End If
End Select
End Sub

Private Sub simpan_maskapai()
Dim a As Boolean
a = False
With Data1.Recordset
If Text2 = "" Or Text3 = "" Then
x = MsgBox("data blm lengkap...", vbOKOnly, "Validasi Data")
    If Text2 = "" Then
        Text2.SetFocus
    ElseIf Text3 = "" Then
        Text3.SetFocus
    End If
Else
    If Label2 = "t" Then
        .MoveFirst
        Do While Not .EOF
            If Text2 = !kode_maskapai Then
                x = MsgBox("kode airlines sudah ada silahkan masukan kode yg lain...", vbOKOnly, "Validasi Data")
                Text2.SetFocus
                a = True
                .MoveLast
            End If
            .MoveNext
        Loop
        If a = False Then
            .AddNew
            !kode_maskapai = Text2
            !nama_maskapai = Text3
            !no_maskapai = Text1
            .Update
            cmd_awal_maskapai
            tutup_maskapai
            Data1.Refresh
        End If
    Else
        !kode_maskapai = Text2
        !nama_maskapai = Text3
        !no_maskapai = Text1
        .Update
        cmd_awal_maskapai
        tutup_maskapai
        Data1.Refresh
    End If
End If
End With
End Sub

Private Sub simpan_lokasi()
Dim a As Boolean
a = False
With Data2.Recordset
If Text7 = "" Or Text8 = "" Or Combo1 = "" Then
    x = MsgBox("data blm lengkap...", vbOKOnly, "Validasi Data")
    If Text7 = "" Then
        Text7.SetFocus
    ElseIf Text8 = "" Then
        Text8.SetFocus
    ElseIf Combo1 = "" Then
        Combo1.SetFocus
    End If
Else
    If Label3 = "t" Then
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If Text7 = !kode_lokasi Then
                    x = MsgBox("Kode lokasi sudah ada, silahkan masukan kode lokasi yg lain...", vbOKOnly, "Validasi Data")
                    Text7.SetFocus
                    .MoveLast
                    a = True
                End If
                .MoveNext
            Loop
        End If
        If a = False Then
            .AddNew
            !kode_lokasi = Text7
            !nama_lokasi = Text8
            !status_lokasi = Combo1
            .Update
            cmd_awal_lokasi
            tutup_lokasi
            Data2.Refresh
        End If
    Else
        !kode_lokasi = Text7
        !nama_lokasi = Text8
        !status_lokasi = Combo1
        .Update
        cmd_awal_lokasi
        tutup_lokasi
        Data2.Refresh
    End If
End If
End With
End Sub

Private Sub simpan_penerbangan()
Dim a As Boolean
a = False
With Data3.Recordset
If Text4 = "" Or Text5 = "" Or Text6 = "" Or DBCombo1 = "" Or DBCombo2 = "" Or DBCombo3 = "" Then
    x = MsgBox("data blm lengkap...", vbOKOnly, "Validasi Data")
    If Text4 = "" Then
        Text4.SetFocus
    ElseIf Text5 = "" Then
        Text5.SetFocus
    ElseIf Text6 = "" Then
        Text6.SetFocus
    ElseIf DBCombo1 = "" Then
        DBCombo1.SetFocus
    ElseIf DBCombo2 = "" Then
        DBCombo2.SetFocus
    ElseIf DBCombo3 = "" Then
        DBCombo3.SetFocus
    End If
Else
    If Label4 = "t" Then
        If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If DBCombo3 = !kode_maskapai And Text4 = !no_penerbangan Then
                x = MsgBox("Data sudah ada, silahkan masukan data yg lain...", vbOKOnly, "Validasi Data")
                Text4.SetFocus
                .MoveLast
                a = True
            End If
            .MoveNext
        Loop
        End If
        If a = False Then
            .AddNew
            !kode_maskapai = DBCombo3
            !no_penerbangan = Text4
            !From = DBCombo1
            !To = DBCombo2
            !dep = Text5
            !arr = Text6
            .Update
            cmd_awal_penerbangan
            tutup_penerbangan
        End If
    Else
        !kode_maskapai = DBCombo3
        !no_penerbangan = Text4
        !From = DBCombo1
        !To = DBCombo2
        !dep = Text5
        !arr = Text6
        .Update
    End If
End If
End With
End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub DBCombo2_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub DBCombo3_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command1(0).Caption <> "Simpan" And Not Data1.Recordset.BOF Then
    isi_maskapai
End If
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command2(0).Caption <> "Simpan" And Not Data2.Recordset.BOF Then
    isi_lokasi
End If
End Sub

Private Sub DBGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command3(0).Caption <> "Simpan" And Not Data3.Recordset.BOF Then
    isi_penerbangan
End If
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
Data3.Refresh
If Data1.Recordset.BOF Then
    x = MsgBox("data masih kosong, silahkan diisi terlebih dahulu...", vbOKOnly, "Blank Data")
    Data1.Recordset.AddNew
    buka_maskapai
    kosong_maskapai
    Call arl_auto
    cmd_simpan_maskapai
    Text2.SetFocus
Else
    isi_maskapai
    tutup_maskapai
    cek_lokasi
End If
End Sub

Private Sub cek_lokasi()
If Data2.Recordset.BOF Then
    x = MsgBox("data masih kosong, silahkan diisi terlebih dahulu...", vbOKOnly, "Blank Data")
    Data2.Recordset.AddNew
    buka_lokasi
    kosong_lokasi
    isi_combo1
    Text7.SetFocus
    cmd_simpan_lokasi
    Label3 = "t"
Else
    isi_lokasi
    tutup_lokasi
    cek_penerbangan
End If
End Sub

Private Sub cek_penerbangan()
If Data3.Recordset.BOF Then
    x = MsgBox("data masih kosong, silahkan diisi terlebih dahulu...", vbOKOnly, "Blank Data")
    Data3.Recordset.AddNew
    buka_penerbangan
    kosong_penerbangan
    DBCombo3.SetFocus
    cmd_simpan_penerbangan
    Label4 = "t"
Else
    isi_penerbangan
    tutup_penerbangan
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
End If
End Sub

Private Sub Form_Load()
kosong_maskapai
kosong_lokasi
kosong_penerbangan
isi_combo1
tutup_maskapai
tutup_lokasi
tutup_penerbangan
Call dbtiket
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
main_form.Show
main_form.Enabled = True
End Sub

Private Sub buka_maskapai()
'Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
End Sub

Private Sub buka_lokasi()
Text7.Enabled = True
Text8.Enabled = True
Combo1.Enabled = True
End Sub

Private Sub buka_penerbangan()
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
DBCombo1.Enabled = True
DBCombo2.Enabled = True
DBCombo3.Enabled = True
End Sub

Private Sub tutup_maskapai()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub

Private Sub tutup_lokasi()
Text7.Enabled = False
Text8.Enabled = False
Combo1.Enabled = False
End Sub

Private Sub tutup_penerbangan()
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
DBCombo1.Enabled = False
DBCombo2.Enabled = False
DBCombo3.Enabled = False
End Sub

Private Sub kosong_maskapai()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Private Sub kosong_lokasi()
Text7 = ""
Text8 = ""
Combo1 = ""
End Sub

Private Sub kosong_penerbangan()
Text4 = ""
Text5 = ""
Text6 = ""
DBCombo1 = ""
DBCombo2 = ""
DBCombo3 = ""
End Sub

Private Sub isi_maskapai()
Text1 = Data1.Recordset!no_maskapai
Text2 = Data1.Recordset!kode_maskapai
Text3 = Data1.Recordset!nama_maskapai
End Sub

Private Sub isi_lokasi()
Text7 = Data2.Recordset!kode_lokasi
Text8 = Data2.Recordset!nama_lokasi
Combo1 = Data2.Recordset!status_lokasi
End Sub

Private Sub isi_penerbangan()
Text4 = Data3.Recordset!no_penerbangan
Text5 = Format(Data3.Recordset!dep, "hh:mm")
Text6 = Format(Data3.Recordset!arr, "hh:mm")
DBCombo1 = Data3.Recordset!From
DBCombo2 = Data3.Recordset!To
DBCombo3 = Data3.Recordset!kode_maskapai
End Sub

Private Sub cmd_simpan_maskapai()
Command1(0).Caption = "Simpan"
Command1(1).Caption = "Batal"
Command1(2).Enabled = False
End Sub

Private Sub cmd_simpan_lokasi()
Command2(0).Caption = "Simpan"
Command2(1).Caption = "Batal"
Command2(2).Enabled = False
End Sub

Private Sub cmd_simpan_penerbangan()
Command3(0).Caption = "Simpan"
Command3(1).Caption = "Batal"
Command3(2).Enabled = False
End Sub

Private Sub cmd_awal_maskapai()
Command1(0).Caption = "Tambah"
Command1(1).Caption = "Edit"
Command1(2).Enabled = True
End Sub

Private Sub cmd_awal_lokasi()
Command2(0).Caption = "Tambah"
Command2(1).Caption = "Edit"
Command2(2).Enabled = True
End Sub

Private Sub cmd_awal_penerbangan()
Command3(0).Caption = "Tambah"
Command3(1).Caption = "Edit"
Command3(2).Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
