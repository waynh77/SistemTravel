VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CariByr_frm 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cari Data"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ClipControls    =   0   'False
   Icon            =   "CariByr_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CariByr_frm.frx":1982
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "CariByr_frm.frx":1996
      TabIndex        =   4
      Top             =   1320
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "BATAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "AMBIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MASUKAN KATA KUNCI :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MASUKAN KATA KUNCI :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "CariByr_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    With Data1.Recordset
        If Text1 <> "" Then
        If Not .BOF Then
            Select Case Label1(1).Caption
                Case "CARI DATA LG"
                    byr_frm.Text11(0) = Data1.Recordset!no_lg
                Case "CARI DATA INVOICE TIKET"
                    byr_frm.Text11(1) = Data1.Recordset!no_invoice
                Case "CARI DATA INVOICE"
                    byr_frm.Text11(2) = Data1.Recordset!no_inv
            End Select
            Unload Me
        Else
            MsgBox "Data Kosong", vbCritical, "Blank Data"
        End If
        Else
            MsgBox "Anda belum memasukan kata kunci", vbCritical, "Validasi Input"
        End If
    End With
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
Call db_cari
Text1 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
byr_frm.Enabled = True
End Sub

Private Sub Text1_Change()
Select Case Label1(1).Caption
Case "CARI DATA LG"
    Data1.RecordSource = "select * from db_log where no_lg like '*" & Text1 & "*' or company like '*" & Text1 & "*'"
Case "CARI DATA INVOICE TIKET"
    Data1.RecordSource = "select * from db_invoice where no_invoice like '*" & Text1 & "*' or company like '*" & Text1 & "*'"
Case "CARI DATA INVOICE"
    Data1.RecordSource = "select * from inv_nt where no_inv like '*" & Text1 & "*' or company like '*" & Text1 & "*'"
End Select
Data1.Refresh
End Sub
