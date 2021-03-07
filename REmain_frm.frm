VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form REmain_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REMAINDER"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ClipControls    =   0   'False
   Icon            =   "REmain_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "REmain_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tgl1 As Date
Dim tgl2 As Date

Private Sub Form_Activate()
Data1.RecordSource = "select * from remainder where status=true" ' and tgl <= " & Date  'and waktu<=time order by tgl,waktu asc"
Data1.Refresh
isi_data
End Sub

Private Sub Form_Load()
Call db_remain2
End Sub

Sub isi_data()
Dim head As ColumnHeader
Dim dtl As ListItem
Set head = ListView1.ColumnHeaders.Add(, , "TANGGAL", ListView1.Width / 4)
Set head = ListView1.ColumnHeaders.Add(, , "JAM", ListView1.Width / 4)
Set head = ListView1.ColumnHeaders.Add(, , "NOMOR", ListView1.Width / 4)
Set head = ListView1.ColumnHeaders.Add(, , "KETERANGAN", ListView1.Width / 4 - 100)
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !tgl <= Date Then
            Set dtl = ListView1.ListItems.Add(, , !tgl)
            dtl.SubItems(1) = Format(!waktu, "hh:mm")
            dtl.SubItems(2) = !nomor
            dtl.SubItems(3) = !keterangan
        End If
        .MoveNext
    Loop
End If
End With
End Sub

