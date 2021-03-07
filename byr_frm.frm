VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form byr_frm 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pembayaran"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "byr_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   14420
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   8421376
      TabCaption(0)   =   "Letter of Guarantee"
      TabPicture(0)   =   "byr_frm.frx":1982
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text11(0)"
      Tab(0).Control(1)=   "Data2(0)"
      Tab(0).Control(2)=   "Command1(2)"
      Tab(0).Control(3)=   "Command1(1)"
      Tab(0).Control(4)=   "Text10(0)"
      Tab(0).Control(5)=   "Text9(0)"
      Tab(0).Control(6)=   "Text8(0)"
      Tab(0).Control(7)=   "Text7(0)"
      Tab(0).Control(8)=   "Text6(0)"
      Tab(0).Control(9)=   "DBGrid1(0)"
      Tab(0).Control(10)=   "Data1(0)"
      Tab(0).Control(11)=   "Command1(0)"
      Tab(0).Control(12)=   "Text5(0)"
      Tab(0).Control(13)=   "Text4(0)"
      Tab(0).Control(14)=   "Text3(0)"
      Tab(0).Control(15)=   "Text2(0)"
      Tab(0).Control(16)=   "Text1(0)"
      Tab(0).Control(17)=   "Label1(7)"
      Tab(0).Control(18)=   "Label1(4)"
      Tab(0).Control(19)=   "Label1(6)"
      Tab(0).Control(20)=   "Label1(5)"
      Tab(0).Control(21)=   "Label1(3)"
      Tab(0).Control(22)=   "Label1(2)"
      Tab(0).Control(23)=   "Label1(1)"
      Tab(0).Control(24)=   "Label1(0)"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Invoice Tiket"
      TabPicture(1)   =   "byr_frm.frx":199E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(8)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(11)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(13)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(14)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(15)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DBGrid1(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text1(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text2(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text3(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text4(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text5(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Command1(3)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Data1(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text6(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text7(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text8(1)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text9(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text10(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Command1(4)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Command1(5)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Data2(1)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text11(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "Invoice Lain-lain"
      TabPicture(2)   =   "byr_frm.frx":19BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(16)"
      Tab(2).Control(1)=   "Label1(17)"
      Tab(2).Control(2)=   "Label1(18)"
      Tab(2).Control(3)=   "Label1(19)"
      Tab(2).Control(4)=   "Label1(20)"
      Tab(2).Control(5)=   "Label1(21)"
      Tab(2).Control(6)=   "Label1(22)"
      Tab(2).Control(7)=   "Label1(23)"
      Tab(2).Control(8)=   "DBGrid1(2)"
      Tab(2).Control(9)=   "Text1(2)"
      Tab(2).Control(10)=   "Text2(2)"
      Tab(2).Control(11)=   "Text3(2)"
      Tab(2).Control(12)=   "Text4(2)"
      Tab(2).Control(13)=   "Text5(2)"
      Tab(2).Control(14)=   "Command1(6)"
      Tab(2).Control(15)=   "Data1(2)"
      Tab(2).Control(16)=   "Text6(2)"
      Tab(2).Control(17)=   "Text7(2)"
      Tab(2).Control(18)=   "Text8(2)"
      Tab(2).Control(19)=   "Text9(2)"
      Tab(2).Control(20)=   "Text10(2)"
      Tab(2).Control(21)=   "Command1(7)"
      Tab(2).Control(22)=   "Command1(8)"
      Tab(2).Control(23)=   "Data2(2)"
      Tab(2).Control(24)=   "Text11(2)"
      Tab(2).ControlCount=   25
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   69
         Text            =   "Text11"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   68
         Text            =   "Text11"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   67
         Text            =   "Text11"
         Top             =   600
         Width           =   3615
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   2
         Left            =   -72120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -72120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CETAK"
         Height          =   735
         Index           =   8
         Left            =   -71040
         Picture         =   "byr_frm.frx":19D6
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BAYAR"
         Height          =   735
         Index           =   7
         Left            =   -72840
         Picture         =   "byr_frm.frx":2A58
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   64
         Text            =   "Text10"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   63
         Text            =   "Text9"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   62
         Text            =   "Text8"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   61
         Text            =   "Text7"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -72120
         TabIndex        =   59
         Text            =   "Text6"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Data Data1 
         BackColor       =   &H00000000&
         Caption         =   "Data Pembayaran Invoice Lain-lain"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   2
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4320
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CARI"
         Height          =   735
         Index           =   6
         Left            =   -74640
         Picture         =   "byr_frm.frx":3ADA
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -72120
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -72120
         TabIndex        =   53
         Text            =   "Text4"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   52
         Text            =   "Text3"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CETAK"
         Height          =   735
         Index           =   5
         Left            =   3960
         Picture         =   "byr_frm.frx":4B5C
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BAYAR"
         Height          =   735
         Index           =   4
         Left            =   2160
         Picture         =   "byr_frm.frx":5BDE
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   42
         Text            =   "Text10"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   41
         Text            =   "Text9"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   40
         Text            =   "Text8"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   39
         Text            =   "Text7"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   37
         Text            =   "Text6"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Data Data1 
         BackColor       =   &H00000000&
         Caption         =   "Data Pembayaran Invoice Tiket"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   1
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4320
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CARI"
         Height          =   735
         Index           =   3
         Left            =   360
         Picture         =   "byr_frm.frx":6C60
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   31
         Text            =   "Text4"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CETAK"
         Height          =   735
         Index           =   2
         Left            =   -71040
         Picture         =   "byr_frm.frx":7CE2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BAYAR"
         Height          =   735
         Index           =   1
         Left            =   -72840
         Picture         =   "byr_frm.frx":8D64
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   20
         Text            =   "Text10"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   19
         Text            =   "Text9"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   18
         Text            =   "Text8"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   17
         Text            =   "Text7"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -72120
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   2760
         Width           =   2655
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "byr_frm.frx":9DE6
         Height          =   3135
         Index           =   0
         Left            =   -74640
         OleObjectBlob   =   "byr_frm.frx":9DFD
         TabIndex        =   13
         Top             =   4800
         Width           =   5175
      End
      Begin VB.Data Data1 
         BackColor       =   &H00000000&
         Caption         =   "Data Pembayaran LG"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   0
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4320
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CARI"
         Height          =   735
         Index           =   0
         Left            =   -74640
         Picture         =   "byr_frm.frx":A7E7
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -72120
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -72120
         TabIndex        =   9
         Text            =   "Text4"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   -73080
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "byr_frm.frx":B869
         Height          =   3135
         Index           =   1
         Left            =   360
         OleObjectBlob   =   "byr_frm.frx":B880
         TabIndex        =   35
         Top             =   4800
         Width           =   5175
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "byr_frm.frx":C276
         Height          =   3135
         Index           =   2
         Left            =   -74640
         OleObjectBlob   =   "byr_frm.frx":C28D
         TabIndex        =   57
         Top             =   4800
         Width           =   5175
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frek. Pembayaran"
         Height          =   255
         Index           =   23
         Left            =   -74640
         TabIndex        =   60
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sisa Bayar"
         Height          =   255
         Index           =   22
         Left            =   -74640
         TabIndex        =   58
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Transaksi"
         Height          =   255
         Index           =   21
         Left            =   -74640
         TabIndex        =   54
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Pembayaran"
         Height          =   255
         Index           =   20
         Left            =   -74640
         TabIndex        =   49
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contact Person"
         Height          =   255
         Index           =   19
         Left            =   -74640
         TabIndex        =   48
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nama Client"
         Height          =   255
         Index           =   18
         Left            =   -74640
         TabIndex        =   47
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Invoice"
         Height          =   255
         Index           =   17
         Left            =   -74640
         TabIndex        =   46
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No. Invoice"
         Height          =   255
         Index           =   16
         Left            =   -74640
         TabIndex        =   45
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frek. Pembayaran"
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   38
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sisa Bayar"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   36
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Transaksi"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   32
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Pembayaran"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contact Person"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nama Client"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Invoice"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No. Invoice"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frek. Pembayaran"
         Height          =   255
         Index           =   7
         Left            =   -74640
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sisa Bayar"
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   14
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Transaksi"
         Height          =   255
         Index           =   6
         Left            =   -74640
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Pembayaran"
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   5
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contact Person"
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nama Supplier"
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal LG"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO. LG"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "byr_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Dim a, b, c As Single
Select Case Index
Case 0
    Me.Enabled = False
    CariByr_frm.Show
    CariByr_frm.Label1(1).Caption = "CARI DATA LG"
Case 1
    If Text11(0) <> "" And Not Data2(0).Recordset.BOF Then
        On Error GoTo salah:
        a = InputBox("Masukan Nominal Pembayaran", "Bayar LG No." & Text11(0))
        If a <> blank Then
            With Data1(0).Recordset
                .AddNew
                !no_lg = Text11(0)
                !tgl_byr = Date
                !curr = Text8(0)
                !Nominal = a
                .Update
                Data1(0).Refresh
                isi_LG
                update_bayar
            End With
        End If
    End If
Case 2
    With CetakByr_frm
        .Label1(0).Caption = "No. LG"
        .Label1(1).Caption = "Tanggal LG"
        .Label1(2).Caption = "Nama Supplier"
        Me.Enabled = False
        .Show
    End With
Case 3
    Me.Enabled = False
    CariByr_frm.Show
    CariByr_frm.Label1(1).Caption = "CARI DATA INVOICE TIKET"
Case 4
    If Text11(1) <> "" And Not Data2(1).Recordset.BOF Then
        On Error GoTo salah:
        b = InputBox("Masukan Nominal Pembayaran", "Bayar Invoice Tiket No." & Text11(1))
        If b <> blank Then
            With Data1(1).Recordset
                .AddNew
                !no_invoice = Text11(1)
                !tgl_byr = Date
                !curr = Text8(1)
                !Nominal = b
                .Update
                Data1(1).Refresh
                isi_invtiket
                update_bayar
            End With
        End If
    End If
Case 5
    With CetakByr_frm
        .Label1(0).Caption = "No. Invoice"
        .Label1(1).Caption = "Tanggal Invoice"
        .Label1(2).Caption = "Nama Client"
        Me.Enabled = False
        .Show
    End With
Case 6
    Me.Enabled = False
    CariByr_frm.Show
    CariByr_frm.Label1(1).Caption = "CARI DATA INVOICE"
Case 7
    If Text11(2) <> "" And Not Data2(2).Recordset.BOF Then
        On Error GoTo salah:
        c = InputBox("Masukan Nominal Pembayaran", "Bayar Invoice No." & Text11(2))
        If c <> blank Then
            With Data1(2).Recordset
                .AddNew
                !no_invoice = Text11(2)
                !tgl_byr = Date
                !curr = Text8(2)
                !Nominal = c
                .Update
                Data1(2).Refresh
                isi_invnt
                update_bayar
            End With
        End If
    End If
Case 8
    With CetakByr_frm
        .Label1(0).Caption = "No. Invoice"
        .Label1(1).Caption = "Tanggal Invoice"
        .Label1(2).Caption = "Nama Client"
        Me.Enabled = False
        .Show
    End With
End Select
salah:
End Sub

Private Sub Form_Load()
Call db_byr
Data2(0).RecordSource = "db_log"
Data2(1).RecordSource = "db_invoice"
Data2(2).RecordSource = "inv_nt"

kosong
End Sub

Sub kosong_lg()
Text11(0) = ""
Text1(0) = ""
Text2(0) = ""
Text3(0) = ""
Text4(0) = ""
Text5(0) = ""
Text6(0) = ""
Text7(0) = ""
Text8(0) = ""
Text9(0) = ""
Text10(0) = ""
End Sub

Sub kosong_Tiket()
Text11(1) = ""
Text1(1) = ""
Text2(1) = ""
Text3(1) = ""
Text4(1) = ""
Text5(1) = ""
Text6(1) = ""
Text7(1) = ""
Text8(1) = ""
Text9(1) = ""
Text10(1) = ""
End Sub

Sub kosong_NT()
Text11(2) = ""
Text1(2) = ""
Text2(2) = ""
Text3(2) = ""
Text4(2) = ""
Text5(2) = ""
Text6(2) = ""
Text7(2) = ""
Text8(2) = ""
Text9(2) = ""
Text10(2) = ""
End Sub

Sub kosong()
kosong_lg
kosong_Tiket
kosong_NT
End Sub

Private Sub Form_Unload(Cancel As Integer)
main_form.Enabled = True
main_form.Show
End Sub

Sub isi_LG()
Dim x, y As Single
Data2(0).Refresh
Data1(0).RecordSource = "select tgl_byr,curr,nominal,no_lg from bayar_lg where no_lg='" & Text11(0) & "'"
Data1(0).Refresh
With Data2(0).Recordset
If Not .BOF Then
    Text1(0) = !tgl_lg
    Text2(0) = !COMPANY
    Text3(0) = !contact_person
    Text8(0) = !Currency
    Text9(0) = Text8(0)
    Text10(0) = Text8(0)
    y = 0
    Do While Not .EOF
        y = y + !harga
        .MoveNext
    Loop
    Text4(0) = Format(y, "###,###,###.00")
    If Not Data1(0).Recordset.BOF Then
        Text7(0) = Data1(0).Recordset.RecordCount
        Data1(0).Recordset.MoveFirst
        x = 0
        Do While Not Data1(0).Recordset.EOF
            x = x + Val(Data1(0).Recordset!Nominal)
            Data1(0).Recordset.MoveNext
        Loop
        Text5(0) = Format(x, "###,###,###.00")
        Text6(0) = Format(y - x, "###,###,###.00")
    Else
        Text7(0) = 0
        Text5(0) = 0
        Text6(0) = 0
    End If
Else
    Text1(0) = ""
    Text2(0) = ""
    Text3(0) = ""
    Text4(0) = ""
    Text5(0) = ""
    Text6(0) = ""
    Text7(0) = ""
    Text8(0) = ""
    Text9(0) = ""
    Text10(0) = ""
End If
End With
End Sub

Sub isi_invtiket()
Dim x, y As Single
Data2(1).Refresh
Data1(1).RecordSource = "select tgl_byr,curr,nominal,no_invoice from byr_invtiket where no_invoice='" & Text11(1) & "'"
Data1(1).Refresh
With Data2(1).Recordset
If Not .BOF Then
    Text1(1) = !tgl_inv
    Text2(1) = !COMPANY
    Text3(1) = !contact_person
    Text8(1) = !Currency
    Text9(1) = Text8(1)
    Text10(1) = Text8(1)
    .MoveFirst
    y = 0
    Do While Not .EOF
        y = y + !hrg_tiket
        .MoveNext
    Loop
    Text4(1) = Format(y, "###,###,###.00")
    If Not Data1(1).Recordset.BOF Then
        Text7(1) = Data1(1).Recordset.RecordCount
        Data1(1).Recordset.MoveFirst
        x = 0
        Do While Not Data1(1).Recordset.EOF
            x = x + Val(Data1(1).Recordset!Nominal)
            Data1(1).Recordset.MoveNext
        Loop
        Text5(1) = Format(x, "###,###,###.00")
        Text6(1) = Format(y - x, "###,###,###.00")
    Else
        Text7(1) = 0
        Text5(1) = 0
        Text6(1) = 0
    End If
Else
    Text1(1) = ""
    Text2(1) = ""
    Text3(1) = ""
    Text4(1) = ""
    Text5(1) = ""
    Text6(1) = ""
    Text7(1) = ""
    Text8(1) = ""
    Text9(1) = ""
    Text10(1) = ""
End If
End With
End Sub

Sub isi_invnt()
Dim x, y As Single
Data2(2).Refresh
Data1(2).RecordSource = "select tgl_byr,curr,nominal,no_invoice from byr_invnt where no_invoice='" & Text11(2) & "'"
Data1(2).Refresh
With Data2(2).Recordset
If Not .BOF Then
    Text1(2) = !tgl
    Text2(2) = !COMPANY
    Text3(2) = !contact_person
    Text8(2) = !Currency
    Text9(2) = Text8(2)
    Text10(2) = Text8(2)
    .MoveFirst
    y = 0
    Do While Not .EOF
        y = y + !Nominal
        .MoveNext
    Loop
    Text4(2) = Format(y, "###,###,###.00")
    If Not Data1(2).Recordset.BOF Then
        Text7(2) = Data1(2).Recordset.RecordCount
        Data1(2).Recordset.MoveFirst
        x = 0
        Do While Not Data1(2).Recordset.EOF
            x = x + Val(Data1(2).Recordset!Nominal)
            Data1(2).Recordset.MoveNext
        Loop
        Text5(2) = Format(x, "###,###,###.00")
        Text6(2) = Format(y - x, "###,###,###.00")
    Else
        Text7(2) = 0
        Text5(2) = 0
        Text6(2) = 0
    End If
Else
    Text1(2) = ""
    Text2(2) = ""
    Text3(2) = ""
    Text4(2) = ""
    Text5(2) = ""
    Text6(2) = ""
    Text7(2) = ""
    Text8(2) = ""
    Text9(2) = ""
    Text10(2) = ""
End If
End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text10_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text11_Change(Index As Integer)
Select Case Index
Case 0
    Data2(0).RecordSource = "select * from db_log where no_lg='" & Text11(0) & "'"
    Data2(0).Refresh
    Data3.RecordSource = "select * from pembayaran where nomor ='" & Text11(0) & "'"
    isi_LG
Case 1
    Data2(1).RecordSource = "select * from db_invoice where no_invoice='" & Text11(1) & "'"
    Data2(1).Refresh
    Data3.RecordSource = "select * from pembayaran where nomor ='" & Text11(1) & "'"
    isi_invtiket
Case 2
    Data2(2).RecordSource = "select * from inv_nt where no_inv='" & Text11(2) & "'"
    Data2(2).Refresh
    Data3.RecordSource = "select * from pembayaran where nomor ='" & Text11(2) & "'"
    isi_invnt
End Select
Data3.Refresh
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text9_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Sub update_bayar()
Dim x As Byte
x = SSTab1.Tab
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    .Edit
    !total_bayar = Val(Format(Text5(x), "###"))
    !sisa = Val(Format(Text6(x), "###"))
    !frek = Val(Text7(x))
    .Update
Else
    .AddNew
    !nomor = Text11(x)
    !curr = Text8(x)
    !nilai_trans = Format(Text4(x), "###")
    !total_bayar = Format(Text5(x), "###")
    !sisa = Format(Text6(x), "###")
    !frek = Val(Text7(x))
    .Update
End If
End With
Data3.Refresh
End Sub
