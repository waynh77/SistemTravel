VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form InvLG_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK INVOICE BERDASARKAN LG"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "InvLG_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12240
      TabIndex        =   96
      Top             =   6240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58589185
      CurrentDate     =   39636
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   960
      TabIndex        =   92
      Top             =   0
      Width           =   3975
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I N V O I C E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   93
         Top             =   240
         Width           =   1740
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "InvLG_frm.frx":1CCA
      Left            =   9960
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Documents and Settings\wahyu\My Documents\waynh project\vb project\travel project\cetak_inv.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "K E L U A R"
      DownPicture     =   "InvLG_frm.frx":1CDE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12600
      MouseIcon       =   "InvLG_frm.frx":29A8
      MousePointer    =   99  'Custom
      Picture         =   "InvLG_frm.frx":2CB2
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   4155
      Left            =   9840
      TabIndex        =   83
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "CETAK INVOICE"
      DownPicture     =   "InvLG_frm.frx":397C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      MouseIcon       =   "InvLG_frm.frx":4646
      MousePointer    =   99  'Custom
      Picture         =   "InvLG_frm.frx":4950
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Pembeli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   120
      TabIndex        =   55
      Top             =   6600
      Width           =   15015
      Begin VB.TextBox Text24 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   7560
         TabIndex        =   91
         Text            =   "Text24"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   7560
         TabIndex        =   90
         Text            =   "Text22"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "EDIT CLIENT"
         Height          =   615
         Index           =   1
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox text9 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   9960
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "InvLG_frm.frx":561A
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox text1 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   9960
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "InvLG_frm.frx":5622
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox text7 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11160
         TabIndex        =   29
         Text            =   "Text7"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox text6 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11160
         TabIndex        =   28
         Text            =   "Text6"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox text2 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox text3 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox text4 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   26
         Text            =   "Text4"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox text5 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox combo1 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "CLIENT BARU"
         DownPicture     =   "InvLG_frm.frx":5628
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
         Left            =   9960
         MouseIcon       =   "InvLG_frm.frx":62F2
         MousePointer    =   99  'Custom
         Picture         =   "InvLG_frm.frx":65FC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3240
         Width           =   4815
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "InvLG_frm.frx":72C6
         Height          =   1935
         Left            =   240
         OleObjectBlob   =   "InvLG_frm.frx":72DA
         TabIndex        =   56
         Top             =   2280
         Width           =   9375
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
         Index           =   46
         Left            =   6360
         TabIndex        =   89
         Top             =   1680
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
         Index           =   45
         Left            =   6360
         TabIndex        =   88
         Top             =   1320
         Width           =   570
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
         Index           =   5
         Left            =   9960
         TabIndex        =   66
         Top             =   2160
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
         Index           =   6
         Left            =   240
         TabIndex        =   65
         Top             =   480
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
         Index           =   7
         Left            =   9960
         TabIndex        =   64
         Top             =   840
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
         Index           =   8
         Left            =   9960
         TabIndex        =   63
         Top             =   480
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
         Index           =   10
         Left            =   240
         TabIndex        =   62
         Top             =   1680
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
         Index           =   26
         Left            =   240
         TabIndex        =   61
         Top             =   1320
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
         Index           =   27
         Left            =   6360
         TabIndex        =   60
         Top             =   480
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
         Index           =   28
         Left            =   9960
         TabIndex        =   59
         Top             =   1320
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
         Index           =   29
         Left            =   240
         TabIndex        =   58
         Top             =   840
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
         Index           =   30
         Left            =   6360
         TabIndex        =   57
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Tiket"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9615
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   8640
         TabIndex        =   12
         Text            =   "Combo6"
         Top             =   960
         Width           =   855
      End
      Begin VB.Data Data1 
         BackColor       =   &H00000000&
         Caption         =   "Data Letter of Guarantee"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   345
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Width           =   4695
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2415
         Left            =   1800
         TabIndex        =   84
         Top             =   2520
         Visible         =   0   'False
         Width           =   3495
         _Version        =   524288
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2007
         Month           =   12
         Day             =   27
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "InvLG_frm.frx":7CAD
         Height          =   1455
         Left            =   4800
         OleObjectBlob   =   "InvLG_frm.frx":7CC1
         TabIndex        =   86
         Top             =   4200
         Width           =   4695
      End
      Begin VB.Data Data4 
         BackColor       =   &H00000000&
         Caption         =   "Data Tambah Tiket"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3720
         Width           =   4695
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   6120
         TabIndex        =   8
         Text            =   "Combo5"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "Combo4"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Combo3"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "TAMBAH TIKET"
         DownPicture     =   "InvLG_frm.frx":8694
         Height          =   735
         Index           =   2
         Left            =   120
         MouseIcon       =   "InvLG_frm.frx":935E
         MousePointer    =   99  'Custom
         Picture         =   "InvLG_frm.frx":9668
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4920
         Width           =   2055
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Text            =   "Text23"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8640
         TabIndex        =   21
         Text            =   "Text21"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   20
         Text            =   "Text20"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Text            =   "Text19"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Text            =   "Text18"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Text            =   "Text17"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Text            =   "Text16"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "EDIT TIKET"
         DownPicture     =   "InvLG_frm.frx":A6EA
         Height          =   735
         Index           =   3
         Left            =   2400
         MouseIcon       =   "InvLG_frm.frx":B3B4
         MousePointer    =   99  'Custom
         Picture         =   "InvLG_frm.frx":B6BE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4920
         Width           =   2055
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Text            =   "Text10"
         Top             =   4200
         Width           =   4335
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Text            =   "Text11"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text12"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "Text13"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6960
         TabIndex        =   9
         Text            =   "Text14"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Text            =   "Text15"
         Top             =   960
         Width           =   1935
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "InvLG_frm.frx":C740
         Height          =   1455
         Left            =   4800
         OleObjectBlob   =   "InvLG_frm.frx":C754
         TabIndex        =   35
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   3840
         TabIndex        =   85
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Curr"
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
         Index           =   44
         Left            =   8160
         TabIndex        =   82
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Inv"
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
         Index           =   43
         Left            =   4680
         TabIndex        =   81
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dep"
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
         Left            =   240
         TabIndex        =   54
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Time"
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
         Index           =   16
         Left            =   3240
         TabIndex        =   53
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Time"
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
         Index           =   15
         Left            =   960
         TabIndex        =   52
         Top             =   3120
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Loc"
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
         Index           =   14
         Left            =   3240
         TabIndex        =   51
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Loc"
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
         Left            =   960
         TabIndex        =   50
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         TabIndex        =   49
         Top             =   3960
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   120
         TabIndex        =   48
         Top             =   3600
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arr"
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
         Index           =   17
         Left            =   2760
         TabIndex        =   47
         Top             =   3120
         Width           =   315
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
         Index           =   18
         Left            =   2760
         TabIndex        =   46
         Top             =   2760
         Width           =   300
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   19
         Left            =   240
         TabIndex        =   45
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flight Date"
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
         Index           =   20
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
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
         Index           =   21
         Left            =   120
         TabIndex        =   43
         Top             =   1920
         Width           =   600
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   22
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1005
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   23
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ticket"
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
         Index           =   24
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1080
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   25
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   1470
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   120
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passanger "
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
         Index           =   31
         Left            =   4680
         TabIndex        =   38
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga LG"
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
         Index           =   32
         Left            =   4680
         TabIndex        =   37
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Curr"
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
         Index           =   33
         Left            =   8160
         TabIndex        =   36
         Top             =   960
         Width           =   435
      End
   End
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Text            =   "Combo8"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Jatuh Tempo"
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
      Index           =   47
      Left            =   9840
      TabIndex        =   95
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   5520
      TabIndex        =   94
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "InvLG_frm.frx":D127
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
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
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   42
      Left            =   13200
      TabIndex        =   80
      Top             =   5880
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal"
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
      Index           =   41
      Left            =   12000
      TabIndex        =   79
      Top             =   5880
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga Invoice"
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
      Index           =   40
      Left            =   9840
      TabIndex        =   78
      Top             =   5880
      Width           =   2070
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
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   39
      Left            =   13200
      TabIndex        =   77
      Top             =   5520
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal"
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
      Index           =   38
      Left            =   12000
      TabIndex        =   76
      Top             =   5520
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga LG"
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
      Index           =   37
      Left            =   9840
      TabIndex        =   75
      Top             =   5520
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl LG"
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
      Index           =   36
      Left            =   2040
      TabIndex        =   74
      Top             =   480
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl LG"
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
      Index           =   35
      Left            =   960
      TabIndex        =   73
      Top             =   480
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. LG"
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
      Index           =   34
      Left            =   960
      TabIndex        =   72
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Date"
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
      Left            =   8160
      TabIndex        =   71
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Invoice :"
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
      Left            =   6720
      TabIndex        =   70
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Number"
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
      Index           =   3
      Left            =   8160
      TabIndex        =   69
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation Detail "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   9840
      TabIndex        =   68
      Top             =   5280
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date             :"
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
      Left            =   6720
      TabIndex        =   67
      Top             =   480
      Width           =   1350
   End
End
Attribute VB_Name = "InvLG_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub isi_dtltambah()
With Data4.Recordset
If Not .BOF Then
    Text16 = !kode_maskapai
    Text17 = !no_tiket
    Text18 = !no_penerbangan
    Text12 = !Class
    Text19 = Format(!tgl_berangkat, "dd mmm yyyy")
    Text13 = !Status
    Text10 = !keterangan
    Text20 = !sex_psg
    Text14 = !nama_psg
    Text15 = !hrg_pokok
    Text23 = !harga_tiket
    Label1(44).Caption = !Currency
End If
End With
End Sub

Sub kosong_lok()
Label1(13) = ""
Label1(14) = ""
Label1(15) = ""
Label1(16) = ""
End Sub

Private Sub Calendar1_Click()
Text19 = Format(Calendar1.Value, "dd mmm yyyy")
End Sub

Private Sub Combo1_Click()
    isi_pembeli3
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Command1(0).Caption = "CLIENT BARU" Then
'    KeyAscii = 0
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Sub isi_combo1()
Combo1.Clear
With Data2.Recordset
If .RecordCount <> 0 Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem (!COMPANY)
        .MoveNext
    Loop
End If
Data2.Refresh
End With
End Sub

Sub cmd_dtlawal()
Command1(2).Caption = "TAMBAH TIKET"
Command1(3).Caption = "EDIT TIKET"
Combo2.Visible = False
Combo3.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Text16.Visible = True
Text17.Visible = True
Text18.Visible = True
Text20.Visible = True
Combo8.Enabled = True
DBGrid2.Enabled = True
Combo8.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
DBGrid3.Enabled = True
Data1.Enabled = True
Data4.Enabled = True
End Sub

Sub cmd_dtlsimpan1()
Data1.Enabled = False
Data4.Enabled = False
Command3.Enabled = False
Command2.Enabled = False
Combo8.Enabled = False
Command1(2).Caption = "SIMPAN"
Command1(3).Caption = "BATAL"
Combo2.Visible = True
Combo3.Visible = True
Combo4.Visible = True
Combo5.Visible = True
Text16.Visible = False
Text17.Visible = False
Text18.Visible = False
Text20.Visible = True
Combo8.Enabled = False
DBGrid2.Enabled = False
End Sub

Sub cmd_dtlsimpan2()
Command3.Enabled = False
Command2.Enabled = False
Command1(2).Caption = "SIMPAN"
Command1(3).Caption = "BATAL"
Combo8.Enabled = False
DBGrid2.Enabled = False
DBGrid3.Enabled = False
Data1.Enabled = False
Data4.Enabled = False
End Sub

Private Sub Combo2_Click()
Text16 = Combo2
nama_maskapai
isi_cmb3
isi_cmb4
End Sub

Sub isi_cmb5()
Combo5.Clear
Combo5.AddItem ("MR")
Combo5.AddItem ("MRS")
Combo5.AddItem ("MS")
Combo5.AddItem ("MSTR")
Combo5.AddItem ("MSS")
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo3_Click()
Text17 = Combo3
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo4_Click()
Text18 = Combo4
dtl_lok
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo6_Change()
Text21 = Combo6
End Sub

Private Sub Combo6_Click()
Text21 = Combo6
Label1(44).Caption = Combo6
End Sub

Private Sub Combo8_Change()
    Data1.RecordSource = "select jenkel_penumpang,passanger_name,tanggal_berangkat,kode_maskapai,no_penerbangan,no_tiket,class,db_log.from,to,status,harga,currency,hrg_jual,tgl_lg,keterangan,company_beli,status_lg from db_log where no_lg='" & Combo8 & "'"
    Data1.Refresh
    Data4.RecordSource = "select * from temp_inv where no_lg='" & Combo8 & "'"
    Data4.Refresh
    isi_dtlLG
    isi_list1
    Data2.RecordSource = "select * from db_client where company ='" & Combo1 & "'"
    Data2.Refresh
    isi_pembeli2
End Sub

Private Sub Combo8_Click()
    Data1.RecordSource = "select jenkel_penumpang,passanger_name,tanggal_berangkat,kode_maskapai,no_penerbangan,no_tiket,class,db_log.from,to,status,harga,currency,hrg_jual,tgl_lg,keterangan,company_beli,status_lg from db_log where no_lg='" & Combo8 & "'"
    Data1.Refresh
    Data4.RecordSource = "select * from temp_inv where no_lg='" & Combo8 & "'"
    Data4.Refresh
    isi_dtlLG
    isi_list1
    Data2.RecordSource = "select * from db_client where company ='" & Combo1 & "'"
    Data2.Refresh
    isi_pembeli2
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Me.Enabled = False
    client_frm.Show
    client_frm.Label3.Caption = "inv"
Case 1
Case 2
    If Command1(2).Caption = "TAMBAH TIKET" Then
        kosong_dtl
        buka_dtl
        cmd_dtlsimpan1
        isi_cmb2
        isi_cmb5
        Label2 = "t"
        kosong_lok
    Else
        simpan_dtltambah
    End If
Case 3
    If Command1(3).Caption = "EDIT TIKET" Then
        If Text11 <> "" Then
            If Label2.Caption = "el" Then
                buka_dtllg
                cmd_dtlsimpan2
                Label2.Caption = "el"
            Else
                buka_dtl
                cmd_dtlsimpan2
                Label2.Caption = "et"
            End If
        Else
            x = MsgBox("Tidak ada data yg akan diedit, mohon periksa kembali", vbOKOnly, "Validasi Data")
        End If
    Else
        tutup_dtl
        isi_dtlLG
        isi_list1
        cmd_dtlawal
        Label2 = "el"
    End If
End Select
End Sub

Sub simpan_dtltambah()
With Data4.Recordset
If Label2.Caption = "t" Then
    If Combo2 = "" Or Text15 = "" Or Combo3 = "" Or Combo4 = "" Or Text12 = "" Or Text19 = "" Or Format(Text19, "yyyymmdd") < Format(Date, "yyyymmdd") Or Text13 = "" Or Combo5 = "" Or Text14 = "" Or Text23 = "" Then
        x = MsgBox("Data belum lengkap...", vbOKOnly, "Validasi Data")
        If Combo2 = "" Then
            Combo2.SetFocus
        ElseIf Combo3 = "" Then
            Combo3.SetFocus
        ElseIf Combo4 = "" Then
            Combo4.SetFocus
        ElseIf Text12 = "" Then
            Text12.SetFocus
        ElseIf Text19 = "" Or Format(Text19, "yyyymmdd") < Format(Date, "yyyymmdd") Then
            x = MsgBox("Periksa kembali tanggal keberangkatan...", vbOKOnly, "Validasi Tanggal")
            Text19.SetFocus
        ElseIf Text13 = "" Then
            Text13.SetFocus
        ElseIf Text15 = "" Then
            Text15.SetFocus
        ElseIf Combo5 = "" Then
            Combo5.SetFocus
        ElseIf Text14 = "" Then
            Text14.SetFocus
        ElseIf Text23 = "" Then
            Text23.SetFocus
        End If
    Else
        .AddNew
        If Label4.Caption = "invlg" Then
            !no_lg = Combo8
        Else
            !no_lg = " "
        End If
        !kode_maskapai = Combo2
        !no_tiket = Combo3
        !no_penerbangan = Combo4
        !Class = Text12
        !tgl_berangkat = Text19
        !Status = Text13
        !keterangan = Text10
        !nama_psg = Text14
        !sex_psg = Combo5
        !harga_tiket = Val(Text23)
        !Currency = Label1(44).Caption
        !hrg_pokok = Val(Text15)
        !From = Label1(13).Caption
        !To = Label1(14).Caption
        .Update
        With Data5.Recordset
            If Not .BOF Then
                .MoveFirst
                Do While Not .EOF
                    If !no_tiket = Combo3 Then
                        .Delete
                        .MoveLast
                    End If
                    If Not .BOF Then
                    .MoveNext
                    End If
                Loop
                Data5.Refresh
            End If
        End With
        Data4.RecordSource = "select * from temp_inv where no_lg = '" & Combo8 & "'"
        Data4.Refresh
        Data1.Refresh
        tutup_dtl
        isi_dtlLG
        isi_list1
        cmd_dtlawal
    End If
ElseIf Label2.Caption = "et" Then
    If Text12 = "" Or Text19 = "" Or Format(Text19, "mm/dd/yyyy") < Format(Date, "mm/dd/yyyy") Or Text13 = "" Or Text23 = "" Then
        x = MsgBox("Data belum lengkap...", vbOKOnly, "Validasi Data")
        If Text12 = "" Then
            Text12.SetFocus
        ElseIf Text19 = "" Or Format(Text19, "dd mmm yyyy") < Format(Date, "dd mmm yyyy") Then
            x = MsgBox("Periksa kembali tanggal keberangkatan...", vbOKOnly, "Validasi Tanggal")
            Text19.SetFocus
        ElseIf Text13 = "" Then
            Text13.SetFocus
        ElseIf Text23 = "" Then
            Text23.SetFocus
        End If
    Else
        .Edit
        !Class = Text12
        !tgl_berangkat = Text19
        !Status = Text13
        !keterangan = Text10
        !harga_tiket = Val(Text23)
        !hrg_pokok = Val(Text15)
        .Update
        Data1.Refresh
        Data4.Refresh
        tutup_dtl
        isi_dtlLG
        isi_list1
        cmd_dtlawal
    End If
ElseIf Label2.Caption = "el" Then
    If Text17 = "" Or Text13 = "" Or Text23 = "" Then
        x = MsgBox("Data belum lengkap...", vbOKOnly, "Validasi Data")
        If Text17 = "" Then
            Text17.SetFocus
        ElseIf Text13 = "" Then
            Text13.SetFocus
        ElseIf Text23 = "" Then
            Text23.SetFocus
        End If
    Else
        With Data1.Recordset
            .Edit
            !no_tiket = Text17
            !Status = Text13
            !keterangan = Text10
            !hrg_jual = Val(Text23)
            !harga = Format(Text15, "###")
            .Update
        End With
    Data1.Refresh
    Data4.Refresh
    tutup_dtl
    isi_dtlLG
    isi_list1
    cmd_dtlawal
    End If
End If
End With
End Sub

Private Sub simpan_datacnt()
With Data2.Recordset
    !COMPANY = Combo1
    !contact_person1 = Text2
    !contact_person2 = Text3
    !telp1 = Text4
    !telp2 = Text5
    !email1 = Text6
    !email2 = Text7
    !address1 = Text8
    !address2 = Text9
    !id_client = Text1
    .Update
End With
End Sub

Private Sub simpan_client()
Dim y As Boolean
y = False
If Combo1 = "" Or Text2 = "" Or Text4 = "" Or Text6 = "" Or Text8 = "" Then
    x = MsgBox("Data belum lengkap...", vbOKOnly, "Validasi Data")
    If Combo1 = "" Then
        Combo1.SetFocus
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
    With Data2.Recordset
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
            simpan_datacnt
            tutup_pembeli
            Data2.Refresh
            Command1(0).Caption = "CLIENT BARU"
            Command2.Caption = "CETAK INVOICE"
        Else
            Combo1.SetFocus
        End If
    Else
        .AddNew
        simpan_datacnt
        tutup_pembeli
        Data2.Refresh
        Command1(0).Caption = "CLIENT BARU"
        Command2.Caption = "CETAK INVOICE"
    End If
    End With
End If
End Sub

Private Sub Command2_Click()
Data1.Refresh
If Not Data1.Recordset.BOF Or Text11 <> "" Then
    Me.Hide
    main_form.Enabled = True
    PrintInv_frm.Show
Else
    x = MsgBox("Mohon data diperiksa kembali...", vbOKOnly, "Validasi Data")
End If
End Sub

Private Sub Command3_Click()
With Data4.Recordset
    If Not .BOF Then
    x = MsgBox("Transaksi belum dicetak... apakah anda yakin untuk keluar?", vbOKCancel, "Exit")
        If x = vbOK Then
            .MoveFirst
            Do While Not .RecordCount = 0
                .Delete
                .MoveNext
            Loop
            Data4.Refresh
            Me.Hide
            main_form.Enabled = True
            main_form.Show
            Call isi_cekdat
        End If
    Else
        Me.Hide
        main_form.Enabled = True
        main_form.Show
        Call isi_cekdat
    End If
End With
With Data5.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
        Data5.Refresh
    End If
End With
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    isi_pembeli
End Sub

Private Sub DBGrid2_Click()
Label2 = "el"
isi_dtlLG
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command1(1).Caption <> "SIMPAN" And Not Data1.Recordset.BOF Then
    isi_dtlLG
End If
End Sub

Private Sub DBGrid3_Click()
isi_dtltambah
Label2 = "et"
End Sub

Private Sub DBGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
isi_dtltambah
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
DTPicker1 = Date
With stoktiket_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Data5.Recordset.AddNew
        Data5.Recordset!kode_maskapai = !kode_maskapai
        Data5.Recordset!no_tiket = !no_tiket
        Data5.Recordset.Update
        .MoveNext
    Loop
    stoktiket_frm.Data1.Refresh
    Data5.Refresh
End If
End With
kosong_pembeli
Label1(44).Caption = ""
kosong_dtl
If Not Data4.Recordset.BOF Then
    isi_dtltambah
End If
Call inv_auto
isi_combo1
If Not Data2.Recordset.RecordCount = 0 Then
    isi_pembeli2
    tutup_pembeli
End If
Command1(1).Enabled = False
isi_cmb6
isi_list1
Calendar1 = Date
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Label1(44).Caption = Text21.Text
End Sub

Sub isi_cmb6()
Combo6.Clear
Combo6.AddItem ("RP")
Combo6.AddItem ("USD")
Combo6.AddItem ("SGD")
Combo6.AddItem ("HKD")
Combo6.AddItem ("CNY")
Combo6.AddItem ("JPY")
Combo6.AddItem ("MYR")
Combo6.AddItem ("SAR")
Combo6.AddItem ("THB")
Combo6.AddItem ("NTD")
Combo6.AddItem ("EURO")
Combo6.AddItem ("GBP")
Combo6.AddItem ("CHF")
Combo6.AddItem ("AUD")
Combo6.AddItem ("NZD")
Combo6.AddItem ("CND")
Combo6.AddItem ("PHP")
Combo6.AddItem ("WON")
Combo6.AddItem ("IND")
Combo6.AddItem ("VND")
Combo6.AddItem ("AED")
Combo6.AddItem ("BND")
Combo6.AddItem ("OMR")
Combo6.AddItem ("EGP")
Combo6.AddItem ("SRI")
Combo6.AddItem ("QTR")
Combo6.AddItem ("ZAR")
End Sub
Private Sub Form_Load()
kosong_pembeli
kosong_dtl
Call db_invlg
Calendar1 = Date
Label1(1) = Format(Date, "dd mmm yyyy")
tutup_dtl
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
main_form.Enabled = True
main_form.Show
End Sub

Private Sub hit_hrg()
Dim hrg, jual As Double
With Data1.Recordset
hrg = 0
jual = 0
If Not .BOF And Label4.Caption = "invlg" Then
    .MoveFirst
    Do While Not .EOF
        hrg = hrg + !harga
        jual = !hrg_jual + jual
        .MoveNext
    Loop
End If
End With
With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        jual = !harga_tiket + jual
        .MoveNext
    Loop
End If
End With
Label1(38).Caption = Format(hrg, "###,###,###.##")
Label1(41).Caption = Format(jual, "###,###,###.##")
Data1.Refresh
Data4.Refresh
End Sub

Private Sub isi_list1()
Dim a As Byte
Dim b, c, d, e As String
With Data1.Recordset
List1.Clear
List1.AddItem ("")
List1.AddItem ("DATA INVOICE")
List1.AddItem ("")
a = 1
If Label4.Caption = "invlg" Then
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            With dbtiket_frm.Data3.Recordset
                .MoveFirst
                Do While Not .EOF
                    If Data1.Recordset!kode_maskapai = !kode_maskapai And Data1.Recordset!no_penerbangan = !no_penerbangan Then
                        b = !From
                        c = !To
                        d = Format(!dep, "hh:mm")
                        e = Format(!arr, "hh:mm")
                        With Data1.Recordset
                            List1.AddItem (a & ". " & !jenkel_penumpang & " " & !passanger_name & ", Flights : " & !kode_maskapai & !no_penerbangan & ", Class : " & !Class)
                            List1.AddItem ("     " & " Date : " & !tanggal_berangkat & ", From : " & b & " To : " & c & ", ETD : " & d & " ETA : " & e)
                            List1.AddItem ("")
                        End With
                        .MoveLast
                    End If
                    .MoveNext
                Loop
            End With
            dbtiket_frm.Data3.Refresh
            .MoveNext
            a = a + 1
        Loop
    Data1.Refresh
    End If
End If
End With
With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If Data4.Recordset!kode_maskapai = !kode_maskapai And Data4.Recordset!no_penerbangan = !no_penerbangan Then
                    b = !From
                    c = !To
                    d = Format(!dep, "hh:mm")
                    e = Format(!arr, "hh:mm")
                    With Data4.Recordset
                        List1.AddItem (a & ". " & !sex_psg & " " & !nama_psg & ", Flights : " & !kode_maskapai & !no_penerbangan & ", Class : " & !Class)
                        List1.AddItem ("     " & " Date : " & !tgl_berangkat & ", From : " & b & " To : " & c & ", ETD : " & d & " ETA : " & e)
                        List1.AddItem ("")
                    End With
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        dbtiket_frm.Data3.Refresh
        .MoveNext
        a = a + 1
    Loop
Data4.Refresh
End If
End With
hit_hrg
End Sub

Private Sub kosong_pembeli()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Combo1 = ""
Text22 = ""
Text24 = ""
End Sub

Private Sub isi_pembeli3()
Dim a As Boolean
a = False
With Data2.Recordset
If Not .RecordCount = 0 Then
    .MoveFirst
    Do While Not a = True
        If !COMPANY = Combo1 Then
            a = True
            .MovePrevious
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Private Sub isi_pembeli()
With Data2.Recordset
If Not .BOF Then
Text1 = !id_client
Text2 = !contact_person1
Text3 = !contact_person2
Text4 = !telp1
Text5 = !telp2
Text6 = !email1
Text7 = !email2
Text8 = !address1
Text9 = !address2
Text22 = !fax1
Text24 = !fax2
Combo1 = !COMPANY
End If
End With
End Sub

Private Sub isi_pembeli2()
Dim a As Boolean
a = False
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF And a = False
        If Combo1 = !COMPANY Then
            a = True
            .MovePrevious
        End If
        .MoveNext
    Loop
    Data2.Refresh
End If
End With
End Sub

Private Sub tutup_pembeli()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text22.Enabled = False
Text24.Enabled = False
End Sub

Private Sub buka_pembeli()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Combo1.Enabled = True
Text22.Enabled = True
Text24.Enabled = True
End Sub

Private Sub tutup_dtl()
Text16.Enabled = False
Text11.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text12.Enabled = False
Text19.Enabled = False
Text13.Enabled = False
Text10.Enabled = False
Text20.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text21.Enabled = False
Text23.Enabled = False
Combo6.Enabled = False
End Sub

Private Sub kosong_dtl()
Text16 = ""
Text11 = ""
Text17 = ""
Text18 = ""
Text12 = ""
Text19 = ""
Text13 = ""
Text10 = ""
Text20 = ""
Text14 = ""
Text15 = ""
'Text21 = ""
Text23 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""
Combo6 = ""
End Sub

Sub isi_cmb2()
Combo2.Clear
With dbtiket_frm.Data1.Recordset
If Not .RecordCount = 0 Then
    .MoveFirst
    Do While Not .EOF
        Combo2.AddItem (!kode_maskapai)
        .MoveNext
    Loop
End If
End With
dbtiket_frm.Data1.Refresh
End Sub

Sub isi_cmb3()
Combo3.Clear
With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo2 = !kode_maskapai Then
            Combo3.AddItem (!no_tiket)
        End If
        .MoveNext
    Loop
End If
End With
Data5.Refresh
End Sub

Sub isi_cmb4()
Combo4.Clear
With dbtiket_frm.Data3.Recordset
If Not .RecordCount = 0 Then
    .MoveFirst
    Do While Not .EOF
        If Combo2 = !kode_maskapai Then
            Combo4.AddItem (!no_penerbangan)
        End If
        .MoveNext
    Loop
End If
End With
dbtiket_frm.Data3.Refresh
End Sub

Private Sub buka_dtllg()
Text17.Enabled = True
Text13.Enabled = True
Text10.Enabled = True
Text15.Enabled = True
Text23.Enabled = True
End Sub

Sub buka_dtl()
Text12.Enabled = True
Text13.Enabled = True
Text10.Enabled = True
Text14.Enabled = True
Text19.Enabled = True
Text23.Enabled = True
If Data4.Recordset.RecordCount = 0 Then
Combo6.Enabled = True
Else
Combo6.Enabled = False
End If
Text15.Enabled = True
End Sub


Private Sub isi_dtlLG()
kosong_dtl
If Label4.Caption = "invlg" Then
    With Data1.Recordset
    If Not .BOF Then
        Label1(36).Caption = Format(!tgl_lg, "dd mmm yyyy")
        Text16 = !kode_maskapai
        nama_maskapai
        Text17 = !no_tiket
        Text18 = !no_penerbangan
        Text12 = !Class
        Text19 = Format(!tanggal_berangkat, "dd mmm yyyy")
        dtl_lok
        Text13 = !Status
        Text10 = !keterangan
        Text20 = !jenkel_penumpang
        Text14 = !passanger_name
        Text15 = !harga
        Text21 = !Currency
        Text23 = !hrg_jual
        Combo1 = !company_beli
    End If
    End With
End If
End Sub

Sub nama_maskapai()
With dbtiket_frm.Data1.Recordset
    .MoveFirst
    Do While Not .EOF
        If !kode_maskapai = Text16 Then
            Text11 = !nama_maskapai
            .MoveLast
        End If
        .MoveNext
    Loop
End With
dbtiket_frm.Data1.Refresh
End Sub

Sub dtl_lok()
With dbtiket_frm.Data3.Recordset
    .MoveFirst
    Do While Not .EOF
        If !kode_maskapai = Text16 And !no_penerbangan = Text18 Then
            Label1(13).Caption = !From
            Label1(14).Caption = !To
            Label1(15).Caption = Format(!dep, "hh:mm")
            Label1(16).Caption = Format(!arr, "hh:mm")
            .MoveLast
        End If
        .MoveNext
    Loop
End With
dbtiket_frm.Data3.Refresh
End Sub

Private Sub Label1_Change(Index As Integer)
Select Case Index
Case 44
Label1(42).Caption = Label1(44).Caption
'Case 39
End Select
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text16_Change()
nama_maskapai
End Sub

Private Sub Text19_GotFocus()
Calendar1.Visible = True
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text19_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub Text21_Change()
Label1(44).Caption = Text21
Label1(39).Caption = Text21
Label1(42).Caption = Text21
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub
