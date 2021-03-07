VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form CetakLg_frm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK LETTER OF GUARANTEE"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   FillColor       =   &H00FFFFFF&
   Icon            =   "CetakLg_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   15270
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "LETTER OF GUARANTEE"
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
      Picture         =   "CetakLg_frm.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   120
      Width           =   7695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2415
      Left            =   120
      TabIndex        =   62
      Top             =   7200
      Width           =   7695
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF80FF&
         Caption         =   "Client Baru"
         DownPicture     =   "CetakLg_frm.frx":3994
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6120
         MouseIcon       =   "CetakLg_frm.frx":465E
         MousePointer    =   99  'Custom
         Picture         =   "CetakLg_frm.frx":4968
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   70
         Text            =   "Text13"
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   69
         Text            =   "Text12"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo9 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   68
         Text            =   "Combo9"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   67
         Text            =   "Text8"
         Top             =   840
         Width           =   3855
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   38
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   990
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   37
         Left            =   120
         TabIndex        =   65
         Top             =   1800
         Width           =   1065
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   36
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   735
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   35
         Left            =   120
         TabIndex        =   63
         Top             =   840
         Width           =   1770
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "CETAK LETTER OF GUARANTEE"
      DownPicture     =   "CetakLg_frm.frx":5632
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      MouseIcon       =   "CetakLg_frm.frx":62FC
      MousePointer    =   99  'Custom
      Picture         =   "CetakLg_frm.frx":6606
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8280
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C0C0&
      Caption         =   "E X I T"
      DownPicture     =   "CetakLg_frm.frx":82D0
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13440
      MouseIcon       =   "CetakLg_frm.frx":8F9A
      MousePointer    =   99  'Custom
      Picture         =   "CetakLg_frm.frx":92A4
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stretches"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   8175
      Left            =   8040
      TabIndex        =   32
      Top             =   0
      Width           =   6975
      Begin MSACAL.Calendar Calendar1 
         Height          =   2415
         Left            =   2160
         TabIndex        =   72
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
         _Version        =   524288
         _ExtentX        =   6376
         _ExtentY        =   4260
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2007
         Month           =   12
         Day             =   26
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
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   5280
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Text            =   "Text10"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   8760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Text            =   "Combo5"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Text            =   "TEXT7"
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   8
         Text            =   "Tgl"
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   9
         Text            =   "Bln"
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Text            =   "Thn"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "CetakLg_frm.frx":9F6E
         Top             =   3480
         Width           =   6735
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   4920
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5160
         TabIndex        =   18
         Text            =   "Combo4"
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "TAMBAH"
         DownPicture     =   "CetakLg_frm.frx":9F74
         Height          =   855
         Index           =   0
         Left            =   120
         MouseIcon       =   "CetakLg_frm.frx":AC3E
         MousePointer    =   99  'Custom
         Picture         =   "CetakLg_frm.frx":AF48
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5640
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "EDIT"
         DownPicture     =   "CetakLg_frm.frx":C8CA
         Height          =   855
         Index           =   1
         Left            =   2520
         MouseIcon       =   "CetakLg_frm.frx":D594
         MousePointer    =   99  'Custom
         Picture         =   "CetakLg_frm.frx":D89E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5640
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "HAPUS"
         DownPicture     =   "CetakLg_frm.frx":F220
         Height          =   855
         Index           =   2
         Left            =   4800
         MouseIcon       =   "CetakLg_frm.frx":FEEA
         MousePointer    =   99  'Custom
         Picture         =   "CetakLg_frm.frx":101F4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5640
         Width           =   2055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "CetakLg_frm.frx":11B76
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "CetakLg_frm.frx":11B8A
         TabIndex        =   33
         Top             =   6600
         Width           =   6735
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
         Index           =   32
         Left            =   4230
         TabIndex        =   59
         Top             =   5280
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
         Index           =   31
         Left            =   120
         TabIndex        =   58
         Top             =   5280
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   5400
         Visible         =   0   'False
         Width           =   975
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
         Index           =   29
         Left            =   4440
         TabIndex        =   54
         Top             =   2400
         Width           =   600
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
         Index           =   25
         Left            =   120
         TabIndex        =   50
         Top             =   1440
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   24
         Left            =   4440
         TabIndex        =   49
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Keberangkatan"
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
         TabIndex        =   48
         Top             =   840
         Width           =   1995
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
         Index           =   22
         Left            =   120
         TabIndex        =   47
         Top             =   2400
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   21
         Left            =   2160
         TabIndex        =   46
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passanger Name"
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
         TabIndex        =   45
         Top             =   480
         Width           =   1815
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   19
         Left            =   120
         TabIndex        =   44
         Top             =   2760
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
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   3
         Left            =   2160
         TabIndex        =   43
         Top             =   2760
         Width           =   465
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
         Index           =   10
         Left            =   4440
         TabIndex        =   42
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Tiket"
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
         TabIndex        =   41
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan "
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
         Index           =   12
         Left            =   120
         TabIndex        =   40
         Top             =   3240
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Beli"
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
         Left            =   120
         TabIndex        =   39
         Top             =   4920
         Width           =   1125
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
         Left            =   840
         TabIndex        =   38
         Top             =   2400
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
         Index           =   15
         Left            =   2760
         TabIndex        =   37
         Top             =   2400
         Width           =   915
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
         Left            =   840
         TabIndex        =   36
         Top             =   2760
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
         Index           =   17
         Left            =   2760
         TabIndex        =   35
         Top             =   2760
         Width           =   1065
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
         Index           =   18
         Left            =   4200
         TabIndex        =   34
         Top             =   4920
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5895
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   7695
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   2010
         Left            =   360
         TabIndex        =   57
         Top             =   2400
         Width           =   6975
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Bindings        =   "CetakLg_frm.frx":1255D
         Left            =   5640
         Top             =   5760
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "D:\waynh proj\travel project\report\cetak_lg.rpt"
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo8 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Combo8"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Combo7"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "Text11"
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "CetakLg_frm.frx":12571
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "Combo6"
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         Height          =   3735
         Left            =   120
         Top             =   1920
         Width           =   7455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   480
         TabIndex        =   55
         Top             =   4680
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   4680
         TabIndex        =   53
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp No."
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
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
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
         TabIndex        =   51
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date : "
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
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. LG :"
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
         Left            =   4560
         TabIndex        =   30
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To."
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
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stretches"
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
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fare"
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
         Left            =   480
         TabIndex        =   27
         Top             =   5160
         Width           =   1095
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   6
         Left            =   4440
         TabIndex        =   26
         Top             =   5160
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal : "
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
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   1065
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   8
         Left            =   5640
         TabIndex        =   24
         Top             =   360
         Width           =   1350
      End
   End
End
Attribute VB_Name = "CetakLg_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calendar1_Click()
Combo3(0) = Calendar1.Day
Combo3(1) = Calendar1.Month
Combo3(2) = Calendar1.Year
End Sub

Private Sub Combo1_Click()
isi_cmb2
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo2_Click()
isi_autodet
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo3_GotFocus(Index As Integer)
Calendar1.Visible = True
End Sub

Private Sub Combo3_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_LostFocus(Index As Integer)
Calendar1.Visible = False
End Sub

Private Sub Combo4_Click()
Label1(32).Caption = Combo4
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo6_Click()
    isi_cmb7
    isi_cmb8
    isi_txt9
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Combo9_Click()
Dim a As Boolean
a = False
With Data3.Recordset
.MoveFirst
Do While Not .EOF
    If !COMPANY = Combo9 Then
        isi_client
        .MoveLast
    End If
    .MoveNext
Loop
End With
End Sub

Private Sub Combo9_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Command1_Click()
'Call uncons
If Combo6 = "" Or Combo7 = "" Or Combo8 = "" Or Text9 = "" Or Data2.Recordset.RecordCount = 0 Or Text11 = "" Then
    x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi Data")
    If Combo6 = "" Then
        Combo6.SetFocus
    ElseIf Combo7 = "" Then
        Combo7.SetFocus
    ElseIf Combo8 = "" Then
        Combo8.SetFocus
    ElseIf Text9 = "" Then
        Text9.SetFocus
    ElseIf Text11 = "" Then
        Text11.SetFocus
    ElseIf Data2.Recordset.RecordCount = 0 Then
        x = MsgBox("Anda belum memasukan detil LG...", vbOKOnly, "Validasi Data")
    End If
Else
    cetak_lg
    Me.Hide
    main_form.Enabled = True
    Call isi_cekdat
End If
End Sub



Sub cetak_lg()
simpan_lg
update_byr
PrintLG_frm.Show
hapus_dtl
kosong_detil
kosong
End Sub

Private Sub hapus_dtl()
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Delete
        .MoveNext
    Loop
End If
End With
Data2.Refresh
End Sub

Private Sub simpan_lg()
With Data1.Recordset
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
    .AddNew
    !no_lg = Label1(8)
    !tgl_lg = Date
    !waktu_lg = Time
    !COMPANY = Combo6
    !contact_person = Combo7
    !telp = Combo8
    !address = Text9
    !jenkel_penumpang = Data2.Recordset!psg_sex
    !passanger_name = Data2.Recordset!psg_name
    !tanggal_berangkat = Data2.Recordset!tgl_berangkat
    !kode_maskapai = Data2.Recordset!kode_maskapai
    !no_penerbangan = Data2.Recordset!flight_no
    !no_tiket = Data2.Recordset!no_tiket
    !Class = Data2.Recordset!Class
    !Status = Data2.Recordset!Status
    !keterangan = Data2.Recordset!keterangan
    !harga = Data2.Recordset!harga
    !hrg_jual = Data2.Recordset!hrg_jual
    !Currency = Data2.Recordset!Currency
    !kode = Text11
    !status_lg = 0
    !company_beli = Combo9
    !From = Data2.Recordset!From
    !To = Data2.Recordset!To
    Data2.Recordset.MoveNext
    .Update
Loop
Data1.Refresh
Data2.Refresh
End With
End Sub

Private Sub hit_hrg()
Dim hrg As Double
With Data2.Recordset
hrg = 0
If Not .BOF Then
    .MoveFirst
    Label1(6) = !Currency
    Do While Not .EOF
        hrg = hrg + !harga
        .MoveNext
    Loop
End If
End With
Text2 = Format(hrg, "###,###,###.##")
Data2.Refresh
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
    If Command2(0).Caption = "TAMBAH" Then
        kosong_detil
        isi_cmb_dtl
        buka_dtl
        cmd_dtl_simpan
        Text7.SetFocus
        Label2 = "t"
        If Not Data2.Recordset.BOF Then
            Data2.Recordset.MoveFirst
            Combo4 = Data2.Recordset!Currency
            Combo4.Enabled = False
            Label1(6) = Combo4
            Data2.Recordset.MoveLast
            isi_dtl
            Text7.SetFocus
        End If
    Else
        cek_data_dtl
    End If
Case 1
    If Not Data2.Recordset.BOF Then
        If Command2(1).Caption = "EDIT" Then
            cmd_dtl_simpan
            buka_dtl
            Combo5.SetFocus
            Label2 = "e"
            Combo4.Enabled = False
        Else
            Data2.Refresh
            tutup_dtl
            cmd_dtl_awal
            kosong_detil
            isi_dtl
            isi_list1
            'isi_listview1
            'hit_hrg
        End If
    Else
        x = MsgBox("Data msh kosong", vbOKOnly, "Validasi Data")
        cmd_dtl_awal
        kosong_detil
        tutup_dtl
    End If
Case 2
    If Not Data2.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin?", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data2.Recordset.Delete
        End If
        Data2.Refresh
        If Not Data2.Recordset.BOF Then
            isi_dtl
            isi_list1
            'isi_listview1
            'hit_hrg
        End If
    Else
        x = MsgBox("Data msh kosong", vbOKOnly, "Validasi Data")
    End If
End Select
End Sub

Private Sub cek_data_dtl()
Dim tgl As Date
With Data2.Recordset
If Combo3(0) <> "Tgl" Then
    tgl = Format(Combo3(0) & "/" & Combo3(1) & "/" & Combo3(2), "dd/mm/yyyy")
End If
If tgl < Date Or Text7 = "" Or Combo1 = "" Or Combo2 = "" Or Text10 = "" Or Text3 = "" Or Text5 = "" Or Combo3(0) = "Tgl" Or Combo3(1) = "Bln" Or Combo3(2) = "Thn" Then
    x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi Data")
    If Text7 = "" Then
        Text7.SetFocus
    ElseIf tgl < Date Or Combo3(0) = "Tgl" Or Combo3(1) = "Bln" Or Combo3(2) = "Thn" Then
        x = MsgBox("Mohon periksa tangal keberangkatan...", vbOKOnly, "Validasi Tanggal")
        Combo3(0).SetFocus
    ElseIf Combo1 = "" Then
        Combo1.SetFocus
    ElseIf Combo2 = "" Then
        Combo2.SetFocus
    ElseIf Text10 = "" Then
        Text10.SetFocus
    ElseIf Text3 = "" Then
        Text3.SetFocus
    ElseIf Text5 = "" Then
        Text5.SetFocus
    End If
Else
    If Label2 = "t" Then
        .AddNew
        If Text1 = "" Then
            Text1 = "N/A"
        End If
        simpan_dtl
        .Update
    Else
        .Edit
        simpan_dtl
        .Update
    End If
    Data2.Refresh
    cmd_dtl_awal
    isi_dtl
    tutup_dtl
    isi_list1
    'isi_listview1
    hit_hrg
End If
End With
End Sub

Private Sub isi_dtl()
With Data2.Recordset
    Combo5 = !psg_sex
    Text7 = !psg_name
    Combo3(0) = Format(!tgl_berangkat, "d")
    Combo3(1) = Format(!tgl_berangkat, "m")
    Combo3(2) = Format(!tgl_berangkat, "yyyy")
    Combo1 = !kode_maskapai
    Combo2 = !flight_no
    Text1 = !no_tiket
    Text10 = !Class
    Text3 = !Status
    Text4 = !keterangan
    Text5 = !harga
    Combo4 = !Currency
    Text6 = !hrg_jual
    isi_autodet
End With
End Sub

Private Sub simpan_dtl()
Dim tgl As Date
With Data2.Recordset
    !psg_sex = Combo5
    !psg_name = Text7
    tgl = Combo3(1) & "/" & Combo3(0) & "/" & Combo3(2)
    !tgl_berangkat = tgl
    !kode_maskapai = Combo1
    !flight_no = Combo2
    !no_tiket = Text1
    !Class = Text10
    !Status = Text3
    !keterangan = Text4
    !harga = Val(Text5)
    !Currency = Combo4
    !hrg_jual = Val(Text6)
    !From = Label1(14).Caption
    !To = Label1(15).Caption
End With
End Sub

Private Sub isi_cmb_dtl()
isi_cmb5
'isi_cmb3
isi_cmb1
isi_cmb4
End Sub

Private Sub Command3_Click()
Me.Hide
Call isi_cekdat
main_form.Show
main_form.Enabled = True
End Sub

Private Sub Command4_Click()
client_frm.Show
'client_frm.Command1(2).Visible = False
'client_frm.Command1(3).Visible = False
Me.Enabled = False
client_frm.Label3.Caption = "lg"
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Command2(0).Caption <> "SIMPAN" And Not Data2.Recordset.RecordCount = 0 Then
    isi_dtl
End If
End Sub

Private Sub isi_list1()
Dim a As Byte
Dim b, c, d, e As String
With Data2.Recordset
If Not .RecordCount = 0 Then
    List1.Clear
    List1.AddItem ("")
    List1.AddItem ("DETAIL LETTER OF GUARANTEE")
    List1.AddItem ("")
    .MoveFirst
    a = 1
    Do While Not .EOF
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If Data2.Recordset!kode_maskapai = !kode_maskapai And Data2.Recordset!flight_no = !no_penerbangan Then
                    b = !From
                    c = !To
                    d = Format(!dep, "hh:mm")
                    e = Format(!arr, "hh:mm")
                    With Data2.Recordset
                        List1.AddItem (a & ". " & !psg_sex & " " & !psg_name & ", Flights : " & !kode_maskapai & !flight_no & ", Class : " & !Class)
                        List1.AddItem ("     " & " Date : " & !tgl_berangkat & ", From : " & b & " To : " & c & ", ETD : " & d & " ETA : " & e)
                        List1.AddItem ("")
                    End With
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        .MoveNext
        a = a + 1
    Loop
    hit_hrg
End If
End With
Data2.Refresh
End Sub

Private Sub Form_Activate()
Call lg_auto
isi_cmb6
isi_cmb5
'isi_cmb3
isi_cmb1
isi_cmb4
isi_combo9
isi_client
tutup_dtl
Text2.Enabled = False
If Not Data2.Recordset.BOF Then
    isi_list1
End If
Calendar1 = Date
    Data1.Refresh
    Data2.Refresh
End Sub

Private Sub Form_Load()
    Label1(7) = Format(Date, "dd mmm yyyy")
    kosong
    kosong_detil
    kosong_client
    Call dblg
End Sub

Private Sub kosong()
Combo6.Clear
Combo7.Clear
Combo8.Clear
Text9 = ""
List1.Clear
Text11 = ""
Text2 = ""
End Sub

Private Sub kosong_detil()
Combo5.Clear
Text7 = ""
Combo1.Clear
Combo2.Clear
Text1 = ""
Text10 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Combo4.Clear
Combo3(0) = "Tgl"
Combo3(1) = "Bln"
Combo3(2) = "Thn"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Call isi_cekdat
main_form.Show
main_form.Enabled = True
End Sub

Private Sub isi_cmb6()
Dim a As String
Combo6.Clear
With dbpartner_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        a = !COMPANY
        Combo6.AddItem (a)
        .MoveNext
    Loop
Combo6.ListIndex = 0
End If
End With
End Sub

Private Sub isi_cmb7()
Dim a, b As String
Combo7.Clear
With dbpartner_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo6 = !COMPANY Then
            a = !contact_person1
            b = !contact_person2
            Combo7.AddItem (a)
            Combo7.AddItem (b)
            .MoveLast
        End If
        .MoveNext
    Loop
    Combo7.ListIndex = 0
End If
End With
End Sub

Private Sub isi_cmb8()
Dim a, b As String
Combo8.Clear
With dbpartner_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo6 = !COMPANY Then
            a = !telp1
            b = !telp2
            Combo8.AddItem (a)
            Combo8.AddItem (b)
            .MoveLast
        End If
        .MoveNext
    Loop
    Combo8.ListIndex = 0
End If
End With
End Sub

Private Sub isi_txt9()
Dim a As String
Text9 = ""
With dbpartner_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo6 = !COMPANY Then
            a = !address1
            Text9 = a
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Private Sub isi_cmb5()
Combo5.Clear
Combo5.AddItem ("MR")
Combo5.AddItem ("MRS")
Combo5.AddItem ("MS")
Combo5.AddItem ("MSTR")
Combo5.AddItem ("MSS")
Combo5.ListIndex = 0
End Sub

Private Sub isi_cmb3()
Dim tgl, bln As Byte
Dim thn As Single
tgl = 1
bln = 1
thn = 2007
Combo3(0).Clear
Combo3(1).Clear
Combo3(2).Clear
Do Until tgl > 31
    Combo3(0).AddItem (tgl)
    tgl = tgl + 1
Loop
Combo3(0) = Format(Date, "d")
Do Until bln > 12
    Combo3(1).AddItem (bln)
    bln = bln + 1
Loop
Combo3(1) = Format(Date, "m")
Do Until thn > 2050
    Combo3(2).AddItem (thn)
    thn = thn + 1
Loop
Combo3(2) = Format(Date, "yyyy")
End Sub

Private Sub isi_cmb1()
Dim a As String
Combo1.Clear
With dbtiket_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        a = !kode_maskapai
        Combo1.AddItem (a)
        .MoveNext
    Loop
End If
End With
End Sub

Private Sub isi_cmb2()
Dim a As String
Combo2.Clear
With dbtiket_frm.Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo1 = !kode_maskapai Then
            a = !no_penerbangan
            Combo2.AddItem (a)
        End If
        .MoveNext
    Loop
'    Combo2.ListIndex = 0
End If
End With
End Sub

Private Sub isi_autodet()
If Combo2 <> "" And Combo1 <> "" Then
    With dbtiket_frm.Data3.Recordset
        .MoveFirst
        Do While Not .EOF
        If Combo1 = !kode_maskapai And Combo2 = !no_penerbangan Then
            Label1(14).Caption = !From
            Label1(15).Caption = !To
            Label1(16).Caption = Format(!dep, "hh:mm")
            Label1(17).Caption = Format(!arr, "hh:mm")
            .MoveLast
        End If
        .MoveNext
        Loop
    End With
End If
End Sub

Private Sub isi_cmb4()
Combo4.Clear
Combo4.AddItem ("Rp")
Combo4.AddItem ("USD")
Combo4.AddItem ("SGD")
Combo4.AddItem ("HKD")
Combo4.AddItem ("CNY")
Combo4.AddItem ("JPY")
Combo4.AddItem ("MYR")
Combo4.AddItem ("SAR")
Combo4.AddItem ("THB")
Combo4.AddItem ("NTD")
Combo4.AddItem ("EURO")
Combo4.AddItem ("GBP")
Combo4.AddItem ("CHF")
Combo4.AddItem ("AUD")
Combo4.AddItem ("NZD")
Combo4.AddItem ("CND")
Combo4.AddItem ("PHP")
Combo4.AddItem ("WON")
Combo4.AddItem ("IND")
Combo4.AddItem ("VND")
Combo4.AddItem ("AED")
Combo4.AddItem ("BND")
Combo4.AddItem ("OMR")
Combo4.AddItem ("EGP")
Combo4.AddItem ("SRI")
Combo4.AddItem ("QTR")
Combo4.AddItem ("ZAR")
Combo4.ListIndex = 0
End Sub

Private Sub cmd_dtl_awal()
Command2(0).Caption = "TAMBAH"
Command2(1).Caption = "EDIT"
Command2(2).Enabled = True
End Sub

Private Sub cmd_dtl_simpan()
Command2(0).Caption = "SIMPAN"
Command2(1).Caption = "BATAL"
Command2(2).Enabled = False
End Sub

Private Sub tutup_dtl()
Combo5.Enabled = False
Text7.Enabled = False
Combo3(0).Enabled = False
Combo3(1).Enabled = False
Combo3(2).Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Text1.Enabled = False
Text10.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Combo4.Enabled = False
End Sub

Private Sub buka_dtl()
Combo5.Enabled = True
Text7.Enabled = True
Combo3(0).Enabled = True
Combo3(1).Enabled = True
Combo3(2).Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Text1.Enabled = True
Text10.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo4.Enabled = True
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub isi_listview1()
Dim clm As ColumnHeader
Dim itm As ListItem
Dim imgx As ListImage
Dim no As Byte
Set clm = ListView1.ColumnHeaders.Add(Text:="No.")
Set clm = ListView1.ColumnHeaders.Add(, , "Nama Penumpang", ListView1.Width / 3)
Set clm = ListView1.ColumnHeaders.Add(, , "Flight No")
Set clm = ListView1.ColumnHeaders.Add(, , "Date")
Set clm = ListView1.ColumnHeaders.Add(, , "From")
Set clm = ListView1.ColumnHeaders.Add(, , "to")
Set clm = ListView1.ColumnHeaders.Add(, , "ETD")
Set clm = ListView1.ColumnHeaders.Add(, , "ETA")
Set clm = ListView1.ColumnHeaders.Add(, , "Status")
Set clm = ListView1.ColumnHeaders.Add(, , "Price")
ListView1.View = lvwReport
With Data2.Recordset
ListView1.ListItems.Clear
no = 1
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set itm = ListView1.ListItems.Add(Text:=no)
        itm.SubItems(1) = !psg_sex & "." & !psg_name
        itm.SubItems(2) = !kode_maskapai & !flight_no
        itm.SubItems(3) = !tgl_berangkat
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If Data2.Recordset!kode_maskapai = !kode_maskapai And Data2.Recordset!flight_no = !no_penerbangan Then
                    itm.SubItems(4) = !From
                    itm.SubItems(5) = !To
                    itm.SubItems(6) = Format(!dep, "hh:mm")
                    itm.SubItems(7) = Format(!arr, "hh:mm")
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        itm.SubItems(8) = !Status
        itm.SubItems(9) = !Currency & !harga
        .MoveNext
        no = no + 1
    Loop
End If
End With
End Sub

Private Sub kosong_client()
Combo9 = ""
Text8 = ""
Text12 = ""
Text13 = ""
End Sub

Private Sub isi_combo9()
Combo9.Clear
With Data3.Recordset
If Not .BOF Then
.MoveFirst
Do While Not .EOF
    Combo9.AddItem (!COMPANY)
    .MoveNext
Loop
End If
'Combo9.ListIndex = 0
End With
Data3.Refresh
End Sub

Private Sub isi_client()
With Data3.Recordset
    Combo9 = !COMPANY
    Text8 = !contact_person1
    Text12 = !telp1
    Text13 = !address1
End With
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Sub update_byr()
Call db_byr
With byr_frm.Data3.Recordset
    .AddNew
    !nomor = Label1(8)
    !curr = Label1(6)
    !nilai_trans = Format(Text2, "###")
    !total_bayar = 0
    !sisa = Format(Text2, "###")
    !frek = 0
    .Update
End With
End Sub
