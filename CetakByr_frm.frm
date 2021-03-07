VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CetakByr_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK DETIL PEMBAYARAN"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ClipControls    =   0   'False
   FillColor       =   &H80000012&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   Icon            =   "CetakByr_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   5880
      TabIndex        =   20
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text7"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text9"
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text10"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text11"
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMINAL"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   27
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CURRENCY"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   26
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2895
      Index           =   2
      Left            =   3000
      TabIndex        =   24
      Top             =   2400
      Width           =   2175
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3836;4868"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2895
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Top             =   2400
      Width           =   1815
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3201;4868"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   2895
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   1695
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;4868"
      MatchEntry      =   0
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PEMBAYARAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NO. LG"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal LG"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Supplier"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pembayaran"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Transaksi"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa Bayar"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Frek. Pembayaran"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   1455
   End
End
Attribute VB_Name = "CetakByr_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub isi_data()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim x As Byte
x = byr_frm.SSTab1.Tab
With byr_frm
Text1 = .Text1(x)
Text2 = .Text2(x)
Text3 = .Text3(x)
Text4 = Format(.Text4(x), "###,###.##")
Text5 = Format(.Text5(x), "###,###.##")
Text6 = Format(.Text6(x), "###,###.##")
Text7 = .Text7(x)
Text8 = .Text8(x)
Text9 = .Text9(x)
Text10 = .Text10(x)
Text11 = .Text11(x)
End With
Set head = ListView1.ColumnHeaders.Add(, , "TANGGAL", ListView1.Width / 4)
Set head = ListView1.ColumnHeaders.Add(, , "CURRENCY", ListView1.Width / 4)
Set head = ListView1.ColumnHeaders.Add(, , "NOMINAL", ListView1.Width / 2 - 100, 1)
ListView1.View = lvwReport
With byr_frm.Data1(x).Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !tgl_byr)
        dtl.SubItems(1) = !curr
        dtl.SubItems(2) = Format(!Nominal, "###,###.##")
        ListBox1(0).AddItem !tgl_byr
        ListBox1(1).AddItem !curr
        ListBox1(2).AddItem Format(!Nominal, "###,###.##")
        .MoveNext
    Loop
End If
End With
End Sub

Private Sub Command1_Click()
Command1.Visible = False
'CommonDialog1.ShowPrinter
Me.PrintForm
Command1.Visible = True
End Sub

Private Sub Form_Load()
isi_data
End Sub

Private Sub Form_Unload(Cancel As Integer)
byr_frm.Enabled = True
End Sub

