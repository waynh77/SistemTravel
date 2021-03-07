VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form PrintInv_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Invoice"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   Icon            =   "PrintInv_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Print"
      DownPicture     =   "PrintInv_frm.frx":1CCA
      Height          =   855
      Left            =   9840
      Picture         =   "PrintInv_frm.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Invoice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   840
      TabIndex        =   35
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   25
      Left            =   8280
      TabIndex        =   34
      Top             =   7200
      Width           =   570
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   7680
      Top             =   720
      Width           =   3375
   End
   Begin VB.Line Line5 
      X1              =   840
      X2              =   840
      Y1              =   2400
      Y2              =   6240
   End
   Begin VB.Line Line4 
      X1              =   11160
      X2              =   11160
      Y1              =   2400
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   8160
      X2              =   8160
      Y1              =   2400
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   8160
      X2              =   11160
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Index           =   2
      X1              =   840
      X2              =   11160
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Index           =   1
      X1              =   840
      X2              =   11160
      Y1              =   2760
      Y2              =   2760
   End
   Begin MSForms.ListBox list3 
      Height          =   3135
      Left            =   9000
      TabIndex        =   33
      Top             =   3000
      Width           =   2055
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3625;5530"
      MatchEntry      =   0
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   10440
      TabIndex        =   31
      Top             =   6600
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   23
      Left            =   9600
      TabIndex        =   30
      Top             =   6600
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft Signature is no longer required"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   22
      Left            =   7440
      TabIndex        =   29
      Top             =   7680
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is a computer generated invoice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   7440
      TabIndex        =   28
      Top             =   7440
      Width           =   2790
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  Bank Mandiri Permata Hijau : ACC No : 102.000.422.969.3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   960
      TabIndex        =   27
      Top             =   8520
      Width           =   4440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   19
      Left            =   1080
      TabIndex        =   26
      Top             =   8280
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  or Bank Transfer to the following accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   960
      TabIndex        =   25
      Top             =   8040
      Width           =   3210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. Cheque should be crossed and made payable to ""WaynhSoft"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   17
      Left            =   840
      TabIndex        =   24
      Top             =   7800
      Width           =   4650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Exchange rate used as date of payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   840
      TabIndex        =   23
      Top             =   7560
      Width           =   3150
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   6960
      Top             =   7080
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   6960
      TabIndex        =   22
      Top             =   6840
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   14
      Left            =   7440
      TabIndex        =   21
      Top             =   6600
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( Name )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   3840
      TabIndex        =   20
      Top             =   6600
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issued by,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   3840
      TabIndex        =   19
      Top             =   6240
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(                               )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   1125
      TabIndex        =   18
      Top             =   6600
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receive by,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   1200
      TabIndex        =   17
      Top             =   6240
      Width           =   1365
   End
   Begin MSForms.ListBox list1 
      Height          =   3135
      Left            =   960
      TabIndex        =   16
      Top             =   3000
      Width           =   7095
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "12515;5212"
      MatchEntry      =   0
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   8280
      TabIndex        =   15
      Top             =   2520
      Width           =   2475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   3825
      TabIndex        =   14
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPWP : 02.098.891.1-013.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   3075
      TabIndex        =   13
      Top             =   2040
      Width           =   2190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Index           =   0
      X1              =   840
      X2              =   11160
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   7920
      TabIndex        =   12
      Top             =   1680
      Width           =   2115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   7920
      TabIndex        =   11
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   7920
      TabIndex        =   10
      Top             =   1200
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   7920
      TabIndex        =   9
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVOICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   10440
      TabIndex        =   8
      Top             =   360
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone : 021 - 570 6753, 570 6754   Fax : 021 - 5712443"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2190
      TabIndex        =   7
      Top             =   1560
      Width           =   4035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RT 10 RW 03 Bendungan Hilir (Benhil) Jakarta Pusat 10210"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1995
      TabIndex        =   6
      Top             =   1320
      Width           =   4230
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jl.Danau Toba Blok F3 No. 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   4215
      TabIndex        =   4
      Top             =   1080
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   840
      Picture         =   "PrintInv_frm.frx":365E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "tour and travel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.wnh-it.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Letter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   810
   End
   Begin MSForms.ListBox list2 
      Height          =   3135
      Left            =   8280
      TabIndex        =   32
      Top             =   3000
      Width           =   1695
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2990;5212"
      MatchEntry      =   0
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "PrintInv_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x As String
'Me.Hide
'main_form.Show
x = InputBox("Masukan No. Kertas", "Nomor Kertas")
If x <> "" Then
    MsgBox "Apakah sudah siap cetak?", vbOKOnly, "Printing"
    Command1.Visible = False
    umpet
    PrintInv_frm.PrintForm
    Command1.Visible = True
    muncul
    With Data1.Recordset
        .AddNew
        !no_sistem = Label1(4).Caption
        !no_kertas = x
        .Update
        Data1.Refresh
    End With
End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbtravel"
Data1.RecordSource = "data_cetak"
Label1(5).Caption = "Tgl Inv : " & Format(Date, "d mmmm yyyy")
If CtkUlang_frm.Visible = False Then
    With InvLG_frm
    Label1(4).Caption = .Label1(3).Caption 'nomor
    Label2(3).Caption = .Combo1 'company
    Label2(4).Caption = .Text2 'contact
    Label2(5).Caption = .Text4 'telp
    Label2(6).Caption = .Text8 'address
    Label2(23).Caption = .Label1(44).Caption 'curr
    Label2(24).Caption = .Label1(41).Caption 'total nominal
    Label2(13).Caption = "( " & main_form.Text1(1) & " )" 'user
    Label2(25) = Format(.DTPicker1, "d mmmm yyyy") 'tgl
    End With
    isi_list
End If
End Sub

Sub isi_list()
Dim no As Byte
Dim a, b As String
List1.Clear
no = 1
With InvLG_frm.Data1.Recordset
    If InvLG_frm.Label4.Caption = "invlg" Then
    .MoveFirst
    Do While Not .EOF
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If InvLG_frm.Data1.Recordset!no_penerbangan = !no_penerbangan And InvLG_frm.Data1.Recordset!kode_maskapai = !kode_maskapai Then
                    a = Format(!dep, "hh:mm")
                    b = Format(!arr, "hh:mm")
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        List1.AddItem (Format(no, "#00") & ". " & !jenkel_penumpang & " " & !passanger_name & "  " & !kode_maskapai & !no_penerbangan & "  " & !Class & "  " & Format(!tanggal_berangkat, "d/m/yyyy") & "  " & !From & "-" & !To & "  " & a & "  " & b & "  " & !Status)
        'list2.AddItem (!Currency)
        list3.AddItem (!Currency & rkanan(!hrg_jual, "###,###,###"))
        no = no + 1
        .MoveNext
    Loop
    End If
End With
With InvLG_frm.Data4.Recordset
    If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If InvLG_frm.Data4.Recordset!no_penerbangan = !no_penerbangan And InvLG_frm.Data4.Recordset!kode_maskapai = !kode_maskapai Then
                    a = Format(!dep, "hh:mm")
                    b = Format(!arr, "hh:mm")
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        List1.AddItem (Format(no, "#00") & ". " & !sex_psg & " " & !nama_psg & "  " & !kode_maskapai & !no_penerbangan & "  " & !Class & "  " & Format(!tgl_berangkat, "d/m/yyyy") & "  " & !From & "-" & !To & "  " & a & "  " & b & "  " & !Status)
        'list2.AddItem (!Currency)
        list3.AddItem (!Currency & rkanan(!harga_tiket, "###,###,###"))
        no = no + 1
        .MoveNext
    Loop
    End If
End With
End Sub

Private Function rkanan(ndata, cformat) As String
    rkanan = Format(ndata, cformat)
    rkanan = Space(Len(cformat) - Len(rkanan)) + rkanan
End Function

Sub umpet()
Image1.Visible = False
Label1(0).Visible = False
Label1(1).Visible = False
Label1(2).Visible = False
Label2(0).Visible = False
Label2(1).Visible = False
Label2(2).Visible = False
Label2(7).Visible = False
Shape1.Visible = False
Label1(3).Visible = False
Line1(0).Visible = False
Label2(8).Visible = False
Line1(1).Visible = False
Line5.Visible = False
Line3.Visible = False
Label2(9).Visible = False
Line4.Visible = False
Line1(2).Visible = False
Label2(14).Visible = False
Label2(15).Visible = False
Line2.Visible = False
Shape2.Visible = False
Label2(21).Visible = False
Label2(22).Visible = False
Label2(10).Visible = False
Label2(12).Visible = False
Label2(11).Visible = False
'Label2(13).Visible = False
Label2(16).Visible = False
Label2(17).Visible = False
Label2(18).Visible = False
Label2(19).Visible = False
Label2(20).Visible = False
End Sub

Sub muncul()
Image1.Visible = True
Label1(0).Visible = True
Label1(1).Visible = True
Label1(2).Visible = True
Label2(0).Visible = True
Label2(1).Visible = True
Label2(2).Visible = True
'Label2(7).Visible = True
Shape1.Visible = True
Label1(3).Visible = True
Line1(0).Visible = True
Label2(8).Visible = True
Line1(1).Visible = True
Line5.Visible = True
Line3.Visible = True
Label2(9).Visible = True
Line4.Visible = True
Line1(2).Visible = True
Label2(14).Visible = True
Label2(15).Visible = True
Line2.Visible = True
Shape2.Visible = True
Label2(21).Visible = True
Label2(22).Visible = True
Label2(10).Visible = True
Label2(12).Visible = True
Label2(11).Visible = True
Label2(13).Visible = True
Label2(16).Visible = True
Label2(17).Visible = True
Label2(18).Visible = True
Label2(19).Visible = True
Label2(20).Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If CtkUlang_frm.Visible = False Then
With InvLG_frm
    If .Label4.Caption = "invlg" Then
            .Data1.Recordset.MoveFirst
            Do While Not .Data1.Recordset.EOF
                .Data3.Recordset.AddNew
                .Data3.Recordset!no_invoice = .Label1(3).Caption
                .Data3.Recordset!tgl_inv = .Label1(1).Caption
                .Data3.Recordset!kode_maskapai = .Data1.Recordset!kode_maskapai
                .Data3.Recordset!no_tiket = .Data1.Recordset!no_tiket
                .Data3.Recordset!no_penerbangan = .Data1.Recordset!no_penerbangan
                .Data3.Recordset!Class = .Data1.Recordset!Class
                .Data3.Recordset!tgl_berangkat = .Data1.Recordset!tanggal_berangkat
                .Data3.Recordset!Status = .Data1.Recordset!Status
                .Data3.Recordset!keterangan = .Data1.Recordset!keterangan
                .Data3.Recordset!sex_psg = .Data1.Recordset!jenkel_penumpang
                .Data3.Recordset!nama_psg = .Data1.Recordset!passanger_name
                .Data3.Recordset!hrg_tiket = .Data1.Recordset!hrg_jual
                .Data3.Recordset!Currency = .Data1.Recordset!Currency
                .Data3.Recordset!id_client = .Text1
                .Data3.Recordset!COMPANY = .Combo1
                .Data3.Recordset!contact_person = .Text2
                .Data3.Recordset!telp = .Text4
                .Data3.Recordset!address = .Text8
                .Data3.Recordset!no_lg = .Combo8
                .Data3.Recordset!hrg_pokok = .Data1.Recordset!harga
                .Data3.Recordset!From = .Data1.Recordset!From
                .Data3.Recordset!To = .Data1.Recordset!To
                .Data3.Recordset!due_date = .DTPicker1
                .Data3.Recordset.Update
                .Data1.Recordset.Edit
                .Data1.Recordset!status_lg = -1
                .Data1.Recordset.Update
                .Data1.Recordset.MoveNext
            Loop
            .Data1.Refresh
            .Data3.Refresh
        End If
    
    If Not .Data4.Recordset.BOF Then
        .Data4.Recordset.MoveFirst
        Do While Not .Data4.Recordset.BOF
            .Data3.Recordset.AddNew
            .Data3.Recordset!no_invoice = .Label1(3)
            .Data3.Recordset!tgl_inv = .Label1(1)
            .Data3.Recordset!kode_maskapai = .Data4.Recordset!kode_maskapai
            .Data3.Recordset!no_tiket = .Data4.Recordset!no_tiket
            .Data3.Recordset!no_penerbangan = .Data4.Recordset!no_penerbangan
            .Data3.Recordset!Class = .Data4.Recordset!Class
            .Data3.Recordset!tgl_berangkat = .Data4.Recordset!tgl_berangkat
            .Data3.Recordset!Status = .Data4.Recordset!Status
            .Data3.Recordset!keterangan = .Data4.Recordset!keterangan
            .Data3.Recordset!sex_psg = .Data4.Recordset!sex_psg
            .Data3.Recordset!nama_psg = .Data4.Recordset!nama_psg
            .Data3.Recordset!hrg_tiket = .Data4.Recordset!harga_tiket
            .Data3.Recordset!Currency = .Data4.Recordset!Currency
            .Data3.Recordset!id_client = .Text1
            .Data3.Recordset!COMPANY = .Combo1
            .Data3.Recordset!contact_person = .Text2
            .Data3.Recordset!telp = .Text4
            .Data3.Recordset!address = .Text8
            .Data3.Recordset!no_lg = ""
            .Data3.Recordset!hrg_pokok = .Data4.Recordset!hrg_pokok
            .Data3.Recordset!From = .Data4.Recordset!From
            .Data3.Recordset!To = .Data4.Recordset!To
            .Data3.Recordset!due_date = .DTPicker1
            .Data3.Recordset.Update
            With stoktiket_frm.Data1.Recordset
                If Not .BOF Then
                .MoveFirst
                Do While Not .EOF
                    If !no_tiket = InvLG_frm.Data4.Recordset!no_tiket Then
                        .Delete
                        .MoveLast
                    End If
                    If Not .BOF Then
                    .MoveNext
                    End If
                Loop
                stoktiket_frm.Data1.Refresh
                End If
            End With
        .Data4.Recordset.Delete
        .Data4.Refresh
        Loop
    End If
With .Data5.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
    End If
End With
.Data5.Refresh
.Data1.Refresh
.Data2.Refresh
.Data3.Refresh
.Data4.Refresh
End With
Call isi_cekdat
If CtkUlang_frm.Visible = False Then
    update_byr
End If
End If
End Sub

Sub hapus_tiket()
End Sub

Sub update_byr()
Call db_byr
With byr_frm.Data3.Recordset
    .AddNew
    !nomor = Label1(4)
    !curr = Label2(23)
    !nilai_trans = Format(Label2(24).Caption, "###")
    !total_bayar = 0
    !sisa = Format(Label2(24).Caption, "###")
    !frek = 0
    .Update
End With
If InvLG_frm.DTPicker1 > Date Then
    With Remainder_frm.Data1.Recordset
        .AddNew
        !nomor = InvLG_frm.Label1(3)
        !tgl = InvLG_frm.DTPicker1
        !waktu = "10:00"
        !Status = True
        !keterangan = "Tagihan Invoice " & InvLG_frm.Combo1 & " " & InvLG_frm.Text2
        .Update
        Remainder_frm.Data1.Refresh
    End With
End If
End Sub


