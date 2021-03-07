VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form PrintLG_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Letter of Guarantee"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12435
   Icon            =   "PrintLG_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Print"
      DownPicture     =   "PrintLG_frm.frx":1CCA
      Height          =   855
      Left            =   9360
      Picture         =   "PrintLG_frm.frx":2994
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____________________________________________________________________________________________________________"
      Height          =   195
      Index           =   29
      Left            =   720
      TabIndex        =   33
      Top             =   8280
      Width           =   9720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____________________________________________________________________________________________________________"
      Height          =   195
      Index           =   28
      Left            =   840
      TabIndex        =   32
      Top             =   1080
      Width           =   9720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"PrintLG_frm.frx":365E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   27
      Left            =   720
      TabIndex        =   31
      Top             =   8640
      Width           =   9060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal"
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
      Index           =   26
      Left            =   3240
      TabIndex        =   30
      Top             =   6360
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr"
      Height          =   195
      Index           =   25
      Left            =   2760
      TabIndex        =   29
      Top             =   6360
      Width           =   285
   End
   Begin MSForms.ListBox list1 
      Height          =   855
      Index           =   1
      Left            =   5520
      TabIndex        =   28
      Top             =   2640
      Width           =   4455
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7858;1428"
      MatchEntry      =   0
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox list2 
      Height          =   1815
      Left            =   960
      TabIndex        =   27
      Top             =   3960
      Width           =   11055
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "19500;3148"
      MatchEntry      =   0
      BorderColor     =   -2147483643
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ListBox list1 
      Height          =   855
      Index           =   0
      Left            =   960
      TabIndex        =   26
      Top             =   2640
      Width           =   4335
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "7646;1428"
      MatchEntry      =   0
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( Nama User )"
      Height          =   195
      Index           =   24
      Left            =   960
      TabIndex        =   25
      Top             =   7920
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   8160
      TabIndex        =   24
      Top             =   3720
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   6960
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   6120
      TabIndex        =   22
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Route"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   4800
      TabIndex        =   21
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   3480
      TabIndex        =   20
      Top             =   3720
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   2400
      TabIndex        =   19
      Top             =   3720
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   1440
      TabIndex        =   18
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      Height          =   195
      Index           =   16
      Left            =   960
      TabIndex        =   17
      Top             =   7200
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your faithfully,"
      Height          =   195
      Index           =   15
      Left            =   960
      TabIndex        =   16
      Top             =   6960
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment guaranteed by WaynhSoft tour & travel, thank you very much for your kind attention and cooperation."
      Height          =   195
      Index           =   14
      Left            =   960
      TabIndex        =   15
      Top             =   6600
      Width           =   7710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nett :"
      Height          =   195
      Index           =   13
      Left            =   1920
      TabIndex        =   14
      Top             =   6360
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fare :"
      Height          =   195
      Index           =   12
      Left            =   1920
      TabIndex        =   13
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gross"
      Height          =   195
      Index           =   11
      Left            =   960
      TabIndex        =   12
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODE :"
      Height          =   195
      Index           =   10
      Left            =   960
      TabIndex        =   11
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For the following route :"
      Height          =   195
      Index           =   9
      Left            =   960
      TabIndex        =   10
      Top             =   3480
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "We kindly request you to issued ticket/s  favour of :"
      Height          =   195
      Index           =   8
      Left            =   960
      TabIndex        =   9
      Top             =   2400
      Width           =   3645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jakarta, "
      Height          =   195
      Index           =   7
      Left            =   10080
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Up."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Letter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LETTER OF GUARANTEE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   9735
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
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   1905
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
      Left            =   6120
      TabIndex        =   1
      Top             =   360
      Width           =   1665
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
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   840
      Picture         =   "PrintLG_frm.frx":36ED
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   9600
      Picture         =   "PrintLG_frm.frx":19B57
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "PrintLG_frm"
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
    y = MsgBox("Cetak header/Kop...???", vbYesNo, "Cetak LG")
    If y = vbNo Then
        umpet
    End If
    MsgBox "Apakah sudah siap mencetak?", vbOKOnly, "Printing"
    Command1.Visible = False
    PrintLG_frm.PrintForm
    Command1.Visible = True
    Image1.Visible = True
    Label1(0).Visible = True
    Label1(1).Visible = True
    Label1(2).Visible = True
    Label1(3).Visible = True
    Image2.Visible = True
    Label1(27).Visible = True
    Label1(28).Visible = True
    Label1(29).Visible = True
        With Data1.Recordset
        .AddNew
        !no_sistem = Label1(4).Caption
        !no_kertas = x
        .Update
        Data1.Refresh
    End With
End If
End Sub

Sub umpet()
    Image1.Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Label1(2).Visible = False
    Label1(3).Visible = False
    Label1(27).Visible = False
    Label1(28).Visible = False
    Label1(29).Visible = False
    Image2.Visible = False
End Sub

Private Sub isi_list1()
Dim no As Byte
Dim a As Byte
Dim nama1, nama2 As String
List1(0).Clear
List1(1).Clear
CetakLg_frm.Data2.RecordSource = "select * from trans_lg order by psg_name asc"
CetakLg_frm.Data2.Refresh
With CetakLg_frm.Data2.Recordset
.MoveFirst
no = 1
nama2 = " "
Do While Not .EOF
    Do While no < 5 And Not .EOF
        nama1 = !psg_name
        If nama1 <> nama2 Then
            List1(0).AddItem (Format(no, "#00") & ". " & !psg_sex & " " & !psg_name)
            no = no + 1
        End If
        nama2 = !psg_name
        .MoveNext
    Loop
    If Not .EOF And no >= 5 Then
    nama1 = !psg_name
        If nama1 <> nama2 Then
            List1(1).AddItem (Format(no, "#00") & ". " & !psg_sex & " " & !psg_name)
            no = no + 1
        End If
    nama2 = !psg_name
    .MoveNext
    End If
Loop
End With
End Sub

Sub isi_list2()
Dim no As Byte
Dim a, b As String
Dim c, d As String
'Dim e, f As String
CetakLg_frm.Data2.RecordSource = "select * from trans_lg order by kode_maskapai,flight_no asc"
'CetakLg_frm.Data2.RecordSource = "select * from trans_lg order by kode_maskapai asc"
CetakLg_frm.Data2.Refresh
With CetakLg_frm.Data2.Recordset
.MoveFirst
no = 1
e = ""
f = ""
Do While Not .EOF
    c = !kode_maskapai & !flight_no
    'd = !kode_maskapai
    If c <> e Then
        With dbtiket_frm.Data3.Recordset
            .MoveFirst
            Do While Not .EOF
                If CetakLg_frm.Data2.Recordset!flight_no = !no_penerbangan And CetakLg_frm.Data2.Recordset!kode_maskapai = !kode_maskapai Then
                    a = Format(!dep, "hh:mm")
                    b = Format(!arr, "hh:mm")
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
        list2.AddItem (Format(no, "#00") & ". " & !kode_maskapai & !flight_no & "     " & vbTab & !Class & vbTab & Format(!tgl_berangkat, "d mmm yyyy") & vbTab & !From & "-" & !To & "   " & vbTab & a & vbTab & b & vbTab & vbTab & !Status)
        no = no + 1
        'f = !kode_maskapai
    End If
    e = !kode_maskapai & !flight_no
    .MoveNext
Loop
End With
'list2.AddItem ""
'list2.AddItem ("Total Pax = " & CetakLg_frm.Data2.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbtravel"
Data1.RecordSource = "data_cetak"
Label1(7).Caption = "Jakarta, " & Format(Date, "d mmmm yyyy")
With CetakLg_frm
Label1(4).Caption = .Label1(8).Caption
Label1(5).Caption = "To. " & .Combo6
Label1(6).Caption = "Up. " & .Combo7
Label1(10).Caption = "CODE : " & .Text11
Label1(24).Caption = "( " & main_form.Text1(1) & " )"
Label1(25).Caption = .Label1(6)
Label1(26).Caption = .Text2
End With
isi_list1
isi_list2
End Sub

