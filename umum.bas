Attribute VB_Name = "Auto_Number"
Public Function user_auto()
Dim urutan As String * 3
Dim hitung As Single
With dbuser_frm.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "001"
    Else
        .MoveLast
        If Val(Left(.Fields("id_user"), 0)) <> "000" Then
            urutan = "001"
        Else
            hitung = Val(Right(.Fields("id_user"), 3)) + 1
            urutan = Right("000" & hitung, 3)
        End If
    End If
    dbuser_frm.Text1.Text = urutan
End With
End Function

Public Function arl_auto()
Dim urutan As String * 3
Dim hitung As Single
With dbtiket_frm.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_maskapai"), 0)) <> "000" Then
            urutan = "001"
        Else
            hitung = Val(Right(.Fields("no_maskapai"), 3)) + 1
            urutan = Right("000" & hitung, 3)
        End If
    End If
    dbtiket_frm.Text1.Text = urutan
End With
End Function

Public Function lg_auto()
Dim urutan As String * 10
Dim hitung As Single
CetakLg_frm.Data1.RecordSource = "select * from db_log WHERE left(no_lg,3)='LoG' order by no_lg"
CetakLg_frm.Data1.Refresh
With CetakLg_frm.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "LoG" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_lg"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_lg"), 7)) + 1
            urutan = "LoG" & Right("0000000" & hitung, 7)
        End If
    End If
    CetakLg_frm.Label1(8).Caption = urutan
End With
End Function

Public Function inv_auto()
Dim urutan As String * 10
Dim hitung As Single
InvLG_frm.Data3.RecordSource = "select * from db_invoice WHERE left(no_invoice,3)='Inv' order by no_invoice"
InvLG_frm.Data3.Refresh
With InvLG_frm.Data3.Recordset
    If .RecordCount = 0 Then
        urutan = "Inv" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_invoice"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_invoice"), 7)) + 1
            urutan = "Inv" & Right("0000000" & hitung, 7)
        End If
    End If
    InvLG_frm.Label1(3).Caption = urutan
End With
End Function

Public Function invnt_auto()
Dim urutan As String * 10
Dim hitung As Single
InvNt_frm.Data1.RecordSource = "select * from inv_nt order by no_inv"
InvNt_frm.Data1.Refresh
With InvNt_frm.Data1.Recordset
    If .BOF Then
        urutan = "Int" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_inv"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_inv"), 7)) + 1
            urutan = "Int" & Right("0000000" & hitung, 7)
        End If
    End If
    InvNt_frm.Label1(7).Caption = urutan
End With
End Function

Public Function client_auto()
Dim urutan As String * 8
Dim hitung As Single
With client_frm.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "CNT" & "00001"
    Else
        .MoveLast
        If Val(Left(.Fields("id_client"), 5)) <> "00000" Then
            urutan = "00000" & "00001"
        Else
            hitung = Val(Right(.Fields("id_client"), 5)) + 1
            urutan = "CNT" & Right("00000" & hitung, 5)
        End If
    End If
    client_frm.Text9 = urutan
    InvLG_frm.Text1 = urutan
End With
End Function

Public Function partner_auto()
Dim urutan As String * 8
Dim hitung As Single
With dbpartner_frm.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "PRT" & "00001"
    Else
        .MoveLast
        If Val(Left(.Fields("id_partner"), 5)) <> "00000" Then
            urutan = "00000" & "00001"
        Else
            hitung = Val(Right(.Fields("id_partner"), 5)) + 1
            urutan = "PRT" & Right("00000" & hitung, 5)
        End If
    End If
    dbpartner_frm.Text9 = urutan
End With
End Function

Public Sub isi_cekdat()
Dim a, b, c, d As String
Dim e, f, g As String
Dim h, k As Single
Dim i, j As String
With main_form
.List1.Clear
.List1.AddItem ("")
.List1.AddItem ("STATUS DATABASE")
.List1.AddItem ("---------------")
dbtiket_frm.Data1.Refresh
If dbtiket_frm.Data1.Recordset.BOF Or dbtiket_frm.Data2.Recordset.BOF Or dbtiket_frm.Data3.Recordset.BOF Then
    a = "Belum Lengkap"
    If dbtiket_frm.Data1.Recordset.BOF Then
        b = "Masih kosong!"
    Else
        b = "Ok"
    End If
    If dbtiket_frm.Data2.Recordset.BOF Then
        c = "Masih kosong!"
    Else
        c = "Ok"
    End If
    If dbtiket_frm.Data2.Recordset.BOF Then
        d = "Masih kosong!"
    Else
        d = "Ok"
    End If
    .Timer8(0).Enabled = True
Else
    a = "Ok"
    b = "Ok"
    c = "Ok"
    d = "Ok"
    .Timer8(0).Enabled = False
    .Timer8(1).Enabled = False
    .Toolbar1.Buttons(1).Image = 1
End If
client_frm.Data1.Refresh
If client_frm.Data1.Recordset.BOF Then
    e = "Masih kosong"
    .Timer6(0).Enabled = True
Else
    e = "Ok"
    .Timer6(0).Enabled = False
    .Timer6(1).Enabled = False
    .Toolbar1.Buttons(3).Image = 3
End If
dbpartner_frm.Data1.Refresh
If dbpartner_frm.Data1.Recordset.BOF Then
    f = "Masih kosong"
    .Timer7(0).Enabled = True
Else
    f = "Ok"
    .Timer7(0).Enabled = False
    .Timer7(1).Enabled = False
    .Toolbar1.Buttons(4).Image = 4
End If
stoktiket_frm.Data1.Refresh
If stoktiket_frm.Data1.Recordset.BOF Then
    g = "Stok kosong"
    .Timer5(0).Enabled = True
    .Timer5(1).Enabled = False
Else
    g = "Ok"
    h = stoktiket_frm.Data1.Recordset.RecordCount
    .Timer5(0).Enabled = False
    .Timer5(1).Enabled = False
    .Toolbar1.Buttons(2).Image = 2
End If

.List1.AddItem ("DB Tiket..." & a)
.List1.AddItem ("   Data Airlines..." & b)
.List1.AddItem ("   Data Lokasi..." & c)
.List1.AddItem ("   Data Penerbangan..." & d)
.List1.AddItem ("DB Client..." & e)
.List1.AddItem ("DB Suplier..." & f)
.List1.AddItem ("")
.List1.AddItem ("STATUS TRANSAKSI")
.List1.AddItem ("----------------")
.List1.AddItem ("Stok Tiket..." & g)
stoktiket_frm.Data1.Refresh
If stoktiket_frm.Data1.Recordset.RecordCount <> 0 Then
    .List1.AddItem ("   Jumlah Tiket..." & h)
End If
CetakLg_frm.Data2.Refresh
If CetakLg_frm.Data2.Recordset.RecordCount = 0 Then
    i = "Ready"
    .Timer2.Enabled = False
    .Timer3.Enabled = False
    .Toolbar1.Buttons(5).Image = 5
Else
    i = "ada data LG yang belum dicetak"
    .Timer2.Enabled = True
End If
.List1.AddItem ("LG..." & i)
CetakLg_frm.Data1.RecordSource = "select*from db_log"
CetakLg_frm.Data1.Refresh
With CetakLg_frm.Data1.Recordset
If Not .BOF Then
    .MoveFirst
    k = 0
    Do While Not .EOF
        If !status_lg = False Then
            k = k + 1
        End If
        .MoveNext
    Loop
End If
End With
If k = 0 Then
    j = "Ready"
    .Timer4(0).Enabled = False
    .Timer4(1).Enabled = False
    .Toolbar1.Buttons(6).Image = 6
Else
    j = k & " Tiket LG blm dicetak invoice"
    .Timer4(0).Enabled = True
End If
.List1.AddItem ("Invoice..." & j)
End With
End Sub

Public Function Auto_Vcr()
Dim urutan As String * 10
Dim hitung As Single
Call db_hotel
With Hotel_frm.Data1.Recordset
    If .BOF Then
        urutan = "Vcr" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("no_voucer"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_voucer"), 7)) + 1
            urutan = "Vcr" & Right("0000000" & hitung, 7)
        End If
    End If
    Hotel_frm.Text1 = urutan
End With
End Function

