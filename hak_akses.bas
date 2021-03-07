Attribute VB_Name = "hak_akses"
Public Sub hak_akses1()
Dim id, jab, hak As String
With dbuser_frm.Data1.Recordset
dbuser_frm.Data1.Refresh
.MoveFirst
Do While Not .EOF
    If main_form.Text1(0) = !id_user Then
    id = !id_user
    jab = !jabatan
        If !hak_pengguna = -1 Then
            main_form.admin_mnu.Visible = True
            main_form.Command2.Visible = True
            stoktiket_frm.Command3.Visible = True
            stoktiket_frm.Command6.Visible = False
            client_frm.Command1(2).Visible = True
            dbpartner_frm.Command1(2).Visible = True
            client_frm.Command1(3).Visible = False
            dbpartner_frm.Command1(3).Visible = False
        Else
            main_form.admin_mnu.Visible = False
            main_form.Command2.Visible = False
            stoktiket_frm.Command3.Visible = False
            stoktiket_frm.Command6.Visible = True
            client_frm.Command1(2).Visible = False
            dbpartner_frm.Command1(2).Visible = False
            client_frm.Command1(3).Visible = True
            dbpartner_frm.Command1(3).Visible = True
        End If
        If !dbtiket = -1 Then
            main_form.dbtiket_mnu.Visible = True
            main_form.Toolbar1.Buttons(1).Visible = True
        Else
            main_form.dbtiket_mnu.Visible = False
            main_form.Toolbar1.Buttons(1).Visible = False
        End If
        If !stok_tiket = -1 Then
            main_form.stok_mnu.Visible = True
            main_form.Toolbar1.Buttons(2).Visible = True
        Else
            main_form.stok_mnu.Visible = False
            main_form.Toolbar1.Buttons(2).Visible = False
        End If
        If !lg = -1 Then
            main_form.LG_mnu.Visible = True
            main_form.Toolbar1.Buttons(5).Visible = True
        Else
            main_form.LG_mnu.Visible = False
            main_form.Toolbar1.Buttons(5).Visible = False
        End If
        If !jual_tiket = -1 Then
            main_form.pesan_mnu.Visible = True
            main_form.Toolbar1.Buttons(6).Visible = True
        Else
            main_form.pesan_mnu.Visible = False
            main_form.Toolbar1.Buttons(6).Visible = False
        End If
        If !client = -1 Then
            main_form.dbClient_mnu.Visible = True
            main_form.Toolbar1.Buttons(3).Visible = True
        Else
            main_form.dbClient_mnu.Visible = False
            main_form.Toolbar1.Buttons(3).Visible = False
        End If
        If !suplier = -1 Then
            main_form.partner_mnu.Visible = True
            main_form.Toolbar1.Buttons(4).Visible = True
        Else
            main_form.partner_mnu.Visible = False
            main_form.Toolbar1.Buttons(4).Visible = False
        End If
        If !ntiket = -1 Then
            main_form.ntiket_mnu.Visible = True
        Else
            main_form.ntiket_mnu.Visible = False
        End If
        If !laporan = -1 Then
            main_form.lap_mnu.Visible = True
            InvNt_frm.lap_nt.Visible = True
        Else
            main_form.lap_mnu.Visible = False
            InvNt_frm.lap_nt.Visible = False
        End If
        If !akuntansi = -1 Then
            main_form.Akt_mnu.Visible = True
        Else
            main_form.Akt_mnu.Visible = False
        End If
        .MoveLast
    End If
    .MoveNext
Loop
End With
End Sub

