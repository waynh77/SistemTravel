Attribute VB_Name = "setdbtravel"
Option Explicit
Public db, db2 As String



Private Sub buka_db()
'    db = "C:\Documents and Settings\wahyu\My Documents\waynh project\vb project\travel project\DBtravel_access\dbtravel.mdb"
    db = App.Path & "\dbtravel.mdb"
End Sub

Private Sub buka_db2()
    db2 = App.Path & "\DBtravel_fox"
'    DB2 = "C:\Documents and Settings\wahyu\My Documents\waynh project\vb project\travel project\DBtravel_fox\db_user.dbf"
End Sub

Public Sub dbuser_awal()
buka_db
'db2 As Database
'UserPass.Data1.Connect = "FoxPro 3.0;"
UserPass.Data1.Connect = "access"
UserPass.Data1.DatabaseName = db
UserPass.Data1.RecordSource = "db_user"
End Sub

Public Sub dbuser()
buka_db
'dbuser_frm.Data1.Connect = "FoxPro 3.0;"
dbuser_frm.Data1.Connect = "access"
dbuser_frm.Data1.DatabaseName = db
dbuser_frm.Data1.RecordSource = "db_user"
End Sub

Public Sub gantipass()
buka_db
'Gantipass_frm.Data1.Connect = "FoxPro 3.0;"
Gantipass_frm.Data1.Connect = "access"
Gantipass_frm.Data1.DatabaseName = db
Gantipass_frm.Data1.RecordSource = "db_user"
End Sub

Public Sub dbtiket()
buka_db
dbtiket_frm.Data1.DatabaseName = db
dbtiket_frm.Data1.RecordSource = "db_maskapai"
dbtiket_frm.Data2.DatabaseName = db
dbtiket_frm.Data2.RecordSource = "db_lokasi"
dbtiket_frm.Data3.DatabaseName = db
dbtiket_frm.Data3.RecordSource = "select * from db_penerbangan  order by kode_maskapai,no_penerbangan asc"
'dbtiket_frm.DBCombo1.DataSource = "data2"
'dbtiket_frm.DBCombo2.DataSource = "data2"
'dbtiket_frm.DBCombo3.DataSource = "data1"
dbtiket_frm.DBCombo1.ListField = "kode_lokasi"
dbtiket_frm.DBCombo2.ListField = "kode_lokasi"
dbtiket_frm.DBCombo3.ListField = "kode_maskapai"
End Sub

Public Sub stock_tiket()
buka_db
stoktiket_frm.Data1.DatabaseName = db
stoktiket_frm.Data1.RecordSource = "stok_tiket"
End Sub

Public Sub dbclient()
buka_db
client_frm.Data1.DatabaseName = db
client_frm.Data1.RecordSource = "select * from db_client order by company "
End Sub

Public Sub dbpartner()
buka_db
dbpartner_frm.Data1.DatabaseName = db
dbpartner_frm.Data1.RecordSource = "db_partner"
End Sub

Public Sub dblg()
buka_db
CetakLg_frm.Data1.DatabaseName = db
CetakLg_frm.Data1.RecordSource = "select * from db_log order by no_lg"
CetakLg_frm.Data2.DatabaseName = db
CetakLg_frm.Data2.RecordSource = "trans_lg"
CetakLg_frm.Data3.DatabaseName = db
CetakLg_frm.Data3.RecordSource = "db_client"
End Sub

Public Sub db_invlg()
buka_db
InvLG_frm.Data1.DatabaseName = db
InvLG_frm.Data1.RecordSource = "db_log"
InvLG_frm.Data2.DatabaseName = db
InvLG_frm.Data2.RecordSource = "db_client"
InvLG_frm.Data3.DatabaseName = db
InvLG_frm.Data3.RecordSource = "select * from db_invoice order by no_invoice"
InvLG_frm.Data4.DatabaseName = db
InvLG_frm.Data4.RecordSource = "temp_inv"
InvLG_frm.Data5.DatabaseName = db
InvLG_frm.Data5.RecordSource = "temp_stok_tiket"
End Sub

Public Sub db_invnt()
buka_db
InvNt_frm.Data1.DatabaseName = db
InvNt_frm.Data1.RecordSource = "select * from inv_nt order by no_inv"
End Sub

Public Sub db_vcr()
'buka_db
'CetakVoucer_frm.Data1.DatabaseName = db
'CetakVoucer_frm.Data1.RecordSource = "vcrHotel"
End Sub

Public Sub db_byr()
buka_db
With byr_frm
    .Data1(0).DatabaseName = db
    .Data1(1).DatabaseName = db
    .Data1(2).DatabaseName = db
    .Data2(0).DatabaseName = db
    .Data2(1).DatabaseName = db
    .Data2(2).DatabaseName = db
    .Data3.DatabaseName = db
    .Data3.RecordSource = "Pembayaran"
End With
End Sub

Public Sub db_cari()
buka_db
CariByr_frm.Data1.DatabaseName = db
End Sub

Public Sub db_remain()
buka_db
With Remainder_frm
    .Data1.DatabaseName = db
    .Data2.DatabaseName = db
    .Data2.RecordSource = "select * from pembayaran where sisa<>0"
    .Data1.RecordSource = "Remainder"
End With
End Sub

Public Sub db_main()
buka_db
main_form.Data1.DatabaseName = db
End Sub

Public Sub db_remain2()
buka_db
REmain_frm.Data1.DatabaseName = db
End Sub

Public Sub db_ctkulang()
buka_db
CtkUlang_frm.Data1.DatabaseName = db
CtkUlang_frm.Data2.DatabaseName = db
End Sub

Public Sub db_hotel()
buka_db
Hotel_frm.Data1.DatabaseName = db
Hotel_frm.Data1.RecordSource = "vcrhotel"
End Sub
