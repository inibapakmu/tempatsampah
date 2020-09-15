Imports MySql.Data.MySqlClient
Imports System.Globalization
Public Class mainmenu

    
    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For Each RunningProcess In Process.GetProcessesByName("pulungsari")
            RunningProcess.Kill()
        Next
        Me.Close()

    End Sub
    Private Sub StokKurangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("Http://localhost:81/programresto/index5.php")
    End Sub


    Dim stokbind, saldobind, pnjbind4, pcus2, pp As New BindingSource
    Dim con As New MySqlConnection("server=localhost;user id=root;password=10092005;database=pulungsari")
    Public Function open() As MySqlConnection
        Try
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
        Catch ex As Exception
            MsgBox("Koneksi ke Database Gagal ! Silakan Coba Kembali Beberapa Saat Lagi", MsgBoxStyle.OkOnly, "Error")
        End Try
        Return con
    End Function
    Public Function conclose() As MySqlConnection
        con.Close()
        Return con
    End Function
    Public Function getdataset(ByVal sql As String) As DataSet
        Dim adapter As New MySqlDataAdapter(sql, con)
        Dim mydata As New DataSet
        adapter.Fill(mydata, "data")
        Return mydata
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        
    End Sub

    


    Private Sub mainmenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub


    Private Sub KalkulatorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim MyApp As New System.Diagnostics.ProcessStartInfo("Calc")
        Process.Start(MyApp)
    End Sub



    Private Sub SupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierToolStripMenuItem.Click
        datasupplier.MdiParent = Me
        datasupplier.Show()

    End Sub

    Private Sub SatuanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        satuan.MdiParent = Me
        satuan.Show()

    End Sub

    Private Sub KategoriToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KategoriToolStripMenuItem.Click
        ketegori.MdiParent = Me
        ketegori.Show()

    End Sub

    Private Sub CustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerToolStripMenuItem.Click
        datacustomer.MdiParent = Me
        datacustomer.Show()

    End Sub

    Private Sub SalesmanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesmanToolStripMenuItem.Click
        datasales.MdiParent = Me
        datasales.Show()

    End Sub

    Private Sub SupervisorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupervisorToolStripMenuItem.Click
        supervisor.MdiParent = Me
        supervisor.Show()

    End Sub

    Private Sub InputBaruToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputBaruToolStripMenuItem.Click
        inputstok.MdiParent = Me
        inputstok.Show()

    End Sub

    Private Sub DataBarangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataBarangToolStripMenuItem.Click
        datastok.MdiParent = Me
        datastok.Show()

    End Sub

    Private Sub UbahPasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UbahPasswordToolStripMenuItem.Click
        editpasskeu.MdiParent = Me
        editpasskeu.Show()

    End Sub

    Private Sub AddUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddUserToolStripMenuItem.Click

    End Sub

    Private Sub LogoutToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        For Each RunningProcess In Process.GetProcessesByName("pulungsari")
            RunningProcess.Kill()
        Next
        Me.Close()

    End Sub

    Private Sub PenyesuaianStokToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenyesuaianStokToolStripMenuItem.Click
        penyesuaian.MdiParent = Me
        penyesuaian.Show()

    End Sub

    Private Sub CetakUlangPenjualanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CetakUlangPenjualanToolStripMenuItem.Click
        cetakulangfaktur.MdiParent = Me
        cetakulangfaktur.Show()

    End Sub

    Private Sub DeletePenjualanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        cetakulangfaktur2.MdiParent = Me
        cetakulangfaktur2.Show()

    End Sub

    Private Sub SejarahTransaksiNotaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SejarahTransaksiNotaToolStripMenuItem.Click
        reknota.MdiParent = Me
        reknota.Show()

    End Sub

    Private Sub InputPembelianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputPembelianToolStripMenuItem.Click
        inputpembeliankredit.MdiParent = Me
        inputpembeliankredit.Show()

    End Sub

    Private Sub DeletePembelianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        lappembelian2.MdiParent = Me
        lappembelian2.Show()

    End Sub

    Private Sub BankToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankToolStripMenuItem.Click
        inputbank.MdiParent = Me
        inputbank.Show()

    End Sub

    Private Sub PerkiraanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerkiraanToolStripMenuItem.Click
        dataperkiraan.MdiParent = Me
        dataperkiraan.Show()

    End Sub

    Private Sub InputKasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputKasToolStripMenuItem.Click
        biayalain.MdiParent = Me
        biayalain.Show()

    End Sub

    Private Sub InputHutangKaryawanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputHutangKaryawanToolStripMenuItem.Click
        pendapatanlain.MdiParent = Me
        pendapatanlain.Show()

    End Sub

    Private Sub LaluLintasKasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaluLintasKasToolStripMenuItem.Click
        lalulintakas.MdiParent = Me
        lalulintakas.Show()

    End Sub

    Private Sub HutangPegawaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HutangPegawaiToolStripMenuItem.Click
        hutangpegawai.MdiParent = Me
        hutangpegawai.Show()

    End Sub

    Private Sub CicilanPegawaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CicilanPegawaiToolStripMenuItem.Click
        cicilan.MdiParent = Me
        cicilan.Show()

    End Sub

    Private Sub PegawaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PegawaiToolStripMenuItem.Click
        datapegawai.MdiParent = Me
        datapegawai.Show()

    End Sub

    Private Sub JabatanPegawaiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JabatanPegawaiToolStripMenuItem.Click
        datajabatan.MdiParent = Me
        datajabatan.Show()

    End Sub

    Private Sub InputPenjualanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputPenjualanToolStripMenuItem.Click
        inputpenjualan.MdiParent = Me
        inputpenjualan.Show()

    End Sub

    Private Sub EditPenjualanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        editpenjualan.MdiParent = Me
        editpenjualan.Show()

    End Sub

    Private Sub EditPembelianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        editpembelian.MdiParent = Me
        editpembelian.Show()

    End Sub

    Private Sub DeleteReturPembelianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ReturPembelianToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturPembelianToolStripMenuItem.Click
        inputreturbeli.MdiParent = Me
        inputreturbeli.Show()

    End Sub

    Private Sub PenerimaanPiutangPelangganToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenerimaanPiutangPelangganToolStripMenuItem.Click
        inputbayarpiutang.MdiParent = Me
        inputbayarpiutang.Show()

    End Sub

    Private Sub PencairanCekGiroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PencairanCekGiroToolStripMenuItem.Click
        pencairancek.MdiParent = Me
        pencairancek.Show()

    End Sub

    Private Sub TolakanCekGiroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TolakanCekGiroToolStripMenuItem.Click
        tolakancek.MdiParent = Me
        tolakancek.Show()

    End Sub

    Private Sub PenggantianCekGiroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenggantianCekGiroToolStripMenuItem.Click
        ganticek.MdiParent = Me
        ganticek.Show()

    End Sub

    Private Sub LaporanPenjualanGlobalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPenjualanGlobalToolStripMenuItem.Click
        lappenjualanglobal.MdiParent = Me
        lappenjualanglobal.Show()

    End Sub

    Private Sub LaporanPenjualanRinciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPenjualanRinciToolStripMenuItem.Click
        lappenjualanprin.MdiParent = Me
        lappenjualanprin.Show()

    End Sub

    Private Sub LaporanPenjualanPerKategoriToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPenjualanPerKategoriToolStripMenuItem.Click
        omsetexppn3.MdiParent = Me
        omsetexppn3.Show()

    End Sub

    Private Sub LaporanPenjualanPerSalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPenjualanPerSalesToolStripMenuItem.Click
        omsetexppn2.MdiParent = Me
        omsetexppn2.Show()

    End Sub

    Private Sub OmsetPerCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OmsetPerCustomerToolStripMenuItem.Click
        lappenjualanpercus.MdiParent = Me
        lappenjualanpercus.Show()

    End Sub

    Private Sub LaporanPembelianGlobalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub LaporanPembelianRinciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPembelianRinciToolStripMenuItem.Click
        lappembelian.MdiParent = Me
        lappembelian.Show()

    End Sub

    Private Sub LaporanStokToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanStokToolStripMenuItem1.Click
        lapstok.MdiParent = Me
        lapstok.Show()

    End Sub

    Private Sub MarginPeriodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarginPeriodeToolStripMenuItem.Click
        lapsellout.MdiParent = Me
        lapsellout.Show()

    End Sub

    Private Sub KasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KasToolStripMenuItem.Click
        lapkas.MdiParent = Me
        lapkas.Show()

    End Sub

    Private Sub BukuBesarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BukuBesarToolStripMenuItem.Click
        bukubesar.MdiParent = Me
        bukubesar.Show()

    End Sub

    Private Sub BukuBankToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BukuBankToolStripMenuItem.Click
        lapbank.MdiParent = Me
        lapbank.Show()

    End Sub

    Private Sub LaporanUmurPiutangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanUmurPiutangToolStripMenuItem.Click
        lapumurpiutang.MdiParent = Me
        lapumurpiutang.Show()

    End Sub

    Private Sub LaporanPenagihanSalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPenagihanSalesToolStripMenuItem.Click
        kartupiutang.MdiParent = Me
        kartupiutang.Show()

    End Sub

    Private Sub LaporanPembayaranPiutangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPembayaranPiutangToolStripMenuItem.Click
        lappembayaranpiutang.MdiParent = Me
        lappembayaranpiutang.Show()

    End Sub

    Private Sub LaporanUmurHutangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanUmurHutangToolStripMenuItem.Click
        lapumurhutang.MdiParent = Me
        lapumurhutang.Show()

    End Sub

    Private Sub RincianHutangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RincianHutangToolStripMenuItem.Click
        kartuhutang.MdiParent = Me
        kartuhutang.Show()

    End Sub

    Private Sub LaporanPembayaranHutangToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaporanPembayaranHutangToolStripMenuItem.Click
        lappembayaranhutang.MdiParent = Me
        lappembayaranhutang.Show()

    End Sub

    Private Sub RumusPayrollToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RumusPayrollToolStripMenuItem.Click
        datarumus.MdiParent = Me
        datarumus.Show()

    End Sub

    Private Sub DataPOOrderHPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataPOOrderHPToolStripMenuItem.Click
        datapo.MdiParent = Me
        datapo.Show()

    End Sub

    Private Sub InputPenjualanPOSalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputPenjualanPOSalesToolStripMenuItem.Click
        inputpenjualan2.MdiParent = Me
        inputpenjualan2.Show()

    End Sub

    Private Sub DaftarPembayaranHariIniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DaftarPembayaranHariIniToolStripMenuItem.Click
        lapharianbayar.MdiParent = Me
        lapharianbayar.Show()

    End Sub

    Private Sub ReturPenjualanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturPenjualanToolStripMenuItem.Click
        inputreturjual.MdiParent = Me
        inputreturjual.Show()

    End Sub

    Private Sub ReturPenjualanRupiahToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturPenjualanRupiahToolStripMenuItem.Click
        inputreturjual2.MdiParent = Me
        inputreturjual2.Show()

    End Sub
End Class
