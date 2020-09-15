Imports MySql.Data.MySqlClient
Imports System.Globalization
Public Class datapegawai
    Dim stokbind, saldobind, pnjbind4 As New BindingSource
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
    Private Sub kosong()
        txtkode.Text = ""
        txtnama.Text = ""
        txtalamat.Text = ""
        txtnoktp.Text = ""

        txtkode.Focus()

    End Sub
    Private Sub isigrid()

        Dim mydata As DataSet = getdataset("select kode,nama,alamat,telp,noktp,user,sidikjari,jabatan,date_format(tglmasuk,'%d-%M-%Y') as tglmasuk from pegawai order by nama ")
        stokbind.DataSource = mydata.Tables("data")
        datagridstok.DataSource = stokbind
        FormatGridWithBothTableAndColumnStyles()
        datagridstok.DataBindings.Clear()

    End Sub
    Private Sub isigrid2()

        Dim mydata As DataSet = getdataset("select kode,nama,alamat,telp,noktp,user,sidikjari,jabatan,date_format(tglmasuk,'%d-%M-%Y') as tglmasuk from pegawai  where " & ComboBox1.Text & " like '%" & txtisi.Text & "%' order by nama ")
        stokbind.DataSource = mydata.Tables("data")
        datagridstok.DataSource = stokbind
        FormatGridWithBothTableAndColumnStyles()
        datagridstok.DataBindings.Clear()

    End Sub

    Private Sub FormatGridWithBothTableAndColumnStyles()
        datagridstok.Columns(0).Width = 50
        datagridstok.Columns(1).Width = 150
        datagridstok.Columns(2).Width = 250
        datagridstok.Columns(3).Width = 150
        datagridstok.Columns(4).Width = 100
        datagridstok.Columns(5).Width = 100
        datagridstok.Columns(6).Width = 180
        datagridstok.Columns(7).Width = 180
        datagridstok.Columns(8).Width = 180
        datagridstok.Columns(0).HeaderText = "Kode"
        datagridstok.Columns(1).HeaderText = "Nama"
        datagridstok.Columns(2).HeaderText = "Alamat"
        datagridstok.Columns(3).HeaderText = "Telp"

        datagridstok.Columns(4).HeaderText = "noktp"

        datagridstok.Columns(5).HeaderText = "user"

        datagridstok.Columns(6).HeaderText = "sidikjari"

        datagridstok.Columns(7).HeaderText = "jabatan"

        datagridstok.Columns(8).HeaderText = "Tgl Masuk"
    End Sub
    Dim jenis = ""
    Public Sub isicombo()
        Dim strSQL As String = "Select distinct ket From jabatan order by ket"
        Dim DA As New MySqlDataAdapter(strSQL, con)
        Dim DS As New DataSet

        DA.Fill(DS, "Codes")
        Dim dt As New DataTable
        dt = DS.Tables("codes")
        Dim dr As DataRow
        '
        ' Populate the Combobox with the Descriptions.
        '
        ComboBox2.Items.Clear()

        For Each dr In dt.Rows()
            ComboBox2.Items.Add(dr("ket"))

        Next
        ComboBox2.SelectedIndex = -1

    End Sub
    Private Sub datacustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        open()
        isigrid()
        isicombo()
        kosong()
        kosong()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        kosong()

        Button1.Visible = True
        Button6.Visible = False

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Button1.Visible = False
        Button6.Visible = True

        txtkode.DataBindings.Add("text", stokbind, "kode")
        txtkode.DataBindings.Clear()
        txtnama.DataBindings.Add("text", stokbind, "nama")
        txtnama.DataBindings.Clear()
        txtalamat.DataBindings.Add("text", stokbind, "alamat")
        txtalamat.DataBindings.Clear()
        ComboBox2.DataBindings.Add("text", stokbind, "jabatan")
        ComboBox2.DataBindings.Clear()
        txtnoktp.DataBindings.Add("text", stokbind, "noktp")
        txtnoktp.DataBindings.Clear()

        txttelp.DataBindings.Add("text", stokbind, "telp")
        txttelp.DataBindings.Clear()





    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If txtkode.Text = "" Then
            MsgBox("Pilih Customer !", MsgBoxStyle.Critical, "Error")
            Return
        End If
        Dim jam = Format(Now, "yyyy-MM-dd HH:mm:ss")
        Dim konfirm As String = MsgBox("Yakin Akan Menghapus Data " & txtnama.Text & " Alamat " & txtalamat.Text & "?", MsgBoxStyle.YesNo, "Konfirmasi")


        If konfirm = vbYes Then
            Try
                open()

                Dim strsql As String = ""
                strsql = "delete from pegawai where kode='" & txtkode.Text & "' "

                Dim mycmd As MySqlCommand = New MySqlCommand(strsql, con)
                mycmd.CommandType = CommandType.Text
                Dim a As MySqlDataReader = mycmd.ExecuteReader()
                a.Close()



                kosong()
                isigrid()

                MsgBox("Data Berhasil Dihapus !")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

    End Sub
    Dim aktif
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim bln, tgl, thn
        tgl = Date.Parse(DateTimePicker1.Text).Day.ToString
        bln = Date.Parse(DateTimePicker1.Text).Month.ToString
        thn = Date.Parse(DateTimePicker1.Text).Year.ToString
        Dim mydate = thn + "-" + bln + "-" + tgl

        Dim user As String
        Dim pwd As String
        Dim pname As String
        Dim cmd As MySqlCommand
        Dim jam = Format(Now, "yyyy-MM-dd HH:mm:ss")

        pname = ""
        open()
        pname = "select * from pegawai where kode='" & txtkode.Text & "'  "

        cmd = New MySqlCommand(pname, con)
        Dim current As String
        current = CStr(cmd.ExecuteScalar)
        If current <> "" Then
            MsgBox("Data ini sudah ada !")

        Else
            Dim strsql5 As String = "insert into pegawai values('" & txtkode.Text & "','" & txtnama.Text & "','" & txtnoktp.Text & "','" & txtalamat.Text & "','" & txttelp.Text & "','','','" & ComboBox2.Text & "','" & mydate & "','')"
            Dim mycmd5 As MySqlCommand = New MySqlCommand(strsql5, con)
            mycmd5.CommandType = CommandType.Text
            Dim a5 As MySqlDataReader = mycmd5.ExecuteReader()
            a5.Close()


        End If
        MsgBox("Data Berhasil Disimpan !")
        isigrid()
        isicombo()

        kosong()

        Button1.Visible = True
        Button6.Visible = False

    End Sub

    Private Sub cmbpbf_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtisi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtisi.TextChanged
        isigrid2()

    End Sub

    Private Sub datagridstok_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles datagridstok.CellContentClick

    End Sub

    Private Sub datagridstok_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles datagridstok.CellEnter
        txtkode.DataBindings.Add("text", stokbind, "kode")
        txtkode.DataBindings.Clear()
        txtnama.DataBindings.Add("text", stokbind, "nama")
        txtnama.DataBindings.Clear()
        txtalamat.DataBindings.Add("text", stokbind, "alamat")
        txtalamat.DataBindings.Clear()
        ComboBox2.DataBindings.Add("text", stokbind, "jabatan")
        ComboBox2.DataBindings.Clear()
        txtnoktp.DataBindings.Add("text", stokbind, "noktp")
        txtnoktp.DataBindings.Clear()

        txttelp.DataBindings.Add("text", stokbind, "telp")
        txttelp.DataBindings.Clear()


    End Sub




    Private Sub cmbrs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            open()

            Dim strsql5 As String = "update pegawai set nama='" & txtnama.Text & "',alamat='" & txtalamat.Text & "',jabatan='" & ComboBox2.Text & "',telp='" & txttelp.Text & "',noktp='" & txtnoktp.Text & "',telp='" & txttelp.Text & "' where kode='" & txtkode.Text & "'"
            Dim mycmd5 As MySqlCommand = New MySqlCommand(strsql5, con)
            mycmd5.CommandType = CommandType.Text
            Dim a5 As MySqlDataReader = mycmd5.ExecuteReader()
            a5.Close()
            MsgBox("Edit Berhasil")
            isigrid()
            isicombo()

            Button1.Visible = True
            Button6.Visible = False
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Button1.Visible = True
        Button6.Visible = False

    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox1.Visible = True
        TextBox1.DataBindings.Add("text", stokbind, "kode")
        TextBox1.DataBindings.Clear()
        TextBox1.Visible = False

        Process.Start("Http://localhost:81/pulungs/index3.php?id='" & TextBox1.Text & "'")
    End Sub
End Class