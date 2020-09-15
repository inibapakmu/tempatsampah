Imports MySql.Data.MySqlClient
Imports System.Globalization
Public Class datacustomer
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
        ComboBox2.Text = ""
        txtnpwp.Text = ""
        txttelp.Text = ""
        txtfax.Text = ""
        txtplafon.Text = "0"
        txttempo.Text = "0"
        txtnamapjk.Text = ""
        txtalamatpjk.Text = ""
        txtnpwpe.Text = ""
        combobox2pjk.Text = ""
        txtkode.Focus()

    End Sub
    Private Sub isigrid()

        Dim mydata As DataSet = getdataset("select kode,nama,alamat,kota,telp,fax,npwp,npwpe,namapjk,alamatpjk,kotapjk,jenis,piutang,plafonhari,sales from customer order by nama ")
        stokbind.DataSource = mydata.Tables("data")
        datagridstok.DataSource = stokbind
        FormatGridWithBothTableAndColumnStyles()
        datagridstok.DataBindings.Clear()

    End Sub
    Private Sub isigrid2()

        Dim mydata As DataSet = getdataset("select kode,nama,alamat,kota,telp,fax,npwp,npwpe,namapjk,alamatpjk,kotapjk,jenis,piutang,plafonhari,sales from customer where " & ComboBox1.Text & " like '%" & txtisi.Text & "%' order by nama ")
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
        datagridstok.Columns(8).Width = 150
        datagridstok.Columns(9).Width = 250
        datagridstok.Columns(10).Width = 150
        datagridstok.Columns(11).Width = 100
        datagridstok.Columns(12).Width = 150
        datagridstok.Columns(13).Width = 150
        datagridstok.Columns(14).Width = 50
        datagridstok.Columns(0).HeaderText = "Kode"
        datagridstok.Columns(1).HeaderText = "Nama"
        datagridstok.Columns(2).HeaderText = "Alamat"
        datagridstok.Columns(3).HeaderText = "Kota"
        datagridstok.Columns(4).HeaderText = "Telp"
        datagridstok.Columns(5).HeaderText = "Fax"
        datagridstok.Columns(6).HeaderText = "NPWP"
        datagridstok.Columns(7).HeaderText = "NPWP-E"
        datagridstok.Columns(8).HeaderText = "Nama Fktr Pjk"
        datagridstok.Columns(9).HeaderText = "Alamat Fktr Pjk"
        datagridstok.Columns(10).HeaderText = "Kota Fktr Pjk"
        datagridstok.Columns(11).HeaderText = "Jenis"
        datagridstok.Columns(12).HeaderText = "Limit Piutang"
        datagridstok.Columns(13).HeaderText = "Plafon Hari"
        datagridstok.Columns(14).HeaderText = "Sales"
    End Sub
    Dim jenis = ""
    Public Sub isicombo()
        Dim strSQL As String = "Select distinct kota From customer order by kota"
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
            ComboBox2.Items.Add(dr("kota"))

        Next
        ComboBox2.SelectedIndex = -1

    End Sub


    Public Sub isicombokel()
        Dim strSQL As String = "Select distinct kelurahan From customer order by kota"
        Dim DA As New MySqlDataAdapter(strSQL, con)
        Dim DS As New DataSet

        DA.Fill(DS, "Codes")
        Dim dt As New DataTable
        dt = DS.Tables("codes")
        Dim dr As DataRow
        '
        ' Populate the Combobox with the Descriptions.
        '
        ComboBox6.Items.Clear()

        For Each dr In dt.Rows()
            ComboBox6.Items.Add(dr("kelurahan"))

        Next
        ComboBox6.SelectedIndex = -1

    End Sub


    Public Sub isicombokec()
        Dim strSQL As String = "Select distinct kecamatan From customer order by kota"
        Dim DA As New MySqlDataAdapter(strSQL, con)
        Dim DS As New DataSet

        DA.Fill(DS, "Codes")
        Dim dt As New DataTable
        dt = DS.Tables("codes")
        Dim dr As DataRow
        '
        ' Populate the Combobox with the Descriptions.
        '
        ComboBox5.Items.Clear()

        For Each dr In dt.Rows()
            ComboBox5.Items.Add(dr("kecamatan"))

        Next
        ComboBox5.SelectedIndex = -1

    End Sub

    Public Sub isicomboprov()
        Dim strSQL As String = "Select distinct provinsi From customer order by kota"
        Dim DA As New MySqlDataAdapter(strSQL, con)
        Dim DS As New DataSet

        DA.Fill(DS, "Codes")
        Dim dt As New DataTable
        dt = DS.Tables("codes")
        Dim dr As DataRow
        '
        ' Populate the Combobox with the Descriptions.
        '
        ComboBox7.Items.Clear()

        For Each dr In dt.Rows()
            ComboBox7.Items.Add(dr("provinsi"))

        Next
        ComboBox7.SelectedIndex = -1

    End Sub


    Public Sub isicombopasar()
        Dim strSQL As String = "Select distinct pasar From customer order by kota"
        Dim DA As New MySqlDataAdapter(strSQL, con)
        Dim DS As New DataSet

        DA.Fill(DS, "Codes")
        Dim dt As New DataTable
        dt = DS.Tables("codes")
        Dim dr As DataRow
        '
        ' Populate the Combobox with the Descriptions.
        '
        ComboBox8.Items.Clear()

        For Each dr In dt.Rows()
            ComboBox8.Items.Add(dr("pasar"))

        Next
        ComboBox8.SelectedIndex = -1

    End Sub


    Public Sub isicombo1()
        Dim strSQL As String = "Select kode,nama From salesman order by nama"
        Dim DA As New MySqlDataAdapter(strSQL, con)
        Dim DS As New DataSet

        DA.Fill(DS, "Codes")
        Dim dt As New DataTable
        dt = DS.Tables("codes")
        Dim dr As DataRow
        '
        ' Populate the Combobox with the Descriptions.
        '
        ComboBox4.Items.Clear()

        For Each dr In dt.Rows()
            ComboBox4.Items.Add(dr("nama") & "#" & dr("kode"))

        Next
        ComboBox4.SelectedIndex = -1
    End Sub
    Public Function CariKodeData(ByVal cString As String) As String
        'FORMAT = .../*** --> Check of a Code
        Dim cAwal As Byte
        Dim lKetemu As Boolean
        Dim cKodeData As String
        cKodeData = ""
        lKetemu = False
        For cAwal = 1 To Len(Trim(cString))
            If Mid(Trim(cString), cAwal, 1) <> "#" Then
                If lKetemu Then
                    cKodeData = cKodeData + Mid(Trim(cString), cAwal, 1)
                End If
            Else
                lKetemu = True
            End If
        Next cAwal
        CariKodeData = Trim(cKodeData)
        Return CariKodeData
    End Function

    Private Sub datacustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        open()
        isigrid()
        isicombo()
        kosong()
        kosong()
        isicombo1()
        isicombokec()
        isicombokel()
        isicombopasar()
        isicomboprov()

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
        ComboBox2.DataBindings.Add("text", stokbind, "kota")
        ComboBox2.DataBindings.Clear()

        ComboBox4.DataBindings.Add("text", stokbind, "sales")
        ComboBox4.DataBindings.Clear()


        TextBox3.DataBindings.Add("text", stokbind, "sales")
        TextBox3.DataBindings.Clear()

        txtnpwp.DataBindings.Add("text", stokbind, "npwp")
        txtnpwp.DataBindings.Clear()

        txttelp.DataBindings.Add("text", stokbind, "telp")
        txttelp.DataBindings.Clear()
        txtfax.DataBindings.Add("text", stokbind, "fax")
        txtfax.DataBindings.Clear()
        txttempo.DataBindings.Add("text", stokbind, "plafonhari")
        txttempo.DataBindings.Clear()
        txtplafon.DataBindings.Add("text", stokbind, "piutang")
        txtplafon.DataBindings.Clear()

        txtnpwpe.DataBindings.Add("text", stokbind, "npwpe")
        txtnpwpe.DataBindings.Clear()

        ComboBox3.DataBindings.Add("text", stokbind, "jenis")
        ComboBox3.DataBindings.Clear()


        txtnamapjk.DataBindings.Add("text", stokbind, "namapjk")
        txtnamapjk.DataBindings.Clear()
        txtalamatpjk.DataBindings.Add("text", stokbind, "alamatpjk")
        txtalamatpjk.DataBindings.Clear()
        combobox2pjk.DataBindings.Add("text", stokbind, "kotapjk")
        combobox2pjk.DataBindings.Clear()

        TextBox1.Visible = True
        TextBox1.DataBindings.Add("text", stokbind, "jenis")
        TextBox1.DataBindings.Clear()
        TextBox1.Visible = False

        
        If TextBox2.Text = "1" Then
            RadioButton1.Checked = True
        ElseIf TextBox2.Text = "" Or TextBox2.Text = "" Then
            RadioButton2.Checked = True

        End If
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
                strsql = "delete from customer where kode='" & txtkode.Text & "' "

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
        If txtkode.Text = "" Then
            MsgBox("Data Masih Ksong")
            Return

        End If
        Dim user As String
        Dim pwd As String
        Dim pname As String
        Dim cmd As MySqlCommand
        Dim jam = Format(Now, "yyyy-MM-dd HH:mm:ss")

        pname = ""
        open()
        pname = "select * from customer where kode='" & txtkode.Text & "'  "

        cmd = New MySqlCommand(pname, con)
        Dim current As String
        current = CStr(cmd.ExecuteScalar)
        If current <> "" Then
            MsgBox("Data ini sudah ada !")
          
        Else
            Dim strsql5 As String = "insert into customer values('" & txtkode.Text & "','" & txtnama.Text & "','" & txtalamat.Text & "','" & txtnamapjk.Text & "','" & txtalamatpjk.Text & "','" & ComboBox2.Text & "','" & combobox2pjk.Text & "','" & txttelp.Text & "','" & txtfax.Text & "','" & txtnpwp.Text & "','" & txtnpwpe.Text & "','" & CDbl(txtplafon.Text) & "','" & txttempo.Text & "','" & aktif & "','" & ComboBox3.Text & "','','','" & TextBox3.Text & "','" & ComboBox5.Text & "','" & ComboBox6.Text & "','" & ComboBox7.Text & "','" & TextBox4.Text & "','" & ComboBox8.Text & "')"
            Dim mycmd5 As MySqlCommand = New MySqlCommand(strsql5, con)
            mycmd5.CommandType = CommandType.Text
            Dim a5 As MySqlDataReader = mycmd5.ExecuteReader()
            a5.Close()

        
        End If
        MsgBox("Data Berhasil Disimpan !")
        isigrid()
        isicombo()
        isicombo1()
        isicombokec()
        isicombokel()
        isicombopasar()
        isicomboprov()

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
        ComboBox2.DataBindings.Add("text", stokbind, "kota")
        ComboBox2.DataBindings.Clear()
        txtnpwp.DataBindings.Add("text", stokbind, "npwp")
        txtnpwp.DataBindings.Clear()

        txttelp.DataBindings.Add("text", stokbind, "telp")
        txttelp.DataBindings.Clear()
        txtfax.DataBindings.Add("text", stokbind, "fax")
        txtfax.DataBindings.Clear()
        txttempo.DataBindings.Add("text", stokbind, "plafonhari")
        txttempo.DataBindings.Clear()
        txtplafon.DataBindings.Add("text", stokbind, "piutang")
        txtplafon.DataBindings.Clear()

        ComboBox3.DataBindings.Add("text", stokbind, "jenis")
        ComboBox3.DataBindings.Clear()

        txtnpwpe.DataBindings.Add("text", stokbind, "npwpe")
        txtnpwpe.DataBindings.Clear()
        txtnamapjk.DataBindings.Add("text", stokbind, "namapjk")
        txtnamapjk.DataBindings.Clear()
        txtalamatpjk.DataBindings.Add("text", stokbind, "alamatpjk")
        txtalamatpjk.DataBindings.Clear()
        combobox2pjk.DataBindings.Add("text", stokbind, "kotapjk")
        combobox2pjk.DataBindings.Clear()

        TextBox1.Visible = True
        TextBox1.DataBindings.Add("text", stokbind, "jenis")
        TextBox1.DataBindings.Clear()
        TextBox1.Visible = False

        
    End Sub

    Private Sub datagridstok_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles datagridstok.CellFormatting

        datagridstok.Columns("piutang").DefaultCellStyle.Format = "#,##"
    End Sub

    Private Sub txtplafon_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtplafon.LostFocus
        If txtplafon.Text = "" Then
            txtplafon.Text = "0"
        End If

        If txtplafon.Text <> "" Then
            Dim amount As Decimal = CType(txtplafon.Text, Decimal) 'say you have entered 1400.10345
            txtplafon.Text = String.Format("{0:n0}", amount)
        End If

    End Sub

    Private Sub txtplafon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtplafon.TextChanged

    End Sub

    Private Sub cmbrs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton1.Checked = True Then
            aktif = "1"
        Else
            aktif = "0"
        End If
    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            aktif = "1"
        Else
            aktif = "0"
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            open()

            Dim strsql5 As String = "update customer set kodepos='" & TextBox4.Text & "',kecamatan='" & ComboBox5.Text & "',kelurahan='" & ComboBox6.Text & "',provinsi='" & ComboBox7.Text & "',pasar='" & ComboBox8.Text & "',sales='" & TextBox3.Text & "',nama='" & txtnama.Text & "',alamat='" & txtalamat.Text & "',kota='" & ComboBox2.Text & "',npwp='" & txtnpwp.Text & "',fax='" & txtfax.Text & "',telp='" & txttelp.Text & "',plafonhari='" & txttempo.Text & "',piutang='" & CDbl(txtplafon.Text) & "',npwpe='" & txtnpwpe.Text & "',namapjk='" & txtnamapjk.Text & "',alamatpjk='" & txtalamatpjk.Text & "',kotapjk='" & combobox2pjk.Text & "' where kode='" & txtkode.Text & "'"
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

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim number As Integer

        Randomize()
        ' The program will generate a number from 0 to 50
        number = Int(Rnd() * 9999999) + 1

        txtkode.Text = number
        txtnama.Focus()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim strValue As String

        strValue = Trim(CariKodeData(Trim(ComboBox4.Text)))

        TextBox3.Visible = True
        TextBox3.Text = strValue.ToString
        TextBox3.Visible = False
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click

    End Sub
End Class