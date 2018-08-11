Imports System.Data
Imports System.Data.OleDb

Public Class Form3

    Dim nomor As String
    Dim total As Double
    Dim total2 As Double
    Dim total3 As Double
    Dim hargasatuan As Double
    Dim harga As Double
    Dim diskon As Long
    Dim powder As Long
    Sub tampilkan()
        If konek.State <> ConnectionState.Closed Then
            konek.Close()
            konek.Open()
            cmd = New OleDbCommand
            da = New OleDbDataAdapter
            dataset = New DataSet
            cmd.Connection = konek
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "select * from tblBeli"
            da.SelectCommand = cmd
            dataset.Tables.Clear()
            da.Fill(dataset, "tblBeli")
            DataGridView1.DataSource = dataset.Tables("tblBeli")
            DataGridView1.Refresh()

        End If
    End Sub

    Private Sub faktur()
        Dim urutan As String
        Dim hitung As Long
        If konek.State <> ConnectionState.Closed Then
            konek.Close()
            konek.Open()
            cmd = New OleDbCommand("Select * From tblCetak WHERE NoFaktur in " & "(select max(NoFaktur)from tblCetak)", konek)
            dr = cmd.ExecuteReader
            dr.Read()
            If Not dr.HasRows Then
                urutan = "0" & Format(Date.Now, "MM") & "001"
                TextBox1.Text = urutan
            ElseIf Microsoft.VisualBasic.Left(dr!NoFaktur, 3) <> "0" & Format(Date.Now, "MM") Then
                urutan = "0" & Format(Date.Now, "MM") & "001"
            Else
                hitung = Microsoft.VisualBasic.Right(dr!NoFaktur, 3) + 1
                urutan = "0" & Format(Date.Now, "MM") & Microsoft.VisualBasic.Right("000" & hitung, 3)
            End If
            TextBox1.Text = urutan

        End If
    End Sub

    Sub atur()

        With DataGridView1.ColumnHeadersDefaultCellStyle
            DataGridView1.Columns(0).Width = 20
            DataGridView1.Columns(0).HeaderText = "No"
            DataGridView1.Columns(1).Width = 130
            DataGridView1.Columns(1).HeaderText = "Nama Barang"
            DataGridView1.Columns(2).Width = 100
            DataGridView1.Columns(3).Width = 50
            DataGridView1.Columns(3).HeaderText = "Jumlah"
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(4).HeaderText = "Harga Satuan"
            DataGridView1.Columns(5).Width = 120
            DataGridView1.Columns(5).HeaderText = "Total Harga"
        End With
    End Sub




    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        buka()
        faktur()
        tampilkan()
        atur()
        auto()
        hitung()

    End Sub

    Sub auto()

        cmd = New OleDbCommand("Select * From tblBeli order by Nomor desc", konek)
        dr = cmd.ExecuteReader
        If dr.Read() Then
            nomor = Microsoft.VisualBasic.Right(dr(0), 2) + 1
        Else
            nomor = "01"
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Then
            MsgBox("Masukan Nomor Faktur!", vbInformation, "Pesan")
            TextBox1.Focus()
        ElseIf txtNamaBarang.Text = "" Then
            MsgBox("Masukan Nomor Nama Barang!", vbInformation, "Pesan")
            txtNamaBarang.Focus()
        ElseIf txtUkuruan1.Text = "" Then
            MsgBox("Masukan Ukuran Barang!", vbInformation, "Pesan")
            txtUkuruan1.Focus()
        ElseIf txtHargaSatuan.Text = "" Then
            MsgBox("Masukan Harga Barang!", vbInformation, "Pesan")
            txtHargaSatuan.Focus()
        ElseIf txtJumlah.Text = "" Then
            MsgBox("Masukan Jumlah Pembelian!", vbInformation, "Pesan")
            txtJumlah.Focus()
        Else
            auto()

            cmd = New OleDbCommand("insert into tblBeli (Nomor,Nama,Ukuran,Jumlah,Harga_Satuan,Total_Harga,Nilai)" & " values('" & nomor & "','" & txtNamaBarang.Text & "','" & txtUkuruan1.Text & "','" & txtJumlah.Text & "','" & Format(Val(txtHargaSatuan.Text), ",###,##,0") & "','" & Format(Val(txtHarga.Text), ",###,##,0") & "','" & harga & "')", konek)
            cmd.ExecuteNonQuery()
            tampilkan()

            hitung()
            clear()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        On Error Resume Next
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox2.Text = DataGridView1.Item(0, i).Value

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        If DataGridView1.Item(0, i).Value = TextBox2.Text Then
            If MsgBox("Anda yakin ingin menghapus nya", MsgBoxStyle.OkCancel, "Pesan") = MsgBoxResult.Ok Then
                cmd = New OleDbCommand("delete from tblBeli where Nomor='" & DataGridView1.Item(0, i).Value & "'", konek)
                cmd.ExecuteNonQuery()
                tampilkan()
                hitung()
            Else
                Me.Show()
            End If
        Else
            MsgBox("Pilihan Data yang di hapus tidak ada", MsgBoxStyle.Information, "Pesan")
        End If
        

    End Sub

    Private Sub txtJumlah_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtJumlah.TextChanged
        harga = Val(txtHargaSatuan.Text) * Val(txtJumlah.Text)

        txtHarga.Text = harga

    End Sub

    Private Sub txtHargaSatuan_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHargaSatuan.TextChanged
        harga = Val(txtHargaSatuan.Text) * Val(txtJumlah.Text)

        txtHarga.Text = harga
    End Sub

    Sub hitung()

        total = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            total = total + Val(DataGridView1.Rows(i).Cells(6).Value)
        Next
        txtTotal.Text = total
        txtTotal.Text = Format(Val(txtTotal.Text), ",###,##,0")
    End Sub
    Sub clear()
        txtNamaBarang.Text = ""
        txtUkuruan1.Text = ""
        txtHargaSatuan.Text = ""
        txtJumlah.Text = ""
        txtHarga.Text = ""
        txtNamaBarang.Focus()
    End Sub

    Private Sub txtDiskon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiskon.TextChanged

        diskon = total * Val(txtDiskon.Text)
        diskon = diskon / 100
        If ComboBox1.Text = "+" Then
            total2 = total + diskon
        ElseIf ComboBox1.Text = "-" Then
            total2 = total - diskon
        Else
            Exit Sub
        End If
        TextBox3.Text = diskon
        txtTotal2.Text = total2
        TextBox3.Text = Format(Val(TextBox3.Text), ",###,##,0")
        txtTotal2.Text = Format(Val(txtTotal2.Text), ",###,##,0")

    End Sub

    Private Sub txtPowder_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowder.TextChanged

        powder = total2 * Val(txtPowder.Text)
        powder = powder / 100
        If ComboBox3.Text = "+" Then
            total3 = total2 + powder
        ElseIf ComboBox3.Text = "-" Then
            total3 = total2 - powder
        Else
            Exit Sub
        End If
        TextBox4.Text = powder
        txtTotal3.Text = total3
        TextBox4.Text = Format(Val(TextBox4.Text), ",###,##,0")
        txtTotal3.Text = Format(Val(txtTotal3.Text), ",###,##,0")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim a As New CrystalReport1
        'Dim b As New Form4
        'b.CrystalReportViewer1.ReportSource = a
        'b.ShowDialog()
        Form4.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "+" Or ComboBox1.Text = "-" Then
            txtDiskon.Text = ""
            TextBox3.Text = ""
            Label12.Text = ""
            txtTotal2.Text = ""
            Dim a As String
            a = InputBox("woodooo", "hi")
            If a = "" Then
                Label12.Text = ""
            Else
                Label12.Text = a
            End If
            Label12.Visible = True
            Label13.Visible = True
            txtDiskon.Visible = True
            TextBox3.Visible = True
            txtTotal2.Visible = True
            ComboBox3.Visible = True
            txtDiskon.Focus()
        Else
            Label12.Visible = False
            Label13.Visible = False
            txtDiskon.Visible = False
            TextBox3.Visible = False
            txtTotal2.Visible = False
            ComboBox3.Visible = False
            txtDiskon.Text = ""
            TextBox3.Text = ""
            Label12.Text = ""
            txtTotal2.Text = ""
            ComboBox3.Text = ".."
        End If
    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "+" Or ComboBox3.Text = "-" Then
            Dim a As String
            a = InputBox("", "")
            If a = "" Then
                Label14.Text = ""
            Else
                Label14.Text = a
            End If
            Label14.Visible = True
            Label15.Visible = True
            txtPowder.Visible = True
            TextBox4.Visible = True
            txtTotal3.Visible = True
            ComboBox4.Visible = True
            txtPowder.Focus()
        Else
            Label14.Visible = False
            Label15.Visible = False
            txtPowder.Visible = False
            TextBox4.Visible = False
            txtTotal3.Visible = False
            ComboBox4.Visible = False
            Label14.Text = ""
            txtPowder.Text = ""
            TextBox4.Text = ""
            txtTotal3.Text = ""
            ComboBox4.Text = ".."
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.Text = "+" Or ComboBox4.Text = "-" Then
            Dim a As String
            a = InputBox("", "")
            If a = "" Then
                Label16.Text = ""
            Else
                Label16.Text = a
            End If
            Label16.Visible = True
            Label17.Visible = True
            TextBox6.Visible = True
            TextBox7.Visible = True
            TextBox8.Visible = True
            ComboBox5.Visible = True
            TextBox6.Focus()
        Else
            Label16.Visible = False
            Label17.Visible = False
            TextBox6.Visible = False
            TextBox7.Visible = False
            TextBox8.Visible = False
            ComboBox5.Visible = False
            Label16.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            ComboBox5.Text = ".."
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "+" Or ComboBox5.Text = "-" Then
            Dim a As String
            a = InputBox("", "")
            If a = "" Then
                Label18.Text = ""
            Else
                Label18.Text = a
            End If
            Label18.Visible = True
            Label19.Visible = True
            TextBox9.Visible = True
            TextBox10.Visible = True
            TextBox11.Visible = True
            TextBox9.Focus()
        Else
            Label18.Visible = False
            Label19.Visible = False
            TextBox9.Visible = False
            TextBox10.Visible = False
            TextBox11.Visible = False
            Label18.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox12.Text = DataGridView1.RowCount


    End Sub
End Class