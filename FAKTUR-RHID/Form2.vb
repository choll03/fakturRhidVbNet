Imports System.Data
Imports System.Data.OleDb

Public Class form2

    Private Sub form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        buka()
       

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ProgressBar1.Visible = True
        Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 20
        If ProgressBar1.Value = 100 Then
            Timer1.Dispose()
            If konek.State <> ConnectionState.Closed Then
                konek.Close()
                konek.Open()
                cmd = New OleDbCommand("Select * From Login WHERE Username ='" & TextBox1.Text & "' and Password='" & TextBox2.Text & "'", konek)
                dr = cmd.ExecuteReader
                If (dr.Read()) Then
                    Form3.Show()
                    Me.Hide()
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox1.Focus()
                    ProgressBar1.Visible = False
                    ProgressBar1.Value = 0
                Else
                    MsgBox("Username dan Password anda Salah !", MsgBoxStyle.OkOnly, "Login Gagal")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox1.Focus()
                    ProgressBar1.Value = 0
                    ProgressBar1.Visible = False
                End If
            End If
        End If
    End Sub

End Class