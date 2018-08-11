Public Class Form5

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        TextBox3.Text = Val(TextBox1.Text) * Val(TextBox2.Text) / 100
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form3.Label12.Text = TextBox4.Text
        Form3.txtDiskon.Text = TextBox2.Text
        Form3.TextBox3.Text = TextBox3.Text
        If Me.Text = "+" Then
            Form3.txtTotal2.Text = Val(TextBox1.Text) + Val(TextBox3.Text)
        Else
            Form3.txtTotal2.Text = Val(TextBox1.Text) - Val(TextBox3.Text)
        End If
    End Sub

End Class