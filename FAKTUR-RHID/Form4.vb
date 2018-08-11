Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Form4

    Dim ReportViewer1 As Object


    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim myre As New CrystalReport1
        myre.SetParameterValue("MParameter", Form3.txtTotal.Text)
        myre.SetParameterValue("NoFaktur", Form3.TextBox1.Text)
        myre.SetParameterValue("txtPenerima", Form3.txtPenerima.Text)
        myre.SetParameterValue("txtAlamat", Form3.txtAlamat.Text)
        myre.SetParameterValue("txtPO", Form3.txtPO.Text)
        myre.SetParameterValue("txtDiskon", Form3.TextBox3.Text)
        myre.SetParameterValue("txtPowder", Form3.TextBox4.Text)
        myre.SetParameterValue("Total", Form3.txtTotal3.Text)
        myre.SetParameterValue("text5", "( " & Form3.TextBox5.Text & " )")
        CrystalReportViewer1.ReportSource = myre
        'Me.ReportViewer1.refresh()
    End Sub

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

    End Sub
End Class