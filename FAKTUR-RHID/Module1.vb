Imports System.Data.OleDb
Module Module1
    Public konek As OleDbConnection
    Public cmd As OleDbCommand
    Public da As OleDbDataAdapter
    Public dr As OleDbDataReader
    Public dataset As DataSet
    Public tabel As DataTable
    Public Str As String
    Sub buka()
        Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\Database1.mdb"
        konek = New OleDbConnection(Str)
        If konek.State = ConnectionState.Closed Then
            konek.Open()
        End If
    End Sub
   


End Module

