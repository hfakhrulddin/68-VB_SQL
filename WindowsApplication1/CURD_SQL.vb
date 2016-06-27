Imports System.Data.Sql
Imports System.Data.SqlClient




Public Class CURD_SQL
    'Connection amd Command
    Private DBCon As New SqlConnection("Server=MINECRAFT\SQLEXPRESS;Database=SQLApps;User=sa;Pwd=Password1;")
    Private DBCmd As SqlCommand

    ' DB DATA
    Public DBDA As SqlDataAdapter
    Public DBDT As DataTable

End Class
