Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class CURD_SQL
    'Connection amd Command
    Private DBCon As New SqlConnection("Server=MINECRAFT\SQLEXPRESS;Database=SQLApps;User=sa;Pwd=Password1;")
    Private DBCmd As SqlCommand

    ' DB DATA
    Public DBDA As SqlDataAdapter
    Public DBDT As DataTable

    ' QUERY STATISTICS 
    Public RecordCount As Integer
    Public Exception As String

    Public Sub New()
    End Sub

    ' ALLOW CONNECTION STRING OVERRIDE
    Public Sub New(ConnectionString As String)
        DBCon = New SqlConnection(ConnectionString)
    End Sub

    ' EXECUTE QUERY SUB
    Public Sub ExecQuery(Query As String)
        ' RESET QUERY STATISTICS 
        RecordCount = 0
        Exception = ""

        Try
            DBCon.Open()

            ' CREATE DB COMMAND
            DBCmd = New SqlCommand(Query, DBCon)

            ' LOAD PARAMS INTO DB COMMAND
            Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

            ' CLEAR PARAM LIST
            Params.Clear()

            ' EXECUTE COMMAND AND FILL DATASET
            DBDT = New DataTable
            DBDA = New SqlDataAdapter(DBCmd)
            RecordCount = DBDA.Fill(DBDT)
        Catch ex As Exception
            ' CAPTURE ERRORS
            Exception = "ExecQuery Error: " & vbNewLine & ex.Message
        End Try

        ' CLOSE CONNECTION
        If DBCon.State = ConnectionState.Open Then DBCon.Close()
    End Sub





    ' ADD PARAM SUB
    ' QUERY PARAMETERS
    Public Params As New List(Of SqlParameter)

    Public Sub AddParam(Name As String, Value As Object)
        Dim NewParam As New SqlParameter(Name, Value)
        Params.Add(NewParam)
    End Sub

End Class
