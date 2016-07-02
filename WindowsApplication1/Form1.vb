Public Class Form1
    Dim Sql As New CURD_SQL

    ' SHORTEN ERROR CHECKS
    Private Function NotEmpty(val As String) As Boolean
        If Not String.IsNullOrEmpty(val) Then Return True Else Return False
    End Function

    Private Sub cmdQuery_Click(sender As System.Object, e As System.EventArgs) Handles cmdQuery.Click
        ' EXECUTE QUERY/COMMAND
        Sql.ExecQuery(txtQuery.Text)

        ' REPORT & ABORT ON ERRORS
        If NotEmpty(Sql.Exception) Then MsgBox(Sql.Exception) : Exit Sub

        ' SEND QUERY RESULTS TO DATAGRIDVIEW
        DGVData.DataSource = Sql.DBDT
    End Sub





    'CURD Operations

    Public Sub CreateUser()
        ' VALIDATE REQUIRED FIELDS
        If Len(txtUser.Text) < 6 Then MsgBox("Username too short.") : Exit Sub
        If Len(txtPass.Text) < 8 Then MsgBox("Password too short.") : Exit Sub

        ' SET INSERT PARAMS
        Sql.AddParam("@user", txtUser.Text)
        Sql.AddParam("@pass", txtPass.Text)
        Sql.AddParam("@email", txtEmail.Text)
        Sql.AddParam("@website", txtWebsite.Text)
        Sql.AddParam("@active", CInt(cbActive.Checked))
        Sql.AddParam("@admin", CInt(cbAdmin.Checked))

        ' EXECUTE INSERT COMMAND
        Sql.ExecQuery("INSERT INTO members (username,password,email,website,active,admin) " &
                      "VALUES(@user,@pass,@email,@website,@active,@admin)")


        ' REPORT ERRORS
        If NotEmpty(Sql.Exception) Then MsgBox(Sql.Exception)
    End Sub

    Public Sub ReadUsers()
        Sql.ExecQuery("SELECT username FROM members")

        ' REPORT & ABORT ON ERRORS
        If NotEmpty(Sql.Exception) Then MsgBox(Sql.Exception) : Exit Sub

        ' POPULATE COMBOBOX WITH USERNAMES
        cbxUsers.Items.Clear()
        For Each r As DataRow In Sql.DBDT.Rows
            cbxUsers.Items.Add(r("username"))
        Next

        ' SELECT FIRST USER IF USERS ARE FOUND
        If cbxUsers.Items.Count > 0 Then cbxUsers.SelectedIndex = 0
    End Sub





    Private Sub ReadUserByName(sender As System.Object, e As System.EventArgs) Handles cmdSave.Click
        ' ADD QUERY PARAMS
        Sql.AddParam("@user", txtUser.Text)

        ' QUERY FOR USER
        Sql.ExecQuery("SELECT username FROM members WHERE username = @user ")

        ' REPORT & ABORT IF USER EXISTS
        If Sql.DBDT.Rows.Count > 0 Then MsgBox("User already exists!") : Exit Sub

        ' CREATE NEW USER
        CreateUser()

        ' CLEAN UP FIELDS
        txtUser.Clear()
        txtPass.Clear()
        txtEmail.Clear()
        txtWebsite.Clear()
        cbActive.Checked = True
        cbAdmin.Checked = False
    End Sub





End Class
