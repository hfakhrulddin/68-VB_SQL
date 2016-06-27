Public Class Form1
    Dim Sql As New CURD_SQL

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







End Class
