Imports System.Data.OleDb

Public Class Complaints
    ReadOnly db As OleDatabaseConnector = Constants.DBConnection
    Dim oForm As New MainScreen()

    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        BunifuPages1.SetPage(0)
        Me.CenterToScreen()
    End Sub

    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        BunifuPages1.SetPage(1)
    End Sub


    Private Sub BunifuButton4_Click(sender As Object, e As EventArgs) Handles BunifuButton4.Click

        If email.Text.Trim().Length = 0 Then
            MessageBox.Show("email field is required.", "Required Field!")
            Return

        ElseIf password.Text.Trim().Length = 0 Then
            MessageBox.Show("password field is required.", "Required Field!")
            Return
        End If

        Dim cmd As OleDbCommand = db.SqlCommand("Select * From Admins Where email = ? And password = ?")

        With cmd.Parameters
            .AddWithValue("@param1", email.Text.Trim())
            .AddWithValue("@param2", password.Text.Trim())
        End With

        Dim task As OleSqlGetTask = db.RunGetCommand(cmd)

        If task.IsSuccessful() And task.GetResult().Read() Then
            Me.Hide()
            oForm.ShowDialog()
            Me.Close()
        Else
            MessageBox.Show("Invalid Email or Password", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub Register_Click(sender As Object, e As EventArgs) Handles register.Click
        If regFullname.Text.Trim().Length = 0 Then
            MessageBox.Show("emplyees's id field is required.", "Required Field!")
            Return
        ElseIf regEmail.Text.Trim().Length = 0 Then
            MessageBox.Show("password field is required.", "Required Field!")
            Return
        ElseIf regPassword.Text.Trim().Length < 7 Then
            MessageBox.Show("password is too short.", "Required Field!")
            Return
        End If

        Dim cmd As OleDbCommand = db.SqlCommand("Insert Into Admins([fullname], [email], [password]) values(?, ?, ?)")

        With cmd.Parameters
            .AddWithValue("@param1", regFullname.Text.Trim())
            .AddWithValue("@param2", regEmail.Text.Trim())
            .AddWithValue("@param3", regPassword.Text.Trim())
        End With

        Dim task As OleSqlSetTask = db.RunSetCommand(cmd)

        If task.IsSuccessful() Then
            MessageBox.Show("your account has been created succesfully.", "success")
            BunifuPages1.SetPage(0)
        Else
            If TypeOf task.GetException() Is OleDbException Then
                If task.GetException().Message.Contains("duplicate") Then
                    MessageBox.Show("a user already exists with the same username!", "sorry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show(task.GetException().Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show(task.GetException().Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
End Class
