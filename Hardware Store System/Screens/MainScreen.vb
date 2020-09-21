Imports System.ComponentModel
Imports System.Data.OleDb

Public Class MainScreen
    ReadOnly db As OleDatabaseConnector = Constants.DBConnection
    Private AdminList As ArrayList = Nothing
    Private SelectedAdmin As DataGridViewRow

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        BunifuFormDock1.SubscribeControlToDragEvents(Panel1)
        BunifuFormDock1.SubscribeControlToDragEvents(Panel2)
        PopulateAdminDataGrid()
    End Sub

    Private Sub PopulateAdminDataGrid()

        Dim Worker As New BackgroundWorker()
        AddHandler Worker.DoWork, AddressOf LoadEntityOnDoWorkHandler
        AddHandler Worker.RunWorkerCompleted, AddressOf LoadAdminEntitiesWorkCompleted
        Worker.RunWorkerAsync(db.SqlCommand("Select * From Admins"))

    End Sub

    Private Sub LoadEntityOnDoWorkHandler(sender As Object, e As DoWorkEventArgs)
        Dim results As New ArrayList()

        Dim cmd As OleDbCommand = CType(e.Argument, OleDbCommand)

        Dim task As OleSqlGetTask = db.RunGetCommand(cmd)

        If task.IsSuccessful() Then
            While task.GetResult().Read()
                Dim admin As Entity = New EntityBuilder().UseOleReader(task.GetResult(), False, False).Build()
                results.Add(admin)
                Console.WriteLine("added item")
            End While
            task.GetResult().Close()
        Else
            Console.WriteLine(task.GetException().Message)
        End If

        e.Result = results

    End Sub

    Private Sub LoadAdminEntitiesWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs)

        Dim results As ArrayList = CType(e.Result, ArrayList)
        AdminList = results
        For i As Integer = 0 To results.Count - 1
            Dim Item As Entity = results(i)
            Dim row As String() = {Item("ID"), Item("fullname"), Item("email"), Item("password")}
            AdminsDataGridView.Rows.Add(row)
        Next

    End Sub

    Private Sub AdminsDataGridView_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles AdminsDataGridView.CellDoubleClick
        SelectedAdmin = AdminsDataGridView.Rows(e.RowIndex)
        'display in form field
        fullname.Text = SelectedAdmin.Cells("Name").Value
        adminEmail.Text = SelectedAdmin.Cells("Email").Value
        password.Text = SelectedAdmin.Cells("Password").Value
    End Sub

    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(0)
    End Sub

    Private Sub BunifuButton3_Click(sender As Object, e As EventArgs) Handles BunifuButton3.Click
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(1)
    End Sub

    Private Sub BunifuButton5_Click(sender As Object, e As EventArgs) Handles BunifuButton5.Click
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(2)
    End Sub

    Private Sub BunifuButton6_Click(sender As Object, e As EventArgs) Handles BunifuButton6.Click
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(3)
    End Sub

    Private Sub BunifuButton4_Click(sender As Object, e As EventArgs) Handles BunifuButton4.Click
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(4)
    End Sub

    Private Sub BunifuButton8_Click(sender As Object, e As EventArgs) Handles BunifuButton8.Click
        Application.[Exit]()
    End Sub

    Private Sub BunifuButton7_Click(sender As Object, e As EventArgs) Handles BunifuButton7.Click
        Dim oForm1 As New AuthScreen()

        Me.Hide()
        oForm1.ShowDialog()

        Me.Close()
    End Sub

    Private Sub BunifuButton20_Click(sender As Object, e As EventArgs) Handles BunifuButton20.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            BunifuPictureBox2.ImageLocation = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub BunifuPictureBox1_Click(sender As Object, e As EventArgs) Handles BunifuPictureBox1.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            BunifuPictureBox1.ImageLocation = OpenFileDialog1.FileName
        End If
    End Sub

End Class