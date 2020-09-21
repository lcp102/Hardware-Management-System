Imports System.ComponentModel
Imports System.Data.OleDb

Public Class MainScreen
    ReadOnly db As OleDatabaseConnector = Constants.DBConnection
    Private AdminList As ArrayList = Nothing
    Private ClientList As ArrayList = Nothing
    Private CarList As ArrayList = Nothing
    Private RetalList As ArrayList = Nothing

    Private SelectedAdmin As DataGridViewRow
    Private SelectedClient As DataGridViewRow
    Private SelectedCar As DataGridViewRow
    Private SelectedRental As DataGridViewRow

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

        AdminList = CType(e.Result, ArrayList)
        SetAdminDatas(AdminList)

    End Sub

    Private Sub AdminsDataGridView_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles AdminsDataGridView.CellClick
        SelectedAdmin = AdminsDataGridView.Rows(e.RowIndex)
        'display in form field
        fullname.Text = SelectedAdmin.Cells("AdminName").Value
        adminEmail.Text = SelectedAdmin.Cells("Email").Value
        password.Text = SelectedAdmin.Cells("AdminPass").Value
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

    Private Sub BunifuButton17_Click(sender As Object, e As EventArgs) Handles BunifuButton17.Click
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(5)
    End Sub

    Private Sub BunifuButton18_Click(sender As Object, e As EventArgs)
        indicator.Top = (CType(sender, Control)).Top
        BunifuPages1.SetPage(6)
    End Sub

    Private Sub BunifuButton8_Click(sender As Object, e As EventArgs) Handles BunifuButton8.Click
        Application.[Exit]()
    End Sub

    Private Sub BunifuButton7_Click(sender As Object, e As EventArgs) Handles BunifuButton7.Click
        Dim oForm1 As New Complaints()

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

    Private Sub BunifuButton10_Click(sender As Object, e As EventArgs) Handles BunifuButton10.Click
        'delete Admin
        If IsNothing(SelectedAdmin) Then
            MessageBox.Show("Please Select An Admin to delete", "No Admin Selected")
            Return
        End If

        If fullname.Text.Trim().Length = 0 Then
            MessageBox.Show("admins's fullname field is required.", "Required Field!")
            Return
        ElseIf adminEmail.Text.Trim().Length = 0 Then
            MessageBox.Show("email field is required.", "Required Field!")
            Return
        ElseIf password.Text.Trim().Length = 0 Then
            MessageBox.Show("password field is required.", "Required Field!")
            Return
        ElseIf (MessageBox.Show("Are you sure you want to delete this admin?", "Confirm Delete", MessageBoxButtons.YesNo) = DialogResult.Yes) Then

            Dim cmd As OleDbCommand = db.SqlCommand("Delete From Admins Where ID = " & SelectedAdmin.Cells("AdminID").Value)

            Dim task As OleSqlSetTask = db.RunSetCommand(cmd)

            If Not task.IsSuccessful() Then
                MessageBox.Show(task.GetException().Message, "Failed Operation!")
            Else
                AdminsDataGridView.Rows.Clear()
                PopulateAdminDataGrid()
            End If

        End If
    End Sub

    Private Sub BunifuButton9_Click(sender As Object, e As EventArgs) Handles BunifuButton9.Click
        'update Admin
        If IsNothing(SelectedAdmin) Then
            MessageBox.Show("Please Select An Admin to update", "No Admin Selected")
            Return
        End If


        If fullname.Text.Trim().Length = 0 Then
            MessageBox.Show("admins's fullname field is required.", "Required Field!")
            Return
        ElseIf adminEmail.Text.Trim().Length = 0 Then
            MessageBox.Show("email field is required.", "Required Field!")
            Return
        ElseIf password.Text.Trim().Length = 0 Then
            MessageBox.Show("password field is required.", "Required Field!")
            Return
        ElseIf (MessageBox.Show("Are you sure you want to update this admin?", "Confirm Update", MessageBoxButtons.YesNo) = DialogResult.Yes) Then

            Dim cmd As OleDbCommand = db.SqlCommand("Update Admins Set [fullname] = ?, [email] = ?, [password] = ? Where ID = ?")

            Console.WriteLine("admin updating " & SelectedAdmin.Cells("AdminID").Value)

            With cmd.Parameters
                .AddWithValue("@param1", fullname.Text.Trim())
                .AddWithValue("@param2", adminEmail.Text.Trim())
                .AddWithValue("@param3", password.Text.Trim())
                .AddWithValue("@param4", SelectedAdmin.Cells("AdminID").Value)
            End With

            Dim task As OleSqlSetTask = db.RunSetCommand(cmd)

            If Not task.IsSuccessful() Then
                MessageBox.Show(task.GetException().Message, "Failed Operation!")
            Else
                AdminsDataGridView.Rows.Clear()
                PopulateAdminDataGrid()
            End If

        End If


    End Sub

    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        'insert Admin
        If fullname.Text.Trim().Length = 0 Then
            MessageBox.Show("admins's fullname field is required.", "Required Field!")
            Return
        ElseIf adminEmail.Text.Trim().Length = 0 Then
            MessageBox.Show("email field is required.", "Required Field!")
            Return
        ElseIf password.Text.Trim().Length = 0 Then
            MessageBox.Show("password field is required.", "Required Field!")
            Return
        ElseIf (MessageBox.Show("Are you sure you want to update this admin?", "Confirm Insert", MessageBoxButtons.YesNo) = DialogResult.Yes) Then

            Dim cmd As OleDbCommand = db.SqlCommand("Insert Into Admins([fullname], [email], [password]) Values(?, ?, ?)")

            With cmd.Parameters
                .AddWithValue("@param1", fullname.Text.Trim())
                .AddWithValue("@param2", adminEmail.Text.Trim())
                .AddWithValue("@param3", password.Text.Trim())
            End With

            Dim task As OleSqlSetTask = db.RunSetCommand(cmd)

            If Not task.IsSuccessful() Then
                MessageBox.Show(task.GetException().Message, "Failed Operation!")
            Else
                AdminsDataGridView.Rows.Clear()
                PopulateAdminDataGrid()
            End If

        End If

    End Sub

    Private Sub BunifuTextBox4_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox4.TextChanged
        Dim searchText As String = BunifuTextBox4.Text.Trim.ToLower()
        If searchText.Length = 0 Then
            SetAdminDatas(AdminList)
            Return
        End If

        Dim results As New ArrayList()

        For Each item As Entity In AdminList
            If item("ID").ToString().ToLower().Contains(searchText) Or item("email").ToString().ToLower().Contains(searchText) Or item("fullname").ToString().ToLower().Contains(searchText) Then
                results.Add(item)
            End If
        Next
        SetAdminDatas(results)
    End Sub

    Private Sub SetAdminDatas(adminList As ArrayList)
        AdminsDataGridView.Rows.Clear()
        For i As Integer = 0 To adminList.Count - 1
            Dim Item As Entity = adminList(i)
            Dim row As String() = {Item("ID"), Item("fullname"), Item("email"), Item("password")}
            AdminsDataGridView.Rows.Add(row)
        Next
    End Sub

    Public Function validateClientForm() As Boolean

        If clientFullname.Text.Trim().Length = 0 Then
            MessageBox.Show("clients's fullname field is required.", "Required Field!")
            Return False
        ElseIf clientEmail.Text.Trim().Length = 0 Then
            MessageBox.Show("clients field is required.", "Required Field!")
            Return False
        ElseIf clientNextKin.Text.Trim().Length = 0 Then
            MessageBox.Show("clients field is required.", "Required Field!")
            Return False
        ElseIf clientPhone.Text.Trim().Length = 0 Then
            MessageBox.Show("clients field is required.", "Required Field!")
            Return False
        ElseIf password.Text.Trim().Length = 0 Then
            MessageBox.Show("password field is required.", "Required Field!")
            Return False
            'check if date in  the future
            'check only one radio button is selected
        ElseIf (MessageBox.Show("Are you sure you want to add this client?", "Confirm Insert", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
            Return True
        Else
            Return False
        End If

    End Function


    Private Sub PopulateClientDataGrid()

        Dim Worker As New BackgroundWorker()
        AddHandler Worker.DoWork, AddressOf LoadEntityOnDoWorkHandler
        AddHandler Worker.RunWorkerCompleted, AddressOf LoadClientEntitiesWorkCompleted
        Worker.RunWorkerAsync(db.SqlCommand("Select * From Clients"))
    End Sub

    Private Sub LoadClientEntitiesWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        ClientList = CType(e.Result, ArrayList)
        SetClientDatas(ClientList)
    End Sub

    Private Sub SetClientDatas(clientList As ArrayList)
        Throw New NotImplementedException()
    End Sub

    Private Sub BunifuButton13_Click(sender As Object, e As EventArgs) Handles BunifuButton13.Click
        If validateClientForm() Then
            Dim cmd As OleDbCommand = db.SqlCommand("Insert Into Clients([fullname], [email], [phone], [gender], [dob], [next_kin]) Values(?, ?, ?, ?, ?, ?)")

            With cmd.Parameters
                .AddWithValue("@param1", clientFullname.Text.Trim())
                .AddWithValue("@param2", clientEmail.Text.Trim())
                .AddWithValue("@param3", clientPhone.Text.Trim())
                .AddWithValue("@param4", clientFullname.Text.Trim()) 'change to combo box and do .selectedValue
                .AddWithValue("@param5", clientDob.Value)
                .AddWithValue("@param6", clientNextKin.Text.Trim())
            End With

            Dim task As OleSqlSetTask = db.RunSetCommand(cmd)

            If Not task.IsSuccessful() Then
                MessageBox.Show(task.GetException().Message, "Failed Operation!")
            Else
                AdminsDataGridView.Rows.Clear()
                PopulateClientDataGrid()
            End If
        End If
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub BunifuTextBox2_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox2.TextChanged

    End Sub

    Private Sub BunifuButton4_Click(sender As Object, e As EventArgs)
        Complains.ShowDialog()
    End Sub

    Private Sub BunifuButton4_Click_1(sender As Object, e As EventArgs) Handles BunifuButton4.Click
        Complains.ShowDialog()
    End Sub

    Private Sub BunifuDropdown1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown1.SelectedIndexChanged

    End Sub

    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Click

    End Sub

    Private Sub BunifuButton25_Click(sender As Object, e As EventArgs) Handles BunifuButton25.Click
        Receipt.ShowDialog()
    End Sub
End Class