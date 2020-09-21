Public Class Functions

    Public Shared Sub CenterInParent(ByVal child As Form, Optional ByVal parent As Form = Nothing)
        Dim x As Integer
        Dim y As Integer
        Dim r As Rectangle

        If IsNothing(parent) Then
            r = Screen.FromPoint(child.Location).WorkingArea
        Else
            r = parent.RectangleToScreen(parent.ClientRectangle)
        End If

        x = r.Left - (r.Width - child.Width) \ 2
        y = r.Top - (r.Height - child.Height) \ 2

        child.Location = New Point(x, y)
    End Sub


    Public Shared Sub SetFormFragment(FragmentContainer As Control, FormFragment As Form)

        FormFragment.TopLevel = False
        FormFragment.FormBorderStyle = FormBorderStyle.None
        FormFragment.Dock = DockStyle.Fill
        FragmentContainer.Controls.Add(FormFragment)
        FragmentContainer.Tag = FormFragment
        FormFragment.BringToFront()
        FormFragment.Show()

    End Sub


End Class
