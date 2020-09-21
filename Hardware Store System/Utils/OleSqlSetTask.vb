
Imports System.Data.OleDb

Public Class OleSqlSetTask
    Inherits Operation(Of Integer)

    Private Command As OleDbCommand

    Public Sub New(sqlCommand As OleDbCommand)
        Command = sqlCommand
        Run()
    End Sub

    Overrides Function DoWork() As Integer
        If Command.Connection.State = ConnectionState.Closed Then
            Command.Connection.Open()
        End If
        Dim Result As Integer = Command.ExecuteNonQuery()
        Command.Connection.Close()
        Return Result
    End Function

End Class