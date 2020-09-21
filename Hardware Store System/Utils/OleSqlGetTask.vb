Imports System.Data.OleDb

Public Class OleSqlGetTask
    Inherits Operation(Of OleDbDataReader)

    Private Command As OleDbCommand

    Public Sub New(sqlCommand As OleDbCommand)
        Command = sqlCommand
        Run()
    End Sub

    Overrides Function DoWork() As OleDbDataReader
        If Command.Connection.State = ConnectionState.Closed Then
            Command.Connection.Open()
        End If
        Dim Result As OleDbDataReader = Command.ExecuteReader()
        Return Result
    End Function

End Class
