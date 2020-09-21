Imports System.Data.OleDb

Public Class OleDatabaseConnector

    Private mConnection As OleDbConnection
    Private mConnectionString As String

    Public Sub New(connectionString As String)
        mConnectionString = connectionString
        mConnection = New OleDbConnection(connectionString)
    End Sub

    Public Sub Open()
        mConnection.Open()
    End Sub

    Public Sub Close()
        mConnection.Close()
    End Sub

    Public Function IsOpen() As Boolean
        Return mConnection.State = ConnectionState.Open
    End Function

    Public Function SqlCommand(command As String) As OleDbCommand
        Return New OleDbCommand(command, mConnection)
    End Function

    Public Function RunGetCommand(sqlCommand As OleDbCommand) As OleSqlGetTask
        Return New OleSqlGetTask(sqlCommand)
    End Function

    Public Function RunSetCommand(sqlCommand As OleDbCommand) As OleSqlSetTask
        Return New OleSqlSetTask(sqlCommand)
    End Function

End Class
