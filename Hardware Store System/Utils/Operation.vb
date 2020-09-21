Public MustInherit Class Operation(Of T)

    Private mResult As T
    Private mException As Exception

    Public MustOverride Function DoWork() As T

    Public Sub Run()
        Try
            mResult = DoWork()
        Catch ex As Exception
            mException = ex
        End Try
    End Sub

    Public Function IsSuccessful() As Boolean
        Return IsNothing(mException)
    End Function

    Public Function GetException() As Exception
        Return mException
    End Function

    Public Function GetResult() As T
        Return mResult
    End Function



End Class
