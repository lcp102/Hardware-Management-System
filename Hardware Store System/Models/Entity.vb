Public Class Entity
    Implements IEnumerable

    Private ReadOnly Properties As New Hashtable()

    Default Public Property Indexer(ByVal key As String)
        Set(value As Object)
            AddProperty(key, value)
        End Set
        Get
            Return GetProperty(key)
        End Get
    End Property

    Public Overridable Sub AddProperty(key As String, value As Object)
        Properties(key) = value
    End Sub

    Public Overridable Function GetProperty(key As String)
        Return Properties(key)
    End Function

    Public Function GetProperties() As Hashtable
        Return Properties
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Properties.GetEnumerator()
    End Function
End Class
