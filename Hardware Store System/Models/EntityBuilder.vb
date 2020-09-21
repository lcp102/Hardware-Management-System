Imports System.Data.OleDb


Public Class EntityBuilder

    Private ReadOnly Item As Entity
    Private Reader As OleDbDataReader
    Private IsClose As Boolean
    Private IsRead As Boolean
    Private IndexTypeMap As Hashtable

    Private Class KeyValuePair(Of K, V)
        Property Key As K
        Property Value As V
    End Class

    Public Sub New(Optional ByVal entityItem As Entity = Nothing)
        If IsNothing(entityItem) Then
            entityItem = New Entity()
        End If
        Item = entityItem
    End Sub

    Public Function UseOleReader(oleReader As OleDbDataReader, Optional ByVal read As Boolean = False, Optional ByVal close As Boolean = True) As EntityBuilder
        Reader = oleReader
        IsRead = read
        IsClose = close
        Return Me
    End Function

    Public Function ReadType(index As Integer, fieldType As Type, Optional ByVal propertyName As String = Nothing) As EntityBuilder
        If IsNothing(IndexTypeMap) Then
            IndexTypeMap = New Hashtable()
        End If

        IndexTypeMap(index) = New KeyValuePair(Of String, Type) With {.Key = propertyName, .Value = fieldType}
        Return Me
    End Function

    Public Function AddProperty(key As String, value As Object) As EntityBuilder
        Item.AddProperty(key, value)
        Return Me
    End Function

    Public Function Build() As Entity
        If Not IsNothing(Reader) Then
            If IsRead Then
                Reader.Read()
            End If

            Dim count As Integer = Reader.FieldCount
            For i As Integer = 0 To count - 1
                Dim key As String = Reader.GetName(i)
                Dim value As Object = Reader.GetValue(i)
                If Not IsNothing(IndexTypeMap) Then
                    If IndexTypeMap.ContainsKey(i) Then
                        Dim kvp As KeyValuePair(Of String, Type) = IndexTypeMap(i)

                        If Not IsNothing(kvp.Key) Then
                            key = kvp.Key
                        End If

                        'cast to type kvp.Value

                    End If
                End If
                AddProperty(key, value)
            Next

            If IsClose Then
                Reader.Close()
            End If
        End If

        Return Item
    End Function

End Class
