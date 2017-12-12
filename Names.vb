Imports JetBrains.Annotations

Public NotInheritable Class Names
    Implements IEnumerable(Of Name)
    Implements IEnumerator(Of Name)
    Implements IDisposable

    Private ReadOnly _names As Object

    Friend Sub New(<NotNull> ByVal names As Object)
        Me._names = names
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._names.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Name
        Get
            Return New Name(Me._names.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal name As String) As Name
        Get
            Return New Name(Me._names.Item(name))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfName() As IEnumerator(Of Name) Implements IEnumerable(Of Name).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfName As Name Implements IEnumerator(Of Name).Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Me._enumeratorPosition += 1
        Return (Me._enumeratorPosition <= Me.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me._enumeratorPosition = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose
        ' nothing to do
    End Sub
#End Region

End Class
