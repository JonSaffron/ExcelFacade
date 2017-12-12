Imports JetBrains.Annotations

Public NotInheritable Class SeriesCollection
    Implements IEnumerable(Of Series)
    Implements IEnumerator(Of Series)
    Implements IDisposable

    Private ReadOnly _seriescollection As Object

    Friend Sub New(<NotNull> ByVal seriescollection As Object)
        Me._seriescollection = seriescollection
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._seriescollection.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Series
        Get
            Return new Series(Me._seriescollection.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal name As String) As Series
        Get
            Return new Series(Me._seriescollection.Item(name))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfSeries() As IEnumerator(Of Series) Implements IEnumerable(Of Series).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfSeries As Series Implements IEnumerator(Of Series).Current
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
