Imports JetBrains.Annotations

Public NotInheritable Class PivotCaches
    Implements IEnumerable(Of PivotCache)
    Implements IEnumerator(Of PivotCache)
    Implements IDisposable

    Private ReadOnly _pivotcaches As Object

    Friend Sub New(<NotNull> ByVal pivotcaches As Object)
        Me._pivotcaches = pivotcaches
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._pivotcaches.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As PivotCache
        Get
            Return New PivotCache(Me._pivotcaches.Item(index))
        End Get
    End Property

    Public Function Add(ByVal sourceType As XlPivotTableSourceType, ByVal sourceData As String) As PivotCache
        Return New PivotCache(Me._pivotcaches.Add(sourceType, sourceData))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfPivotCache() As IEnumerator(Of PivotCache) Implements IEnumerable(Of PivotCache).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfPivotCache As PivotCache Implements IEnumerator(Of PivotCache).Current
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
