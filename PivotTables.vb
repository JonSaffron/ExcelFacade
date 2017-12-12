Imports JetBrains.Annotations

Public NotInheritable Class PivotTables
    Implements IEnumerable(Of PivotTable)
    Implements IEnumerator(Of PivotTable)
    Implements IDisposable

    Private ReadOnly _pivottables As Object

    Friend Sub New(<NotNull> ByVal pivottables As Object)
        Me._pivottables = pivottables
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._pivottables.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As PivotTable
        Get
            Return New PivotTable(Me._pivottables.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal pivotTableName As String) As PivotTable
        Get
            Return New PivotTable(Me._pivottables.Item(pivotTableName))
        End Get
    End Property

    Public Function Add(ByVal pivotCache As PivotCache, ByVal tableDestination As String, ByVal tableName As String) As PivotTable
        Return New PivotTable(Me._pivottables.Add(pivotCache._pivotcache, tableDestination, tableName))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfPivotTable() As IEnumerator(Of PivotTable) Implements IEnumerable(Of PivotTable).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfPivotTable As PivotTable Implements IEnumerator(Of PivotTable).Current
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
