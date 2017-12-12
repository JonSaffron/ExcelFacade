Imports JetBrains.Annotations

Public NotInheritable Class PivotFields
    Implements IEnumerable(Of PivotField)
    Implements IEnumerator(Of PivotField)
    Implements IDisposable

    Private ReadOnly _pivotfields As Object

    Friend Sub New(<NotNull> ByVal pivotfields As Object)
        Me._pivotfields = pivotfields
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._pivotfields.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As PivotField
        Get
            Return New PivotField(Me._pivotfields.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal pivotfieldname As String) As PivotField
        Get
            Return New PivotField(Me._pivotfields.Item(pivotfieldname))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfPivotField() As IEnumerator(Of PivotField) Implements IEnumerable(Of PivotField).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfPivotField As PivotField Implements IEnumerator(Of PivotField).Current
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
