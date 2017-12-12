Imports JetBrains.Annotations

Public NotInheritable Class Points
    Implements IEnumerable(Of Point)
    Implements IEnumerator(Of Point)
    Implements IDisposable

    Private ReadOnly _points As Object

    Friend Sub New(<NotNull> ByVal points As Object)
        Me._points = points
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._points.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Point
        Get
            Return new Point(Me._points.Item(index))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfPoint() As IEnumerator(Of Point) Implements IEnumerable(Of Point).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfPoint As Point Implements IEnumerator(Of Point).Current
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
