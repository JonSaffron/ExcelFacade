'<JetBrains.Annotations.PublicAPI> _
Imports JetBrains.Annotations

Public NotInheritable Class Axes
    Implements IEnumerable(Of Axis)
    Implements IEnumerator(Of Axis)

    Private ReadOnly _axes As Object

    Friend Sub New(<NotNull> ByVal axes As Object)
        Me._axes = axes
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._axes.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal axisType As XlAxisType) As Axis
        Get
            Return New Axis(Me._axes.Item(axisType))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfAxis() As IEnumerator(Of Axis) Implements IEnumerable(Of Axis).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfAxis As Axis Implements IEnumerator(Of Axis).Current
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
