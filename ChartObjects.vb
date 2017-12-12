' ChartObjects is a 1 based collection of ChartObject objects
Imports JetBrains.Annotations

Public NotInheritable Class ChartObjects
    Implements IEnumerable(Of ChartObject)
    Implements IEnumerator(Of ChartObject)

    Private ReadOnly _chartObjects As Object

    Friend Sub New(<NotNull> ByVal chartObjects As Object)
        Me._chartObjects = chartObjects
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._chartObjects.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As ChartObject
        Get
            Return New ChartObject(Me._chartObjects.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal name As String) As ChartObject
        Get
            Return New ChartObject(Me._chartObjects.item(name))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfChartObject() As IEnumerator(Of ChartObject) Implements IEnumerable(Of ChartObject).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfChartObject As ChartObject Implements IEnumerator(Of ChartObject).Current
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
