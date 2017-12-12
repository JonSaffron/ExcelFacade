Imports JetBrains.Annotations

Public NotInheritable Class ColorStops
    Implements IEnumerable(Of ColorStop)
    Implements IEnumerator(Of ColorStop)
    Implements IDisposable

    Private ReadOnly _colorStops As Object

    Friend Sub New(<NotNull> ByVal colorStops As Object)
        Me._colorStops = colorStops
    End sub

    public ReadOnly Property Count as Integer
        get
            return Me._colorStops.Count
        End Get
    End Property

    default Public ReadOnly Property Item(byval index As integer)
        Get
            Return new ColorStop(Me._colorStops.Item(index))
        End Get
    End Property

    Public Function Add(byval position As Double) As ColorStop
        Dim result as New ColorStop(Me._colorStops.Add(position))
        Return result
    End Function

    Public sub Clear
        Call Me._colorStops.Clear()
    End sub

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfColorStop() As IEnumerator(Of ColorStop) Implements IEnumerable(Of ColorStop).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfColorStop As ColorStop Implements IEnumerator(Of ColorStop).Current
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
end Class
