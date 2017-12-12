' HPageBreaks is a 1 based collection of HPageBreak objects
Imports JetBrains.Annotations

Public NotInheritable Class HPageBreaks
    Implements IEnumerable(Of HPageBreak)
    Implements IEnumerator(Of HPageBreak)
    Implements IDisposable

    Private ReadOnly _hpagebreaks As Object

    Friend Sub New(<NotNull> ByVal hpagebreaks As Object)
        Me._hpagebreaks = hpagebreaks
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._hpagebreaks.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As HPageBreak
        Get
            Return New HPageBreak(Me._hpagebreaks.Item(index))
        End Get
    End Property

    Public Function Add(ByVal before As Range) As HPageBreak
        Return New HPageBreak(Me._hpagebreaks.Add(before.underlyingComObject))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfHPageBreak() As IEnumerator(Of HPageBreak) Implements IEnumerable(Of HPageBreak).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfHPageBreak As HPageBreak Implements IEnumerator(Of HPageBreak).Current
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
