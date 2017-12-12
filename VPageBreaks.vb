' VPageBreaks is a 1 based collection of VPageBreak objects
Imports JetBrains.Annotations

Public NotInheritable Class VPageBreaks
    Implements IEnumerable(Of VPageBreak)
    Implements IEnumerator(Of VPageBreak)
    Implements IDisposable

    Private ReadOnly _vpagebreaks As Object

    Friend Sub New(<NotNull> ByVal vpagebreaks As Object)
        Me._vpagebreaks = vpagebreaks
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._vpagebreaks.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As VPageBreak
        Get
            Return New VPageBreak(Me._vpagebreaks.Item(index))
        End Get
    End Property

    Public Function Add(ByVal before As Range) As VPageBreak
        Return New VPageBreak(Me._vpagebreaks.Add(before.underlyingComObject))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfVPageBreak() As IEnumerator(Of VPageBreak) Implements IEnumerable(Of VPageBreak).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfVPageBreak As VPageBreak Implements IEnumerator(Of VPageBreak).Current
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
