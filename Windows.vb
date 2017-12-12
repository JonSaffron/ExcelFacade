' Windows is a 1 based collection of Window objects
Imports JetBrains.Annotations

Public NotInheritable Class Windows
    Implements IEnumerable(Of Window)
    Implements IEnumerator(Of Window)
    Implements IDisposable

    Private ReadOnly _windows As Object

    Friend Sub New(<NotNull> ByVal windows As Object)
        Me._windows = windows
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._windows.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Window
        Get
            Return New Window(Me._windows.Item(index))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfWindow() As IEnumerator(Of Window) Implements IEnumerable(Of Window).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfWindow As Window Implements IEnumerator(Of Window).Current
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
