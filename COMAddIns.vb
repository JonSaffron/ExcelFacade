' ReSharper disable InconsistentNaming
Imports JetBrains.Annotations

Public NotInheritable Class COMAddIns
    Implements IEnumerable(Of COMAddIn)
    Implements IEnumerator(Of COMAddIn)
    Implements IDisposable

    Private ReadOnly _COMAddIns As Object

    Friend Sub New(<NotNull> ByVal COMAddIns As Object)
        Me._COMAddIns = COMAddIns
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._COMAddIns.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As COMAddIn
        Get
            Return New COMAddIn(Me._COMAddIns.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal progId As String) As COMAddIn
        Get
            Return New COMAddIn(Me._COMAddIns.Item(progId))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfCOMAddIn() As IEnumerator(Of COMAddIn) Implements IEnumerable(Of COMAddIn).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfCOMAddIn As COMAddIn Implements IEnumerator(Of COMAddIn).Current
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
' ReSharper restore InconsistentNaming
