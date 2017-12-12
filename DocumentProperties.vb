' Bookmarks is a 1 based collection of object DocumentProperty
Imports JetBrains.Annotations

Public NotInheritable Class DocumentProperties
    Implements IEnumerable(Of DocumentProperty)
    Implements IEnumerator(Of DocumentProperty)
    Implements IDisposable

    Private ReadOnly _documentproperties As Object

    Friend Sub New(<NotNull> ByVal documentproperties As Object)
        Me._documentproperties = documentproperties
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._documentproperties.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As DocumentProperty
        Get
            Return New DocumentProperty(Me._documentproperties.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal propertyname As String) As DocumentProperty
        Get
            Return New DocumentProperty(Me._documentproperties.Item(propertyname))
        End Get
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfDocumentProperty() As IEnumerator(Of DocumentProperty) Implements IEnumerable(Of DocumentProperty).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfDocumentProperty As DocumentProperty Implements IEnumerator(Of DocumentProperty).Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Me._enumeratorPostion += 1
        Return (Me._enumeratorPostion <= Me.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me._enumeratorPostion = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose
        ' nothing to do
    End Sub
#End Region

End Class
