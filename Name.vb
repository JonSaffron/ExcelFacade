Imports JetBrains.Annotations

Public NotInheritable Class Name
    Private ReadOnly _name As Object

    Friend Sub New(<NotNull> ByVal name As Object)
        Me._name = name
    End Sub

    Public Property Name As String
        Get
            Return Me._name.Name
        End Get
        Set
            Me._name.Name = value
        End Set
    End Property

    Public ReadOnly Property RefersToRange As Range
        Get
            Return New Range(Me._name.RefersToRange)
        End Get
    End Property

    Public Sub Delete()
        Call Me._name.Delete()
    End Sub
End Class
