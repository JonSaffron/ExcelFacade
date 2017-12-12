Imports JetBrains.Annotations

Public NotInheritable Class [Error]
    Private ReadOnly _error As Object

    Friend Sub New(<NotNull> ByVal [error] As Object)
        Me._error = [error]
    End Sub

    Public Property Ignore As Boolean
        Get
            Return Me._error.Ignore
        End Get
        Set
            Me._error.Ignore = value
        End Set
    End Property

    Public ReadOnly Property Value As Boolean
        Get
            Return Me._error.Value
        End Get
    End Property
End Class
