Imports JetBrains.Annotations

' This is not a collection, oddly enough

Public NotInheritable Class Errors
    Private ReadOnly _errors As Object

    Friend Sub New(<NotNull> ByVal shapes As Object)
        Me._errors = shapes
    End Sub

    Default Public ReadOnly Property Item(ByVal errortype As XlErrorChecks) As [Error]
        Get
            Return New [Error](Me._errors.Item(errortype))
        End Get
    End Property
End Class
