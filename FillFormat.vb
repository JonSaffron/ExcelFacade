Imports JetBrains.Annotations

Public NotInheritable Class FillFormat
    Private ReadOnly _fillFormat As Object

    Friend Sub New(<NotNull> ByVal fillFormat As Object)
        Me._fillFormat = fillFormat
    End Sub

    Public ReadOnly Property ForeColor As ColorFormat
        Get
            Return New ColorFormat(Me._fillFormat.ForeColor)
        End Get
    End Property
End Class
