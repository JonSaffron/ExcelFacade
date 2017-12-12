Imports JetBrains.Annotations

Public NotInheritable Class Point
    Private ReadOnly _point As Object

    Friend Sub New(<NotNull> ByVal point As Object)
        Me._point = point
    End Sub

    public property HasDataLabel as boolean
        Get
            return me._point.HasDataLabel
        End Get
        Set
            Me._point.HasDataLabel = value
        End Set
    End Property

    Public ReadOnly Property Format As ChartFormat
        Get
            Return New ChartFormat(Me._point.Format)
        End Get
    End Property
End Class
