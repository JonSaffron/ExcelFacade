Imports JetBrains.Annotations

Public NotInheritable Class Shape
    Private ReadOnly _shape As Object

    Friend Sub New(<NotNull> ByVal shape As Object)
        Me._shape = shape
    End Sub

    Public Property Name As String
        Get
            Return Me._shape.Name
        End Get
        Set
            Me._shape.Name = value
        End Set
    End Property

    Public Property Placement As XlPlacement
        Get
            Return Me._shape.Placement
        End Get
        Set
            Me._shape.Placement = value
        End Set
    End Property

    Public ReadOnly Property Chart As Chart
        Get
            Return New Chart(Me._shape.Chart)
        End Get
    End Property
End Class
