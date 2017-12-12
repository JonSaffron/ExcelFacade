Imports JetBrains.Annotations

Public NotInheritable Class ChartObject
    Private ReadOnly _chartObject As Object

    Friend Sub New(<NotNull> ByVal chartObject As Object)
        Me._chartObject = chartObject
    End Sub

    Public ReadOnly Property Chart As Chart
        Get
            Return New Chart(Me._chartObject.Chart)
        End Get
    End Property

    Public Property Name As String
        Get
            Return Me._chartObject.Name
        End Get
        Set
            Me._chartObject.Name = value
        End Set
    End Property
End Class
