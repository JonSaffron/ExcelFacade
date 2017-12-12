Imports JetBrains.Annotations

Public NotInheritable Class ChartFormat
    Private ReadOnly _chartFormat As Object

    Friend Sub New(<NotNull> ByVal chartFormat As Object)
        Me._chartFormat = chartFormat
    End Sub

    Public ReadOnly Property Fill As FillFormat
        Get
            Return New FillFormat(Me._chartFormat.Fill)
        End Get
    End Property

    Public ReadOnly Property Line As LineFormat
        Get
            Return New LineFormat(Me._chartFormat.Line)
        End Get
    End Property
End Class
