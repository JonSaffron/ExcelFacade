Imports JetBrains.Annotations

Public NotInheritable Class Series
    Private ReadOnly _series As Object

    Friend Sub New(<NotNull> ByVal series As Object)
        Me._series = series
    End Sub

    public property HasDataLabels as boolean
        Get
            return me._series.HasDataLabels
        End Get
        Set
            me._series.HasDataLabels = value
        End Set
    End Property

    public readonly property Points as Points
        Get
            return new Points(me._series.Points)
        End Get
    End Property

    public readonly property Values as Array
        Get
            Dim x as array = me._series.Values
            return x
        End Get
    End Property

    public property Name as string
        Get
            return me._series.Name
        End Get
        Set
            me._series.Name = value
        End Set
    End Property

    Public Property FormulaR1C1 As String
        Get
            Return Me._series.FormulaR1C1
        End Get
        Set
            Me._series.FormulaR1C1 = value
        End Set
    End Property

    Public ReadOnly Property Format As ChartFormat
        Get
            Return New ChartFormat(Me._series.Format)
        End Get
    End Property

    Public Property MarkerSize As Integer
        Get
            Return Me._series.MarkerSize
        End Get
        Set
            Me._series.MarkerSize = value
        End Set
    End Property
End Class
