Imports JetBrains.Annotations

Public NotInheritable Class Axis
    Private ReadOnly _axis As Object

    Friend Sub New(<NotNull> ByVal axis As Object)
        Me._axis = axis
    End Sub

    Public Property MaximumScale As Double
        Get
            Return Me._axis.MaximumScale
        End Get
        Set
            Me._axis.MaximumScale = value
        End Set
    End Property

    Public Property MajorTickMark As XlTickMark
        Get
            Return Me._axis.MajorTickMark
        End Get
        Set
            Me._axis.MajorTickMark = value
        End Set
    End Property

    Public Property MinorTickMark As XlTickMark
        Get
            Return Me._axis.MinorTickMark
        End Get
        Set
            Me._axis.MinorTickMark = value
        End Set
    End Property

    Public Property ReversePlotOrder As Boolean
        Get
            Return Me._axis.ReversePlotOrder
        End Get
        Set
            Me._axis.ReversePlotOrder = value
        End Set
    End Property

    Public Property ScaleType As XlScaleType
        Get
            Return Me._axis.ScaleType
        End Get
        Set
            Me._axis.ScaleType = value
        End Set
    End Property


End Class
