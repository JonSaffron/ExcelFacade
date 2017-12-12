Imports JetBrains.Annotations

Public NotInheritable Class RectangularGradient
    inherits Gradient

    Friend Sub New(<NotNull> ByVal rectangularGradient As Object)
        Call MyBase.New(rectangularGradient)
    End sub

    public Property RectangleBottom As Double
        get
            Return Me._gradient.RectangleBottom
        End Get
        Set
            Me._gradient.RectangleBottom = value
        End Set
    End Property

    public Property RectangleLeft As Double
        get
            Return Me._gradient.RectangleLeft
        End Get
        Set
            Me._gradient.RectangleLeft = value
        End Set
    End Property

    public Property RectangleRight As Double
        get
            Return Me._gradient.RectangleRight
        End Get
        Set
            Me._gradient.RectangleRight = value
        End Set
    End Property

    public Property RectangleTop As Double
        get
            Return Me._gradient.RectangleTop
        End Get
        Set
            Me._gradient.RectangleTop = value
        End Set
    End Property
End Class
