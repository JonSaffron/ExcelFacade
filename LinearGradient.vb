Imports JetBrains.Annotations

Public NotInheritable Class LinearGradient
    inherits Gradient

    Friend Sub New(<NotNull> ByVal linearGradient As Object)
        Call MyBase.New(linearGradient)
    End sub
    
    public Property Degree As Double
        get
            Return Me._gradient.Degree
        End Get
        Set
            Me._gradient.Degree = value
        End Set
    End Property
End Class
