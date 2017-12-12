Public MustInherit Class Gradient
' ReSharper disable InconsistentNaming
    protected ReadOnly _gradient as Object
' ReSharper restore InconsistentNaming

    protected sub New(byval gradient As object)
        me._gradient = gradient
    End sub

    Public ReadOnly Property ColorStops as ColorStops
        get
            Return New ColorStops(Me._gradient.ColorStops)
        End Get
    End Property
End Class
