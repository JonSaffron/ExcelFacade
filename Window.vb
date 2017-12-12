Imports JetBrains.Annotations

Public NotInheritable Class Window
    Private ReadOnly _window As Object

    Friend Sub New(<NotNull> ByVal window As Object)
        Me._window = window
    End Sub

    Public Property FreezePanes As Boolean
        Get
            Return Me._window.FreezePanes
        End Get
        Set
            Me._window.FreezePanes = value
        End Set
    End Property

    Public Property SplitRow As Integer
        Get
            Return Me._window.SplitRow
        End Get
        Set
            Me._window.SplitRow = value
        End Set
    End Property

    Public Property SplitColumn As Integer
        Get
            Return Me._window.SplitColumn
        End Get
        Set
            Me._window.SplitColumn = value
        End Set
    End Property

    Public Property DisplayZeros As Boolean
        Get
            Return Me._window.DisplayZeros
        End Get
        Set
            Me._window.DisplayZeros = value
        End Set
    End Property

    Public Property DisplayGridlines As Boolean
        Get
            Return Me._window.DisplayGridlines
        End Get
        Set
            Me._window.DisplayGridlines = value
        End Set
    End Property

    Public Property View As XlWindowView
        Get
            Return Me._window.View
        End Get
        Set
            Me._window.View = value
        End Set
    End Property
End Class
