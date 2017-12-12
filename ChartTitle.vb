Imports JetBrains.Annotations

Public NotInheritable Class ChartTitle
    Private ReadOnly _charttitle As Object

    Friend Sub New(<NotNull> ByVal charttitle As Object)
        Me._charttitle = charttitle
    End Sub

    Public Property Text As String
        Get
            Return Me._charttitle.Text
        End Get
        Set
            Me._charttitle.Text = value
        End Set
    End Property
End Class
