Imports JetBrains.Annotations

Public NotInheritable Class LineFormat
    Private ReadOnly _lineFormat As Object

    Friend Sub New(<NotNull> ByVal lineFormat As Object)
        Me._lineFormat = lineFormat
    End Sub

    ' this is actually of type MsoTriState, however only True and False are valid
    Public Property Visible As Boolean
        Get
            Return Me._lineFormat.Visible
        End Get
        Set
            Me._lineFormat.Visible = value
        End Set
    End Property
End Class
