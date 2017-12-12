Imports System.Drawing
Imports JetBrains.Annotations

Public NotInheritable Class ColorStop
    Private ReadOnly _colorStop As Object

    Friend Sub New(<NotNull> ByVal colorStop As Object)
        Me._colorStop = colorStop
    End Sub

    Public Property Color As Color
        get
            Dim returnValue As Integer = Me._colorStop.Color
            Dim c As Color = Color.FromArgb(returnValue)
            Return Color.FromArgb(255, c.B, c.G, c.R)    ' fully opaque
        End Get
        Set
            Dim c As Color = Color.FromArgb(0, value.B, value.G, value.R)
            Me._colorStop.Color = c.ToArgb()
        End Set
    End Property
End Class
