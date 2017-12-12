Imports System.Drawing
Imports JetBrains.Annotations

Public NotInheritable Class Border
    Private ReadOnly _border As Object

    Friend Sub New(<NotNull> ByVal border As Object)
        Me._border = border
    End Sub

    Public Property LineStyle As XlLineStyle
        Get
            Return Me._border.LineStyle
        End Get
        Set
            Me._border.LineStyle = value
        End Set
    End Property

    Public Property Weight As XlBorderWeight
        Get
            Return Me._border.Weight
        End Get
        Set
            Me._border.Weight = value
        End Set
    End Property

    Public Property Color As Color
        Get
            Dim returnValue As Integer = Me._border.Color
            Dim c As Color = Color.FromArgb(returnValue)
            Return Color.FromArgb(255, c.B, c.G, c.R)    ' fully opaque
        End Get
        Set
            Dim c As Color = Color.FromArgb(0, value.B, value.G, value.R)
            Me._border.Color = c.ToArgb()
        End Set
    End Property

    Public WriteOnly Property ColorIndex As XlColorIndex
        Set
            Me._border.ColorIndex = value
        End Set
    End Property

    Public Property ColorIndexInPalette As Integer
        Get
            Return Me._border.ColorIndex
        End Get
        Set
            Me._border.ColorIndex = value
        End Set
    End Property
End Class
