Imports System.Drawing
Imports JetBrains.Annotations

Public NotInheritable Class ColorFormat
    Private ReadOnly _colorFormat As Object

    Friend Sub New(<NotNull> ByVal colorFormat As Object)
        Me._colorFormat = colorFormat
    End Sub

' ReSharper disable InconsistentNaming
    Public Property RGB As Color
' ReSharper restore InconsistentNaming
        Get
            Dim returnValue As Integer = Me._colorFormat.RGB
            Dim c As Color = Color.FromArgb(ReturnValue)
            Return Color.FromArgb(255, c.B, c.G, c.R)
        End Get
        Set
            Dim c As Color = Color.FromArgb(0, value.B, value.G, value.R)
            Me._colorFormat.RGB = c.ToArgb()
        End Set
    End Property
End Class
