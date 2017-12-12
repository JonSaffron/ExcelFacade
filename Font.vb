Imports System.Drawing
Imports JetBrains.Annotations

Public NotInheritable Class Font
    Private ReadOnly _font As Object

    Friend Sub New(<NotNull> ByVal font As Object)
        Me._font = font
    End Sub

    Public Property Bold As Boolean?
        Get
            Return ToNullable(of Boolean)(Me._font.Bold)
        End Get
        Set
            Me._font.Bold = value.NullableToNull()
        End Set
    End Property

    Public Property Italic As Boolean?
        Get
            Return ToNullable(of Boolean)(Me._font.Italic)
        End Get
        Set
            Me._font.Italic = value.NullableToNull()
        End Set
    End Property

    Public Property Strikethrough As Boolean?
        Get
            Return ToNullable(of Boolean)(Me._font.Strikethrough)
        End Get
        Set
            Me._font.Strikethrough = value.NullableToNull()
        End Set
    End Property

    Public Property Underline As XlUnderlineStyle?
        Get
            Return ToNullable(Of XlUnderlineStyle)(Me._font.Underline)
        End Get
        Set
            Me._font.Underline = value.NullableToNull()
        End Set
    End Property

    Public Property Size As Double?
        Get
            Return ToNullable(of Double)(Me._font.Size)
        End Get
        Set
            Me._font.Size = value.NullableToNull()
        End Set
    End Property

    Public Property Name As String
        Get
            Dim returnValue As Object = Me._font.Name
            Return If(TypeOf returnValue Is DBNull, nothing, returnValue)
        End Get
        Set
            Me._font.Name = If(value, DBNull.Value)
        End Set
    End Property

    Public Property Color As Color?
        Get
            Return ToColor(Me._font.Color)
        End Get
        Set
            Me._font.Color = color.ToVbaColor()
        End Set
    End Property

    Public WriteOnly Property ColorIndex As XlColorIndex?
        Set
            Me._font.ColorIndex = value.NullableToNull()
        End Set
    End Property

    Public Property ColorIndexInPalette As Integer?
        Get
            Return ToNullable(of integer)(Me._font.ColorIndex)
        End Get
        Set
            Me._font.ColorIndex = value.NullableToNull()
        End Set
    End Property
End Class
