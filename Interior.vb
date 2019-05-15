Imports System.Drawing
Imports JetBrains.Annotations

Public NotInheritable Class Interior
    Private ReadOnly _interior As Object

    Friend Sub New(<NotNull> ByVal interior As Object)
        Me._interior = interior
    End Sub

    Public Property Color As Color?
        Get
            Return ToColor(Me._interior.Color)
        End Get
        Set
            Me._interior.Color = color.ToVbaColor()
        End Set
    End Property

    Public WriteOnly Property ColorIndex As XlColorIndex?
        Set
            Me._interior.ColorIndex = value.NullableToNull()
        End Set
    End Property

    Public Property ColorIndexInPalette As Integer?
        Get
            Return ToNullable(of integer)(Me._interior.ColorIndex)
        End Get
        Set
            Me._interior.ColorIndex = value.NullableToNull()
        End Set
    End Property

    Public Property Pattern As XlPattern?
        Get
            Return ToNullable(of XlPattern)(Me._interior.Pattern)
        End Get
        Set
            Me._interior.Pattern = value.NullableToNull()
        End Set
    End Property

    ''' <summary>
    ''' Returns the Gradient object for the Interior
    ''' </summary>
    ''' <returns>A LinearGradient object, or a RectangularGradient object.</returns>
    Public ReadOnly Property Gradient as Gradient
        Get
            Dim comObject As Object = Me._interior.Gradient
            If comObject Is Nothing Then
                Return Nothing
            End If
            If GetComTypeName(comObject) = "LinearGradient" Then
                ' Should be true if Pattern = xlPatternLinearGradient
                Return New LinearGradient(comObject)
            End If
            If GetComTypeName(comObject) = "RectangularGradient" Then
                ' Should be true if Pattern = xlPatternRectangularGradient
                Return New RectangularGradient(comObject)
            End If
            Throw New InvalidOperationException("Gradient property is not of a recognised type.")
        End Get
    End Property

    ''' <summary>
    ''' Returns the correctly typed Gradient object for the Interior
    ''' </summary>
    ''' <returns>A LinearGradient object if Interior.Pattern = xlPatternLinearGradient, otherwise null</returns>
    ''' <remarks>This does not match an Excel property, but is provided for convenience.</remarks>
    public ReadOnly Property LinearGradient As LinearGradient
        get
            Return TryCast(Me.Gradient, LinearGradient)
        End Get
    end property

    ''' <summary>
    ''' Returns the correctly typed Gradient object for the Interior
    ''' </summary>
    ''' <returns>A RectangularGradient object if Interior.Pattern = xlPatternRectangularGradient</returns>
    ''' <remarks>This does not match an Excel property, but is provided for convenience.</remarks>
    public ReadOnly Property RectangularGradient As RectangularGradient
        get
            Return TryCast(Me.Gradient, RectangularGradient)
        End Get
    end property
End Class
