Imports JetBrains.Annotations

Public NotInheritable Class VPageBreak
    Private ReadOnly _vpagebreak As Object

    Friend Sub New(<NotNull> ByVal vpagebreak As Object)
        Me._vpagebreak = vpagebreak
    End Sub

    Public Property Extent As XlPageBreakExtent
        Get
            Return Me._vpagebreak.Extent
        End Get
        Set
            Me._vpagebreak.Extent = value
        End Set
    End Property

    Public Property Location As Range
        Get
            Return New Range(Me._vpagebreak.Location)
        End Get
        Set
            Me._vpagebreak.Location = value.underlyingComObject
        End Set
    End Property

    Public Property Type As XlPageBreak
        Get
            Return Me._vpagebreak.Type
        End Get
        Set
            Me._vpagebreak.Type = value
        End Set
    End Property

    Public Sub Delete()
        Call Me._vpagebreak.Delete()
    End Sub
End Class
