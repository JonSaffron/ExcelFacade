Imports JetBrains.Annotations

Public NotInheritable Class HPageBreak
    Private ReadOnly _hpagebreak As Object

    Friend Sub New(<NotNull> ByVal hpagebreak As Object)
        Me._hpagebreak = hpagebreak
    End Sub

    Public Property Extent As XlPageBreakExtent
        Get
            Return Me._hpagebreak.Extent
        End Get
        Set
            Me._hpagebreak.Extent = value
        End Set
    End Property

    Public Property Location As Range
        Get
            Return New Range(Me._hpagebreak.Location)
        End Get
        Set
            Me._hpagebreak.Location = value.underlyingComObject
        End Set
    End Property

    Public Property Type As XlPageBreak
        Get
            Return Me._hpagebreak.Type
        End Get
        Set
            Me._hpagebreak.Type = value
        End Set
    End Property

    Public Sub Delete()
        Call Me._hpagebreak.Delete()
    End Sub
End Class
