Imports JetBrains.Annotations

Public NotInheritable Class PivotField
    Private ReadOnly _pivotfield As Object

    Friend Sub New(<NotNull> ByVal pivotfield As Object)
        Me._pivotfield = pivotfield
    End Sub

    Public Property Orientation As XlPivotFieldOrientation
        Get
            Return Me._pivotfield.Orientation
        End Get
        Set
            Me._pivotfield.Orientation = value
        End Set
    End Property

    Public Property NumberFormat As String
        Get
            Return Me._pivotfield.NumberFormat
        End Get
        Set
            Me._pivotfield.NumberFormat = value
        End Set
    End Property

    ' returns a 12 element array for each xlSubtotals type
    Public Property Subtotals As Boolean()
        Get
            Return Me._pivotfield.Subtotals
        End Get
        Set
            Me._pivotfield.Subtotals = value
        End Set
    End Property

    Public Property Subtotals(ByVal index As XlSubtotals) As Boolean
        Get
            Return Me._pivotfield.Subtotals(index)
        End Get
        Set
            Me._pivotfield.Subtotals(index) = value
        End Set
    End Property
End Class
