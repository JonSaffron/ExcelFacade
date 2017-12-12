Imports JetBrains.Annotations

Public NotInheritable Class PivotTable
    Private ReadOnly _pivottable As Object

    Friend Sub New(<NotNull> ByVal pivottable As Object)
        Me._pivottable = pivottable
    End Sub

    Public Sub AddFields(ByVal rowFields As String, ByVal columnFields As String)
        Call Me._pivottable.AddFields(rowFields, columnFields)
    End Sub

    Public Sub AddFields(ByVal rowFields As String(), ByVal columnFields As String)
        Call Me._pivottable.AddFields(rowFields, columnFields)
    End Sub

    Public Sub AddFields(ByVal rowFields As String, ByVal columnFields As String())
        Call Me._pivottable.AddFields(rowFields, columnFields)
    End Sub

    Public Sub AddFields(ByVal rowFields As String(), ByVal columnFields As String())
        Call Me._pivottable.AddFields(rowFields, columnFields)
    End Sub

    Public Function PivotFields(ByVal index As Integer) As PivotField
        Return New PivotField(Me._pivottable.PivotFields(index))
    End Function

    Public Function PivotFields(ByVal pivotfieldname As String) As PivotField
        Return New PivotField(Me._pivottable.PivotFields(pivotfieldname))
    End Function
End Class
