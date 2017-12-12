Imports JetBrains.Annotations

Public NotInheritable Class PivotCache
' ReSharper disable InconsistentNaming
    Friend ReadOnly _pivotcache As Object
' ReSharper restore InconsistentNaming

    Friend Sub New(<NotNull> ByVal pivotcache As Object)
        Me._pivotcache = pivotcache
    End Sub

    Public Function CreatePivotTable(ByVal tableDestination As String, ByVal tableName As String) As PivotTable
        Return New PivotTable(Me._pivotcache.CreatePivotTable(tableDestination, tableName))
    End Function
End Class
