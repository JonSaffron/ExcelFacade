Imports JetBrains.Annotations

Public NotInheritable Class Worksheets
    Inherits Sheets

    Friend Sub New(<NotNull> ByVal worksheets As Object)
        Call MyBase.New(worksheets)
    End Sub

    Default Public Shadows ReadOnly Property Item(ByVal index As Integer) As Worksheet
        Get
            Return MyBase.Item(index)
        End Get
    End Property

    Default Public Shadows ReadOnly Property Item(ByVal worksheetname As String) As Worksheet
        Get
            Return MyBase.Item(worksheetname)
        End Get
    End Property
End Class
