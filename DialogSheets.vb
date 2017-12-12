Imports JetBrains.Annotations

Public NotInheritable Class DialogSheets
    Inherits Sheets

    Friend Sub New(<NotNull> ByVal dialogsheets As Object)
        Call MyBase.New(dialogsheets)
    End Sub

    Default Public Shadows ReadOnly Property Item(ByVal index As Integer) As DialogSheet
        Get
            Return MyBase.Item(index)
        End Get
    End Property

    Default Public Shadows ReadOnly Property Item(ByVal dialogsheetname As String) As DialogSheet
        Get
            Return MyBase.Item(dialogsheetname)
        End Get
    End Property
End Class
