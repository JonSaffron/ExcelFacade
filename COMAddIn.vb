' ReSharper disable InconsistentNaming
Imports JetBrains.Annotations

Public NotInheritable Class COMAddIn
    Private ReadOnly _COMAddIn As Object

    Friend Sub New(<NotNull> ByVal COMAddIn As Object)
        Me._COMAddIn = COMAddIn
    End Sub

    Public Property Connect As Boolean
        Get
            Return Me._COMAddIn.Connect
        End Get
        Set
            Me._COMAddIn.Connect = value
        End Set
    End Property

    Public ReadOnly Property Description As String
        Get
            Return Me._COMAddIn.Description
        End Get
    End Property

    Public ReadOnly Property Guid As Guid
        Get
            Dim g As String = Me._COMAddIn.Guid
            Return New Guid(g)
        End Get
    End Property

    Public ReadOnly Property ProdId As String
        Get
            Return Me._COMAddIn.ProgId
        End Get
    End Property
End Class
' ReSharper restore InconsistentNaming
