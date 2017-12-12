Imports JetBrains.Annotations

Public NotInheritable Class Characters
    Private ReadOnly _characters As Object

    Friend Sub New(<NotNull> ByVal characters As Object)
        Me._characters = characters
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._characters.Count
        End Get
    End Property

    Public ReadOnly Property Font As Font
        Get
            Return New Font(Me._characters.Font)
        End Get
    End Property

    Public Property Text As String
        Get
            Return Me._characters.Text
        End Get
        Set
            Me._characters.Text = value
        End Set
    End Property

    Public Sub Insert(ByVal textString As String)
        Call Me._characters.Insert(textString)
    End Sub

    Public Sub Delete()
        Call Me._characters.Delete()
    End Sub
End Class
