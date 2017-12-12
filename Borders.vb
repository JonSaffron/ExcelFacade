Imports System.Drawing
Imports JetBrains.Annotations

Public NotInheritable Class Borders
    Implements IEnumerable(Of Border)
    Implements IEnumerator(Of Border)

    Private ReadOnly _borders As Object

    Friend Sub New(<NotNull> ByVal borders As Object)
        Me._borders = borders
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._borders.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As XlBordersIndex) As Border
        Get
            Return New Border(Me._borders.Item(index))
        End Get
    End Property

    Public Property LineStyle As XlLineStyle
        Get
            Return Me._borders.LineStyle
        End Get
        Set
            Me._borders.LineStyle = value
        End Set
    End Property

    Public Property Weight As XlBorderWeight
        Get
            Return Me._borders.Weight
        End Get
        Set
            Me._borders.Weight = value
        End Set
    End Property

    Public Property Color As Color
        Get
            Dim returnValue As Integer = Me._borders.Color
            Dim c As Color = Color.FromArgb(returnValue)
            Return Color.FromArgb(255, c.B, c.G, c.R)
        End Get
        Set
            Dim c As Color = Color.FromArgb(0, value.B, value.G, value.R)
            Me._borders.Color = c.ToArgb()
        End Set
    End Property

    Public WriteOnly Property ColorIndex As XlColorIndex
        Set
            Me._borders.ColorIndex = value
        End Set
    End Property

    Public Property ColorIndexInPalette As Integer
        Get
            Return Me._borders.ColorIndex
        End Get
        Set
            Me._borders.ColorIndex = value
        End Set
    End Property

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfBorder() As IEnumerator(Of Border) Implements IEnumerable(Of Border).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfBorder As Border Implements IEnumerator(Of Border).Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPosition)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Me._enumeratorPosition += 1
        Return (Me._enumeratorPosition <= Me.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me._enumeratorPosition = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose
        ' nothing to do
    End Sub
#End Region

End Class
