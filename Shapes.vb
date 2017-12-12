Imports System.Reflection
Imports JetBrains.Annotations

Public NotInheritable Class Shapes
    Implements IEnumerable(Of Shape)
    Implements IEnumerator(Of Shape)
    Implements IDisposable

    Private ReadOnly _shapes As Object

    Friend Sub New(<NotNull> ByVal shapes As Object)
        Me._shapes = shapes
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._shapes.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Shape
        Get
            Return Me._shapes.Item(index)
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal shapename As String) As Shape
        Get
            Return Me._shapes.Item(shapename)
        End Get
    End Property

    Public Function AddPicture(ByVal filename As String, ByVal linktofile As Boolean, ByVal savewithdocument As Boolean, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single) As Shape
        Return New Shape(Me._shapes.AddPicture(filename, linktofile, savewithdocument, left, top, width, height))
    End Function

    public Function AddChart(Optional ByVal chartType As XlChartType? = Nothing, Optional ByVal left As Single? = Nothing, Optional ByVal top As Single? = Nothing, Optional ByVal width As Single? = Nothing, Optional ByVal height As Single? = Nothing) As Shape
        if Decimal.Parse(Me._shapes.Application.Version) > 14d Then
            Throw New NotImplementedException("Sorry, this property should only be used in Excel 2010 and earlier. Use AddChart2 from Excel 2013 onwards.")
        End If

        Dim chartType2 As Object = If(chartType.HasValue, chartType.Value, Missing.Value)
        Dim left2 As Object = If(left.HasValue, left.Value, Missing.Value)
        Dim top2 As Object = If(top.HasValue, top.Value, Missing.Value)
        Dim width2 As Object = If(width.HasValue, width.Value, Missing.Value)
        Dim height2 As Object = If(height.HasValue, height.Value, Missing.Value)
        Return New Shape(Me._shapes.AddChart(chartType2, left2, top2, width2, height2))
    End Function

    Public Function AddChart2(Optional ByVal style As Integer? = Nothing, Optional ByVal chartType As XlChartType? = Nothing, Optional ByVal left As Single? = Nothing, Optional ByVal top As Single? = Nothing, Optional ByVal width As Single? = Nothing, Optional ByVal height As Single? = Nothing, Optional ByVal newLayout As Boolean? = Nothing) As Shape
        if Decimal.Parse(Me._shapes.Application.Version) <= 14d Then
            Throw New NotImplementedException("Sorry, this property only became available in Excel 2013.")
        End If

        Dim style2 As Object = If(style.HasValue, style.Value, Missing.Value)
        Dim chartType2 As Object = If(chartType.HasValue, chartType.Value, Missing.Value)
        Dim left2 As Object = If(left.HasValue, left.Value, Missing.Value)
        Dim top2 As Object = If(top.HasValue, top.Value, Missing.Value)
        Dim width2 As Object = If(width.HasValue, width.Value, Missing.Value)
        Dim height2 As Object = If(height.HasValue, height.Value, Missing.Value)
        Dim newLayout2 As Object = If(newLayout.HasValue, newLayout.Value, Missing.Value)
        Return New Shape(Me._shapes.AddChart2(style2, chartType2, left2, top2, width2, height2, newLayout2))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfShape() As IEnumerator(Of Shape) Implements IEnumerable(Of Shape).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfShape As Shape Implements IEnumerator(Of Shape).Current
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
