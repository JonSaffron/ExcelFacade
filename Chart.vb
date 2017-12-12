Imports System.Reflection
Imports JetBrains.Annotations

Public NotInheritable Class Chart
    Inherits Sheet

    Friend Sub New(<NotNull> ByVal chart As Object)
        Call MyBase.New(chart)
    End Sub

    Public ReadOnly Property ChartTitle As ChartTitle
        Get
            Return New ChartTitle(Me.underlyingComObject.ChartTitle)
        End Get
    End Property

    Public ReadOnly Property SeriesCollection As SeriesCollection
        Get
            Return New SeriesCollection(Me.underlyingComObject.SeriesCollection)
        End Get
    End Property

    Public Property ChartStyle As Object
        Get
            Return Me.underlyingComObject.ChartStyle
        End Get
        Set
            Me.underlyingComObject.ChartStyle = value
        End Set
    End Property

    Public Property ChartType As XlChartType
        Get
            Return Me.underlyingComObject.ChartType
        End Get
        Set
            Me.underlyingComObject.ChartType = value
        End Set
    End Property

    Public Sub SetSourceData(ByVal range As Range)
        Call Me.underlyingComObject.SetSourceData(range.underlyingComObject)
    End Sub

    Public Function Location(ByVal chartLocation As XlChartLocation, Optional ByVal sheetName As String = Nothing) As Chart
        Return New Chart(Me.underlyingComObject.Location(chartLocation, If(sheetName IsNot Nothing, sheetName, Missing.Value)))
    End Function

    Public ReadOnly Property Axes As Axes
        Get
            Return New Axes(Me.underlyingComObject.Axes)
        End Get
    End Property
End Class
