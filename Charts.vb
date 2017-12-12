Imports System.Reflection
Imports JetBrains.Annotations

Public NotInheritable Class Charts
    Inherits Sheets

    Friend Sub New(<NotNull> ByVal charts As Object)
        Call MyBase.New(charts)
    End Sub

    Public Overrides Function Add() As Sheet
        Return Add(Nothing, Nothing, 1)
    End Function

    Public Overrides Function Add(ByVal before As Sheet) As Sheet
        Return Add(before, Nothing, 1)
    End Function

    Public Overrides Function Add(ByVal before As Sheet, ByVal after As Sheet)
        Return Add(before, after, 1)
    End Function

    Public Overrides Function Add(ByVal before As Sheet, ByVal after As Sheet, ByVal countOfCharts As Integer)
        Dim b As Object
        If before Is Nothing Then
            b = Missing.Value
        Else
            b = before.underlyingComObject
        End If
        Dim a As Object
        If after Is Nothing Then
            a = Missing.Value
        Else
            a = after.underlyingComObject
        End If
        Return Sheet.CreateSheetObject(Me._sheets.Add(b, a, countOfCharts))
    End Function

    <Obsolete>
    Public Overrides Function Add(ByVal before As Sheet, ByVal after As Sheet, ByVal countOfCharts As Integer, ByVal type As XlSheetType) As Sheet
        Throw New InvalidOperationException()
    End Function

    Default Public Shadows ReadOnly Property Item(ByVal index As Integer) As Chart
        Get
            Return MyBase.Item(index)
        End Get
    End Property

    Default Public Shadows ReadOnly Property Item(ByVal chartname As String) As Chart
        Get
            Return MyBase.Item(chartname)
        End Get
    End Property
End Class
