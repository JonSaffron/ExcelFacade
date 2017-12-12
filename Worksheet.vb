Imports JetBrains.Annotations

Public NotInheritable Class Worksheet
    Inherits Sheet

    Friend Sub New(<NotNull> ByVal worksheet As Object)
        Call MyBase.New(worksheet)
    End Sub

    Public ReadOnly Property Cells As Range
        Get
            Return New Range(Me.underlyingComObject.Cells)
        End Get
    End Property

    Public ReadOnly Property Columns As Range
        Get
            Return New Range(Me.underlyingComObject.Columns)
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As Range) As Range
        Get
            Return New Range(Me.underlyingComObject.Range(cell1.underlyingComObject))
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As String) As Range
        Get
            Return New Range(Me.underlyingComObject.Range(cell1))
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As Range, ByVal cell2 As Range) As Range
        Get
            Return New Range(Me.underlyingComObject.Range(cell1.underlyingComObject, cell2.underlyingComObject))
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As String, ByVal cell2 As String) As Range
        Get
            Return New Range(Me.underlyingComObject.Range(cell1, cell2))
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As Range, ByVal cell2 As String) As Range
        Get
            Return New Range(Me.underlyingComObject.Range(cell1.underlyingComObject, cell2))
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As String, ByVal cell2 As Range) As Range
        Get
            Return New Range(Me.underlyingComObject.Range(cell1, cell2.underlyingComObject))
        End Get
    End Property

    Public ReadOnly Property Rows As Range
        Get
            Return New Range(Me.underlyingComObject.Rows)
        End Get
    End Property

    Public ReadOnly Property Names As Names
        Get
            Return New Names(Me.underlyingComObject.Names)
        End Get
    End Property

    Public ReadOnly Property Shapes As Shapes
        Get
            Return New Shapes(Me.underlyingComObject.Shapes)
        End Get
    End Property

    Public ReadOnly Property UsedRange As Range
        Get
            Return New Range(Me.underlyingComObject.UsedRange)
        End Get
    End Property

    Public ReadOnly Property Application As Application
        Get
            Return New Application(Me.underlyingComObject.Application)
        End Get
    End Property

    Public ReadOnly Property HPageBreaks As HPageBreaks
        Get
            Return New HPageBreaks(Me.underlyingComObject.HPageBreaks)
        End Get
    End Property

    Public ReadOnly Property VPageBreaks As VPageBreaks
        Get
            Return New VPageBreaks(Me.underlyingComObject.VPageBreaks)
        End Get
    End Property

    Public ReadOnly Property PivotTables As PivotTables
        Get
            Return New PivotTables(Me.underlyingComObject.PivotTables)
        End Get
    End Property

    Public Property DisplayPageBreaks As Boolean
        Get
            Return Me.underlyingComObject.DisplayPageBreaks
        End Get
        Set
            Me.underlyingComObject.DisplayPageBreaks = value
        End Set
    End Property

    Public ReadOnly Property ChartObjects As ChartObjects
        Get
            Return New ChartObjects(Me.underlyingComObject.ChartObjects)
        End Get
    End Property
End Class
