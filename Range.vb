Imports System.Runtime.InteropServices
Imports JetBrains.Annotations

Public NotInheritable Class Range
    Implements IEnumerable(Of Range)
    Implements IEnumerator(Of Range)
    Implements IDisposable

    Private ReadOnly _range As Object

    Friend Sub New(<NotNull> ByVal range As Object)
        Me._range = range
        Call Me.Reset()
    End Sub

' ReSharper disable InconsistentNaming
    Friend ReadOnly Property underlyingComObject As Object
' ReSharper restore InconsistentNaming
        Get
            Return Me._range
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Range
        Get
            Return New Range(Me._range.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal rowindex As Integer, ByVal columnindex As Integer) As Range
        Get
            Return New Range(Me._range.Item(rowindex, columnindex))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal rowindex As Integer, ByVal columnName As String) As Range
        Get
            Return New Range(Me._range.Item(rowindex, columnName))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As String) As Range
        Get
            Return New Range(Me._range.Item(index))
        End Get
    End Property

    Public ReadOnly Property Cells As Range
        Get
            Return New Range(Me._range.Cells)
        End Get
    End Property

    Public ReadOnly Property Columns As Range
        Get
            Return New Range(Me._range.Columns)
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As String) As Range
        Get
            Return New Range(Me._range.Range(cell1))
        End Get
    End Property

    Public ReadOnly Property Range(ByVal cell1 As String, ByVal cell2 As String) As Range
        Get
            Return New Range(Me._range.Range(cell1, cell2))
        End Get
    End Property

    Public ReadOnly Property Rows As Range
        Get
            Return New Range(Me._range.Rows)
        End Get
    End Property

    Public ReadOnly Property Column As Integer
        Get
            Return Me._range.Column
        End Get
    End Property

    Public ReadOnly Property Row As Integer
        Get
            Return Me._range.Row
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return Me._range.Count
        End Get
    End Property

    Public ReadOnly Property EntireRow As Range
        Get
            Return New Range(Me._range.EntireRow)
        End Get
    End Property

    Public ReadOnly Property EntireColumn As Range
        Get
            Return New Range(Me._range.EntireColumn)
        End Get
    End Property

    Public Property ColumnWidth As Double
        Get
            Return Me._range.ColumnWidth
        End Get
        Set
            Me._range.ColumnWidth = value
        End Set
    End Property

    Public Property RowHeight As Double
        Get
            Return Me._range.RowHeight
        End Get
        Set
            Me._range.RowHeight = value
        End Set
    End Property

    Public Property Value As Object
        Get
            Return Me._range.Value
        End Get
        Set
            Me._range.Value = value
        End Set
    End Property

    Public Property Formula As Object
        Get
            Return Me._range.Formula
        End Get
        Set
            Me._range.Formula = value
        End Set
    End Property

    Public ReadOnly Property Characters As Characters
        Get
            Return New Characters(Me._range.Characters)
        End Get
    End Property

    Public ReadOnly Property Characters(ByVal start As Integer) As Characters
        Get
            Return New Characters(Me._range.Characters(start))
        End Get
    End Property

    Public ReadOnly Property Characters(ByVal start As Integer, ByVal length As Integer) As Characters
        Get
            Return New Characters(Me._range.Characters(start, length))
        End Get
    End Property

    Public ReadOnly Property Font As Font
        Get
            Return New Font(Me._range.Font)
        End Get
    End Property

    Public ReadOnly Property Errors As Errors
        Get
            Return New Errors(Me._range.Errors)
        End Get
    End Property

    Public Sub Clear()
        Call Me._range.Clear()
    End Sub

    Public Sub ClearContents()
        Call Me._range.ClearContents()
    End Sub

    Public Function Insert() As Boolean
        Return Me._range.Insert()
    End Function

    Public Function Insert(ByVal shift As XlInsertShiftDirection) As Boolean
        Return Me._range.Insert(shift)
    End Function

    Public Function Delete() As Boolean
        Return Me._range.Delete()
    End Function

    Public Function Delete(ByVal shift As XlDeleteShiftDirection) As Boolean
        Return Me._range.Delete(shift)
    End Function

    Public Sub Copy()
        Call Me._range.Copy()
    End Sub

    Public Sub Copy(ByVal destination As Range)
        Call Me._range.Copy(destination.underlyingComObject)
    End Sub

    Public ReadOnly Property Parent As Worksheet
        Get
            Return New Worksheet(Me._range.Parent)
        End Get
    End Property

    Public ReadOnly Property Left As Double
        Get
            Return Me._range.Left
        End Get
    End Property

    Public ReadOnly Property Top As Double
        Get
            Return Me._range.Top
        End Get
    End Property

    Public ReadOnly Property Width As Double
        Get
            Return Me._range.Width
        End Get
    End Property

    Public ReadOnly Property Height As Double
        Get
            Return Me._range.Height
        End Get
    End Property

    Public Function [Select]() As Boolean
        ' Be aware: You cannot select a range in a sheet that is not the active one
        Return Me._range.Select()
    End Function

    Public ReadOnly Property Application As Application
        Get
            Return New Application(Me._range.Application)
        End Get
    End Property

    Public Property VerticalAlignment As XlVAlign?
        Get
            Return ToNullable(of XlVAlign)(Me._range.VerticalAlignment)
        End Get
        Set
            Me._range.VerticalAlignment = value.NullableToNull()
        End Set
    End Property

    Public Property HorizontalAlignment As XlHAlign?
        Get
            Return ToNullable(Of XlHAlign)(Me._range.HorizontalAlignment)
        End Get
        Set
            Me._range.HorizontalAlignment = value.NullableToNull()
        End Set
    End Property

    Public Sub CopyFromRecordset(ByVal data As Object)
        If data Is Nothing Then
            Throw New ArgumentNullException(NameOf(data))
        End If

' ReSharper disable VBPossibleMistakenCallToGetType.2
        If data.GetType().FullName <> "ADODB.RecordsetClass" Then
' ReSharper restore VBPossibleMistakenCallToGetType.2
            Throw New ArgumentException("CopyFromRecordset only accepts an ADODB.Recordset as a source of data.", "data")
        End If

        ' workaround for a problem in excel 2003 sp3 where it auto formats cells in the active sheet rather than the sheet where the data is being placed
        Dim currentlyActiveSheet As Sheet = Me.Application.ActiveSheet
        Dim swappingSheets As Boolean = False
        If Not Me.Worksheet.Equals(currentlyActiveSheet) Then
            swappingSheets = True
            Call Me.Worksheet.Activate()
        End If

        ' There are many reasons why the call to CopyFromRecordset might fail:
        ' - an excel cell cannot contain more than around 900 characters.
        ' - a decimal value runs to too many decimal places. excel only seems to cope with a maximum of 4 dp.
        ' - Excel is busy doing something else (which is why it's worth disabling COM Add Ins)
        Try
            Call Me._range.CopyFromRecordset(data)
        Catch ex As COMException
            Dim msg As String = String.Format("CopyFromRecordset failed. {0}", ex.Message)
            Throw New InvalidOperationException(msg, ex)
        End Try

        If swappingSheets Then
            Call currentlyActiveSheet.Activate()
        End If
    End Sub

    Public ReadOnly Property Worksheet As Worksheet
        Get
            Return New Worksheet(Me._range.Worksheet)
        End Get
    End Property

    Public ReadOnly Property Interior As Interior
        Get
            Return New Interior(Me._range.Interior)
        End Get
    End Property

    Public Property NumberFormat As String
        Get
            Dim returnValue As Object = Me._range.NumberFormat
            Return If(typeof returnValue Is DBNull, Nothing, returnValue)
        End Get
        Set
            Me._range.NumberFormat = If(value Is Nothing, DBNull.Value, value)
        End Set
    End Property

    Public Property WrapText As Boolean?
        Get
            Return ToNullable(of boolean)(Me._range.WrapText)
        End Get
        Set
            Me._range.WrapText = value.NullableToNull()
        End Set
    End Property

    public Property ShrinkToFit As Boolean?
        get
            Return ToNullable(of Boolean)(Me._range.ShrinkToFit)
        End Get
        Set
            Me._range.ShrinkToFit = value.NullableToNull()
        End Set
    End Property

    Public Sub SetOrientation(byval newOrientation as XlOrientation)
        Me._range.Orientation = newOrientation
    end Sub

    public Property Orientation As Integer?
        get
            return ToNullable(of Integer)(Me._range.Orientation)
        End Get
        Set
            Me._range.Orientation = value.NullableToNull()
        End Set
    End Property

    Public ReadOnly Property Borders As Borders
        Get
            Return New Borders(Me._range.Borders)
        End Get
    End Property

    Public Property MergeCells As Boolean?
        Get
            Return ToNullable(Of Boolean)(Me._range.MergeCells)
        End Get
        Set
            Me._range.MergeCells = value.NullableToNull()
        End Set
    End Property

    Public Sub Merge()
        Call Me._range.Merge()
    End Sub

    Public Sub Merge(ByVal across As Boolean)
        Call Me._range.Merge(across)
    End Sub

    Public ReadOnly Property Text As String
        Get
            Dim returnValue as Object = Me._range.Text
            Return If(typeof returnValue Is DBNull, Nothing, returnValue)
        End Get
    End Property

    Public Property Hidden As Boolean
        Get
            Return Me._range.Hidden
        End Get
        Set
            Me._range.Hidden = value
        End Set
    End Property

    Public Property PageBreak As XlPageBreak
        Get
            Return Me._range.PageBreak
        End Get
        Set
            Me._range.PageBreak = value
        End Set
    End Property

    Public Sub PasteSpecial()
        Call Me._range.PasteSpecial()
    End Sub

    Public Sub PasteSpecial(ByVal paste As XlPasteType)
        Call Me._range.PasteSpecial(paste)
    End Sub

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfRange() As IEnumerator(Of Range) Implements IEnumerable(Of Range).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer
    Private _items() As Range

    Public ReadOnly Property CurrentOfRange As Range Implements IEnumerator(Of Range).Current
        Get
            Return _items(_enumeratorPosition)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return _items(_enumeratorPosition)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        If _enumeratorPosition = 0 Then
            ReDim Me._items(Me.Count + 1)
            Dim i As Integer = 1
            For Each innerRange As Object In Me._range
                Dim r As Range = New Range(innerRange)
                _items(i) = r
                i += 1
            Next
        End If
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
