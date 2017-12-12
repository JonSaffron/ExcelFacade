' Sheets is a 1 based collection of Worksheet, Chart and DialogSheet objects
Imports System.Reflection
Imports JetBrains.Annotations

Public Class Sheets
    Implements IEnumerable(Of Sheet)
    Implements IEnumerator(Of Sheet)
    Implements IDisposable

' ReSharper disable InconsistentNaming
    Protected ReadOnly _sheets As Object
' ReSharper restore InconsistentNaming

    Friend Sub New(<NotNull> ByVal sheets As Object)
        Me._sheets = sheets
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._sheets.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Sheet
        Get
            Dim underlyingComObject As Object = Me._sheets.Item(index)
            Dim result = Sheet.CreateSheetObject(underlyingComObject)
            Return result
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal sheetname As String) As Sheet
        Get
            Dim underlyingComObject As Object = Me._sheets.Item(sheetname)
            Dim result = Sheet.CreateSheetObject(underlyingComObject)
            Return result
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal sheetnames As String()) As Sheets
        Get
            Dim underlyingComObject As Object = Me._sheets.Item(sheetnames)
            Dim result As Sheets = New Sheets(underlyingComObject)
            Return result
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal sheets As Integer()) As Sheets
        Get
            Dim underlyingComObject As Object = Me._sheets.Item(sheets)
            Dim result As Sheets = New Sheets(underlyingComObject)
            Return result
        End Get
    End Property

    Public Sub [Select]()
        Call Me.Select(True)
    End Sub

    Public Sub [Select](ByVal replace As Boolean)
        Call Me._sheets.[Select](replace)
    End Sub

    Public Overridable Function Add() As Sheet
        Return Add(Nothing, Nothing, 1, XlSheetType.xlWorksheet)
    End Function

    Public Overridable Function Add(ByVal before As Sheet) As Sheet
        Return Add(before, Nothing, 1, XlSheetType.xlWorksheet)
    End Function

    Public Overridable Function Add(ByVal before As Sheet, ByVal after As Sheet)
        Return Add(before, after, 1, XlSheetType.xlWorksheet)
    End Function

    Public Overridable Function Add(ByVal before As Sheet, ByVal after As Sheet, ByVal countOfSheets As Integer)
        Return Add(before, after, countOfSheets, XlSheetType.xlWorksheet)
    End Function

    Public Overridable Function Add(ByVal before As Sheet, ByVal after As Sheet, ByVal countOfSheets As Integer, ByVal type As XlSheetType) As Sheet
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
        Return Sheet.CreateSheetObject(Me._sheets.Add(b, a, countOfSheets, type))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfType() As IEnumerator(Of Sheet) Implements IEnumerable(Of Sheet).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPosition As Integer

    Public ReadOnly Property CurrentOfType As Sheet Implements IEnumerator(Of Sheet).Current
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
