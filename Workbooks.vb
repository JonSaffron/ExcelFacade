' Workbooks is a 1 based collection of object Workbook
Imports JetBrains.Annotations

Public NotInheritable Class Workbooks
    Implements IEnumerable(Of Workbook)
    Implements IEnumerator(Of Workbook)
    Implements IDisposable

    Private ReadOnly _workbooks As Object

    Friend Sub New(<NotNull> ByVal workbooks As Object)
        Me._workbooks = workbooks
        Call Me.Reset()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return Me._workbooks.Count
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Workbook
        Get
            Return New Workbook(Me._workbooks.Item(index))
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal workbookname As String) As Workbook
        Get
            Return New Workbook(Me._workbooks.Item(workbookname))
        End Get
    End Property

    Public Function Open(ByVal filename As String) As Workbook
        Return New Workbook(Me._workbooks.Open(filename))
    End Function

    Public Function Open(ByVal filename As String, ByVal updatelinks As XlUpdateLinks) As Workbook
        Return New Workbook(Me._workbooks.Open(filename, updatelinks))
    End Function

    Public Function Open(ByVal filename As String, ByVal updatelinks As XlUpdateLinks, ByVal [readonly] As Boolean) As Workbook
        Return New Workbook(Me._workbooks.Open(filename, updatelinks, [readonly]))
    End Function

    Public Function Add() As Workbook
        Return New Workbook(Me._workbooks.Add())
    End Function

    Public Function Add(ByVal template As String) As Workbook
        Return New Workbook(Me._workbooks.Add(template))
    End Function

    Public Function Add(ByVal sheettype As XlWBATemplate) As Workbook
        Return New Workbook(Me._workbooks.Add(sheettype))
    End Function

#Region " IEnumerable implementation"
    Public Function GetEnumeratorOfWorkbook() As IEnumerator(Of Workbook) Implements IEnumerable(Of Workbook).GetEnumerator
        Return Me
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me
    End Function
#End Region

#Region " IEnumerator implementation"
    Private _enumeratorPostion As Integer

    Public ReadOnly Property CurrentOfWorkbook As Workbook Implements IEnumerator(Of Workbook).Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return Me.Item(Me._enumeratorPostion)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        Me._enumeratorPostion += 1
        Return (Me._enumeratorPostion <= Me.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        Me._enumeratorPostion = 0
    End Sub
#End Region

#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose
        ' nothing to do
    End Sub
#End Region

End Class
