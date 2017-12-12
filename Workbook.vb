Imports JetBrains.Annotations

Public NotInheritable Class Workbook
    Private ReadOnly _workbook As Object

    Friend Sub New(<NotNull> ByVal workbook As Object)
        Me._workbook = workbook
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return Me._workbook.Name
        End Get
    End Property

    Public ReadOnly Property Worksheets As Worksheets
        Get
            Return New Worksheets(Me._workbook.Worksheets)
        End Get
    End Property

    Public ReadOnly Property DialogSheets As DialogSheets
        Get
            Return New DialogSheets(Me._workbook.DialogSheets)
        End Get
    End Property

    Public ReadOnly Property Charts As Charts
        Get
            Return New Charts(Me._workbook.Charts)
        End Get
    End Property

    Public ReadOnly Property Sheets As Sheets
        Get
            Return New Sheets(Me._workbook.Sheets)
        End Get
    End Property

    Public ReadOnly Property PivotCaches As PivotCaches
        Get
            Return New PivotCaches(Me._workbook.PivotCaches)
        End Get
    End Property

    Public Property ShowPivotTableFieldList As Boolean
        Get
            Return Me._workbook.ShowPivotTableFieldList
        End Get
        Set
            Me._workbook.ShowPivotTableFieldList = value
        End Set
    End Property

    Public Sub Save()
        Call Me._workbook.Save()
    End Sub

    Public Sub SaveCopyAs(ByVal filename As String)
        Call Me._workbook.SaveCopyAs(filename)
    End Sub

    <Obsolete("Use override specifying fileformat. In Excel 2007 the file extension must be appropriate to the file format.")> _
    Public Sub SaveAs(ByVal filename As String)
        Call Me._workbook.SaveAs(filename)
    End Sub

    Public Sub SaveAs(ByVal filename As String, ByVal format As XlFileFormat)
        Call Me._workbook.SaveAs(filename, format)
    End Sub

    Public Sub Close()
        Call Me._workbook.Close()
    End Sub

    Public Sub Close(ByVal savechanges As Boolean)
        Call Me._workbook.close()
    End Sub

    Public Property Saved As Boolean
        Get
            Return Me._workbook.Saved
        End Get
        Set
            Me._workbook.Saved = value
        End Set
    End Property

    Public ReadOnly Property Names As Names
        Get
            Return New Names(Me._workbook.Names)
        End Get
    End Property

    Public ReadOnly Property Windows As Windows
        Get
            Return New Windows(Me._workbook.Windows)
        End Get
    End Property

    Public ReadOnly Property Application As Application
        Get
            Return New Application(Me._workbook.Application)
        End Get
    End Property

    Public ReadOnly Property BuiltInDocumentProperties As DocumentProperties
        Get
            Return New DocumentProperties(Me._workbook.BuiltInDocumentProperties)
        End Get
    End Property

    Public ReadOnly Property CustomDocumentProperties As DocumentProperties
        Get
            Return New DocumentProperties(Me._workbook.CustomDocumentProperties)
        End Get
    End Property

    Public ReadOnly Property ActiveSheet As Sheet
        Get
            Return Sheet.CreateSheetObject(Me._workbook.ActiveSheet)
        End Get
    End Property

    Public ReadOnly Property FullName As String
        Get
            Return Me._workbook.FullName
        End Get
    End Property

    Public ReadOnly Property Path As String
        Get
            Return Me._workbook.Path
        End Get
    End Property

    Public ReadOnly Property FileFormat As XlFileFormat
        Get
            Return Me._workbook.FileFormat
        End Get
    End Property

    Public Overloads Overrides Function Equals(ByVal secondobject As Object) As Boolean
        Dim secondworkbook = TryCast(secondobject, Workbook)
        Dim result = secondworkbook IsNot Nothing AndAlso Me.Application.Equals(secondworkbook.Application) AndAlso Me.Name = secondworkbook.Name
        return result
    End Function
End Class
