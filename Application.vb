Imports JetBrains.Annotations

Public NotInheritable Class Application
    Private ReadOnly _app As Object

    Public Sub New()
        Try
            Dim typeExcel As Type = Type.GetTypeFromProgID("Excel.Application")
            Me._app = Activator.CreateInstance(typeExcel)
        Catch ex As Exception
            Throw New InvalidOperationException("It was not possible to start Excel - " & ex.Message, ex)
        End Try
    End Sub

    Friend Sub New(<NotNull> ByVal application As Object)
        Me._app = application
    End Sub

    Public Property Visible As Boolean
        Get
            Return Me._app.Visible
        End Get
        Set
            Me._app.visible = value
        End Set
    End Property

    Public ReadOnly Property Version As String
        Get
            Return Me._app.Version
        End Get
    End Property

    Public Sub Quit()
        Call Me._app.Quit()
    End Sub

    Public Sub Run(ByVal macroname As String)
        Call Me._app.Run(macroname)
    End Sub

    Public Sub Run(ByVal macroname As String, ByVal varg1 As Object)
        Call Me._app.Run(macroname, varg1)
    End Sub

    Public ReadOnly Property Workbooks As Workbooks
        Get
            Return New Workbooks(Me._app.Workbooks)
        End Get
    End Property

    Public Property Interactive As Boolean
        Get
            Return Me._app.Interactive
        End Get
        Set
            Me._app.Interactive = value
        End Set
    End Property

    Public Property DisplayAlerts As Boolean
        Get
            Return Me._app.DisplayAlerts
        End Get
        Set
            Me._app.DisplayAlerts = value
        End Set
    End Property

    Public Property WindowState As XlWindowState
        Get
            Return Me._app.WindowState
        End Get
        Set
            Me._app.WindowState = value
        End Set
    End Property

    Public Function Union(ByVal arg1 As Range, ByVal arg2 As Range) As Range
        Dim r1 As Object = arg1.underlyingComObject
        Dim r2 As Object = arg2.underlyingComObject
        Return New Range(Me._app.Union(r1, r2))
    End Function

    Public ReadOnly Property ActiveWindow As Window
        Get
            Dim comresult As Object = Me._app.ActiveWindow
            If comresult Is Nothing Then Return Nothing
            Return New Window(comresult)
        End Get
    End Property

    Public ReadOnly Property ActiveSheet As Sheet
        Get
            Return Sheet.CreateSheetObject(Me._app.ActiveSheet)
        End Get
    End Property

    Public ReadOnly Property Windows As Windows
        Get
            Return New Windows(Me._app.Windows)
        End Get
    End Property

    Public Property ScreenUpdating As Boolean
        Get
            Return Me._app.ScreeenUpdating
        End Get
        Set
            Me._app.ScreenUpdating = value
        End Set
    End Property

    Public Property ActivePrinter As String
        Get
            Return Me._app.ActivePrinter
        End Get
        Set
            Me._app.ActivePrinter = value
        End Set
    End Property

    Public Function CentimetersToPoints(ByVal centimeters As Double) As Double
        Return Me._app.CentimetersToPoints(centimeters)
    End Function

    Public ReadOnly Property Selection As Range
        Get
            Return New Range(Me._app.Selection)
        End Get
    End Property

    Public Property CutCopyMode As XlCutCopyMode
        Get
            Return Me._app.CutCopyMode
        End Get
        Set
            ' Be aware that it doesn't matter what value is used;
            ' setting CutCopyMode always cancels cut/copy mode and removes the moving border.
            Me._app.CutCopyMode = value
        End Set
    End Property

' ReSharper disable InconsistentNaming
    Public ReadOnly Property COMAddIns As COMAddIns
' ReSharper restore InconsistentNaming
        Get
            Return New COMAddIns(Me._app.COMAddIns)
        End Get
    End Property

    Public ReadOnly Property Hinstance As Integer
        Get
            Return Me._app.Hinstance
        End Get
    End Property

    Public Property PrintCommunication As Boolean
        Get
            If Decimal.Parse(Me.Version) < 14 Then Throw New NotImplementedException("Sorry, this property only became available in Excel 2010.")
            Return Me._app.PrintCommunication
        End Get
        Set
            If Decimal.Parse(Me.Version) < 14 Then Throw New NotImplementedException("Sorry, this property only became available in Excel 2010.")
            Me._app.PrintCommunication = value
        End Set
    End Property

    Public Function ConvertFormula(ByVal formula As String, ByVal fromReferenceStyle As XlReferenceStyle, ByVal toReferenceStyle As XlReferenceStyle) As String
        Dim result = Me._app.ConvertFormula(formula, fromReferenceStyle, toReferenceStyle)
        Return result
    End Function

    Public Overloads Overrides Function Equals(ByVal secondobject As Object) As Boolean
        Dim secondApplication = TryCast(secondobject, Application)
        Dim result = secondApplication Isnot Nothing andalso Me.Hinstance = secondApplication.Hinstance
        Return result
    End Function
End Class
