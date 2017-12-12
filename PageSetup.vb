Imports JetBrains.Annotations

Public NotInheritable Class PageSetup
    Private ReadOnly _pagesetup As Object

    Friend Sub New(<NotNull> ByVal pagesetup As Object)
        Me._pagesetup = pagesetup
    End Sub

    Public Property LeftFooter As String
        Get
            Return Me._pagesetup.LeftFooter
        End Get
        Set
            Me._pagesetup.LeftFooter = value
        End Set
    End Property

    Public Property CenterFooter As String
        Get
            Return Me._pagesetup.CenterFooter
        End Get
        Set
            Me._pagesetup.CenterFooter = value
        End Set
    End Property

    Public Property RightFooter As String
        Get
            Return Me._pagesetup.RightFooter
        End Get
        Set
            Me._pagesetup.RightFooter = value
        End Set
    End Property

    Public Property LeftMargin As Double
        Get
            Return Me._pagesetup.LeftMargin
        End Get
        Set
            Me._pagesetup.LeftMargin = value
        End Set
    End Property

    Public Property RightMargin As Double
        Get
            Return Me._pagesetup.RightMargin
        End Get
        Set
            Me._pagesetup.RightMargin = value
        End Set
    End Property

    Public Property TopMargin As Double
        Get
            Return Me._pagesetup.TopMargin
        End Get
        Set
            Me._pagesetup.TopMargin = value
        End Set
    End Property

    Public Property BottomMargin As Double
        Get
            Return Me._pagesetup.BottomMargin
        End Get
        Set
            Me._pagesetup.BottomMargin = value
        End Set
    End Property

    Public Property HeaderMargin As Double
        Get
            Return Me._pagesetup.HeaderMargin
        End Get
        Set
            Me._pagesetup.HeaderMargin = value
        End Set
    End Property

    Public Property FooterMargin As Double
        Get
            Return Me._pagesetup.FooterMargin
        End Get
        Set
            Me._pagesetup.FooterMargin = value
        End Set
    End Property

    Public Property PrintTitleRows As String
        Get
            Return Me._pagesetup.PrintTitleRows
        End Get
        Set
            Me._pagesetup.PrintTitleRows = value
        End Set
    End Property

    Public Property PrintTitleColumns As String
        Get
            Return Me._pagesetup.PrintTitleColumns
        End Get
        Set
            Me._pagesetup.PrintTitleColumns = value
        End Set
    End Property

    Public Property PrintGridlines As Boolean
        Get
            Return Me._pagesetup.PrintGridlines
        End Get
        Set
            Me._pagesetup.PrintGridlines = value
        End Set
    End Property

    Public Property Orientation As XlPageOrientation
        Get
            Return Me._pagesetup.Orientation
        End Get
        Set
            Me._pagesetup.Orientation = value
        End Set
    End Property

    Public Property Draft As Boolean
        Get
            Return Me._pagesetup.Draft
        End Get
        Set
            Me._pagesetup.Draft = value
        End Set
    End Property

    ' Set to nothing to allow FitToPagesWide and FitToPagesTall to do their work
    Public Property Zoom As Integer?
        Get
            Dim returnValue As Object = Me._pagesetup.Zoom
            Return If(TypeOf returnValue Is Boolean AndAlso returnValue = False, nothing, Convert.ToInt32(returnValue))
        End Get
        Set
            Me._pagesetup.Zoom = If(value.HasValue, Value.Value, False)
        End Set
    End Property

    Public Property FitToPagesWide As Integer?
        Get
            Dim returnValue As Object = Me._pagesetup.FitToPagesWide
            Return If(TypeOf returnValue Is Boolean AndAlso returnValue = false, nothing, Convert.ToInt32(returnValue))
        End Get
        Set
            Me._pagesetup.FitToPagesWide = If(value.HasValue, Value.Value, False)
        End Set
    End Property

    Public Property FitToPagesTall As Integer?
        Get
            Dim returnValue As Object = Me._pagesetup.FitToPagesTall
            Return If(TypeOf returnValue Is Boolean AndAlso returnValue = false, nothing, Convert.ToInt32(returnValue))
        End Get
        Set
            Me._pagesetup.FitToPagesTall = If(value.HasValue, Value.Value, False)
        End Set
    End Property

    Public Property CenterHorizontally As Boolean
        Get
            Return Me._pagesetup.CenterHorizontally
        End Get
        Set
            Me._pagesetup.CenterHorizontally = value
        End Set
    End Property

    Public Property CenterVertically As Boolean
        Get
            Return Me._pagesetup.CenterVertically
        End Get
        Set
            Me._pagesetup.CenterVertically = value
        End Set
    End Property

    Public Property LeftHeader As String
        Get
            Return Me._pagesetup.LeftHeader
        End Get
        Set
            Me._pagesetup.LeftHeader = value
        End Set
    End Property

    Public Property CenterHeader As String
        Get
            Return Me._pagesetup.CenterHeader
        End Get
        Set
            Me._pagesetup.CenterHeader = value
        End Set
    End Property

    Public Property RightHeader As String
        Get
            Return Me._pagesetup.RightHeader
        End Get
        Set
            Me._pagesetup.RightHeader = value
        End Set
    End Property

    Public Property PaperSize As XlPaperSize
        Get
            Return Me._pagesetup.PaperSize
        End Get
        Set
            Me._pagesetup.PaperSize = value
        End Set
    End Property
End Class
