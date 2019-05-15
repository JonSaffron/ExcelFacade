Imports System.Reflection
Imports JetBrains.Annotations

Public MustInherit Class Sheet
    Private ReadOnly _sheet As Object

    Friend Sub New(<NotNull> ByVal sheet As Object)
        Me._sheet = sheet
    End Sub

' ReSharper disable InconsistentNaming
    Friend ReadOnly Property underlyingComObject As Object
' ReSharper restore InconsistentNaming
        Get
            Return Me._sheet
        End Get
    End Property

    Friend Shared Function CreateSheetObject(ByVal underlyingComObject As Object) As Sheet
        Dim objectTypeName As String = GetComTypeName(underlyingComObject)
        Dim result As Sheet
        Select Case objectTypeName
            Case "Worksheet" : result = New Worksheet(underlyingComObject)
            Case "Chart" : result = New Chart(underlyingComObject)
            Case "DialogSheet" : result = New DialogSheet(underlyingComObject)
            Case Else : Throw New InvalidOperationException("Unknown type of excel object returned from sheets.item method.")
        End Select
        Return result
    End Function

    Public Property Name As String
        Get
            Return Me._sheet.Name
        End Get
        Set
            Me._sheet.Name = value
        End Set
    End Property

    Public ReadOnly Property Index As Integer
        Get
            Return Me._sheet.Index
        End Get
    End Property

    Public ReadOnly Property Parent As Workbook
        Get
            Return New Workbook(Me._sheet.Parent)
        End Get
    End Property

    Public Sub Activate()
        Call Me._sheet.Activate()
    End Sub

    Public Sub [Select]()
        Call Me.Select(True)
    End Sub

    Public Sub [Select](ByVal replace As Boolean)
        Call Me._sheet.[Select](replace)
    End Sub

    Public ReadOnly Property PageSetup As PageSetup
        Get
            Return New PageSetup(Me._sheet.PageSetup)
        End Get
    End Property

    Public Property Visible As XlSheetVisibility
        Get
            Return Me._sheet.Visible
        End Get
        Set
            Me._sheet.Visible = value
        End Set
    End Property

    Public Sub Move()
        Call Move(Nothing, Nothing)
    End Sub

    Public Sub Move(ByVal before As Sheet)
        Call Move(before, Nothing)
    End Sub

    Public Sub Move(ByVal before As Sheet, ByVal after As Sheet)
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
        Call Me._sheet.Move(b, a)
    End Sub

    Public Sub Copy()
        Call Copy(Nothing, Nothing)
    End Sub

    Public Sub Copy(ByVal before As Sheet)
        Call Copy(before, Nothing)
    End Sub

    Public Sub Copy(ByVal before As Sheet, ByVal after As Sheet)
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
        Call Me._sheet.Copy(b, a)
    End Sub

    Public Sub Delete()
        Call Me._sheet.Delete()
    End Sub

    Public Overloads Overrides Function Equals(ByVal secondobject As Object) As Boolean
        Dim secondsheet = TryCast(secondobject, Sheet)
        dim result = secondsheet IsNot Nothing andalso Me.Parent.Equals(secondsheet.Parent) AndAlso Me.Name = secondsheet.Name
        Return result
    End Function
End Class
