Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Friend Module Extensions
    Public Function ToNullable(Of T As Structure)(ByVal value As Object) As Nullable(Of T)
        Return If(TypeOf value Is DBNull, Nothing, New T?(value))
    End Function

    <Extension>
    Public Function NullableToNull(Of T As Structure)(ByVal value As Nullable(Of T)) As Object
        Return If(value, DBNull.Value)
    End Function

    Public Function ToColor(ByVal value As Object) As Color?
        If TypeOf value Is DBNull Then Return Nothing
        Dim colourValue = Convert.ToUInt32(value)
        Dim red As Byte = colourValue And Byte.MaxValue
        Dim green As Byte = (colourValue >> 8) And Byte.MaxValue
        Dim blue As Byte = (colourValue >> 16) And Byte.MaxValue
        Return Color.FromArgb(255, red, green, blue)
    End Function

    <Extension>
    Public Function ToVbaColor(ByVal color As Color?) As Object
        If Not color.HasValue Then Return DBNull.Value
        Return (color.Value.B << 16) Or (color.Value.G << 8) Or color.Value.R
    End Function

    ' https://stackoverflow.com/questions/1429548/how-to-get-type-of-com-object

    Public Function GetComTypeName(ByVal comObject As Object) As String
        Dim dispatch As IDispatch = TryCast(comObject, IDispatch)

        If dispatch Is Nothing Then
            Return Nothing
        End If

        Dim pTypeInfo = dispatch.GetTypeInfo(0, 1033)

        Dim pBstrName As String = String.Empty
        Dim pBstrDocString As String = String.Empty
        Dim pdwHelpContext As Integer
        Dim pBstrHelpFile As String = String.Empty
        Call pTypeInfo.GetDocumentation(-1, pBstrName, pBstrDocString, pdwHelpContext, pBstrHelpFile)

        Dim str As String = pBstrName
        If str(0) = "_"c Then
            ' remove leading '_'
            str = str.Substring(1)
        End If

        Return str
    End Function

    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("00020400-0000-0000-C000-000000000046")>
    Private Interface IDispatch
        Function GetTypeInfoCount() As Integer

        Function GetTypeInfo(
            <[In], MarshalAs(UnmanagedType.U4)> iTInfo As Integer,
            <[In], MarshalAs(UnmanagedType.U4)> lcid As Integer) _
                As <MarshalAs(UnmanagedType.Interface)> ITypeInfo

        Sub GetIDsOfNames(
            <[In]> ByRef riid As Guid,
            <[In], MarshalAs(UnmanagedType.LPArray)> rgszNames As String(),
            <[In], MarshalAs(UnmanagedType.U4)> cNames As Integer,
            <[In], MarshalAs(UnmanagedType.U4)> lcid As Integer,
            <Out, MarshalAs(UnmanagedType.LPArray)> rgDispId As Integer())
    End Interface
End Module
