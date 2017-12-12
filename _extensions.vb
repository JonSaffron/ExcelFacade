Imports System.Drawing
Imports System.Runtime.CompilerServices

friend module Extensions
    public Function ToNullable(of T As structure)(ByVal value As object) As Nullable(Of T)
        Return If(TypeOf value Is DBNull, Nothing, New T?(value))
    End Function

    <Extension>
    Public Function NullableToNull(Of T As structure)(byval value As Nullable(of T)) As Object
        Return If(value, DBNull.Value)
    End Function

    public Function ToColor(byval value As object) As Color?
        If TypeOf value Is DBNull Then Return Nothing
        Dim colourValue = Convert.ToUInt32(value)
        Dim red As byte = colourValue and byte.MaxValue
        Dim green As byte = (colourValue >> 8) and byte.MaxValue
        Dim blue as Byte = (colourValue >> 16) And byte.MaxValue
        Return Color.FromArgb(255, red, green, blue)
    End Function

    <Extension>
    Public Function ToVbaColor(byval color As Color?) As Object
        if Not color.HasValue then Return DBNull.Value
        return (color.Value.B << 16) or (color.Value.G << 8) or color.Value.R
    End Function
End module
