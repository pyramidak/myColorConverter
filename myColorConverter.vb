Class myColorConverter

#Region " Light or Dark Color "

    Public Shared Function LighterColor(inColor As Color, dFactor As Double) As System.Windows.Media.Color
        inColor.R = CByte(inColor.R + (255 - inColor.R) * dFactor)
        inColor.G = CByte(inColor.G + (255 - inColor.G) * dFactor)
        inColor.B = CByte(inColor.B + (255 - inColor.B) * dFactor)
        Return inColor
    End Function

    Public Shared Function DarkerColor(inColor As Color, dFactor As Double) As System.Windows.Media.Color
        inColor.R = CByte(inColor.R * (1 - dFactor))
        inColor.G = CByte(inColor.G * (1 - dFactor))
        inColor.B = CByte(inColor.B * (1 - dFactor))
        Return inColor
    End Function

#End Region

#Region " Drawing.Color and Media.Color "

    Public Shared Function ColorDrawingToMedia(dColor As System.Drawing.Color) As System.Windows.Media.Color
        Return System.Windows.Media.Color.FromArgb(dColor.A, dColor.R, dColor.G, dColor.B)
    End Function

    Public Shared Function ColorMediaToDrawing(mColor As System.Drawing.Color) As System.Drawing.Color
        Return System.Drawing.Color.FromArgb(mColor.A, mColor.R, mColor.G, mColor.B)
    End Function
#End Region

#Region " Brush and String "

    Public Shared Function BrushToColor(ByVal Brush As System.Windows.Media.Brush) As System.Windows.Media.Color
        Dim SCB As SolidColorBrush = DirectCast(Brush, SolidColorBrush)
        Return SCB.Color
    End Function

    Public Shared Function ColorToBrush(ByVal mColor As System.Windows.Media.Color) As System.Windows.Media.Brush
        Return New SolidColorBrush(mColor)
    End Function

    Public Shared Function BrushToString(ByVal Brush As System.Windows.Media.Brush) As String
        Return BrushToColor(Brush).ToString
    End Function

    Public Shared Function StringToBrush(ByVal sValue As String) As SolidColorBrush
        Try
            Dim mColor As System.Windows.Media.Color = CType(System.Windows.Media.ColorConverter.ConvertFromString(sValue), System.Windows.Media.Color)
            Return New SolidColorBrush(mColor)
        Catch ex As Exception
            Return New SolidColorBrush(Colors.Red)
        End Try
    End Function

#End Region

#Region " Convert Media.Colour to Integer "

    Public Shared Function ColorToInt(ByVal mColor As System.Windows.Media.Color) As Integer
        Dim dColor As System.Drawing.Color = System.Drawing.Color.FromArgb(mColor.A, mColor.R, mColor.G, mColor.B)
        Return System.Drawing.ColorTranslator.ToOle(dColor)
    End Function

    Public Shared Function BrushToInt(ByVal Brush As System.Windows.Media.Brush) As Integer
        Dim SCB As SolidColorBrush = DirectCast(Brush, SolidColorBrush)
        Return ColorToInt(SCB.Color)
    End Function

#End Region

#Region " Convert Integer to Media.Colour to String "

    Public Shared Function IntToColor(ByVal iColor As Integer) As System.Windows.Media.Color
        Dim winFormsColor As System.Drawing.Color = System.Drawing.ColorTranslator.FromOle(iColor)
        Return System.Windows.Media.Color.FromArgb(255, winFormsColor.B, winFormsColor.G, winFormsColor.R)
    End Function

    Public Shared Function IntToBrush(ByVal iColor As Integer) As SolidColorBrush
        Return New SolidColorBrush(IntToColor(iColor))
    End Function

    Public Shared Function IntToString(ByVal iColor As Integer) As String
        Return IntToColor(iColor).ToString
    End Function
#End Region

#Region " Get Color Name "

    Public Shared Function ColorToName(ByVal mColor As System.Windows.Media.Color) As String
        Dim clrKnownColor As System.Windows.Media.Color

        'Use reflection to get all known colors
        Dim ColorType As Type = GetType(System.Windows.Media.Colors)
        Dim arrPiColors As System.Reflection.PropertyInfo() = ColorType.GetProperties(System.Reflection.BindingFlags.[Public] Or System.Reflection.BindingFlags.[Static])

        'Iterate over all known colors, convert each to a <Color> and then compare
        'that color to the passed color.
        For Each pi As System.Reflection.PropertyInfo In arrPiColors
            clrKnownColor = DirectCast(pi.GetValue(Nothing, Nothing), System.Windows.Media.Color)
            If clrKnownColor = mColor Then
                Return pi.Name
            End If
        Next

        Return String.Empty
    End Function

    Public Shared Function NameToColor(ByVal sName As String) As System.Windows.Media.Color
        Dim mColor As System.Windows.Media.Color = Nothing
        Try
            Dim objValue As Object = System.Windows.Media.ColorConverter.ConvertFromString(sName)
            If (objValue IsNot Nothing) Then mColor = DirectCast(objValue, System.Windows.Media.Color)
            Return mColor
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function GetKnownColors() As List(Of KeyValuePair(Of String, System.Windows.Media.Color))
        Dim lst As New List(Of KeyValuePair(Of String, System.Windows.Media.Color))()
        Dim ColorType As Type = GetType(System.Windows.Media.Colors)
        Dim arrPiColors As System.Reflection.PropertyInfo() = ColorType.GetProperties(System.Reflection.BindingFlags.[Public] Or System.Reflection.BindingFlags.[Static])

        For Each pi As System.Reflection.PropertyInfo In arrPiColors
            lst.Add(New KeyValuePair(Of String, System.Windows.Media.Color)(pi.Name, DirectCast(pi.GetValue(Nothing, Nothing), System.Windows.Media.Color)))
        Next
        Return lst
    End Function

#End Region

End Class