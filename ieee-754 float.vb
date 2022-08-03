' Writen by Selene Tabacchini (2022.07.15)
'
' Converts a binary string into an IEEE-754 floating point data type

Module ieee_754_float
    Public Function int2float(ByVal iInteger As Int32) As Double
        ' Convert the integer to a binary string
        Dim sBinary As String = Convert.ToString(iInteger, 2)
        ' Make sure binary string is padded to 32 characters
        fixBinaryLength(sBinary)
        ' Use the bin2float function to complete the rest
        int2float = bin2float(sBinary)
    End Function

    Public Function bin2float(ByVal sBinary As String) As Double
        ' Converts binary string into IEEE-754 float value

        ' Prepare the string variables for the whole numbers and fraction
        Dim sWhole As String = "0"
        Dim sFraction As String = String.Empty

        ' Find the decimal placement
        Dim decPlace As Integer = CInt(Convert.ToByte(Mid(sBinary, 2, 8), 2) - 127)
        If decPlace < 0 Then
            ' With a negative decimal placement, pad the fraction with 0's before the starting 1
            sFraction = New String(CChar("0"), Math.Abs(decPlace) - 1) & "1" & Mid(sBinary, 10)
        Else
            ' With non-negative decimal placement, place the starting 1 and grab the whole numbers part of the binary string (if any)
            sWhole = "1" & Mid(sBinary, 10, decPlace)
            ' The rest of the binary string is for the fraction
            sFraction = Mid(sBinary, 10 + decPlace)
        End If

        ' Start of with the 1 or 0
        bin2float = CDbl(Mid(sWhole, 1, 1))

        ' Step through the whole number string, doubling at every binary digit, increasing by 1 if it is a 1
        ' Note we start at 2 since we already grabbed the starting digit
        For i = 2 To sWhole.Length
            bin2float = bin2float * 2
            If Mid(sWhole, i, 1) = "1" Then bin2float = bin2float + 1
        Next

        ' Step through the fraction number string, adding the shrinking fraction value for each 1
        Dim dFraction As Double = 1.0
        For i = 1 To sFraction.Length
            dFraction = dFraction / 2
            If Mid(sFraction, i, 1) = "1" Then bin2float = bin2float + dFraction
        Next

        ' Finally, check the first signed flag to see if it is a negative value
        If Mid(sBinary, 1, 1) = "1" Then bin2float = bin2float * -1
    End Function

    Public Sub fixBinaryLength(ByRef bin1 As String, Optional ByRef bin2 As String = "", Optional ByRef bin3 As String = "")
        While bin1.Length < 32
            bin1 = "0" & bin1
        End While
        If bin2 = "" Then Exit Sub
        While bin2.Length < 32
            bin2 = "0" & bin2
        End While
        If bin3 = "" Then Exit Sub
        While bin3.Length < 32
            bin3 = "0" & bin3
        End While
    End Sub
End Module