Module ModulePrint

    Public Function getXInteger(iXBegin As Integer, iNumber As Integer) As Integer
        If ((iNumber >= 10) And (iNumber < 100)) Then
            iXBegin = iXBegin - 10
        ElseIf ((iNumber >= 100) And (iNumber < 1000)) Then
            iXBegin = iXBegin - 20
        ElseIf ((iNumber >= 1000) And (iNumber < 10000)) Then
            iXBegin = iXBegin - 30
        ElseIf (iNumber >= 10000) Then
            iXBegin = iXBegin - 40
        End If
        Return iXBegin
    End Function

    Public Function getXCurrency(ByVal iXBegin As Integer, ByVal dbTemp As Double) As Integer

        If ((dbTemp >= 1.0) And (dbTemp < 10.0)) Then
            iXBegin = iXBegin + 13
        ElseIf ((dbTemp >= 100.0) And (dbTemp < 1000.0)) Then
            iXBegin = iXBegin - 10
        ElseIf ((dbTemp >= 1000.0) And (dbTemp < 10000.0)) Then
            iXBegin = iXBegin - 30
        ElseIf ((dbTemp >= 10000.0) And (dbTemp < 100000.0)) Then
            iXBegin = iXBegin - 42
        ElseIf ((dbTemp >= 100000.0) And (dbTemp < 1000000.0)) Then
            iXBegin = iXBegin - 55
        ElseIf ((dbTemp >= 1000000.0) And (dbTemp < 10000000.0)) Then
            iXBegin = iXBegin - 75
        ElseIf ((dbTemp >= 10000000.0) And (dbTemp < 100000000.0)) Then
            iXBegin = iXBegin - 87
        ElseIf ((dbTemp >= 100000000.0) And (dbTemp < 1000000000.0)) Then
            iXBegin = iXBegin - 99
        ElseIf ((dbTemp >= 1000000000.0) And (dbTemp < 10000000000.0)) Then
            iXBegin = iXBegin - 118
        End If
        Return iXBegin
    End Function

    Public Function getStringCurrency(ByVal dbCurrency As Double) As String
        Return String.Format("{0:N}", dbCurrency)
    End Function

    Public Function getXNumber(iXBegin As Integer, iNumber As Integer) As Integer
        If iNumber < 10 Then
            iXBegin = iXBegin + 2
        ElseIf ((iNumber >= 10) And (iNumber < 100)) Then
            iXBegin = iXBegin - 10
        ElseIf ((iNumber >= 100) And (iNumber < 1000)) Then
            iXBegin = iXBegin - 20
        ElseIf ((iNumber >= 1000) And (iNumber < 10000)) Then
            iXBegin = iXBegin - 38
        ElseIf (iNumber >= 10000) And (iNumber < 100000) Then
            iXBegin = iXBegin - 50
        ElseIf (iNumber >= 100000) And (iNumber < 1000000) Then
            iXBegin = iXBegin - 62
        End If
        Return iXBegin
    End Function

    Public Function getStringNumber(ByVal iNumber As Integer) As String

        If iNumber < 10 Then
            Return iNumber.ToString()
        End If

        Dim dbTemp As Double = 0D
        dbTemp = dbTemp + iNumber

        Return String.Format("{0:0,0}", dbTemp)
    End Function
End Module
