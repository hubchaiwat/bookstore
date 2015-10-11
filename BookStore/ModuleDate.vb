Module ModuleDate
    'Function สำหรับสร้างวันที่ในการพิมพ์
    Function makeDatePrint(ByVal intDay As Integer, ByVal intMonth As Integer, ByVal intYear As Integer) As String
        Dim strDate As String = ""

        'สร้างวันที่
        If intDay < 10 Then
            strDate = strDate + "0"
        End If

        strDate = strDate + intDay.ToString() + "/"

        'สร้างเดือน
        If intMonth < 10 Then
            strDate = strDate + "0"
        End If

        strDate = strDate + intMonth.ToString() + "/"

        'สร้างปี
        intYear = intYear + 543
        If intYear < 10 Then
            strDate = strDate + "000"
        ElseIf intYear < 100 Then
            strDate = strDate + "00"
        ElseIf intYear < 1000 Then
            strDate = strDate + "0"
        End If

        strDate = strDate + intYear.ToString()

        Return strDate
    End Function

    Function makeDateRentAndReturn(ByVal intDay As Integer, ByVal intMonth As Integer, ByVal intYear As Integer) As String
        Dim strDate As String = ""

        'สร้างวันที่
        If intDay < 10 Then
            strDate = strDate + "0"
        End If

        strDate = strDate + intDay.ToString() + "/"

        'สร้างเดือน
        If intMonth < 10 Then
            strDate = strDate + "0"
        End If

        strDate = strDate + intMonth.ToString() + "/"

        'สร้างปี
        intYear = intYear + 543
        Dim strYear As String = ""
        If intYear < 10 Then
            strYear = strYear + "000"
        ElseIf intYear < 100 Then
            strYear = strYear + "00"
        ElseIf intYear < 1000 Then
            strYear = strYear + "0"
        End If

        strYear = strYear + intYear.ToString()
        strYear = strYear.Substring(2, 2)

        strDate = strDate + strYear

        Return strDate
    End Function

    'Function สำหรับสร้างข้อความวันที่เพื่อ Transaction
    Function makeDateTran(ByVal intDay As Integer, ByVal intMonth As Integer, ByVal intYear As Integer) As String
        Dim strDate As String = ""

        strDate = strDate + "#" + intYear.ToString() + "/" + intMonth.ToString() + "/" + intDay.ToString() + "#"

        Return strDate
    End Function

    'Function สำหรับสกัดวันที่ชนิด Integer ออกจากข้อความวันที่ชนิด String
    Function getDayFromTextDate(ByVal strDate As String) As Integer
        Dim intDay As Integer = 0
        Dim strDay As String = ""

        'สกัดวันที่ออกจากข้อความ
        strDay = strDate.Substring(0, 2)

        'แปลงค่าจากตัวอักษรเป็นตัวเลขให้กับ intDay
        intDay = Integer.Parse(strDay)

        Return intDay
    End Function

    'Function สำหรับสกัดเดือนชนิด Integer ออกจากข้อความวันที่ชนิด String
    Function getMonthFromTextDate(ByVal strDate As String) As Integer
        Dim intMonth As Integer = 0
        Dim strMonth As String = ""

        'สกัดวันที่ออกจากข้อความ
        strMonth = strDate.Substring(3, 2)

        'แปลงค่าจากตัวอักษรเป็นตัวเลขให้กับ intDay
        intMonth = Integer.Parse(strMonth)

        Return intMonth
    End Function

    'Function สำหรับสกัดเดือนชนิด Integer ออกจากข้อความวันที่ชนิด String
    Function getYearFromTextDate(ByVal strDate As String) As Integer
        Dim intYear As Integer
        Dim strYear As String = ""

        'สกัดวันที่ออกจากข้อความ
        strYear = strDate.Substring(6, 4)

        'แปลงค่าจากตัวอักษรเป็นตัวเลขให้กับ intDay
        intYear = Integer.Parse(strYear)

        Return intYear
    End Function

    Function calDiffDay(ByRef datePresent As DateTime, ByRef dateReturn As DateTime) As Integer
        Dim travelTime As TimeSpan = datePresent - dateReturn
        Return Convert.ToInt32(travelTime.TotalDays)
    End Function

    'Function สำหรับคืนค่าชื่อเดือน 12 เดือน
    Function getNameOfMonth(ByVal iMonth As Integer) As String
        Dim strMonthName As String = ""

        Select Case iMonth
            Case 1
                strMonthName = "มกราคม"
            Case 2
                strMonthName = "กุมภาพันธ์"
            Case 3
                strMonthName = "มีนาคม"
            Case 4
                strMonthName = "เมษายน"
            Case 5
                strMonthName = "พฤษภาคม"
            Case 6
                strMonthName = "มิถุนายน"
            Case 7
                strMonthName = "กรกฎาคม"
            Case 8
                strMonthName = "สิงหาคม"
            Case 9
                strMonthName = "กันยายน"
            Case 10
                strMonthName = "ตุลาคม"
            Case 11
                strMonthName = "พฤศจิกายน"
            Case 12
                strMonthName = "ธันวาคม"

        End Select

        Return strMonthName
    End Function
End Module
