Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization

Public Class FormReportOver

#Region "Connecting"
    '/////////////////////////////ส่วนเชื่อมต่อฐานข้อมูล/////////////////////////////////////
    Private Const strConn As String = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=C:\\Database\\mybookstore.mdb"
    Private Conn As OleDbConnection = New OleDbConnection()

    Private Sub OpenConnection()
        Try
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
            Conn.ConnectionString = strConn
            Conn.Open()
        Catch ex As Exception
            Dim strError As String = ""
            strError = strError + "ไฟล์ Database ของโปรแกรมได้รับความเสียหายค่ะ" + Environment.NewLine
            strError = strError + "กรุณาตรวจสอบไฟล์ Database ของท่านอีกครั้งว่าไฟล์อยู่ในตำแหน่งที่ถูกต้องหรือไม่" + Environment.NewLine
            strError = strError + "ตำแหน่งที่ถูกต้องของ Database คือ C:\\Database\\mybookstore.mdb" + Environment.NewLine
            MessageBox.Show(strError, cstWarning)
            Me.Close()
        End Try

    End Sub

    Private Sub CloseConnection()
        Try
            If (Conn.State = ConnectionState.Open) Then
                Conn.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "คำเตือน")
        End Try

    End Sub
#End Region

#Region "Constant"
    Private Const cstWarning As String = "คำเตือน"
    Private Const cstFontName As String = "Tahoma"
#End Region

    Private Function getStoreName() As String
        Dim strName As String = ""

        Dim sqlSelect As String = "SELECT StoreName FROM StoreDetail"

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            Dim dt As New DataTable()
            dt.Load(dr)

            strName = dt.Rows(0)("StoreName")

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        Return strName
    End Function

    Private Sub FormReportOver_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PrintPreviewDialog1.Document = PrintDocument1
        tbYear.Text = (DateTime.Now.Year + 543).ToString()

        Dim iMonth As Integer = DateTime.Now.Month
        iMonth = iMonth - 1
        cbMonth.SelectedIndex = iMonth
        cbMonth.Text = cbMonth.Items(iMonth).ToString()
        'PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub tbYear_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles tbYear.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub cbMonth_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles cbMonth.KeyPress
        e.Handled = True
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        If tbYear.Text = "" Then
            MessageBox.Show("คุณยังไม่กรอกปีที่ต้องการพิมพ์รายงานค่ะ")
            Return
        End If

        Dim iYear As Integer = Integer.Parse(tbYear.Text)

        If iYear < 544 Then
            MessageBox.Show("ปีพ.ศ. ต้องมีค่ามากกว่าหรือเท่ากับ 544 ค่ะ")
            Return
        End If

        PrintPreviewDialog1.ShowDialog()

    End Sub

    Private Function getAuthor(ByVal strBookID As String) As String
        Dim strAuthor As String = ""

        Dim sb As New StringBuilder()
        sb.Append("SELECT Author")
        sb.Append(" FROM Books")
        sb.Append(" WHERE BookID=@BookID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = strBookID

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)
                strAuthor = dt.Rows(0)("Author").ToString()
            Else
                strAuthor = "ไม่มีข้อมูลผู้แต่ง"
            End If
            
            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message)
        End Try

        Return strAuthor
    End Function

    Private Function getXDayOver(iXBegin As Integer, iDayOver As Integer) As Integer
        If ((iDayOver >= 10) And (iDayOver < 100)) Then
            iXBegin = iXBegin - 10
        ElseIf ((iDayOver >= 100) And (iDayOver < 1000)) Then
            iXBegin = iXBegin - 20
        ElseIf ((iDayOver >= 1000) And (iDayOver < 10000)) Then
            iXBegin = iXBegin - 30
        ElseIf (iDayOver >= 10000) Then
            iXBegin = iXBegin - 40
        End If
        Return iXBegin
    End Function

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim iYear As Integer = Integer.Parse(tbYear.Text)
        iYear = iYear - 543

        If (CultureInfo.CurrentCulture.ToString() = "th-TH") Then
            iYear = iYear + 543
        End If

        Dim iMonth As Integer = 0
        Dim strMonthName As String = ""
        If cbMonth.SelectedIndex = -1 Then
            MessageBox.Show("คุณยังไม่ได้เลือกเดือนค่ะ")
            Return
        Else
            iMonth = cbMonth.SelectedIndex
            strMonthName = cbMonth.SelectedItem
        End If
        iMonth = iMonth + 1

        Dim iMaxDay As Integer = DateTime.DaysInMonth(iYear, iMonth)

        Dim strBegin As String = ModuleDate.makeDateTran(1, iMonth, iYear)
        Dim strEnd As String = ModuleDate.makeDateTran(iMaxDay, iMonth, iYear)

        Dim fnt As New Font(cstFontName, 16)
        Dim pPrintDate As New Point(550, 40)
        Dim pLogo As New Point(300, 80)
        Dim pReportAbout As New Point(200, 120)

        Dim strLinePoint As String = "------------------------------------------------------------------------------------------"
        Dim strDatePrint As String = ModuleDate.makeDatePrint(DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year)
        Dim strPrintDate As String = "วันที่พิมพ์ " + strDatePrint
        Dim strLogo As String = getStoreName()
        Dim strReportAbout As String = "รายงานแสดงรายชื่อหนังสือที่เช่าเกินกำหนด"
        Dim strMonthPrint As String = "ประจำเดือน " + strMonthName
        Dim strHeader As String = "รหัสหนังสือ    ชื่อหนังสือ                   ผู้แต่ง                   วันที่เช่า      วันที่คืน จำนวนวันที่เกิน"

        e.Graphics.DrawString(strPrintDate, fnt, Brushes.Black, pPrintDate)

        fnt = New Font(cstFontName, 22, FontStyle.Bold)
        e.Graphics.DrawString(strLogo, fnt, Brushes.Black, pLogo)

        fnt = New Font(cstFontName, 18)
        e.Graphics.DrawString(strReportAbout, fnt, Brushes.Black, pReportAbout)

        Dim yTarget As Integer = 120 + fnt.Height
        e.Graphics.DrawString(strMonthPrint, fnt, Brushes.Black, 300, yTarget)

        yTarget = yTarget + fnt.Height
        fnt = New Font(cstFontName, 14)


        Dim sb As New StringBuilder()
        sb.Append("SELECT BookID,BookName,")
        sb.Append("FORMAT(RentDate,'dd/MM/yyyy') AS RentDate,")
        sb.Append("FORMAT(ReturnDate,'dd/MM/yyyy') AS ReturnDate,ReturnStatus")
        sb.Append(" FROM RentNote")
        sb.Append(" WHERE (RentDate BETWEEN @Begin AND @END)")
        sb.Append(" ORDER BY BookID,RentDate")

        Dim sqlSelect As String = sb.ToString()
        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@Begin", OleDbType.Date).Value = strBegin
        com.Parameters.Add("@End", OleDbType.Date).Value = strEnd

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()
            If dr.HasRows = True Then

                'พิมพ์หัวคอลัมภ์
                fnt = New Font(cstFontName, 14, FontStyle.Bold)
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 20, yTarget)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString(strHeader, fnt, Brushes.Black, 20, yTarget)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 20, yTarget)
                fnt = New Font(cstFontName, 14, FontStyle.Regular)

                Dim dt As New DataTable()
                dt.Load(dr)
                dr.Close()
                CloseConnection()
                For I As Integer = 0 To dt.Rows.Count - 1
                    fnt = New Font(cstFontName, 14)
                    Dim strRentDate As String = dt.Rows(I)("RentDate").ToString()
                    Dim iRentDay As Integer = getDayFromTextDate(strRentDate)
                    Dim iRentMonth As Integer = getMonthFromTextDate(strRentDate)
                    Dim iRentYear As Integer = getYearFromTextDate(strRentDate)
                    Dim dtRent As DateTime = New DateTime(iRentYear, iRentMonth, iRentDay)

                    Dim strReturnDate As String = dt.Rows(I)("ReturnDate").ToString()
                    If strReturnDate <> "" Then
                        Dim iReturnDay As Integer = getDayFromTextDate(strReturnDate)
                        Dim iReturnMonth As Integer = getMonthFromTextDate(strReturnDate)
                        Dim iReturnYear As Integer = getYearFromTextDate(strReturnDate)
                        Dim dtReturn As DateTime = New DateTime(iReturnYear, iReturnMonth, iReturnDay)
                        Dim iDiffDay As Integer = calDiffDay(dtReturn, dtRent) - 1

                        If iDiffDay > 0 Then
                            yTarget = yTarget + fnt.Height

                            'พิมพ์รหัสหนังสือ'
                            Dim strBookID As String = dt.Rows(I)("BookID")
                            e.Graphics.DrawString(strBookID, fnt, Brushes.Black, 40, yTarget)

                            'พิมพ์ชื่อหนังสือ
                            Dim strBookName As String = dt.Rows(I)("BookName")
                            If strBookName.Length >= 20 Then
                                strBookName = strBookName.Substring(0, 20)
                            End If
                            e.Graphics.DrawString(strBookName, fnt, Brushes.Black, 115, yTarget)

                            'พิมพ์ชื่อผู้แต่ง
                            Dim strAuthor As String = getAuthor(strBookID)
                            If strAuthor.Length >= 20 Then
                                strAuthor = strAuthor.Substring(0, 20)
                            End If
                            e.Graphics.DrawString(strAuthor, fnt, Brushes.Black, 325, yTarget)

                            'พิมพ์วันที่เช่า
                            Dim strRentPrint As String = ModuleDate.makeDateRentAndReturn(iRentDay, iRentMonth, iRentYear)
                            e.Graphics.DrawString(strRentPrint, fnt, Brushes.Black, 495, yTarget)

                            'พิมพ์วันที่คืน
                            Dim strReturnPrint As String = ModuleDate.makeDateRentAndReturn(iReturnDay, iReturnMonth, iReturnYear)
                            e.Graphics.DrawString(strReturnPrint, fnt, Brushes.Black, 600, yTarget)

                            'พิมพ์จำนวนวันที่เกิน
                            Dim iXDayOver As Integer = getXDayOver(750, iDiffDay)
                            e.Graphics.DrawString(iDiffDay.ToString(), fnt, Brushes.Black, iXDayOver, yTarget)

                        End If
                    End If

                Next
            Else
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString("ไม่มีข้อมูลหนังสือที่เช่าเกินกำหนดในเดือนนี้ค่ะ", fnt, Brushes.Black, 195, yTarget)
            End If

        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

    End Sub
End Class