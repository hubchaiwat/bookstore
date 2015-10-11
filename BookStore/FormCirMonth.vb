Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization

Public Class FormCirMonth
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

    Private Sub FormCirMonth_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
            tbYear.Focus()
            Return
        End If

        Dim iYear As Integer = Integer.Parse(tbYear.Text)

        If iYear < 544 Then
            MessageBox.Show("ปีพ.ศ. ต้องมีค่ามากกว่าหรือเท่ากับ 544 ค่ะ")
            tbYear.Focus()
            Return
        End If

        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As System.Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        If tbYear.Text = "" Then
            MessageBox.Show("คุณยังไม่กรอกปีที่ต้องการพิมพ์รายงานค่ะ")
            Return
        End If

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
        Dim pReportAbout As New Point(320, 120)

        Dim strDatePrint As String = ModuleDate.makeDatePrint(DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year)
        Dim strLinePoint As String = "----------------------------------------------------------------------------"
        Dim strPrintDate As String = "วันที่พิมพ์ " + strDatePrint
        Dim strLogo As String = getStoreName()
        Dim strReportAbout As String = "รายงานแสดงรายได้"
        Dim strMonthPrint As String = "ประจำเดือน " + strMonthName
        Dim strHeader As String = "วันที่         รายได้บริการ(ค่าเช่า)          รายได้บริการ(ค่าปรับ)               รวม"

        e.Graphics.DrawString(strPrintDate, fnt, Brushes.Black, pPrintDate)

        fnt = New Font(cstFontName, 22, FontStyle.Bold)
        e.Graphics.DrawString(strLogo, fnt, Brushes.Black, pLogo)

        fnt = New Font(cstFontName, 18)
        e.Graphics.DrawString(strReportAbout, fnt, Brushes.Black, pReportAbout)

        Dim yTarget As Integer = 120 + fnt.Height
        e.Graphics.DrawString(strMonthPrint, fnt, Brushes.Black, 315, yTarget)

        yTarget = yTarget + fnt.Height
        fnt = New Font(cstFontName, 16)


        Dim sb As New StringBuilder()
        sb.Append("SELECT FORMAT(DateDetail,'dd/MM/yyyy') AS DateDetail,")
        sb.Append(" RentIncome,FineIncome")
        sb.Append(" FROM DayDetail")
        sb.Append(" WHERE DateDetail BETWEEN @Begin AND @END")

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
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString(strHeader, fnt, Brushes.Black, 30, yTarget)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)
                fnt = New Font(cstFontName, 16, FontStyle.Regular)

                Dim dt As New DataTable()
                dt.Load(dr)
                dr.Close()
                CloseConnection()

                Dim dbSumRent As Double = 0D
                Dim dbSumFine As Double = 0D
                Dim dbSumAll As Double = 0D

                For I As Integer = 0 To dt.Rows.Count - 1
                    yTarget = yTarget + fnt.Height
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    'พิมพ์วันที่
                    Dim strDateDetail As String = dt.Rows(I)("DateDetail").ToString()
                    Dim iDay As Integer = ModuleDate.getDayFromTextDate(strDateDetail)
                    Dim iTargetDay As Integer = ModulePrint.getXInteger(50, iDay)
                    e.Graphics.DrawString(iDay.ToString(), fnt, Brushes.Black, iTargetDay, yTarget)

                    'พิมพ์รายได้ค่าเช่า
                    Dim dbRentIncome As Double = Double.Parse(dt.Rows(I)("RentIncome").ToString())
                    dbSumRent = dbSumRent + dbRentIncome
                    Dim iTargetRentIncome As Integer = ModulePrint.getXCurrency(240, dbRentIncome)
                    Dim strRentIncome As String = ModulePrint.getStringCurrency(dbRentIncome)
                    e.Graphics.DrawString(strRentIncome, fnt, Brushes.Black, iTargetRentIncome, yTarget)

                    'พิมพ์รายได้ค่าปรับ
                    Dim dbFineIncome As Double = Double.Parse(dt.Rows(I)("FineIncome").ToString())
                    dbSumFine = dbSumFine + dbFineIncome
                    Dim iTargetFineIncome As Integer = ModulePrint.getXCurrency(520, dbFineIncome)
                    Dim strFineIncome As String = ModulePrint.getStringCurrency(dbFineIncome)
                    e.Graphics.DrawString(strFineIncome, fnt, Brushes.Black, iTargetFineIncome, yTarget)

                    'พิมพ์รายได้รวมเฉพาะวัน
                    Dim dbAllIncome As Double = dbRentIncome + dbFineIncome
                    Dim iTargetAllIncome As Integer = ModulePrint.getXCurrency(725, dbAllIncome)
                    Dim strAllIncome As String = ModulePrint.getStringCurrency(dbAllIncome)
                    e.Graphics.DrawString(strAllIncome, fnt, Brushes.Black, iTargetAllIncome, yTarget)

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Next

                'พิมพ์เส้น ------------------
                yTarget = yTarget + fnt.Height
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)

                'พิมพ์คำว่ารวม
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString("รวม", fnt, Brushes.Black, 30, yTarget)

                'พิมพ์รายได้รวมค่าเช่า
                Dim iTargetSumRent As Integer = ModulePrint.getXCurrency(220, dbSumRent)
                Dim strSumRent As String = ModulePrint.getStringCurrency(dbSumRent)
                e.Graphics.DrawString(strSumRent, fnt, Brushes.Black, iTargetSumRent, yTarget)

                'พิมพ์รายได้รวมค่าปรับ
                Dim iTargetSumFine As Integer = ModulePrint.getXCurrency(500, dbSumFine)
                Dim strSumFine As String = ModulePrint.getStringCurrency(dbSumFine)
                e.Graphics.DrawString(strSumFine, fnt, Brushes.Black, iTargetSumFine, yTarget)

                'พิมพ์รายได้รวมค่าปรับ
                dbSumAll = dbSumRent + dbSumFine
                Dim iTargetSumAll As Integer = ModulePrint.getXCurrency(705, dbSumAll)
                Dim strSumAll As String = ModulePrint.getStringCurrency(dbSumAll)
                e.Graphics.DrawString(strSumAll, fnt, Brushes.Black, iTargetSumAll, yTarget)

            Else
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString("ไม่มีข้อมูลรายได้ในเดือนที่กำหนดค่ะ", fnt, Brushes.Black, 250, yTarget)
            End If

        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    
End Class