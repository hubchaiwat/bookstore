Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization

Public Class FormAmountYear

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

#Region "Structure"
    Private Structure CirMonth
        Dim iAmountMember As Integer
        Dim dbRentIncome As Double
        Dim dbFineIncome As Double
        Dim bActivated As Boolean

        Sub Init()
            iAmountMember = 0
            dbRentIncome = 0D
            dbFineIncome = 0D
            bActivated = False
        End Sub

    End Structure
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

    Private Sub FormAmountYear_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        PrintPreviewDialog1.Document = PrintDocument1
        tbYear.Text = (DateTime.Now.Year + 543).ToString()
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

        Dim iCurrentMonth As Integer = DateTime.Now.Month

        Dim recCirMonth(iCurrentMonth - 1) As CirMonth

        For I As Integer = 0 To recCirMonth.Length - 1
            recCirMonth(I) = New CirMonth()
            recCirMonth(I).Init()
        Next

        Dim strBegin As String = ModuleDate.makeDateTran(1, 1, iYear)
        Dim strEnd As String = ModuleDate.makeDateTran(31, 12, iYear)

        Dim fnt As New Font(cstFontName, 16)
        Dim pPrintDate As New Point(550, 40)
        Dim pLogo As New Point(300, 80)
        Dim pReportAbout As New Point(245, 120)

        Dim strDatePrint As String = ModuleDate.makeDatePrint(DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year)
        Dim strLinePoint As String = "----------------------------------------------------------------------------"
        Dim strPrintDate As String = "วันที่พิมพ์ " + strDatePrint
        Dim strLogo As String = getStoreName()
        Dim strReportAbout As String = "รายงานแสดงการใช้บริการของสมาชิก"

        Dim iYearPrint As Integer = iYear

        If (CultureInfo.CurrentCulture.ToString() = "th-TH") Then
            iYearPrint = iYearPrint - 543
        End If

        Dim strMonthPrint As String = "ประจำปี พ.ศ. " + (iYearPrint + 543).ToString()
        Dim strHeadMonth As String = "เดือน"
        Dim strHeadAmount As String = "จำนวนครั้งที่ใช้บริการ"

        'พิมพ์วันที่พิมพ์รายงาน
        e.Graphics.DrawString(strPrintDate, fnt, Brushes.Black, pPrintDate)

        'พิมพ์ Bank of Cartoon
        fnt = New Font(cstFontName, 22, FontStyle.Bold)
        e.Graphics.DrawString(strLogo, fnt, Brushes.Black, pLogo)

        'พิมพ์หัวข้อรายงาน
        fnt = New Font(cstFontName, 18)
        e.Graphics.DrawString(strReportAbout, fnt, Brushes.Black, pReportAbout)

        'พิมพ์ประจำประจำปี
        Dim yTarget As Integer = 120 + fnt.Height
        e.Graphics.DrawString(strMonthPrint, fnt, Brushes.Black, 320, yTarget)

        yTarget = yTarget + fnt.Height
        fnt = New Font(cstFontName, 16)


        Dim sb As New StringBuilder()
        sb.Append("SELECT FORMAT(DateDetail,'dd/MM/yyyy') AS DateDetail,")
        sb.Append("AmountMember")
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
                e.Graphics.DrawString(strHeadMonth, fnt, Brushes.Black, 200, yTarget)
                e.Graphics.DrawString(strHeadAmount, fnt, Brushes.Black, 450, yTarget)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)
                fnt = New Font(cstFontName, 16, FontStyle.Regular)

                Dim dt As New DataTable()
                dt.Load(dr)
                dr.Close()
                CloseConnection()

                Dim iSumAmount As Integer = 0

                For I As Integer = 0 To dt.Rows.Count - 1
                    Dim strDateDetail As String = dt.Rows(I)("DateDetail").ToString()
                    Dim iAmountMember As Integer = dt.Rows(I)("AmountMember").ToString()
                    iSumAmount = iSumAmount + iAmountMember
                    Dim iMonth As Integer = ModuleDate.getMonthFromTextDate(strDateDetail)
                    iMonth = iMonth - 1

                    recCirMonth(iMonth).bActivated = True
                    recCirMonth(iMonth).iAmountMember = recCirMonth(iMonth).iAmountMember + iAmountMember
                Next

                For I As Integer = 0 To recCirMonth.Length - 1
                    yTarget = yTarget + fnt.Height
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    'พิมพ์เดือน
                    Dim strMonthName As String = ModuleDate.getNameOfMonth(I + 1)
                    e.Graphics.DrawString(strMonthName, fnt, Brushes.Black, 170, yTarget)

                    If recCirMonth(I).bActivated <> False Then
                        'พิมพ์จำนวนครั้งที่ใช้บริการ
                        Dim iTargetAmountMember As Integer = ModulePrint.getXNumber(570, recCirMonth(I).iAmountMember)
                        Dim strAmountMember As String = ModulePrint.getStringNumber(recCirMonth(I).iAmountMember)
                        e.Graphics.DrawString(strAmountMember, fnt, Brushes.Black, iTargetAmountMember, yTarget)


                    Else
                        e.Graphics.DrawString("ไม่มีข้อมูลการใช้บริการของสมาชิกในเดือนนี้ค่ะ", fnt, Brushes.Black, 350, yTarget)
                    End If

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Next

                'พิมพ์เส้น ------------------
                yTarget = yTarget + fnt.Height
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)

                'พิมพ์คำว่ารวม
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString("รวม", fnt, Brushes.Black, 170, yTarget)

                'พิมพ์รวมจำนวนการใช้บริการของสมาชิก
                Dim iTargetSumAmount As Integer = ModulePrint.getXNumber(561, iSumAmount)
                Dim strSumAmount As String = ModulePrint.getStringNumber(iSumAmount)
                e.Graphics.DrawString(strSumAmount, fnt, Brushes.Black, iTargetSumAmount, yTarget)


            Else
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString("ไม่มีข้อมูลการใช้บริการของสมาชิกในปีที่กำหนดค่ะ", fnt, Brushes.Black, 180, yTarget)
            End If

        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub
End Class