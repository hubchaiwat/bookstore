Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections
Imports System.Globalization

Public Class FormAmountMonth

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

#Region "Structure"
    Private Structure MemberAmout
        Dim strMemberID As String
        Dim strMemberName As String
        Dim strMemberSurname As String
        Dim iAmount As Integer

        Sub Init()
            strMemberID = ""
            strMemberName = ""
            iAmount = 0
        End Sub
    End Structure
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

    Private Sub FormAmountMonth_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
        Dim pReportAbout As New Point(245, 120)

        Dim strDatePrint As String = ModuleDate.makeDatePrint(DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year)
        Dim strLinePoint As String = "----------------------------------------------------------------------------"
        Dim strPrintDate As String = "วันที่พิมพ์ " + strDatePrint

        Dim strLogo As String = getStoreName()
        Dim strReportAbout As String = "รายงานแสดงการใช้บริการของสมาชิก"
        Dim strMonthPrint As String = "ประจำเดือน " + strMonthName

        Dim strHeadMember As String = "รหัสสมาชิก"
        Dim strHeadName As String = "ชื่อ-นามสกุล"
        Dim strHeadAmount As String = "จำนวนครั้งที่ใช้บริการ"

        'พิมพ์วันที่พิมพ์รายงาน
        e.Graphics.DrawString(strPrintDate, fnt, Brushes.Black, pPrintDate)

        'พิมพ์ Bank of Cartoon
        fnt = New Font(cstFontName, 22, FontStyle.Bold)
        e.Graphics.DrawString(strLogo, fnt, Brushes.Black, pLogo)

        'พิมพ์รายละเอียดรายงาน
        fnt = New Font(cstFontName, 18)
        e.Graphics.DrawString(strReportAbout, fnt, Brushes.Black, pReportAbout)

        'พิมพ์ประจำเดือน
        Dim yTarget As Integer = 120 + fnt.Height
        e.Graphics.DrawString(strMonthPrint, fnt, Brushes.Black, 315, yTarget)

        yTarget = yTarget + fnt.Height
        fnt = New Font(cstFontName, 16)

        'เก็บข้อมูลรายละเอียดสมาชิกแต่ละคน
        Dim sb As New StringBuilder()
        sb.Append("SELECT DISTINCT MemberID,MemberName,Surname")
        sb.Append(" FROM RentNote")
        sb.Append(" WHERE RentDate BETWEEN @Begin AND @END")
        sb.Append(" ORDER BY MemberID")


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

                'พิมพ์เส้นประ
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)
                yTarget = yTarget + fnt.Height

                'พิมพ์หัวคอลัมภ์
                e.Graphics.DrawString(strHeadMember, fnt, Brushes.Black, 40, yTarget)
                e.Graphics.DrawString(strHeadName, fnt, Brushes.Black, 250, yTarget)
                e.Graphics.DrawString(strHeadAmount, fnt, Brushes.Black, 550, yTarget)
                yTarget = yTarget + fnt.Height

                'พิมพ์เส้นประ
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)
                fnt = New Font(cstFontName, 16, FontStyle.Regular)

                Dim dt As New DataTable()
                dt.Load(dr)
                dr.Close()
                CloseConnection()

                Dim dbSumRent As Double = 0D
                Dim dbSumFine As Double = 0D
                Dim dbSumAll As Double = 0D

                Dim arrMember(dt.Rows.Count - 1) As MemberAmout

                For I As Integer = 0 To dt.Rows.Count - 1
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    arrMember(I).strMemberID = dt.Rows(I)("MemberID").ToString()
                    arrMember(I).strMemberName = dt.Rows(I)("MemberName").ToString()
                    arrMember(I).strMemberSurname = dt.Rows(I)("Surname").ToString()


                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Next

                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                'เก็บจำนวนการใช้บริการของสมาชิกแต่ละคน

                sb = New StringBuilder()
                sb.Append("SELECT MemberID")
                sb.Append(" FROM RentNote")
                sb.Append(" WHERE RentDate BETWEEN @Begin AND @END")

                sqlSelect = sb.ToString()
                OpenConnection()
                com = New OleDbCommand()
                com.CommandType = CommandType.Text
                com.CommandText = sqlSelect
                com.Connection = Conn

                com.Parameters.Add("@Begin", OleDbType.Date).Value = strBegin
                com.Parameters.Add("@End", OleDbType.Date).Value = strEnd

                Try
                    dr = com.ExecuteReader()

                    If dr.HasRows = True Then
                        dt = New DataTable()
                        dt.Load(dr)

                        Dim strPastMemberID As String = ""

                        For I As Integer = 0 To dt.Rows.Count - 1
                            Dim strCurrentMemberID As String = dt.Rows(I)("MemberID").ToString()

                            If strCurrentMemberID <> strPastMemberID Then
                                For J As Integer = 0 To arrMember.Length - 1
                                    If arrMember(J).strMemberID = strCurrentMemberID Then
                                        arrMember(J).iAmount = arrMember(J).iAmount + 1
                                    End If
                                Next
                            End If

                            strPastMemberID = strCurrentMemberID
                        Next

                    End If

                    dr.Close()
                    CloseConnection()
                Catch ex As Exception
                    CloseConnection()
                    MessageBox.Show(ex.Message, cstWarning)
                End Try
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                '******************************************************************************************
                'พิมพ์จำนวนการใช้บริการของสมาชิก

                Dim iSumAmount As Integer = 0
                For I As Integer = 0 To arrMember.Length - 1
                    yTarget = yTarget + fnt.Height

                    'พิมพ์รหัสสมาชิก
                    e.Graphics.DrawString(arrMember(I).strMemberID, fnt, Brushes.Black, 40, yTarget)

                    'พิมพ์ชื่อและนามสกุลสมาชิก
                    e.Graphics.DrawString(arrMember(I).strMemberName + " " + arrMember(I).strMemberSurname, fnt, Brushes.Black, 250, yTarget)

                    'พิมพ์จำนวนครั้งที่ใช้บริการ
                    Dim iTargetAmount As Integer = ModulePrint.getXNumber(675, arrMember(I).iAmount)
                    iSumAmount = iSumAmount + arrMember(I).iAmount
                    Dim strAmount As String = ModulePrint.getStringNumber(arrMember(I).iAmount)
                    e.Graphics.DrawString(strAmount, fnt, Brushes.Black, iTargetAmount, yTarget)
                Next

                'พิมพ์เส้น ------------------
                yTarget = yTarget + fnt.Height
                fnt = New Font(cstFontName, 16, FontStyle.Bold)
                e.Graphics.DrawString(strLinePoint, fnt, Brushes.Black, 30, yTarget)

                'พิมพ์คำว่ารวม
                yTarget = yTarget + fnt.Height
                e.Graphics.DrawString("รวม", fnt, Brushes.Black, 300, yTarget)

                'พิมพ์รวมจำนวนการใช้บริการของสมาชิก
                Dim iTargetSumAmount As Integer = ModulePrint.getXNumber(673, iSumAmount)
                Dim strSumAmount As String = ModulePrint.getStringNumber(iSumAmount)
                e.Graphics.DrawString(strSumAmount, fnt, Brushes.Black, iTargetSumAmount, yTarget)

                '******************************************************************************************

            Else
                e.Graphics.DrawString("ไม่มีข้อมูลการใช้บริการของสมาชิกในเดือนที่กำหนดค่ะ", fnt, Brushes.Black, 190, yTarget)
            End If

        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub
End Class