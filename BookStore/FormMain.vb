Imports System.Data
Imports System.Data.OleDb
Imports System.Text
Imports System.Globalization

Public Class FormMain

#Region "Constant"
    Private Const cstTitle As String = "ผลการดำเนินงาน"
    Private Const cstWarning As String = "คำเตือน"
    Private Const maxRent As Integer = 5
#End Region

#Region "Structure"
    Private Structure Member
        Dim MemberID As String
        Dim MemberName As String
        Dim Surname As String
        Dim PastRent As Integer

        Sub Init()
            MemberID = ""
            MemberName = ""
            Surname = ""
            PastRent = 0
        End Sub
    End Structure

    Private Structure BookRent
        Dim BookID As String
        Dim BookName As String
        Dim RentPrice As Double

        Sub Init()
            BookID = ""
            BookName = ""
            RentPrice = 0D
        End Sub
    End Structure

    Private Structure RentDetail
        Dim amountRent As Integer
        Dim rentDate As String
        Dim totalPrice As Double

        Sub Init()
            amountRent = 0
            rentDate = ""
            totalPrice = 0D
        End Sub

        Sub setToPresentDate()
            Dim day As Integer = DateTime.Now.Day
            Dim month As Integer = DateTime.Now.Month
            Dim year As Integer = DateTime.Now.Year

            rentDate = ModuleDate.makeDateTran(day, month, year)
        End Sub
    End Structure

    Private Structure BookBeforeReturn
        Dim BookID As String
        Dim BookName As String
        Dim RentDate As String
        Dim RentPrice As Double

        Sub Init()
            BookID = ""
            BookName = ""
            RentDate = ""
            RentPrice = 0D
        End Sub
    End Structure

    Private Structure BookAfterReturn
        Dim BookID As String
        Dim MemberID As String

        Sub Init()
            BookID = ""
            MemberID = ""
        End Sub

    End Structure

    Private Structure ReturnDetail
        Dim MemberID As String
        Dim presentAmount As Integer
        Dim returnDate As String
        Dim totalFine As Double

        Sub Init()
            MemberID = ""
            presentAmount = 0
            returnDate = ""
            totalFine = 0D
        End Sub

        Sub setToPresentDate()
            Dim day As Integer = DateTime.Now.Day
            Dim month As Integer = DateTime.Now.Month
            Dim year As Integer = DateTime.Now.Year

            If (CultureInfo.CurrentCulture.ToString() = "th-TH") Then
                year = year + 543
            End If

            returnDate = ModuleDate.makeDateTran(day, month, year)
        End Sub
    End Structure

    Private Structure SaveStat
        Dim lRecordID As Long
        Dim strDateDetail As String
        Dim iAmountMember As Integer
        Dim dbAmountIncome As Double
        Dim dbAmountFine As Double

        Sub Init()
            lRecordID = 0L
            strDateDetail = ""
            iAmountMember = 0
            dbAmountIncome = 0D
            dbAmountFine = 0D
        End Sub

        Sub setToPresentDate()
            Dim day As Integer = DateTime.Now.Day
            Dim month As Integer = DateTime.Now.Month
            Dim year As Integer = DateTime.Now.Year

            strDateDetail = ModuleDate.makeDateTran(day, month, year)
        End Sub
    End Structure

#End Region

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

#Region "LoadAndClosing"
    Private Sub FormMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OpenConnection()
        dgvReturn.Columns.Add("RentDate", "RentDate")
        dgvReturn.Columns.Add("RentPrice", "RentPrice")
        dgvReturn.Columns("RentDate").Visible = False
        dgvReturn.Columns("RentPrice").Visible = False

        'Dim lMaxRecordID As Long = getMaxRecordID()
        'MessageBox.Show(lMaxRecordID)
    End Sub

    Private Sub FormMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        CloseConnection()
    End Sub
#End Region

#Region "UI Control"
    '//////////////////////////////ส่วนควบคุม UI //////////////////////////////////
    'โพรซีเยอร์เอาไว้เปิด Panel ยืมคืน
    Private Sub OpenPanelBorrowReturn()
        PanelBorrowReturn.Visible = True
        PanelBook.Visible = False
        PanelMember.Visible = False
        tbRentMemberID0.Focus()
    End Sub

    'โพรซีเยอร์เอาไว้เปิด Panel ข้อมูลหนังสือ
    Private Sub OpenPanelBook()
        PanelBorrowReturn.Visible = False
        PanelBook.Visible = True
        PanelMember.Visible = False
    End Sub

    'โพรซีเยอร์เอาไว้เปิด Panel ข้อมูลสมาชิก
    Private Sub OpenPanelMember()
        PanelBorrowReturn.Visible = False
        PanelBook.Visible = False
        PanelMember.Visible = True
        tbMemberID0.Focus()
    End Sub

    'Control ที่เอาไว้เปิด Form รายละเอียดของร้าน
    Private Sub menuShopDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuShopDetail.Click
        Dim frmDetail As New FormStoreDetail
        frmDetail.ShowDialog()

    End Sub

    'Control ที่เอาไว้เปิด Panel ยืมคืน
    Private Sub imgBorrowReturn_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgBorrowReturn.MouseHover
        linkBorrowReturn.ForeColor = Color.Blue
    End Sub

    Private Sub imgBorrowReturn_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgBorrowReturn.MouseLeave
        linkBorrowReturn.ForeColor = Color.Black
    End Sub


    Private Sub imgBorrowReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgBorrowReturn.Click
        OpenPanelBorrowReturn()
    End Sub

    Private Sub linkBorrowReturn_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkBorrowReturn.MouseHover
        linkBorrowReturn.ForeColor = Color.Blue
    End Sub

    Private Sub linkBorrowReturn_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkBorrowReturn.MouseLeave
        linkBorrowReturn.ForeColor = Color.Black
    End Sub

    Private Sub linkBorrowReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkBorrowReturn.Click
        OpenPanelBorrowReturn()
    End Sub

    Private Sub menuBorrowReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuBorrowReturn.Click
        OpenPanelBorrowReturn()
    End Sub

    'Control ที่เอาไว้เปิด Panel ข้อมูลหนังสือ
    Private Sub imgBook_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgBook.MouseHover
        linkBook.ForeColor = Color.Blue
    End Sub

    Private Sub imgBook_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgBook.MouseLeave
        linkBook.ForeColor = Color.Black
    End Sub

    Private Sub imgBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgBook.Click
        OpenPanelBook()
    End Sub

    Private Sub linkBook_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkBook.MouseHover
        linkBook.ForeColor = Color.Blue
    End Sub

    Private Sub linkBook_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkBook.MouseLeave
        linkBook.ForeColor = Color.Black
    End Sub

    Private Sub linkBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkBook.Click
        OpenPanelBook()
    End Sub

    Private Sub menuBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuBook.Click
        OpenPanelBook()
    End Sub

    'Control ที่เอาไว้เปิด Panel สมาชิก
    Private Sub imgMember_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgMember.MouseHover
        linkMember.ForeColor = Color.Blue
    End Sub

    Private Sub imgMember_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgMember.MouseLeave
        linkMember.ForeColor = Color.Black
    End Sub

    Private Sub imgMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgMember.Click
        OpenPanelMember()
    End Sub

    Private Sub linkMember_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkMember.MouseHover
        linkMember.ForeColor = Color.Blue
    End Sub

    Private Sub linkMember_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkMember.MouseLeave
        linkMember.ForeColor = Color.Black
    End Sub

    Private Sub linkMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles linkMember.Click
        OpenPanelMember()
    End Sub

    Private Sub menuMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuMember.Click
        OpenPanelMember()
    End Sub

    'Control MenuBar ที่เอาไว้ปิดโปรแกรม
    Private Sub menuClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuClose.Click
        Me.Close()
    End Sub

    Private Sub menuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuAbout.Click
        Dim frmAboout As FormAbout = New FormAbout()
        frmAboout.ShowDialog()

    End Sub

    Private Sub cbType_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles cbType.KeyPress
        e.Handled = True
    End Sub

    Private Sub cbSearchBook_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles cbSearchBook.KeyPress
        e.Handled = True
    End Sub

    Private Sub cbBookSerach_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles cbMemberSerach.KeyPress
        e.Handled = True
    End Sub

#End Region

#Region "Report"

#End Region

    '///////////////////////////////////////เข้าสู่ขั้นตอนการติดต่อฐานข้อมูล//////////////////////////////////////////////////
#Region "Member"
    'Event ที่เกิดจากการกรอกรหัสสมาชิก
    Private Sub tbMemberID0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMemberID0.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            tbMemberID0.Focus()
        ElseIf Char.IsNumber(e.KeyChar) = True Then
            tbMemberID1.Focus()
        Else
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub tbMemberID1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMemberID1.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbMemberID1.Text.Length = 0 Then
                tbMemberID0.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
            Return
        End If

        If tbMemberID1.Text.Length = 3 Then
            tbMemberID2.Focus()
        End If
    End Sub

    Private Sub tbMemberID2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMemberID2.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbMemberID2.Text.Length = 0 Then
                tbMemberID1.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbMemberID2.Text.Length = 4 Then
            tbMemberID3.Focus()
        End If
    End Sub

    Private Sub tbMemberID3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMemberID3.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbMemberID3.Text.Length = 0 Then
                tbMemberID2.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbMemberID3.Text.Length = 1 Then
            tbMemberID4.Focus()
        End If
    End Sub

    Private Sub tbMemberID4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMemberID4.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbMemberID4.Text.Length = 0 Then
                tbMemberID3.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    'Procedure สำหรับรวมรหัสสมาชิกจากช่องต่างๆ
    Private Function getMergeMemberID() As String
        Return tbMemberID0.Text.Trim() + tbMemberID1.Text.Trim() + tbMemberID2.Text.Trim() + tbMemberID3.Text.Trim() + tbMemberID4.Text.Trim()
    End Function

    'Procedure สำหรับสกัดรหัสสมาชิกในช่วงต่างๆ
    Private Function getHyphenMemberID(ByVal strMemberID As String) As String
        Dim strNewMemberID As String = ""
        Dim rgMemberID(5) As String

        rgMemberID(0) = strMemberID.Substring(0, 1)
        rgMemberID(1) = strMemberID.Substring(1, 4)
        rgMemberID(2) = strMemberID.Substring(5, 5)
        rgMemberID(3) = strMemberID.Substring(10, 2)
        rgMemberID(4) = strMemberID.Substring(12, 1)

        strNewMemberID = strNewMemberID + rgMemberID(0) + "-"
        strNewMemberID = strNewMemberID + rgMemberID(1) + "-"
        strNewMemberID = strNewMemberID + rgMemberID(2) + "-"
        strNewMemberID = strNewMemberID + rgMemberID(3) + "-"
        strNewMemberID = strNewMemberID + rgMemberID(4)

        Return strNewMemberID
    End Function

    'Procudure สำหรับใส่ * label สมาชิก
    Private Sub PlusStarMember()
        lbMember1.Text = lbMember1.Text + "*"
        lbMember2.Text = lbMember2.Text + "*"
        lbMember3.Text = lbMember3.Text + "*"
        lbMember4.Text = lbMember4.Text + "*"
        lbMember5.Text = lbMember5.Text + "*"

    End Sub

    'Procedure สำหรับลบ * label สมาชิก
    Private Sub DeleteStarMember()
        If lbMember1.Text(lbMember1.Text.Length - 1) = "*" Then
            lbMember1.Text = lbMember1.Text.Substring(0, lbMember1.Text.Length - 1)
            lbMember2.Text = lbMember2.Text.Substring(0, lbMember2.Text.Length - 1)
            lbMember3.Text = lbMember3.Text.Substring(0, lbMember3.Text.Length - 1)
            lbMember4.Text = lbMember4.Text.Substring(0, lbMember4.Text.Length - 1)
            lbMember5.Text = lbMember5.Text.Substring(0, lbMember5.Text.Length - 1)
        End If
    End Sub

    'Procedure ไว้ลบข้อมูลในฟิลด์หน้ารายละเอียดสมาชิก
    Private Sub clearForNewMember()
        tbMemberName.Text = ""
        tbSurname.Text = ""
        tbAddress.Text = ""
        tbTel.Text = ""
        tbMemberNote.Text = ""
        btnEditMember.Enabled = False
        btnSearchEachMember.Visible = True
        btnChkDupMemberID.Visible = False
        tbMemberID0.ReadOnly = False
    End Sub

    'Procedure set Enable value ฟิลด์สมาชิก
    Private Sub setEnableFieldMember(ByVal bValue As Boolean)
        tbMemberName.Enabled = bValue
        tbSurname.Enabled = bValue
        tbAddress.Enabled = bValue
        tbTel.Enabled = bValue
        tbMemberNote.Enabled = bValue
    End Sub

    'Procedure set Readonly value ฟิลด์สมาชิก
    Private Sub setReadOnlyFieldMember(ByVal bValue As Boolean)
        tbMemberName.ReadOnly = bValue
        tbSurname.ReadOnly = bValue
        tbAddress.ReadOnly = bValue
        tbTel.ReadOnly = bValue
        tbMemberNote.ReadOnly = bValue
    End Sub

    'Event สำหรับตรวจสอบรหัสสมาชิกขณะเพิ่มข้อมูลสมาชิก
    Private Sub btnChkDupMemberID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChkDupMemberID.Click
        Dim strMemberID As String = getMergeMemberID()

        If strMemberID = "" Then
            MessageBox.Show("คุณยังไม่ได้กรอกรหัสสมาชิกเพื่อตรวจสอบค่ะ", cstWarning)
            tbMemberID0.Focus()
            Return
        End If

        If strMemberID.Length < 13 Then
            MessageBox.Show("คุณกรอกรหัสสมาชิกไม่ครบ 13 ตัวค่ะ", cstWarning)
            Return
        End If

        Dim sb As New StringBuilder()
        sb.Append("SELECT COUNT(MemberID) AS CountMemberID")
        sb.Append(" FROM Members")
        sb.Append(" WHERE MemberID=@MemberID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        com.Parameters.Add("MemberID", OleDbType.VarChar).Value = strMemberID

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                Dim iCount = dt.Rows(0)("CountMemberID")
                If iCount <> 0 Then
                    MessageBox.Show("มีสมาชิกที่ตรงกับรหัสนี้แล้ว กรุณากรอกใหม่อีกครั้งค่ะ")
                    tbMemberID0.Text = ""
                    tbMemberID1.Text = ""
                    tbMemberID2.Text = ""
                    tbMemberID3.Text = ""
                    tbMemberID4.Text = ""
                    tbMemberID0.Focus()
                Else
                    tbMemberID0.ReadOnly = True
                    tbMemberID1.ReadOnly = True
                    tbMemberID2.ReadOnly = True
                    tbMemberID3.ReadOnly = True
                    tbMemberID4.ReadOnly = True
                    setReadOnlyFieldMember(False)
                    setEnableFieldMember(True)
                    btnChkDupMemberID.Enabled = False
                    tbMemberName.Focus()
                End If

            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    'Procedure เพิ่มสมาชิก
    Private Function InsertToMembers() As Boolean
        Dim bComplete As Boolean = False

        'สร้างคำสั่ง SQL ในการ Insert 
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("INSERT INTO Members(MemberID,MemberName,Surname,Address,Tel,MemberNote,AmountRent)")
        sb.Append(" VALUES(@MemberID,@MemberName,@Surname,@Address,@Tel,@MemberNote,@AmountRent)")

        Dim sqlInsert As String = sb.ToString()

        OpenConnection()

        Dim com As OleDbCommand = New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlInsert
        com.Connection = Conn
        com.Transaction = tr

        Dim strMemberID As String = getMergeMemberID()

        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = strMemberID
        com.Parameters.Add("@MemberName", OleDbType.VarChar).Value = tbMemberName.Text.Trim()
        com.Parameters.Add("@Surname", OleDbType.VarChar).Value = tbSurname.Text.Trim()
        com.Parameters.Add("@Address", OleDbType.VarChar).Value = tbAddress.Text.Trim()
        com.Parameters.Add("@Tel", OleDbType.VarChar).Value = tbTel.Text.Trim()
        com.Parameters.Add("@MemberNote", OleDbType.VarChar).Value = tbMemberNote.Text.Trim()
        com.Parameters.Add("@AmountRent", OleDbType.Integer).Value = 0

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
            bComplete = True
            MessageBox.Show("เพิ่มสมาชิกท่านนี้เรียบร้อยแล้วค่ะ", cstTitle)
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            bComplete = False
            MessageBox.Show("มีสมาชิกที่ตรงกับรหัสนี้แล้วค่ะ กรุณาแก้รหัสสมาชิกใหม่นะคะ", cstWarning)
        End Try

        Return bComplete
    End Function

    'Procudure อัพเดทข้อมูลสมาชิก
    Private Function UpdateToMember() As Boolean
        Dim rowAffect As Integer = 0
        Dim bComplete As Boolean = False

        'สร้างคำสั่ง SQL ในการ Insert 
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("UPDATE Members")
        sb.Append(" SET MemberID=@MemberID,")
        sb.Append("MemberName=@MemberName,")
        sb.Append("Surname=@Surname,")
        sb.Append("Address=@Address,")
        sb.Append("Tel=@Tel,")
        sb.Append("MemberNote=@MemberNote")
        sb.Append(" WHERE MemberID=@MemberID")

        Dim sqlUpdate As String = sb.ToString()

        OpenConnection()

        Dim com As OleDbCommand = New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn
        com.Transaction = tr

        Dim strMemberID As String = getMergeMemberID()

        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = strMemberID
        com.Parameters.Add("@MemberName", OleDbType.VarChar).Value = tbMemberName.Text.Trim()
        com.Parameters.Add("@Surname", OleDbType.VarChar).Value = tbSurname.Text.Trim()
        com.Parameters.Add("@Address", OleDbType.VarChar).Value = tbAddress.Text.Trim()
        com.Parameters.Add("@Tel", OleDbType.VarChar).Value = tbTel.Text.Trim()
        com.Parameters.Add("@MemberNote", OleDbType.VarChar).Value = tbMemberNote.Text.Trim()
        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = strMemberID

        Try
            rowAffect = com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
            If rowAffect = 0 Then
                bComplete = False
                MessageBox.Show("ไม่สามารถอัพเดทข้อมูลสมาชิกได้ค่ะ เนื่องจากไม่มีสมาชิกที่มีรหัสตรงกับรหัสนี้", cstWarning)
            Else
                bComplete = True
                MessageBox.Show("อัพเดทข้อมูลสมาชิกเรียบร้อยแล้วค่ะ", cstTitle)
            End If

        Catch ex As Exception
            bComplete = False
            tr.Rollback()
            CloseConnection()
        End Try

        Return bComplete
    End Function

    'Procedure เอาไว้รับข้อมูลสมาชิกจากตาราง Members และโชว์บนฟอร์ม
    Private Sub ShowEachMembers(ByVal strMemberID As String)
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("SELECT MemberName, Surname, Address, Tel, MemberNote")
        sb.Append(" FROM Members")
        sb.Append(" WHERE MemberID=@MemberID")

        Dim sqlSelect As String = sb.ToString()

        Dim com As OleDbCommand = New OleDbCommand()

        OpenConnection()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = strMemberID

        Dim dr As OleDbDataReader

        Try
            dr = com.ExecuteReader()

            If (dr.HasRows = True) Then
                Dim dt As DataTable = New DataTable()
                dt.Load(dr)

                Dim dataRow As DataRow
                dataRow = dt.Rows(0)

                tbMemberName.Text = dataRow("MemberName").ToString()
                tbSurname.Text = dataRow("Surname").ToString()
                tbAddress.Text = dataRow("Address").ToString()
                tbTel.Text = dataRow("Tel").ToString()
                tbMemberNote.Text = dataRow("MemberNote").ToString()
                setEnableFieldMember(True)
                setReadOnlyFieldMember(True)
                btnEditMember.Enabled = True
                tbMemberID4.Focus()
            Else
                clearForNewMember()
                tbMemberID0.Text = ""
                tbMemberID1.Text = ""
                tbMemberID2.Text = ""
                tbMemberID3.Text = ""
                tbMemberID4.Text = ""
                tbMemberID0.Focus()
                MessageBox.Show("ไม่มีสมาชิกที่ตรงกับรหัสสมาชิกนี้ค่ะ", cstWarning)
            End If
            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnMemberSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMemberSave.Click
        Dim strOut As String = ""

        Dim strMemberID As String = getMergeMemberID()

        If strMemberID < 13 Then
            strOut = strOut + " รหัสสมาชิก"

        End If

        If tbMemberName.Text.Trim() = "" Then
            strOut = strOut + " ชื่อ"
        End If

        If tbSurname.Text.Trim() = "" Then
            strOut = strOut + " นามสกุล"
        End If

        If tbAddress.Text.Trim() = "" Then
            strOut = strOut + " ที่อยู่"
        End If

        If tbTel.Text.Trim() = "" Then
            strOut = strOut + " โทรศัพท์"
        End If

        If strOut <> "" Then
            MessageBox.Show("กรุณากรอก " + strOut + "ของสมาชิกด้วยค่ะ", cstWarning)
            Return
        End If

        Dim bComplete As Boolean = False
        If btnSearchEachMember.Enabled = False Then
            bComplete = InsertToMembers()
        Else
            bComplete = UpdateToMember()
        End If

        If bComplete = False Then
            Return
        Else
            clearForNewMember()
            tbMemberID0.ReadOnly = False
            tbMemberID1.ReadOnly = False
            tbMemberID2.ReadOnly = False
            tbMemberID3.ReadOnly = False
            tbMemberID4.ReadOnly = False
            tbMemberID0.Text = ""
            tbMemberID1.Text = ""
            tbMemberID2.Text = ""
            tbMemberID3.Text = ""
            tbMemberID4.Text = ""
            tbMemberID0.Focus()
            DeleteStarMember()
            btnAddMember.Enabled = True
            btnEditMember.Enabled = False
            btnSearchEachMember.Enabled = True
            btnSearchEachMember.Visible = True
            btnChkDupMemberID.Enabled = False
            btnChkDupMemberID.Visible = False
            btnMemberSave.Enabled = False
            btnMemberClear.Enabled = False
            setEnableFieldMember(False)
        End If
    End Sub

    'Event ที่เกิดจากปุ่มค้นหาสมาชิกแต่ละคน
    Private Sub btnSearchEachMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchEachMember.Click
        Dim strMemberID As String = getMergeMemberID()

        If strMemberID.Length = 0 Then
            MessageBox.Show("คุณยังไม่ได้กรอกรหัสสมาชิกค่ะ")
            tbMemberID0.Focus()
            Return
        ElseIf strMemberID.Length < 13 Then
            MessageBox.Show("คุณกรอกรหัสสมาชิกไม่ครบ 13 ตัวอักษรค่ะ")
            Return
        End If

        ShowEachMembers(strMemberID)
    End Sub

    Private Sub btnDeleteMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteMember.Click
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("DELETE FROM Members WHERE MemberID=@MemberID")

        Dim sqlDelete As String = sb.ToString()

        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlDelete
        com.Connection = Conn
        com.Transaction = tr
        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = getMergeDelMemberID()

        Dim iRowAffect As Integer = 0
        Try
            iRowAffect = com.ExecuteNonQuery()

            tr.Commit()
            CloseConnection()

            If (iRowAffect <> 0) Then
                tbDelMemberID0.Text = ""
                tbDelMemberID1.Text = ""
                tbDelMemberID2.Text = ""
                tbDelMemberID3.Text = ""
                tbDelMemberID4.Text = ""
                tbDelMemberName.Enabled = False
                tbDelMemberSurname.Enabled = False
                tbDelMemberName.Text = ""
                tbDelMemberSurname.Text = ""
                btnDeleteMember.Enabled = False
                MessageBox.Show("ลบสมาชิกเรียบร้อยแล้วค่ะ", cstTitle)
                tbDelMemberID0.Focus()
            Else
                MessageBox.Show("ไม่มีสมาชิกที่ตรงตามรหัสสมาชิกนี้ค่ะ", cstTitle)
            End If
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub tbMemberName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMemberName.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If (tbMemberName.Text.Trim.Length = 0) Then
                MessageBox.Show("กรุณากรอกชื่อสมาชิกด้วยค่ะ", cstWarning)
            Else
                tbSurname.Focus()
            End If
        End If
    End Sub

    Private Sub tbSurname_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbSurname.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If (tbSurname.Text.Trim.Length = 0) Then
                MessageBox.Show("กรุณากรอกนามสกุลสมาชิกด้วยค่ะ", cstWarning)
            Else
                tbAddress.Focus()
            End If
        End If
    End Sub

    Private Sub tbAddress_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbAddress.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If (tbAddress.Text.Trim.Length = 0) Then
                MessageBox.Show("กรุณากรอกที่อยู่ด้วยค่ะ", cstWarning)
            Else
                tbTel.Focus()
            End If
        End If
    End Sub

    Private Sub tbTel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbTel.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If (tbTel.Text.Trim.Length = 0) Then
                MessageBox.Show("กรุณากรอกหมายเลขโทรศัพท์ด้วยค่ะ", cstWarning)
            Else
                tbMemberNote.Focus()
            End If
        End If
    End Sub

    Private Sub btnMemberSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMemberSearch.Click
        Dim bChoose As Boolean = False
        dgvMember.Rows.Clear()

        Dim sqlSelect As String = ""
        sqlSelect = sqlSelect + "SELECT MemberID,MemberName,Surname"
        sqlSelect = sqlSelect + " FROM Members"
        If (cbMemberSerach.SelectedItem = "รหัสสมาชิก") Then
            sqlSelect = sqlSelect + " WHERE MemberID LIKE '" + tbMemberSearch.Text.Trim() + "%'"
            bChoose = True
        ElseIf (cbMemberSerach.SelectedItem = "ชื่อ") Then
            sqlSelect = sqlSelect + " WHERE MemberName LIKE '" + tbMemberSearch.Text.Trim() + "%'"
            bChoose = True
        ElseIf (cbMemberSerach.SelectedItem = "นามสกุล") Then
            sqlSelect = sqlSelect + " WHERE Surname LIKE '" + tbMemberSearch.Text.Trim() + "%'"
            bChoose = True
        End If

        If (bChoose = False) Then
            MessageBox.Show("กรุณาเลือกรูปแบบการค้นหาในช่องค้นหาจากก่อนค่ะ", cstWarning)
            Return
        End If

        sqlSelect = sqlSelect + " ORDER BY MemberID"

        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        Dim dr As OleDbDataReader
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            dr = com.ExecuteReader
            If (dr.HasRows) Then
                Dim dt As New DataTable()
                dt.Load(dr)

                For I As Integer = 0 To dt.Rows.Count - 1
                    dgvMember.Rows.Add()
                    Dim iTarget As Integer = dgvMember.Rows.Count - 2
                    Dim strMemberID As String = dt.Rows(I)("MemberID").ToString()
                    Dim strNewMemberID As String = getHyphenMemberID(strMemberID)

                    dgvMember.Rows(iTarget).Cells(0).Value = strNewMemberID
                    dgvMember.Rows(iTarget).Cells(1).Value = dt.Rows(I)("MemberName")
                    dgvMember.Rows(iTarget).Cells(2).Value = dt.Rows(I)("Surname")

                Next
            End If
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub btnAddMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddMember.Click
        clearForNewMember()
        btnSearchEachMember.Enabled = False
        btnSearchEachMember.Visible = False
        btnChkDupMemberID.Visible = True
        btnChkDupMemberID.Enabled = True
        btnAddMember.Enabled = False
        btnMemberSave.Enabled = True
        btnMemberClear.Enabled = True
        PlusStarMember()
        setEnableFieldMember(False)
        setReadOnlyFieldMember(True)
        tbMemberID0.Text = ""
        tbMemberID1.Text = ""
        tbMemberID2.Text = ""
        tbMemberID3.Text = ""
        tbMemberID4.Text = ""
        tbMemberID0.Focus()
    End Sub

    Private Function getMergeDelMemberID() As String
        Return tbDelMemberID0.Text.Trim() + tbDelMemberID1.Text.Trim() + tbDelMemberID2.Text.Trim() + tbDelMemberID3.Text.Trim() + tbDelMemberID4.Text.Trim()
    End Function

    Private Sub btnDelSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSearchMember.Click
        Dim sb As New StringBuilder()
        sb.Append("SELECT MemberName,Surname")
        sb.Append(" FROM Members")
        sb.Append(" WHERE MemberID=@MemberID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = getMergeDelMemberID()

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                tbDelMemberName.Enabled = True
                tbDelMemberSurname.Enabled = True
                tbDelMemberName.Text = dt.Rows(0)("MemberName").ToString()
                tbDelMemberSurname.Text = dt.Rows(0)("Surname").ToString()

                btnDeleteMember.Enabled = True
            Else
                MessageBox.Show("ไม่พบสมาชิกที่ตรงกับรหัสนี้ค่ะ", cstWarning)
                btnDeleteMember.Enabled = False
                tbDelMemberID0.Text = ""
                tbDelMemberID1.Text = ""
                tbDelMemberID2.Text = ""
                tbDelMemberID3.Text = ""
                tbDelMemberID4.Text = ""
                tbDelMemberName.Enabled = False
                tbDelMemberSurname.Enabled = False
                tbDelMemberName.Text = ""
                tbDelMemberSurname.Text = ""
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub btnMemberClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMemberClear.Click
        clearForNewMember()
        DeleteStarMember()
        btnSearchEachMember.Enabled = True
        btnAddMember.Enabled = True
        btnEditMember.Enabled = False
        btnMemberSave.Enabled = False
        btnMemberClear.Enabled = False
        setEnableFieldMember(False)
        tbMemberID0.Text = ""
        tbMemberID1.Text = ""
        tbMemberID2.Text = ""
        tbMemberID3.Text = ""
        tbMemberID4.Text = ""
        tbMemberID0.ReadOnly = False
        tbMemberID1.ReadOnly = False
        tbMemberID2.ReadOnly = False
        tbMemberID3.ReadOnly = False
        tbMemberID4.ReadOnly = False
        tbMemberID0.Focus()
    End Sub

    Private Sub btnEditMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditMember.Click
        tbMemberID0.ReadOnly = True
        tbMemberID1.ReadOnly = True
        tbMemberID2.ReadOnly = True
        tbMemberID3.ReadOnly = True
        tbMemberID4.ReadOnly = True
        btnAddMember.Enabled = False
        btnEditMember.Enabled = False
        btnMemberSave.Enabled = True
        btnMemberClear.Enabled = True
        setReadOnlyFieldMember(False)
        btnSearchEachMember.Enabled = True
    End Sub

    Private Sub tbDelMemberID0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbDelMemberID0.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            tbDelMemberID0.Focus()
        ElseIf Char.IsNumber(e.KeyChar) = True Then
            tbDelMemberID1.Focus()
        Else
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub tbDelMemberID1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbDelMemberID1.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbDelMemberID1.Text.Length = 0 Then
                tbDelMemberID0.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
            Return
        End If

        If tbDelMemberID1.Text.Length = 3 Then
            tbDelMemberID2.Focus()
        End If
    End Sub

    Private Sub tbDelMemberID2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbDelMemberID2.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbDelMemberID2.Text.Length = 0 Then
                tbDelMemberID1.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbDelMemberID2.Text.Length = 4 Then
            tbDelMemberID3.Focus()
        End If
    End Sub

    Private Sub tbDelMemberID3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbDelMemberID3.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbDelMemberID3.Text.Length = 0 Then
                tbDelMemberID2.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbDelMemberID3.Text.Length = 1 Then
            tbDelMemberID4.Focus()
        End If
    End Sub

    Private Sub tbDelMemberID4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbDelMemberID4.KeyPress

        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbDelMemberID4.Text.Length = 0 Then
                tbDelMemberID3.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub tbDelMemberID0_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelMemberID0.TextChanged
        btnDeleteMember.Enabled = False
    End Sub

    Private Sub tbDelMemberID1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelMemberID1.TextChanged
        btnDeleteMember.Enabled = False
    End Sub

    Private Sub tbDelMemberID2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelMemberID2.TextChanged
        btnDeleteMember.Enabled = False
    End Sub

    Private Sub tbDelMemberID3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelMemberID3.TextChanged
        btnDeleteMember.Enabled = False
    End Sub

    Private Sub tbDelMemberID4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDelMemberID4.TextChanged
        btnDeleteMember.Enabled = False
    End Sub

    Private Sub cbMemberSerach_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbMemberSerach.SelectedIndexChanged
        Select Case cbMemberSerach.SelectedIndex
            Case 0
                tbMemberSearch.Text = ""
                tbMemberSearch.MaxLength = 13
                tbMemberSearch.Focus()
            Case 1
                tbMemberSearch.Text = ""
                tbMemberSearch.Focus()
                tbMemberSearch.MaxLength = 40
            Case 2
                tbMemberSearch.Text = ""
                tbMemberSearch.MaxLength = 40
                tbMemberSearch.Focus()
        End Select
    End Sub
#End Region

#Region "Book"
    '//////////////////////////////////////////////////////เข้าสู่ส่วนจัดการหนังสือ////////////////////////////////////////////////////////////
    'Procedure สำหรับลบข้อมูลทุกฟิลด์ในหน้าจอเพิ่มหนังสือ
    Private Sub clearBookField()
        tbBookID.Text = ""
        tbISBN.Text = ""
        tbBookName.Text = ""
        tbAuthor.Text = ""
        cbType.SelectedItem = -1
        cbType.Text = "กรุณาเลือกหมวดหมู่"
        tbPublisher.Text = ""
        tbPrice.Text = ""
        tbRentPrice.Text = ""
        tbRentDay.Text = ""
        tbFine.Text = ""
        tbBookNote.Text = ""
        btnSearchBook.Visible = True
        btnChkDupBookID.Visible = False
        btnChkDupBookID.Enabled = True
        btnAddBook.Enabled = True
    End Sub

    'Procedure สำหรับใส่ * label Book
    Private Sub PlusStarBook()
        If lbBook1.Text(lbBook1.Text.Length - 1) <> "*" Then
            lbBook1.Text = lbBook1.Text + "*"
            lbBook2.Text = lbBook2.Text + "*"
            lbBook3.Text = lbBook3.Text + "*"
            lbBook4.Text = lbBook4.Text + "*"
            lbBook5.Text = lbBook5.Text + "*"
            lbBook6.Text = lbBook6.Text + "*"
            lbBook7.Text = lbBook7.Text + "*"
        End If
    End Sub

    'Procedure สำหรับลบ * label Book
    Private Sub DeleteStarBook()
        If lbBook1.Text(lbBook1.Text.Length - 1) = "*" Then
            lbBook1.Text = lbBook1.Text.Substring(0, lbBook1.Text.Length - 1)
            lbBook2.Text = lbBook2.Text.Substring(0, lbBook2.Text.Length - 1)
            lbBook3.Text = lbBook3.Text.Substring(0, lbBook3.Text.Length - 1)
            lbBook4.Text = lbBook4.Text.Substring(0, lbBook4.Text.Length - 1)
            lbBook5.Text = lbBook5.Text.Substring(0, lbBook5.Text.Length - 1)
            lbBook6.Text = lbBook6.Text.Substring(0, lbBook6.Text.Length - 1)
            lbBook7.Text = lbBook7.Text.Substring(0, lbBook7.Text.Length - 1)
        End If
    End Sub

    'Procedure สำหรับ set Enable Value Field Book
    Private Sub setEnableFieldBook(ByVal bValue As Boolean)
        tbISBN.Enabled = bValue
        tbBookName.Enabled = bValue
        tbAuthor.Enabled = bValue
        cbType.Enabled = bValue
        tbPublisher.Enabled = bValue
        tbPrice.Enabled = bValue
        tbRentPrice.Enabled = bValue
        tbRentDay.Enabled = bValue
        tbFine.Enabled = bValue
        tbBookNote.Enabled = bValue
    End Sub

    'Procedure สำหรับ set Readonly Value Field Book
    Private Sub setReadOnlyFieldBook(ByVal bValue As Boolean)
        tbISBN.ReadOnly = bValue
        tbBookName.ReadOnly = bValue
        tbAuthor.ReadOnly = bValue
        tbPublisher.ReadOnly = bValue
        tbPrice.ReadOnly = bValue
        tbBookNote.ReadOnly = bValue
    End Sub

    Private Sub tbCancelBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelBook.Click
        clearBookField()
        DeleteStarBook()
        setEnableFieldBook(False)
        setReadOnlyFieldBook(True)
        btnEditBook.Enabled = False
        btnSaveBook.Enabled = False
        btnCancelBook.Enabled = False
        btnCalRent.Enabled = False
        tbBookID.Text = ""
        tbBookID.Enabled = True
        tbBookID.Focus()
    End Sub

    Private Function getCodeType() As String
        Dim strCodeType As String = ""
        If cbType.SelectedItem = "หนังสือการ์ตูน" Then
            strCodeType = "01"
        ElseIf cbType.SelectedItem = "นิยายไทย" Then
            strCodeType = "02"
        ElseIf cbType.SelectedItem = "นิยายแปล" Then
            strCodeType = "03"
        ElseIf cbType.SelectedItem = "พ็อกเก็ตบุ๊ค" Then
            strCodeType = "04"
        Else
            strCodeType = "Not Select"
        End If

        Return strCodeType
    End Function

    Private Function getRentDay(ByVal strCodeType As String) As Integer
        Dim iDayRent As Integer = 0

        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("SELECT DayRent")
        sb.Append(" FROM BookTypes")
        sb.Append(" WHERE TypeCode=@TypeCode")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        com.Parameters.Add("@TypeCode", OleDbType.VarChar).Value = strCodeType

        Dim dr As OleDbDataReader

        Try
            dr = com.ExecuteReader()

            If (dr.HasRows = True) Then
                Dim dt As DataTable = New DataTable()
                dt.Load(dr)

                Dim dataRow As DataRow
                dataRow = dt.Rows(0)

                iDayRent = Integer.Parse(dataRow("DayRent").ToString())

            End If

            dr.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        Return iDayRent
    End Function

    Private Sub btnSearchBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchBook.Click
        If tbBookID.Text.Trim() = "" Then
            MessageBox.Show("กรุณากรอกรหัสหนังสือก่อนค่ะ", cstWarning)
            Return
        ElseIf tbBookID.Text.Trim().Length < 5 Then
            MessageBox.Show("คุณกรอกรหัสหนังสือไม่ครบ 5 ตัวค่ะ", cstWarning)
            Return
        End If

        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("SELECT BookID,ISBN,BookName,Author,TypeCode,Publisher,Price,RentPrice,BookNote")
        sb.Append(" FROM Books")
        sb.Append(" WHERE BookID=@BookID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = tbBookID.Text.Trim()

        Dim dr As OleDbDataReader

        Try
            dr = com.ExecuteReader()

            If (dr.HasRows = True) Then
                Dim dt As DataTable = New DataTable()

                dt.Load(dr)

                Dim dataRow As DataRow = dt.Rows(0)

                dataRow = dt.Rows(0)
                tbBookID.Text = dataRow("BookID").ToString()
                tbISBN.Text = dataRow("ISBN").ToString()
                tbBookName.Text = dataRow("BookName").ToString()
                tbAuthor.Text = dataRow("Author").ToString()
                Dim strType As String = ""
                strType = dataRow("TypeCode").ToString()
                If strType = "01" Then
                    cbType.SelectedItem = "หนังสือการ์ตูน"
                ElseIf strType = "02" Then
                    cbType.SelectedItem = "นิยายไทย"
                ElseIf strType = "03" Then
                    cbType.SelectedItem = "นิยายแปล"
                ElseIf strType = "04" Then
                    cbType.SelectedItem = "พ็อกเก็ตบุ๊ค"
                End If
                tbPublisher.Text = dataRow("Publisher").ToString()
                Dim dbPrice As Double = 0D
                dbPrice = Double.Parse(dataRow("Price").ToString())
                tbPrice.Text = dbPrice.ToString()
                Dim dbRentPrice As Double = 0D
                dbRentPrice = Double.Parse(dataRow("RentPrice").ToString())
                tbRentPrice.Text = dbRentPrice.ToString()
                tbFine.Text = tbRentPrice.Text
                tbBookNote.Text = dataRow("BookNote").ToString()

                sqlSelect = ""
                sqlSelect = sqlSelect + "SELECT DayRent"
                sqlSelect = sqlSelect + " FROM BookTypes"
                sqlSelect = sqlSelect + " WHERE TypeCode = '" + strType + "'"
                com = New OleDbCommand()
                com.CommandType = CommandType.Text
                com.CommandText = sqlSelect
                com.Connection = Conn
                dr = com.ExecuteReader()
                dt = New DataTable()
                dt.Load(dr)
                dataRow = dt.Rows(0)
                Dim iDayRent As Integer = 0
                iDayRent = Integer.Parse(dataRow("DayRent").ToString())
                tbRentDay.Text = iDayRent.ToString()
                setEnableFieldBook(True)
                setReadOnlyFieldBook(True)
                btnEditBook.Enabled = True
            Else
                clearBookField()
                MessageBox.Show("ไม่มีหนังสือที่ตรงกับรหัสหนังสือนี้ค่ะ", cstWarning)
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub tbBookID_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbBookID.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If tbBookID.Text.Trim() = "" Then
                tbBookID.Focus()
                MessageBox.Show("กรุณากรอกรหัสสมาชิกก่อนค่ะ", cstWarning)
            Else
                tbISBN.Focus()
            End If
        End If
    End Sub

    Private Sub tbISBN_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbISBN.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If tbISBN.Text.Trim() = "" Then
                tbISBN.Focus()
                MessageBox.Show("กรุณากรอกหมายเลข ISBN ก่อนค่ะ", cstWarning)
            Else
                tbBookName.Focus()
            End If
        End If
    End Sub

    Private Sub tbBookName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbBookName.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If tbBookName.Text.Trim() = "" Then
                tbBookName.Focus()
                MessageBox.Show("กรุณากรอกหนังสือก่อนค่ะ", cstWarning)
            Else
                tbAuthor.Focus()
            End If
        End If
    End Sub

    Private Sub tbAuthor_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbAuthor.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If tbAuthor.Text.Trim() = "" Then
                tbAuthor.Focus()
                MessageBox.Show("กรุณากรอกผู้แต่งก่อนค่ะ", cstWarning)
            Else
                cbType.Focus()
            End If
        End If
    End Sub

    Private Sub cbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbType.SelectedIndexChanged
        tbPublisher.Focus()
    End Sub

    Private Sub tbPublisher_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbPublisher.KeyPress
        If (e.KeyChar = Convert.ToChar(Keys.Enter) Or e.KeyChar = Convert.ToChar(Keys.Tab)) Then
            If tbPublisher.Text.Trim() = "" Then
                tbPublisher.Focus()
                MessageBox.Show("กรุณากรอกสำนักพิมพ์ก่อนค่ะ", cstWarning)
            Else
                tbPrice.Focus()
            End If
        End If
    End Sub

    'Procedure สำหรับ Insert หนังสือเล่มใหม่สู่ตาราง Books 
    Private Function InsertToBooks() As Boolean
        Dim bComplete As Boolean = False
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("INSERT INTO Books(BookID,ISBN,BookName,Author,TypeCode,Publisher,Price,")
        sb.Append("RentPrice,Rent,BookNote)")
        sb.Append(" VALUES(@BookID,@ISBN,@BookName,@Author,@TypeCode,@Publisher,@Price,")
        sb.Append("@RentPrice,@Rent,@BookNote)")

        Dim sqlAdd As String = sb.ToString()
        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlAdd
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = tbBookID.Text.Trim()
        com.Parameters.Add("@ISBN", OleDbType.VarChar).Value = tbISBN.Text.Trim()
        com.Parameters.Add("@BookName", OleDbType.VarChar).Value = tbBookName.Text.Trim()
        com.Parameters.Add("@Author", OleDbType.VarChar).Value = tbAuthor.Text.Trim()

        Dim strCodeType As String = getCodeType()
        com.Parameters.Add("@TypeCode", OleDbType.VarChar).Value = strCodeType
        com.Parameters.Add("@Publisher", OleDbType.VarChar).Value = tbAuthor.Text.Trim()

        Dim strPrice As String = tbPrice.Text.Trim()
        Dim dbPrice As Double = Double.Parse(strPrice)
        com.Parameters.Add("@Price", OleDbType.Single).Value = dbPrice

        Dim strRentPrice = tbRentPrice.Text.Trim()
        Dim dbRentPrice As Double = Double.Parse(strRentPrice)
        com.Parameters.Add("@RentPrice", OleDbType.Single).Value = dbRentPrice
        com.Parameters.Add("@Rent", OleDbType.Integer).Value = 0
        com.Parameters.Add("@BookNote", OleDbType.VarChar).Value = tbBookNote.Text.Trim()

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
            bComplete = True
            MessageBox.Show("เพิ่มข้อมูลหนังสือเรียบร้อยแล้ว", cstTitle)
        Catch
            tr.Rollback()
            CloseConnection()
            bComplete = False
            MessageBox.Show("มีหนังสือที่่ใช้รหัสนี้อยู่แล้วค่ะ กรุณากรอกรหัสหนังสือใหม่นะคะ", cstWarning)
        End Try
        Return bComplete
    End Function

    'Procedure สำหรับ Update ข้อมูลหนังสือในตาราง Books
    Private Function UpdateBooks() As Boolean
        Dim bComplete As Boolean = False
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("UPDATE Books")
        sb.Append(" SET BookID=@BookID,")
        sb.Append("ISBN=@ISBN,")
        sb.Append("BookName=@BookName,")
        sb.Append("Author=@Author,")
        sb.Append("TypeCode=@TypeCode,")
        sb.Append("Publisher=@Publisher,")
        sb.Append("Price=@Price,")
        sb.Append("RentPrice=@RentPrice,")
        sb.Append("BookNote=@BookNote")
        sb.Append(" WHERE BookID=@BookID")

        Dim sqlUpdate As String = sb.ToString()

        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = tbBookID.Text.Trim()
        com.Parameters.Add("@ISBN", OleDbType.VarChar).Value = tbISBN.Text.Trim()
        com.Parameters.Add("@BookName", OleDbType.VarChar).Value = tbBookName.Text.Trim()
        com.Parameters.Add("@Author", OleDbType.VarChar).Value = tbAuthor.Text.Trim()

        Dim strCodeType As String = getCodeType()
        com.Parameters.Add("@TypeCode", OleDbType.VarChar).Value = strCodeType
        com.Parameters.Add("@Publisher", OleDbType.VarChar).Value = tbAuthor.Text.Trim()

        Dim strPrice As String = tbPrice.Text.Trim()
        Dim dbPrice As Double = Double.Parse(strPrice)
        com.Parameters.Add("@Price", OleDbType.Single).Value = dbPrice

        Dim strRentPrice = tbRentPrice.Text.Trim()
        Dim dbRentPrice As Double = Double.Parse(strRentPrice)
        com.Parameters.Add("@RentPrice", OleDbType.Single).Value = dbRentPrice
        com.Parameters.Add("@BookNote", OleDbType.VarChar).Value = tbBookNote.Text.Trim()
        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = tbBookID.Text.Trim()

        Dim iAffect As Integer = 0
        Try
            iAffect = com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
            If iAffect <> 0 Then
                bComplete = True
                MessageBox.Show("อัพเดทข้อมูลหนังสือเรียบร้อยแล้วค่ะ", cstTitle)
            Else
                bComplete = False
                MessageBox.Show("ไม่สามารถอัพเดทข้อมูลหนังสือได้ค่ะ", cstWarning)
            End If

        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            bComplete = False
            MessageBox.Show("การดำเนินการอัพเดทล้มเหลวเนื่องจาก " + ex.Message, cstWarning)
        End Try

        Return bComplete
    End Function

    Private Sub tbSaveBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveBook.Click
        Dim strOut As String = ""
        If tbBookID.Text.Trim() = "" Then
            strOut = strOut + "รหัสหนังสือ"
        End If

        If tbISBN.Text.Trim() = "" Then
            strOut = strOut + " หมายเลข ISBN"
        End If

        If tbBookName.Text.Trim() = "" Then
            strOut = strOut + " ชื่อหนังสือ"
        End If

        If tbAuthor.Text.Trim() = "" Then
            strOut = strOut + " ผู้แต่ง"
        End If

        If tbPublisher.Text.Trim() = "" Then
            strOut = strOut + " สำนักพิมพ์"
        End If

        If tbPrice.Text.Trim() = "" Then
            strOut = strOut + " ราคา"
        End If

        Dim strTypeCode As String = getCodeType()
        If strTypeCode = "Not Select" Then
            strTypeCode = ""
            strTypeCode = strTypeCode + Environment.NewLine + "กรุณาเลือกประเภทหนังสือด้วยค่ะ"
        End If

        Dim strOut2 As String = ""
        If tbRentPrice.Text.Trim() = "" Or tbRentDay.Text.Trim() = "" Or tbFine.Text.Trim() = "" Then
            strOut2 = Environment.NewLine + "กรุณาคลิกที่ช่องราคาปกและกดปุ่มคำนวณราคาเช่าอีกครั้งค่ะ"
        End If

        If strOut <> "" Then
            MessageBox.Show("กรุณากรอก" + strOut + "ด้วยค่ะ" + strOut2, cstWarning)
            Return
        End If

        Dim bComplete As Boolean = False
        If btnSearchBook.Enabled = False Then
            bComplete = UpdateBooks()
        Else
            bComplete = InsertToBooks()
        End If

        If bComplete = True Then
            clearBookField()
            DeleteStarBook()
            setEnableFieldBook(False)
            setReadOnlyFieldBook(True)
            btnSearchBook.Enabled = True
            btnEditBook.Enabled = False
            btnSaveBook.Enabled = False
            btnCancelBook.Enabled = False
            btnCalRent.Enabled = False
            tbBookID.Text = ""
            tbBookID.ReadOnly = False
            tbBookID.Focus()
        End If

    End Sub

    Private Sub btnBookDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBookDelete.Click
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("DELETE FROM Books WHERE BookID=@BookID")

        Dim sqlDelete As String = sb.ToString()
        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlDelete
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = tbDeleteBook.Text.Trim()

        Dim iRowAffect As Integer

        Try
            iRowAffect = com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()

            If (iRowAffect <> 0) Then
                tbDeleteBook.Text = ""
                tbDelBookName.Text = ""
                tbDelAuthor.Text = ""
                btnBookDelete.Enabled = False
                MessageBox.Show("ลบข้อมูลหนังสือเรียบร้อยแล้วค่ะ", cstTitle)
                tbDeleteBook.Text = ""
                tbDeleteBook.Focus()
            Else
                MessageBox.Show("ไม่มีหนังสือที่ตรงตามรหัสหนังสือนี้ค่ะ", cstTitle)
            End If
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub btnAddBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddBook.Click
        btnSearchBook.Visible = False
        btnChkDupBookID.Visible = True
        btnChkDupBookID.Enabled = True
        btnAddBook.Enabled = False
        PlusStarBook()
        setEnableFieldBook(False)
        setReadOnlyFieldBook(True)
        btnEditBook.Enabled = False
        btnSaveBook.Enabled = True
        btnCancelBook.Enabled = True
        tbBookID.Text = ""
        tbBookID.Enabled = True
        tbBookID.Focus()
    End Sub

    Private Sub btnDelSearchBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSearchBook.Click
        Dim sb As New StringBuilder()
        sb.Append("SELECT BookName,Author")
        sb.Append(" FROM Books")
        sb.Append(" WHERE BookID=@BookID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = tbDeleteBook.Text

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                tbDelBookName.Text = dt.Rows(0)("BookName").ToString()
                tbDelAuthor.Text = dt.Rows(0)("Author").ToString()
                btnBookDelete.Enabled = True
            Else
                btnBookDelete.Enabled = False
                tbDelBookName.Text = ""
                tbDelAuthor.Text = ""
                MessageBox.Show("ไม่พบหนังสือที่ตรงกับรหัสนี้ค่ะ", cstWarning)
                tbDeleteBook.Text = ""
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
        End Try
    End Sub

    Private Sub btnChkDupBookID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChkDupBookID.Click
        If tbBookID.Text.Trim() = "" Then
            MessageBox.Show("คุณยังไม่ได้กรอกรหัสหนังสือเพื่อตรวจสอบค่ะ", cstWarning)
            tbMemberID0.Text = ""
            tbMemberID0.Focus()
            Return
        ElseIf tbBookID.Text.Trim().Length < 5 Then
            MessageBox.Show("คุณกรอกรหัสหนังสือไม่ครบ 5 ตัวค่ะ", cstWarning)
            tbMemberID0.Text = ""
            tbMemberID0.Focus()
            Return
        End If

        Dim sb As New StringBuilder()
        sb.Append("SELECT COUNT(BookID) AS CountBookID")
        sb.Append(" FROM Books")
        sb.Append(" WHERE BookID=@BookID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        com.Parameters.Add("BookID", OleDbType.VarChar).Value = tbBookID.Text

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                Dim iCount = dt.Rows(0)("CountBookID")
                If iCount <> 0 Then
                    MessageBox.Show("มีหนังสือที่ตรงกับรหัสนี้แล้ว กรุณากรอกใหม่อีกครั้งค่ะ")
                    tbBookID.Text = ""
                    tbBookID.Focus()
                Else
                    tbBookID.ReadOnly = True
                    setReadOnlyFieldBook(False)
                    setEnableFieldBook(True)
                    btnChkDupBookID.Enabled = False
                    btnCalRent.Enabled = True
                    tbISBN.Focus()
                End If

            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub btnCalRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalRent.Click
        If tbPrice.Text.Trim() = "" Then
            tbPrice.Focus()
            MessageBox.Show("กรุณากรอกราคาหนังสือก่อนค่ะ", cstWarning)
            Return
        Else
            Dim dbPrice = 0D
            Try
                dbPrice = Double.Parse(tbPrice.Text)
            Catch
                MessageBox.Show("ค่าเช่าต้องเป็นตัวเลขเท่านั้นค่ะ")
                Return
            End Try

            If dbPrice = 0 Or dbPrice < 0 Then
                tbPrice.Text = ""
                tbPrice.Focus()
                MessageBox.Show("ราคาหนังสือต้องมีค่ามากกว่า 0 บาทค่ะ", cstWarning)
            Else

                Dim strCodeType As String = ""
                strCodeType = getCodeType()

                If strCodeType = "Not Select" Then
                    cbType.Focus()
                    MessageBox.Show("กรุณาเลือกประเภทหนังสือด้วยค่ะ", cstWarning)
                    Return
                Else
                    dbPrice = dbPrice * 0.1
                    dbPrice = Math.Ceiling(dbPrice)
                    tbRentPrice.Text = dbPrice.ToString()
                    tbFine.Text = dbPrice.ToString()

                    'Dim iDayRent As Integer = 0
                    'iDayRent = getRentDay(strCodeType)
                    tbRentDay.Text = "1"

                    tbBookNote.Focus()
                    CloseConnection()
                End If

            End If
        End If
    End Sub

    Private Sub btnEditBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditBook.Click
        btnSearchBook.Enabled = False
        btnAddBook.Enabled = False
        btnEditBook.Enabled = False
        btnSaveBook.Enabled = True
        btnCancelBook.Enabled = True
        tbBookID.ReadOnly = True
        setReadOnlyFieldBook(False)
        PlusStarBook()
        btnCalRent.Enabled = True
    End Sub

    Private Sub btnBookSearch_Click(sender As System.Object, e As System.EventArgs) Handles btnBookSearch.Click
        Dim bChoose As Boolean = False
        dgvBook.Rows.Clear()

        Dim sqlSelect As String = ""
        sqlSelect = sqlSelect + "SELECT BookID,ISBN,BookName,Author,Publisher,TypeCode"
        sqlSelect = sqlSelect + " FROM Books"
        If (cbSearchBook.SelectedItem = "รหัสหนังสือ") Then
            sqlSelect = sqlSelect + " WHERE BookID LIKE '" + tbBookSearch.Text.Trim() + "%'"
            bChoose = True
        ElseIf (cbSearchBook.SelectedItem = "ISBN") Then
            sqlSelect = sqlSelect + " WHERE ISBN LIKE '" + tbBookSearch.Text.Trim() + "%'"
            bChoose = True
        ElseIf (cbSearchBook.SelectedItem = "ชื่อหนังสือ") Then
            sqlSelect = sqlSelect + " WHERE BookName LIKE '" + tbBookSearch.Text.Trim() + "%'"
            bChoose = True
        ElseIf (cbSearchBook.SelectedItem = "ผู้แต่ง") Then
            sqlSelect = sqlSelect + " WHERE Author LIKE '" + tbBookSearch.Text.Trim() + "%'"
            bChoose = True
        ElseIf (cbSearchBook.SelectedItem = "สำนักพิมพ์") Then
            sqlSelect = sqlSelect + " WHERE Publisher LIKE '" + tbBookSearch.Text.Trim() + "%'"
            bChoose = True
        End If

        If (bChoose = False) Then
            MessageBox.Show("กรุณาเลือกรูปแบบการค้นหาในช่องค้นหาจากก่อนค่ะ", cstWarning)
            Return
        End If

        sqlSelect = sqlSelect + " ORDER BY BookID"

        OpenConnection()
        Dim com As OleDbCommand = New OleDbCommand()
        Dim dr As OleDbDataReader
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            dr = com.ExecuteReader
            If (dr.HasRows) Then
                Dim dt As New DataTable()
                dt.Load(dr)

                For I As Integer = 0 To dt.Rows.Count - 1
                    dgvBook.Rows.Add()
                    Dim iTarget As Integer = dgvBook.Rows.Count - 2

                    dgvBook.Rows(iTarget).Cells(0).Value = dt.Rows(I)("BookID")
                    dgvBook.Rows(iTarget).Cells(1).Value = dt.Rows(I)("ISBN")
                    dgvBook.Rows(iTarget).Cells(2).Value = dt.Rows(I)("BookName")
                    dgvBook.Rows(iTarget).Cells(3).Value = dt.Rows(I)("Author")
                    dgvBook.Rows(iTarget).Cells(4).Value = dt.Rows(I)("Publisher")
                    dgvBook.Rows(iTarget).Cells(5).Value = dt.Rows(I)("TypeCode")

                Next
            End If
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub cbSearchBook_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbSearchBook.SelectedIndexChanged
        Select Case cbSearchBook.SelectedIndex
            Case 0
                tbBookSearch.Text = ""
                tbBookSearch.MaxLength = 5
                tbBookSearch.Focus()
            Case 1
                tbBookSearch.Text = ""
                tbBookSearch.MaxLength = 13
                tbBookSearch.Focus()
            Case 2
                tbBookSearch.Text = ""
                tbBookSearch.MaxLength = 100
                tbBookSearch.Focus()
            Case 3
                tbBookSearch.Text = ""
                tbBookSearch.MaxLength = 40
                tbBookSearch.Focus()
            Case 4
                tbBookSearch.Text = ""
                tbBookSearch.MaxLength = 40
                tbBookSearch.Focus()
        End Select
    End Sub
#End Region

#Region "RentBook"
    '///////////////////////////////////////////////////เข้าสู่ส่วนจัดการระบบเช่า///////////////////////////////////
    Private Sub tbRentMemberID0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbRentMemberID0.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            tbRentMemberID0.Focus()
        ElseIf Char.IsNumber(e.KeyChar) = True Then
            tbRentMemberID1.Focus()
        Else
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub tbRentMemberID1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbRentMemberID1.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbRentMemberID1.Text.Length = 0 Then
                tbRentMemberID0.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
            Return
        End If

        If tbRentMemberID1.Text.Length = 3 Then
            tbRentMemberID2.Focus()
        End If
    End Sub

    Private Sub tbRentMemberID2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbRentMemberID2.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbRentMemberID2.Text.Length = 0 Then
                tbRentMemberID1.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbRentMemberID2.Text.Length = 4 Then
            tbRentMemberID3.Focus()
        End If
    End Sub

    Private Sub tbRentMemberID3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbRentMemberID3.KeyPress

        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbRentMemberID3.Text.Length = 0 Then
                tbRentMemberID2.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbRentMemberID3.Text.Length = 1 Then
            tbRentMemberID4.Focus()
        End If
    End Sub

    Private Sub tbRentMemberID4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbRentMemberID4.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbRentMemberID4.Text.Length = 0 Then
                tbRentMemberID3.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub clearForNewRent()
        tbRentMemberName.Enabled = False
        tbRentMemberName.Text = ""
        tbRentMemberSurname.Enabled = False
        tbRentMemberSurname.Text = ""
        tbPastRent1.Enabled = False
        tbPastRent1.Text = ""
        tbRentBookID.Text = ""
        tbPastRent1.Text = ""
        tbRentBookID.Enabled = False
        tbRentBookID.Text = ""
        tbRentBookID.Enabled = False
        tbRentBookName.Text = ""
        tbRentAuthor.Text = ""
        tbRentAuthor.Enabled = False
        btnSearchRentBook.Enabled = False
        btnAddRent.Enabled = False
        tbRentAmount.Enabled = False
        tbRentTotalPrice.Enabled = False

        dgvRent.Rows.Clear()

        tbRentAmount.Text = "0"
        tbRentTotalPrice.Text = "0"

        btnSaveRent.Enabled = False
        btnCancelRentAll.Enabled = False
        btnCalcelEachRent.Enabled = False
        tbMemberID0.Focus()
    End Sub

    Private Function getMergeRentMemberID() As String
        Return tbRentMemberID0.Text.Trim() + tbRentMemberID1.Text.Trim() + tbRentMemberID2.Text.Trim() + tbRentMemberID3.Text.Trim() + tbRentMemberID4.Text.Trim()
    End Function

    Private Sub btnRentSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRentSearchMember.Click

        clearForNewRent()

        Dim strMemberID As String = getMergeRentMemberID()

        Select Case strMemberID.Length
            Case 0
                MessageBox.Show("คุณยังไม่ได้กรอกรหัสสมาชิกที่ใช้ในการค้นหาค่ะ", cstWarning)
                Return
            Case 1 To 12
                MessageBox.Show("คุณกรอกรหัสสมาชิกไม่ครบ 13 ตัวค่ะ", cstWarning)
                Return
        End Select

        Dim sb As New StringBuilder()
        sb.Append("SELECT MemberName,Surname,AmountRent")
        sb.Append(" FROM Members")
        sb.Append(" WHERE MemberID=@MemberID")

        Dim sqlSelect As String = sb.ToString()
        OpenConnection()

        Dim com As New OleDbCommand()
        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = strMemberID

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()
            If (dr.HasRows) Then
                Dim dt As New DataTable()
                dt.Load(dr)
                Dim dataRow As DataRow = dt.Rows(0)

                tbRentMemberName.Text = dataRow("MemberName").ToString()
                tbRentMemberSurname.Text = dataRow("Surname").ToString()
                tbPastRent1.Text = dataRow("AmountRent").ToString()
                tbRentMemberName.Enabled = True
                tbRentMemberSurname.Enabled = True
                tbPastRent1.Enabled = True
                tbRentBookID.Enabled = True
                btnSearchRentBook.Enabled = True
                btnCancelRentAll.Enabled = True
                tbRentAmount.Enabled = True
                tbRentTotalPrice.Enabled = True
                tbRentBookID.Focus()
            Else
                MessageBox.Show("ไม่มีสมาชิกที่ตรงกับรหัสสมาชิกนี้ค่ะ" + Environment.NewLine + "กรุณากรอกใหม่อีกครั้งค่ะ", cstWarning)
                tbRentMemberID0.Text = ""
                tbRentMemberID1.Text = ""
                tbRentMemberID2.Text = ""
                tbRentMemberID3.Text = ""
                tbRentMemberID4.Text = ""
                tbRentMemberName.Text = ""
                tbRentMemberSurname.Text = ""
                tbPastRent1.Text = ""
                tbRentMemberID0.Focus()
            End If
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub btnRentOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Function checkRentDuplicate(ByVal strRentBookID As String) As Boolean
        Dim bCheck As Boolean = False

        For I As Integer = 0 To dgvRent.Rows.Count - 2
            If dgvRent.Rows(I).Cells(1).Value = strRentBookID Then
                bCheck = True
            End If
        Next

        Return bCheck
    End Function

    Private Sub btnSearchRentBook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchRentBook.Click
        Select Case tbRentBookID.Text.Trim().Length
            Case 0
                MessageBox.Show("คุณยังไม่ได้กรอกหนังสือที่จะให้เช่าค่ะ", cstWarning)
                Return
            Case 1 To 4
                MessageBox.Show("คุณกรอกรหัสหนังสือไม่ครบ 5 ตัวค่ะ", cstWarning)
                Return
        End Select

        Dim bCheckDuplicate As Boolean = checkRentDuplicate(tbRentBookID.Text.Trim())

        If bCheckDuplicate = True Then
            MessageBox.Show("ขออภัยค่ะ หนังสือที่ตรงกับรหัสหนังสือนี้มีอยู่ในรายการที่จะให้เช่าแล้วค่ะ", cstWarning)
            Return
        End If

        Dim sqlSelect As String = ""
        Dim com As OleDbCommand
        Dim dr As OleDbDataReader
        Dim dt As DataTable
        Dim dataRow As DataRow

        sqlSelect = ""
        sqlSelect = sqlSelect + "SELECT BookName,Author,Rent"
        sqlSelect = sqlSelect + " FROM Books"
        sqlSelect = sqlSelect + " WHERE BookID = '" + tbRentBookID.Text.Trim() + "'"
        OpenConnection()
        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        Try
            dr = com.ExecuteReader()
            If dr.HasRows = True Then
                dt = New DataTable()
                dt.Load(dr)
                dataRow = dt.Rows(0)

                tbRentBookName.Text = dt.Rows(0)("BookName").ToString()
                tbRentBookName.Enabled = True
                tbRentAuthor.Text = dt.Rows(0)("Author").ToString()
                tbRentAuthor.Enabled = True
                Dim strAvilable As String = dt.Rows(0)("Rent").ToString()
                If strAvilable = "False" Then
                    btnAddRent.Enabled = True
                Else
                    btnAddRent.Enabled = False
                    MessageBox.Show("หนังสือเล่มนี้ถูกเช่าแล้วค่ะ", cstWarning)
                    tbRentBookID.Text = ""
                    tbRentBookName.Text = ""
                    tbRentAuthor.Text = ""
                    tbRentBookID.Focus()
                End If
            Else
                tbRentBookID.Text = ""
                tbRentBookID.Focus()
                MessageBox.Show("ไม่มีหนังสือที่ตรงกับรหัสหนังสือนี้ค่ะ" + Environment.NewLine + "กรุณากรอกใหม่อีกครั้งค่ะ", cstWarning)
                tbRentBookName.Text = ""
                tbRentBookName.Enabled = False
                tbRentAuthor.Text = ""
                tbRentAuthor.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub addTodgvRent(ByRef bookRent As BookRent)
        dgvRent.Rows.Add()
        dgvRent.Rows(dgvRent.Rows.Count - 1).Cells(0).Value = False
        dgvRent.Rows(dgvRent.Rows.Count - 1).Cells(1).Value = bookRent.BookID
        dgvRent.Rows(dgvRent.Rows.Count - 1).Cells(2).Value = bookRent.BookName
        dgvRent.Rows(dgvRent.Rows.Count - 1).Cells(3).Value = bookRent.RentPrice

    End Sub

    Private Sub btnAddRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddRent.Click

        Dim sqlSelect As String = ""
        Dim com As OleDbCommand
        Dim dr As OleDbDataReader
        Dim dt As DataTable
        Dim dataRow As DataRow

        sqlSelect = ""
        sqlSelect = sqlSelect + "SELECT BookID,BookName,RentPrice"
        sqlSelect = sqlSelect + " FROM Books"
        sqlSelect = sqlSelect + " WHERE BookID = '" + tbRentBookID.Text.Trim() + "'"
        OpenConnection()
        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        Try
            dr = com.ExecuteReader()
            If dr.HasRows = True Then
                dt = New DataTable()
                dt.Load(dr)
                dataRow = dt.Rows(0)

                Dim bookRent As BookRent = New BookRent()
                bookRent.Init()
                bookRent.BookID = dataRow("BookID").ToString()
                bookRent.BookName = dataRow("BookName").ToString()
                bookRent.RentPrice = Double.Parse(dataRow("RentPrice").ToString())

                Dim iPastRent As Integer = Integer.Parse(tbPastRent1.Text.Trim().ToString())
                Dim iWillRent As Integer = Integer.Parse(tbRentAmount.Text.Trim().ToString())
                Dim iPresentRent As Integer = iPastRent + iWillRent

                If iPresentRent < maxRent Then
                    addTodgvRent(bookRent)
                    IncreaseAmountRent()
                    PlusRentTotalPrice(bookRent.RentPrice)
                    btnSaveRent.Enabled = True
                    btnCancelRentAll.Enabled = True
                    tbRentBookName.Text = ""
                    tbRentBookName.Enabled = False
                    tbRentAuthor.Text = ""
                    tbRentAuthor.Enabled = False
                    tbRentBookID.Text = ""
                    tbRentBookID.Focus()
                    btnCancelRentAll.Enabled = True
                    btnCalcelEachRent.Enabled = True
                Else
                    MessageBox.Show("ลูกค้าไม่สามารถยืมหนังสือได้เกิน 5 เล่มค่ะ", cstWarning)
                End If
            Else
                tbRentBookName.Text = ""
                tbRentBookName.Enabled = False
                tbRentAuthor.Text = ""
                tbRentAuthor.Enabled = False
                tbRentBookID.Text = ""
                tbRentBookID.Focus()
                MessageBox.Show("ไม่มีหนังสือที่ตรงกับรหัสหนังสือนี้ค่ะ" + Environment.NewLine + "กรุณากรอกใหม่อีกครั้งค่ะ", cstWarning)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub insertToRentNote(ByRef rentDeatil As RentDetail, ByRef memberRent As Member)

        OpenConnection()

        For I As Integer = 0 To dgvRent.Rows.Count - 1

            'Database Variable
            Dim iRecordID As String
            Dim sqlSelect As String = "SELECT Last(RecordID) AS LastRecordID FROM RentNote"
            Dim sqlInsert As String
            Dim com As OleDbCommand
            Dim dr As OleDbDataReader
            Dim dt As DataTable
            Dim dataRow As DataRow

            com = New OleDbCommand
            com.CommandType = CommandType.Text
            com.CommandText = sqlSelect
            com.Connection = Conn
            Try
                dr = com.ExecuteReader()
                dt = New DataTable()
                dt.Load(dr)
                dataRow = dt.Rows(0)

                If dataRow("LastRecordID").ToString = "" Then
                    iRecordID = 0
                Else
                    iRecordID = Integer.Parse(dataRow("LastRecordID").ToString())

                End If
                iRecordID = iRecordID + 1
                sqlInsert = ""
                sqlInsert = sqlInsert + "INSERT INTO RentNote(RecordID,MemberID,MemberName,Surname,RentDate,ReturnStatus,BookID,BookName,Rent)"
                sqlInsert = sqlInsert + " VALUES(" + iRecordID + ","
                sqlInsert = sqlInsert + "'" + memberRent.MemberID + "',"
                sqlInsert = sqlInsert + "'" + memberRent.MemberName + "',"
                sqlInsert = sqlInsert + "'" + memberRent.Surname + "',"
                sqlInsert = sqlInsert + rentDeatil.rentDate + ","
                'sqlInsert = sqlInsert + rentDeatil.returnDate + ","
                sqlInsert = sqlInsert + "0,"
                sqlInsert = sqlInsert + "'" + dgvRent.Rows(I).Cells(1).Value + "',"
                sqlInsert = sqlInsert + "'" + dgvRent.Rows(I).Cells(2).Value + "',"
                Dim dbRentPrice As Double = Double.Parse(dgvRent.Rows(I).Cells(3).Value.ToString())
                sqlInsert = sqlInsert + dbRentPrice.ToString()
                sqlInsert = sqlInsert + ")"

                com = New OleDbCommand
                com.CommandType = CommandType.Text
                com.CommandText = sqlInsert
                com.Connection = Conn
                com.ExecuteNonQuery()
            Catch ex As Exception
                CloseConnection()
                MessageBox.Show(ex.Message, cstWarning)
            End Try

        Next
        CloseConnection()
    End Sub

    Private Sub updateForRentToMembers(ByRef rentDeatil As RentDetail, ByRef memberRent As Member)
        Dim newAmountRent As Integer = memberRent.PastRent + rentDeatil.amountRent
        Dim sqlUpdate As String = ""
        sqlUpdate = sqlUpdate + "Update Members"
        sqlUpdate = sqlUpdate + " SET AmountRent = " + newAmountRent.ToString()
        sqlUpdate = sqlUpdate + " WHERE MemberID = '" + memberRent.MemberID + "'"
        Dim com As OleDbCommand = New OleDbCommand()
        OpenConnection()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn

        Try
            com.ExecuteNonQuery()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub saveStatRent(ByVal rentDate As String, ByVal dbTotalPlus As Double)
        Dim sqlSelect As String
        Dim com As OleDbCommand
        Dim dr As OleDbDataReader

        sqlSelect = ""
        sqlSelect = sqlSelect + "SELECT Last(RecordID) AS LastRecordID FROM DayDetail"

        OpenConnection()

        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            dr = com.ExecuteReader()
            If dr.HasRows = True Then
                Dim dt As DataTable = New DataTable()
                Dim dataRow As DataRow
                Dim iRecordID As Integer = 0
                dt.Load(dr)
                dataRow = dt.Rows(0)

                'กรณีที่ไม่มีข้อมูลอยู่เลย
                If dataRow("LastRecordID").ToString = "" Then
                    iRecordID = 1
                    Dim sqlInsert As String = ""
                    sqlInsert = sqlInsert + "INSERT INTO DayDetail(RecordID,DateDetail,AmountMember,RentIncome,FineIncome)"
                    sqlInsert = sqlInsert + " VALUES("
                    sqlInsert = sqlInsert + iRecordID.ToString() + ","
                    sqlInsert = sqlInsert + rentDate + ","
                    sqlInsert = sqlInsert + "1,"
                    sqlInsert = sqlInsert + dbTotalPlus.ToString() + ","
                    sqlInsert = sqlInsert + "0"
                    sqlInsert = sqlInsert + ")"

                    com = New OleDbCommand()
                    com.CommandType = CommandType.Text
                    com.CommandText = sqlInsert
                    com.Connection = Conn
                    com.ExecuteNonQuery()

                Else
                    'ในกรณีที่มีข้อมูลอยู่อย่างน้อย 1 Record
                    sqlSelect = ""
                    sqlSelect = sqlSelect + "SELECT RecordID,DateDetail,AmountMember,RentIncome"
                    sqlSelect = sqlSelect + " FROM DayDetail"
                    sqlSelect = sqlSelect + " WHERE DateDetail = " + rentDate

                    com = New OleDbCommand()
                    com.CommandType = CommandType.Text
                    com.CommandText = sqlSelect
                    com.Connection = Conn

                    dr = com.ExecuteReader()
                    If dr.HasRows = True Then
                        dt = New DataTable()
                        dt.Load(dr)
                        dataRow = dt.Rows(0)
                        Dim iAmountMember As Integer = Integer.Parse(dataRow("AmountMember").ToString())
                        iAmountMember = iAmountMember + 1
                        Dim dbTotalIncome As Double = Double.Parse(dataRow("RentIncome").ToString())
                        dbTotalIncome = dbTotalIncome + dbTotalPlus
                        Dim sqlUpdate As String = ""
                        sqlUpdate = sqlUpdate + "UPDATE DayDetail"
                        sqlUpdate = sqlUpdate + " SET AmountMember = " + iAmountMember.ToString() + ","
                        sqlUpdate = sqlUpdate + " RentIncome = " + dbTotalIncome.ToString()
                        sqlUpdate = sqlUpdate + " WHERE DateDetail = " + rentDate

                        com = New OleDbCommand()
                        com.CommandType = CommandType.Text
                        com.CommandText = sqlUpdate
                        com.Connection = Conn
                        com.ExecuteNonQuery()
                    Else
                        iRecordID = Integer.Parse(dataRow("LastRecordID").ToString())
                        iRecordID = iRecordID + 1

                        Dim sqlInsert As String = ""
                        sqlInsert = sqlInsert + "INSERT INTO DayDetail(RecordID,DateDetail,AmountMember,RentIncome,FineIncome)"
                        sqlInsert = sqlInsert + " VALUES("
                        sqlInsert = sqlInsert + iRecordID.ToString() + ","
                        sqlInsert = sqlInsert + rentDate + ","
                        sqlInsert = sqlInsert + "1,"
                        sqlInsert = sqlInsert + dbTotalPlus.ToString() + ","
                        sqlInsert = sqlInsert + "0"
                        sqlInsert = sqlInsert + ")"

                        com = New OleDbCommand()
                        com.CommandType = CommandType.Text
                        com.CommandText = sqlInsert
                        com.Connection = Conn
                        com.ExecuteNonQuery()
                    End If

                End If
            End If
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        CloseConnection()

    End Sub

    Private Sub rentToBook()


        For I As Integer = 0 To dgvRent.Rows.Count - 1

            Dim sb As New StringBuilder()
            sb.Append("UPDATE Books")
            sb.Append(" SET Rent=@Rent")
            sb.Append(" WHERE BookID=@BookID")

            Dim sqlUpdate As String = sb.ToString()

            OpenConnection()
            Dim com As New OleDbCommand()
            Dim tr As OleDbTransaction = Conn.BeginTransaction
            com.CommandType = CommandType.Text
            com.CommandText = sqlUpdate
            com.Connection = Conn
            com.Transaction = tr

            com.Parameters.Add("@Rent", OleDbType.Integer).Value = "1"
            Dim strBoookID As String = dgvRent.Rows(I).Cells(1).Value
            com.Parameters.Add("@BookID", OleDbType.VarChar).Value = strBoookID

            Try
                com.ExecuteNonQuery()
                tr.Commit()
                CloseConnection()
            Catch ex As Exception
                tr.Rollback()
                CloseConnection()
                MessageBox.Show(ex.Message, cstWarning)
            End Try
        Next

    End Sub

    Private Sub btnSaveRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveRent.Click
        Dim memberRent As Member = New Member()
        memberRent.Init()
        memberRent.MemberID = getMergeRentMemberID()
        memberRent.MemberName = tbRentMemberName.Text
        memberRent.Surname = tbRentMemberSurname.Text
        memberRent.PastRent = Integer.Parse(tbPastRent1.Text.Trim())

        Dim rentDetail As RentDetail = New RentDetail()
        rentDetail.Init()
        rentDetail.setToPresentDate()
        rentDetail.amountRent = Integer.Parse(tbRentAmount.Text.Trim())
        rentDetail.totalPrice = Double.Parse(tbRentTotalPrice.Text.Trim())

        insertToRentNote(rentDetail, memberRent)
        rentToBook()
        updateForRentToMembers(rentDetail, memberRent)
        saveStatRent(rentDetail.rentDate, rentDetail.totalPrice)

        MessageBox.Show("บันทึกการเช่าหนังสือเรียบร้อยแล้วค่ะ", cstTitle)
        tbRentMemberID0.Text = ""
        tbRentMemberID1.Text = ""
        tbRentMemberID2.Text = ""
        tbRentMemberID3.Text = ""
        tbRentMemberID4.Text = ""
        clearForNewRent()
        tbRentMemberID0.Focus()
    End Sub

    Private Sub IncreaseAmountRent()
        Dim iAmount As Integer = Integer.Parse(tbRentAmount.Text.Trim())
        iAmount = iAmount + 1
        tbRentAmount.Text = iAmount.ToString()
    End Sub

    Private Sub PlusRentTotalPrice(ByVal dbPrice As Double)
        Dim dbTotalPrice As Double = Double.Parse(tbRentTotalPrice.Text.Trim())
        dbTotalPrice = dbTotalPrice + dbPrice
        tbRentTotalPrice.Text = dbTotalPrice.ToString()
    End Sub

    Private Sub btnCancelRentAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelRentAll.Click
        tbRentMemberID0.Text = ""
        tbRentMemberID1.Text = ""
        tbRentMemberID2.Text = ""
        tbRentMemberID3.Text = ""
        tbRentMemberID4.Text = ""
        clearForNewRent()
        tbRentMemberID0.Focus()
    End Sub

    Private Sub btnCalcelEachRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalcelEachRent.Click
        Dim iIndex As Integer = 0
        Dim iLength As Integer = dgvRent.Rows.Count

        While iIndex < iLength
            If dgvRent.Rows(iIndex).Cells(0).Value.ToString() = "True" Then

                Dim oldRentAmount As Integer = Integer.Parse(tbRentAmount.Text)
                tbRentAmount.Text = (oldRentAmount - 1).ToString()
                Dim oldTotalPrice As Double = Double.Parse(tbRentTotalPrice.Text)
                Dim tempPrice As Double = Double.Parse(dgvRent.Rows(iIndex).Cells(3).Value)
                Dim newTotalPrice As Double = oldTotalPrice - tempPrice
                tbRentTotalPrice.Text = newTotalPrice.ToString()

                dgvRent.Rows.RemoveAt(iIndex)
                iLength = iLength - 1

            Else
                iIndex = iIndex + 1
            End If
        End While
    End Sub

#End Region

#Region "ReturnBook"
    '//////////////////////////////////////////////เข้าสู่ส่วนจัดการระบบคืน////////////////////////////////////////////////
    Private Sub tbReturnMemberID0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbReturnMemberID0.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            tbReturnMemberID0.Focus()
        ElseIf Char.IsNumber(e.KeyChar) = True Then
            tbReturnMemberID1.Focus()
        Else
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub tbReturnMemberID1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbReturnMemberID1.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbReturnMemberID1.Text.Length = 0 Then
                tbReturnMemberID0.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
            Return
        End If

        If tbReturnMemberID1.Text.Length = 3 Then
            tbReturnMemberID2.Focus()
        End If
    End Sub

    Private Sub tbReturnMemberID2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbReturnMemberID2.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbReturnMemberID2.Text.Length = 0 Then
                tbReturnMemberID1.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbReturnMemberID2.Text.Length = 4 Then
            tbReturnMemberID3.Focus()
        End If
    End Sub

    Private Sub tbReturnMemberID3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbReturnMemberID3.KeyPress

        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbReturnMemberID3.Text.Length = 0 Then
                tbReturnMemberID2.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If

        If tbReturnMemberID3.Text.Length = 1 Then
            tbReturnMemberID4.Focus()
        End If
    End Sub

    Private Sub tbReturnMemberID4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbReturnMemberID4.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Back) Then
            If tbReturnMemberID4.Text.Length = 0 Then
                tbReturnMemberID3.Focus()
            End If
            Return
        ElseIf Char.IsNumber(e.KeyChar) = False Then
            e.Handled = True
            MessageBox.Show("กรุณาใส่เฉพาะตัวเลข 0-9 เท่านั้นค่ะ")
        End If
    End Sub

    Private Sub clearForNewReturn()
        dgvReturn.Rows.Clear()
        btnReturnSearchMember.Enabled = False
        tbReturnMemberName.Enabled = False
        tbReturnMemberSurname.Enabled = False
        tbPastRent2.Enabled = False
        dgvReturn.Enabled = False
        tbAmountReturn.Enabled = False
        tbTotalFine.Enabled = False
        btnCancelAllReturn.Enabled = False
        btnSaveReturn.Enabled = False

        tbReturnMemberName.Text = ""
        tbReturnMemberSurname.Text = ""
        tbPastRent2.Text = ""
        tbAmountReturn.Text = "0"
        tbTotalFine.Text = "0"

        For I As Integer = 0 To dgvReturn.Rows.Count - 2
            dgvReturn.Rows.RemoveAt(0)
        Next

        dgvReturn.Enabled = True

        tbReturnMemberID0.Enabled = True
        btnReturnSearchMember.Enabled = True

    End Sub

    Private Sub showReturnMemberDetail(ByRef member As Member)
        Dim com As OleDbCommand
        Dim dr As OleDbDataReader
        Dim dt As DataTable
        Dim dataRow As DataRow
        Dim sqlSelect As String

        sqlSelect = ""
        sqlSelect = sqlSelect + "SELECT MemberName,Surname,AmountRent"
        sqlSelect = sqlSelect + " FROM Members"
        sqlSelect = sqlSelect + " WHERE MemberID = '" + member.MemberID + "'"

        OpenConnection()

        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            dr = com.ExecuteReader()
            If (dr.HasRows) Then
                dt = New DataTable()
                dt.Load(dr)
                dataRow = dt.Rows(0)

                member.MemberName = dataRow("MemberName").ToString()
                tbReturnMemberName.Text = member.MemberName

                member.Surname = dataRow("Surname").ToString()
                tbReturnMemberSurname.Text = member.Surname

                member.PastRent = Integer.Parse(dataRow("AmountRent").ToString())
                If member.PastRent = 0 Then
                    btnCancelAllReturn.Enabled = False
                    btnSaveReturn.Enabled = False
                    tbTotalFine.Enabled = False
                    tbAmountReturn.Enabled = False
                Else
                    btnCancelAllReturn.Enabled = True
                    btnSaveReturn.Enabled = True
                    tbTotalFine.Enabled = True
                    tbAmountReturn.Enabled = True
                End If
                tbPastRent2.Text = member.PastRent.ToString()
                tbReturnMemberName.Enabled = True
                tbReturnMemberSurname.Enabled = True
                tbPastRent2.Enabled = True
                tbReturnMemberID4.Focus()
            Else
                MessageBox.Show("ไม่มีสมาชิกที่ตรงกับรหัสสมาชิกนี้ค่ะ" + Environment.NewLine + "กรุณากรอกใหม่อีกครั้งค่ะ", cstWarning)
                tbReturnMemberID0.Text = ""
                tbReturnMemberID1.Text = ""
                tbReturnMemberID2.Text = ""
                tbReturnMemberID3.Text = ""
                tbReturnMemberID4.Text = ""
                tbReturnMemberName.Text = ""
                tbReturnMemberSurname.Text = ""
                tbPastRent2.Text = ""
                tbReturnMemberID0.Focus()
            End If

            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub addTodgvReturn(ByRef bookReturn As BookBeforeReturn)
        dgvReturn.Rows.Add()
        Dim iTarget As Integer = dgvReturn.Rows.Count - 1
        dgvReturn.Rows(iTarget).Cells(0).Value = False
        dgvReturn.Rows(iTarget).Cells(1).Value = bookReturn.BookID
        dgvReturn.Rows(iTarget).Cells(2).Value = bookReturn.BookName
        dgvReturn.Rows(iTarget).Cells(3).Value = bookReturn.RentPrice
        dgvReturn.Rows(iTarget).Cells(4).Value = ""
        dgvReturn.Rows(iTarget).Cells(5).Value = ""
        dgvReturn.Rows(iTarget).Cells("RentDate").Value = bookReturn.RentDate
        dgvReturn.Rows(iTarget).Cells("RentPrice").Value = bookReturn.RentPrice
    End Sub

    Private Sub showRentBook(ByRef member As Member)
        Dim com As OleDbCommand
        Dim dr As OleDbDataReader
        Dim dt As DataTable
        Dim dataRow As DataRow
        Dim sqlSelect As String

        sqlSelect = ""
        sqlSelect = sqlSelect + "SELECT BookID,BookName,FORMAT(RentDate,'dd/mm/yyyy') AS RentDate ,Rent AS RentPrice"
        sqlSelect = sqlSelect + " FROM RentNote"
        sqlSelect = sqlSelect + " WHERE MemberID = '" + member.MemberID + "' AND ReturnStatus = 0"
        OpenConnection()

        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            dr = com.ExecuteReader()
            If dr.HasRows = True Then
                dt = New DataTable()
                dt.Load(dr)

                Dim bookReturn As BookBeforeReturn = New BookBeforeReturn()
                bookReturn.Init()

                For I As Integer = 0 To dt.Rows.Count - 1
                    dataRow = dt.Rows(I)
                    bookReturn.BookID = dataRow("BookID").ToString()
                    bookReturn.BookName = dataRow("BookName").ToString()
                    bookReturn.RentDate = dataRow("RentDate").ToString()
                    bookReturn.RentPrice = Double.Parse(dataRow("RentPrice").ToString())
                    addTodgvReturn(bookReturn)
                Next
            End If

        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        CloseConnection()
    End Sub

    Private Function getMergeReturnMemberID() As String
        Return tbReturnMemberID0.Text.Trim() + tbReturnMemberID1.Text.Trim() + tbReturnMemberID2.Text.Trim() + tbReturnMemberID3.Text.Trim() + tbReturnMemberID4.Text.Trim()
    End Function

    Private Sub btnReturnSearchMember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturnSearchMember.Click
        clearForNewReturn()

        Dim member As Member = New Member()
        member.Init()
        member.MemberID = getMergeReturnMemberID()
        showReturnMemberDetail(member)

        If member.MemberName <> "" Then
            showRentBook(member)
            If member.PastRent = 0 Then
                MessageBox.Show("สมาชิกท่านนี้ไม่มีหนังสือที่เช่าไว้ค่ะ", cstWarning)
                tbReturnMemberID0.Text = ""
                tbReturnMemberID1.Text = ""
                tbReturnMemberID2.Text = ""
                tbReturnMemberID3.Text = ""
                tbReturnMemberID4.Text = ""
                tbReturnMemberName.Text = ""
                tbReturnMemberSurname.Text = ""
                tbPastRent2.Text = ""
                tbReturnMemberID0.Focus()
            End If
        End If
    End Sub

    Private Sub tbCancelAllReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelAllReturn.Click
        clearForNewReturn()
        tbReturnMemberID0.Text = ""
        tbReturnMemberID1.Text = ""
        tbReturnMemberID2.Text = ""
        tbReturnMemberID3.Text = ""
        tbReturnMemberID4.Text = ""
        tbReturnMemberID0.Focus()
    End Sub

    Private Sub updateRentNoteForReturn(ByVal bookReturn As BookAfterReturn, ByVal returnDetail As ReturnDetail)
        Dim sb As New StringBuilder()
        sb.Append("UPDATE RentNote")
        sb.Append(" SET ReturnStatus=@ReturnStatus,")
        sb.Append("ReturnDate=@ReturnDate")
        sb.Append(" WHERE MemberID=@MemberID AND BookID=@BookID")

        Dim sqlUpdate As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn

        com.Parameters.Add("@ReturnStatus", OleDbType.Integer).Value = "1"

        com.Parameters.Add("@ReturnDate", OleDbType.Date).Value = returnDetail.returnDate
        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = bookReturn.MemberID
        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = bookReturn.BookID

        Try
            com.ExecuteNonQuery()

        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
        CloseConnection()
    End Sub

    Private Sub updateBookForReturn(ByVal bookReturn As BookAfterReturn)
        Dim sb As New StringBuilder()
        sb.Append("UPDATE Books")
        sb.Append(" SET Rent=@Rent")
        sb.Append(" WHERE BookID=@BookID")

        Dim sqlUpdate As String = sb.ToString()
        OpenConnection()
        Dim com As New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@Rent", OleDbType.Integer).Value = "0"
        com.Parameters.Add("@BookID", OleDbType.VarChar).Value = bookReturn.BookID

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Sub updateMemberForReturn(ByVal returnDetail As ReturnDetail)
        Dim sb As New StringBuilder()
        sb.Append("UPDATE Members")
        sb.Append(" SET AmountRent=@AmountRent")
        sb.Append(" WHERE MemberID=@MemberID")

        Dim sqlUpdate As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@AmountRent", OleDbType.Integer).Value = returnDetail.presentAmount
        com.Parameters.Add("@MemberID", OleDbType.VarChar).Value = returnDetail.MemberID

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
    End Sub

    Private Function getMaxRecordID() As Long
        Dim lMaxRecord As Long = 0

        Dim sb As New StringBuilder()
        sb.Append("SELECT Max(RecordID) AS MaxRecordID")
        sb.Append(" FROM DayDetail")

        Dim sqlSelect As String = sb.ToString()
        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            Dim dt As New DataTable()
            dt.Load(dr)

            Dim strMaxRecordID As String = dt.Rows(0)("MaxRecordID").ToString()

            If strMaxRecordID <> "" Then
                lMaxRecord = Long.Parse(strMaxRecordID)
            Else
                lMaxRecord = 0
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        Return lMaxRecord
    End Function

    Private Function getPastRecordID(ByVal strDateDetail As String) As Long
        Dim lRecordID As Long = 0

        Dim sb As New StringBuilder()
        sb.Append("SELECT RecordID")
        sb.Append(" FROM DayDetail")
        sb.Append(" WHERE DateDetail=@DateDetail")

        Dim sqlSelect As String = sb.ToString()
        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@DateDetail", OleDbType.Date).Value = strDateDetail

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows Then
                Dim dt As New DataTable()
                dt.Load(dr)
                lRecordID = Long.Parse(dt.Rows(0)("RecordID").ToString())
            End If

            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try
        Return lRecordID
    End Function

    Private Sub insertReturnStatToDayDetail(ByVal returnStat As SaveStat)
        Dim sb As New StringBuilder()
        sb.Append("INSERT INTO DayDetail(RecordID,DateDetail,AmountMember,")
        sb.Append("RentIncome,FineIncome)")
        sb.Append(" VALUES(@RecordID,@DateDetail,@AmountMember,")
        sb.Append("@RentIncome,@FineIncome)")

        Dim sqlInsert As String = sb.ToString()
        OpenConnection()
        Dim com As New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlInsert
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@RecordID", OleDbType.BigInt).Value = returnStat.lRecordID
        com.Parameters.Add("@DateDetail", OleDbType.Date).Value = returnStat.strDateDetail
        com.Parameters.Add("@AmountMember", OleDbType.Integer).Value = 1
        com.Parameters.Add("@RentIncome", OleDbType.Double).Value = 0
        com.Parameters.Add("@FineIncome", OleDbType.Double).Value = returnStat.dbAmountFine

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

    End Sub

    Private Function getPastAmountMember(ByVal lRecordID As Long) As Integer
        Dim iAmountMember As Integer = 0

        Dim sb As New StringBuilder()
        sb.Append("SELECT AmountMember")
        sb.Append(" FROM DayDetail")
        sb.Append(" WHERE RecordID=@RecordID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@RecordID", OleDbType.BigInt).Value = lRecordID

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                iAmountMember = Integer.Parse(dt.Rows(0)("AmountMember").ToString())
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        Return iAmountMember
    End Function

    Private Function getPastFineIncome(ByVal lRecordID As Long) As Double
        Dim dbPastFine As Double = 0D

        Dim sb As New StringBuilder()
        sb.Append("SELECT FineIncome")
        sb.Append(" FROM DayDetail")
        sb.Append(" WHERE RecordID=@RecordID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@RecordID", OleDbType.BigInt).Value = lRecordID

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                dbPastFine = Double.Parse(dt.Rows(0)("FineIncome").ToString())
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        Return dbPastFine
    End Function

    Private Function getPastStatReturn(ByVal lRecordID As Long) As SaveStat
        Dim pastStat As New SaveStat()
        pastStat.Init()

        Dim sb As New StringBuilder()
        sb.Append("SELECT AmountMember,FineIncome")
        sb.Append(" FROM DayDetail")
        sb.Append(" WHERE RecordID=@RecordID")

        Dim sqlSelect As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        com.Parameters.Add("@RecordID", OleDbType.BigInt).Value = lRecordID

        Try
            Dim dr As OleDbDataReader = com.ExecuteReader()

            If dr.HasRows = True Then
                Dim dt As New DataTable()
                dt.Load(dr)

                pastStat.iAmountMember = Integer.Parse(dt.Rows(0)("AmountMember").ToString())
                pastStat.dbAmountFine = Double.Parse(dt.Rows(0)("FineIncome").ToString())
            End If

            dr.Close()
            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

        Return pastStat
    End Function

    Private Sub updateReturnStatInDayDetail(ByVal returnStat As SaveStat)
        Dim pastStat = getPastStatReturn(returnStat.lRecordID)
        returnStat.iAmountMember = pastStat.iAmountMember + 1
        returnStat.dbAmountFine = pastStat.dbAmountFine + returnStat.dbAmountFine

        Dim sb As New StringBuilder()
        sb.Append("UPDATE DayDetail")
        sb.Append(" SET AmountMember=@AmountMember,")
        sb.Append("FineIncome=@FineIncome")
        sb.Append(" WHERE RecordID=@RecordID")

        Dim sqlUpdate As String = sb.ToString()

        OpenConnection()
        Dim com As New OleDbCommand()
        Dim tr As OleDbTransaction = Conn.BeginTransaction()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn
        com.Transaction = tr

        com.Parameters.Add("@AmountMember", OleDbType.Integer).Value = returnStat.iAmountMember
        com.Parameters.Add("@FineIncome", OleDbType.Double).Value = returnStat.dbAmountFine
        com.Parameters.Add("@RecordID", OleDbType.BigInt).Value = returnStat.lRecordID

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

    End Sub

    Private Sub saveStatReturn(ByVal returnDetail As ReturnDetail)
        Dim returnStat As New SaveStat()
        returnStat.Init()

        returnStat.strDateDetail = returnDetail.returnDate
        returnStat.dbAmountFine = returnDetail.totalFine

        Dim lMaxRecordID As Long = getMaxRecordID()
        If lMaxRecordID <> 0 Then
            Dim lRecordID As Long = getPastRecordID(returnStat.strDateDetail)

            If lRecordID <> 0 Then
                returnStat.lRecordID = lRecordID
                updateReturnStatInDayDetail(returnStat)
            Else
                lMaxRecordID = lMaxRecordID + 1
                returnStat.lRecordID = lMaxRecordID
                insertReturnStatToDayDetail(returnStat)
            End If
        Else
            returnStat.lRecordID = 1
            insertReturnStatToDayDetail(returnStat)
        End If

    End Sub

    Private Sub btnSaveReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveReturn.Click
        Dim returnDetail As New ReturnDetail()
        returnDetail.Init()
        returnDetail.setToPresentDate()

        returnDetail.MemberID = getMergeReturnMemberID()
        returnDetail.presentAmount = Integer.Parse(tbPastRent2.Text)
        returnDetail.totalFine = Double.Parse(tbTotalFine.Text)

        Dim bookReturn As New BookAfterReturn()
        bookReturn.Init()
        bookReturn.MemberID = getMergeReturnMemberID()

        Dim strBookStatus As String = ""
        For I As Integer = 0 To dgvReturn.Rows.Count - 1
            bookReturn.BookID = dgvReturn.Rows(I).Cells(1).Value.ToString()
            strBookStatus = dgvReturn.Rows(I).Cells(4).Value.ToString()
            If strBookStatus <> "" Then

                updateRentNoteForReturn(bookReturn, returnDetail)
                updateBookForReturn(bookReturn)
                returnDetail.presentAmount = returnDetail.presentAmount - 1
            End If
        Next

        updateMemberForReturn(returnDetail)
        saveStatReturn(returnDetail)
        MessageBox.Show("บันทึกการคืนหนังสือเรียบร้อยแล้วค่ะ", cstTitle)
        clearForNewReturn()
        tbReturnMemberID0.Text = ""
        tbReturnMemberID1.Text = ""
        tbReturnMemberID2.Text = ""
        tbReturnMemberID3.Text = ""
        tbReturnMemberID4.Text = ""
        tbReturnMemberID0.Focus()
    End Sub

    Private Sub dgvReturn_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReturn.CellContentClick
        'MessageBox.Show(e.RowIndex.ToString())
        Dim I As Integer = e.RowIndex
        If e.ColumnIndex = 0 Then
            If dgvReturn.Rows(I).Cells(4).Value.ToString() = "" Then
                Dim iReturnDay As Integer = getDayFromTextDate(dgvReturn.Rows(I).Cells("RentDate").Value.ToString())
                Dim iReturnMonth As Integer = getMonthFromTextDate(dgvReturn.Rows(I).Cells("RentDate").Value.ToString())
                Dim iReturnYear As Integer = getYearFromTextDate(dgvReturn.Rows(I).Cells("RentDate").Value.ToString())

                Dim iCurrentDay As Integer = DateTime.Now.Day
                Dim iCurrentMonth As Integer = DateTime.Now.Month
                Dim iCurrentYear As Integer = DateTime.Now.Year

                If (CultureInfo.CurrentCulture.ToString() = "th-TH") Then
                    iCurrentYear = iCurrentYear + 543
                End If

                Dim datePresent As DateTime = New DateTime(iCurrentYear, iCurrentMonth, iCurrentDay)
                Dim dateRent As DateTime = New DateTime(iReturnYear, iReturnMonth, iReturnDay)
                Dim iLateReturn As Integer = calDiffDay(datePresent, dateRent)
                'MessageBox.Show(iLateReturn)

                Dim dbFine As Double = 0D
                If iLateReturn = 1 Or iLateReturn = 0 Then
                    dgvReturn.Rows(I).Cells(4).Value = "0"
                    dbFine = 0D
                Else
                    iLateReturn = iLateReturn - 1
                    dgvReturn.Rows(I).Cells(4).Value = iLateReturn.ToString()
                    dbFine = Double.Parse((dgvReturn.Rows(I).Cells(3).Value.ToString()))
                End If
                dbFine = dbFine * iLateReturn
                dgvReturn.Rows(I).Cells(5).Value = dbFine

                Dim iAmountReturn As Integer = Integer.Parse(tbAmountReturn.Text.Trim())
                iAmountReturn = iAmountReturn + 1
                tbAmountReturn.Text = iAmountReturn.ToString()

                Dim dbTotalFine As Double = Double.Parse(tbTotalFine.Text.Trim())
                dbTotalFine = dbTotalFine + dbFine
                tbTotalFine.Text = dbTotalFine
                dgvReturn.Rows(I).Cells(0).Value = True

            ElseIf dgvReturn.Rows(I).Cells(4).Value.ToString() <> "" Then
                Dim iTotalReturn As Integer = Integer.Parse(tbAmountReturn.Text.Trim())
                iTotalReturn = iTotalReturn - 1
                tbAmountReturn.Text = iTotalReturn.ToString()

                Dim dbTotalFine As Double = Double.Parse(tbTotalFine.Text.Trim())
                Dim dbTemp As Double = 0D
                dbTemp = Double.Parse(dgvReturn.Rows(I).Cells(5).Value)
                dbTotalFine = dbTotalFine - dbTemp
                tbTotalFine.Text = dbTotalFine.ToString()

                dgvReturn.Rows(I).Cells(4).Value = ""
                dgvReturn.Rows(I).Cells(5).Value = ""
                dgvReturn.Rows(I).Cells(0).Value = False
            End If
        End If
    End Sub
#End Region

#Region "Report"
    Private Sub reportOver_Click(sender As System.Object, e As System.EventArgs) Handles reportOver.Click
        Dim frmOver As New FormReportOver()
        frmOver.ShowDialog()
    End Sub

    Private Sub reportCirMonth_Click(sender As System.Object, e As System.EventArgs) Handles reportCirMonth.Click
        Dim frmCirMonth As New FormCirMonth()
        frmCirMonth.ShowDialog()
    End Sub

    Private Sub reportCirYear_Click(sender As System.Object, e As System.EventArgs) Handles reportCirYear.Click
        Dim frmCirYear As New FormCirYear()
        frmCirYear.ShowDialog()
    End Sub

    Private Sub reportAmountMonth_Click(sender As System.Object, e As System.EventArgs) Handles reportAmountMonth.Click
        Dim frmAmountMonth As New FormAmountMonth()
        frmAmountMonth.ShowDialog()
    End Sub

    Private Sub reportAmountYear_Click(sender As System.Object, e As System.EventArgs) Handles reportAmountYear.Click
        Dim frmAmountYear As New FormAmountYear()
        frmAmountYear.ShowDialog()
    End Sub
#End Region

    
End Class


