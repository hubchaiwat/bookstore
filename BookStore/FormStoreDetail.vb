Imports System.Data
Imports System.Data.OleDb

Public Class FormStoreDetail

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
#End Region

    Private Sub FormStoreDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sqlSelect As String = ""
        Dim com As OleDbCommand
        Dim dr As OleDbDataReader
        Dim dt As DataTable
        Dim dataRow As DataRow

        sqlSelect = sqlSelect + "SELECT * FROM StoreDetail"
        OpenConnection()
        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn

        Try
            dr = com.ExecuteReader()
            dt = New DataTable
            dt.Load(dr)

            dataRow = dt.Rows(0)
            tbName.Text = dataRow("StoreName").ToString()
            tbAddress.Text = dataRow("Address").ToString()
            tbProvince.Text = dataRow("Province").ToString()
            tbPostalCode.Text = dataRow("PostalCode").ToString()
            tbTel.Text = dataRow("Tel").ToString()
            tbFax.Text = dataRow("Fax").ToString()

            CloseConnection()
        Catch ex As Exception
            CloseConnection()
            MessageBox.Show(ex.Message, cstWarning)
        End Try

    End Sub

    Private Function checkAllField()
        Dim bCheckField As Boolean = True
        Dim strOut As String = ""

        If (tbName.Text.Trim() = "") Then
            strOut = strOut + "กรุณาใส่ชื่อร้าน"
        End If

        If (tbAddress.Text.Trim() = "") Then
            strOut = strOut + " ที่อยู่ของร้าน"
        End If

        If (tbProvince.Text.Trim() = "") Then
            strOut = strOut + " จังหวัดที่ร้านตั้งอยู่"
        End If

        If (tbPostalCode.Text.Trim() = "") Then
            strOut = strOut + " รหัสไปรษณีย์"
        End If

        If (tbTel.Text.Trim() = "") Then
            strOut = strOut + " หมายเลขโทรศัพท์"
        End If

        If (tbFax.Text.Trim() = "") Then
            strOut = strOut + " หมายเลขแฟกซ์"
        End If

        If (strOut <> "") Then
            bCheckField = False
            MessageBox.Show(strOut + "ด้วยค่ะ", "คำเตือน")
        End If

        Return bCheckField
    End Function

    Private Sub setReadOnlyField(ByVal bValue As Boolean)
        tbName.ReadOnly = bValue
        tbAddress.ReadOnly = bValue
        tbProvince.ReadOnly = bValue
        tbPostalCode.ReadOnly = bValue
        tbTel.ReadOnly = bValue
        tbFax.ReadOnly = bValue
    End Sub

    Private Sub btnEdit_Click(sender As System.Object, e As System.EventArgs) Handles btnEdit.Click
        setReadOnlyField(False)
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnClose.Visible = False
        btnCancel.Visible = True
    End Sub

    Private Function updateStorDetail() As Boolean
        Dim bComplete As Boolean = False

        'สร้างคำสั่ง SQL สำหรับ SELECT
        Dim sqlSelect As String = ""
        sqlSelect = sqlSelect + "SELECT * FROM StoreDetail"

        Dim com As OleDbCommand
        Dim dr As OleDbDataReader
        Dim dt As DataTable
        Dim tr As OleDbTransaction

        OpenConnection()
        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlSelect
        com.Connection = Conn
        dr = com.ExecuteReader()
        dt = New DataTable()
        dt.Load(dr)

        Dim dataRow = dt.Rows(0)
        Dim strName As String = dataRow("StoreName").ToString()

        'สร้างคำสั่ง SQL
        Dim sqlUpdate As String = ""
        sqlUpdate = sqlUpdate + "UPDATE StoreDetail"
        sqlUpdate = sqlUpdate + " SET StoreName = '" + tbName.Text.Trim() + "',"
        sqlUpdate = sqlUpdate + "Address = '" + tbAddress.Text.Trim() + "',"
        sqlUpdate = sqlUpdate + "Province = '" + tbProvince.Text.Trim() + "',"
        sqlUpdate = sqlUpdate + "PostalCode = '" + tbPostalCode.Text.Trim() + "',"
        sqlUpdate = sqlUpdate + "Tel = '" + tbTel.Text.Trim() + "',"
        sqlUpdate = sqlUpdate + "Fax = '" + tbFax.Text.Trim() + "'"

        sqlUpdate = sqlUpdate + " WHERE StoreName = '" + strName + "'"
        OpenConnection()
        tr = Conn.BeginTransaction()
        com = New OleDbCommand()
        com.CommandType = CommandType.Text
        com.CommandText = sqlUpdate
        com.Connection = Conn
        com.Transaction = tr

        Try
            com.ExecuteNonQuery()
            tr.Commit()
            CloseConnection()
            bComplete = True
        Catch ex As Exception
            tr.Rollback()
            CloseConnection()
            bComplete = False
            MessageBox.Show("เกิดข้อผิดพลาดเนื่องจาก" + ex.Message, "ผลการทำงาน")
        End Try

        Return bComplete
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim bCheck = checkAllField()

        If (bCheck = False) Then
            Return
        End If

        Dim bUpdateComplete As Boolean = updateStorDetail()

        If bUpdateComplete = True Then
            setReadOnlyField(True)
            btnEdit.Enabled = True
            btnSave.Enabled = False
            btnCancel.Visible = False
            btnClose.Visible = True
            MessageBox.Show("อัพเดทข้อมูลของร้านเสร็จสมบูรณ์", "ผลการทำงาน")
            tbName.Focus()
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        btnEdit.Enabled = True
        btnSave.Enabled = False
        btnCancel.Visible = False
        btnClose.Visible = True
        setReadOnlyField(True)
    End Sub
End Class