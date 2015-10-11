<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormStoreDetail
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormStoreDetail))
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.tbFax = New System.Windows.Forms.TextBox()
        Me.label6 = New System.Windows.Forms.Label()
        Me.tbTel = New System.Windows.Forms.TextBox()
        Me.label5 = New System.Windows.Forms.Label()
        Me.tbPostalCode = New System.Windows.Forms.TextBox()
        Me.label4 = New System.Windows.Forms.Label()
        Me.tbProvince = New System.Windows.Forms.TextBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.tbAddress = New System.Windows.Forms.TextBox()
        Me.tbName = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(306, 176)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 41
        Me.btnClose.Text = "ปิด"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(225, 176)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 40
        Me.btnSave.Text = "บันทึก"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'tbFax
        '
        Me.tbFax.BackColor = System.Drawing.Color.White
        Me.tbFax.Location = New System.Drawing.Point(272, 150)
        Me.tbFax.MaxLength = 20
        Me.tbFax.Name = "tbFax"
        Me.tbFax.ReadOnly = True
        Me.tbFax.Size = New System.Drawing.Size(109, 20)
        Me.tbFax.TabIndex = 39
        '
        'label6
        '
        Me.label6.AutoSize = True
        Me.label6.Location = New System.Drawing.Point(229, 153)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(37, 13)
        Me.label6.TabIndex = 38
        Me.label6.Text = "แฟกซ์"
        '
        'tbTel
        '
        Me.tbTel.BackColor = System.Drawing.Color.White
        Me.tbTel.Location = New System.Drawing.Point(87, 150)
        Me.tbTel.MaxLength = 20
        Me.tbTel.Name = "tbTel"
        Me.tbTel.ReadOnly = True
        Me.tbTel.Size = New System.Drawing.Size(100, 20)
        Me.tbTel.TabIndex = 37
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Location = New System.Drawing.Point(33, 153)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(48, 13)
        Me.label5.TabIndex = 36
        Me.label5.Text = "โทรศัพท์"
        '
        'tbPostalCode
        '
        Me.tbPostalCode.BackColor = System.Drawing.Color.White
        Me.tbPostalCode.Location = New System.Drawing.Point(272, 125)
        Me.tbPostalCode.MaxLength = 5
        Me.tbPostalCode.Name = "tbPostalCode"
        Me.tbPostalCode.ReadOnly = True
        Me.tbPostalCode.Size = New System.Drawing.Size(109, 20)
        Me.tbPostalCode.TabIndex = 35
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(197, 128)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(69, 13)
        Me.label4.TabIndex = 34
        Me.label4.Text = "รหัสไปรษณีย์"
        '
        'tbProvince
        '
        Me.tbProvince.BackColor = System.Drawing.Color.White
        Me.tbProvince.Location = New System.Drawing.Point(87, 125)
        Me.tbProvince.MaxLength = 30
        Me.tbProvince.Name = "tbProvince"
        Me.tbProvince.ReadOnly = True
        Me.tbProvince.Size = New System.Drawing.Size(104, 20)
        Me.tbProvince.TabIndex = 33
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(46, 128)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(38, 13)
        Me.label3.TabIndex = 32
        Me.label3.Text = "จังหวัด"
        '
        'tbAddress
        '
        Me.tbAddress.BackColor = System.Drawing.Color.White
        Me.tbAddress.Location = New System.Drawing.Point(87, 45)
        Me.tbAddress.MaxLength = 100
        Me.tbAddress.Multiline = True
        Me.tbAddress.Name = "tbAddress"
        Me.tbAddress.ReadOnly = True
        Me.tbAddress.Size = New System.Drawing.Size(294, 75)
        Me.tbAddress.TabIndex = 31
        '
        'tbName
        '
        Me.tbName.BackColor = System.Drawing.Color.White
        Me.tbName.Location = New System.Drawing.Point(87, 19)
        Me.tbName.MaxLength = 40
        Me.tbName.Name = "tbName"
        Me.tbName.ReadOnly = True
        Me.tbName.Size = New System.Drawing.Size(294, 20)
        Me.tbName.TabIndex = 30
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(54, 48)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(27, 13)
        Me.label2.TabIndex = 29
        Me.label2.Text = "ที่อยู่"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(42, 22)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(39, 13)
        Me.label1.TabIndex = 28
        Me.label1.Text = "ชื่อร้าน"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(36, 176)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(75, 23)
        Me.btnEdit.TabIndex = 42
        Me.btnEdit.Text = "แก้ไขข้อมูล"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(306, 176)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 43
        Me.btnCancel.Text = "ยกเลิก"
        Me.btnCancel.UseVisualStyleBackColor = True
        Me.btnCancel.Visible = False
        '
        'FormStoreDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(411, 222)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.tbFax)
        Me.Controls.Add(Me.label6)
        Me.Controls.Add(Me.tbTel)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.tbPostalCode)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.tbProvince)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.tbAddress)
        Me.Controls.Add(Me.tbName)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(427, 260)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(427, 260)
        Me.Name = "FormStoreDetail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "กำหนดข้อมูลร้านหนังสือ"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents tbFax As System.Windows.Forms.TextBox
    Private WithEvents label6 As System.Windows.Forms.Label
    Private WithEvents tbTel As System.Windows.Forms.TextBox
    Private WithEvents label5 As System.Windows.Forms.Label
    Private WithEvents tbPostalCode As System.Windows.Forms.TextBox
    Private WithEvents label4 As System.Windows.Forms.Label
    Private WithEvents tbProvince As System.Windows.Forms.TextBox
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents tbAddress As System.Windows.Forms.TextBox
    Private WithEvents tbName As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
