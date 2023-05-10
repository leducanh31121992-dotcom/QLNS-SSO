<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKeHoachLaoDong
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.labAlert = New System.Windows.Forms.Label()
        Me.bntNew = New System.Windows.Forms.Button()
        Me.bntUpdate = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.bntDelete = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.gridKHLD = New System.Windows.Forms.DataGridView()
        Me.idKHLD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nam = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Thang = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Id_DonVi = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdPhong = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DaiHan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NganHan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtGhiChu = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.txtSoNganHan = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtSoDaiHan = New System.Windows.Forms.TextBox()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.labNam = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.numThang = New System.Windows.Forms.NumericUpDown()
        Me.dtpkNam = New System.Windows.Forms.DateTimePicker()
        Me.rdThang = New System.Windows.Forms.RadioButton()
        Me.rdNam = New System.Windows.Forms.RadioButton()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboPhong = New System.Windows.Forms.ComboBox()
        Me.labPhong = New System.Windows.Forms.Label()
        Me.Panel10.SuspendLayout()
        CType(Me.gridKHLD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.numThang, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.Color.DarkGray
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel10.Controls.Add(Me.labAlert)
        Me.Panel10.Controls.Add(Me.bntNew)
        Me.Panel10.Controls.Add(Me.bntUpdate)
        Me.Panel10.Controls.Add(Me.btnCancel)
        Me.Panel10.Controls.Add(Me.bntDelete)
        Me.Panel10.Controls.Add(Me.bntClose)
        Me.Panel10.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel10.Location = New System.Drawing.Point(0, 537)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(787, 26)
        Me.Panel10.TabIndex = 6
        '
        'labAlert
        '
        Me.labAlert.Dock = System.Windows.Forms.DockStyle.Left
        Me.labAlert.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labAlert.ForeColor = System.Drawing.Color.Maroon
        Me.labAlert.Location = New System.Drawing.Point(0, 0)
        Me.labAlert.Name = "labAlert"
        Me.labAlert.Size = New System.Drawing.Size(399, 22)
        Me.labAlert.TabIndex = 8
        Me.labAlert.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'bntNew
        '
        Me.bntNew.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntNew.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntNew.Location = New System.Drawing.Point(466, 0)
        Me.bntNew.Name = "bntNew"
        Me.bntNew.Size = New System.Drawing.Size(73, 22)
        Me.bntNew.TabIndex = 4
        Me.bntNew.Text = "&Thêm mới"
        Me.bntNew.UseVisualStyleBackColor = True
        '
        'bntUpdate
        '
        Me.bntUpdate.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntUpdate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntUpdate.Location = New System.Drawing.Point(539, 0)
        Me.bntUpdate.Name = "bntUpdate"
        Me.bntUpdate.Size = New System.Drawing.Size(61, 22)
        Me.bntUpdate.TabIndex = 1
        Me.bntUpdate.Text = "&Ghi"
        Me.bntUpdate.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(600, 0)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(61, 22)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Bỏ qua"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'bntDelete
        '
        Me.bntDelete.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntDelete.Enabled = False
        Me.bntDelete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntDelete.Location = New System.Drawing.Point(661, 0)
        Me.bntDelete.Name = "bntDelete"
        Me.bntDelete.Size = New System.Drawing.Size(61, 22)
        Me.bntDelete.TabIndex = 3
        Me.bntDelete.Text = "&Xoá"
        Me.bntDelete.UseVisualStyleBackColor = True
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntClose.Location = New System.Drawing.Point(722, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(61, 22)
        Me.bntClose.TabIndex = 5
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'gridKHLD
        '
        Me.gridKHLD.AllowUserToAddRows = False
        Me.gridKHLD.AllowUserToDeleteRows = False
        Me.gridKHLD.AllowUserToResizeColumns = False
        Me.gridKHLD.AllowUserToResizeRows = False
        Me.gridKHLD.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridKHLD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridKHLD.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridKHLD.ColumnHeadersHeight = 37
        Me.gridKHLD.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.idKHLD, Me.Nam, Me.Thang, Me.Id_DonVi, Me.IdPhong, Me.DaiHan, Me.NganHan})
        Me.gridKHLD.Dock = System.Windows.Forms.DockStyle.Top
        Me.gridKHLD.Location = New System.Drawing.Point(0, 0)
        Me.gridKHLD.MultiSelect = False
        Me.gridKHLD.Name = "gridKHLD"
        Me.gridKHLD.ReadOnly = True
        Me.gridKHLD.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridKHLD.RowHeadersVisible = False
        Me.gridKHLD.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridKHLD.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridKHLD.Size = New System.Drawing.Size(787, 367)
        Me.gridKHLD.TabIndex = 123
        Me.gridKHLD.Tag = ""
        '
        'idKHLD
        '
        Me.idKHLD.DataPropertyName = "idKHLD"
        Me.idKHLD.HeaderText = "IdLDV"
        Me.idKHLD.Name = "idKHLD"
        Me.idKHLD.ReadOnly = True
        Me.idKHLD.Visible = False
        '
        'Nam
        '
        Me.Nam.DataPropertyName = "Nam"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.Nam.DefaultCellStyle = DataGridViewCellStyle1
        Me.Nam.HeaderText = "Năm"
        Me.Nam.Name = "Nam"
        Me.Nam.ReadOnly = True
        Me.Nam.Width = 60
        '
        'Thang
        '
        Me.Thang.DataPropertyName = "Thang"
        Me.Thang.HeaderText = "Tháng"
        Me.Thang.Name = "Thang"
        Me.Thang.ReadOnly = True
        Me.Thang.Width = 50
        '
        'Id_DonVi
        '
        Me.Id_DonVi.DataPropertyName = "IdDonVi"
        Me.Id_DonVi.HeaderText = "Đơn vị"
        Me.Id_DonVi.Name = "Id_DonVi"
        Me.Id_DonVi.ReadOnly = True
        Me.Id_DonVi.Width = 294
        '
        'IdPhong
        '
        Me.IdPhong.DataPropertyName = "IdPhong"
        Me.IdPhong.HeaderText = "Phòng"
        Me.IdPhong.Name = "IdPhong"
        Me.IdPhong.ReadOnly = True
        Me.IdPhong.Width = 220
        '
        'DaiHan
        '
        Me.DaiHan.DataPropertyName = "DaiHan"
        Me.DaiHan.HeaderText = "Số lao động dài hạn"
        Me.DaiHan.Name = "DaiHan"
        Me.DaiHan.ReadOnly = True
        Me.DaiHan.Width = 80
        '
        'NganHan
        '
        Me.NganHan.DataPropertyName = "NganHan"
        Me.NganHan.HeaderText = "Số lao động ngắn hạn"
        Me.NganHan.Name = "NganHan"
        Me.NganHan.ReadOnly = True
        Me.NganHan.Width = 80
        '
        'txtGhiChu
        '
        Me.txtGhiChu.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhiChu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhiChu.Location = New System.Drawing.Point(114, 492)
        Me.txtGhiChu.Multiline = True
        Me.txtGhiChu.Name = "txtGhiChu"
        Me.txtGhiChu.Size = New System.Drawing.Size(671, 44)
        Me.txtGhiChu.TabIndex = 9
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(65, 495)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(56, 18)
        Me.Label39.TabIndex = 122
        Me.Label39.Text = "Ghi chú"
        '
        'txtSoNganHan
        '
        Me.txtSoNganHan.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoNganHan.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoNganHan.Location = New System.Drawing.Point(460, 469)
        Me.txtSoNganHan.Name = "txtSoNganHan"
        Me.txtSoNganHan.Size = New System.Drawing.Size(325, 26)
        Me.txtSoNganHan.TabIndex = 8
        Me.txtSoNganHan.Text = "0"
        Me.txtSoNganHan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(-2, 472)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(137, 18)
        Me.Label42.TabIndex = 121
        Me.Label42.Text = "Số lao động dài hạn"
        '
        'txtSoDaiHan
        '
        Me.txtSoDaiHan.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoDaiHan.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoDaiHan.Location = New System.Drawing.Point(114, 469)
        Me.txtSoDaiHan.Name = "txtSoDaiHan"
        Me.txtSoDaiHan.Size = New System.Drawing.Size(210, 26)
        Me.txtSoDaiHan.TabIndex = 7
        Me.txtSoDaiHan.Text = "0"
        Me.txtSoDaiHan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(333, 473)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(151, 18)
        Me.Label44.TabIndex = 120
        Me.Label44.Text = "Số lao động ngắn hạn"
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(114, 446)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(341, 26)
        Me.cboDonVi.TabIndex = 5
        Me.cboDonVi.ValueMember = "Value"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(72, 450)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(49, 18)
        Me.Label35.TabIndex = 119
        Me.Label35.Text = "Đơn vị"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(243, 18)
        Me.Label1.TabIndex = 124
        Me.Label1.Text = "Số kế hoạch lao động trong năm"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(243, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(219, 18)
        Me.Label2.TabIndex = 125
        Me.Label2.Text = "(được thông báo đến 31/12)"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.labNam)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(112, 389)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(592, 22)
        Me.Panel1.TabIndex = 126
        '
        'labNam
        '
        Me.labNam.AutoSize = True
        Me.labNam.Dock = System.Windows.Forms.DockStyle.Left
        Me.labNam.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labNam.Location = New System.Drawing.Point(243, 0)
        Me.labNam.Name = "labNam"
        Me.labNam.Size = New System.Drawing.Size(0, 18)
        Me.labNam.TabIndex = 126
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(82, 427)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 18)
        Me.Label3.TabIndex = 128
        Me.Label3.Text = "Năm"
        '
        'numThang
        '
        Me.numThang.Location = New System.Drawing.Point(460, 423)
        Me.numThang.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.numThang.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numThang.Name = "numThang"
        Me.numThang.Size = New System.Drawing.Size(39, 26)
        Me.numThang.TabIndex = 4
        Me.numThang.Value = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numThang.Visible = False
        '
        'dtpkNam
        '
        Me.dtpkNam.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpkNam.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpkNam.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpkNam.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpkNam.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpkNam.CustomFormat = "yyyy"
        Me.dtpkNam.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpkNam.Location = New System.Drawing.Point(114, 423)
        Me.dtpkNam.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam.Name = "dtpkNam"
        Me.dtpkNam.ShowUpDown = True
        Me.dtpkNam.Size = New System.Drawing.Size(106, 26)
        Me.dtpkNam.TabIndex = 129
        '
        'rdThang
        '
        Me.rdThang.AutoSize = True
        Me.rdThang.Location = New System.Drawing.Point(387, 426)
        Me.rdThang.Name = "rdThang"
        Me.rdThang.Size = New System.Drawing.Size(85, 22)
        Me.rdThang.TabIndex = 3
        Me.rdThang.Text = "từ tháng"
        Me.rdThang.UseVisualStyleBackColor = True
        '
        'rdNam
        '
        Me.rdNam.AutoSize = True
        Me.rdNam.Checked = True
        Me.rdNam.Location = New System.Drawing.Point(323, 426)
        Me.rdNam.Name = "rdNam"
        Me.rdNam.Size = New System.Drawing.Size(78, 22)
        Me.rdNam.TabIndex = 2
        Me.rdNam.TabStop = True
        Me.rdNam.Text = "cả năm"
        Me.rdNam.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox2.Location = New System.Drawing.Point(329, 422)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(66, 26)
        Me.TextBox2.TabIndex = 134
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(224, 427)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(116, 18)
        Me.Label5.TabIndex = 135
        Me.Label5.Text = "Số kế hoạch cho"
        '
        'cboPhong
        '
        Me.cboPhong.BackColor = System.Drawing.SystemColors.Window
        Me.cboPhong.DisplayMember = "Display"
        Me.cboPhong.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPhong.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhong.FormattingEnabled = True
        Me.cboPhong.Location = New System.Drawing.Point(499, 446)
        Me.cboPhong.Name = "cboPhong"
        Me.cboPhong.Size = New System.Drawing.Size(286, 26)
        Me.cboPhong.TabIndex = 6
        Me.cboPhong.ValueMember = "Value"
        '
        'labPhong
        '
        Me.labPhong.AutoSize = True
        Me.labPhong.Location = New System.Drawing.Point(457, 450)
        Me.labPhong.Name = "labPhong"
        Me.labPhong.Size = New System.Drawing.Size(48, 18)
        Me.labPhong.TabIndex = 137
        Me.labPhong.Text = "Phòng"
        '
        'frmKeHoachLaoDong
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.cboPhong)
        Me.Controls.Add(Me.labPhong)
        Me.Controls.Add(Me.numThang)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.rdThang)
        Me.Controls.Add(Me.rdNam)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.dtpkNam)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.gridKHLD)
        Me.Controls.Add(Me.Panel10)
        Me.Controls.Add(Me.txtGhiChu)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.txtSoNganHan)
        Me.Controls.Add(Me.Label42)
        Me.Controls.Add(Me.txtSoDaiHan)
        Me.Controls.Add(Me.Label44)
        Me.Controls.Add(Me.cboDonVi)
        Me.Controls.Add(Me.Label35)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmKeHoachLaoDong"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: KẾ HOẠCH LAO ĐỘNG HÀNG NĂM CỦA ĐƠN VỊ"
        Me.Panel10.ResumeLayout(False)
        CType(Me.gridKHLD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.numThang, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents bntNew As System.Windows.Forms.Button
    Friend WithEvents bntUpdate As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents bntDelete As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents gridKHLD As System.Windows.Forms.DataGridView
    Friend WithEvents txtGhiChu As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents txtSoNganHan As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtSoDaiHan As System.Windows.Forms.TextBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents labNam As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents numThang As System.Windows.Forms.NumericUpDown
    Friend WithEvents dtpkNam As System.Windows.Forms.DateTimePicker
    Friend WithEvents labAlert As System.Windows.Forms.Label
    Friend WithEvents rdThang As System.Windows.Forms.RadioButton
    Friend WithEvents rdNam As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboPhong As System.Windows.Forms.ComboBox
    Friend WithEvents labPhong As System.Windows.Forms.Label
    Friend WithEvents idKHLD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nam As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Thang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Id_DonVi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdPhong As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DaiHan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NganHan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
