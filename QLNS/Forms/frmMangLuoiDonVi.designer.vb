<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMangLuoiDonVi
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.labAlert = New System.Windows.Forms.Label()
        Me.bntNew = New System.Windows.Forms.Button()
        Me.bntUpdate = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.bntDelete = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.gridMLDV = New System.Windows.Forms.DataGridView()
        Me.IdML = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nam = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Quy = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdPGD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SoXaPhuong = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SoDiemGD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DuNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtGhiChu = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.txtDuNo = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtSoDiemGD = New System.Windows.Forms.TextBox()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.cboDV = New System.Windows.Forms.ComboBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.numQuy = New System.Windows.Forms.NumericUpDown()
        Me.dtpkNam = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtXaPhuong = New System.Windows.Forms.TextBox()
        Me.Panel10.SuspendLayout()
        CType(Me.gridMLDV, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numQuy, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.labAlert.Size = New System.Drawing.Size(444, 22)
        Me.labAlert.TabIndex = 7
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
        'gridMLDV
        '
        Me.gridMLDV.AllowUserToAddRows = False
        Me.gridMLDV.AllowUserToDeleteRows = False
        Me.gridMLDV.AllowUserToResizeColumns = False
        Me.gridMLDV.AllowUserToResizeRows = False
        Me.gridMLDV.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridMLDV.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridMLDV.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridMLDV.ColumnHeadersHeight = 37
        Me.gridMLDV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdML, Me.Nam, Me.Quy, Me.IdPGD, Me.SoXaPhuong, Me.SoDiemGD, Me.DuNo})
        Me.gridMLDV.Dock = System.Windows.Forms.DockStyle.Top
        Me.gridMLDV.Location = New System.Drawing.Point(0, 0)
        Me.gridMLDV.MultiSelect = False
        Me.gridMLDV.Name = "gridMLDV"
        Me.gridMLDV.ReadOnly = True
        Me.gridMLDV.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridMLDV.RowHeadersVisible = False
        Me.gridMLDV.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridMLDV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridMLDV.Size = New System.Drawing.Size(787, 415)
        Me.gridMLDV.TabIndex = 123
        Me.gridMLDV.Tag = ""
        '
        'IdML
        '
        Me.IdML.DataPropertyName = "IdMangLuoi"
        Me.IdML.HeaderText = "IdML"
        Me.IdML.Name = "IdML"
        Me.IdML.ReadOnly = True
        Me.IdML.Visible = False
        '
        'Nam
        '
        Me.Nam.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Nam.DataPropertyName = "Nam"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.Nam.DefaultCellStyle = DataGridViewCellStyle1
        Me.Nam.FillWeight = 50.0!
        Me.Nam.HeaderText = "Năm"
        Me.Nam.Name = "Nam"
        Me.Nam.ReadOnly = True
        Me.Nam.Width = 56
        '
        'Quy
        '
        Me.Quy.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Quy.DataPropertyName = "Quy"
        Me.Quy.FillWeight = 50.0!
        Me.Quy.HeaderText = "Quý"
        Me.Quy.Name = "Quy"
        Me.Quy.ReadOnly = True
        Me.Quy.Width = 54
        '
        'IdPGD
        '
        Me.IdPGD.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.IdPGD.DataPropertyName = "IdPGD"
        Me.IdPGD.HeaderText = "Đơn vị"
        Me.IdPGD.Name = "IdPGD"
        Me.IdPGD.ReadOnly = True
        '
        'SoXaPhuong
        '
        Me.SoXaPhuong.DataPropertyName = "SoXaPhuong"
        Me.SoXaPhuong.HeaderText = "Số xã, phường"
        Me.SoXaPhuong.Name = "SoXaPhuong"
        Me.SoXaPhuong.ReadOnly = True
        '
        'SoDiemGD
        '
        Me.SoDiemGD.DataPropertyName = "SoDiemGD"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.Format = "N0"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.SoDiemGD.DefaultCellStyle = DataGridViewCellStyle2
        Me.SoDiemGD.HeaderText = "Số điểm giao dịch"
        Me.SoDiemGD.Name = "SoDiemGD"
        Me.SoDiemGD.ReadOnly = True
        '
        'DuNo
        '
        Me.DuNo.DataPropertyName = "DuNo"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Format = "N0"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.DuNo.DefaultCellStyle = DataGridViewCellStyle3
        Me.DuNo.HeaderText = "Dư nợ"
        Me.DuNo.Name = "DuNo"
        Me.DuNo.ReadOnly = True
        Me.DuNo.Width = 180
        '
        'txtGhiChu
        '
        Me.txtGhiChu.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhiChu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhiChu.Location = New System.Drawing.Point(152, 490)
        Me.txtGhiChu.Multiline = True
        Me.txtGhiChu.Name = "txtGhiChu"
        Me.txtGhiChu.Size = New System.Drawing.Size(503, 44)
        Me.txtGhiChu.TabIndex = 7
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(103, 493)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(48, 14)
        Me.Label39.TabIndex = 122
        Me.Label39.Text = "Ghi chú"
        '
        'txtDuNo
        '
        Me.txtDuNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDuNo.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDuNo.Location = New System.Drawing.Point(485, 467)
        Me.txtDuNo.Name = "txtDuNo"
        Me.txtDuNo.Size = New System.Drawing.Size(170, 22)
        Me.txtDuNo.TabIndex = 6
        Me.txtDuNo.Text = "0"
        Me.txtDuNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(250, 470)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(103, 14)
        Me.Label42.TabIndex = 121
        Me.Label42.Text = "Số điểm giao dịch"
        '
        'txtSoDiemGD
        '
        Me.txtSoDiemGD.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoDiemGD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoDiemGD.Location = New System.Drawing.Point(354, 467)
        Me.txtSoDiemGD.Name = "txtSoDiemGD"
        Me.txtSoDiemGD.Size = New System.Drawing.Size(69, 22)
        Me.txtSoDiemGD.TabIndex = 5
        Me.txtSoDiemGD.Text = "0"
        Me.txtSoDiemGD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(441, 471)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(41, 14)
        Me.Label44.TabIndex = 120
        Me.Label44.Text = "Dư nợ"
        '
        'cboDV
        '
        Me.cboDV.BackColor = System.Drawing.SystemColors.Window
        Me.cboDV.DisplayMember = "Display"
        Me.cboDV.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDV.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDV.FormattingEnabled = True
        Me.cboDV.Location = New System.Drawing.Point(152, 444)
        Me.cboDV.Name = "cboDV"
        Me.cboDV.Size = New System.Drawing.Size(503, 22)
        Me.cboDV.TabIndex = 3
        Me.cboDV.ValueMember = "Value"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(110, 448)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(41, 14)
        Me.Label35.TabIndex = 119
        Me.Label35.Text = "Đơn vị"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(120, 425)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 14)
        Me.Label3.TabIndex = 128
        Me.Label3.Text = "Năm"
        '
        'numQuy
        '
        Me.numQuy.Location = New System.Drawing.Point(354, 421)
        Me.numQuy.Maximum = New Decimal(New Integer() {4, 0, 0, 0})
        Me.numQuy.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numQuy.Name = "numQuy"
        Me.numQuy.Size = New System.Drawing.Size(69, 22)
        Me.numQuy.TabIndex = 2
        Me.numQuy.Value = New Decimal(New Integer() {1, 0, 0, 0})
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
        Me.dtpkNam.Location = New System.Drawing.Point(152, 421)
        Me.dtpkNam.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam.Name = "dtpkNam"
        Me.dtpkNam.ShowUpDown = True
        Me.dtpkNam.Size = New System.Drawing.Size(69, 22)
        Me.dtpkNam.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(324, 425)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 14)
        Me.Label4.TabIndex = 131
        Me.Label4.Text = "Quý"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(655, 470)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(129, 14)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "(đơn vị tính: Tỷ đồng)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(30, 470)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 14)
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "Tổng số xã, phường"
        '
        'txtXaPhuong
        '
        Me.txtXaPhuong.BackColor = System.Drawing.SystemColors.Window
        Me.txtXaPhuong.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtXaPhuong.Location = New System.Drawing.Point(152, 467)
        Me.txtXaPhuong.Name = "txtXaPhuong"
        Me.txtXaPhuong.Size = New System.Drawing.Size(69, 22)
        Me.txtXaPhuong.TabIndex = 4
        Me.txtXaPhuong.Text = "0"
        Me.txtXaPhuong.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'frmMangLuoiDonVi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtXaPhuong)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.numQuy)
        Me.Controls.Add(Me.dtpkNam)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.gridMLDV)
        Me.Controls.Add(Me.Panel10)
        Me.Controls.Add(Me.txtGhiChu)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.txtDuNo)
        Me.Controls.Add(Me.Label42)
        Me.Controls.Add(Me.txtSoDiemGD)
        Me.Controls.Add(Me.Label44)
        Me.Controls.Add(Me.cboDV)
        Me.Controls.Add(Me.Label35)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmMangLuoiDonVi"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: MẠNG LƯỚI HOẠT ĐỘNG CỦA ĐƠN VỊ"
        Me.Panel10.ResumeLayout(False)
        CType(Me.gridMLDV, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numQuy, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents bntNew As System.Windows.Forms.Button
    Friend WithEvents bntUpdate As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents bntDelete As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents gridMLDV As System.Windows.Forms.DataGridView
    Friend WithEvents txtGhiChu As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents txtDuNo As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtSoDiemGD As System.Windows.Forms.TextBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents cboDV As System.Windows.Forms.ComboBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents numQuy As System.Windows.Forms.NumericUpDown
    Friend WithEvents dtpkNam As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents labAlert As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtXaPhuong As System.Windows.Forms.TextBox
    Friend WithEvents IdML As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nam As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Quy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdPGD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SoXaPhuong As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SoDiemGD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DuNo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
