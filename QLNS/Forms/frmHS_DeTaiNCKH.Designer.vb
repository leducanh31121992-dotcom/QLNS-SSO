<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHS_DeTaiNCKH
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHS_DeTaiNCKH))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtGhichu_NC = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.dpkDenNgay_NC = New System.Windows.Forms.DateTimePicker()
        Me.dpkTuNgay_NC = New System.Windows.Forms.DateTimePicker()
        Me.gridDeTai = New System.Windows.Forms.DataGridView()
        Me.cln_cbNCKH = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.IdDeTaiNCKH = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdCapDeTai = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TenDeTai = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DonVi_QL = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtDonViQL_NC = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.cboCapDeTai_NC = New System.Windows.Forms.ComboBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtNoiDung_NC = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.dpkNgayNT_NC = New System.Windows.Forms.DateTimePicker()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtTenDeTai_NC = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.bntNew = New System.Windows.Forms.Button()
        Me.bntUpdate = New System.Windows.Forms.Button()
        Me.bntCancel = New System.Windows.Forms.Button()
        Me.bntDelete = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.Panel2.SuspendLayout()
        CType(Me.gridDeTai, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel11.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtGhichu_NC)
        Me.Panel2.Controls.Add(Me.Label26)
        Me.Panel2.Controls.Add(Me.Label25)
        Me.Panel2.Controls.Add(Me.dpkDenNgay_NC)
        Me.Panel2.Controls.Add(Me.dpkTuNgay_NC)
        Me.Panel2.Controls.Add(Me.gridDeTai)
        Me.Panel2.Controls.Add(Me.txtDonViQL_NC)
        Me.Panel2.Controls.Add(Me.Label18)
        Me.Panel2.Controls.Add(Me.Label19)
        Me.Panel2.Controls.Add(Me.cboCapDeTai_NC)
        Me.Panel2.Controls.Add(Me.Label21)
        Me.Panel2.Controls.Add(Me.txtNoiDung_NC)
        Me.Panel2.Controls.Add(Me.Label22)
        Me.Panel2.Controls.Add(Me.dpkNgayNT_NC)
        Me.Panel2.Controls.Add(Me.Label23)
        Me.Panel2.Controls.Add(Me.txtTenDeTai_NC)
        Me.Panel2.Controls.Add(Me.Label24)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(592, 481)
        Me.Panel2.TabIndex = 1
        '
        'txtGhichu_NC
        '
        Me.txtGhichu_NC.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhichu_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhichu_NC.Location = New System.Drawing.Point(100, 414)
        Me.txtGhichu_NC.Multiline = True
        Me.txtGhichu_NC.Name = "txtGhichu_NC"
        Me.txtGhichu_NC.Size = New System.Drawing.Size(490, 40)
        Me.txtGhichu_NC.TabIndex = 92
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(51, 417)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(48, 14)
        Me.Label26.TabIndex = 99
        Me.Label26.Text = "Ghi chú"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(246, 325)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(59, 14)
        Me.Label25.TabIndex = 98
        Me.Label25.Text = "Đến ngày"
        '
        'dpkDenNgay_NC
        '
        Me.dpkDenNgay_NC.CustomFormat = "dd/MM/yyyy"
        Me.dpkDenNgay_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkDenNgay_NC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkDenNgay_NC.Location = New System.Drawing.Point(307, 321)
        Me.dpkDenNgay_NC.Name = "dpkDenNgay_NC"
        Me.dpkDenNgay_NC.Size = New System.Drawing.Size(93, 22)
        Me.dpkDenNgay_NC.TabIndex = 84
        '
        'dpkTuNgay_NC
        '
        Me.dpkTuNgay_NC.CustomFormat = "dd/MM/yyyy"
        Me.dpkTuNgay_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkTuNgay_NC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkTuNgay_NC.Location = New System.Drawing.Point(100, 321)
        Me.dpkTuNgay_NC.Name = "dpkTuNgay_NC"
        Me.dpkTuNgay_NC.Size = New System.Drawing.Size(93, 22)
        Me.dpkTuNgay_NC.TabIndex = 83
        '
        'gridDeTai
        '
        Me.gridDeTai.AllowUserToAddRows = False
        Me.gridDeTai.AllowUserToDeleteRows = False
        Me.gridDeTai.AllowUserToResizeColumns = False
        Me.gridDeTai.AllowUserToResizeRows = False
        Me.gridDeTai.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridDeTai.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridDeTai.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridDeTai.ColumnHeadersHeight = 34
        Me.gridDeTai.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.cln_cbNCKH, Me.IdDeTaiNCKH, Me.IdCapDeTai, Me.TenDeTai, Me.DonVi_QL})
        Me.gridDeTai.Location = New System.Drawing.Point(0, 0)
        Me.gridDeTai.MultiSelect = False
        Me.gridDeTai.Name = "gridDeTai"
        Me.gridDeTai.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridDeTai.RowHeadersVisible = False
        Me.gridDeTai.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridDeTai.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridDeTai.Size = New System.Drawing.Size(591, 277)
        Me.gridDeTai.TabIndex = 97
        Me.gridDeTai.Tag = ""
        '
        'cln_cbNCKH
        '
        Me.cln_cbNCKH.HeaderText = "Chọn xoá"
        Me.cln_cbNCKH.Name = "cln_cbNCKH"
        Me.cln_cbNCKH.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.cln_cbNCKH.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.cln_cbNCKH.Width = 65
        '
        'IdDeTaiNCKH
        '
        Me.IdDeTaiNCKH.DataPropertyName = "IdDeTai"
        Me.IdDeTaiNCKH.HeaderText = "ID"
        Me.IdDeTaiNCKH.Name = "IdDeTaiNCKH"
        Me.IdDeTaiNCKH.ReadOnly = True
        Me.IdDeTaiNCKH.Visible = False
        '
        'IdCapDeTai
        '
        Me.IdCapDeTai.DataPropertyName = "IdCapDeTai"
        Me.IdCapDeTai.HeaderText = "Cấp đề tài"
        Me.IdCapDeTai.Name = "IdCapDeTai"
        Me.IdCapDeTai.ReadOnly = True
        Me.IdCapDeTai.Width = 110
        '
        'TenDeTai
        '
        Me.TenDeTai.DataPropertyName = "TenDeTai"
        Me.TenDeTai.HeaderText = "Tên đề tài"
        Me.TenDeTai.Name = "TenDeTai"
        Me.TenDeTai.ReadOnly = True
        Me.TenDeTai.Width = 367
        '
        'DonVi_QL
        '
        Me.DonVi_QL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DonVi_QL.DataPropertyName = "DonVi_QL"
        Me.DonVi_QL.FillWeight = 150.0!
        Me.DonVi_QL.HeaderText = "Đơn vị chủ trì"
        Me.DonVi_QL.Name = "DonVi_QL"
        Me.DonVi_QL.ReadOnly = True
        Me.DonVi_QL.Width = 87
        '
        'txtDonViQL_NC
        '
        Me.txtDonViQL_NC.BackColor = System.Drawing.SystemColors.Window
        Me.txtDonViQL_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDonViQL_NC.Location = New System.Drawing.Point(307, 391)
        Me.txtDonViQL_NC.Name = "txtDonViQL_NC"
        Me.txtDonViQL_NC.Size = New System.Drawing.Size(283, 22)
        Me.txtDonViQL_NC.TabIndex = 90
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(46, 325)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(53, 14)
        Me.Label18.TabIndex = 96
        Me.Label18.Text = "Từ ngày"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(225, 393)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(80, 14)
        Me.Label19.TabIndex = 95
        Me.Label19.Text = "Đơn vị chủ trì"
        '
        'cboCapDeTai_NC
        '
        Me.cboCapDeTai_NC.BackColor = System.Drawing.SystemColors.Window
        Me.cboCapDeTai_NC.DisplayMember = "Display"
        Me.cboCapDeTai_NC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCapDeTai_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCapDeTai_NC.FormattingEnabled = True
        Me.cboCapDeTai_NC.Location = New System.Drawing.Point(100, 277)
        Me.cboCapDeTai_NC.Name = "cboCapDeTai_NC"
        Me.cboCapDeTai_NC.Size = New System.Drawing.Size(490, 22)
        Me.cboCapDeTai_NC.TabIndex = 85
        Me.cboCapDeTai_NC.ValueMember = "Value"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(37, 280)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(62, 14)
        Me.Label21.TabIndex = 94
        Me.Label21.Text = "Cấp đề tài"
        '
        'txtNoiDung_NC
        '
        Me.txtNoiDung_NC.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoiDung_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNoiDung_NC.Location = New System.Drawing.Point(100, 343)
        Me.txtNoiDung_NC.Multiline = True
        Me.txtNoiDung_NC.Name = "txtNoiDung_NC"
        Me.txtNoiDung_NC.Size = New System.Drawing.Size(490, 48)
        Me.txtNoiDung_NC.TabIndex = 87
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(8, 348)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(91, 14)
        Me.Label22.TabIndex = 93
        Me.Label22.Text = "Nội dung đề tài"
        '
        'dpkNgayNT_NC
        '
        Me.dpkNgayNT_NC.Checked = False
        Me.dpkNgayNT_NC.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayNT_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayNT_NC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayNT_NC.Location = New System.Drawing.Point(100, 391)
        Me.dpkNgayNT_NC.Name = "dpkNgayNT_NC"
        Me.dpkNgayNT_NC.ShowCheckBox = True
        Me.dpkNgayNT_NC.Size = New System.Drawing.Size(108, 22)
        Me.dpkNgayNT_NC.TabIndex = 89
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(-1, 395)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(101, 14)
        Me.Label23.TabIndex = 91
        Me.Label23.Text = "Ngày nghiệm thu"
        '
        'txtTenDeTai_NC
        '
        Me.txtTenDeTai_NC.BackColor = System.Drawing.SystemColors.Window
        Me.txtTenDeTai_NC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTenDeTai_NC.Location = New System.Drawing.Point(100, 299)
        Me.txtTenDeTai_NC.Name = "txtTenDeTai_NC"
        Me.txtTenDeTai_NC.Size = New System.Drawing.Size(490, 22)
        Me.txtTenDeTai_NC.TabIndex = 86
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(35, 304)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(64, 14)
        Me.Label24.TabIndex = 88
        Me.Label24.Text = "Tên đề tài"
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.DarkGray
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel11.Controls.Add(Me.bntNew)
        Me.Panel11.Controls.Add(Me.bntUpdate)
        Me.Panel11.Controls.Add(Me.bntCancel)
        Me.Panel11.Controls.Add(Me.bntDelete)
        Me.Panel11.Controls.Add(Me.bntClose)
        Me.Panel11.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel11.Location = New System.Drawing.Point(0, 455)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(592, 26)
        Me.Panel11.TabIndex = 10
        '
        'bntNew
        '
        Me.bntNew.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntNew.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntNew.Location = New System.Drawing.Point(271, 0)
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
        Me.bntUpdate.Location = New System.Drawing.Point(344, 0)
        Me.bntUpdate.Name = "bntUpdate"
        Me.bntUpdate.Size = New System.Drawing.Size(61, 22)
        Me.bntUpdate.TabIndex = 1
        Me.bntUpdate.Text = "&Ghi"
        Me.bntUpdate.UseVisualStyleBackColor = True
        '
        'bntCancel
        '
        Me.bntCancel.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCancel.Location = New System.Drawing.Point(405, 0)
        Me.bntCancel.Name = "bntCancel"
        Me.bntCancel.Size = New System.Drawing.Size(61, 22)
        Me.bntCancel.TabIndex = 2
        Me.bntCancel.Text = "&Bỏ qua"
        Me.bntCancel.UseVisualStyleBackColor = True
        '
        'bntDelete
        '
        Me.bntDelete.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntDelete.Enabled = False
        Me.bntDelete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntDelete.Location = New System.Drawing.Point(466, 0)
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
        Me.bntClose.Location = New System.Drawing.Point(527, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(61, 22)
        Me.bntClose.TabIndex = 5
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'frmHS_DeTaiNCKH
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(592, 481)
        Me.Controls.Add(Me.Panel11)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmHS_DeTaiNCKH"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ĐỀ TÀI NGHIÊN CỨU KHOA HỌC"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.gridDeTai, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel11.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtGhichu_NC As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents dpkDenNgay_NC As System.Windows.Forms.DateTimePicker
    Friend WithEvents dpkTuNgay_NC As System.Windows.Forms.DateTimePicker
    Friend WithEvents gridDeTai As System.Windows.Forms.DataGridView
    Friend WithEvents txtDonViQL_NC As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents cboCapDeTai_NC As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtNoiDung_NC As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents dpkNgayNT_NC As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtTenDeTai_NC As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents bntNew As System.Windows.Forms.Button
    Friend WithEvents bntUpdate As System.Windows.Forms.Button
    Friend WithEvents bntCancel As System.Windows.Forms.Button
    Friend WithEvents bntDelete As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents cln_cbNCKH As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents IdDeTaiNCKH As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdCapDeTai As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TenDeTai As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DonVi_QL As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
