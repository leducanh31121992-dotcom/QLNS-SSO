<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKhenThuongGroup
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.treeCocau = New System.Windows.Forms.TreeView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlGroup = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.gridKT = New System.Windows.Forms.DataGridView()
        Me.IdKhenThuong_CB = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cln_cbKT = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Ngay_QD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdDanhHieuTD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.So_QD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdCanBo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.labListBC = New System.Windows.Forms.Label()
        Me.cbAllDelete = New System.Windows.Forms.CheckBox()
        Me.cboDanhhieuTD = New System.Windows.Forms.ComboBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.bntNewKT = New System.Windows.Forms.Button()
        Me.bntUpdateKT = New System.Windows.Forms.Button()
        Me.bntCancelKT = New System.Windows.Forms.Button()
        Me.bntDeleteKT = New System.Windows.Forms.Button()
        Me.bntCloseKT = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtGhichu_KT = New System.Windows.Forms.TextBox()
        Me.txtSoQD_KT = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtMuc_KT = New System.Windows.Forms.TextBox()
        Me.dpkNgayQD_KT = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtNguoiQD_KT = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboHinhthuc_KT = New System.Windows.Forms.ComboBox()
        Me.cboCap_KT = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.pnlGroup.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.gridKT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.treeCocau)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(204, 563)
        Me.Panel1.TabIndex = 4
        '
        'treeCocau
        '
        Me.treeCocau.CheckBoxes = True
        Me.treeCocau.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeCocau.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.treeCocau.Location = New System.Drawing.Point(0, 21)
        Me.treeCocau.Name = "treeCocau"
        Me.treeCocau.Size = New System.Drawing.Size(204, 542)
        Me.treeCocau.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(204, 21)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "CHỌN DANH SÁCH CÁN BỘ"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlGroup
        '
        Me.pnlGroup.Controls.Add(Me.Panel2)
        Me.pnlGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlGroup.Location = New System.Drawing.Point(204, 0)
        Me.pnlGroup.Name = "pnlGroup"
        Me.pnlGroup.Padding = New System.Windows.Forms.Padding(3)
        Me.pnlGroup.Size = New System.Drawing.Size(583, 563)
        Me.pnlGroup.TabIndex = 5
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.gridKT)
        Me.Panel2.Controls.Add(Me.Panel4)
        Me.Panel2.Controls.Add(Me.cboDanhhieuTD)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.txtGhichu_KT)
        Me.Panel2.Controls.Add(Me.txtSoQD_KT)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.txtMuc_KT)
        Me.Panel2.Controls.Add(Me.dpkNgayQD_KT)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.txtNguoiQD_KT)
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.cboHinhthuc_KT)
        Me.Panel2.Controls.Add(Me.cboCap_KT)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Location = New System.Drawing.Point(4, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(575, 558)
        Me.Panel2.TabIndex = 124
        '
        'gridKT
        '
        Me.gridKT.AllowUserToAddRows = False
        Me.gridKT.AllowUserToDeleteRows = False
        Me.gridKT.AllowUserToResizeColumns = False
        Me.gridKT.AllowUserToResizeRows = False
        Me.gridKT.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridKT.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridKT.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridKT.ColumnHeadersHeight = 34
        Me.gridKT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdKhenThuong_CB, Me.cln_cbKT, Me.Ngay_QD, Me.IdDanhHieuTD, Me.So_QD, Me.IdCanBo})
        Me.gridKT.Location = New System.Drawing.Point(1, 21)
        Me.gridKT.MultiSelect = False
        Me.gridKT.Name = "gridKT"
        Me.gridKT.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridKT.RowHeadersVisible = False
        Me.gridKT.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridKT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridKT.Size = New System.Drawing.Size(574, 378)
        Me.gridKT.TabIndex = 126
        Me.gridKT.Tag = ""
        '
        'IdKhenThuong_CB
        '
        Me.IdKhenThuong_CB.DataPropertyName = "IdKhenThuong_CB"
        Me.IdKhenThuong_CB.HeaderText = "ID"
        Me.IdKhenThuong_CB.Name = "IdKhenThuong_CB"
        Me.IdKhenThuong_CB.ReadOnly = True
        Me.IdKhenThuong_CB.Visible = False
        '
        'cln_cbKT
        '
        Me.cln_cbKT.HeaderText = "Chọn xoá"
        Me.cln_cbKT.Name = "cln_cbKT"
        Me.cln_cbKT.Width = 65
        '
        'Ngay_QD
        '
        Me.Ngay_QD.DataPropertyName = "Ngay_QD"
        Me.Ngay_QD.HeaderText = "Ngày"
        Me.Ngay_QD.Name = "Ngay_QD"
        Me.Ngay_QD.ReadOnly = True
        Me.Ngay_QD.Width = 85
        '
        'IdDanhHieuTD
        '
        Me.IdDanhHieuTD.DataPropertyName = "IdDanhHieuTD"
        Me.IdDanhHieuTD.HeaderText = "Danh hiệu thi đua"
        Me.IdDanhHieuTD.Name = "IdDanhHieuTD"
        Me.IdDanhHieuTD.ReadOnly = True
        Me.IdDanhHieuTD.Width = 200
        '
        'So_QD
        '
        Me.So_QD.DataPropertyName = "So_QD"
        Me.So_QD.HeaderText = "Số quyết định"
        Me.So_QD.Name = "So_QD"
        Me.So_QD.Width = 90
        '
        'IdCanBo
        '
        Me.IdCanBo.DataPropertyName = "IdCanBo"
        Me.IdCanBo.HeaderText = "Cán bộ"
        Me.IdCanBo.Name = "IdCanBo"
        Me.IdCanBo.ReadOnly = True
        Me.IdCanBo.Width = 160
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkGray
        Me.Panel4.Controls.Add(Me.labListBC)
        Me.Panel4.Controls.Add(Me.cbAllDelete)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(575, 21)
        Me.Panel4.TabIndex = 125
        '
        'labListBC
        '
        Me.labListBC.BackColor = System.Drawing.Color.DarkGray
        Me.labListBC.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labListBC.ForeColor = System.Drawing.Color.Maroon
        Me.labListBC.Location = New System.Drawing.Point(123, 0)
        Me.labListBC.Name = "labListBC"
        Me.labListBC.Size = New System.Drawing.Size(452, 21)
        Me.labListBC.TabIndex = 106
        Me.labListBC.Text = "Danh sách Khen thưởng"
        Me.labListBC.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbAllDelete
        '
        Me.cbAllDelete.AutoSize = True
        Me.cbAllDelete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAllDelete.ForeColor = System.Drawing.Color.DarkRed
        Me.cbAllDelete.Location = New System.Drawing.Point(5, 3)
        Me.cbAllDelete.Name = "cbAllDelete"
        Me.cbAllDelete.Size = New System.Drawing.Size(124, 17)
        Me.cbAllDelete.TabIndex = 0
        Me.cbAllDelete.Text = "Chọn xoá toàn bộ"
        Me.cbAllDelete.UseVisualStyleBackColor = True
        '
        'cboDanhhieuTD
        '
        Me.cboDanhhieuTD.BackColor = System.Drawing.SystemColors.Window
        Me.cboDanhhieuTD.DisplayMember = "Display"
        Me.cboDanhhieuTD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDanhhieuTD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDanhhieuTD.FormattingEnabled = True
        Me.cboDanhhieuTD.Location = New System.Drawing.Point(141, 399)
        Me.cboDanhhieuTD.Name = "cboDanhhieuTD"
        Me.cboDanhhieuTD.Size = New System.Drawing.Size(292, 22)
        Me.cboDanhhieuTD.TabIndex = 107
        Me.cboDanhhieuTD.ValueMember = "Value"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.bntNewKT)
        Me.Panel3.Controls.Add(Me.bntUpdateKT)
        Me.Panel3.Controls.Add(Me.bntCancelKT)
        Me.Panel3.Controls.Add(Me.bntDeleteKT)
        Me.Panel3.Controls.Add(Me.bntCloseKT)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 532)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(575, 26)
        Me.Panel3.TabIndex = 117
        '
        'bntNewKT
        '
        Me.bntNewKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntNewKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntNewKT.Location = New System.Drawing.Point(254, 0)
        Me.bntNewKT.Name = "bntNewKT"
        Me.bntNewKT.Size = New System.Drawing.Size(73, 22)
        Me.bntNewKT.TabIndex = 4
        Me.bntNewKT.Text = "&Thêm mới"
        Me.bntNewKT.UseVisualStyleBackColor = True
        '
        'bntUpdateKT
        '
        Me.bntUpdateKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntUpdateKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntUpdateKT.Location = New System.Drawing.Point(327, 0)
        Me.bntUpdateKT.Name = "bntUpdateKT"
        Me.bntUpdateKT.Size = New System.Drawing.Size(61, 22)
        Me.bntUpdateKT.TabIndex = 1
        Me.bntUpdateKT.Text = "&Ghi"
        Me.bntUpdateKT.UseVisualStyleBackColor = True
        '
        'bntCancelKT
        '
        Me.bntCancelKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCancelKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCancelKT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCancelKT.Location = New System.Drawing.Point(388, 0)
        Me.bntCancelKT.Name = "bntCancelKT"
        Me.bntCancelKT.Size = New System.Drawing.Size(61, 22)
        Me.bntCancelKT.TabIndex = 2
        Me.bntCancelKT.Text = "&Bỏ qua"
        Me.bntCancelKT.UseVisualStyleBackColor = True
        '
        'bntDeleteKT
        '
        Me.bntDeleteKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntDeleteKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntDeleteKT.Location = New System.Drawing.Point(449, 0)
        Me.bntDeleteKT.Name = "bntDeleteKT"
        Me.bntDeleteKT.Size = New System.Drawing.Size(61, 22)
        Me.bntDeleteKT.TabIndex = 3
        Me.bntDeleteKT.Text = "&Xoá"
        Me.bntDeleteKT.UseVisualStyleBackColor = True
        '
        'bntCloseKT
        '
        Me.bntCloseKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCloseKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCloseKT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCloseKT.Location = New System.Drawing.Point(510, 0)
        Me.bntCloseKT.Name = "bntCloseKT"
        Me.bntCloseKT.Size = New System.Drawing.Size(61, 22)
        Me.bntCloseKT.TabIndex = 5
        Me.bntCloseKT.Text = "&Quay ra"
        Me.bntCloseKT.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(55, 447)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(84, 14)
        Me.Label10.TabIndex = 112
        Me.Label10.Text = "Số quyết định"
        '
        'txtGhichu_KT
        '
        Me.txtGhichu_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhichu_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhichu_KT.Location = New System.Drawing.Point(141, 509)
        Me.txtGhichu_KT.Name = "txtGhichu_KT"
        Me.txtGhichu_KT.Size = New System.Drawing.Size(433, 22)
        Me.txtGhichu_KT.TabIndex = 116
        '
        'txtSoQD_KT
        '
        Me.txtSoQD_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoQD_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoQD_KT.Location = New System.Drawing.Point(141, 443)
        Me.txtSoQD_KT.Name = "txtSoQD_KT"
        Me.txtSoQD_KT.Size = New System.Drawing.Size(86, 22)
        Me.txtSoQD_KT.TabIndex = 109
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(91, 512)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 14)
        Me.Label14.TabIndex = 123
        Me.Label14.Text = "Ghi chú"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(231, 447)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(34, 14)
        Me.Label9.TabIndex = 115
        Me.Label9.Text = "Ngày"
        '
        'txtMuc_KT
        '
        Me.txtMuc_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtMuc_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMuc_KT.Location = New System.Drawing.Point(141, 487)
        Me.txtMuc_KT.Name = "txtMuc_KT"
        Me.txtMuc_KT.Size = New System.Drawing.Size(433, 22)
        Me.txtMuc_KT.TabIndex = 114
        '
        'dpkNgayQD_KT
        '
        Me.dpkNgayQD_KT.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayQD_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayQD_KT.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayQD_KT.Location = New System.Drawing.Point(265, 443)
        Me.dpkNgayQD_KT.Name = "dpkNgayQD_KT"
        Me.dpkNgayQD_KT.Size = New System.Drawing.Size(90, 22)
        Me.dpkNgayQD_KT.TabIndex = 110
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(33, 490)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(106, 14)
        Me.Label13.TabIndex = 122
        Me.Label13.Text = "Mức khen thưởng"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(357, 447)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 14)
        Me.Label8.TabIndex = 118
        Me.Label8.Text = "Người QĐ"
        '
        'txtNguoiQD_KT
        '
        Me.txtNguoiQD_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtNguoiQD_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNguoiQD_KT.Location = New System.Drawing.Point(417, 443)
        Me.txtNguoiQD_KT.Name = "txtNguoiQD_KT"
        Me.txtNguoiQD_KT.Size = New System.Drawing.Size(157, 22)
        Me.txtNguoiQD_KT.TabIndex = 111
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(77, 402)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(62, 14)
        Me.Label12.TabIndex = 121
        Me.Label12.Text = "Danh hiệu"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(36, 424)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(103, 14)
        Me.Label7.TabIndex = 119
        Me.Label7.Text = "Cấp khen thưởng"
        '
        'cboHinhthuc_KT
        '
        Me.cboHinhthuc_KT.BackColor = System.Drawing.SystemColors.Window
        Me.cboHinhthuc_KT.DisplayMember = "Display"
        Me.cboHinhthuc_KT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboHinhthuc_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboHinhthuc_KT.FormattingEnabled = True
        Me.cboHinhthuc_KT.Location = New System.Drawing.Point(141, 465)
        Me.cboHinhthuc_KT.Name = "cboHinhthuc_KT"
        Me.cboHinhthuc_KT.Size = New System.Drawing.Size(433, 22)
        Me.cboHinhthuc_KT.TabIndex = 113
        Me.cboHinhthuc_KT.ValueMember = "Value"
        '
        'cboCap_KT
        '
        Me.cboCap_KT.BackColor = System.Drawing.SystemColors.Window
        Me.cboCap_KT.DisplayMember = "Display"
        Me.cboCap_KT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCap_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCap_KT.FormattingEnabled = True
        Me.cboCap_KT.Location = New System.Drawing.Point(141, 421)
        Me.cboCap_KT.Name = "cboCap_KT"
        Me.cboCap_KT.Size = New System.Drawing.Size(433, 22)
        Me.cboCap_KT.TabIndex = 108
        Me.cboCap_KT.ValueMember = "Value"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(2, 468)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(137, 14)
        Me.Label11.TabIndex = 120
        Me.Label11.Text = "Hình thức khen thưởng"
        '
        'frmKhenThuongGroup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.pnlGroup)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmKhenThuongGroup"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: KHEN THƯỞNG CHO NHIỀU CÁN BỘ"
        Me.Panel1.ResumeLayout(False)
        Me.pnlGroup.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.gridKT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents treeCocau As System.Windows.Forms.TreeView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlGroup As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents bntNewKT As System.Windows.Forms.Button
    Friend WithEvents bntUpdateKT As System.Windows.Forms.Button
    Friend WithEvents bntCancelKT As System.Windows.Forms.Button
    Friend WithEvents bntDeleteKT As System.Windows.Forms.Button
    Friend WithEvents bntCloseKT As System.Windows.Forms.Button
    Friend WithEvents txtGhichu_KT As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtMuc_KT As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboDanhhieuTD As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cboHinhthuc_KT As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboCap_KT As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtNguoiQD_KT As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dpkNgayQD_KT As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtSoQD_KT As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents labListBC As System.Windows.Forms.Label
    Friend WithEvents cbAllDelete As System.Windows.Forms.CheckBox
    Friend WithEvents gridKT As System.Windows.Forms.DataGridView
    Friend WithEvents IdKhenThuong_CB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cln_cbKT As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Ngay_QD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdDanhHieuTD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents So_QD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdCanBo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
