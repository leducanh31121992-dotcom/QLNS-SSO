<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLuongBosungGroup
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
        Me.treeCocau = New System.Windows.Forms.TreeView()
        Me.pnlGroup = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.labListBC = New System.Windows.Forms.Label()
        Me.cbAllDelete = New System.Windows.Forms.CheckBox()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.bntNewBS = New System.Windows.Forms.Button()
        Me.bntUpdateBS = New System.Windows.Forms.Button()
        Me.bntCancelBS = New System.Windows.Forms.Button()
        Me.bntDeleteBS = New System.Windows.Forms.Button()
        Me.bntCloseBS = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtGhichuBS = New System.Windows.Forms.TextBox()
        Me.Label100 = New System.Windows.Forms.Label()
        Me.txtNamBS = New System.Windows.Forms.TextBox()
        Me.Label94 = New System.Windows.Forms.Label()
        Me.Label99 = New System.Windows.Forms.Label()
        Me.txtHesochi = New System.Windows.Forms.TextBox()
        Me.Label95 = New System.Windows.Forms.Label()
        Me.cboLoaiLuongBS = New System.Windows.Forms.ComboBox()
        Me.Label93 = New System.Windows.Forms.Label()
        Me.txtThuclinhBS = New System.Windows.Forms.TextBox()
        Me.txtThangBS = New System.Windows.Forms.TextBox()
        Me.Label90 = New System.Windows.Forms.Label()
        Me.gridLuongBS = New System.Windows.Forms.DataGridView()
        Me.Id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cln_cbBS = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.IdCanBo_LBS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Thang = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdLoaiLuongBS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdCanbo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Hesochi = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ThucLinh = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Ghichu = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlGroup.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel11.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.gridLuongBS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
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
        'pnlGroup
        '
        Me.pnlGroup.Controls.Add(Me.Panel2)
        Me.pnlGroup.Controls.Add(Me.Panel11)
        Me.pnlGroup.Controls.Add(Me.GroupBox1)
        Me.pnlGroup.Controls.Add(Me.gridLuongBS)
        Me.pnlGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlGroup.Location = New System.Drawing.Point(204, 0)
        Me.pnlGroup.Name = "pnlGroup"
        Me.pnlGroup.Padding = New System.Windows.Forms.Padding(3)
        Me.pnlGroup.Size = New System.Drawing.Size(583, 563)
        Me.pnlGroup.TabIndex = 2
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkGray
        Me.Panel2.Controls.Add(Me.labListBC)
        Me.Panel2.Controls.Add(Me.cbAllDelete)
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(583, 21)
        Me.Panel2.TabIndex = 106
        '
        'labListBC
        '
        Me.labListBC.BackColor = System.Drawing.Color.DarkGray
        Me.labListBC.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labListBC.ForeColor = System.Drawing.Color.Maroon
        Me.labListBC.Location = New System.Drawing.Point(123, 0)
        Me.labListBC.Name = "labListBC"
        Me.labListBC.Size = New System.Drawing.Size(457, 21)
        Me.labListBC.TabIndex = 106
        Me.labListBC.Text = "Danh sách Lương bổ sung"
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
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.DarkGray
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel11.Controls.Add(Me.bntNewBS)
        Me.Panel11.Controls.Add(Me.bntUpdateBS)
        Me.Panel11.Controls.Add(Me.bntCancelBS)
        Me.Panel11.Controls.Add(Me.bntDeleteBS)
        Me.Panel11.Controls.Add(Me.bntCloseBS)
        Me.Panel11.Location = New System.Drawing.Point(4, 533)
        Me.Panel11.Margin = New System.Windows.Forms.Padding(5)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(575, 26)
        Me.Panel11.TabIndex = 7
        '
        'bntNewBS
        '
        Me.bntNewBS.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntNewBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntNewBS.Location = New System.Drawing.Point(254, 0)
        Me.bntNewBS.Name = "bntNewBS"
        Me.bntNewBS.Size = New System.Drawing.Size(73, 22)
        Me.bntNewBS.TabIndex = 4
        Me.bntNewBS.Text = "&Thêm mới"
        Me.bntNewBS.UseVisualStyleBackColor = True
        '
        'bntUpdateBS
        '
        Me.bntUpdateBS.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntUpdateBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntUpdateBS.Location = New System.Drawing.Point(327, 0)
        Me.bntUpdateBS.Name = "bntUpdateBS"
        Me.bntUpdateBS.Size = New System.Drawing.Size(61, 22)
        Me.bntUpdateBS.TabIndex = 1
        Me.bntUpdateBS.Text = "&Ghi"
        Me.bntUpdateBS.UseVisualStyleBackColor = True
        '
        'bntCancelBS
        '
        Me.bntCancelBS.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCancelBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCancelBS.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCancelBS.Location = New System.Drawing.Point(388, 0)
        Me.bntCancelBS.Name = "bntCancelBS"
        Me.bntCancelBS.Size = New System.Drawing.Size(61, 22)
        Me.bntCancelBS.TabIndex = 2
        Me.bntCancelBS.Text = "&Bỏ qua"
        Me.bntCancelBS.UseVisualStyleBackColor = True
        '
        'bntDeleteBS
        '
        Me.bntDeleteBS.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntDeleteBS.Enabled = False
        Me.bntDeleteBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntDeleteBS.Location = New System.Drawing.Point(449, 0)
        Me.bntDeleteBS.Name = "bntDeleteBS"
        Me.bntDeleteBS.Size = New System.Drawing.Size(61, 22)
        Me.bntDeleteBS.TabIndex = 3
        Me.bntDeleteBS.Text = "&Xoá"
        Me.bntDeleteBS.UseVisualStyleBackColor = True
        '
        'bntCloseBS
        '
        Me.bntCloseBS.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCloseBS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCloseBS.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCloseBS.Location = New System.Drawing.Point(510, 0)
        Me.bntCloseBS.Name = "bntCloseBS"
        Me.bntCloseBS.Size = New System.Drawing.Size(61, 22)
        Me.bntCloseBS.TabIndex = 5
        Me.bntCloseBS.Text = "&Quay ra"
        Me.bntCloseBS.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtGhichuBS)
        Me.GroupBox1.Controls.Add(Me.Label100)
        Me.GroupBox1.Controls.Add(Me.txtNamBS)
        Me.GroupBox1.Controls.Add(Me.Label94)
        Me.GroupBox1.Controls.Add(Me.Label99)
        Me.GroupBox1.Controls.Add(Me.txtHesochi)
        Me.GroupBox1.Controls.Add(Me.Label95)
        Me.GroupBox1.Controls.Add(Me.cboLoaiLuongBS)
        Me.GroupBox1.Controls.Add(Me.Label93)
        Me.GroupBox1.Controls.Add(Me.txtThuclinhBS)
        Me.GroupBox1.Controls.Add(Me.txtThangBS)
        Me.GroupBox1.Controls.Add(Me.Label90)
        Me.GroupBox1.Location = New System.Drawing.Point(1, 403)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(579, 132)
        Me.GroupBox1.TabIndex = 105
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Thông tin lương bổ sung"
        '
        'txtGhichuBS
        '
        Me.txtGhichuBS.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhichuBS.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhichuBS.Location = New System.Drawing.Point(87, 85)
        Me.txtGhichuBS.Multiline = True
        Me.txtGhichuBS.Name = "txtGhichuBS"
        Me.txtGhichuBS.Size = New System.Drawing.Size(490, 44)
        Me.txtGhichuBS.TabIndex = 6
        '
        'Label100
        '
        Me.Label100.AutoSize = True
        Me.Label100.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label100.Location = New System.Drawing.Point(54, 22)
        Me.Label100.Name = "Label100"
        Me.Label100.Size = New System.Drawing.Size(31, 14)
        Me.Label100.TabIndex = 66
        Me.Label100.Text = "Năm"
        '
        'txtNamBS
        '
        Me.txtNamBS.BackColor = System.Drawing.SystemColors.Window
        Me.txtNamBS.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNamBS.Location = New System.Drawing.Point(87, 19)
        Me.txtNamBS.Name = "txtNamBS"
        Me.txtNamBS.Size = New System.Drawing.Size(55, 22)
        Me.txtNamBS.TabIndex = 1
        Me.txtNamBS.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label94
        '
        Me.Label94.AutoSize = True
        Me.Label94.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label94.Location = New System.Drawing.Point(5, 66)
        Me.Label94.Name = "Label94"
        Me.Label94.Size = New System.Drawing.Size(81, 14)
        Me.Label94.TabIndex = 73
        Me.Label94.Text = "Hệ số chi/thu"
        '
        'Label99
        '
        Me.Label99.AutoSize = True
        Me.Label99.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label99.Location = New System.Drawing.Point(144, 22)
        Me.Label99.Name = "Label99"
        Me.Label99.Size = New System.Drawing.Size(42, 14)
        Me.Label99.TabIndex = 67
        Me.Label99.Text = "Tháng"
        '
        'txtHesochi
        '
        Me.txtHesochi.BackColor = System.Drawing.SystemColors.Window
        Me.txtHesochi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHesochi.Location = New System.Drawing.Point(87, 63)
        Me.txtHesochi.Name = "txtHesochi"
        Me.txtHesochi.Size = New System.Drawing.Size(146, 22)
        Me.txtHesochi.TabIndex = 4
        Me.txtHesochi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label95
        '
        Me.Label95.AutoSize = True
        Me.Label95.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label95.Location = New System.Drawing.Point(9, 44)
        Me.Label95.Name = "Label95"
        Me.Label95.Size = New System.Drawing.Size(76, 14)
        Me.Label95.TabIndex = 68
        Me.Label95.Text = "Loại bổ sung"
        '
        'cboLoaiLuongBS
        '
        Me.cboLoaiLuongBS.DisplayMember = "Display"
        Me.cboLoaiLuongBS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLoaiLuongBS.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLoaiLuongBS.FormattingEnabled = True
        Me.cboLoaiLuongBS.Location = New System.Drawing.Point(87, 41)
        Me.cboLoaiLuongBS.Name = "cboLoaiLuongBS"
        Me.cboLoaiLuongBS.Size = New System.Drawing.Size(490, 22)
        Me.cboLoaiLuongBS.TabIndex = 3
        Me.cboLoaiLuongBS.ValueMember = "Value"
        '
        'Label93
        '
        Me.Label93.AutoSize = True
        Me.Label93.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label93.Location = New System.Drawing.Point(285, 66)
        Me.Label93.Name = "Label93"
        Me.Label93.Size = New System.Drawing.Size(46, 14)
        Me.Label93.TabIndex = 69
        Me.Label93.Text = "Số tiền"
        '
        'txtThuclinhBS
        '
        Me.txtThuclinhBS.BackColor = System.Drawing.SystemColors.Window
        Me.txtThuclinhBS.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtThuclinhBS.Location = New System.Drawing.Point(332, 63)
        Me.txtThuclinhBS.Name = "txtThuclinhBS"
        Me.txtThuclinhBS.Size = New System.Drawing.Size(245, 22)
        Me.txtThuclinhBS.TabIndex = 5
        Me.txtThuclinhBS.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtThangBS
        '
        Me.txtThangBS.BackColor = System.Drawing.SystemColors.Window
        Me.txtThangBS.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtThangBS.Location = New System.Drawing.Point(188, 19)
        Me.txtThangBS.Name = "txtThangBS"
        Me.txtThangBS.Size = New System.Drawing.Size(45, 22)
        Me.txtThangBS.TabIndex = 2
        Me.txtThangBS.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label90
        '
        Me.Label90.AutoSize = True
        Me.Label90.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label90.Location = New System.Drawing.Point(37, 86)
        Me.Label90.Name = "Label90"
        Me.Label90.Size = New System.Drawing.Size(48, 14)
        Me.Label90.TabIndex = 70
        Me.Label90.Text = "Ghi chú"
        '
        'gridLuongBS
        '
        Me.gridLuongBS.AllowUserToAddRows = False
        Me.gridLuongBS.AllowUserToDeleteRows = False
        Me.gridLuongBS.AllowUserToResizeColumns = False
        Me.gridLuongBS.AllowUserToResizeRows = False
        Me.gridLuongBS.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridLuongBS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridLuongBS.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridLuongBS.ColumnHeadersHeight = 34
        Me.gridLuongBS.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Id, Me.cln_cbBS, Me.IdCanBo_LBS, Me.Thang, Me.IdLoaiLuongBS, Me.IdCanbo, Me.Hesochi, Me.ThucLinh, Me.Ghichu})
        Me.gridLuongBS.Location = New System.Drawing.Point(5, 21)
        Me.gridLuongBS.MultiSelect = False
        Me.gridLuongBS.Name = "gridLuongBS"
        Me.gridLuongBS.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridLuongBS.RowHeadersVisible = False
        Me.gridLuongBS.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridLuongBS.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridLuongBS.Size = New System.Drawing.Size(575, 376)
        Me.gridLuongBS.TabIndex = 71
        Me.gridLuongBS.Tag = ""
        '
        'Id
        '
        Me.Id.DataPropertyName = "Id"
        Me.Id.HeaderText = "IDBS"
        Me.Id.Name = "Id"
        Me.Id.ReadOnly = True
        Me.Id.Visible = False
        '
        'cln_cbBS
        '
        Me.cln_cbBS.HeaderText = "Chọn xoá"
        Me.cln_cbBS.Name = "cln_cbBS"
        Me.cln_cbBS.Width = 65
        '
        'IdCanBo_LBS
        '
        Me.IdCanBo_LBS.DataPropertyName = "IdCanBo"
        Me.IdCanBo_LBS.HeaderText = "IdCanBo_LBS"
        Me.IdCanBo_LBS.Name = "IdCanBo_LBS"
        Me.IdCanBo_LBS.ReadOnly = True
        Me.IdCanBo_LBS.Visible = False
        '
        'Thang
        '
        Me.Thang.DataPropertyName = "Thang"
        Me.Thang.HeaderText = "Tháng"
        Me.Thang.Name = "Thang"
        Me.Thang.ReadOnly = True
        Me.Thang.Width = 50
        '
        'IdLoaiLuongBS
        '
        Me.IdLoaiLuongBS.DataPropertyName = "IdLoaiLuongBS"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.IdLoaiLuongBS.DefaultCellStyle = DataGridViewCellStyle1
        Me.IdLoaiLuongBS.HeaderText = "HĐNH/TS"
        Me.IdLoaiLuongBS.Name = "IdLoaiLuongBS"
        Me.IdLoaiLuongBS.ReadOnly = True
        Me.IdLoaiLuongBS.Width = 60
        '
        'IdCanbo
        '
        Me.IdCanbo.DataPropertyName = "IdCanBo"
        Me.IdCanbo.HeaderText = "Cán bộ"
        Me.IdCanbo.Name = "IdCanbo"
        Me.IdCanbo.ReadOnly = True
        Me.IdCanbo.Width = 180
        '
        'Hesochi
        '
        Me.Hesochi.DataPropertyName = "Hesochi"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Hesochi.DefaultCellStyle = DataGridViewCellStyle2
        Me.Hesochi.HeaderText = "Hệ số chi/thu"
        Me.Hesochi.Name = "Hesochi"
        Me.Hesochi.ReadOnly = True
        Me.Hesochi.Width = 90
        '
        'ThucLinh
        '
        Me.ThucLinh.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.ThucLinh.DataPropertyName = "ThucLinh"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ThucLinh.DefaultCellStyle = DataGridViewCellStyle3
        Me.ThucLinh.HeaderText = "Số tiền"
        Me.ThucLinh.Name = "ThucLinh"
        Me.ThucLinh.ReadOnly = True
        Me.ThucLinh.Width = 66
        '
        'Ghichu
        '
        Me.Ghichu.DataPropertyName = "Ghichu"
        Me.Ghichu.HeaderText = "Ghi chú"
        Me.Ghichu.Name = "Ghichu"
        Me.Ghichu.ReadOnly = True
        Me.Ghichu.Width = 200
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.treeCocau)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(204, 563)
        Me.Panel1.TabIndex = 3
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
        'frmLuongBosungGroup
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
        Me.Name = "frmLuongBosungGroup"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: CHI LƯƠNG BỔ SUNG CHO NHIỀU CÁN BỘ"
        Me.pnlGroup.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel11.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.gridLuongBS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents treeCocau As System.Windows.Forms.TreeView
    Friend WithEvents pnlGroup As System.Windows.Forms.Panel
    Friend WithEvents bntDeleteBS As System.Windows.Forms.Button
    Friend WithEvents Label94 As System.Windows.Forms.Label
    Friend WithEvents txtHesochi As System.Windows.Forms.TextBox
    Friend WithEvents cboLoaiLuongBS As System.Windows.Forms.ComboBox
    Friend WithEvents txtThangBS As System.Windows.Forms.TextBox
    Friend WithEvents gridLuongBS As System.Windows.Forms.DataGridView
    Friend WithEvents bntUpdateBS As System.Windows.Forms.Button
    Friend WithEvents bntNewBS As System.Windows.Forms.Button
    Friend WithEvents Label90 As System.Windows.Forms.Label
    Friend WithEvents txtGhichuBS As System.Windows.Forms.TextBox
    Friend WithEvents txtThuclinhBS As System.Windows.Forms.TextBox
    Friend WithEvents Label93 As System.Windows.Forms.Label
    Friend WithEvents Label95 As System.Windows.Forms.Label
    Friend WithEvents Label99 As System.Windows.Forms.Label
    Friend WithEvents txtNamBS As System.Windows.Forms.TextBox
    Friend WithEvents Label100 As System.Windows.Forms.Label
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents bntCloseBS As System.Windows.Forms.Button
    Friend WithEvents bntCancelBS As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents labListBC As System.Windows.Forms.Label
    Friend WithEvents Id As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cln_cbBS As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents IdCanBo_LBS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Thang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdLoaiLuongBS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdCanbo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Hesochi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ThucLinh As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Ghichu As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cbAllDelete As System.Windows.Forms.CheckBox
End Class
