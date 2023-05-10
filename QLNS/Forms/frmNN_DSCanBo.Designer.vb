<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNN_DSCanBo
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.dgv = New System.Windows.Forms.DataGridView()
        Me.STT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.HoTen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NgaySinh = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PhongBan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TamThoi = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.BaoCao = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.rbtChuyenNganh = New System.Windows.Forms.RadioButton()
        Me.rbtNganHan = New System.Windows.Forms.RadioButton()
        Me.btn_Luu = New System.Windows.Forms.Button()
        Me.btn_back = New System.Windows.Forms.Button()
        Me.chkAll_TT = New System.Windows.Forms.CheckBox()
        Me.chkAll_BC = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbtToanBo = New System.Windows.Forms.RadioButton()
        Me.rbtTheoDonVi = New System.Windows.Forms.RadioButton()
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboDonVi
        '
        Me.cboDonVi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(166, 15)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(313, 21)
        Me.cboDonVi.TabIndex = 2
        '
        'dgv
        '
        Me.dgv.AllowUserToAddRows = False
        Me.dgv.AllowUserToDeleteRows = False
        Me.dgv.AllowUserToResizeRows = False
        Me.dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgv.BackgroundColor = System.Drawing.Color.White
        Me.dgv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgv.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.STT, Me.ID, Me.HoTen, Me.NgaySinh, Me.PhongBan, Me.TamThoi, Me.BaoCao})
        Me.dgv.Location = New System.Drawing.Point(12, 59)
        Me.dgv.Name = "dgv"
        Me.dgv.RowHeadersVisible = False
        Me.dgv.Size = New System.Drawing.Size(671, 405)
        Me.dgv.TabIndex = 2
        '
        'STT
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.STT.DefaultCellStyle = DataGridViewCellStyle1
        Me.STT.HeaderText = "STT"
        Me.STT.Name = "STT"
        Me.STT.ReadOnly = True
        Me.STT.Width = 53
        '
        'ID
        '
        Me.ID.HeaderText = "ID"
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Width = 43
        '
        'HoTen
        '
        Me.HoTen.HeaderText = "Họ tên"
        Me.HoTen.Name = "HoTen"
        Me.HoTen.ReadOnly = True
        Me.HoTen.Width = 64
        '
        'NgaySinh
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.Format = "dd/MM/yyyy"
        Me.NgaySinh.DefaultCellStyle = DataGridViewCellStyle2
        Me.NgaySinh.HeaderText = "Ngày sinh"
        Me.NgaySinh.Name = "NgaySinh"
        Me.NgaySinh.ReadOnly = True
        Me.NgaySinh.Width = 79
        '
        'PhongBan
        '
        Me.PhongBan.HeaderText = "Phòng/Ban/PGD"
        Me.PhongBan.Name = "PhongBan"
        Me.PhongBan.ReadOnly = True
        Me.PhongBan.Width = 115
        '
        'TamThoi
        '
        Me.TamThoi.HeaderText = "Tạm thời"
        Me.TamThoi.Name = "TamThoi"
        Me.TamThoi.Width = 54
        '
        'BaoCao
        '
        Me.BaoCao.HeaderText = "Báo cáo"
        Me.BaoCao.Name = "BaoCao"
        Me.BaoCao.Width = 53
        '
        'rbtChuyenNganh
        '
        Me.rbtChuyenNganh.AutoSize = True
        Me.rbtChuyenNganh.Checked = True
        Me.rbtChuyenNganh.Location = New System.Drawing.Point(8, 16)
        Me.rbtChuyenNganh.Name = "rbtChuyenNganh"
        Me.rbtChuyenNganh.Size = New System.Drawing.Size(94, 17)
        Me.rbtChuyenNganh.TabIndex = 0
        Me.rbtChuyenNganh.TabStop = True
        Me.rbtChuyenNganh.Text = "Chuyển ngành"
        Me.rbtChuyenNganh.UseVisualStyleBackColor = True
        '
        'rbtNganHan
        '
        Me.rbtNganHan.AutoSize = True
        Me.rbtNganHan.Location = New System.Drawing.Point(108, 16)
        Me.rbtNganHan.Name = "rbtNganHan"
        Me.rbtNganHan.Size = New System.Drawing.Size(72, 17)
        Me.rbtNganHan.TabIndex = 1
        Me.rbtNganHan.Text = "Ngắn hạn"
        Me.rbtNganHan.UseVisualStyleBackColor = True
        '
        'btn_Luu
        '
        Me.btn_Luu.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Luu.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Luu.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btn_Luu.Location = New System.Drawing.Point(542, 477)
        Me.btn_Luu.Name = "btn_Luu"
        Me.btn_Luu.Size = New System.Drawing.Size(77, 23)
        Me.btn_Luu.TabIndex = 5
        Me.btn_Luu.Text = "&Lưu"
        Me.btn_Luu.UseVisualStyleBackColor = True
        '
        'btn_back
        '
        Me.btn_back.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_back.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_back.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btn_back.Location = New System.Drawing.Point(622, 477)
        Me.btn_back.Name = "btn_back"
        Me.btn_back.Size = New System.Drawing.Size(61, 23)
        Me.btn_back.TabIndex = 6
        Me.btn_back.Text = "&Quay ra"
        Me.btn_back.UseVisualStyleBackColor = True
        '
        'chkAll_TT
        '
        Me.chkAll_TT.AutoSize = True
        Me.chkAll_TT.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAll_TT.Location = New System.Drawing.Point(15, 477)
        Me.chkAll_TT.Name = "chkAll_TT"
        Me.chkAll_TT.Size = New System.Drawing.Size(152, 17)
        Me.chkAll_TT.TabIndex = 3
        Me.chkAll_TT.Text = "Chọn toàn bộ cột Tạm thời"
        Me.chkAll_TT.UseVisualStyleBackColor = True
        '
        'chkAll_BC
        '
        Me.chkAll_BC.AutoSize = True
        Me.chkAll_BC.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAll_BC.Location = New System.Drawing.Point(173, 477)
        Me.chkAll_BC.Name = "chkAll_BC"
        Me.chkAll_BC.Size = New System.Drawing.Size(151, 17)
        Me.chkAll_BC.TabIndex = 4
        Me.chkAll_BC.Text = "Chọn toàn bộ cột Báo cáo"
        Me.chkAll_BC.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbtChuyenNganh)
        Me.GroupBox1.Controls.Add(Me.rbtNganHan)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(182, 42)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbtToanBo)
        Me.GroupBox2.Controls.Add(Me.rbtTheoDonVi)
        Me.GroupBox2.Controls.Add(Me.cboDonVi)
        Me.GroupBox2.Location = New System.Drawing.Point(198, 7)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(485, 42)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'rbtToanBo
        '
        Me.rbtToanBo.AutoSize = True
        Me.rbtToanBo.Location = New System.Drawing.Point(7, 16)
        Me.rbtToanBo.Name = "rbtToanBo"
        Me.rbtToanBo.Size = New System.Drawing.Size(65, 17)
        Me.rbtToanBo.TabIndex = 0
        Me.rbtToanBo.Text = "Toàn bộ"
        Me.rbtToanBo.UseVisualStyleBackColor = True
        '
        'rbtTheoDonVi
        '
        Me.rbtTheoDonVi.AutoSize = True
        Me.rbtTheoDonVi.Checked = True
        Me.rbtTheoDonVi.Location = New System.Drawing.Point(78, 16)
        Me.rbtTheoDonVi.Name = "rbtTheoDonVi"
        Me.rbtTheoDonVi.Size = New System.Drawing.Size(86, 17)
        Me.rbtTheoDonVi.TabIndex = 1
        Me.rbtTheoDonVi.TabStop = True
        Me.rbtTheoDonVi.Text = "Theo đơn vị:"
        Me.rbtTheoDonVi.UseVisualStyleBackColor = True
        '
        'frmNN_DSCanBo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(695, 512)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.chkAll_BC)
        Me.Controls.Add(Me.chkAll_TT)
        Me.Controls.Add(Me.btn_Luu)
        Me.Controls.Add(Me.btn_back)
        Me.Controls.Add(Me.dgv)
        Me.Name = "frmNN_DSCanBo"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Danh sách cán bộ báo cáo"
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents dgv As System.Windows.Forms.DataGridView
    Friend WithEvents rbtChuyenNganh As System.Windows.Forms.RadioButton
    Friend WithEvents rbtNganHan As System.Windows.Forms.RadioButton
    Friend WithEvents btn_Luu As System.Windows.Forms.Button
    Friend WithEvents btn_back As System.Windows.Forms.Button
    Friend WithEvents chkAll_TT As System.Windows.Forms.CheckBox
    Friend WithEvents chkAll_BC As System.Windows.Forms.CheckBox
    Friend WithEvents STT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HoTen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NgaySinh As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PhongBan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TamThoi As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents BaoCao As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbtToanBo As System.Windows.Forms.RadioButton
    Friend WithEvents rbtTheoDonVi As System.Windows.Forms.RadioButton
End Class
