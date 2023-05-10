<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLuongDonVi
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
        Me.labTuNgay = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.labHeSoNganh = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.labLuongCoBan = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.labAlert = New System.Windows.Forms.Label()
        Me.bntNew_CN = New System.Windows.Forms.Button()
        Me.bntUpdate_CN = New System.Windows.Forms.Button()
        Me.btnCancel_CN = New System.Windows.Forms.Button()
        Me.bntDelete_CN = New System.Windows.Forms.Button()
        Me.bntClose_CN = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.gridLuongDonVi = New System.Windows.Forms.DataGridView()
        Me.idLDV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.idChiNhanh = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.idTienLuong_CN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TuNgay = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtGhiChu_CN = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.dpkNgayApDung_CN = New System.Windows.Forms.DateTimePicker()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.cboCVNguoiQD_CN = New System.Windows.Forms.ComboBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.txtNguoiQD_CN = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.dpkNgayQD_CN = New System.Windows.Forms.DateTimePicker()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.txtSoQD_CN = New System.Windows.Forms.TextBox()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.cboLuongCoBan = New System.Windows.Forms.ComboBox()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.grpDonVi = New System.Windows.Forms.GroupBox()
        Me.Panel7.SuspendLayout()
        Me.Panel10.SuspendLayout()
        CType(Me.gridLuongDonVi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.grpDonVi.SuspendLayout()
        Me.SuspendLayout()
        '
        'labTuNgay
        '
        Me.labTuNgay.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labTuNgay.ForeColor = System.Drawing.SystemColors.Highlight
        Me.labTuNgay.Location = New System.Drawing.Point(200, 64)
        Me.labTuNgay.Name = "labTuNgay"
        Me.labTuNgay.Size = New System.Drawing.Size(576, 19)
        Me.labTuNgay.TabIndex = 21
        Me.labTuNgay.Text = "áp dung tu"
        Me.labTuNgay.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label5.Location = New System.Drawing.Point(83, 65)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(117, 18)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Áp dụng từ ngày"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'labHeSoNganh
        '
        Me.labHeSoNganh.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labHeSoNganh.ForeColor = System.Drawing.SystemColors.Highlight
        Me.labHeSoNganh.Location = New System.Drawing.Point(200, 39)
        Me.labHeSoNganh.Name = "labHeSoNganh"
        Me.labHeSoNganh.Size = New System.Drawing.Size(573, 24)
        Me.labHeSoNganh.TabIndex = 19
        Me.labHeSoNganh.Text = "He so luong"
        Me.labHeSoNganh.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label3.Location = New System.Drawing.Point(10, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(190, 24)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Hệ số ngành"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'labLuongCoBan
        '
        Me.labLuongCoBan.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labLuongCoBan.ForeColor = System.Drawing.SystemColors.Highlight
        Me.labLuongCoBan.Location = New System.Drawing.Point(200, 16)
        Me.labLuongCoBan.Name = "labLuongCoBan"
        Me.labLuongCoBan.Size = New System.Drawing.Size(573, 24)
        Me.labLuongCoBan.TabIndex = 17
        Me.labLuongCoBan.Text = "Lương cơ bản"
        Me.labLuongCoBan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label4.Location = New System.Drawing.Point(13, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(187, 24)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Lương tối thiểu"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.SystemColors.Control
        Me.Panel7.Controls.Add(Me.Panel10)
        Me.Panel7.Controls.Add(Me.Panel2)
        Me.Panel7.Controls.Add(Me.Label38)
        Me.Panel7.Controls.Add(Me.gridLuongDonVi)
        Me.Panel7.Controls.Add(Me.txtGhiChu_CN)
        Me.Panel7.Controls.Add(Me.Label39)
        Me.Panel7.Controls.Add(Me.dpkNgayApDung_CN)
        Me.Panel7.Controls.Add(Me.Label40)
        Me.Panel7.Controls.Add(Me.cboCVNguoiQD_CN)
        Me.Panel7.Controls.Add(Me.Label41)
        Me.Panel7.Controls.Add(Me.txtNguoiQD_CN)
        Me.Panel7.Controls.Add(Me.Label42)
        Me.Panel7.Controls.Add(Me.dpkNgayQD_CN)
        Me.Panel7.Controls.Add(Me.Label43)
        Me.Panel7.Controls.Add(Me.txtSoQD_CN)
        Me.Panel7.Controls.Add(Me.Label44)
        Me.Panel7.Controls.Add(Me.cboLuongCoBan)
        Me.Panel7.Controls.Add(Me.Label37)
        Me.Panel7.Controls.Add(Me.cboDonVi)
        Me.Panel7.Controls.Add(Me.Label35)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel7.Location = New System.Drawing.Point(0, 91)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(787, 472)
        Me.Panel7.TabIndex = 22
        '
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.Color.DarkGray
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel10.Controls.Add(Me.labAlert)
        Me.Panel10.Controls.Add(Me.bntNew_CN)
        Me.Panel10.Controls.Add(Me.bntUpdate_CN)
        Me.Panel10.Controls.Add(Me.btnCancel_CN)
        Me.Panel10.Controls.Add(Me.bntDelete_CN)
        Me.Panel10.Controls.Add(Me.bntClose_CN)
        Me.Panel10.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel10.Location = New System.Drawing.Point(0, 446)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(787, 26)
        Me.Panel10.TabIndex = 9
        '
        'labAlert
        '
        Me.labAlert.Dock = System.Windows.Forms.DockStyle.Left
        Me.labAlert.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labAlert.ForeColor = System.Drawing.Color.Maroon
        Me.labAlert.Location = New System.Drawing.Point(0, 0)
        Me.labAlert.Name = "labAlert"
        Me.labAlert.Size = New System.Drawing.Size(427, 22)
        Me.labAlert.TabIndex = 7
        Me.labAlert.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'bntNew_CN
        '
        Me.bntNew_CN.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntNew_CN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntNew_CN.Location = New System.Drawing.Point(466, 0)
        Me.bntNew_CN.Name = "bntNew_CN"
        Me.bntNew_CN.Size = New System.Drawing.Size(73, 22)
        Me.bntNew_CN.TabIndex = 4
        Me.bntNew_CN.Text = "&Thêm mới"
        Me.bntNew_CN.UseVisualStyleBackColor = True
        '
        'bntUpdate_CN
        '
        Me.bntUpdate_CN.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntUpdate_CN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntUpdate_CN.Location = New System.Drawing.Point(539, 0)
        Me.bntUpdate_CN.Name = "bntUpdate_CN"
        Me.bntUpdate_CN.Size = New System.Drawing.Size(61, 22)
        Me.bntUpdate_CN.TabIndex = 1
        Me.bntUpdate_CN.Text = "&Ghi"
        Me.bntUpdate_CN.UseVisualStyleBackColor = True
        '
        'btnCancel_CN
        '
        Me.btnCancel_CN.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCancel_CN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel_CN.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel_CN.Location = New System.Drawing.Point(600, 0)
        Me.btnCancel_CN.Name = "btnCancel_CN"
        Me.btnCancel_CN.Size = New System.Drawing.Size(61, 22)
        Me.btnCancel_CN.TabIndex = 2
        Me.btnCancel_CN.Text = "&Bỏ qua"
        Me.btnCancel_CN.UseVisualStyleBackColor = True
        '
        'bntDelete_CN
        '
        Me.bntDelete_CN.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntDelete_CN.Enabled = False
        Me.bntDelete_CN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntDelete_CN.Location = New System.Drawing.Point(661, 0)
        Me.bntDelete_CN.Name = "bntDelete_CN"
        Me.bntDelete_CN.Size = New System.Drawing.Size(61, 22)
        Me.bntDelete_CN.TabIndex = 3
        Me.bntDelete_CN.Text = "&Xoá"
        Me.bntDelete_CN.UseVisualStyleBackColor = True
        '
        'bntClose_CN
        '
        Me.bntClose_CN.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose_CN.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose_CN.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntClose_CN.Location = New System.Drawing.Point(722, 0)
        Me.bntClose_CN.Name = "bntClose_CN"
        Me.bntClose_CN.Size = New System.Drawing.Size(61, 22)
        Me.bntClose_CN.TabIndex = 5
        Me.bntClose_CN.Text = "&Quay ra"
        Me.bntClose_CN.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(787, 1)
        Me.Panel2.TabIndex = 114
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.DarkGray
        Me.Label38.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.Color.Maroon
        Me.Label38.Location = New System.Drawing.Point(0, 0)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(791, 21)
        Me.Label38.TabIndex = 113
        Me.Label38.Text = "Danh sách các thông số lương của đơn vị"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'gridLuongDonVi
        '
        Me.gridLuongDonVi.AllowUserToAddRows = False
        Me.gridLuongDonVi.AllowUserToDeleteRows = False
        Me.gridLuongDonVi.AllowUserToResizeColumns = False
        Me.gridLuongDonVi.AllowUserToResizeRows = False
        Me.gridLuongDonVi.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridLuongDonVi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridLuongDonVi.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridLuongDonVi.ColumnHeadersHeight = 37
        Me.gridLuongDonVi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.idLDV, Me.idChiNhanh, Me.idTienLuong_CN, Me.TuNgay})
        Me.gridLuongDonVi.Location = New System.Drawing.Point(0, 21)
        Me.gridLuongDonVi.MultiSelect = False
        Me.gridLuongDonVi.Name = "gridLuongDonVi"
        Me.gridLuongDonVi.ReadOnly = True
        Me.gridLuongDonVi.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridLuongDonVi.RowHeadersVisible = False
        Me.gridLuongDonVi.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridLuongDonVi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridLuongDonVi.Size = New System.Drawing.Size(789, 269)
        Me.gridLuongDonVi.TabIndex = 112
        Me.gridLuongDonVi.Tag = ""
        '
        'idLDV
        '
        Me.idLDV.DataPropertyName = "id"
        Me.idLDV.HeaderText = "IdLDV"
        Me.idLDV.Name = "idLDV"
        Me.idLDV.ReadOnly = True
        Me.idLDV.Visible = False
        '
        'idChiNhanh
        '
        Me.idChiNhanh.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.idChiNhanh.DataPropertyName = "idChiNhanh"
        Me.idChiNhanh.HeaderText = "Đơn vị"
        Me.idChiNhanh.Name = "idChiNhanh"
        Me.idChiNhanh.ReadOnly = True
        Me.idChiNhanh.Width = 54
        '
        'idTienLuong_CN
        '
        Me.idTienLuong_CN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.idTienLuong_CN.DataPropertyName = "idTienLuong"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.idTienLuong_CN.DefaultCellStyle = DataGridViewCellStyle1
        Me.idTienLuong_CN.HeaderText = "Thanh toán lương"
        Me.idTienLuong_CN.Name = "idTienLuong_CN"
        Me.idTienLuong_CN.ReadOnly = True
        '
        'TuNgay
        '
        Me.TuNgay.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.TuNgay.DataPropertyName = "TuNgay"
        Me.TuNgay.HeaderText = "Ngày áp dụng"
        Me.TuNgay.Name = "TuNgay"
        Me.TuNgay.ReadOnly = True
        Me.TuNgay.Width = 99
        '
        'txtGhiChu_CN
        '
        Me.txtGhiChu_CN.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhiChu_CN.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhiChu_CN.Location = New System.Drawing.Point(109, 401)
        Me.txtGhiChu_CN.Multiline = True
        Me.txtGhiChu_CN.Name = "txtGhiChu_CN"
        Me.txtGhiChu_CN.Size = New System.Drawing.Size(597, 44)
        Me.txtGhiChu_CN.TabIndex = 8
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(59, 404)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(48, 14)
        Me.Label39.TabIndex = 111
        Me.Label39.Text = "Ghi chú"
        '
        'dpkNgayApDung_CN
        '
        Me.dpkNgayApDung_CN.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayApDung_CN.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayApDung_CN.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayApDung_CN.Location = New System.Drawing.Point(109, 335)
        Me.dpkNgayApDung_CN.Name = "dpkNgayApDung_CN"
        Me.dpkNgayApDung_CN.Size = New System.Drawing.Size(110, 22)
        Me.dpkNgayApDung_CN.TabIndex = 3
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(6, 338)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(101, 14)
        Me.Label40.TabIndex = 110
        Me.Label40.Text = "Áp dụng từ ngày"
        '
        'cboCVNguoiQD_CN
        '
        Me.cboCVNguoiQD_CN.BackColor = System.Drawing.SystemColors.Window
        Me.cboCVNguoiQD_CN.DisplayMember = "Display"
        Me.cboCVNguoiQD_CN.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCVNguoiQD_CN.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCVNguoiQD_CN.FormattingEnabled = True
        Me.cboCVNguoiQD_CN.Location = New System.Drawing.Point(431, 379)
        Me.cboCVNguoiQD_CN.Name = "cboCVNguoiQD_CN"
        Me.cboCVNguoiQD_CN.Size = New System.Drawing.Size(275, 22)
        Me.cboCVNguoiQD_CN.TabIndex = 7
        Me.cboCVNguoiQD_CN.ValueMember = "Value"
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(377, 382)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(52, 14)
        Me.Label41.TabIndex = 109
        Me.Label41.Text = "Chức vụ"
        '
        'txtNguoiQD_CN
        '
        Me.txtNguoiQD_CN.BackColor = System.Drawing.SystemColors.Window
        Me.txtNguoiQD_CN.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNguoiQD_CN.Location = New System.Drawing.Point(109, 379)
        Me.txtNguoiQD_CN.Name = "txtNguoiQD_CN"
        Me.txtNguoiQD_CN.Size = New System.Drawing.Size(221, 22)
        Me.txtNguoiQD_CN.TabIndex = 6
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(47, 382)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(60, 14)
        Me.Label42.TabIndex = 108
        Me.Label42.Text = "Người QĐ"
        '
        'dpkNgayQD_CN
        '
        Me.dpkNgayQD_CN.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayQD_CN.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayQD_CN.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayQD_CN.Location = New System.Drawing.Point(431, 357)
        Me.dpkNgayQD_CN.Name = "dpkNgayQD_CN"
        Me.dpkNgayQD_CN.Size = New System.Drawing.Size(116, 22)
        Me.dpkNgayQD_CN.TabIndex = 5
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(332, 361)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(97, 14)
        Me.Label43.TabIndex = 105
        Me.Label43.Text = "Ngày quyết định"
        '
        'txtSoQD_CN
        '
        Me.txtSoQD_CN.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoQD_CN.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoQD_CN.Location = New System.Drawing.Point(109, 357)
        Me.txtSoQD_CN.Name = "txtSoQD_CN"
        Me.txtSoQD_CN.Size = New System.Drawing.Size(221, 22)
        Me.txtSoQD_CN.TabIndex = 4
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(23, 361)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(84, 14)
        Me.Label44.TabIndex = 101
        Me.Label44.Text = "Số quyết định"
        '
        'cboLuongCoBan
        '
        Me.cboLuongCoBan.BackColor = System.Drawing.SystemColors.Window
        Me.cboLuongCoBan.DisplayMember = "ThongTinChung"
        Me.cboLuongCoBan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLuongCoBan.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLuongCoBan.FormattingEnabled = True
        Me.cboLuongCoBan.Location = New System.Drawing.Point(109, 313)
        Me.cboLuongCoBan.Name = "cboLuongCoBan"
        Me.cboLuongCoBan.Size = New System.Drawing.Size(438, 22)
        Me.cboLuongCoBan.TabIndex = 2
        Me.cboLuongCoBan.ValueMember = "IdTienLuong"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.Location = New System.Drawing.Point(1, 316)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(106, 14)
        Me.Label37.TabIndex = 97
        Me.Label37.Text = "Thanh toán lương"
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(109, 291)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(438, 22)
        Me.cboDonVi.TabIndex = 1
        Me.cboDonVi.ValueMember = "Value"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(66, 295)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(41, 14)
        Me.Label35.TabIndex = 92
        Me.Label35.Text = "Đơn vị"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.grpDonVi)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(787, 91)
        Me.Panel1.TabIndex = 23
        '
        'grpDonVi
        '
        Me.grpDonVi.Controls.Add(Me.Label3)
        Me.grpDonVi.Controls.Add(Me.labHeSoNganh)
        Me.grpDonVi.Controls.Add(Me.labLuongCoBan)
        Me.grpDonVi.Controls.Add(Me.labTuNgay)
        Me.grpDonVi.Controls.Add(Me.Label4)
        Me.grpDonVi.Controls.Add(Me.Label5)
        Me.grpDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpDonVi.Location = New System.Drawing.Point(2, 3)
        Me.grpDonVi.Name = "grpDonVi"
        Me.grpDonVi.Size = New System.Drawing.Size(782, 86)
        Me.grpDonVi.TabIndex = 22
        Me.grpDonVi.TabStop = False
        Me.grpDonVi.Text = "Thông số lương hiện hành của đơn vị"
        '
        'frmLuongDonVi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.Panel7)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmLuongDonVi"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: THANH TOÁN LƯƠNG ĐƠN VỊ"
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout()
        Me.Panel10.ResumeLayout(False)
        CType(Me.gridLuongDonVi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.grpDonVi.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents labTuNgay As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents labHeSoNganh As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents labLuongCoBan As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents bntDelete_CN As System.Windows.Forms.Button
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents gridLuongDonVi As System.Windows.Forms.DataGridView
    Friend WithEvents bntUpdate_CN As System.Windows.Forms.Button
    Friend WithEvents bntNew_CN As System.Windows.Forms.Button
    Friend WithEvents txtGhiChu_CN As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents dpkNgayApDung_CN As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents cboCVNguoiQD_CN As System.Windows.Forms.ComboBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents txtNguoiQD_CN As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents dpkNgayQD_CN As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents txtSoQD_CN As System.Windows.Forms.TextBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents cboLuongCoBan As System.Windows.Forms.ComboBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents btnCancel_CN As System.Windows.Forms.Button
    Friend WithEvents bntClose_CN As System.Windows.Forms.Button
    Friend WithEvents grpDonVi As System.Windows.Forms.GroupBox
    Friend WithEvents labAlert As System.Windows.Forms.Label
    Friend WithEvents idLDV As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents idChiNhanh As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents idTienLuong_CN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TuNgay As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
