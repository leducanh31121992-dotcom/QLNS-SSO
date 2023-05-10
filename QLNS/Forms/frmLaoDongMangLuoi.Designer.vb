<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaoDongMangLuoi
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.tabMain = New System.Windows.Forms.TabControl()
        Me.tabKeHoachLD = New System.Windows.Forms.TabPage()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.labNam = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cboPhong = New System.Windows.Forms.ComboBox()
        Me.labPhong = New System.Windows.Forms.Label()
        Me.numThang = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.rdThang = New System.Windows.Forms.RadioButton()
        Me.rdNam = New System.Windows.Forms.RadioButton()
        Me.dtpkNam_LD = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.gridKHLD = New System.Windows.Forms.DataGridView()
        Me.idKHLD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nam = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Thang = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Id_DonVi = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Id_Phong = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DaiHan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NganHan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.labAlert = New System.Windows.Forms.Label()
        Me.btnNew_LD = New System.Windows.Forms.Button()
        Me.btnUpdate_LD = New System.Windows.Forms.Button()
        Me.btnCancel_LD = New System.Windows.Forms.Button()
        Me.btnDelete_LD = New System.Windows.Forms.Button()
        Me.btnClose_LD = New System.Windows.Forms.Button()
        Me.txtGhiChu = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.txtSoNganHan = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtSoDaiHan = New System.Windows.Forms.TextBox()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.tabMangLuoiDV = New System.Windows.Forms.TabPage()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtXaPhuong = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.numQuy = New System.Windows.Forms.NumericUpDown()
        Me.dtpkNam_ML = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.gridMLDV = New System.Windows.Forms.DataGridView()
        Me.IdML = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Quy = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdPGD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SoXaPhuong = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SoDiemGD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DuNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnNew_ML = New System.Windows.Forms.Button()
        Me.btnUpdate_ML = New System.Windows.Forms.Button()
        Me.btnCancel_ML = New System.Windows.Forms.Button()
        Me.btnDelete_ML = New System.Windows.Forms.Button()
        Me.btnClose_ML = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtDuNo = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtSoDiemGD = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cboDV = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.tabMain.SuspendLayout()
        Me.tabKeHoachLD.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.numThang, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridKHLD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel10.SuspendLayout()
        Me.tabMangLuoiDV.SuspendLayout()
        CType(Me.numQuy, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridMLDV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.tabKeHoachLD)
        Me.tabMain.Controls.Add(Me.tabMangLuoiDV)
        Me.tabMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabMain.ItemSize = New System.Drawing.Size(120, 22)
        Me.tabMain.Location = New System.Drawing.Point(0, 0)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SelectedIndex = 0
        Me.tabMain.Size = New System.Drawing.Size(787, 563)
        Me.tabMain.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.tabMain.TabIndex = 1
        '
        'tabKeHoachLD
        '
        Me.tabKeHoachLD.Controls.Add(Me.Panel2)
        Me.tabKeHoachLD.Controls.Add(Me.cboPhong)
        Me.tabKeHoachLD.Controls.Add(Me.labPhong)
        Me.tabKeHoachLD.Controls.Add(Me.numThang)
        Me.tabKeHoachLD.Controls.Add(Me.Label5)
        Me.tabKeHoachLD.Controls.Add(Me.rdThang)
        Me.tabKeHoachLD.Controls.Add(Me.rdNam)
        Me.tabKeHoachLD.Controls.Add(Me.dtpkNam_LD)
        Me.tabKeHoachLD.Controls.Add(Me.Label3)
        Me.tabKeHoachLD.Controls.Add(Me.gridKHLD)
        Me.tabKeHoachLD.Controls.Add(Me.Panel10)
        Me.tabKeHoachLD.Controls.Add(Me.txtGhiChu)
        Me.tabKeHoachLD.Controls.Add(Me.Label39)
        Me.tabKeHoachLD.Controls.Add(Me.txtSoNganHan)
        Me.tabKeHoachLD.Controls.Add(Me.Label42)
        Me.tabKeHoachLD.Controls.Add(Me.txtSoDaiHan)
        Me.tabKeHoachLD.Controls.Add(Me.Label44)
        Me.tabKeHoachLD.Controls.Add(Me.cboDonVi)
        Me.tabKeHoachLD.Controls.Add(Me.Label35)
        Me.tabKeHoachLD.Controls.Add(Me.Panel3)
        Me.tabKeHoachLD.Location = New System.Drawing.Point(4, 26)
        Me.tabKeHoachLD.Name = "tabKeHoachLD"
        Me.tabKeHoachLD.Padding = New System.Windows.Forms.Padding(3)
        Me.tabKeHoachLD.Size = New System.Drawing.Size(779, 533)
        Me.tabKeHoachLD.TabIndex = 0
        Me.tabKeHoachLD.Text = "Kế hoạch lao động"
        Me.tabKeHoachLD.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.labNam)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Location = New System.Drawing.Point(-1, 378)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(777, 30)
        Me.Panel2.TabIndex = 156
        '
        'labNam
        '
        Me.labNam.AutoSize = True
        Me.labNam.Dock = System.Windows.Forms.DockStyle.Left
        Me.labNam.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labNam.Location = New System.Drawing.Point(313, 0)
        Me.labNam.Name = "labNam"
        Me.labNam.Size = New System.Drawing.Size(0, 14)
        Me.labNam.TabIndex = 126
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(0, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(313, 14)
        Me.Label13.TabIndex = 124
        Me.Label13.Text = "Số kế hoạch lao động trong năm của toàn đơn vị:"
        '
        'cboPhong
        '
        Me.cboPhong.BackColor = System.Drawing.SystemColors.Window
        Me.cboPhong.DisplayMember = "Display"
        Me.cboPhong.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPhong.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhong.FormattingEnabled = True
        Me.cboPhong.Location = New System.Drawing.Point(484, 433)
        Me.cboPhong.Name = "cboPhong"
        Me.cboPhong.Size = New System.Drawing.Size(299, 22)
        Me.cboPhong.TabIndex = 143
        Me.cboPhong.ValueMember = "Value"
        '
        'labPhong
        '
        Me.labPhong.AutoSize = True
        Me.labPhong.Location = New System.Drawing.Point(392, 437)
        Me.labPhong.Name = "labPhong"
        Me.labPhong.Size = New System.Drawing.Size(86, 14)
        Me.labPhong.TabIndex = 155
        Me.labPhong.Text = "ĐV trực thuộc"
        '
        'numThang
        '
        Me.numThang.Location = New System.Drawing.Point(426, 410)
        Me.numThang.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.numThang.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numThang.Name = "numThang"
        Me.numThang.Size = New System.Drawing.Size(39, 22)
        Me.numThang.TabIndex = 140
        Me.numThang.Value = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numThang.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(190, 414)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(99, 14)
        Me.Label5.TabIndex = 154
        Me.Label5.Text = "Số kế hoạch cho"
        '
        'rdThang
        '
        Me.rdThang.AutoSize = True
        Me.rdThang.Location = New System.Drawing.Point(353, 413)
        Me.rdThang.Name = "rdThang"
        Me.rdThang.Size = New System.Drawing.Size(74, 18)
        Me.rdThang.TabIndex = 139
        Me.rdThang.Text = "từ tháng"
        Me.rdThang.UseVisualStyleBackColor = True
        '
        'rdNam
        '
        Me.rdNam.AutoSize = True
        Me.rdNam.Checked = True
        Me.rdNam.Location = New System.Drawing.Point(289, 413)
        Me.rdNam.Name = "rdNam"
        Me.rdNam.Size = New System.Drawing.Size(64, 18)
        Me.rdNam.TabIndex = 138
        Me.rdNam.TabStop = True
        Me.rdNam.Text = "cả năm"
        Me.rdNam.UseVisualStyleBackColor = True
        '
        'dtpkNam_LD
        '
        Me.dtpkNam_LD.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpkNam_LD.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpkNam_LD.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpkNam_LD.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpkNam_LD.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpkNam_LD.CustomFormat = "yyyy"
        Me.dtpkNam_LD.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpkNam_LD.Location = New System.Drawing.Point(88, 410)
        Me.dtpkNam_LD.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam_LD.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam_LD.Name = "dtpkNam_LD"
        Me.dtpkNam_LD.ShowUpDown = True
        Me.dtpkNam_LD.Size = New System.Drawing.Size(96, 22)
        Me.dtpkNam_LD.TabIndex = 153
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(53, 414)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 14)
        Me.Label3.TabIndex = 152
        Me.Label3.Text = "Năm"
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
        Me.gridKHLD.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.idKHLD, Me.Nam, Me.Thang, Me.Id_DonVi, Me.Id_Phong, Me.DaiHan, Me.NganHan})
        Me.gridKHLD.Dock = System.Windows.Forms.DockStyle.Top
        Me.gridKHLD.Location = New System.Drawing.Point(3, 4)
        Me.gridKHLD.MultiSelect = False
        Me.gridKHLD.Name = "gridKHLD"
        Me.gridKHLD.ReadOnly = True
        Me.gridKHLD.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridKHLD.RowHeadersVisible = False
        Me.gridKHLD.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridKHLD.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridKHLD.Size = New System.Drawing.Size(773, 367)
        Me.gridKHLD.TabIndex = 151
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
        Me.Id_DonVi.DataPropertyName = "IdDonVi_KH"
        Me.Id_DonVi.HeaderText = "Đơn vị"
        Me.Id_DonVi.Name = "Id_DonVi"
        Me.Id_DonVi.ReadOnly = True
        Me.Id_DonVi.Width = 350
        '
        'Id_Phong
        '
        Me.Id_Phong.DataPropertyName = "IdPhong"
        Me.Id_Phong.HeaderText = "Phòng/Ban"
        Me.Id_Phong.Name = "Id_Phong"
        Me.Id_Phong.ReadOnly = True
        Me.Id_Phong.Width = 180
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
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.Color.DarkGray
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel10.Controls.Add(Me.labAlert)
        Me.Panel10.Controls.Add(Me.btnNew_LD)
        Me.Panel10.Controls.Add(Me.btnUpdate_LD)
        Me.Panel10.Controls.Add(Me.btnCancel_LD)
        Me.Panel10.Controls.Add(Me.btnDelete_LD)
        Me.Panel10.Controls.Add(Me.btnClose_LD)
        Me.Panel10.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel10.Location = New System.Drawing.Point(3, 504)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(773, 26)
        Me.Panel10.TabIndex = 142
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
        'btnNew_LD
        '
        Me.btnNew_LD.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnNew_LD.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew_LD.Location = New System.Drawing.Point(452, 0)
        Me.btnNew_LD.Name = "btnNew_LD"
        Me.btnNew_LD.Size = New System.Drawing.Size(73, 22)
        Me.btnNew_LD.TabIndex = 4
        Me.btnNew_LD.Text = "&Thêm mới"
        Me.btnNew_LD.UseVisualStyleBackColor = True
        '
        'btnUpdate_LD
        '
        Me.btnUpdate_LD.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnUpdate_LD.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate_LD.Location = New System.Drawing.Point(525, 0)
        Me.btnUpdate_LD.Name = "btnUpdate_LD"
        Me.btnUpdate_LD.Size = New System.Drawing.Size(61, 22)
        Me.btnUpdate_LD.TabIndex = 1
        Me.btnUpdate_LD.Text = "&Ghi"
        Me.btnUpdate_LD.UseVisualStyleBackColor = True
        '
        'btnCancel_LD
        '
        Me.btnCancel_LD.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCancel_LD.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel_LD.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel_LD.Location = New System.Drawing.Point(586, 0)
        Me.btnCancel_LD.Name = "btnCancel_LD"
        Me.btnCancel_LD.Size = New System.Drawing.Size(61, 22)
        Me.btnCancel_LD.TabIndex = 2
        Me.btnCancel_LD.Text = "&Bỏ qua"
        Me.btnCancel_LD.UseVisualStyleBackColor = True
        '
        'btnDelete_LD
        '
        Me.btnDelete_LD.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnDelete_LD.Enabled = False
        Me.btnDelete_LD.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete_LD.Location = New System.Drawing.Point(647, 0)
        Me.btnDelete_LD.Name = "btnDelete_LD"
        Me.btnDelete_LD.Size = New System.Drawing.Size(61, 22)
        Me.btnDelete_LD.TabIndex = 3
        Me.btnDelete_LD.Text = "&Xoá"
        Me.btnDelete_LD.UseVisualStyleBackColor = True
        '
        'btnClose_LD
        '
        Me.btnClose_LD.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnClose_LD.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose_LD.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose_LD.Location = New System.Drawing.Point(708, 0)
        Me.btnClose_LD.Name = "btnClose_LD"
        Me.btnClose_LD.Size = New System.Drawing.Size(61, 22)
        Me.btnClose_LD.TabIndex = 5
        Me.btnClose_LD.Text = "&Quay ra"
        Me.btnClose_LD.UseVisualStyleBackColor = True
        '
        'txtGhiChu
        '
        Me.txtGhiChu.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhiChu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhiChu.Location = New System.Drawing.Point(88, 479)
        Me.txtGhiChu.Name = "txtGhiChu"
        Me.txtGhiChu.Size = New System.Drawing.Size(691, 22)
        Me.txtGhiChu.TabIndex = 146
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(36, 482)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(48, 14)
        Me.Label39.TabIndex = 150
        Me.Label39.Text = "Ghi chú"
        '
        'txtSoNganHan
        '
        Me.txtSoNganHan.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoNganHan.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoNganHan.Location = New System.Drawing.Point(484, 456)
        Me.txtSoNganHan.Name = "txtSoNganHan"
        Me.txtSoNganHan.Size = New System.Drawing.Size(299, 22)
        Me.txtSoNganHan.TabIndex = 145
        Me.txtSoNganHan.Text = "0"
        Me.txtSoNganHan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(2, 459)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(82, 14)
        Me.Label42.TabIndex = 149
        Me.Label42.Text = "Số LĐ dài hạn"
        '
        'txtSoDaiHan
        '
        Me.txtSoDaiHan.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoDaiHan.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoDaiHan.Location = New System.Drawing.Point(88, 456)
        Me.txtSoDaiHan.Name = "txtSoDaiHan"
        Me.txtSoDaiHan.Size = New System.Drawing.Size(210, 22)
        Me.txtSoDaiHan.TabIndex = 144
        Me.txtSoDaiHan.Text = "0"
        Me.txtSoDaiHan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(351, 460)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(127, 14)
        Me.Label44.TabIndex = 148
        Me.Label44.Text = "Số lao động ngắn hạn"
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(88, 433)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(297, 22)
        Me.cboDonVi.TabIndex = 141
        Me.cboDonVi.ValueMember = "Value"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(43, 437)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(41, 14)
        Me.Label35.TabIndex = 147
        Me.Label35.Text = "Đơn vị"
        '
        'Panel3
        '
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(3, 3)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(773, 1)
        Me.Panel3.TabIndex = 0
        '
        'tabMangLuoiDV
        '
        Me.tabMangLuoiDV.Controls.Add(Me.Label2)
        Me.tabMangLuoiDV.Controls.Add(Me.txtXaPhuong)
        Me.tabMangLuoiDV.Controls.Add(Me.Label4)
        Me.tabMangLuoiDV.Controls.Add(Me.Label6)
        Me.tabMangLuoiDV.Controls.Add(Me.numQuy)
        Me.tabMangLuoiDV.Controls.Add(Me.dtpkNam_ML)
        Me.tabMangLuoiDV.Controls.Add(Me.Label7)
        Me.tabMangLuoiDV.Controls.Add(Me.gridMLDV)
        Me.tabMangLuoiDV.Controls.Add(Me.Panel1)
        Me.tabMangLuoiDV.Controls.Add(Me.TextBox1)
        Me.tabMangLuoiDV.Controls.Add(Me.Label8)
        Me.tabMangLuoiDV.Controls.Add(Me.txtDuNo)
        Me.tabMangLuoiDV.Controls.Add(Me.Label9)
        Me.tabMangLuoiDV.Controls.Add(Me.txtSoDiemGD)
        Me.tabMangLuoiDV.Controls.Add(Me.Label10)
        Me.tabMangLuoiDV.Controls.Add(Me.cboDV)
        Me.tabMangLuoiDV.Controls.Add(Me.Label11)
        Me.tabMangLuoiDV.Controls.Add(Me.Panel4)
        Me.tabMangLuoiDV.Location = New System.Drawing.Point(4, 26)
        Me.tabMangLuoiDV.Name = "tabMangLuoiDV"
        Me.tabMangLuoiDV.Padding = New System.Windows.Forms.Padding(3)
        Me.tabMangLuoiDV.Size = New System.Drawing.Size(779, 533)
        Me.tabMangLuoiDV.TabIndex = 1
        Me.tabMangLuoiDV.Text = "Mạng lưới đơn vị"
        Me.tabMangLuoiDV.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(-4, 459)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 14)
        Me.Label2.TabIndex = 151
        Me.Label2.Text = "Số xã, phường"
        '
        'txtXaPhuong
        '
        Me.txtXaPhuong.BackColor = System.Drawing.SystemColors.Window
        Me.txtXaPhuong.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtXaPhuong.Location = New System.Drawing.Point(88, 456)
        Me.txtXaPhuong.Name = "txtXaPhuong"
        Me.txtXaPhuong.Size = New System.Drawing.Size(69, 22)
        Me.txtXaPhuong.TabIndex = 138
        Me.txtXaPhuong.Text = "0"
        Me.txtXaPhuong.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(651, 459)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(129, 14)
        Me.Label4.TabIndex = 150
        Me.Label4.Text = "(đơn vị tính: Tỷ đồng)"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(320, 414)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 14)
        Me.Label6.TabIndex = 149
        Me.Label6.Text = "Quý"
        '
        'numQuy
        '
        Me.numQuy.Location = New System.Drawing.Point(350, 410)
        Me.numQuy.Maximum = New Decimal(New Integer() {4, 0, 0, 0})
        Me.numQuy.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numQuy.Name = "numQuy"
        Me.numQuy.Size = New System.Drawing.Size(69, 22)
        Me.numQuy.TabIndex = 136
        Me.numQuy.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'dtpkNam_ML
        '
        Me.dtpkNam_ML.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpkNam_ML.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpkNam_ML.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpkNam_ML.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpkNam_ML.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpkNam_ML.CustomFormat = "yyyy"
        Me.dtpkNam_ML.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpkNam_ML.Location = New System.Drawing.Point(88, 410)
        Me.dtpkNam_ML.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam_ML.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam_ML.Name = "dtpkNam_ML"
        Me.dtpkNam_ML.ShowUpDown = True
        Me.dtpkNam_ML.Size = New System.Drawing.Size(69, 22)
        Me.dtpkNam_ML.TabIndex = 135
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(53, 414)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(31, 14)
        Me.Label7.TabIndex = 148
        Me.Label7.Text = "Năm"
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
        Me.gridMLDV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdML, Me.DataGridViewTextBoxColumn1, Me.Quy, Me.IdPGD, Me.SoXaPhuong, Me.SoDiemGD, Me.DuNo})
        Me.gridMLDV.Dock = System.Windows.Forms.DockStyle.Top
        Me.gridMLDV.Location = New System.Drawing.Point(3, 11)
        Me.gridMLDV.MultiSelect = False
        Me.gridMLDV.Name = "gridMLDV"
        Me.gridMLDV.ReadOnly = True
        Me.gridMLDV.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridMLDV.RowHeadersVisible = False
        Me.gridMLDV.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridMLDV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridMLDV.Size = New System.Drawing.Size(773, 393)
        Me.gridMLDV.TabIndex = 147
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
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "Nam"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewTextBoxColumn1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewTextBoxColumn1.FillWeight = 50.0!
        Me.DataGridViewTextBoxColumn1.HeaderText = "Năm"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 56
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
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Format = "N0"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.SoDiemGD.DefaultCellStyle = DataGridViewCellStyle3
        Me.SoDiemGD.HeaderText = "Số điểm giao dịch"
        Me.SoDiemGD.Name = "SoDiemGD"
        Me.SoDiemGD.ReadOnly = True
        '
        'DuNo
        '
        Me.DuNo.DataPropertyName = "DuNo"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.Format = "N0"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.DuNo.DefaultCellStyle = DataGridViewCellStyle4
        Me.DuNo.HeaderText = "Dư nợ"
        Me.DuNo.Name = "DuNo"
        Me.DuNo.ReadOnly = True
        Me.DuNo.Width = 180
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGray
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnNew_ML)
        Me.Panel1.Controls.Add(Me.btnUpdate_ML)
        Me.Panel1.Controls.Add(Me.btnCancel_ML)
        Me.Panel1.Controls.Add(Me.btnDelete_ML)
        Me.Panel1.Controls.Add(Me.btnClose_ML)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(3, 504)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(773, 26)
        Me.Panel1.TabIndex = 141
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Maroon
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(444, 22)
        Me.Label1.TabIndex = 7
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnNew_ML
        '
        Me.btnNew_ML.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnNew_ML.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew_ML.Location = New System.Drawing.Point(452, 0)
        Me.btnNew_ML.Name = "btnNew_ML"
        Me.btnNew_ML.Size = New System.Drawing.Size(73, 22)
        Me.btnNew_ML.TabIndex = 4
        Me.btnNew_ML.Text = "&Thêm mới"
        Me.btnNew_ML.UseVisualStyleBackColor = True
        '
        'btnUpdate_ML
        '
        Me.btnUpdate_ML.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnUpdate_ML.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate_ML.Location = New System.Drawing.Point(525, 0)
        Me.btnUpdate_ML.Name = "btnUpdate_ML"
        Me.btnUpdate_ML.Size = New System.Drawing.Size(61, 22)
        Me.btnUpdate_ML.TabIndex = 1
        Me.btnUpdate_ML.Text = "&Ghi"
        Me.btnUpdate_ML.UseVisualStyleBackColor = True
        '
        'btnCancel_ML
        '
        Me.btnCancel_ML.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCancel_ML.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel_ML.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel_ML.Location = New System.Drawing.Point(586, 0)
        Me.btnCancel_ML.Name = "btnCancel_ML"
        Me.btnCancel_ML.Size = New System.Drawing.Size(61, 22)
        Me.btnCancel_ML.TabIndex = 2
        Me.btnCancel_ML.Text = "&Bỏ qua"
        Me.btnCancel_ML.UseVisualStyleBackColor = True
        '
        'btnDelete_ML
        '
        Me.btnDelete_ML.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnDelete_ML.Enabled = False
        Me.btnDelete_ML.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete_ML.Location = New System.Drawing.Point(647, 0)
        Me.btnDelete_ML.Name = "btnDelete_ML"
        Me.btnDelete_ML.Size = New System.Drawing.Size(61, 22)
        Me.btnDelete_ML.TabIndex = 3
        Me.btnDelete_ML.Text = "&Xoá"
        Me.btnDelete_ML.UseVisualStyleBackColor = True
        '
        'btnClose_ML
        '
        Me.btnClose_ML.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnClose_ML.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose_ML.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose_ML.Location = New System.Drawing.Point(708, 0)
        Me.btnClose_ML.Name = "btnClose_ML"
        Me.btnClose_ML.Size = New System.Drawing.Size(61, 22)
        Me.btnClose_ML.TabIndex = 5
        Me.btnClose_ML.Text = "&Quay ra"
        Me.btnClose_ML.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(88, 479)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(691, 22)
        Me.TextBox1.TabIndex = 142
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(36, 482)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 14)
        Me.Label8.TabIndex = 146
        Me.Label8.Text = "Ghi chú"
        '
        'txtDuNo
        '
        Me.txtDuNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDuNo.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDuNo.Location = New System.Drawing.Point(481, 456)
        Me.txtDuNo.Name = "txtDuNo"
        Me.txtDuNo.Size = New System.Drawing.Size(170, 22)
        Me.txtDuNo.TabIndex = 140
        Me.txtDuNo.Text = "0"
        Me.txtDuNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(246, 459)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(103, 14)
        Me.Label9.TabIndex = 145
        Me.Label9.Text = "Số điểm giao dịch"
        '
        'txtSoDiemGD
        '
        Me.txtSoDiemGD.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoDiemGD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoDiemGD.Location = New System.Drawing.Point(350, 456)
        Me.txtSoDiemGD.Name = "txtSoDiemGD"
        Me.txtSoDiemGD.Size = New System.Drawing.Size(69, 22)
        Me.txtSoDiemGD.TabIndex = 139
        Me.txtSoDiemGD.Text = "0"
        Me.txtSoDiemGD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(437, 459)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(41, 14)
        Me.Label10.TabIndex = 144
        Me.Label10.Text = "Dư nợ"
        '
        'cboDV
        '
        Me.cboDV.BackColor = System.Drawing.SystemColors.Window
        Me.cboDV.DisplayMember = "Display"
        Me.cboDV.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDV.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDV.FormattingEnabled = True
        Me.cboDV.Location = New System.Drawing.Point(88, 433)
        Me.cboDV.Name = "cboDV"
        Me.cboDV.Size = New System.Drawing.Size(563, 22)
        Me.cboDV.TabIndex = 137
        Me.cboDV.ValueMember = "Value"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(43, 437)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(41, 14)
        Me.Label11.TabIndex = 143
        Me.Label11.Text = "Đơn vị"
        '
        'Panel4
        '
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(3, 3)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(773, 8)
        Me.Panel4.TabIndex = 1
        '
        'frmLaoDongMangLuoi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.tabMain)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmLaoDongMangLuoi"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: HỒ SƠ LAO ĐỘNG - MẠNG LƯỚI HÀNG NĂM CỦA ĐƠN VỊ"
        Me.tabMain.ResumeLayout(False)
        Me.tabKeHoachLD.ResumeLayout(False)
        Me.tabKeHoachLD.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.numThang, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridKHLD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel10.ResumeLayout(False)
        Me.tabMangLuoiDV.ResumeLayout(False)
        Me.tabMangLuoiDV.PerformLayout()
        CType(Me.numQuy, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridMLDV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabMain As System.Windows.Forms.TabControl
    Friend WithEvents tabKeHoachLD As System.Windows.Forms.TabPage
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents tabMangLuoiDV As System.Windows.Forms.TabPage
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents cboPhong As System.Windows.Forms.ComboBox
    Friend WithEvents labPhong As System.Windows.Forms.Label
    Friend WithEvents numThang As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rdThang As System.Windows.Forms.RadioButton
    Friend WithEvents rdNam As System.Windows.Forms.RadioButton
    Friend WithEvents dtpkNam_LD As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents gridKHLD As System.Windows.Forms.DataGridView
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents labAlert As System.Windows.Forms.Label
    Friend WithEvents btnNew_LD As System.Windows.Forms.Button
    Friend WithEvents btnUpdate_LD As System.Windows.Forms.Button
    Friend WithEvents btnCancel_LD As System.Windows.Forms.Button
    Friend WithEvents btnDelete_LD As System.Windows.Forms.Button
    Friend WithEvents btnClose_LD As System.Windows.Forms.Button
    Friend WithEvents txtGhiChu As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents txtSoNganHan As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtSoDaiHan As System.Windows.Forms.TextBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtXaPhuong As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents numQuy As System.Windows.Forms.NumericUpDown
    Friend WithEvents dtpkNam_ML As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents gridMLDV As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnNew_ML As System.Windows.Forms.Button
    Friend WithEvents btnUpdate_ML As System.Windows.Forms.Button
    Friend WithEvents btnCancel_ML As System.Windows.Forms.Button
    Friend WithEvents btnDelete_ML As System.Windows.Forms.Button
    Friend WithEvents btnClose_ML As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtDuNo As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtSoDiemGD As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cboDV As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents labNam As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents idKHLD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nam As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Thang As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Id_DonVi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Id_Phong As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DaiHan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NganHan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdML As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Quy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IdPGD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SoXaPhuong As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SoDiemGD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DuNo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
