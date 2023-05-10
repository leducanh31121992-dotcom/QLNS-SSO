<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTDKTGroup
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTDKTGroup))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.treeCocau = New System.Windows.Forms.TreeView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lstKhenThuong = New System.Windows.Forms.ListBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rdDeNghi = New System.Windows.Forms.RadioButton()
        Me.rdKhenThuong = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdDX = New System.Windows.Forms.RadioButton()
        Me.rdDK = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpkNam = New System.Windows.Forms.DateTimePicker()
        Me.lnkChoiseCB = New System.Windows.Forms.LinkLabel()
        Me.gridKT = New System.Windows.Forms.DataGridView()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.labTitle = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.labAlert = New System.Windows.Forms.Label()
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
        Me.txtNoiDung_KT = New System.Windows.Forms.TextBox()
        Me.dpkNgayQD_KT = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtNguoiQD_KT = New System.Windows.Forms.TextBox()
        Me.labCapKT = New System.Windows.Forms.Label()
        Me.cboChucVuKyQD = New System.Windows.Forms.ComboBox()
        Me.cboCap_KT = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
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
        Me.Panel1.TabIndex = 5
        '
        'treeCocau
        '
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
        Me.Label1.Text = "DANH SÁCH QĐ KHEN THƯỞNG"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lstKhenThuong)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.dtpkNam)
        Me.Panel2.Controls.Add(Me.lnkChoiseCB)
        Me.Panel2.Controls.Add(Me.gridKT)
        Me.Panel2.Controls.Add(Me.Panel4)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.txtGhichu_KT)
        Me.Panel2.Controls.Add(Me.txtSoQD_KT)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.txtNoiDung_KT)
        Me.Panel2.Controls.Add(Me.dpkNgayQD_KT)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.txtNguoiQD_KT)
        Me.Panel2.Controls.Add(Me.labCapKT)
        Me.Panel2.Controls.Add(Me.cboChucVuKyQD)
        Me.Panel2.Controls.Add(Me.cboCap_KT)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(204, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(583, 563)
        Me.Panel2.TabIndex = 125
        '
        'lstKhenThuong
        '
        Me.lstKhenThuong.FormattingEnabled = True
        Me.lstKhenThuong.HorizontalScrollbar = True
        Me.lstKhenThuong.ItemHeight = 14
        Me.lstKhenThuong.Location = New System.Drawing.Point(114, 426)
        Me.lstKhenThuong.Name = "lstKhenThuong"
        Me.lstKhenThuong.Size = New System.Drawing.Size(466, 88)
        Me.lstKhenThuong.TabIndex = 144
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(237, 320)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(137, 18)
        Me.Label5.TabIndex = 143
        Me.Label5.Text = "Thời gian khen thưởng:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(1, 429)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 37)
        Me.Label2.TabIndex = 141
        Me.Label2.Text = "Danh sách           cá nhân/tập thể"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rdDeNghi)
        Me.GroupBox2.Controls.Add(Me.rdKhenThuong)
        Me.GroupBox2.Location = New System.Drawing.Point(119, 289)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox2.Size = New System.Drawing.Size(259, 25)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'rdDeNghi
        '
        Me.rdDeNghi.AutoSize = True
        Me.rdDeNghi.Location = New System.Drawing.Point(108, 7)
        Me.rdDeNghi.Name = "rdDeNghi"
        Me.rdDeNghi.Size = New System.Drawing.Size(143, 18)
        Me.rdDeNghi.TabIndex = 2
        Me.rdDeNghi.Text = "Đề nghị khen thưởng"
        Me.rdDeNghi.UseVisualStyleBackColor = True
        '
        'rdKhenThuong
        '
        Me.rdKhenThuong.AutoSize = True
        Me.rdKhenThuong.Checked = True
        Me.rdKhenThuong.Location = New System.Drawing.Point(10, 7)
        Me.rdKhenThuong.Name = "rdKhenThuong"
        Me.rdKhenThuong.Size = New System.Drawing.Size(98, 18)
        Me.rdKhenThuong.TabIndex = 1
        Me.rdKhenThuong.TabStop = True
        Me.rdKhenThuong.Text = "Khen thưởng"
        Me.rdKhenThuong.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdDX)
        Me.GroupBox1.Controls.Add(Me.rdDK)
        Me.GroupBox1.Location = New System.Drawing.Point(369, 313)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Size = New System.Drawing.Size(148, 25)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'rdDX
        '
        Me.rdDX.AutoSize = True
        Me.rdDX.Location = New System.Drawing.Point(71, 7)
        Me.rdDX.Name = "rdDX"
        Me.rdDX.Size = New System.Drawing.Size(73, 18)
        Me.rdDX.TabIndex = 2
        Me.rdDX.Text = "Đột xuất"
        Me.rdDX.UseVisualStyleBackColor = True
        '
        'rdDK
        '
        Me.rdDK.AutoSize = True
        Me.rdDK.Checked = True
        Me.rdDK.Location = New System.Drawing.Point(5, 7)
        Me.rdDK.Name = "rdDK"
        Me.rdDK.Size = New System.Drawing.Size(65, 18)
        Me.rdDK.TabIndex = 1
        Me.rdDK.TabStop = True
        Me.rdDK.Text = "Định kỳ"
        Me.rdDK.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1, 320)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 14)
        Me.Label3.TabIndex = 133
        Me.Label3.Text = "Năm khen thưởng"
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
        Me.dtpkNam.Location = New System.Drawing.Point(114, 316)
        Me.dtpkNam.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam.Name = "dtpkNam"
        Me.dtpkNam.ShowUpDown = True
        Me.dtpkNam.Size = New System.Drawing.Size(118, 22)
        Me.dtpkNam.TabIndex = 5
        '
        'lnkChoiseCB
        '
        Me.lnkChoiseCB.Location = New System.Drawing.Point(1, 461)
        Me.lnkChoiseCB.Margin = New System.Windows.Forms.Padding(0)
        Me.lnkChoiseCB.Name = "lnkChoiseCB"
        Me.lnkChoiseCB.Size = New System.Drawing.Size(103, 18)
        Me.lnkChoiseCB.TabIndex = 2
        Me.lnkChoiseCB.TabStop = True
        Me.lnkChoiseCB.Text = "Chọn danh sách"
        Me.lnkChoiseCB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        Me.gridKT.Location = New System.Drawing.Point(1, 21)
        Me.gridKT.MultiSelect = False
        Me.gridKT.Name = "gridKT"
        Me.gridKT.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.gridKT.RowHeadersVisible = False
        Me.gridKT.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.gridKT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridKT.Size = New System.Drawing.Size(579, 268)
        Me.gridKT.TabIndex = 126
        Me.gridKT.Tag = ""
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.DarkGray
        Me.Panel4.Controls.Add(Me.labTitle)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(583, 21)
        Me.Panel4.TabIndex = 125
        '
        'labTitle
        '
        Me.labTitle.BackColor = System.Drawing.Color.DarkGray
        Me.labTitle.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labTitle.ForeColor = System.Drawing.Color.Maroon
        Me.labTitle.Location = New System.Drawing.Point(0, 0)
        Me.labTitle.Name = "labTitle"
        Me.labTitle.Size = New System.Drawing.Size(581, 21)
        Me.labTitle.TabIndex = 106
        Me.labTitle.Text = "Danh sách Khen thưởng"
        Me.labTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.labAlert)
        Me.Panel3.Controls.Add(Me.bntNewKT)
        Me.Panel3.Controls.Add(Me.bntUpdateKT)
        Me.Panel3.Controls.Add(Me.bntCancelKT)
        Me.Panel3.Controls.Add(Me.bntDeleteKT)
        Me.Panel3.Controls.Add(Me.bntCloseKT)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 537)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(583, 26)
        Me.Panel3.TabIndex = 15
        '
        'labAlert
        '
        Me.labAlert.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labAlert.ForeColor = System.Drawing.Color.Blue
        Me.labAlert.Location = New System.Drawing.Point(-2, 5)
        Me.labAlert.Name = "labAlert"
        Me.labAlert.Size = New System.Drawing.Size(258, 20)
        Me.labAlert.TabIndex = 6
        '
        'bntNewKT
        '
        Me.bntNewKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntNewKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntNewKT.Location = New System.Drawing.Point(262, 0)
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
        Me.bntUpdateKT.Location = New System.Drawing.Point(335, 0)
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
        Me.bntCancelKT.Location = New System.Drawing.Point(396, 0)
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
        Me.bntDeleteKT.Location = New System.Drawing.Point(457, 0)
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
        Me.bntCloseKT.Location = New System.Drawing.Point(518, 0)
        Me.bntCloseKT.Name = "bntCloseKT"
        Me.bntCloseKT.Size = New System.Drawing.Size(61, 22)
        Me.bntCloseKT.TabIndex = 5
        Me.bntCloseKT.Text = "&Quay ra"
        Me.bntCloseKT.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(1, 364)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(21, 14)
        Me.Label10.TabIndex = 112
        Me.Label10.Text = "Số"
        '
        'txtGhichu_KT
        '
        Me.txtGhichu_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtGhichu_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGhichu_KT.Location = New System.Drawing.Point(114, 514)
        Me.txtGhichu_KT.Name = "txtGhichu_KT"
        Me.txtGhichu_KT.Size = New System.Drawing.Size(466, 22)
        Me.txtGhichu_KT.TabIndex = 14
        '
        'txtSoQD_KT
        '
        Me.txtSoQD_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtSoQD_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoQD_KT.Location = New System.Drawing.Point(114, 360)
        Me.txtSoQD_KT.Name = "txtSoQD_KT"
        Me.txtSoQD_KT.Size = New System.Drawing.Size(118, 22)
        Me.txtSoQD_KT.TabIndex = 8
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(1, 517)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 14)
        Me.Label14.TabIndex = 123
        Me.Label14.Text = "Ghi chú"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(237, 364)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(34, 14)
        Me.Label9.TabIndex = 115
        Me.Label9.Text = "Ngày"
        '
        'txtNoiDung_KT
        '
        Me.txtNoiDung_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoiDung_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNoiDung_KT.Location = New System.Drawing.Point(114, 404)
        Me.txtNoiDung_KT.Name = "txtNoiDung_KT"
        Me.txtNoiDung_KT.Size = New System.Drawing.Size(466, 22)
        Me.txtNoiDung_KT.TabIndex = 12
        '
        'dpkNgayQD_KT
        '
        Me.dpkNgayQD_KT.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayQD_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayQD_KT.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayQD_KT.Location = New System.Drawing.Point(271, 360)
        Me.dpkNgayQD_KT.Name = "dpkNgayQD_KT"
        Me.dpkNgayQD_KT.Size = New System.Drawing.Size(89, 22)
        Me.dpkNgayQD_KT.TabIndex = 9
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(1, 407)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 14)
        Me.Label13.TabIndex = 122
        Me.Label13.Text = "Nội dung"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(377, 385)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(55, 14)
        Me.Label8.TabIndex = 118
        Me.Label8.Text = "Người ký"
        '
        'txtNguoiQD_KT
        '
        Me.txtNguoiQD_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtNguoiQD_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNguoiQD_KT.Location = New System.Drawing.Point(432, 382)
        Me.txtNguoiQD_KT.Name = "txtNguoiQD_KT"
        Me.txtNguoiQD_KT.Size = New System.Drawing.Size(148, 22)
        Me.txtNguoiQD_KT.TabIndex = 11
        '
        'labCapKT
        '
        Me.labCapKT.AutoSize = True
        Me.labCapKT.Location = New System.Drawing.Point(1, 341)
        Me.labCapKT.Name = "labCapKT"
        Me.labCapKT.Size = New System.Drawing.Size(103, 14)
        Me.labCapKT.TabIndex = 119
        Me.labCapKT.Text = "Cấp khen thưởng"
        '
        'cboChucVuKyQD
        '
        Me.cboChucVuKyQD.BackColor = System.Drawing.SystemColors.Window
        Me.cboChucVuKyQD.DisplayMember = "Display"
        Me.cboChucVuKyQD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChucVuKyQD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboChucVuKyQD.FormattingEnabled = True
        Me.cboChucVuKyQD.Location = New System.Drawing.Point(114, 382)
        Me.cboChucVuKyQD.Name = "cboChucVuKyQD"
        Me.cboChucVuKyQD.Size = New System.Drawing.Size(246, 22)
        Me.cboChucVuKyQD.TabIndex = 10
        Me.cboChucVuKyQD.ValueMember = "Value"
        '
        'cboCap_KT
        '
        Me.cboCap_KT.BackColor = System.Drawing.SystemColors.Window
        Me.cboCap_KT.DisplayMember = "Display"
        Me.cboCap_KT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCap_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCap_KT.FormattingEnabled = True
        Me.cboCap_KT.Location = New System.Drawing.Point(114, 338)
        Me.cboCap_KT.Name = "cboCap_KT"
        Me.cboCap_KT.Size = New System.Drawing.Size(466, 22)
        Me.cboCap_KT.TabIndex = 7
        Me.cboCap_KT.ValueMember = "Value"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(1, 385)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(103, 14)
        Me.Label11.TabIndex = 120
        Me.Label11.Text = "Chức vụ người ký"
        '
        'frmTDKTGroup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmTDKTGroup"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: THI ĐUA KHEN THƯỞNG CHO CÁ NHÂN/TẬP THỂ"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.gridKT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents treeCocau As System.Windows.Forms.TreeView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents gridKT As System.Windows.Forms.DataGridView
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents labTitle As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents bntNewKT As System.Windows.Forms.Button
    Friend WithEvents bntUpdateKT As System.Windows.Forms.Button
    Friend WithEvents bntCancelKT As System.Windows.Forms.Button
    Friend WithEvents bntDeleteKT As System.Windows.Forms.Button
    Friend WithEvents bntCloseKT As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtGhichu_KT As System.Windows.Forms.TextBox
    Friend WithEvents txtSoQD_KT As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtNoiDung_KT As System.Windows.Forms.TextBox
    Friend WithEvents dpkNgayQD_KT As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtNguoiQD_KT As System.Windows.Forms.TextBox
    Friend WithEvents labCapKT As System.Windows.Forms.Label
    Friend WithEvents cboChucVuKyQD As System.Windows.Forms.ComboBox
    Friend WithEvents cboCap_KT As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents rdDeNghi As System.Windows.Forms.RadioButton
    Friend WithEvents rdKhenThuong As System.Windows.Forms.RadioButton
    Friend WithEvents dtpkNam As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lnkChoiseCB As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdDX As System.Windows.Forms.RadioButton
    Friend WithEvents rdDK As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lstKhenThuong As System.Windows.Forms.ListBox
    Friend WithEvents labAlert As System.Windows.Forms.Label
End Class
