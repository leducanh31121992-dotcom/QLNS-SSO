<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBC_BaoCao
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBC_BaoCao))
        Me.crpv_main = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.gb_main = New System.Windows.Forms.GroupBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.dtpk_ngaylap_bc = New System.Windows.Forms.DateTimePicker()
        Me.lbl_ngaylap = New System.Windows.Forms.Label()
        Me.edt_so_bc = New System.Windows.Forms.TextBox()
        Me.lbl_sobc = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.edt_nguoilap = New System.Windows.Forms.TextBox()
        Me.lbl_nguoilap = New System.Windows.Forms.Label()
        Me.edt_truongphong = New System.Windows.Forms.TextBox()
        Me.lbl_truongphong = New System.Windows.Forms.Label()
        Me.edt_giamdoc = New System.Windows.Forms.TextBox()
        Me.lbl_giamdoc = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.pnl_2 = New System.Windows.Forms.Panel()
        Me.dtpk_nam_bc = New System.Windows.Forms.DateTimePicker()
        Me.lbl_nambc = New System.Windows.Forms.Label()
        Me.numDenThang = New System.Windows.Forms.NumericUpDown()
        Me.labDenThang = New System.Windows.Forms.Label()
        Me.numTuThang = New System.Windows.Forms.NumericUpDown()
        Me.labTuThang = New System.Windows.Forms.Label()
        Me.cb_kybc = New System.Windows.Forms.ComboBox()
        Me.lbl_kybc = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.pnl_01 = New System.Windows.Forms.Panel()
        Me.ckb_ChiTietPGD = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cb_chinhanh = New System.Windows.Forms.ComboBox()
        Me.lbl_chinhanh = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btn_exportexcel = New System.Windows.Forms.Button()
        Me.btn_view_report = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btn_reset = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_close = New System.Windows.Forms.Button()
        Me.rdt_reportBC3 = New QLNS.BC03()
        Me.rdt_reportBC4 = New QLNS.BC04()
        Me.rdt_reportBC1 = New QLNS.BC01()
        Me.rdt_reportBC2 = New QLNS.BC02()
        Me.rpt_BC02QT = New QLNS.BC02_TQ()
        Me.rptBC02_HSC = New QLNS.BC02_HSC()
        Me.rptBC02SGD = New QLNS.BC02SGD()
        Me.rdt_reportBC2TTDT = New QLNS.BC02TTDT()
        Me.rdt_reportBC2CNTT = New QLNS.BC02CNTT()
        Me.gb_main.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.pnl_2.SuspendLayout()
        CType(Me.numDenThang, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numTuThang, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnl_01.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'crpv_main
        '
        Me.crpv_main.ActiveViewIndex = -1
        Me.crpv_main.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crpv_main.Cursor = System.Windows.Forms.Cursors.Default
        Me.crpv_main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crpv_main.Location = New System.Drawing.Point(0, 134)
        Me.crpv_main.Name = "crpv_main"
        Me.crpv_main.SelectionFormula = ""
        Me.crpv_main.Size = New System.Drawing.Size(1006, 519)
        Me.crpv_main.TabIndex = 0
        Me.crpv_main.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        Me.crpv_main.ViewTimeSelectionFormula = ""
        '
        'gb_main
        '
        Me.gb_main.Controls.Add(Me.Panel2)
        Me.gb_main.Controls.Add(Me.Panel6)
        Me.gb_main.Controls.Add(Me.Label7)
        Me.gb_main.Controls.Add(Me.pnl_2)
        Me.gb_main.Controls.Add(Me.Label3)
        Me.gb_main.Controls.Add(Me.pnl_01)
        Me.gb_main.Dock = System.Windows.Forms.DockStyle.Top
        Me.gb_main.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gb_main.Location = New System.Drawing.Point(0, 24)
        Me.gb_main.Name = "gb_main"
        Me.gb_main.Size = New System.Drawing.Size(1006, 110)
        Me.gb_main.TabIndex = 2
        Me.gb_main.TabStop = False
        Me.gb_main.Text = "Khai báo tham số báo cáo"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dtpk_ngaylap_bc)
        Me.Panel2.Controls.Add(Me.lbl_ngaylap)
        Me.Panel2.Controls.Add(Me.edt_so_bc)
        Me.Panel2.Controls.Add(Me.lbl_sobc)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 90)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1000, 22)
        Me.Panel2.TabIndex = 5
        '
        'dtpk_ngaylap_bc
        '
        Me.dtpk_ngaylap_bc.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpk_ngaylap_bc.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpk_ngaylap_bc.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpk_ngaylap_bc.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpk_ngaylap_bc.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpk_ngaylap_bc.CustomFormat = "dd/MM/yyyy"
        Me.dtpk_ngaylap_bc.Dock = System.Windows.Forms.DockStyle.Left
        Me.dtpk_ngaylap_bc.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpk_ngaylap_bc.Location = New System.Drawing.Point(309, 0)
        Me.dtpk_ngaylap_bc.Name = "dtpk_ngaylap_bc"
        Me.dtpk_ngaylap_bc.Size = New System.Drawing.Size(101, 26)
        Me.dtpk_ngaylap_bc.TabIndex = 2
        '
        'lbl_ngaylap
        '
        Me.lbl_ngaylap.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_ngaylap.Location = New System.Drawing.Point(191, 0)
        Me.lbl_ngaylap.Name = "lbl_ngaylap"
        Me.lbl_ngaylap.Size = New System.Drawing.Size(118, 22)
        Me.lbl_ngaylap.TabIndex = 6
        Me.lbl_ngaylap.Text = "Ngày lập biểu "
        Me.lbl_ngaylap.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'edt_so_bc
        '
        Me.edt_so_bc.BackColor = System.Drawing.Color.White
        Me.edt_so_bc.Dock = System.Windows.Forms.DockStyle.Left
        Me.edt_so_bc.ForeColor = System.Drawing.Color.Black
        Me.edt_so_bc.Location = New System.Drawing.Point(112, 0)
        Me.edt_so_bc.Name = "edt_so_bc"
        Me.edt_so_bc.Size = New System.Drawing.Size(79, 26)
        Me.edt_so_bc.TabIndex = 1
        '
        'lbl_sobc
        '
        Me.lbl_sobc.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_sobc.Location = New System.Drawing.Point(0, 0)
        Me.lbl_sobc.Name = "lbl_sobc"
        Me.lbl_sobc.Size = New System.Drawing.Size(112, 22)
        Me.lbl_sobc.TabIndex = 8
        Me.lbl_sobc.Text = " Số báo cáo "
        Me.lbl_sobc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.edt_nguoilap)
        Me.Panel6.Controls.Add(Me.lbl_nguoilap)
        Me.Panel6.Controls.Add(Me.edt_truongphong)
        Me.Panel6.Controls.Add(Me.lbl_truongphong)
        Me.Panel6.Controls.Add(Me.edt_giamdoc)
        Me.Panel6.Controls.Add(Me.lbl_giamdoc)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel6.Location = New System.Drawing.Point(3, 68)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1000, 22)
        Me.Panel6.TabIndex = 4
        '
        'edt_nguoilap
        '
        Me.edt_nguoilap.BackColor = System.Drawing.Color.White
        Me.edt_nguoilap.Dock = System.Windows.Forms.DockStyle.Fill
        Me.edt_nguoilap.ForeColor = System.Drawing.Color.Black
        Me.edt_nguoilap.Location = New System.Drawing.Point(651, 0)
        Me.edt_nguoilap.Name = "edt_nguoilap"
        Me.edt_nguoilap.Size = New System.Drawing.Size(349, 26)
        Me.edt_nguoilap.TabIndex = 5
        '
        'lbl_nguoilap
        '
        Me.lbl_nguoilap.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_nguoilap.Location = New System.Drawing.Point(541, 0)
        Me.lbl_nguoilap.Name = "lbl_nguoilap"
        Me.lbl_nguoilap.Size = New System.Drawing.Size(110, 22)
        Me.lbl_nguoilap.TabIndex = 4
        Me.lbl_nguoilap.Text = "Người lập biểu "
        Me.lbl_nguoilap.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'edt_truongphong
        '
        Me.edt_truongphong.BackColor = System.Drawing.Color.White
        Me.edt_truongphong.Dock = System.Windows.Forms.DockStyle.Left
        Me.edt_truongphong.ForeColor = System.Drawing.Color.Black
        Me.edt_truongphong.Location = New System.Drawing.Point(391, 0)
        Me.edt_truongphong.Name = "edt_truongphong"
        Me.edt_truongphong.Size = New System.Drawing.Size(150, 26)
        Me.edt_truongphong.TabIndex = 3
        '
        'lbl_truongphong
        '
        Me.lbl_truongphong.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_truongphong.Location = New System.Drawing.Point(262, 0)
        Me.lbl_truongphong.Name = "lbl_truongphong"
        Me.lbl_truongphong.Size = New System.Drawing.Size(129, 22)
        Me.lbl_truongphong.TabIndex = 2
        Me.lbl_truongphong.Text = "Trưởng phòng HC-TC "
        Me.lbl_truongphong.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'edt_giamdoc
        '
        Me.edt_giamdoc.BackColor = System.Drawing.Color.White
        Me.edt_giamdoc.Dock = System.Windows.Forms.DockStyle.Left
        Me.edt_giamdoc.ForeColor = System.Drawing.Color.Black
        Me.edt_giamdoc.Location = New System.Drawing.Point(112, 0)
        Me.edt_giamdoc.Name = "edt_giamdoc"
        Me.edt_giamdoc.Size = New System.Drawing.Size(150, 26)
        Me.edt_giamdoc.TabIndex = 1
        '
        'lbl_giamdoc
        '
        Me.lbl_giamdoc.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_giamdoc.Location = New System.Drawing.Point(0, 0)
        Me.lbl_giamdoc.Name = "lbl_giamdoc"
        Me.lbl_giamdoc.Size = New System.Drawing.Size(112, 22)
        Me.lbl_giamdoc.TabIndex = 0
        Me.lbl_giamdoc.Text = "Tổng giám đốc"
        Me.lbl_giamdoc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(3, 67)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(1000, 1)
        Me.Label7.TabIndex = 3
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnl_2
        '
        Me.pnl_2.Controls.Add(Me.dtpk_nam_bc)
        Me.pnl_2.Controls.Add(Me.lbl_nambc)
        Me.pnl_2.Controls.Add(Me.numDenThang)
        Me.pnl_2.Controls.Add(Me.labDenThang)
        Me.pnl_2.Controls.Add(Me.numTuThang)
        Me.pnl_2.Controls.Add(Me.labTuThang)
        Me.pnl_2.Controls.Add(Me.cb_kybc)
        Me.pnl_2.Controls.Add(Me.lbl_kybc)
        Me.pnl_2.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_2.Location = New System.Drawing.Point(3, 45)
        Me.pnl_2.Name = "pnl_2"
        Me.pnl_2.Size = New System.Drawing.Size(1000, 22)
        Me.pnl_2.TabIndex = 2
        '
        'dtpk_nam_bc
        '
        Me.dtpk_nam_bc.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpk_nam_bc.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpk_nam_bc.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpk_nam_bc.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpk_nam_bc.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpk_nam_bc.CustomFormat = "yyyy"
        Me.dtpk_nam_bc.Dock = System.Windows.Forms.DockStyle.Left
        Me.dtpk_nam_bc.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpk_nam_bc.Location = New System.Drawing.Point(722, 0)
        Me.dtpk_nam_bc.Name = "dtpk_nam_bc"
        Me.dtpk_nam_bc.ShowUpDown = True
        Me.dtpk_nam_bc.Size = New System.Drawing.Size(60, 26)
        Me.dtpk_nam_bc.TabIndex = 15
        '
        'lbl_nambc
        '
        Me.lbl_nambc.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_nambc.Location = New System.Drawing.Point(628, 0)
        Me.lbl_nambc.Name = "lbl_nambc"
        Me.lbl_nambc.Size = New System.Drawing.Size(94, 22)
        Me.lbl_nambc.TabIndex = 14
        Me.lbl_nambc.Text = "Năm báo cáo "
        Me.lbl_nambc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'numDenThang
        '
        Me.numDenThang.Dock = System.Windows.Forms.DockStyle.Left
        Me.numDenThang.Location = New System.Drawing.Point(588, 0)
        Me.numDenThang.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.numDenThang.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numDenThang.Name = "numDenThang"
        Me.numDenThang.Size = New System.Drawing.Size(40, 26)
        Me.numDenThang.TabIndex = 3
        Me.numDenThang.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'labDenThang
        '
        Me.labDenThang.Dock = System.Windows.Forms.DockStyle.Left
        Me.labDenThang.Location = New System.Drawing.Point(516, 0)
        Me.labDenThang.Name = "labDenThang"
        Me.labDenThang.Size = New System.Drawing.Size(72, 22)
        Me.labDenThang.TabIndex = 12
        Me.labDenThang.Text = "Đến tháng"
        Me.labDenThang.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'numTuThang
        '
        Me.numTuThang.Dock = System.Windows.Forms.DockStyle.Left
        Me.numTuThang.Location = New System.Drawing.Point(476, 0)
        Me.numTuThang.Maximum = New Decimal(New Integer() {12, 0, 0, 0})
        Me.numTuThang.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numTuThang.Name = "numTuThang"
        Me.numTuThang.Size = New System.Drawing.Size(40, 26)
        Me.numTuThang.TabIndex = 2
        Me.numTuThang.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'labTuThang
        '
        Me.labTuThang.Dock = System.Windows.Forms.DockStyle.Left
        Me.labTuThang.Location = New System.Drawing.Point(410, 0)
        Me.labTuThang.Name = "labTuThang"
        Me.labTuThang.Size = New System.Drawing.Size(66, 22)
        Me.labTuThang.TabIndex = 10
        Me.labTuThang.Text = "Từ tháng"
        Me.labTuThang.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cb_kybc
        '
        Me.cb_kybc.BackColor = System.Drawing.Color.White
        Me.cb_kybc.Dock = System.Windows.Forms.DockStyle.Left
        Me.cb_kybc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_kybc.ForeColor = System.Drawing.Color.Black
        Me.cb_kybc.FormattingEnabled = True
        Me.cb_kybc.Items.AddRange(New Object() {"Báo cáo số liệu tích lũy đến thời điểm hiện tại", "Phát sinh Kỳ 1 (thời điểm từ ngày 01/01 đến ngày 30/06)", "Phát sinh Kỳ 2 (thời điểm từ ngày 01/07 đến ngày 31/12)"})
        Me.cb_kybc.Location = New System.Drawing.Point(112, 0)
        Me.cb_kybc.Name = "cb_kybc"
        Me.cb_kybc.Size = New System.Drawing.Size(298, 26)
        Me.cb_kybc.TabIndex = 1
        '
        'lbl_kybc
        '
        Me.lbl_kybc.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_kybc.Location = New System.Drawing.Point(0, 0)
        Me.lbl_kybc.Name = "lbl_kybc"
        Me.lbl_kybc.Size = New System.Drawing.Size(112, 22)
        Me.lbl_kybc.TabIndex = 0
        Me.lbl_kybc.Text = "Kỳ báo cáo "
        Me.lbl_kybc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(3, 44)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(1000, 1)
        Me.Label3.TabIndex = 1
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnl_01
        '
        Me.pnl_01.Controls.Add(Me.ckb_ChiTietPGD)
        Me.pnl_01.Controls.Add(Me.Label2)
        Me.pnl_01.Controls.Add(Me.cb_chinhanh)
        Me.pnl_01.Controls.Add(Me.lbl_chinhanh)
        Me.pnl_01.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_01.Location = New System.Drawing.Point(3, 22)
        Me.pnl_01.Name = "pnl_01"
        Me.pnl_01.Size = New System.Drawing.Size(1000, 22)
        Me.pnl_01.TabIndex = 0
        '
        'ckb_ChiTietPGD
        '
        Me.ckb_ChiTietPGD.Dock = System.Windows.Forms.DockStyle.Left
        Me.ckb_ChiTietPGD.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ckb_ChiTietPGD.ForeColor = System.Drawing.Color.Maroon
        Me.ckb_ChiTietPGD.Location = New System.Drawing.Point(476, 0)
        Me.ckb_ChiTietPGD.Name = "ckb_ChiTietPGD"
        Me.ckb_ChiTietPGD.Size = New System.Drawing.Size(215, 22)
        Me.ckb_ChiTietPGD.TabIndex = 3
        Me.ckb_ChiTietPGD.Text = "Chi tiết tới Phòng giao dịch"
        Me.ckb_ChiTietPGD.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label2.Location = New System.Drawing.Point(410, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 22)
        Me.Label2.TabIndex = 2
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cb_chinhanh
        '
        Me.cb_chinhanh.BackColor = System.Drawing.Color.White
        Me.cb_chinhanh.Dock = System.Windows.Forms.DockStyle.Left
        Me.cb_chinhanh.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_chinhanh.ForeColor = System.Drawing.Color.Black
        Me.cb_chinhanh.FormattingEnabled = True
        Me.cb_chinhanh.Location = New System.Drawing.Point(112, 0)
        Me.cb_chinhanh.Name = "cb_chinhanh"
        Me.cb_chinhanh.Size = New System.Drawing.Size(298, 26)
        Me.cb_chinhanh.TabIndex = 1
        '
        'lbl_chinhanh
        '
        Me.lbl_chinhanh.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_chinhanh.Location = New System.Drawing.Point(0, 0)
        Me.lbl_chinhanh.Name = "lbl_chinhanh"
        Me.lbl_chinhanh.Size = New System.Drawing.Size(112, 22)
        Me.lbl_chinhanh.TabIndex = 0
        Me.lbl_chinhanh.Text = "Chi nhánh "
        Me.lbl_chinhanh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGray
        Me.Panel1.Controls.Add(Me.btn_exportexcel)
        Me.Panel1.Controls.Add(Me.btn_view_report)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.btn_reset)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btn_close)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1006, 24)
        Me.Panel1.TabIndex = 3
        '
        'btn_exportexcel
        '
        Me.btn_exportexcel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_exportexcel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_exportexcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_exportexcel.Location = New System.Drawing.Point(602, 0)
        Me.btn_exportexcel.Name = "btn_exportexcel"
        Me.btn_exportexcel.Size = New System.Drawing.Size(115, 24)
        Me.btn_exportexcel.TabIndex = 19
        Me.btn_exportexcel.Text = "Xuất file &excel"
        Me.btn_exportexcel.UseVisualStyleBackColor = True
        Me.btn_exportexcel.Visible = False
        '
        'btn_view_report
        '
        Me.btn_view_report.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_view_report.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_view_report.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_view_report.Location = New System.Drawing.Point(717, 0)
        Me.btn_view_report.Name = "btn_view_report"
        Me.btn_view_report.Size = New System.Drawing.Size(95, 24)
        Me.btn_view_report.TabIndex = 14
        Me.btn_view_report.Text = "&Xem báo cáo"
        Me.btn_view_report.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label5.Location = New System.Drawing.Point(812, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(2, 24)
        Me.Label5.TabIndex = 15
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_reset
        '
        Me.btn_reset.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_reset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_reset.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_reset.Location = New System.Drawing.Point(814, 0)
        Me.btn_reset.Name = "btn_reset"
        Me.btn_reset.Size = New System.Drawing.Size(95, 24)
        Me.btn_reset.TabIndex = 16
        Me.btn_reset.Text = "&Bỏ qua"
        Me.btn_reset.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label1.Location = New System.Drawing.Point(909, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(2, 24)
        Me.Label1.TabIndex = 17
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_close
        '
        Me.btn_close.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_close.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_close.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_close.Location = New System.Drawing.Point(911, 0)
        Me.btn_close.Name = "btn_close"
        Me.btn_close.Size = New System.Drawing.Size(95, 24)
        Me.btn_close.TabIndex = 18
        Me.btn_close.Text = "&Quay ra"
        Me.btn_close.UseVisualStyleBackColor = True
        '
        'frmBC_BaoCao
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1006, 653)
        Me.Controls.Add(Me.crpv_main)
        Me.Controls.Add(Me.gb_main)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "frmBC_BaoCao"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự"
        Me.gb_main.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.pnl_2.ResumeLayout(False)
        CType(Me.numDenThang, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numTuThang, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnl_01.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gb_main As System.Windows.Forms.GroupBox
    Friend WithEvents pnl_01 As System.Windows.Forms.Panel
    Friend WithEvents cb_chinhanh As System.Windows.Forms.ComboBox
    Friend WithEvents lbl_chinhanh As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btn_view_report As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btn_reset As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_close As System.Windows.Forms.Button
    Friend WithEvents rdt_reportBC1 As QLNS.BC01
    Friend WithEvents rdt_reportBC2 As QLNS.BC02
    Friend WithEvents rdt_reportBC3 As QLNS.BC03
    Friend WithEvents rdt_reportBC4 As QLNS.BC04
    Friend WithEvents rdt_reportBC11 As QLNS.BC01_NN
    Friend WithEvents pnl_2 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cb_kybc As System.Windows.Forms.ComboBox
    Friend WithEvents lbl_kybc As System.Windows.Forms.Label
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents edt_truongphong As System.Windows.Forms.TextBox
    Friend WithEvents lbl_truongphong As System.Windows.Forms.Label
    Friend WithEvents edt_nguoilap As System.Windows.Forms.TextBox
    Friend WithEvents lbl_nguoilap As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rdt_reportBC2TTDT As QLNS.BC02TTDT
    Friend WithEvents rdt_reportBC2CNTT As QLNS.BC02CNTT
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dtpk_ngaylap_bc As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbl_ngaylap As System.Windows.Forms.Label
    Friend WithEvents edt_giamdoc As System.Windows.Forms.TextBox
    Friend WithEvents lbl_giamdoc As System.Windows.Forms.Label
    Friend WithEvents edt_so_bc As System.Windows.Forms.TextBox
    Friend WithEvents lbl_sobc As System.Windows.Forms.Label
    Friend WithEvents labTuThang As System.Windows.Forms.Label
    Friend WithEvents dtpk_nam_bc As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbl_nambc As System.Windows.Forms.Label
    Friend WithEvents labDenThang As System.Windows.Forms.Label
    Friend WithEvents numDenThang As System.Windows.Forms.NumericUpDown
    Friend WithEvents numTuThang As System.Windows.Forms.NumericUpDown
    Friend WithEvents rpt_BC02QT As QLNS.BC02_TQ
    Private WithEvents crpv_main As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents rptBC02_HSC As QLNS.BC02_HSC
    Friend WithEvents rptBC02SGD As QLNS.BC02SGD
    Friend WithEvents btn_exportexcel As System.Windows.Forms.Button
    Friend WithEvents ckb_ChiTietPGD As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
