<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTimKiem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTimKiem))
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.labStatusProcess = New System.Windows.Forms.Label()
        Me.bntSearch = New System.Windows.Forms.Button()
        Me.bntExportFile = New System.Windows.Forms.Button()
        Me.bntRefresh = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dpkNgayNHCS_Den = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.dpkNgayNHCS_Tu = New System.Windows.Forms.DateTimePicker()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.cbNgayNHCS = New System.Windows.Forms.CheckBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cbSoBHXH = New System.Windows.Forms.CheckBox()
        Me.txtHuuNam = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtHuuNu = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.dpkHuuTri = New System.Windows.Forms.DateTimePicker()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cbHuuTri = New System.Windows.Forms.CheckBox()
        Me.cboToanTu = New System.Windows.Forms.ComboBox()
        Me.txtHeSo = New System.Windows.Forms.TextBox()
        Me.cbHSL = New System.Windows.Forms.CheckBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cbChucvu = New System.Windows.Forms.CheckBox()
        Me.cbPhong = New System.Windows.Forms.CheckBox()
        Me.cbTonGiao = New System.Windows.Forms.CheckBox()
        Me.cbDT = New System.Windows.Forms.CheckBox()
        Me.cboChucvu = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cboPhong = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.dpkDenNgay = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cbNS = New System.Windows.Forms.CheckBox()
        Me.dpkTuNgay = New System.Windows.Forms.DateTimePicker()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cboTonGiao = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboDanToc = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbNu = New System.Windows.Forms.CheckBox()
        Me.cbNam = New System.Windows.Forms.CheckBox()
        Me.txtMaCB = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtHoten = New System.Windows.Forms.TextBox()
        Me.labGD = New System.Windows.Forms.Label()
        Me.grpResult = New System.Windows.Forms.GroupBox()
        Me.gridResult = New System.Windows.Forms.DataGridView()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Panel5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.grpResult.SuspendLayout()
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.DarkGray
        Me.Panel5.Controls.Add(Me.labStatusProcess)
        Me.Panel5.Controls.Add(Me.bntSearch)
        Me.Panel5.Controls.Add(Me.bntExportFile)
        Me.Panel5.Controls.Add(Me.bntRefresh)
        Me.Panel5.Controls.Add(Me.bntClose)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(1006, 24)
        Me.Panel5.TabIndex = 2
        '
        'labStatusProcess
        '
        Me.labStatusProcess.Dock = System.Windows.Forms.DockStyle.Left
        Me.labStatusProcess.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labStatusProcess.ForeColor = System.Drawing.Color.Maroon
        Me.labStatusProcess.Location = New System.Drawing.Point(0, 0)
        Me.labStatusProcess.Name = "labStatusProcess"
        Me.labStatusProcess.Size = New System.Drawing.Size(136, 24)
        Me.labStatusProcess.TabIndex = 18
        Me.labStatusProcess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'bntSearch
        '
        Me.bntSearch.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntSearch.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntSearch.Location = New System.Drawing.Point(626, 0)
        Me.bntSearch.Name = "bntSearch"
        Me.bntSearch.Size = New System.Drawing.Size(95, 24)
        Me.bntSearch.TabIndex = 0
        Me.bntSearch.Text = "&Tra cứu"
        Me.bntSearch.UseVisualStyleBackColor = True
        '
        'bntExportFile
        '
        Me.bntExportFile.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntExportFile.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntExportFile.Location = New System.Drawing.Point(721, 0)
        Me.bntExportFile.Name = "bntExportFile"
        Me.bntExportFile.Size = New System.Drawing.Size(95, 24)
        Me.bntExportFile.TabIndex = 1
        Me.bntExportFile.Text = "Xuất &file Excel"
        Me.bntExportFile.UseVisualStyleBackColor = True
        '
        'bntRefresh
        '
        Me.bntRefresh.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntRefresh.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntRefresh.Location = New System.Drawing.Point(816, 0)
        Me.bntRefresh.Name = "bntRefresh"
        Me.bntRefresh.Size = New System.Drawing.Size(95, 24)
        Me.bntRefresh.TabIndex = 2
        Me.bntRefresh.Text = "Tra cứu &mới"
        Me.bntRefresh.UseVisualStyleBackColor = True
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntClose.Location = New System.Drawing.Point(911, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(95, 24)
        Me.bntClose.TabIndex = 3
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.dpkNgayNHCS_Den)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.dpkNgayNHCS_Tu)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.cbNgayNHCS)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.cbSoBHXH)
        Me.GroupBox1.Controls.Add(Me.txtHuuNam)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.txtHuuNu)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.dpkHuuTri)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.cbHuuTri)
        Me.GroupBox1.Controls.Add(Me.cboToanTu)
        Me.GroupBox1.Controls.Add(Me.txtHeSo)
        Me.GroupBox1.Controls.Add(Me.cbHSL)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.cbChucvu)
        Me.GroupBox1.Controls.Add(Me.cbPhong)
        Me.GroupBox1.Controls.Add(Me.cbTonGiao)
        Me.GroupBox1.Controls.Add(Me.cbDT)
        Me.GroupBox1.Controls.Add(Me.cboChucvu)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.cboPhong)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.cboDonVi)
        Me.GroupBox1.Controls.Add(Me.cbAll)
        Me.GroupBox1.Controls.Add(Me.Label40)
        Me.GroupBox1.Controls.Add(Me.dpkDenNgay)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.cbNS)
        Me.GroupBox1.Controls.Add(Me.dpkTuNgay)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.cboTonGiao)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.cboDanToc)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cbNu)
        Me.GroupBox1.Controls.Add(Me.cbNam)
        Me.GroupBox1.Controls.Add(Me.txtMaCB)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtHoten)
        Me.GroupBox1.Controls.Add(Me.labGD)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1006, 240)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Thông tin tra cứu"
        '
        'dpkNgayNHCS_Den
        '
        Me.dpkNgayNHCS_Den.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayNHCS_Den.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayNHCS_Den.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayNHCS_Den.Location = New System.Drawing.Point(289, 125)
        Me.dpkNgayNHCS_Den.Name = "dpkNgayNHCS_Den"
        Me.dpkNgayNHCS_Den.Size = New System.Drawing.Size(92, 26)
        Me.dpkNgayNHCS_Den.TabIndex = 178
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(262, 130)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(33, 18)
        Me.Label9.TabIndex = 180
        Me.Label9.Text = "đến"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dpkNgayNHCS_Tu
        '
        Me.dpkNgayNHCS_Tu.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayNHCS_Tu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayNHCS_Tu.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayNHCS_Tu.Location = New System.Drawing.Point(164, 125)
        Me.dpkNgayNHCS_Tu.Name = "dpkNgayNHCS_Tu"
        Me.dpkNgayNHCS_Tu.Size = New System.Drawing.Size(92, 26)
        Me.dpkNgayNHCS_Tu.TabIndex = 177
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(145, 130)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(22, 18)
        Me.Label17.TabIndex = 179
        Me.Label17.Text = "từ"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbNgayNHCS
        '
        Me.cbNgayNHCS.AutoSize = True
        Me.cbNgayNHCS.Location = New System.Drawing.Point(102, 132)
        Me.cbNgayNHCS.Name = "cbNgayNHCS"
        Me.cbNgayNHCS.Size = New System.Drawing.Size(18, 17)
        Me.cbNgayNHCS.TabIndex = 176
        Me.cbNgayNHCS.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(2, 126)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(92, 24)
        Me.Label8.TabIndex = 175
        Me.Label8.Text = "Ngày vào NHCS"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbSoBHXH
        '
        Me.cbSoBHXH.AutoSize = True
        Me.cbSoBHXH.Location = New System.Drawing.Point(479, 127)
        Me.cbSoBHXH.Name = "cbSoBHXH"
        Me.cbSoBHXH.Size = New System.Drawing.Size(155, 22)
        Me.cbSoBHXH.TabIndex = 173
        Me.cbSoBHXH.Text = "Chưa làm sổ BHXH"
        Me.cbSoBHXH.UseVisualStyleBackColor = True
        '
        'txtHuuNam
        '
        Me.txtHuuNam.BackColor = System.Drawing.Color.White
        Me.txtHuuNam.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHuuNam.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHuuNam.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtHuuNam.Location = New System.Drawing.Point(351, 174)
        Me.txtHuuNam.Name = "txtHuuNam"
        Me.txtHuuNam.ReadOnly = True
        Me.txtHuuNam.Size = New System.Drawing.Size(30, 26)
        Me.txtHuuNam.TabIndex = 12
        Me.txtHuuNam.Text = "60"
        Me.txtHuuNam.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(274, 178)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(94, 18)
        Me.Label16.TabIndex = 172
        Me.Label16.Text = ", đối với nam"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHuuNu
        '
        Me.txtHuuNu.BackColor = System.Drawing.Color.White
        Me.txtHuuNu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHuuNu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHuuNu.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtHuuNu.Location = New System.Drawing.Point(244, 174)
        Me.txtHuuNu.Name = "txtHuuNu"
        Me.txtHuuNu.ReadOnly = True
        Me.txtHuuNu.Size = New System.Drawing.Size(30, 26)
        Me.txtHuuNu.TabIndex = 11
        Me.txtHuuNu.Text = "55"
        Me.txtHuuNu.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(22, 178)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(220, 18)
        Me.Label15.TabIndex = 170
        Me.Label15.Text = "Tuổi nghỉ hưu cơ bản: đối với nữ"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dpkHuuTri
        '
        Me.dpkHuuTri.CustomFormat = "dd/MM/yyyy"
        Me.dpkHuuTri.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkHuuTri.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkHuuTri.Location = New System.Drawing.Point(244, 152)
        Me.dpkHuuTri.Name = "dpkHuuTri"
        Me.dpkHuuTri.Size = New System.Drawing.Size(137, 26)
        Me.dpkHuuTri.TabIndex = 10
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(145, 156)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(98, 18)
        Me.Label14.TabIndex = 168
        Me.Label14.Text = "tính đến ngày"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(49, 157)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(52, 18)
        Me.Label13.TabIndex = 166
        Me.Label13.Text = "Hưu trí"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbHuuTri
        '
        Me.cbHuuTri.AutoSize = True
        Me.cbHuuTri.Location = New System.Drawing.Point(102, 157)
        Me.cbHuuTri.Name = "cbHuuTri"
        Me.cbHuuTri.Size = New System.Drawing.Size(18, 17)
        Me.cbHuuTri.TabIndex = 9
        Me.cbHuuTri.UseVisualStyleBackColor = True
        '
        'cboToanTu
        '
        Me.cboToanTu.BackColor = System.Drawing.SystemColors.Window
        Me.cboToanTu.DisplayMember = "Display"
        Me.cboToanTu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboToanTu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboToanTu.FormattingEnabled = True
        Me.cboToanTu.Location = New System.Drawing.Point(498, 147)
        Me.cboToanTu.Name = "cboToanTu"
        Me.cboToanTu.Size = New System.Drawing.Size(130, 26)
        Me.cboToanTu.TabIndex = 18
        Me.cboToanTu.ValueMember = "Value"
        '
        'txtHeSo
        '
        Me.txtHeSo.BackColor = System.Drawing.Color.White
        Me.txtHeSo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHeSo.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHeSo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtHeSo.Location = New System.Drawing.Point(639, 147)
        Me.txtHeSo.Name = "txtHeSo"
        Me.txtHeSo.Size = New System.Drawing.Size(145, 26)
        Me.txtHeSo.TabIndex = 19
        '
        'cbHSL
        '
        Me.cbHSL.AutoSize = True
        Me.cbHSL.Location = New System.Drawing.Point(479, 151)
        Me.cbHSL.Name = "cbHSL"
        Me.cbHSL.Size = New System.Drawing.Size(18, 17)
        Me.cbHSL.TabIndex = 17
        Me.cbHSL.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(398, 144)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(74, 24)
        Me.Label12.TabIndex = 164
        Me.Label12.Text = "Hệ số lương"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbChucvu
        '
        Me.cbChucvu.AutoSize = True
        Me.cbChucvu.Location = New System.Drawing.Point(479, 65)
        Me.cbChucvu.Name = "cbChucvu"
        Me.cbChucvu.Size = New System.Drawing.Size(18, 17)
        Me.cbChucvu.TabIndex = 15
        Me.cbChucvu.UseVisualStyleBackColor = True
        '
        'cbPhong
        '
        Me.cbPhong.AutoSize = True
        Me.cbPhong.Location = New System.Drawing.Point(479, 44)
        Me.cbPhong.Name = "cbPhong"
        Me.cbPhong.Size = New System.Drawing.Size(18, 17)
        Me.cbPhong.TabIndex = 13
        Me.cbPhong.UseVisualStyleBackColor = True
        '
        'cbTonGiao
        '
        Me.cbTonGiao.AutoSize = True
        Me.cbTonGiao.Location = New System.Drawing.Point(479, 108)
        Me.cbTonGiao.Name = "cbTonGiao"
        Me.cbTonGiao.Size = New System.Drawing.Size(18, 17)
        Me.cbTonGiao.TabIndex = 22
        Me.cbTonGiao.UseVisualStyleBackColor = True
        '
        'cbDT
        '
        Me.cbDT.AutoSize = True
        Me.cbDT.Location = New System.Drawing.Point(479, 87)
        Me.cbDT.Name = "cbDT"
        Me.cbDT.Size = New System.Drawing.Size(18, 17)
        Me.cbDT.TabIndex = 20
        Me.cbDT.UseVisualStyleBackColor = True
        '
        'cboChucvu
        '
        Me.cboChucvu.BackColor = System.Drawing.SystemColors.Window
        Me.cboChucvu.DisplayMember = "Display"
        Me.cboChucvu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChucvu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboChucvu.FormattingEnabled = True
        Me.cboChucvu.Location = New System.Drawing.Point(498, 60)
        Me.cboChucvu.Name = "cboChucvu"
        Me.cboChucvu.Size = New System.Drawing.Size(286, 26)
        Me.cboChucvu.TabIndex = 16
        Me.cboChucvu.ValueMember = "Value"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(398, 59)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(74, 24)
        Me.Label11.TabIndex = 159
        Me.Label11.Text = "Chức vụ"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboPhong
        '
        Me.cboPhong.BackColor = System.Drawing.SystemColors.Window
        Me.cboPhong.DisplayMember = "Display"
        Me.cboPhong.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPhong.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhong.FormattingEnabled = True
        Me.cboPhong.Location = New System.Drawing.Point(498, 38)
        Me.cboPhong.Name = "cboPhong"
        Me.cboPhong.Size = New System.Drawing.Size(286, 26)
        Me.cboPhong.TabIndex = 14
        Me.cboPhong.ValueMember = "Value"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(398, 38)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 24)
        Me.Label7.TabIndex = 157
        Me.Label7.Text = "Phòng"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(102, 16)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(526, 26)
        Me.cboDonVi.TabIndex = 0
        Me.cboDonVi.ValueMember = "Value"
        '
        'cbAll
        '
        Me.cbAll.AutoSize = True
        Me.cbAll.Location = New System.Drawing.Point(639, 18)
        Me.cbAll.Name = "cbAll"
        Me.cbAll.Size = New System.Drawing.Size(171, 22)
        Me.cbAll.TabIndex = 1
        Me.cbAll.Text = "Thống kê toàn đơn vị"
        Me.cbAll.UseVisualStyleBackColor = True
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(48, 15)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(46, 22)
        Me.Label40.TabIndex = 154
        Me.Label40.Text = "Đơn vị"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dpkDenNgay
        '
        Me.dpkDenNgay.CustomFormat = "dd/MM/yyyy"
        Me.dpkDenNgay.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkDenNgay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkDenNgay.Location = New System.Drawing.Point(289, 103)
        Me.dpkDenNgay.Name = "dpkDenNgay"
        Me.dpkDenNgay.Size = New System.Drawing.Size(92, 26)
        Me.dpkDenNgay.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(262, 108)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(33, 18)
        Me.Label6.TabIndex = 152
        Me.Label6.Text = "đến"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(3, 102)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(92, 24)
        Me.Label5.TabIndex = 150
        Me.Label5.Text = "Ngày sinh"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbNS
        '
        Me.cbNS.AutoSize = True
        Me.cbNS.Location = New System.Drawing.Point(102, 109)
        Me.cbNS.Name = "cbNS"
        Me.cbNS.Size = New System.Drawing.Size(18, 17)
        Me.cbNS.TabIndex = 6
        Me.cbNS.UseVisualStyleBackColor = True
        '
        'dpkTuNgay
        '
        Me.dpkTuNgay.CustomFormat = "dd/MM/yyyy"
        Me.dpkTuNgay.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkTuNgay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkTuNgay.Location = New System.Drawing.Point(164, 103)
        Me.dpkTuNgay.Name = "dpkTuNgay"
        Me.dpkTuNgay.Size = New System.Drawing.Size(92, 26)
        Me.dpkTuNgay.TabIndex = 7
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(145, 108)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(22, 18)
        Me.Label10.TabIndex = 148
        Me.Label10.Text = "từ"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cboTonGiao
        '
        Me.cboTonGiao.BackColor = System.Drawing.SystemColors.Window
        Me.cboTonGiao.DisplayMember = "Display"
        Me.cboTonGiao.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTonGiao.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTonGiao.FormattingEnabled = True
        Me.cboTonGiao.Location = New System.Drawing.Point(498, 104)
        Me.cboTonGiao.Name = "cboTonGiao"
        Me.cboTonGiao.Size = New System.Drawing.Size(286, 26)
        Me.cboTonGiao.TabIndex = 23
        Me.cboTonGiao.ValueMember = "Value"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(398, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 24)
        Me.Label4.TabIndex = 146
        Me.Label4.Text = "Tôn giáo"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboDanToc
        '
        Me.cboDanToc.BackColor = System.Drawing.SystemColors.Window
        Me.cboDanToc.DisplayMember = "Display"
        Me.cboDanToc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDanToc.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDanToc.FormattingEnabled = True
        Me.cboDanToc.Location = New System.Drawing.Point(498, 82)
        Me.cboDanToc.Name = "cboDanToc"
        Me.cboDanToc.Size = New System.Drawing.Size(286, 26)
        Me.cboDanToc.TabIndex = 21
        Me.cboDanToc.ValueMember = "Value"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(401, 81)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 24)
        Me.Label3.TabIndex = 144
        Me.Label3.Text = "Dân tộc"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(0, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 24)
        Me.Label2.TabIndex = 142
        Me.Label2.Text = "Giới tính"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbNu
        '
        Me.cbNu.AutoSize = True
        Me.cbNu.Location = New System.Drawing.Point(164, 83)
        Me.cbNu.Name = "cbNu"
        Me.cbNu.Size = New System.Drawing.Size(49, 22)
        Me.cbNu.TabIndex = 5
        Me.cbNu.Text = "Nữ"
        Me.cbNu.UseVisualStyleBackColor = True
        '
        'cbNam
        '
        Me.cbNam.AutoSize = True
        Me.cbNam.Location = New System.Drawing.Point(102, 83)
        Me.cbNam.Name = "cbNam"
        Me.cbNam.Size = New System.Drawing.Size(61, 22)
        Me.cbNam.TabIndex = 4
        Me.cbNam.Text = "Nam"
        Me.cbNam.UseVisualStyleBackColor = True
        '
        'txtMaCB
        '
        Me.txtMaCB.BackColor = System.Drawing.Color.White
        Me.txtMaCB.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMaCB.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaCB.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtMaCB.Location = New System.Drawing.Point(102, 60)
        Me.txtMaCB.Name = "txtMaCB"
        Me.txtMaCB.Size = New System.Drawing.Size(279, 26)
        Me.txtMaCB.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(-3, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 24)
        Me.Label1.TabIndex = 139
        Me.Label1.Text = "Mã cán bộ"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHoten
        '
        Me.txtHoten.BackColor = System.Drawing.Color.White
        Me.txtHoten.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHoten.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHoten.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtHoten.Location = New System.Drawing.Point(102, 38)
        Me.txtHoten.Name = "txtHoten"
        Me.txtHoten.Size = New System.Drawing.Size(279, 26)
        Me.txtHoten.TabIndex = 2
        '
        'labGD
        '
        Me.labGD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labGD.Location = New System.Drawing.Point(-6, 36)
        Me.labGD.Name = "labGD"
        Me.labGD.Size = New System.Drawing.Size(100, 24)
        Me.labGD.TabIndex = 137
        Me.labGD.Text = "Họ tên cán bộ"
        Me.labGD.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpResult
        '
        Me.grpResult.Controls.Add(Me.gridResult)
        Me.grpResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpResult.Location = New System.Drawing.Point(0, 264)
        Me.grpResult.Name = "grpResult"
        Me.grpResult.Size = New System.Drawing.Size(1006, 389)
        Me.grpResult.TabIndex = 18
        Me.grpResult.TabStop = False
        Me.grpResult.Text = "Kết quả tra cứu"
        '
        'gridResult
        '
        Me.gridResult.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridResult.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.gridResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gridResult.Location = New System.Drawing.Point(3, 22)
        Me.gridResult.Name = "gridResult"
        Me.gridResult.ReadOnly = True
        Me.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridResult.Size = New System.Drawing.Size(1000, 364)
        Me.gridResult.TabIndex = 105
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Maroon
        Me.Label18.Location = New System.Drawing.Point(387, 182)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(442, 18)
        Me.Label18.TabIndex = 181
        Me.Label18.Text = "Khi tìm kiếm phần mềm tự động tính toán theo luật lao động 2019"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmTimKiem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1006, 653)
        Me.Controls.Add(Me.grpResult)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel5)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmTimKiem"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: BÁO CÁO NHANH"
        Me.Panel5.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grpResult.ResumeLayout(False)
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents labStatusProcess As System.Windows.Forms.Label
    Friend WithEvents bntSearch As System.Windows.Forms.Button
    Friend WithEvents bntExportFile As System.Windows.Forms.Button
    Friend WithEvents bntRefresh As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbNu As System.Windows.Forms.CheckBox
    Friend WithEvents cbNam As System.Windows.Forms.CheckBox
    Friend WithEvents txtMaCB As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHoten As System.Windows.Forms.TextBox
    Friend WithEvents labGD As System.Windows.Forms.Label
    Friend WithEvents cboTonGiao As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboDanToc As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbNS As System.Windows.Forms.CheckBox
    Friend WithEvents dpkTuNgay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents dpkDenNgay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboPhong As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents cbDT As System.Windows.Forms.CheckBox
    Friend WithEvents cboChucvu As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboToanTu As System.Windows.Forms.ComboBox
    Friend WithEvents txtHeSo As System.Windows.Forms.TextBox
    Friend WithEvents cbHSL As System.Windows.Forms.CheckBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cbChucvu As System.Windows.Forms.CheckBox
    Friend WithEvents cbPhong As System.Windows.Forms.CheckBox
    Friend WithEvents cbTonGiao As System.Windows.Forms.CheckBox
    Friend WithEvents grpResult As System.Windows.Forms.GroupBox
    Friend WithEvents gridResult As System.Windows.Forms.DataGridView
    Friend WithEvents txtHuuNu As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents dpkHuuTri As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cbHuuTri As System.Windows.Forms.CheckBox
    Friend WithEvents txtHuuNam As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cbSoBHXH As System.Windows.Forms.CheckBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbNgayNHCS As System.Windows.Forms.CheckBox
    Friend WithEvents dpkNgayNHCS_Den As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dpkNgayNHCS_Tu As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
End Class
