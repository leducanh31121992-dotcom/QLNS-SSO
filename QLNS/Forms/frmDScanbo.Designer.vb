<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDScanbo
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDScanbo))
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.labStatusProcess = New System.Windows.Forms.Label()
        Me.bntSearch = New System.Windows.Forms.Button()
        Me.bntExportFile = New System.Windows.Forms.Button()
        Me.bntRefresh = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.dtpkNam = New System.Windows.Forms.DateTimePicker()
        Me.dpkDenNgay = New System.Windows.Forms.DateTimePicker()
        Me.labThoiDiem = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.grpResult = New System.Windows.Forms.GroupBox()
        Me.gridResult = New System.Windows.Forms.DataGridView()
        Me.Panel5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
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
        Me.Panel5.Size = New System.Drawing.Size(785, 24)
        Me.Panel5.TabIndex = 3
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
        Me.bntSearch.Location = New System.Drawing.Point(405, 0)
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
        Me.bntExportFile.Location = New System.Drawing.Point(500, 0)
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
        Me.bntRefresh.Location = New System.Drawing.Point(595, 0)
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
        Me.bntClose.Location = New System.Drawing.Point(690, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(95, 24)
        Me.bntClose.TabIndex = 3
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Panel2)
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(785, 67)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Thông tin tra cứu"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dtpkNam)
        Me.Panel2.Controls.Add(Me.dpkDenNgay)
        Me.Panel2.Controls.Add(Me.labThoiDiem)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 39)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(779, 22)
        Me.Panel2.TabIndex = 171
        '
        'dtpkNam
        '
        Me.dtpkNam.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpkNam.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpkNam.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpkNam.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpkNam.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpkNam.CustomFormat = "yyyy"
        Me.dtpkNam.Dock = System.Windows.Forms.DockStyle.Left
        Me.dtpkNam.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpkNam.Location = New System.Drawing.Point(218, 0)
        Me.dtpkNam.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam.Name = "dtpkNam"
        Me.dtpkNam.ShowUpDown = True
        Me.dtpkNam.Size = New System.Drawing.Size(48, 21)
        Me.dtpkNam.TabIndex = 169
        '
        'dpkDenNgay
        '
        Me.dpkDenNgay.CustomFormat = "dd/MM/yyyy"
        Me.dpkDenNgay.Dock = System.Windows.Forms.DockStyle.Left
        Me.dpkDenNgay.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkDenNgay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkDenNgay.Location = New System.Drawing.Point(81, 0)
        Me.dpkDenNgay.Name = "dpkDenNgay"
        Me.dpkDenNgay.Size = New System.Drawing.Size(137, 21)
        Me.dpkDenNgay.TabIndex = 10
        '
        'labThoiDiem
        '
        Me.labThoiDiem.Dock = System.Windows.Forms.DockStyle.Left
        Me.labThoiDiem.Location = New System.Drawing.Point(0, 0)
        Me.labThoiDiem.Name = "labThoiDiem"
        Me.labThoiDiem.Size = New System.Drawing.Size(81, 22)
        Me.labThoiDiem.TabIndex = 168
        Me.labThoiDiem.Text = "Đến ngày"
        Me.labThoiDiem.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cboDonVi)
        Me.Panel1.Controls.Add(Me.Label40)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(3, 17)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(779, 22)
        Me.Panel1.TabIndex = 170
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(81, 0)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(698, 21)
        Me.cboDonVi.TabIndex = 0
        Me.cboDonVi.ValueMember = "Value"
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(0, 0)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(81, 22)
        Me.Label40.TabIndex = 154
        Me.Label40.Text = "Đơn vị"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpResult
        '
        Me.grpResult.Controls.Add(Me.gridResult)
        Me.grpResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpResult.Location = New System.Drawing.Point(0, 91)
        Me.grpResult.Name = "grpResult"
        Me.grpResult.Size = New System.Drawing.Size(785, 470)
        Me.grpResult.TabIndex = 19
        Me.grpResult.TabStop = False
        Me.grpResult.Text = "Danh sách cán bộ viên chức"
        '
        'gridResult
        '
        Me.gridResult.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridResult.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.gridResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gridResult.Location = New System.Drawing.Point(3, 17)
        Me.gridResult.Name = "gridResult"
        Me.gridResult.ReadOnly = True
        Me.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridResult.Size = New System.Drawing.Size(779, 450)
        Me.gridResult.TabIndex = 105
        '
        'frmDScanbo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(785, 561)
        Me.Controls.Add(Me.grpResult)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel5)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmDScanbo"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: DANH SÁCH CÁN BỘ VIÊN CHỨC"
        Me.Panel5.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
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
    Friend WithEvents dpkDenNgay As System.Windows.Forms.DateTimePicker
    Friend WithEvents labThoiDiem As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents grpResult As System.Windows.Forms.GroupBox
    Friend WithEvents gridResult As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents dtpkNam As System.Windows.Forms.DateTimePicker
End Class
