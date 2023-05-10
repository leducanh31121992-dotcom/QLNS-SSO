<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLaoDongThuNhapReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLaoDongThuNhapReport))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtNam = New System.Windows.Forms.TextBox()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.dpkNgayLapBieu = New System.Windows.Forms.DateTimePicker()
        Me.txtSoBC = New System.Windows.Forms.TextBox()
        Me.txtGD = New System.Windows.Forms.TextBox()
        Me.txtLapBieu = New System.Windows.Forms.TextBox()
        Me.labGD = New System.Windows.Forms.Label()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.labHCTC = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.txtHCTC = New System.Windows.Forms.TextBox()
        Me.cmdView = New System.Windows.Forms.Button()
        Me.rpt_View = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.cmdCreate = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.labStatusProcess = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.rpt_BC05 = New QLNS.BC05()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 24)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(787, 117)
        Me.Panel1.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtNam)
        Me.GroupBox1.Controls.Add(Me.cbAll)
        Me.GroupBox1.Controls.Add(Me.Label41)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label39)
        Me.GroupBox1.Controls.Add(Me.dpkNgayLapBieu)
        Me.GroupBox1.Controls.Add(Me.txtSoBC)
        Me.GroupBox1.Controls.Add(Me.txtGD)
        Me.GroupBox1.Controls.Add(Me.txtLapBieu)
        Me.GroupBox1.Controls.Add(Me.labGD)
        Me.GroupBox1.Controls.Add(Me.Label38)
        Me.GroupBox1.Controls.Add(Me.cboDonVi)
        Me.GroupBox1.Controls.Add(Me.labHCTC)
        Me.GroupBox1.Controls.Add(Me.Label40)
        Me.GroupBox1.Controls.Add(Me.txtHCTC)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(787, 112)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Khai báo tham số"
        '
        'txtNam
        '
        Me.txtNam.BackColor = System.Drawing.Color.White
        Me.txtNam.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNam.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNam.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtNam.Location = New System.Drawing.Point(338, 40)
        Me.txtNam.Name = "txtNam"
        Me.txtNam.ReadOnly = True
        Me.txtNam.Size = New System.Drawing.Size(71, 22)
        Me.txtNam.TabIndex = 4
        '
        'cbAll
        '
        Me.cbAll.AutoSize = True
        Me.cbAll.Location = New System.Drawing.Point(627, 19)
        Me.cbAll.Name = "cbAll"
        Me.cbAll.Size = New System.Drawing.Size(129, 17)
        Me.cbAll.TabIndex = 2
        Me.cbAll.Text = "Thống kê toàn đơn vị"
        Me.cbAll.UseVisualStyleBackColor = True
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.SystemColors.Control
        Me.Label41.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.Location = New System.Drawing.Point(0, 83)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(146, 24)
        Me.Label41.TabIndex = 111
        Me.Label41.Text = "Người lập biểu"
        Me.Label41.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(474, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 24)
        Me.Label2.TabIndex = 124
        Me.Label2.Text = "Ngày lập biểu"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.SystemColors.Control
        Me.Label39.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.Location = New System.Drawing.Point(56, 39)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(90, 24)
        Me.Label39.TabIndex = 110
        Me.Label39.Text = "Số báo cáo"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dpkNgayLapBieu
        '
        Me.dpkNgayLapBieu.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayLapBieu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayLapBieu.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayLapBieu.Location = New System.Drawing.Point(555, 86)
        Me.dpkNgayLapBieu.Name = "dpkNgayLapBieu"
        Me.dpkNgayLapBieu.Size = New System.Drawing.Size(105, 22)
        Me.dpkNgayLapBieu.TabIndex = 8
        '
        'txtSoBC
        '
        Me.txtSoBC.BackColor = System.Drawing.Color.White
        Me.txtSoBC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSoBC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSoBC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtSoBC.Location = New System.Drawing.Point(146, 40)
        Me.txtSoBC.Name = "txtSoBC"
        Me.txtSoBC.Size = New System.Drawing.Size(139, 22)
        Me.txtSoBC.TabIndex = 3
        '
        'txtGD
        '
        Me.txtGD.BackColor = System.Drawing.Color.White
        Me.txtGD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGD.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtGD.Location = New System.Drawing.Point(146, 63)
        Me.txtGD.Name = "txtGD"
        Me.txtGD.Size = New System.Drawing.Size(263, 22)
        Me.txtGD.TabIndex = 5
        '
        'txtLapBieu
        '
        Me.txtLapBieu.BackColor = System.Drawing.Color.White
        Me.txtLapBieu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLapBieu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLapBieu.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtLapBieu.Location = New System.Drawing.Point(146, 86)
        Me.txtLapBieu.Name = "txtLapBieu"
        Me.txtLapBieu.Size = New System.Drawing.Size(263, 22)
        Me.txtLapBieu.TabIndex = 7
        '
        'labGD
        '
        Me.labGD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labGD.Location = New System.Drawing.Point(44, 62)
        Me.labGD.Name = "labGD"
        Me.labGD.Size = New System.Drawing.Size(101, 24)
        Me.labGD.TabIndex = 122
        Me.labGD.Text = "Giám đốc"
        Me.labGD.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label38
        '
        Me.Label38.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.Location = New System.Drawing.Point(291, 39)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(45, 24)
        Me.Label38.TabIndex = 116
        Me.Label38.Text = "Năm"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(146, 17)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(475, 22)
        Me.cboDonVi.TabIndex = 1
        Me.cboDonVi.ValueMember = "Value"
        '
        'labHCTC
        '
        Me.labHCTC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labHCTC.Location = New System.Drawing.Point(409, 60)
        Me.labHCTC.Name = "labHCTC"
        Me.labHCTC.Size = New System.Drawing.Size(146, 24)
        Me.labHCTC.TabIndex = 117
        Me.labHCTC.Text = "Trưởng phòng HC-TC"
        Me.labHCTC.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(6, 14)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(139, 24)
        Me.Label40.TabIndex = 113
        Me.Label40.Text = "Đơn vị"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHCTC
        '
        Me.txtHCTC.BackColor = System.Drawing.Color.White
        Me.txtHCTC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHCTC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHCTC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtHCTC.Location = New System.Drawing.Point(555, 63)
        Me.txtHCTC.Name = "txtHCTC"
        Me.txtHCTC.Size = New System.Drawing.Size(226, 22)
        Me.txtHCTC.TabIndex = 6
        '
        'cmdView
        '
        Me.cmdView.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdView.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdView.Location = New System.Drawing.Point(403, 0)
        Me.cmdView.Name = "cmdView"
        Me.cmdView.Size = New System.Drawing.Size(95, 24)
        Me.cmdView.TabIndex = 1
        Me.cmdView.Text = "&Xem báo cáo"
        Me.cmdView.UseVisualStyleBackColor = True
        '
        'rpt_View
        '
        Me.rpt_View.ActiveViewIndex = -1
        Me.rpt_View.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rpt_View.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.rpt_View.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rpt_View.Location = New System.Drawing.Point(0, 141)
        Me.rpt_View.Name = "rpt_View"
        Me.rpt_View.SelectionFormula = ""
        Me.rpt_View.ShowRefreshButton = False
        Me.rpt_View.Size = New System.Drawing.Size(787, 496)
        Me.rpt_View.TabIndex = 4
        Me.rpt_View.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        Me.rpt_View.UseWaitCursor = True
        Me.rpt_View.ViewTimeSelectionFormula = ""
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.Controls.Add(Me.cmdView)
        Me.Panel3.Controls.Add(Me.cmdCreate)
        Me.Panel3.Controls.Add(Me.cmdSave)
        Me.Panel3.Controls.Add(Me.labStatusProcess)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.cmdClose)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(787, 24)
        Me.Panel3.TabIndex = 11
        '
        'cmdCreate
        '
        Me.cmdCreate.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdCreate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCreate.Location = New System.Drawing.Point(498, 0)
        Me.cmdCreate.Name = "cmdCreate"
        Me.cmdCreate.Size = New System.Drawing.Size(95, 24)
        Me.cmdCreate.TabIndex = 2
        Me.cmdCreate.Text = "&Tạo báo cáo"
        Me.cmdCreate.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdSave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.Location = New System.Drawing.Point(593, 0)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(95, 24)
        Me.cmdSave.TabIndex = 3
        Me.cmdSave.Text = "&Lưu báo cáo"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'labStatusProcess
        '
        Me.labStatusProcess.AutoSize = True
        Me.labStatusProcess.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labStatusProcess.ForeColor = System.Drawing.Color.Maroon
        Me.labStatusProcess.Location = New System.Drawing.Point(0, 7)
        Me.labStatusProcess.Name = "labStatusProcess"
        Me.labStatusProcess.Size = New System.Drawing.Size(0, 13)
        Me.labStatusProcess.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label5.Location = New System.Drawing.Point(688, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(2, 24)
        Me.Label5.TabIndex = 15
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label4.Location = New System.Drawing.Point(690, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(2, 24)
        Me.Label4.TabIndex = 17
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdClose
        '
        Me.cmdClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClose.Location = New System.Drawing.Point(692, 0)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(95, 24)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "&Quay ra"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmLaoDongThuNhapReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 637)
        Me.Controls.Add(Me.rpt_View)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmLaoDongThuNhapReport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Báo cáo 05: TÌNH HÌNH THỰC HIỆN LAO ĐỘNG THU NHẬP"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dpkNgayLapBieu As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmdView As System.Windows.Forms.Button
    Friend WithEvents txtGD As System.Windows.Forms.TextBox
    Friend WithEvents labGD As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents txtHCTC As System.Windows.Forms.TextBox
    Friend WithEvents txtNam As System.Windows.Forms.TextBox
    Friend WithEvents labHCTC As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents txtLapBieu As System.Windows.Forms.TextBox
    Friend WithEvents txtSoBC As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents rpt_View As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents rpt_BC05 As QLNS.BC05
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents labStatusProcess As System.Windows.Forms.Label
    Friend WithEvents cmdCreate As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
End Class
