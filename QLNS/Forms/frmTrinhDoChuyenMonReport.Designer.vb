<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTrinhDoChuyenMonReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTrinhDoChuyenMonReport))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dpkThoiDiemBC = New System.Windows.Forms.DateTimePicker()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.txtLapBieu = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.dpkNgayLapBieu = New System.Windows.Forms.DateTimePicker()
        Me.labHCTC = New System.Windows.Forms.Label()
        Me.txtGD = New System.Windows.Forms.TextBox()
        Me.txtHCTC = New System.Windows.Forms.TextBox()
        Me.labGD = New System.Windows.Forms.Label()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.labStatusProcess = New System.Windows.Forms.Label()
        Me.cmdReport = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.rpt_View = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.rpt_BC09 = New QLNS.BC09()
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
        Me.Panel1.Size = New System.Drawing.Size(787, 115)
        Me.Panel1.TabIndex = 5
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dpkThoiDiemBC)
        Me.GroupBox1.Controls.Add(Me.cbAll)
        Me.GroupBox1.Controls.Add(Me.txtLapBieu)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label41)
        Me.GroupBox1.Controls.Add(Me.dpkNgayLapBieu)
        Me.GroupBox1.Controls.Add(Me.labHCTC)
        Me.GroupBox1.Controls.Add(Me.txtGD)
        Me.GroupBox1.Controls.Add(Me.txtHCTC)
        Me.GroupBox1.Controls.Add(Me.labGD)
        Me.GroupBox1.Controls.Add(Me.Label40)
        Me.GroupBox1.Controls.Add(Me.cboDonVi)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(787, 115)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Khai báo tham sô"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(429, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 24)
        Me.Label1.TabIndex = 127
        Me.Label1.Text = "Ngày báo cáo"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dpkThoiDiemBC
        '
        Me.dpkThoiDiemBC.CustomFormat = "dd/MM/yyyy"
        Me.dpkThoiDiemBC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkThoiDiemBC.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkThoiDiemBC.Location = New System.Drawing.Point(146, 40)
        Me.dpkThoiDiemBC.Name = "dpkThoiDiemBC"
        Me.dpkThoiDiemBC.Size = New System.Drawing.Size(105, 22)
        Me.dpkThoiDiemBC.TabIndex = 3
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
        'txtLapBieu
        '
        Me.txtLapBieu.BackColor = System.Drawing.Color.White
        Me.txtLapBieu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLapBieu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLapBieu.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtLapBieu.Location = New System.Drawing.Point(146, 86)
        Me.txtLapBieu.Name = "txtLapBieu"
        Me.txtLapBieu.Size = New System.Drawing.Size(263, 22)
        Me.txtLapBieu.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(28, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(116, 24)
        Me.Label2.TabIndex = 124
        Me.Label2.Text = "Thời điểm báo cáo"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.SystemColors.Control
        Me.Label41.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.Location = New System.Drawing.Point(17, 85)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(127, 24)
        Me.Label41.TabIndex = 111
        Me.Label41.Text = "Người lập biểu"
        Me.Label41.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dpkNgayLapBieu
        '
        Me.dpkNgayLapBieu.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayLapBieu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayLapBieu.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayLapBieu.Location = New System.Drawing.Point(547, 86)
        Me.dpkNgayLapBieu.Name = "dpkNgayLapBieu"
        Me.dpkNgayLapBieu.Size = New System.Drawing.Size(105, 22)
        Me.dpkNgayLapBieu.TabIndex = 7
        '
        'labHCTC
        '
        Me.labHCTC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labHCTC.Location = New System.Drawing.Point(418, 61)
        Me.labHCTC.Name = "labHCTC"
        Me.labHCTC.Size = New System.Drawing.Size(127, 24)
        Me.labHCTC.TabIndex = 117
        Me.labHCTC.Text = "Trưởng phòng HC-TC"
        Me.labHCTC.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        Me.txtGD.TabIndex = 4
        '
        'txtHCTC
        '
        Me.txtHCTC.BackColor = System.Drawing.Color.White
        Me.txtHCTC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHCTC.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHCTC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtHCTC.Location = New System.Drawing.Point(547, 63)
        Me.txtHCTC.Name = "txtHCTC"
        Me.txtHCTC.Size = New System.Drawing.Size(234, 22)
        Me.txtHCTC.TabIndex = 5
        '
        'labGD
        '
        Me.labGD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labGD.Location = New System.Drawing.Point(43, 61)
        Me.labGD.Name = "labGD"
        Me.labGD.Size = New System.Drawing.Size(101, 24)
        Me.labGD.TabIndex = 122
        Me.labGD.Text = "Giám đốc"
        Me.labGD.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(16, 15)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(128, 24)
        Me.Label40.TabIndex = 113
        Me.Label40.Text = "Đơn vị"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.Controls.Add(Me.labStatusProcess)
        Me.Panel3.Controls.Add(Me.cmdReport)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.cmdClose)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(787, 24)
        Me.Panel3.TabIndex = 2
        '
        'labStatusProcess
        '
        Me.labStatusProcess.AutoSize = True
        Me.labStatusProcess.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labStatusProcess.ForeColor = System.Drawing.Color.Maroon
        Me.labStatusProcess.Location = New System.Drawing.Point(0, 7)
        Me.labStatusProcess.Name = "labStatusProcess"
        Me.labStatusProcess.Size = New System.Drawing.Size(0, 13)
        Me.labStatusProcess.TabIndex = 21
        '
        'cmdReport
        '
        Me.cmdReport.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReport.Location = New System.Drawing.Point(593, 0)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(95, 24)
        Me.cmdReport.TabIndex = 1
        Me.cmdReport.Text = "Xem báo cáo"
        Me.cmdReport.UseVisualStyleBackColor = True
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
        Me.cmdClose.TabIndex = 2
        Me.cmdClose.Text = "&Quay ra"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'rpt_View
        '
        Me.rpt_View.ActiveViewIndex = -1
        Me.rpt_View.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rpt_View.Cursor = System.Windows.Forms.Cursors.Default
        Me.rpt_View.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rpt_View.Location = New System.Drawing.Point(0, 139)
        Me.rpt_View.Name = "rpt_View"
        Me.rpt_View.SelectionFormula = ""
        Me.rpt_View.ShowRefreshButton = False
        Me.rpt_View.Size = New System.Drawing.Size(787, 424)
        Me.rpt_View.TabIndex = 10
        Me.rpt_View.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        Me.rpt_View.ViewTimeSelectionFormula = ""
        '
        'frmTrinhDoChuyenMonReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.rpt_View)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmTrinhDoChuyenMonReport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Báo cáo 09: THỐNG KÊ TRÌNH ĐỘ CHUYÊN MÔN - CHUYÊN NGÀNH THEO CHỨC DANH CÔNG VIỆC"
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
    Friend WithEvents txtGD As System.Windows.Forms.TextBox
    Friend WithEvents labGD As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents txtHCTC As System.Windows.Forms.TextBox
    Friend WithEvents labHCTC As System.Windows.Forms.Label
    Friend WithEvents txtLapBieu As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents rpt_BC09 As QLNS.BC09
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents labStatusProcess As System.Windows.Forms.Label
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dpkThoiDiemBC As System.Windows.Forms.DateTimePicker
    Private WithEvents rpt_View As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
