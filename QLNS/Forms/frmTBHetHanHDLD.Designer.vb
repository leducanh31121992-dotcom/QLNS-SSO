<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTBHetHanHDLD
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
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.labStatusProcess = New System.Windows.Forms.Label()
        Me.cmdReport = New System.Windows.Forms.Button()
        Me.bntExportFile = New System.Windows.Forms.Button()
        Me.bntRefresh = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dpkDateline = New System.Windows.Forms.DateTimePicker()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.grpResult = New System.Windows.Forms.GroupBox()
        Me.gridResult = New System.Windows.Forms.DataGridView()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.grpResult.SuspendLayout()
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.Controls.Add(Me.labStatusProcess)
        Me.Panel3.Controls.Add(Me.cmdReport)
        Me.Panel3.Controls.Add(Me.bntExportFile)
        Me.Panel3.Controls.Add(Me.bntRefresh)
        Me.Panel3.Controls.Add(Me.cmdClose)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(787, 24)
        Me.Panel3.TabIndex = 3
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
        'cmdReport
        '
        Me.cmdReport.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReport.Location = New System.Drawing.Point(397, 0)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(100, 24)
        Me.cmdReport.TabIndex = 1
        Me.cmdReport.Text = "&Tra cứu"
        Me.cmdReport.UseVisualStyleBackColor = True
        '
        'bntExportFile
        '
        Me.bntExportFile.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntExportFile.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntExportFile.Location = New System.Drawing.Point(497, 0)
        Me.bntExportFile.Name = "bntExportFile"
        Me.bntExportFile.Size = New System.Drawing.Size(95, 24)
        Me.bntExportFile.TabIndex = 22
        Me.bntExportFile.Text = "Xuất &file Excel"
        Me.bntExportFile.UseVisualStyleBackColor = True
        '
        'bntRefresh
        '
        Me.bntRefresh.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntRefresh.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntRefresh.Location = New System.Drawing.Point(592, 0)
        Me.bntRefresh.Name = "bntRefresh"
        Me.bntRefresh.Size = New System.Drawing.Size(95, 24)
        Me.bntRefresh.TabIndex = 23
        Me.bntRefresh.Text = "Tra cứu &mới"
        Me.bntRefresh.UseVisualStyleBackColor = True
        '
        'cmdClose
        '
        Me.cmdClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClose.Location = New System.Drawing.Point(687, 0)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(100, 24)
        Me.cmdClose.TabIndex = 2
        Me.cmdClose.Text = "&Quay ra"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 24)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(787, 74)
        Me.Panel1.TabIndex = 7
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dpkDateline)
        Me.GroupBox1.Controls.Add(Me.Label40)
        Me.GroupBox1.Controls.Add(Me.cboDonVi)
        Me.GroupBox1.Controls.Add(Me.cbAll)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(787, 70)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Khai báo tham số"
        '
        'dpkDateline
        '
        Me.dpkDateline.CustomFormat = "dd/MM/yyyy"
        Me.dpkDateline.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkDateline.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkDateline.Location = New System.Drawing.Point(146, 40)
        Me.dpkDateline.Name = "dpkDateline"
        Me.dpkDateline.Size = New System.Drawing.Size(105, 22)
        Me.dpkDateline.TabIndex = 7
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(-1, 13)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(146, 24)
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
        'cbAll
        '
        Me.cbAll.AutoSize = True
        Me.cbAll.Location = New System.Drawing.Point(627, 19)
        Me.cbAll.Name = "cbAll"
        Me.cbAll.Size = New System.Drawing.Size(145, 18)
        Me.cbAll.TabIndex = 2
        Me.cbAll.Text = "Thống kê toàn đơn vị"
        Me.cbAll.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(52, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 24)
        Me.Label2.TabIndex = 124
        Me.Label2.Text = "Tính đến ngày"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpResult
        '
        Me.grpResult.Controls.Add(Me.gridResult)
        Me.grpResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpResult.Location = New System.Drawing.Point(0, 98)
        Me.grpResult.Name = "grpResult"
        Me.grpResult.Size = New System.Drawing.Size(787, 465)
        Me.grpResult.TabIndex = 19
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
        Me.gridResult.Location = New System.Drawing.Point(3, 18)
        Me.gridResult.Name = "gridResult"
        Me.gridResult.ReadOnly = True
        Me.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridResult.Size = New System.Drawing.Size(781, 444)
        Me.gridResult.TabIndex = 105
        '
        'frmTBHetHanHDLD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.grpResult)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Name = "frmTBHetHanHDLD"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Quản lý nhân sự: DANH SÁCH HẾT HẠN HỢP ĐỒNG LAO ĐỘNG"
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grpResult.ResumeLayout(False)
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents labStatusProcess As System.Windows.Forms.Label
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dpkDateline As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents grpResult As System.Windows.Forms.GroupBox
    Friend WithEvents gridResult As System.Windows.Forms.DataGridView
    Friend WithEvents bntExportFile As System.Windows.Forms.Button
    Friend WithEvents bntRefresh As System.Windows.Forms.Button
End Class
