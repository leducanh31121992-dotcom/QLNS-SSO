<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDS_TDKT
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDS_TDKT))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rdDeNghi = New System.Windows.Forms.RadioButton()
        Me.rdKhenThuong = New System.Windows.Forms.RadioButton()
        Me.dtpkNam = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rdDX = New System.Windows.Forms.RadioButton()
        Me.rdDK = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.rdTT = New System.Windows.Forms.RadioButton()
        Me.rdCN = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.bntCreate = New System.Windows.Forms.Button()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.labAlert = New System.Windows.Forms.Label()
        Me.bntUpdate = New System.Windows.Forms.Button()
        Me.bntPrint = New System.Windows.Forms.Button()
        Me.bntCloseKT = New System.Windows.Forms.Button()
        Me.gridResult = New System.Windows.Forms.DataGridView()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel5.SuspendLayout()
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel8.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(785, 73)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Panel6)
        Me.Panel2.Controls.Add(Me.Panel4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(695, 73)
        Me.Panel2.TabIndex = 2
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.cboDonVi)
        Me.Panel6.Controls.Add(Me.Panel7)
        Me.Panel6.Controls.Add(Me.Label40)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel6.Location = New System.Drawing.Point(0, 44)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(695, 29)
        Me.Panel6.TabIndex = 1
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.SystemColors.Window
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(55, 0)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(632, 22)
        Me.cboDonVi.TabIndex = 140
        Me.cboDonVi.ValueMember = "Value"
        '
        'Panel7
        '
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel7.Location = New System.Drawing.Point(687, 0)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(8, 29)
        Me.Panel7.TabIndex = 141
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(0, 0)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(55, 29)
        Me.Label40.TabIndex = 138
        Me.Label40.Text = "Đơn vị "
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Label3)
        Me.Panel4.Controls.Add(Me.GroupBox2)
        Me.Panel4.Controls.Add(Me.dtpkNam)
        Me.Panel4.Controls.Add(Me.GroupBox1)
        Me.Panel4.Controls.Add(Me.GroupBox3)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(695, 44)
        Me.Panel4.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(597, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 14)
        Me.Label3.TabIndex = 135
        Me.Label3.Text = "Năm"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rdDeNghi)
        Me.GroupBox2.Controls.Add(Me.rdKhenThuong)
        Me.GroupBox2.Location = New System.Drawing.Point(2, 3)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox2.Size = New System.Drawing.Size(262, 35)
        Me.GroupBox2.TabIndex = 136
        Me.GroupBox2.TabStop = False
        '
        'rdDeNghi
        '
        Me.rdDeNghi.AutoSize = True
        Me.rdDeNghi.Location = New System.Drawing.Point(111, 12)
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
        Me.rdKhenThuong.Location = New System.Drawing.Point(12, 12)
        Me.rdKhenThuong.Name = "rdKhenThuong"
        Me.rdKhenThuong.Size = New System.Drawing.Size(98, 18)
        Me.rdKhenThuong.TabIndex = 1
        Me.rdKhenThuong.TabStop = True
        Me.rdKhenThuong.Text = "Khen thưởng"
        Me.rdKhenThuong.UseVisualStyleBackColor = True
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
        Me.dtpkNam.Location = New System.Drawing.Point(628, 13)
        Me.dtpkNam.MaxDate = New Date(2100, 12, 31, 0, 0, 0, 0)
        Me.dtpkNam.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpkNam.Name = "dtpkNam"
        Me.dtpkNam.ShowUpDown = True
        Me.dtpkNam.Size = New System.Drawing.Size(59, 22)
        Me.dtpkNam.TabIndex = 134
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdDX)
        Me.GroupBox1.Controls.Add(Me.rdDK)
        Me.GroupBox1.Location = New System.Drawing.Point(269, 3)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.Size = New System.Drawing.Size(156, 35)
        Me.GroupBox1.TabIndex = 137
        Me.GroupBox1.TabStop = False
        '
        'rdDX
        '
        Me.rdDX.AutoSize = True
        Me.rdDX.Location = New System.Drawing.Point(78, 12)
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
        Me.rdDK.Location = New System.Drawing.Point(12, 12)
        Me.rdDK.Name = "rdDK"
        Me.rdDK.Size = New System.Drawing.Size(65, 18)
        Me.rdDK.TabIndex = 1
        Me.rdDK.TabStop = True
        Me.rdDK.Text = "Định kỳ"
        Me.rdDK.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rdTT)
        Me.GroupBox3.Controls.Add(Me.rdCN)
        Me.GroupBox3.Location = New System.Drawing.Point(430, 3)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox3.Size = New System.Drawing.Size(158, 35)
        Me.GroupBox3.TabIndex = 137
        Me.GroupBox3.TabStop = False
        '
        'rdTT
        '
        Me.rdTT.AutoSize = True
        Me.rdTT.Location = New System.Drawing.Point(83, 12)
        Me.rdTT.Name = "rdTT"
        Me.rdTT.Size = New System.Drawing.Size(69, 18)
        Me.rdTT.TabIndex = 2
        Me.rdTT.Text = "Tập thể"
        Me.rdTT.UseVisualStyleBackColor = True
        '
        'rdCN
        '
        Me.rdCN.AutoSize = True
        Me.rdCN.Checked = True
        Me.rdCN.Location = New System.Drawing.Point(12, 12)
        Me.rdCN.Name = "rdCN"
        Me.rdCN.Size = New System.Drawing.Size(69, 18)
        Me.rdCN.TabIndex = 1
        Me.rdCN.TabStop = True
        Me.rdCN.Text = "Cá nhân"
        Me.rdCN.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.bntCreate)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel3.Location = New System.Drawing.Point(695, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(90, 73)
        Me.Panel3.TabIndex = 1
        '
        'bntCreate
        '
        Me.bntCreate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCreate.Location = New System.Drawing.Point(12, 16)
        Me.bntCreate.Name = "bntCreate"
        Me.bntCreate.Size = New System.Drawing.Size(66, 41)
        Me.bntCreate.TabIndex = 2
        Me.bntCreate.Text = "&Hiển thị"
        Me.bntCreate.UseVisualStyleBackColor = True
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.DarkGray
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel5.Controls.Add(Me.labAlert)
        Me.Panel5.Controls.Add(Me.bntUpdate)
        Me.Panel5.Controls.Add(Me.bntPrint)
        Me.Panel5.Controls.Add(Me.bntCloseKT)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel5.Location = New System.Drawing.Point(0, 535)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(785, 26)
        Me.Panel5.TabIndex = 16
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
        'bntUpdate
        '
        Me.bntUpdate.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntUpdate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntUpdate.Location = New System.Drawing.Point(480, 0)
        Me.bntUpdate.Name = "bntUpdate"
        Me.bntUpdate.Size = New System.Drawing.Size(120, 22)
        Me.bntUpdate.TabIndex = 1
        Me.bntUpdate.Text = "&Duyệt danh sách"
        Me.bntUpdate.UseVisualStyleBackColor = True
        '
        'bntPrint
        '
        Me.bntPrint.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntPrint.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntPrint.Location = New System.Drawing.Point(600, 0)
        Me.bntPrint.Name = "bntPrint"
        Me.bntPrint.Size = New System.Drawing.Size(120, 22)
        Me.bntPrint.TabIndex = 3
        Me.bntPrint.Text = "&In danh sách"
        Me.bntPrint.UseVisualStyleBackColor = True
        '
        'bntCloseKT
        '
        Me.bntCloseKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCloseKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCloseKT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCloseKT.Location = New System.Drawing.Point(720, 0)
        Me.bntCloseKT.Name = "bntCloseKT"
        Me.bntCloseKT.Size = New System.Drawing.Size(61, 22)
        Me.bntCloseKT.TabIndex = 5
        Me.bntCloseKT.Text = "&Quay ra"
        Me.bntCloseKT.UseVisualStyleBackColor = True
        '
        'gridResult
        '
        Me.gridResult.AllowUserToAddRows = False
        Me.gridResult.AllowUserToDeleteRows = False
        Me.gridResult.AllowUserToResizeRows = False
        Me.gridResult.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.gridResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridResult.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
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
        Me.gridResult.Location = New System.Drawing.Point(0, 73)
        Me.gridResult.MultiSelect = False
        Me.gridResult.Name = "gridResult"
        Me.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridResult.Size = New System.Drawing.Size(785, 420)
        Me.gridResult.TabIndex = 107
        '
        'Panel8
        '
        Me.Panel8.Controls.Add(Me.Label1)
        Me.Panel8.Controls.Add(Me.Label5)
        Me.Panel8.Controls.Add(Me.Label4)
        Me.Panel8.Controls.Add(Me.Label2)
        Me.Panel8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel8.Location = New System.Drawing.Point(0, 493)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(785, 42)
        Me.Panel8.TabIndex = 108
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(358, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "x    : Đã được duyệt và QĐ khen thưởng đã nhập vào hệ thống"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(461, 25)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(260, 14)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "?   : Đề nghị khen thưởng không đủ điều kiện"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(462, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(224, 14)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "!    : Đề nghị khen thưởng đủ điều kiện"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(370, 14)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "x(!): Đã được duyệt và QĐ khen thưởng chưa nhập vào hệ thống"
        '
        'frmDS_TDKT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(785, 561)
        Me.Controls.Add(Me.gridResult)
        Me.Controls.Add(Me.Panel8)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmDS_TDKT"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: DANH SÁCH KHEN THƯỞNG CÁ NHÂN/TẬP THỂ HÀNG NĂM"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents labAlert As System.Windows.Forms.Label
    Friend WithEvents bntUpdate As System.Windows.Forms.Button
    Friend WithEvents bntPrint As System.Windows.Forms.Button
    Friend WithEvents bntCloseKT As System.Windows.Forms.Button
    Friend WithEvents gridResult As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpkNam As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rdDeNghi As System.Windows.Forms.RadioButton
    Friend WithEvents rdKhenThuong As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdDX As System.Windows.Forms.RadioButton
    Friend WithEvents rdDK As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents rdTT As System.Windows.Forms.RadioButton
    Friend WithEvents rdCN As System.Windows.Forms.RadioButton
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents bntCreate As System.Windows.Forms.Button
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
