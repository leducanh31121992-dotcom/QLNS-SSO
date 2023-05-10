<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChonDS_KhenThuong
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChonDS_KhenThuong))
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.bntChoise = New System.Windows.Forms.Button()
        Me.bntCancelKT = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.treeCocau = New System.Windows.Forms.TreeView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstKhenThuong = New System.Windows.Forms.ListBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtNoiDung_KT = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cboDanhhieuTD = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.labTitle = New System.Windows.Forms.Label()
        Me.rdTT = New System.Windows.Forms.RadioButton()
        Me.rdCN = New System.Windows.Forms.RadioButton()
        Me.rdDV = New System.Windows.Forms.RadioButton()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtTenNhom = New System.Windows.Forms.TextBox()
        Me.Panel11.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.DarkGray
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel11.Controls.Add(Me.bntChoise)
        Me.Panel11.Controls.Add(Me.bntCancelKT)
        Me.Panel11.Controls.Add(Me.bntClose)
        Me.Panel11.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel11.Location = New System.Drawing.Point(0, 533)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(361, 28)
        Me.Panel11.TabIndex = 12
        '
        'bntChoise
        '
        Me.bntChoise.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntChoise.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntChoise.Location = New System.Drawing.Point(144, 0)
        Me.bntChoise.Name = "bntChoise"
        Me.bntChoise.Size = New System.Drawing.Size(71, 24)
        Me.bntChoise.TabIndex = 1
        Me.bntChoise.Text = "&Chọn"
        Me.bntChoise.UseVisualStyleBackColor = True
        '
        'bntCancelKT
        '
        Me.bntCancelKT.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntCancelKT.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntCancelKT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntCancelKT.Location = New System.Drawing.Point(215, 0)
        Me.bntCancelKT.Name = "bntCancelKT"
        Me.bntCancelKT.Size = New System.Drawing.Size(71, 24)
        Me.bntCancelKT.TabIndex = 2
        Me.bntCancelKT.Text = "&Bỏ qua"
        Me.bntCancelKT.UseVisualStyleBackColor = True
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntClose.Location = New System.Drawing.Point(286, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(71, 24)
        Me.bntClose.TabIndex = 3
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.treeCocau)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.lstKhenThuong)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(361, 533)
        Me.Panel1.TabIndex = 137
        '
        'treeCocau
        '
        Me.treeCocau.CheckBoxes = True
        Me.treeCocau.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeCocau.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.treeCocau.Location = New System.Drawing.Point(0, 122)
        Me.treeCocau.Name = "treeCocau"
        Me.treeCocau.Size = New System.Drawing.Size(361, 275)
        Me.treeCocau.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(0, 397)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(361, 20)
        Me.Label1.TabIndex = 127
        Me.Label1.Text = "Danh sách đã chọn"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lstKhenThuong
        '
        Me.lstKhenThuong.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lstKhenThuong.FormattingEnabled = True
        Me.lstKhenThuong.HorizontalScrollbar = True
        Me.lstKhenThuong.ItemHeight = 14
        Me.lstKhenThuong.Location = New System.Drawing.Point(0, 417)
        Me.lstKhenThuong.Name = "lstKhenThuong"
        Me.lstKhenThuong.Size = New System.Drawing.Size(361, 116)
        Me.lstKhenThuong.TabIndex = 5
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtNoiDung_KT)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.cboDanhhieuTD)
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.labTitle)
        Me.Panel2.Controls.Add(Me.rdTT)
        Me.Panel2.Controls.Add(Me.rdCN)
        Me.Panel2.Controls.Add(Me.rdDV)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.txtTenNhom)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(361, 122)
        Me.Panel2.TabIndex = 4
        '
        'txtNoiDung_KT
        '
        Me.txtNoiDung_KT.BackColor = System.Drawing.SystemColors.Window
        Me.txtNoiDung_KT.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNoiDung_KT.Location = New System.Drawing.Point(65, 47)
        Me.txtNoiDung_KT.Name = "txtNoiDung_KT"
        Me.txtNoiDung_KT.Size = New System.Drawing.Size(294, 22)
        Me.txtNoiDung_KT.TabIndex = 2
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(1, 50)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 14)
        Me.Label13.TabIndex = 128
        Me.Label13.Text = "Nội dung"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cboDanhhieuTD
        '
        Me.cboDanhhieuTD.BackColor = System.Drawing.SystemColors.Window
        Me.cboDanhhieuTD.DisplayMember = "Display"
        Me.cboDanhhieuTD.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDanhhieuTD.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDanhhieuTD.FormattingEnabled = True
        Me.cboDanhhieuTD.Location = New System.Drawing.Point(65, 24)
        Me.cboDanhhieuTD.Name = "cboDanhhieuTD"
        Me.cboDanhhieuTD.Size = New System.Drawing.Size(294, 22)
        Me.cboDanhhieuTD.TabIndex = 1
        Me.cboDanhhieuTD.ValueMember = "Value"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(1, 27)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(62, 14)
        Me.Label12.TabIndex = 126
        Me.Label12.Text = "Danh hiệu"
        '
        'labTitle
        '
        Me.labTitle.BackColor = System.Drawing.Color.DarkGray
        Me.labTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.labTitle.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labTitle.ForeColor = System.Drawing.Color.Maroon
        Me.labTitle.Location = New System.Drawing.Point(0, 0)
        Me.labTitle.Name = "labTitle"
        Me.labTitle.Size = New System.Drawing.Size(361, 21)
        Me.labTitle.TabIndex = 107
        Me.labTitle.Text = "Cho QĐ số "
        Me.labTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'rdTT
        '
        Me.rdTT.AutoSize = True
        Me.rdTT.Location = New System.Drawing.Point(260, 73)
        Me.rdTT.Name = "rdTT"
        Me.rdTT.Size = New System.Drawing.Size(104, 18)
        Me.rdTT.TabIndex = 6
        Me.rdTT.Text = "Tập thể nhóm"
        Me.rdTT.UseVisualStyleBackColor = True
        '
        'rdCN
        '
        Me.rdCN.AutoSize = True
        Me.rdCN.Location = New System.Drawing.Point(157, 73)
        Me.rdCN.Name = "rdCN"
        Me.rdCN.Size = New System.Drawing.Size(69, 18)
        Me.rdCN.TabIndex = 5
        Me.rdCN.Text = "Cá nhân"
        Me.rdCN.UseVisualStyleBackColor = True
        '
        'rdDV
        '
        Me.rdDV.AutoSize = True
        Me.rdDV.Checked = True
        Me.rdDV.Location = New System.Drawing.Point(3, 73)
        Me.rdDV.Name = "rdDV"
        Me.rdDV.Size = New System.Drawing.Size(127, 18)
        Me.rdDV.TabIndex = 3
        Me.rdDV.TabStop = True
        Me.rdDV.Text = "Đơn vị/ Phòng ban"
        Me.rdDV.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(1, 98)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 14)
        Me.Label8.TabIndex = 120
        Me.Label8.Text = "Tên nhóm"
        '
        'txtTenNhom
        '
        Me.txtTenNhom.BackColor = System.Drawing.SystemColors.Window
        Me.txtTenNhom.Enabled = False
        Me.txtTenNhom.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTenNhom.Location = New System.Drawing.Point(65, 95)
        Me.txtTenNhom.Name = "txtTenNhom"
        Me.txtTenNhom.Size = New System.Drawing.Size(294, 22)
        Me.txtTenNhom.TabIndex = 7
        '
        'frmChonDS_KhenThuong
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(361, 561)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel11)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmChonDS_KhenThuong"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chọn danh sách khen thưởng"
        Me.Panel11.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents bntChoise As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents treeCocau As System.Windows.Forms.TreeView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtTenNhom As System.Windows.Forms.TextBox
    Friend WithEvents rdTT As System.Windows.Forms.RadioButton
    Friend WithEvents rdCN As System.Windows.Forms.RadioButton
    Friend WithEvents rdDV As System.Windows.Forms.RadioButton
    Friend WithEvents labTitle As System.Windows.Forms.Label
    Friend WithEvents cboDanhhieuTD As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtNoiDung_KT As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lstKhenThuong As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents bntCancelKT As System.Windows.Forms.Button
End Class
