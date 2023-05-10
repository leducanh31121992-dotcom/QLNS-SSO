<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHT_ThanhVien
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btn_grmember = New System.Windows.Forms.Button()
        Me.btn_accept = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btn_cancel = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btn_back = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.tv_main = New System.Windows.Forms.TreeView()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.splitt_main = New System.Windows.Forms.Splitter()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.clb_nhomquyen = New System.Windows.Forms.CheckedListBox()
        Me.ckb_nhquyen = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.clb_tapquyen = New System.Windows.Forms.CheckedListBox()
        Me.ckb_tapquyen = New System.Windows.Forms.CheckBox()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGray
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.ForeColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(789, 24)
        Me.Panel1.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label9.Location = New System.Drawing.Point(783, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(3, 24)
        Me.Label9.TabIndex = 2
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label10
        '
        Me.Label10.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label10.Location = New System.Drawing.Point(786, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(3, 24)
        Me.Label10.TabIndex = 4
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.DarkGray
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Maroon
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(181, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = " QUYỀN SỬ DỤNG"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.btn_grmember)
        Me.Panel3.Controls.Add(Me.btn_accept)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.btn_cancel)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.btn_back)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 538)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(789, 27)
        Me.Panel3.TabIndex = 1
        '
        'btn_grmember
        '
        Me.btn_grmember.Dock = System.Windows.Forms.DockStyle.Left
        Me.btn_grmember.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_grmember.Location = New System.Drawing.Point(0, 0)
        Me.btn_grmember.Name = "btn_grmember"
        Me.btn_grmember.Size = New System.Drawing.Size(115, 23)
        Me.btn_grmember.TabIndex = 0
        Me.btn_grmember.Text = "&Nhóm người dùng"
        Me.btn_grmember.UseVisualStyleBackColor = True
        '
        'btn_accept
        '
        Me.btn_accept.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_accept.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_accept.Location = New System.Drawing.Point(554, 0)
        Me.btn_accept.Name = "btn_accept"
        Me.btn_accept.Size = New System.Drawing.Size(103, 23)
        Me.btn_accept.TabIndex = 1
        Me.btn_accept.Text = "&Thiết lập quyền"
        Me.btn_accept.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label8.Location = New System.Drawing.Point(657, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(3, 23)
        Me.Label8.TabIndex = 2
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_cancel
        '
        Me.btn_cancel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_cancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_cancel.Location = New System.Drawing.Point(660, 0)
        Me.btn_cancel.Name = "btn_cancel"
        Me.btn_cancel.Size = New System.Drawing.Size(61, 23)
        Me.btn_cancel.TabIndex = 3
        Me.btn_cancel.Text = "&Bỏ qua"
        Me.btn_cancel.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label7.Location = New System.Drawing.Point(721, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(3, 23)
        Me.Label7.TabIndex = 4
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_back
        '
        Me.btn_back.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_back.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_back.Location = New System.Drawing.Point(724, 0)
        Me.btn_back.Name = "btn_back"
        Me.btn_back.Size = New System.Drawing.Size(61, 23)
        Me.btn_back.TabIndex = 5
        Me.btn_back.Text = "Quay &ra"
        Me.btn_back.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.tv_main)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel2.Location = New System.Drawing.Point(0, 24)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(221, 514)
        Me.Panel2.TabIndex = 2
        '
        'tv_main
        '
        Me.tv_main.BackColor = System.Drawing.Color.White
        Me.tv_main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tv_main.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.tv_main.Location = New System.Drawing.Point(0, 22)
        Me.tv_main.Name = "tv_main"
        Me.tv_main.Size = New System.Drawing.Size(221, 492)
        Me.tv_main.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(221, 22)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Nhóm người dùng"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'splitt_main
        '
        Me.splitt_main.BackColor = System.Drawing.Color.DarkGray
        Me.splitt_main.Location = New System.Drawing.Point(221, 24)
        Me.splitt_main.Name = "splitt_main"
        Me.splitt_main.Size = New System.Drawing.Size(3, 514)
        Me.splitt_main.TabIndex = 3
        Me.splitt_main.TabStop = False
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.DarkGray
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Maroon
        Me.Label5.Location = New System.Drawing.Point(224, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(565, 23)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = " CHI TIẾT THÔNG TIN NHÓM QUYỀN"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.clb_nhomquyen)
        Me.GroupBox1.Controls.Add(Me.ckb_nhquyen)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(224, 47)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(265, 491)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Nhóm quyền sử dụng"
        '
        'clb_nhomquyen
        '
        Me.clb_nhomquyen.BackColor = System.Drawing.Color.White
        Me.clb_nhomquyen.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clb_nhomquyen.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.clb_nhomquyen.FormattingEnabled = True
        Me.clb_nhomquyen.Location = New System.Drawing.Point(3, 18)
        Me.clb_nhomquyen.Name = "clb_nhomquyen"
        Me.clb_nhomquyen.Size = New System.Drawing.Size(259, 449)
        Me.clb_nhomquyen.TabIndex = 0
        '
        'ckb_nhquyen
        '
        Me.ckb_nhquyen.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ckb_nhquyen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ckb_nhquyen.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ckb_nhquyen.Location = New System.Drawing.Point(3, 467)
        Me.ckb_nhquyen.Name = "ckb_nhquyen"
        Me.ckb_nhquyen.Size = New System.Drawing.Size(259, 21)
        Me.ckb_nhquyen.TabIndex = 1
        Me.ckb_nhquyen.Text = "Chọn toàn bộ"
        Me.ckb_nhquyen.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.clb_tapquyen)
        Me.GroupBox2.Controls.Add(Me.ckb_tapquyen)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(492, 47)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(297, 491)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Quyền sử dụng chi tiết"
        '
        'clb_tapquyen
        '
        Me.clb_tapquyen.BackColor = System.Drawing.Color.White
        Me.clb_tapquyen.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clb_tapquyen.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.clb_tapquyen.FormattingEnabled = True
        Me.clb_tapquyen.Location = New System.Drawing.Point(3, 18)
        Me.clb_tapquyen.Name = "clb_tapquyen"
        Me.clb_tapquyen.Size = New System.Drawing.Size(291, 449)
        Me.clb_tapquyen.TabIndex = 0
        '
        'ckb_tapquyen
        '
        Me.ckb_tapquyen.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ckb_tapquyen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ckb_tapquyen.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ckb_tapquyen.Location = New System.Drawing.Point(3, 467)
        Me.ckb_tapquyen.Name = "ckb_tapquyen"
        Me.ckb_tapquyen.Size = New System.Drawing.Size(291, 21)
        Me.ckb_tapquyen.TabIndex = 1
        Me.ckb_tapquyen.Text = "Chọn toàn bộ"
        Me.ckb_tapquyen.UseVisualStyleBackColor = True
        '
        'Splitter1
        '
        Me.Splitter1.BackColor = System.Drawing.Color.DarkGray
        Me.Splitter1.Location = New System.Drawing.Point(489, 47)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 491)
        Me.Splitter1.TabIndex = 8
        Me.Splitter1.TabStop = False
        '
        'frmHT_ThanhVien
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 565)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.splitt_main)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MinimizeBox = False
        Me.Name = "frmHT_ThanhVien"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Thiết lập danh sách quyền sử dụng"
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btn_grmember As System.Windows.Forms.Button
    Friend WithEvents btn_accept As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btn_cancel As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btn_back As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents tv_main As System.Windows.Forms.TreeView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents splitt_main As System.Windows.Forms.Splitter
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents clb_nhomquyen As System.Windows.Forms.CheckedListBox
    Friend WithEvents ckb_nhquyen As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents clb_tapquyen As System.Windows.Forms.CheckedListBox
    Friend WithEvents ckb_tapquyen As System.Windows.Forms.CheckBox
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
End Class
