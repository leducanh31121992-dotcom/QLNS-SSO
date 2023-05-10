<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DMLuongForm
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btn_add = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btn_edit = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btn_delete = New System.Windows.Forms.Button()
        Me.lbl_header = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tv_main = New System.Windows.Forms.TreeView()
        Me.splitt_main = New System.Windows.Forms.Splitter()
        Me.gb_main = New System.Windows.Forms.GroupBox()
        Me.edt_maso = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.edt_tengoi = New System.Windows.Forms.TextBox()
        Me.pnl_label = New System.Windows.Forms.Panel()
        Me.lbl_code = New System.Windows.Forms.Label()
        Me.lbl_name = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.ckb_chonca = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.lbl_note = New System.Windows.Forms.Label()
        Me.btn_quayra = New System.Windows.Forms.Button()
        Me.dgv_main = New System.Windows.Forms.DataGridView()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.gb_main.SuspendLayout()
        Me.pnl_label.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.dgv_main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkGray
        Me.Panel2.Controls.Add(Me.btn_add)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.btn_edit)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.btn_delete)
        Me.Panel2.Controls.Add(Me.lbl_header)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(789, 24)
        Me.Panel2.TabIndex = 0
        '
        'btn_add
        '
        Me.btn_add.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_add.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_add.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_add.Location = New System.Drawing.Point(592, 0)
        Me.btn_add.Name = "btn_add"
        Me.btn_add.Size = New System.Drawing.Size(71, 24)
        Me.btn_add.TabIndex = 1
        Me.btn_add.Text = "Thêm &mới"
        Me.btn_add.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label11.Location = New System.Drawing.Point(663, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(2, 24)
        Me.Label11.TabIndex = 2
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_edit
        '
        Me.btn_edit.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_edit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_edit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_edit.Location = New System.Drawing.Point(665, 0)
        Me.btn_edit.Name = "btn_edit"
        Me.btn_edit.Size = New System.Drawing.Size(61, 24)
        Me.btn_edit.TabIndex = 3
        Me.btn_edit.Text = "&Sửa đổi"
        Me.btn_edit.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label13.Location = New System.Drawing.Point(726, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(2, 24)
        Me.Label13.TabIndex = 4
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_delete
        '
        Me.btn_delete.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_delete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_delete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_delete.Location = New System.Drawing.Point(728, 0)
        Me.btn_delete.Name = "btn_delete"
        Me.btn_delete.Size = New System.Drawing.Size(61, 24)
        Me.btn_delete.TabIndex = 5
        Me.btn_delete.Text = "&Xoá"
        Me.btn_delete.UseVisualStyleBackColor = True
        '
        'lbl_header
        '
        Me.lbl_header.BackColor = System.Drawing.Color.DarkGray
        Me.lbl_header.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_header.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_header.ForeColor = System.Drawing.Color.Maroon
        Me.lbl_header.Location = New System.Drawing.Point(0, 0)
        Me.lbl_header.Name = "lbl_header"
        Me.lbl_header.Size = New System.Drawing.Size(227, 24)
        Me.lbl_header.TabIndex = 0
        Me.lbl_header.Text = "HỆ THỐNG DANH MỤC LƯƠNG"
        Me.lbl_header.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label4.Location = New System.Drawing.Point(785, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(4, 537)
        Me.Label4.TabIndex = 2
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label3.Location = New System.Drawing.Point(0, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(4, 537)
        Me.Label3.TabIndex = 1
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label2.Location = New System.Drawing.Point(0, 561)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(789, 4)
        Me.Label2.TabIndex = 3
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tv_main
        '
        Me.tv_main.BackColor = System.Drawing.Color.White
        Me.tv_main.Dock = System.Windows.Forms.DockStyle.Left
        Me.tv_main.ForeColor = System.Drawing.Color.Black
        Me.tv_main.Location = New System.Drawing.Point(4, 24)
        Me.tv_main.Name = "tv_main"
        Me.tv_main.Size = New System.Drawing.Size(220, 537)
        Me.tv_main.TabIndex = 4
        '
        'splitt_main
        '
        Me.splitt_main.BackColor = System.Drawing.SystemColors.Control
        Me.splitt_main.Location = New System.Drawing.Point(224, 24)
        Me.splitt_main.Name = "splitt_main"
        Me.splitt_main.Size = New System.Drawing.Size(3, 537)
        Me.splitt_main.TabIndex = 5
        Me.splitt_main.TabStop = False
        '
        'gb_main
        '
        Me.gb_main.Controls.Add(Me.edt_maso)
        Me.gb_main.Controls.Add(Me.Label7)
        Me.gb_main.Controls.Add(Me.edt_tengoi)
        Me.gb_main.Controls.Add(Me.pnl_label)
        Me.gb_main.Dock = System.Windows.Forms.DockStyle.Top
        Me.gb_main.Location = New System.Drawing.Point(227, 24)
        Me.gb_main.Name = "gb_main"
        Me.gb_main.Size = New System.Drawing.Size(558, 68)
        Me.gb_main.TabIndex = 6
        Me.gb_main.TabStop = False
        Me.gb_main.Text = " Nội dung tìm kiếm "
        '
        'edt_maso
        '
        Me.edt_maso.BackColor = System.Drawing.Color.White
        Me.edt_maso.Dock = System.Windows.Forms.DockStyle.Fill
        Me.edt_maso.ForeColor = System.Drawing.Color.Black
        Me.edt_maso.Location = New System.Drawing.Point(114, 41)
        Me.edt_maso.Name = "edt_maso"
        Me.edt_maso.Size = New System.Drawing.Size(441, 22)
        Me.edt_maso.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label7.Location = New System.Drawing.Point(114, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(441, 1)
        Me.Label7.TabIndex = 2
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'edt_tengoi
        '
        Me.edt_tengoi.BackColor = System.Drawing.Color.White
        Me.edt_tengoi.Dock = System.Windows.Forms.DockStyle.Top
        Me.edt_tengoi.ForeColor = System.Drawing.Color.Black
        Me.edt_tengoi.Location = New System.Drawing.Point(114, 18)
        Me.edt_tengoi.Name = "edt_tengoi"
        Me.edt_tengoi.Size = New System.Drawing.Size(441, 22)
        Me.edt_tengoi.TabIndex = 1
        '
        'pnl_label
        '
        Me.pnl_label.Controls.Add(Me.lbl_code)
        Me.pnl_label.Controls.Add(Me.lbl_name)
        Me.pnl_label.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnl_label.Location = New System.Drawing.Point(3, 18)
        Me.pnl_label.Name = "pnl_label"
        Me.pnl_label.Size = New System.Drawing.Size(111, 47)
        Me.pnl_label.TabIndex = 0
        '
        'lbl_code
        '
        Me.lbl_code.BackColor = System.Drawing.Color.Transparent
        Me.lbl_code.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_code.Location = New System.Drawing.Point(0, 22)
        Me.lbl_code.Name = "lbl_code"
        Me.lbl_code.Size = New System.Drawing.Size(111, 23)
        Me.lbl_code.TabIndex = 1
        Me.lbl_code.Text = "  (*) Mã số"
        Me.lbl_code.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lbl_name
        '
        Me.lbl_name.BackColor = System.Drawing.Color.Transparent
        Me.lbl_name.Dock = System.Windows.Forms.DockStyle.Top
        Me.lbl_name.Location = New System.Drawing.Point(0, 0)
        Me.lbl_name.Name = "lbl_name"
        Me.lbl_name.Size = New System.Drawing.Size(111, 22)
        Me.lbl_name.TabIndex = 0
        Me.lbl_name.Text = "  (*) Tên gọi"
        Me.lbl_name.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.ckb_chonca)
        Me.Panel4.Controls.Add(Me.Label9)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(227, 92)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(558, 22)
        Me.Panel4.TabIndex = 7
        '
        'ckb_chonca
        '
        Me.ckb_chonca.Dock = System.Windows.Forms.DockStyle.Right
        Me.ckb_chonca.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ckb_chonca.ForeColor = System.Drawing.Color.Maroon
        Me.ckb_chonca.Location = New System.Drawing.Point(487, 0)
        Me.ckb_chonca.Name = "ckb_chonca"
        Me.ckb_chonca.Size = New System.Drawing.Size(71, 22)
        Me.ckb_chonca.TabIndex = 1
        Me.ckb_chonca.Text = "&Chọn cả"
        Me.ckb_chonca.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Maroon
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(232, 22)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = " KẾT QUẢ TÌM KIẾM"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.lbl_note)
        Me.Panel3.Controls.Add(Me.btn_quayra)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(227, 536)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(558, 25)
        Me.Panel3.TabIndex = 10
        '
        'lbl_note
        '
        Me.lbl_note.BackColor = System.Drawing.Color.DarkGray
        Me.lbl_note.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_note.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_note.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.lbl_note.Location = New System.Drawing.Point(0, 0)
        Me.lbl_note.Name = "lbl_note"
        Me.lbl_note.Size = New System.Drawing.Size(241, 21)
        Me.lbl_note.TabIndex = 8
        Me.lbl_note.Text = " Các đối tượng có ký tự * hỗ trợ tìm kiếm"
        Me.lbl_note.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_quayra
        '
        Me.btn_quayra.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_quayra.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_quayra.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_quayra.Location = New System.Drawing.Point(484, 0)
        Me.btn_quayra.Name = "btn_quayra"
        Me.btn_quayra.Size = New System.Drawing.Size(70, 21)
        Me.btn_quayra.TabIndex = 7
        Me.btn_quayra.Text = "&Quay ra"
        Me.btn_quayra.UseVisualStyleBackColor = True
        '
        'dgv_main
        '
        Me.dgv_main.AllowUserToAddRows = False
        Me.dgv_main.AllowUserToDeleteRows = False
        Me.dgv_main.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(254, Byte), Integer))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(22, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(124, Byte), Integer))
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(254, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(22, Byte), Integer), CType(CType(56, Byte), Integer), CType(CType(124, Byte), Integer))
        Me.dgv_main.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgv_main.BackgroundColor = System.Drawing.Color.White
        Me.dgv_main.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgv_main.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv_main.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgv_main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(237, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(250, Byte), Integer), CType(CType(237, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgv_main.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgv_main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_main.Location = New System.Drawing.Point(227, 114)
        Me.dgv_main.Name = "dgv_main"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(178, Byte), Integer), CType(CType(219, Byte), Integer), CType(CType(255, Byte), Integer))
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv_main.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgv_main.RowHeadersVisible = False
        Me.dgv_main.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.dgv_main.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.SystemColors.Highlight
        Me.dgv_main.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White
        Me.dgv_main.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgv_main.Size = New System.Drawing.Size(558, 420)
        Me.dgv_main.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label8.Location = New System.Drawing.Point(227, 534)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(558, 2)
        Me.Label8.TabIndex = 9
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DMLuongForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 565)
        Me.Controls.Add(Me.dgv_main)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.gb_main)
        Me.Controls.Add(Me.splitt_main)
        Me.Controls.Add(Me.tv_main)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "DMLuongForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự"
        Me.Panel2.ResumeLayout(False)
        Me.gb_main.ResumeLayout(False)
        Me.gb_main.PerformLayout()
        Me.pnl_label.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.dgv_main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btn_add As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btn_edit As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btn_delete As System.Windows.Forms.Button
    Public WithEvents lbl_header As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tv_main As System.Windows.Forms.TreeView
    Friend WithEvents splitt_main As System.Windows.Forms.Splitter
    Friend WithEvents gb_main As System.Windows.Forms.GroupBox
    Friend WithEvents edt_maso As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents edt_tengoi As System.Windows.Forms.TextBox
    Friend WithEvents pnl_label As System.Windows.Forms.Panel
    Friend WithEvents lbl_code As System.Windows.Forms.Label
    Friend WithEvents lbl_name As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents ckb_chonca As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents lbl_note As System.Windows.Forms.Label
    Friend WithEvents btn_quayra As System.Windows.Forms.Button
    Friend WithEvents dgv_main As System.Windows.Forms.DataGridView
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
