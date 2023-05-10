<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BaoCaoForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BaoCaoForm))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lbl_waitting = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btn_xuatexcel = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btn_quayra = New System.Windows.Forms.Button()
        Me.gb_main = New System.Windows.Forms.GroupBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.pnl_tonghop_cn_hay_pgd = New System.Windows.Forms.Panel()
        Me.rb_danhsach_chitiet = New System.Windows.Forms.RadioButton()
        Me.rb_tonghop_pgd = New System.Windows.Forms.RadioButton()
        Me.rb_tonghop_cn = New System.Windows.Forms.RadioButton()
        Me.pnl_opt_tonghop_chitiet = New System.Windows.Forms.Panel()
        Me.rb_chitiet = New System.Windows.Forms.RadioButton()
        Me.rb_tonghop = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.cb_donvi = New System.Windows.Forms.ComboBox()
        Me.lbl_tungay = New System.Windows.Forms.Label()
        Me.dtpk_tungay = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lbl_ngaybc = New System.Windows.Forms.Label()
        Me.dtpk_ngaybc = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cb_baocao = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.gb_main.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.pnl_tonghop_cn_hay_pgd.SuspendLayout()
        Me.pnl_opt_tonghop_chitiet.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.Panel2.Controls.Add(Me.lbl_waitting)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.btn_xuatexcel)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.btn_quayra)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(878, 25)
        Me.Panel2.TabIndex = 1
        '
        'lbl_waitting
        '
        Me.lbl_waitting.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbl_waitting.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_waitting.ForeColor = System.Drawing.Color.Maroon
        Me.lbl_waitting.Location = New System.Drawing.Point(10, 0)
        Me.lbl_waitting.Name = "lbl_waitting"
        Me.lbl_waitting.Size = New System.Drawing.Size(522, 25)
        Me.lbl_waitting.TabIndex = 1
        Me.lbl_waitting.Text = "Waitting ..."
        Me.lbl_waitting.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label5.Location = New System.Drawing.Point(0, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(10, 25)
        Me.Label5.TabIndex = 0
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btn_xuatexcel
        '
        Me.btn_xuatexcel.BackColor = System.Drawing.Color.Transparent
        Me.btn_xuatexcel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_xuatexcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_xuatexcel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_xuatexcel.ForeColor = System.Drawing.Color.AliceBlue
        Me.btn_xuatexcel.Image = Global.QLNS.My.Resources.Resources.export
        Me.btn_xuatexcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_xuatexcel.Location = New System.Drawing.Point(683, 0)
        Me.btn_xuatexcel.Name = "btn_xuatexcel"
        Me.btn_xuatexcel.Size = New System.Drawing.Size(108, 25)
        Me.btn_xuatexcel.TabIndex = 2
        Me.btn_xuatexcel.Text = "&Xuất excel"
        Me.btn_xuatexcel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_xuatexcel.UseVisualStyleBackColor = False
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label13.Location = New System.Drawing.Point(791, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(2, 25)
        Me.Label13.TabIndex = 3
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_quayra
        '
        Me.btn_quayra.BackColor = System.Drawing.Color.Transparent
        Me.btn_quayra.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_quayra.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_quayra.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_quayra.ForeColor = System.Drawing.Color.AliceBlue
        Me.btn_quayra.Image = Global.QLNS.My.Resources.Resources.icon_back
        Me.btn_quayra.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_quayra.Location = New System.Drawing.Point(793, 0)
        Me.btn_quayra.Name = "btn_quayra"
        Me.btn_quayra.Size = New System.Drawing.Size(85, 25)
        Me.btn_quayra.TabIndex = 4
        Me.btn_quayra.Text = "&Quay ra"
        Me.btn_quayra.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_quayra.UseVisualStyleBackColor = False
        '
        'gb_main
        '
        Me.gb_main.Controls.Add(Me.Panel4)
        Me.gb_main.Controls.Add(Me.Label3)
        Me.gb_main.Controls.Add(Me.Panel3)
        Me.gb_main.Controls.Add(Me.Label6)
        Me.gb_main.Controls.Add(Me.Panel1)
        Me.gb_main.Controls.Add(Me.Label1)
        Me.gb_main.Dock = System.Windows.Forms.DockStyle.Top
        Me.gb_main.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gb_main.Location = New System.Drawing.Point(0, 25)
        Me.gb_main.Name = "gb_main"
        Me.gb_main.Size = New System.Drawing.Size(878, 103)
        Me.gb_main.TabIndex = 2
        Me.gb_main.TabStop = False
        Me.gb_main.Text = " Khai báo tham số "
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.pnl_tonghop_cn_hay_pgd)
        Me.Panel4.Controls.Add(Me.pnl_opt_tonghop_chitiet)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(3, 73)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(872, 22)
        Me.Panel4.TabIndex = 5
        '
        'pnl_tonghop_cn_hay_pgd
        '
        Me.pnl_tonghop_cn_hay_pgd.BackColor = System.Drawing.Color.White
        Me.pnl_tonghop_cn_hay_pgd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_tonghop_cn_hay_pgd.Controls.Add(Me.rb_danhsach_chitiet)
        Me.pnl_tonghop_cn_hay_pgd.Controls.Add(Me.rb_tonghop_pgd)
        Me.pnl_tonghop_cn_hay_pgd.Controls.Add(Me.rb_tonghop_cn)
        Me.pnl_tonghop_cn_hay_pgd.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnl_tonghop_cn_hay_pgd.Location = New System.Drawing.Point(104, 0)
        Me.pnl_tonghop_cn_hay_pgd.Name = "pnl_tonghop_cn_hay_pgd"
        Me.pnl_tonghop_cn_hay_pgd.Size = New System.Drawing.Size(513, 22)
        Me.pnl_tonghop_cn_hay_pgd.TabIndex = 5
        '
        'rb_danhsach_chitiet
        '
        Me.rb_danhsach_chitiet.Dock = System.Windows.Forms.DockStyle.Left
        Me.rb_danhsach_chitiet.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb_danhsach_chitiet.Location = New System.Drawing.Point(340, 0)
        Me.rb_danhsach_chitiet.Name = "rb_danhsach_chitiet"
        Me.rb_danhsach_chitiet.Size = New System.Drawing.Size(150, 18)
        Me.rb_danhsach_chitiet.TabIndex = 2
        Me.rb_danhsach_chitiet.Text = "Chi tiết danh sách"
        Me.rb_danhsach_chitiet.UseVisualStyleBackColor = True
        '
        'rb_tonghop_pgd
        '
        Me.rb_tonghop_pgd.Dock = System.Windows.Forms.DockStyle.Left
        Me.rb_tonghop_pgd.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb_tonghop_pgd.Location = New System.Drawing.Point(190, 0)
        Me.rb_tonghop_pgd.Name = "rb_tonghop_pgd"
        Me.rb_tonghop_pgd.Size = New System.Drawing.Size(150, 18)
        Me.rb_tonghop_pgd.TabIndex = 1
        Me.rb_tonghop_pgd.Text = "Tổng hợp đến PGD"
        Me.rb_tonghop_pgd.UseVisualStyleBackColor = True
        '
        'rb_tonghop_cn
        '
        Me.rb_tonghop_cn.Checked = True
        Me.rb_tonghop_cn.Dock = System.Windows.Forms.DockStyle.Left
        Me.rb_tonghop_cn.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb_tonghop_cn.Location = New System.Drawing.Point(0, 0)
        Me.rb_tonghop_cn.Name = "rb_tonghop_cn"
        Me.rb_tonghop_cn.Size = New System.Drawing.Size(190, 18)
        Me.rb_tonghop_cn.TabIndex = 0
        Me.rb_tonghop_cn.TabStop = True
        Me.rb_tonghop_cn.Text = "Tổng hợp đến chi nhánh"
        Me.rb_tonghop_cn.UseVisualStyleBackColor = True
        '
        'pnl_opt_tonghop_chitiet
        '
        Me.pnl_opt_tonghop_chitiet.BackColor = System.Drawing.Color.White
        Me.pnl_opt_tonghop_chitiet.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_opt_tonghop_chitiet.Controls.Add(Me.rb_chitiet)
        Me.pnl_opt_tonghop_chitiet.Controls.Add(Me.rb_tonghop)
        Me.pnl_opt_tonghop_chitiet.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnl_opt_tonghop_chitiet.Location = New System.Drawing.Point(617, 0)
        Me.pnl_opt_tonghop_chitiet.Name = "pnl_opt_tonghop_chitiet"
        Me.pnl_opt_tonghop_chitiet.Size = New System.Drawing.Size(255, 22)
        Me.pnl_opt_tonghop_chitiet.TabIndex = 4
        '
        'rb_chitiet
        '
        Me.rb_chitiet.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.rb_chitiet.Dock = System.Windows.Forms.DockStyle.Right
        Me.rb_chitiet.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb_chitiet.Location = New System.Drawing.Point(151, 0)
        Me.rb_chitiet.Name = "rb_chitiet"
        Me.rb_chitiet.Size = New System.Drawing.Size(100, 18)
        Me.rb_chitiet.TabIndex = 1
        Me.rb_chitiet.TabStop = True
        Me.rb_chitiet.Text = "Theo vùng"
        Me.rb_chitiet.UseVisualStyleBackColor = True
        '
        'rb_tonghop
        '
        Me.rb_tonghop.Dock = System.Windows.Forms.DockStyle.Left
        Me.rb_tonghop.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rb_tonghop.Location = New System.Drawing.Point(0, 0)
        Me.rb_tonghop.Name = "rb_tonghop"
        Me.rb_tonghop.Size = New System.Drawing.Size(100, 18)
        Me.rb_tonghop.TabIndex = 0
        Me.rb_tonghop.TabStop = True
        Me.rb_tonghop.Text = "Tổng hợp"
        Me.rb_tonghop.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(119, Byte), Integer), CType(CType(182, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label3.ForeColor = System.Drawing.Color.Maroon
        Me.Label3.Location = New System.Drawing.Point(3, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(872, 1)
        Me.Label3.TabIndex = 4
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.cb_donvi)
        Me.Panel3.Controls.Add(Me.lbl_tungay)
        Me.Panel3.Controls.Add(Me.dtpk_tungay)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.lbl_ngaybc)
        Me.Panel3.Controls.Add(Me.dtpk_ngaybc)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(3, 50)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(872, 22)
        Me.Panel3.TabIndex = 2
        '
        'cb_donvi
        '
        Me.cb_donvi.BackColor = System.Drawing.Color.White
        Me.cb_donvi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cb_donvi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_donvi.ForeColor = System.Drawing.Color.Black
        Me.cb_donvi.FormattingEnabled = True
        Me.cb_donvi.Location = New System.Drawing.Point(215, 0)
        Me.cb_donvi.Name = "cb_donvi"
        Me.cb_donvi.Size = New System.Drawing.Size(267, 26)
        Me.cb_donvi.TabIndex = 1
        '
        'lbl_tungay
        '
        Me.lbl_tungay.Dock = System.Windows.Forms.DockStyle.Right
        Me.lbl_tungay.Location = New System.Drawing.Point(482, 0)
        Me.lbl_tungay.Name = "lbl_tungay"
        Me.lbl_tungay.Size = New System.Drawing.Size(71, 22)
        Me.lbl_tungay.TabIndex = 2
        Me.lbl_tungay.Text = "Từ ngày "
        Me.lbl_tungay.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtpk_tungay
        '
        Me.dtpk_tungay.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpk_tungay.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpk_tungay.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpk_tungay.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpk_tungay.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpk_tungay.CustomFormat = "dd/MM/yyyy"
        Me.dtpk_tungay.Dock = System.Windows.Forms.DockStyle.Right
        Me.dtpk_tungay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpk_tungay.Location = New System.Drawing.Point(553, 0)
        Me.dtpk_tungay.Name = "dtpk_tungay"
        Me.dtpk_tungay.Size = New System.Drawing.Size(105, 26)
        Me.dtpk_tungay.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label7.Location = New System.Drawing.Point(0, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(215, 22)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Đơn vị "
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbl_ngaybc
        '
        Me.lbl_ngaybc.Dock = System.Windows.Forms.DockStyle.Right
        Me.lbl_ngaybc.Location = New System.Drawing.Point(658, 0)
        Me.lbl_ngaybc.Name = "lbl_ngaybc"
        Me.lbl_ngaybc.Size = New System.Drawing.Size(109, 22)
        Me.lbl_ngaybc.TabIndex = 4
        Me.lbl_ngaybc.Text = "Ngày báo cáo "
        Me.lbl_ngaybc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtpk_ngaybc
        '
        Me.dtpk_ngaybc.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer))
        Me.dtpk_ngaybc.CalendarMonthBackground = System.Drawing.Color.FromArgb(CType(CType(227, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.dtpk_ngaybc.CalendarTitleBackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(214, Byte), Integer))
        Me.dtpk_ngaybc.CalendarTitleForeColor = System.Drawing.Color.Lavender
        Me.dtpk_ngaybc.CalendarTrailingForeColor = System.Drawing.Color.Black
        Me.dtpk_ngaybc.CustomFormat = "dd/MM/yyyy"
        Me.dtpk_ngaybc.Dock = System.Windows.Forms.DockStyle.Right
        Me.dtpk_ngaybc.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpk_ngaybc.Location = New System.Drawing.Point(767, 0)
        Me.dtpk_ngaybc.Name = "dtpk_ngaybc"
        Me.dtpk_ngaybc.Size = New System.Drawing.Size(105, 26)
        Me.dtpk_ngaybc.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(CType(119, Byte), Integer), CType(CType(182, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label6.ForeColor = System.Drawing.Color.Maroon
        Me.Label6.Location = New System.Drawing.Point(3, 49)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(872, 1)
        Me.Label6.TabIndex = 1
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cb_baocao)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(3, 27)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(872, 22)
        Me.Panel1.TabIndex = 0
        '
        'cb_baocao
        '
        Me.cb_baocao.BackColor = System.Drawing.Color.White
        Me.cb_baocao.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cb_baocao.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_baocao.ForeColor = System.Drawing.Color.Black
        Me.cb_baocao.FormattingEnabled = True
        Me.cb_baocao.Location = New System.Drawing.Point(215, 0)
        Me.cb_baocao.Name = "cb_baocao"
        Me.cb_baocao.Size = New System.Drawing.Size(657, 26)
        Me.cb_baocao.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label2.Location = New System.Drawing.Point(0, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(215, 22)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Báo cáo/Sao kê theo yêu cầu "
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Location = New System.Drawing.Point(3, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(872, 5)
        Me.Label1.TabIndex = 3
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BaoCaoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(206, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(878, 249)
        Me.Controls.Add(Me.gb_main)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Navy
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BaoCaoForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: Báo cáo, sao kê khác theo yêu cầu"
        Me.Panel2.ResumeLayout(False)
        Me.gb_main.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.pnl_tonghop_cn_hay_pgd.ResumeLayout(False)
        Me.pnl_opt_tonghop_chitiet.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lbl_waitting As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btn_xuatexcel As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btn_quayra As System.Windows.Forms.Button
    Friend WithEvents gb_main As System.Windows.Forms.GroupBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lbl_ngaybc As System.Windows.Forms.Label
    Friend WithEvents dtpk_ngaybc As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cb_baocao As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lbl_tungay As System.Windows.Forms.Label
    Friend WithEvents dtpk_tungay As System.Windows.Forms.DateTimePicker
    Friend WithEvents cb_donvi As System.Windows.Forms.ComboBox
    Friend WithEvents pnl_opt_tonghop_chitiet As System.Windows.Forms.Panel
    Friend WithEvents rb_chitiet As System.Windows.Forms.RadioButton
    Friend WithEvents rb_tonghop As System.Windows.Forms.RadioButton
    Friend WithEvents pnl_tonghop_cn_hay_pgd As System.Windows.Forms.Panel
    Friend WithEvents rb_tonghop_pgd As System.Windows.Forms.RadioButton
    Friend WithEvents rb_tonghop_cn As System.Windows.Forms.RadioButton
    Friend WithEvents rb_danhsach_chitiet As System.Windows.Forms.RadioButton
End Class
