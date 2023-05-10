<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTBNangBacLuong
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTBNangBacLuong))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rpt_View = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.gb_main_thamso = New System.Windows.Forms.GroupBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.dpkNgayLapBieu = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNam = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboKy = New System.Windows.Forms.ComboBox()
        Me.labKy = New System.Windows.Forms.Label()
        Me.rdNgach = New System.Windows.Forms.RadioButton()
        Me.rdBac = New System.Windows.Forms.RadioButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cboDonVi = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbAll = New System.Windows.Forms.CheckBox()
        Me.chk_TrunguongQL = New System.Windows.Forms.CheckBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.cmdReport = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.lbl_waitting = New System.Windows.Forms.Label()
        Me.btn_exportexcel = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.labStatusProcess = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.rpt_TBNangBac = New QLNS.TBNangLuong()
        Me.rpt_TBNangNgach = New QLNS.TBNangNgachLuong()
        Me.Panel1.SuspendLayout()
        Me.gb_main_thamso.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rpt_View)
        Me.Panel1.Controls.Add(Me.gb_main_thamso)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 25)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(789, 590)
        Me.Panel1.TabIndex = 6
        '
        'rpt_View
        '
        Me.rpt_View.ActiveViewIndex = -1
        Me.rpt_View.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rpt_View.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.rpt_View.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rpt_View.Location = New System.Drawing.Point(0, 77)
        Me.rpt_View.Name = "rpt_View"
        Me.rpt_View.SelectionFormula = ""
        Me.rpt_View.ShowRefreshButton = False
        Me.rpt_View.Size = New System.Drawing.Size(789, 513)
        Me.rpt_View.TabIndex = 9
        Me.rpt_View.UseWaitCursor = True
        Me.rpt_View.ViewTimeSelectionFormula = ""
        '
        'gb_main_thamso
        '
        Me.gb_main_thamso.Controls.Add(Me.Panel4)
        Me.gb_main_thamso.Controls.Add(Me.Label8)
        Me.gb_main_thamso.Controls.Add(Me.Panel2)
        Me.gb_main_thamso.Dock = System.Windows.Forms.DockStyle.Top
        Me.gb_main_thamso.Location = New System.Drawing.Point(0, 0)
        Me.gb_main_thamso.Name = "gb_main_thamso"
        Me.gb_main_thamso.Size = New System.Drawing.Size(789, 77)
        Me.gb_main_thamso.TabIndex = 1
        Me.gb_main_thamso.TabStop = False
        Me.gb_main_thamso.Text = "Khai báo tham số"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.dpkNgayLapBieu)
        Me.Panel4.Controls.Add(Me.Label2)
        Me.Panel4.Controls.Add(Me.txtNam)
        Me.Panel4.Controls.Add(Me.Label3)
        Me.Panel4.Controls.Add(Me.cboKy)
        Me.Panel4.Controls.Add(Me.labKy)
        Me.Panel4.Controls.Add(Me.rdNgach)
        Me.Panel4.Controls.Add(Me.rdBac)
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel4.Location = New System.Drawing.Point(3, 46)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(783, 22)
        Me.Panel4.TabIndex = 3
        '
        'dpkNgayLapBieu
        '
        Me.dpkNgayLapBieu.CustomFormat = "dd/MM/yyyy"
        Me.dpkNgayLapBieu.Dock = System.Windows.Forms.DockStyle.Left
        Me.dpkNgayLapBieu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dpkNgayLapBieu.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dpkNgayLapBieu.Location = New System.Drawing.Point(656, 0)
        Me.dpkNgayLapBieu.Name = "dpkNgayLapBieu"
        Me.dpkNgayLapBieu.Size = New System.Drawing.Size(105, 26)
        Me.dpkNgayLapBieu.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(563, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 22)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Ngày lập "
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtNam
        '
        Me.txtNam.BackColor = System.Drawing.Color.White
        Me.txtNam.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNam.Dock = System.Windows.Forms.DockStyle.Left
        Me.txtNam.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNam.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtNam.Location = New System.Drawing.Point(512, 0)
        Me.txtNam.Name = "txtNam"
        Me.txtNam.Size = New System.Drawing.Size(51, 26)
        Me.txtNam.TabIndex = 6
        Me.txtNam.Text = "2019"
        Me.txtNam.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(454, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 22)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Năm "
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboKy
        '
        Me.cboKy.BackColor = System.Drawing.Color.White
        Me.cboKy.DisplayMember = "Display"
        Me.cboKy.Dock = System.Windows.Forms.DockStyle.Left
        Me.cboKy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboKy.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboKy.ForeColor = System.Drawing.Color.Black
        Me.cboKy.FormattingEnabled = True
        Me.cboKy.Items.AddRange(New Object() {"1", "2", "3", "4"})
        Me.cboKy.Location = New System.Drawing.Point(390, 0)
        Me.cboKy.Name = "cboKy"
        Me.cboKy.Size = New System.Drawing.Size(64, 26)
        Me.cboKy.TabIndex = 4
        Me.cboKy.ValueMember = "Value"
        '
        'labKy
        '
        Me.labKy.Dock = System.Windows.Forms.DockStyle.Left
        Me.labKy.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labKy.Location = New System.Drawing.Point(330, 0)
        Me.labKy.Name = "labKy"
        Me.labKy.Size = New System.Drawing.Size(60, 22)
        Me.labKy.TabIndex = 3
        Me.labKy.Text = "Quý "
        Me.labKy.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'rdNgach
        '
        Me.rdNgach.Dock = System.Windows.Forms.DockStyle.Left
        Me.rdNgach.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdNgach.Location = New System.Drawing.Point(215, 0)
        Me.rdNgach.Name = "rdNgach"
        Me.rdNgach.Size = New System.Drawing.Size(115, 22)
        Me.rdNgach.TabIndex = 2
        Me.rdNgach.TabStop = True
        Me.rdNgach.Text = "Nâng ngạch"
        Me.rdNgach.UseVisualStyleBackColor = True
        '
        'rdBac
        '
        Me.rdBac.Checked = True
        Me.rdBac.Dock = System.Windows.Forms.DockStyle.Left
        Me.rdBac.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdBac.Location = New System.Drawing.Point(100, 0)
        Me.rdBac.Name = "rdBac"
        Me.rdBac.Size = New System.Drawing.Size(115, 22)
        Me.rdBac.TabIndex = 1
        Me.rdBac.TabStop = True
        Me.rdBac.Text = "Nâng bậc"
        Me.rdBac.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(85, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(15, 22)
        Me.Label6.TabIndex = 134
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Xét "
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(3, 44)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(783, 2)
        Me.Label8.TabIndex = 2
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.cboDonVi)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.cbAll)
        Me.Panel2.Controls.Add(Me.chk_TrunguongQL)
        Me.Panel2.Controls.Add(Me.Label40)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 22)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(783, 22)
        Me.Panel2.TabIndex = 1
        '
        'cboDonVi
        '
        Me.cboDonVi.BackColor = System.Drawing.Color.White
        Me.cboDonVi.DisplayMember = "Display"
        Me.cboDonVi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cboDonVi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDonVi.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDonVi.ForeColor = System.Drawing.Color.Black
        Me.cboDonVi.FormattingEnabled = True
        Me.cboDonVi.Location = New System.Drawing.Point(85, 0)
        Me.cboDonVi.Name = "cboDonVi"
        Me.cboDonVi.Size = New System.Drawing.Size(299, 26)
        Me.cboDonVi.TabIndex = 1
        Me.cboDonVi.ValueMember = "Value"
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(384, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(10, 22)
        Me.Label7.TabIndex = 2
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbAll
        '
        Me.cbAll.Dock = System.Windows.Forms.DockStyle.Right
        Me.cbAll.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAll.ForeColor = System.Drawing.Color.Maroon
        Me.cbAll.Location = New System.Drawing.Point(394, 0)
        Me.cbAll.Name = "cbAll"
        Me.cbAll.Size = New System.Drawing.Size(205, 22)
        Me.cbAll.TabIndex = 3
        Me.cbAll.Text = "Thống kê toàn đơn vị"
        Me.cbAll.UseVisualStyleBackColor = True
        '
        'chk_TrunguongQL
        '
        Me.chk_TrunguongQL.AutoSize = True
        Me.chk_TrunguongQL.Dock = System.Windows.Forms.DockStyle.Right
        Me.chk_TrunguongQL.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_TrunguongQL.ForeColor = System.Drawing.Color.Maroon
        Me.chk_TrunguongQL.Location = New System.Drawing.Point(599, 0)
        Me.chk_TrunguongQL.Name = "chk_TrunguongQL"
        Me.chk_TrunguongQL.Size = New System.Drawing.Size(184, 22)
        Me.chk_TrunguongQL.TabIndex = 4
        Me.chk_TrunguongQL.Text = "&Cán bộ do TW Quản lý"
        Me.chk_TrunguongQL.UseVisualStyleBackColor = True
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.SystemColors.Control
        Me.Label40.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label40.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.Location = New System.Drawing.Point(0, 0)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(85, 22)
        Me.Label40.TabIndex = 0
        Me.Label40.Text = "Đơn vị "
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdReport
        '
        Me.cmdReport.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdReport.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReport.Location = New System.Drawing.Point(535, 0)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(150, 25)
        Me.cmdReport.TabIndex = 2
        Me.cmdReport.Text = "&Xem danh sách"
        Me.cmdReport.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.Controls.Add(Me.lbl_waitting)
        Me.Panel3.Controls.Add(Me.btn_exportexcel)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.labStatusProcess)
        Me.Panel3.Controls.Add(Me.cmdReport)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.cmdClose)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(789, 25)
        Me.Panel3.TabIndex = 2
        '
        'lbl_waitting
        '
        Me.lbl_waitting.BackColor = System.Drawing.Color.Transparent
        Me.lbl_waitting.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbl_waitting.Font = New System.Drawing.Font("Tahoma", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_waitting.ForeColor = System.Drawing.Color.Maroon
        Me.lbl_waitting.Location = New System.Drawing.Point(0, 0)
        Me.lbl_waitting.Name = "lbl_waitting"
        Me.lbl_waitting.Size = New System.Drawing.Size(416, 25)
        Me.lbl_waitting.TabIndex = 21
        Me.lbl_waitting.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_exportexcel
        '
        Me.btn_exportexcel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_exportexcel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_exportexcel.Location = New System.Drawing.Point(416, 0)
        Me.btn_exportexcel.Name = "btn_exportexcel"
        Me.btn_exportexcel.Size = New System.Drawing.Size(117, 25)
        Me.btn_exportexcel.TabIndex = 0
        Me.btn_exportexcel.Text = "Xuất &excel"
        Me.btn_exportexcel.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label9.Location = New System.Drawing.Point(533, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(2, 25)
        Me.Label9.TabIndex = 1
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'labStatusProcess
        '
        Me.labStatusProcess.AutoSize = True
        Me.labStatusProcess.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labStatusProcess.ForeColor = System.Drawing.Color.Maroon
        Me.labStatusProcess.Location = New System.Drawing.Point(0, 7)
        Me.labStatusProcess.Name = "labStatusProcess"
        Me.labStatusProcess.Size = New System.Drawing.Size(0, 17)
        Me.labStatusProcess.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label5.Location = New System.Drawing.Point(685, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(2, 25)
        Me.Label5.TabIndex = 3
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label4.Location = New System.Drawing.Point(687, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(2, 25)
        Me.Label4.TabIndex = 17
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdClose
        '
        Me.cmdClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.cmdClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClose.Location = New System.Drawing.Point(689, 0)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(100, 25)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "&Quay ra"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'frmTBNangBacLuong
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 615)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmTBNangBacLuong"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = ""
        Me.Text = "Quản lý nhân sự: Thông báo danh sách cán bộ đến kỳ nâng bậc/ngạch lương"
        Me.Panel1.ResumeLayout(False)
        Me.gb_main_thamso.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents rpt_TBNangBac As QLNS.TBNangLuong
    Friend WithEvents rpt_TBNangNgach As QLNS.TBNangNgachLuong
    Friend WithEvents gb_main_thamso As System.Windows.Forms.GroupBox
    Friend WithEvents labStatusProcess As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Private WithEvents rpt_View As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents dpkNgayLapBieu As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNam As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboKy As System.Windows.Forms.ComboBox
    Friend WithEvents labKy As System.Windows.Forms.Label
    Friend WithEvents rdNgach As System.Windows.Forms.RadioButton
    Friend WithEvents rdBac As System.Windows.Forms.RadioButton
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btn_exportexcel As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lbl_waitting As System.Windows.Forms.Label
    Friend WithEvents cboDonVi As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbAll As System.Windows.Forms.CheckBox
    Friend WithEvents chk_TrunguongQL As System.Windows.Forms.CheckBox
End Class
