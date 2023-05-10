<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHT_NhomThVien
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHT_NhomThVien))
        Me.Label3 = New System.Windows.Forms.Label()
        Me.gb_main = New System.Windows.Forms.GroupBox()
        Me.edt_note = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.edt_name = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.edt_code = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btn_add = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.ckb_choiceall = New System.Windows.Forms.CheckBox()
        Me.btn_save = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.btn_cancel = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btn_delete = New System.Windows.Forms.Button()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.btn_back = New System.Windows.Forms.Button()
        Me.dgv_main = New System.Windows.Forms.DataGridView()
        Me.gb_main.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.dgv_main, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label3.Location = New System.Drawing.Point(0, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(459, 4)
        Me.Label3.TabIndex = 1
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'gb_main
        '
        Me.gb_main.Controls.Add(Me.edt_note)
        Me.gb_main.Controls.Add(Me.Label10)
        Me.gb_main.Controls.Add(Me.edt_name)
        Me.gb_main.Controls.Add(Me.Label8)
        Me.gb_main.Controls.Add(Me.edt_code)
        Me.gb_main.Controls.Add(Me.Panel2)
        Me.gb_main.Dock = System.Windows.Forms.DockStyle.Top
        Me.gb_main.Location = New System.Drawing.Point(0, 26)
        Me.gb_main.Name = "gb_main"
        Me.gb_main.Size = New System.Drawing.Size(459, 144)
        Me.gb_main.TabIndex = 2
        Me.gb_main.TabStop = False
        Me.gb_main.Text = " Thông tin nhóm"
        '
        'edt_note
        '
        Me.edt_note.BackColor = System.Drawing.Color.White
        Me.edt_note.Dock = System.Windows.Forms.DockStyle.Fill
        Me.edt_note.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_note.Location = New System.Drawing.Point(114, 64)
        Me.edt_note.Multiline = True
        Me.edt_note.Name = "edt_note"
        Me.edt_note.Size = New System.Drawing.Size(342, 77)
        Me.edt_note.TabIndex = 5
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label10.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label10.Location = New System.Drawing.Point(114, 63)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(342, 1)
        Me.Label10.TabIndex = 4
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'edt_name
        '
        Me.edt_name.BackColor = System.Drawing.Color.White
        Me.edt_name.Dock = System.Windows.Forms.DockStyle.Top
        Me.edt_name.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_name.Location = New System.Drawing.Point(114, 41)
        Me.edt_name.Name = "edt_name"
        Me.edt_name.Size = New System.Drawing.Size(342, 22)
        Me.edt_name.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label8.Location = New System.Drawing.Point(114, 40)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(342, 1)
        Me.Label8.TabIndex = 2
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'edt_code
        '
        Me.edt_code.BackColor = System.Drawing.Color.White
        Me.edt_code.Dock = System.Windows.Forms.DockStyle.Top
        Me.edt_code.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_code.Location = New System.Drawing.Point(114, 18)
        Me.edt_code.Name = "edt_code"
        Me.edt_code.Size = New System.Drawing.Size(342, 22)
        Me.edt_code.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel2.Location = New System.Drawing.Point(3, 18)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(111, 123)
        Me.Panel2.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label9.Location = New System.Drawing.Point(0, 46)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(111, 22)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Ghi chú"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label6.Location = New System.Drawing.Point(0, 45)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(111, 1)
        Me.Label6.TabIndex = 3
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label7.Location = New System.Drawing.Point(0, 23)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(111, 22)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Tên nhóm"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label5.Location = New System.Drawing.Point(0, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(111, 1)
        Me.Label5.TabIndex = 1
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label4.Location = New System.Drawing.Point(0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 22)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Mã hiệu nhóm"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGray
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(459, 22)
        Me.Panel1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.DarkGray
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Maroon
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "THÔNG TIN NHÓM NGƯỜI DÙNG"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label11.Location = New System.Drawing.Point(0, 170)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(459, 4)
        Me.Label11.TabIndex = 3
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.DarkGray
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Maroon
        Me.Label2.Location = New System.Drawing.Point(0, 174)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(459, 21)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Danh sách nhóm người dùng"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkGray
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.btn_add)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Controls.Add(Me.ckb_choiceall)
        Me.Panel3.Controls.Add(Me.btn_save)
        Me.Panel3.Controls.Add(Me.Label12)
        Me.Panel3.Controls.Add(Me.btn_cancel)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Controls.Add(Me.btn_delete)
        Me.Panel3.Controls.Add(Me.Label20)
        Me.Panel3.Controls.Add(Me.btn_back)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.ForeColor = System.Drawing.Color.Black
        Me.Panel3.Location = New System.Drawing.Point(0, 359)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(459, 27)
        Me.Panel3.TabIndex = 5
        '
        'btn_add
        '
        Me.btn_add.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_add.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_add.Location = New System.Drawing.Point(142, 0)
        Me.btn_add.Name = "btn_add"
        Me.btn_add.Size = New System.Drawing.Size(61, 23)
        Me.btn_add.TabIndex = 0
        Me.btn_add.Text = "&Thêm"
        Me.btn_add.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label14.Location = New System.Drawing.Point(203, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(2, 23)
        Me.Label14.TabIndex = 1
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ckb_choiceall
        '
        Me.ckb_choiceall.Dock = System.Windows.Forms.DockStyle.Left
        Me.ckb_choiceall.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ckb_choiceall.ForeColor = System.Drawing.Color.Maroon
        Me.ckb_choiceall.Location = New System.Drawing.Point(0, 0)
        Me.ckb_choiceall.Name = "ckb_choiceall"
        Me.ckb_choiceall.Size = New System.Drawing.Size(77, 23)
        Me.ckb_choiceall.TabIndex = 4
        Me.ckb_choiceall.Text = "&Chọn cả"
        Me.ckb_choiceall.UseVisualStyleBackColor = True
        '
        'btn_save
        '
        Me.btn_save.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_save.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_save.Location = New System.Drawing.Point(205, 0)
        Me.btn_save.Name = "btn_save"
        Me.btn_save.Size = New System.Drawing.Size(61, 23)
        Me.btn_save.TabIndex = 2
        Me.btn_save.Text = "&Ghi"
        Me.btn_save.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label12.Location = New System.Drawing.Point(266, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(2, 23)
        Me.Label12.TabIndex = 3
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_cancel
        '
        Me.btn_cancel.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_cancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_cancel.Location = New System.Drawing.Point(268, 0)
        Me.btn_cancel.Name = "btn_cancel"
        Me.btn_cancel.Size = New System.Drawing.Size(61, 23)
        Me.btn_cancel.TabIndex = 4
        Me.btn_cancel.Text = "&Bỏ qua"
        Me.btn_cancel.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label13.Location = New System.Drawing.Point(329, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(2, 23)
        Me.Label13.TabIndex = 5
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_delete
        '
        Me.btn_delete.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_delete.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_delete.Location = New System.Drawing.Point(331, 0)
        Me.btn_delete.Name = "btn_delete"
        Me.btn_delete.Size = New System.Drawing.Size(61, 23)
        Me.btn_delete.TabIndex = 6
        Me.btn_delete.Text = "&Xoá"
        Me.btn_delete.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label20.Location = New System.Drawing.Point(392, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(2, 23)
        Me.Label20.TabIndex = 7
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_back
        '
        Me.btn_back.Dock = System.Windows.Forms.DockStyle.Right
        Me.btn_back.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_back.Location = New System.Drawing.Point(394, 0)
        Me.btn_back.Name = "btn_back"
        Me.btn_back.Size = New System.Drawing.Size(61, 23)
        Me.btn_back.TabIndex = 8
        Me.btn_back.Text = "&Quay ra"
        Me.btn_back.UseVisualStyleBackColor = True
        '
        'dgv_main
        '
        Me.dgv_main.AllowUserToAddRows = False
        Me.dgv_main.AllowUserToDeleteRows = False
        Me.dgv_main.AllowUserToResizeRows = False
        Me.dgv_main.BackgroundColor = System.Drawing.Color.White
        Me.dgv_main.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgv_main.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.DarkGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv_main.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgv_main.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgv_main.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgv_main.Location = New System.Drawing.Point(0, 195)
        Me.dgv_main.Name = "dgv_main"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgv_main.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgv_main.RowHeadersVisible = False
        Me.dgv_main.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        Me.dgv_main.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(119, Byte), Integer), CType(CType(182, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.dgv_main.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black
        Me.dgv_main.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgv_main.Size = New System.Drawing.Size(459, 164)
        Me.dgv_main.TabIndex = 6
        '
        'frmHT_NhomThVien
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(459, 386)
        Me.Controls.Add(Me.dgv_main)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.gb_main)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHT_NhomThVien"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự"
        Me.gb_main.ResumeLayout(False)
        Me.gb_main.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.dgv_main, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents gb_main As System.Windows.Forms.GroupBox
    Friend WithEvents edt_note As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents edt_name As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents edt_code As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btn_save As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btn_cancel As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents btn_back As System.Windows.Forms.Button
    Friend WithEvents dgv_main As System.Windows.Forms.DataGridView
    Friend WithEvents ckb_choiceall As System.Windows.Forms.CheckBox
    Friend WithEvents btn_add As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btn_delete As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
End Class
