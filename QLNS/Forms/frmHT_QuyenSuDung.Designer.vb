<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHT_QuyenSuDung
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
        Me.btn_accept = New System.Windows.Forms.Button()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.btn_back = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lbl_fullname = New System.Windows.Forms.Label()
        Me.edt_confirm = New System.Windows.Forms.TextBox()
        Me.edt_pass_new = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txt_Username = New System.Windows.Forms.TextBox()
        Me.cboNhomQuyen = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkChangePass = New System.Windows.Forms.CheckBox()
        Me.txtMaPOS = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkGray
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.chkChangePass)
        Me.Panel1.Controls.Add(Me.btn_accept)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.btn_back)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.ForeColor = System.Drawing.Color.Navy
        Me.Panel1.Location = New System.Drawing.Point(0, 201)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(371, 27)
        Me.Panel1.TabIndex = 6
        '
        'btn_accept
        '
        Me.btn_accept.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_accept.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_accept.Location = New System.Drawing.Point(216, 0)
        Me.btn_accept.Name = "btn_accept"
        Me.btn_accept.Size = New System.Drawing.Size(77, 23)
        Me.btn_accept.TabIndex = 1
        Me.btn_accept.Text = "&Chấp nhận"
        Me.btn_accept.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label20.Location = New System.Drawing.Point(364, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(3, 23)
        Me.Label20.TabIndex = 1
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_back
        '
        Me.btn_back.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_back.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_back.Location = New System.Drawing.Point(296, 0)
        Me.btn_back.Name = "btn_back"
        Me.btn_back.Size = New System.Drawing.Size(61, 23)
        Me.btn_back.TabIndex = 2
        Me.btn_back.Text = "&Quay ra"
        Me.btn_back.UseVisualStyleBackColor = True
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.DarkGray
        Me.lblTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Maroon
        Me.lblTitle.Location = New System.Drawing.Point(0, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(371, 24)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = " THÔNG TIN ĐĂNG NHẬP"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 115)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(113, 22)
        Me.Label12.TabIndex = 8
        Me.Label12.Text = "Xác nhận mật khẩu"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 89)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(113, 22)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "Mật khẩu"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(113, 22)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Tên đăng nhập"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 37)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 22)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Họ và tên"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lbl_fullname
        '
        Me.lbl_fullname.BackColor = System.Drawing.SystemColors.Control
        Me.lbl_fullname.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_fullname.Location = New System.Drawing.Point(127, 38)
        Me.lbl_fullname.Name = "lbl_fullname"
        Me.lbl_fullname.Size = New System.Drawing.Size(226, 22)
        Me.lbl_fullname.TabIndex = 0
        Me.lbl_fullname.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'edt_confirm
        '
        Me.edt_confirm.BackColor = System.Drawing.Color.White
        Me.edt_confirm.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_confirm.Location = New System.Drawing.Point(127, 116)
        Me.edt_confirm.Name = "edt_confirm"
        Me.edt_confirm.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.edt_confirm.Size = New System.Drawing.Size(226, 22)
        Me.edt_confirm.TabIndex = 3
        '
        'edt_pass_new
        '
        Me.edt_pass_new.BackColor = System.Drawing.Color.White
        Me.edt_pass_new.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_pass_new.Location = New System.Drawing.Point(127, 90)
        Me.edt_pass_new.Name = "edt_pass_new"
        Me.edt_pass_new.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.edt_pass_new.Size = New System.Drawing.Size(226, 22)
        Me.edt_pass_new.TabIndex = 2
        '
        'Label14
        '
        Me.Label14.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label14.Location = New System.Drawing.Point(366, 24)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(5, 177)
        Me.Label14.TabIndex = 12
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label15
        '
        Me.Label15.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label15.Location = New System.Drawing.Point(0, 24)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(5, 177)
        Me.Label15.TabIndex = 13
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txt_Username
        '
        Me.txt_Username.BackColor = System.Drawing.SystemColors.Control
        Me.txt_Username.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txt_Username.Location = New System.Drawing.Point(127, 64)
        Me.txt_Username.Name = "txt_Username"
        Me.txt_Username.ReadOnly = True
        Me.txt_Username.Size = New System.Drawing.Size(226, 22)
        Me.txt_Username.TabIndex = 1
        '
        'cboNhomQuyen
        '
        Me.cboNhomQuyen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboNhomQuyen.FormattingEnabled = True
        Me.cboNhomQuyen.Location = New System.Drawing.Point(127, 168)
        Me.cboNhomQuyen.Name = "cboNhomQuyen"
        Me.cboNhomQuyen.Size = New System.Drawing.Size(226, 22)
        Me.cboNhomQuyen.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 22)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Nhóm quyền"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkChangePass
        '
        Me.chkChangePass.AutoSize = True
        Me.chkChangePass.Location = New System.Drawing.Point(10, 2)
        Me.chkChangePass.Name = "chkChangePass"
        Me.chkChangePass.Size = New System.Drawing.Size(128, 18)
        Me.chkChangePass.TabIndex = 0
        Me.chkChangePass.Text = "Thay đổi mật khẩu"
        Me.chkChangePass.UseVisualStyleBackColor = True
        '
        'txtMaPOS
        '
        Me.txtMaPOS.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaPOS.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.txtMaPOS.Location = New System.Drawing.Point(127, 142)
        Me.txtMaPOS.Name = "txtMaPOS"
        Me.txtMaPOS.Size = New System.Drawing.Size(226, 22)
        Me.txtMaPOS.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 142)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 22)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Mã POS"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmHT_QuyenSuDung
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(371, 228)
        Me.Controls.Add(Me.txtMaPOS)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboNhomQuyen)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txt_Username)
        Me.Controls.Add(Me.edt_confirm)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.edt_pass_new)
        Me.Controls.Add(Me.lbl_fullname)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblTitle)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHT_QuyenSuDung"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

        '
        ' btn_sso
        '
        Me.btn_sso.BackColor = System.Drawing.Color.Transparent
        Me.btn_sso.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_sso.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_sso.ForeColor = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(204, Byte), Integer))
        ' Đặt vị trí X=154 (cùng lề trái với các Label), Y=229
        Me.btn_sso.Location = New System.Drawing.Point(154, 229)
        Me.btn_sso.Name = "btn_sso"
        Me.btn_sso.Size = New System.Drawing.Size(100, 20)
        Me.btn_sso.TabIndex = 6
        Me.btn_sso.Text = "Đăng nhập SSO"
        Me.btn_sso.UseVisualStyleBackColor = False
        Me.Controls.Add(Me.btn_sso)
    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btn_accept As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents btn_back As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbl_fullname As System.Windows.Forms.Label
    Friend WithEvents edt_confirm As System.Windows.Forms.TextBox
    Friend WithEvents edt_pass_new As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txt_Username As System.Windows.Forms.TextBox
    Friend WithEvents cboNhomQuyen As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkChangePass As System.Windows.Forms.CheckBox
    Friend WithEvents txtMaPOS As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_sso As System.Windows.Forms.Button
End Class
