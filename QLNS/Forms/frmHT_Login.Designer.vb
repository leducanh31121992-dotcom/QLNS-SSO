<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHT_Login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHT_Login))
        Me.edt_username = New System.Windows.Forms.TextBox()
        Me.edt_password = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btn_reset = New System.Windows.Forms.Button()
        Me.btn_login = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.edt_maPOS = New System.Windows.Forms.TextBox()
        Me.chkLuuThongTin = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'edt_username
        '
        Me.edt_username.BackColor = System.Drawing.Color.FromArgb(CType(CType(219, Byte), Integer), CType(CType(242, Byte), Integer), CType(CType(222, Byte), Integer))
        Me.edt_username.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.edt_username.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_username.Location = New System.Drawing.Point(259, 137)
        Me.edt_username.Name = "edt_username"
        Me.edt_username.Size = New System.Drawing.Size(151, 22)
        Me.edt_username.TabIndex = 0
        '
        'edt_password
        '
        Me.edt_password.BackColor = System.Drawing.Color.FromArgb(CType(CType(219, Byte), Integer), CType(CType(242, Byte), Integer), CType(CType(222, Byte), Integer))
        Me.edt_password.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_password.Location = New System.Drawing.Point(259, 161)
        Me.edt_password.Name = "edt_password"
        Me.edt_password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.edt_password.Size = New System.Drawing.Size(151, 22)
        Me.edt_password.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(154, 138)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Tên đăng nhập"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(154, 162)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 22)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Mật khẩu"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btn_reset
        '
        Me.btn_reset.BackColor = System.Drawing.Color.Transparent
        Me.btn_reset.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_reset.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_reset.Image = Global.QLNS.My.Resources.Resources.nhap_lai
        Me.btn_reset.Location = New System.Drawing.Point(342, 229)
        Me.btn_reset.Name = "btn_reset"
        Me.btn_reset.Size = New System.Drawing.Size(70, 20)
        Me.btn_reset.TabIndex = 5
        Me.btn_reset.UseVisualStyleBackColor = False
        '
        'btn_login
        '
        Me.btn_login.BackColor = System.Drawing.Color.Transparent
        Me.btn_login.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_login.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_login.Image = Global.QLNS.My.Resources.Resources.dang_nhap
        Me.btn_login.Location = New System.Drawing.Point(258, 229)
        Me.btn_login.Name = "btn_login"
        Me.btn_login.Size = New System.Drawing.Size(80, 20)
        Me.btn_login.TabIndex = 4
        Me.btn_login.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(154, 186)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 22)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Mã POS"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'edt_maPOS
        '
        Me.edt_maPOS.BackColor = System.Drawing.Color.FromArgb(CType(CType(219, Byte), Integer), CType(CType(242, Byte), Integer), CType(CType(222, Byte), Integer))
        Me.edt_maPOS.ForeColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(37, Byte), Integer), CType(CType(127, Byte), Integer))
        Me.edt_maPOS.Location = New System.Drawing.Point(259, 186)
        Me.edt_maPOS.Name = "edt_maPOS"
        Me.edt_maPOS.Size = New System.Drawing.Size(151, 22)
        Me.edt_maPOS.TabIndex = 2
        '
        'chkLuuThongTin
        '
        Me.chkLuuThongTin.AutoSize = True
        Me.chkLuuThongTin.BackColor = System.Drawing.Color.Transparent
        Me.chkLuuThongTin.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLuuThongTin.ForeColor = System.Drawing.Color.FromArgb(CType(CType(3, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.chkLuuThongTin.Location = New System.Drawing.Point(258, 210)
        Me.chkLuuThongTin.Name = "chkLuuThongTin"
        Me.chkLuuThongTin.Size = New System.Drawing.Size(144, 17)
        Me.chkLuuThongTin.TabIndex = 3
        Me.chkLuuThongTin.Text = "Lưu thông tin đăng nhập"
        Me.chkLuuThongTin.UseVisualStyleBackColor = False
        '
        'frmHT_Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.QLNS.My.Resources.Resources.loginmain
        Me.ClientSize = New System.Drawing.Size(433, 254)
        Me.Controls.Add(Me.chkLuuThongTin)
        Me.Controls.Add(Me.edt_maPOS)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btn_reset)
        Me.Controls.Add(Me.btn_login)
        Me.Controls.Add(Me.edt_password)
        Me.Controls.Add(Me.edt_username)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Transparent
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHT_Login"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Đăng nhập"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents edt_username As System.Windows.Forms.TextBox
    Friend WithEvents edt_password As System.Windows.Forms.TextBox
    Friend WithEvents btn_login As System.Windows.Forms.Button
    Friend WithEvents btn_reset As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents edt_maPOS As System.Windows.Forms.TextBox
    Friend WithEvents chkLuuThongTin As System.Windows.Forms.CheckBox
End Class
