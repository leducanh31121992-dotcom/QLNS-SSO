<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCheckDL_DanhMuc
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
        Me.bnt_CheckDM = New System.Windows.Forms.Button()
        Me.labInfo = New System.Windows.Forms.Label()
        Me.cmd_INS_D = New System.Windows.Forms.Button()
        Me.cmdChonTM = New System.Windows.Forms.Button()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'bnt_CheckDM
        '
        Me.bnt_CheckDM.Location = New System.Drawing.Point(25, 136)
        Me.bnt_CheckDM.Name = "bnt_CheckDM"
        Me.bnt_CheckDM.Size = New System.Drawing.Size(235, 23)
        Me.bnt_CheckDM.TabIndex = 0
        Me.bnt_CheckDM.Text = "Kiểm tra dữ liệu danh mục giữ TW và CN"
        Me.bnt_CheckDM.UseVisualStyleBackColor = True
        '
        'labInfo
        '
        Me.labInfo.AutoSize = True
        Me.labInfo.Location = New System.Drawing.Point(12, 9)
        Me.labInfo.Name = "labInfo"
        Me.labInfo.Size = New System.Drawing.Size(52, 13)
        Me.labInfo.TabIndex = 1
        Me.labInfo.Text = "Thong tin"
        '
        'cmd_INS_D
        '
        Me.cmd_INS_D.Location = New System.Drawing.Point(25, 73)
        Me.cmd_INS_D.Name = "cmd_INS_D"
        Me.cmd_INS_D.Size = New System.Drawing.Size(158, 23)
        Me.cmd_INS_D.TabIndex = 2
        Me.cmd_INS_D.Text = "Insert data into Database"
        Me.cmd_INS_D.UseVisualStyleBackColor = True
        '
        'cmdChonTM
        '
        Me.cmdChonTM.Image = Global.QLNS.My.Resources.Resources.chon_file
        Me.cmdChonTM.Location = New System.Drawing.Point(252, 37)
        Me.cmdChonTM.Name = "cmdChonTM"
        Me.cmdChonTM.Size = New System.Drawing.Size(28, 22)
        Me.cmdChonTM.TabIndex = 10
        Me.cmdChonTM.Text = "..."
        Me.cmdChonTM.UseVisualStyleBackColor = True
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(15, 37)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(231, 20)
        Me.txtPath.TabIndex = 9
        '
        'frmCheckDL_DanhMuc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.cmdChonTM)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.cmd_INS_D)
        Me.Controls.Add(Me.labInfo)
        Me.Controls.Add(Me.bnt_CheckDM)
        Me.Name = "frmCheckDL_DanhMuc"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmCheckDL_DanhMuc"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents bnt_CheckDM As System.Windows.Forms.Button
    Friend WithEvents labInfo As System.Windows.Forms.Label
    Friend WithEvents cmd_INS_D As System.Windows.Forms.Button
    Friend WithEvents cmdChonTM As System.Windows.Forms.Button
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
End Class
