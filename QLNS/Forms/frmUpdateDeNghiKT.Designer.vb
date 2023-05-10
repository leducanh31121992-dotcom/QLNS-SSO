<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateDeNghiKT
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdateDeNghiKT))
        Me.pnlHoTen = New System.Windows.Forms.Panel()
        Me.lab = New System.Windows.Forms.Label()
        Me.txtDoiTuong = New System.Windows.Forms.TextBox()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.bntUpdate = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.gridResult = New System.Windows.Forms.DataGridView()
        Me.pnlHoTen.SuspendLayout()
        Me.Panel11.SuspendLayout()
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlHoTen
        '
        Me.pnlHoTen.Controls.Add(Me.lab)
        Me.pnlHoTen.Controls.Add(Me.txtDoiTuong)
        Me.pnlHoTen.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlHoTen.Location = New System.Drawing.Point(0, 0)
        Me.pnlHoTen.Name = "pnlHoTen"
        Me.pnlHoTen.Size = New System.Drawing.Size(451, 35)
        Me.pnlHoTen.TabIndex = 5
        '
        'lab
        '
        Me.lab.AutoSize = True
        Me.lab.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lab.Location = New System.Drawing.Point(4, 11)
        Me.lab.Name = "lab"
        Me.lab.Size = New System.Drawing.Size(42, 14)
        Me.lab.TabIndex = 2
        Me.lab.Text = "Label1"
        '
        'txtDoiTuong
        '
        Me.txtDoiTuong.BackColor = System.Drawing.Color.WhiteSmoke
        Me.txtDoiTuong.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDoiTuong.Location = New System.Drawing.Point(76, 8)
        Me.txtDoiTuong.Name = "txtDoiTuong"
        Me.txtDoiTuong.ReadOnly = True
        Me.txtDoiTuong.Size = New System.Drawing.Size(373, 22)
        Me.txtDoiTuong.TabIndex = 1
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.DarkGray
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel11.Controls.Add(Me.bntUpdate)
        Me.Panel11.Controls.Add(Me.bntClose)
        Me.Panel11.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel11.Location = New System.Drawing.Point(0, 142)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(451, 25)
        Me.Panel11.TabIndex = 6
        '
        'bntUpdate
        '
        Me.bntUpdate.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntUpdate.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntUpdate.Location = New System.Drawing.Point(325, 0)
        Me.bntUpdate.Name = "bntUpdate"
        Me.bntUpdate.Size = New System.Drawing.Size(61, 21)
        Me.bntUpdate.TabIndex = 1
        Me.bntUpdate.Text = "&Ghi"
        Me.bntUpdate.UseVisualStyleBackColor = True
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntClose.Location = New System.Drawing.Point(386, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(61, 21)
        Me.bntClose.TabIndex = 2
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
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
        Me.gridResult.Location = New System.Drawing.Point(0, 35)
        Me.gridResult.MultiSelect = False
        Me.gridResult.Name = "gridResult"
        Me.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridResult.Size = New System.Drawing.Size(451, 107)
        Me.gridResult.TabIndex = 108
        '
        'frmUpdateDeNghiKT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(451, 167)
        Me.Controls.Add(Me.gridResult)
        Me.Controls.Add(Me.Panel11)
        Me.Controls.Add(Me.pnlHoTen)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdateDeNghiKT"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Duyệt đề nghị khen thưởng"
        Me.pnlHoTen.ResumeLayout(False)
        Me.pnlHoTen.PerformLayout()
        Me.Panel11.ResumeLayout(False)
        CType(Me.gridResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlHoTen As System.Windows.Forms.Panel
    Friend WithEvents lab As System.Windows.Forms.Label
    Friend WithEvents txtDoiTuong As System.Windows.Forms.TextBox
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents bntUpdate As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents gridResult As System.Windows.Forms.DataGridView
End Class
