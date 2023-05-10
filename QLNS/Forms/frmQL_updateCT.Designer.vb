<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQL_updateCT
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQL_updateCT))
        Me.gridDS = New System.Windows.Forms.DataGridView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.labDateUp = New System.Windows.Forms.Label()
        Me.labDateVer = New System.Windows.Forms.Label()
        Me.labVer = New System.Windows.Forms.Label()
        Me.labDS = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.bntClose = New System.Windows.Forms.Button()
        CType(Me.gridDS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gridDS
        '
        Me.gridDS.AllowUserToAddRows = False
        Me.gridDS.AllowUserToDeleteRows = False
        Me.gridDS.AllowUserToResizeRows = False
        Me.gridDS.BackgroundColor = System.Drawing.Color.White
        Me.gridDS.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridDS.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.DarkGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridDS.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.gridDS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.gridDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gridDS.Location = New System.Drawing.Point(0, 148)
        Me.gridDS.Name = "gridDS"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridDS.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.gridDS.RowHeadersVisible = False
        Me.gridDS.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        Me.gridDS.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(119, Byte), Integer), CType(CType(182, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.gridDS.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black
        Me.gridDS.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridDS.Size = New System.Drawing.Size(420, 193)
        Me.gridDS.TabIndex = 12
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.labDateUp)
        Me.Panel2.Controls.Add(Me.labDateVer)
        Me.Panel2.Controls.Add(Me.labVer)
        Me.Panel2.Controls.Add(Me.labDS)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(420, 148)
        Me.Panel2.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(24, 91)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Date update:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Date version:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Version:"
        '
        'labDateUp
        '
        Me.labDateUp.AutoSize = True
        Me.labDateUp.Location = New System.Drawing.Point(90, 91)
        Me.labDateUp.Name = "labDateUp"
        Me.labDateUp.Size = New System.Drawing.Size(66, 13)
        Me.labDateUp.TabIndex = 15
        Me.labDateUp.Text = "Date update"
        '
        'labDateVer
        '
        Me.labDateVer.AutoSize = True
        Me.labDateVer.Location = New System.Drawing.Point(90, 69)
        Me.labDateVer.Name = "labDateVer"
        Me.labDateVer.Size = New System.Drawing.Size(67, 13)
        Me.labDateVer.TabIndex = 14
        Me.labDateVer.Text = "Date version"
        '
        'labVer
        '
        Me.labVer.AutoSize = True
        Me.labVer.Location = New System.Drawing.Point(90, 46)
        Me.labVer.Name = "labVer"
        Me.labVer.Size = New System.Drawing.Size(42, 13)
        Me.labVer.TabIndex = 13
        Me.labVer.Text = "Version"
        '
        'labDS
        '
        Me.labDS.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.labDS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labDS.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.labDS.Location = New System.Drawing.Point(0, 131)
        Me.labDS.Name = "labDS"
        Me.labDS.Size = New System.Drawing.Size(420, 17)
        Me.labDS.TabIndex = 12
        Me.labDS.Text = "Danh sách các lần update chương trình:"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.bntClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 341)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(420, 24)
        Me.Panel1.TabIndex = 10
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 0)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(357, 24)
        Me.ProgressBar1.TabIndex = 3
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.Location = New System.Drawing.Point(357, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(63, 24)
        Me.bntClose.TabIndex = 2
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'frmQL_updateCT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(420, 365)
        Me.Controls.Add(Me.gridDS)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmQL_updateCT"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Thông tin cập nhật chương trình"
        CType(Me.gridDS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gridDS As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents labVer As System.Windows.Forms.Label
    Friend WithEvents labDS As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents labDateUp As System.Windows.Forms.Label
    Friend WithEvents labDateVer As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
