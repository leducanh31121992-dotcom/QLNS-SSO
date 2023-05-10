<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChuyenHSCB_CN_ve_TW
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChuyenHSCB_CN_ve_TW))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.labDS = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.bntConvert = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.gridDS_CB = New System.Windows.Forms.DataGridView()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.gridDS_CB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.labDS)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(808, 34)
        Me.Panel2.TabIndex = 2
        '
        'labDS
        '
        Me.labDS.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.labDS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labDS.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.labDS.Location = New System.Drawing.Point(0, 17)
        Me.labDS.Name = "labDS"
        Me.labDS.Size = New System.Drawing.Size(808, 17)
        Me.labDS.TabIndex = 12
        Me.labDS.Text = "Danh sách cán bộ chuyển hồ sơ về Hội sở chính quản lý"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.bntConvert)
        Me.Panel1.Controls.Add(Me.bntClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 537)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(808, 24)
        Me.Panel1.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 0)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(601, 24)
        Me.ProgressBar1.TabIndex = 3
        '
        'bntConvert
        '
        Me.bntConvert.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntConvert.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntConvert.Location = New System.Drawing.Point(601, 0)
        Me.bntConvert.Name = "bntConvert"
        Me.bntConvert.Size = New System.Drawing.Size(144, 24)
        Me.bntConvert.TabIndex = 1
        Me.bntConvert.Text = "&Chuyển"
        Me.bntConvert.UseVisualStyleBackColor = True
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.Location = New System.Drawing.Point(745, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(63, 24)
        Me.bntClose.TabIndex = 2
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'gridDS_CB
        '
        Me.gridDS_CB.AllowUserToAddRows = False
        Me.gridDS_CB.AllowUserToDeleteRows = False
        Me.gridDS_CB.AllowUserToResizeRows = False
        Me.gridDS_CB.BackgroundColor = System.Drawing.Color.White
        Me.gridDS_CB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.gridDS_CB.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.DarkGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridDS_CB.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.gridDS_CB.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.gridDS_CB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gridDS_CB.Location = New System.Drawing.Point(0, 34)
        Me.gridDS_CB.Name = "gridDS_CB"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.gridDS_CB.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.gridDS_CB.RowHeadersVisible = False
        Me.gridDS_CB.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        Me.gridDS_CB.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(119, Byte), Integer), CType(CType(182, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.gridDS_CB.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black
        Me.gridDS_CB.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gridDS_CB.Size = New System.Drawing.Size(808, 503)
        Me.gridDS_CB.TabIndex = 10
        '
        'frmChuyenHSCB_CN_ve_TW
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(808, 561)
        Me.Controls.Add(Me.gridDS_CB)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmChuyenHSCB_CN_ve_TW"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chuyển hồ sơ cán bộ thuộc TW quản lý về HSC"
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.gridDS_CB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents labDS As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bntConvert As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents gridDS_CB As System.Windows.Forms.DataGridView
End Class
