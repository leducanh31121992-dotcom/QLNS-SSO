<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportExcel_CanBoInf
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportExcel_CanBoInf))
        Me.gridDS_CB = New System.Windows.Forms.DataGridView()
        Me.bntImportMaCB = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.labDS = New System.Windows.Forms.Label()
        Me.bntImport = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmdChonTM = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboLoaiHS = New System.Windows.Forms.ComboBox()
        Me.txtPath = New System.Windows.Forms.TextBox()
        CType(Me.gridDS_CB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
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
        Me.gridDS_CB.Location = New System.Drawing.Point(0, 0)
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
        Me.gridDS_CB.Size = New System.Drawing.Size(787, 457)
        Me.gridDS_CB.TabIndex = 8
        '
        'bntImportMaCB
        '
        Me.bntImportMaCB.Location = New System.Drawing.Point(677, 34)
        Me.bntImportMaCB.Name = "bntImportMaCB"
        Me.bntImportMaCB.Size = New System.Drawing.Size(90, 23)
        Me.bntImportMaCB.TabIndex = 12
        Me.bntImportMaCB.Text = "Import MaDV"
        Me.bntImportMaCB.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.gridDS_CB)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 88)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(787, 457)
        Me.Panel3.TabIndex = 5
        '
        'labDS
        '
        Me.labDS.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.labDS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labDS.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.labDS.Location = New System.Drawing.Point(0, 71)
        Me.labDS.Name = "labDS"
        Me.labDS.Size = New System.Drawing.Size(787, 17)
        Me.labDS.TabIndex = 11
        Me.labDS.Text = "Danh sách cán bộ được import vào CSDL"
        '
        'bntImport
        '
        Me.bntImport.Location = New System.Drawing.Point(578, 34)
        Me.bntImport.Name = "bntImport"
        Me.bntImport.Size = New System.Drawing.Size(75, 23)
        Me.bntImport.TabIndex = 10
        Me.bntImport.Text = "Import"
        Me.bntImport.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "File excel cán bộ"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 3)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(787, 15)
        Me.ProgressBar1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.ProgressBar1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 545)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(787, 18)
        Me.Panel2.TabIndex = 4
        '
        'cmdChonTM
        '
        Me.cmdChonTM.Image = Global.QLNS.My.Resources.Resources.chon_file
        Me.cmdChonTM.Location = New System.Drawing.Point(527, 36)
        Me.cmdChonTM.Name = "cmdChonTM"
        Me.cmdChonTM.Size = New System.Drawing.Size(30, 22)
        Me.cmdChonTM.TabIndex = 8
        Me.cmdChonTM.Text = "..."
        Me.cmdChonTM.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cboLoaiHS)
        Me.Panel1.Controls.Add(Me.bntImportMaCB)
        Me.Panel1.Controls.Add(Me.labDS)
        Me.Panel1.Controls.Add(Me.bntImport)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cmdChonTM)
        Me.Panel1.Controls.Add(Me.txtPath)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(787, 88)
        Me.Panel1.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 17)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Loại hồ sơ cán bộ"
        '
        'cboLoaiHS
        '
        Me.cboLoaiHS.FormattingEnabled = True
        Me.cboLoaiHS.Items.AddRange(New Object() {"1. Thông tin chung của cán bộ", "2. Quyết định Tuyển dụng, Điều động, Bổ nhiệm, Miễn nhiệm", "3. Quyết định lương", "4. Hồ sơ đào tạo, văn bằng, chứng chỉ", "5. Thông tin gia đình cán bộ", "6. Các quyết định khác", "7. Hợp đồng lao động", "8. Thông tin tham gia lực lượng vũ trang", "9. Hồ sơ kỉ luật", " "})
        Me.cboLoaiHS.Location = New System.Drawing.Point(110, 13)
        Me.cboLoaiHS.Name = "cboLoaiHS"
        Me.cboLoaiHS.Size = New System.Drawing.Size(401, 21)
        Me.cboLoaiHS.TabIndex = 1
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(110, 37)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(401, 21)
        Me.txtPath.TabIndex = 7
        '
        'frmImportExcel_CanBoInf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmImportExcel_CanBoInf"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "QUẢN LÝ NHÂN SỰ: Nhập thông tin cán bộ từ file excel"
        CType(Me.gridDS_CB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gridDS_CB As System.Windows.Forms.DataGridView
    Friend WithEvents bntImportMaCB As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents labDS As System.Windows.Forms.Label
    Friend WithEvents bntImport As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdChonTM As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboLoaiHS As System.Windows.Forms.ComboBox
End Class
