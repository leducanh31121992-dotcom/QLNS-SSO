<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExportHSCB
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExportHSCB))
        Me.labCB = New System.Windows.Forms.Label()
        Me.labCN = New System.Windows.Forms.Label()
        Me.bntExportHS = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.labErr = New System.Windows.Forms.Label()
        Me.labDir = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'labCB
        '
        Me.labCB.AutoSize = True
        Me.labCB.Location = New System.Drawing.Point(20, 13)
        Me.labCB.Name = "labCB"
        Me.labCB.Size = New System.Drawing.Size(41, 13)
        Me.labCB.TabIndex = 0
        Me.labCB.Text = "Cán bộ"
        '
        'labCN
        '
        Me.labCN.AutoSize = True
        Me.labCN.Location = New System.Drawing.Point(20, 38)
        Me.labCN.Name = "labCN"
        Me.labCN.Size = New System.Drawing.Size(55, 13)
        Me.labCN.TabIndex = 1
        Me.labCN.Text = "Chi nhánh"
        '
        'bntExportHS
        '
        Me.bntExportHS.Location = New System.Drawing.Point(23, 82)
        Me.bntExportHS.Name = "bntExportHS"
        Me.bntExportHS.Size = New System.Drawing.Size(175, 23)
        Me.bntExportHS.TabIndex = 2
        Me.bntExportHS.Text = "Xuất hồ sơ cán bộ để chuyển đi"
        Me.bntExportHS.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(23, 64)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(307, 12)
        Me.ProgressBar1.TabIndex = 7
        '
        'labErr
        '
        Me.labErr.AutoSize = True
        Me.labErr.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labErr.Location = New System.Drawing.Point(20, 117)
        Me.labErr.Name = "labErr"
        Me.labErr.Size = New System.Drawing.Size(0, 13)
        Me.labErr.TabIndex = 8
        '
        'labDir
        '
        Me.labDir.AutoSize = True
        Me.labDir.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labDir.Location = New System.Drawing.Point(20, 139)
        Me.labDir.Name = "labDir"
        Me.labDir.Size = New System.Drawing.Size(0, 13)
        Me.labDir.TabIndex = 9
        '
        'frmExportHSCB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(355, 161)
        Me.Controls.Add(Me.labDir)
        Me.Controls.Add(Me.labErr)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.bntExportHS)
        Me.Controls.Add(Me.labCN)
        Me.Controls.Add(Me.labCB)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmExportHSCB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Xuất hồ sơ cho cán bộ chuyển chi nhánh"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents labCB As System.Windows.Forms.Label
    Friend WithEvents labCN As System.Windows.Forms.Label
    Friend WithEvents bntExportHS As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents labErr As System.Windows.Forms.Label
    Friend WithEvents labDir As System.Windows.Forms.Label
End Class
