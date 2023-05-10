<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportHSCB
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportHSCB))
        Me.lab1 = New System.Windows.Forms.Label()
        Me.lab2 = New System.Windows.Forms.Label()
        Me.lab3 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.bntImportHS = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lab1
        '
        Me.lab1.AutoSize = True
        Me.lab1.Location = New System.Drawing.Point(20, 13)
        Me.lab1.Name = "lab1"
        Me.lab1.Size = New System.Drawing.Size(0, 13)
        Me.lab1.TabIndex = 0
        '
        'lab2
        '
        Me.lab2.AutoSize = True
        Me.lab2.Location = New System.Drawing.Point(20, 38)
        Me.lab2.Name = "lab2"
        Me.lab2.Size = New System.Drawing.Size(0, 13)
        Me.lab2.TabIndex = 1
        '
        'lab3
        '
        Me.lab3.AutoSize = True
        Me.lab3.Location = New System.Drawing.Point(20, 118)
        Me.lab3.Name = "lab3"
        Me.lab3.Size = New System.Drawing.Size(0, 13)
        Me.lab3.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(22, 64)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(310, 12)
        Me.ProgressBar1.TabIndex = 9
        '
        'bntImportHS
        '
        Me.bntImportHS.Location = New System.Drawing.Point(22, 82)
        Me.bntImportHS.Name = "bntImportHS"
        Me.bntImportHS.Size = New System.Drawing.Size(175, 23)
        Me.bntImportHS.TabIndex = 8
        Me.bntImportHS.Text = "Nhập hồ sơ cán bộ chuyển tới"
        Me.bntImportHS.UseVisualStyleBackColor = True
        '
        'frmImportHSCB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(355, 148)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.bntImportHS)
        Me.Controls.Add(Me.lab3)
        Me.Controls.Add(Me.lab2)
        Me.Controls.Add(Me.lab1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmImportHSCB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Nhập hồ sơ cán bộ chuyển tới"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lab1 As System.Windows.Forms.Label
    Friend WithEvents lab2 As System.Windows.Forms.Label
    Friend WithEvents lab3 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bntImportHS As System.Windows.Forms.Button
End Class
