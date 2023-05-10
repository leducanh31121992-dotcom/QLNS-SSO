<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExChamCong
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExChamCong))
        Me.bntExBangChamCong = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'bntExBangChamCong
        '
        Me.bntExBangChamCong.Location = New System.Drawing.Point(65, 169)
        Me.bntExBangChamCong.Name = "bntExBangChamCong"
        Me.bntExBangChamCong.Size = New System.Drawing.Size(147, 23)
        Me.bntExBangChamCong.TabIndex = 0
        Me.bntExBangChamCong.Text = "Export Bảng chấm công"
        Me.bntExBangChamCong.UseVisualStyleBackColor = True
        '
        'frmExChamCong
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.bntExBangChamCong)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmExChamCong"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "frmExChamCong"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents bntExBangChamCong As System.Windows.Forms.Button
End Class
