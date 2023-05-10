<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctlTest
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.lab = New System.Windows.Forms.Label
        Me.cbDuyet = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'lab
        '
        Me.lab.AutoSize = True
        Me.lab.Location = New System.Drawing.Point(4, 6)
        Me.lab.Name = "lab"
        Me.lab.Size = New System.Drawing.Size(21, 13)
        Me.lab.TabIndex = 0
        Me.lab.Text = "lab"
        '
        'cbDuyet
        '
        Me.cbDuyet.AutoSize = True
        Me.cbDuyet.Location = New System.Drawing.Point(28, 5)
        Me.cbDuyet.Name = "cbDuyet"
        Me.cbDuyet.Size = New System.Drawing.Size(54, 17)
        Me.cbDuyet.TabIndex = 1
        Me.cbDuyet.Text = "Duyệt"
        Me.cbDuyet.UseVisualStyleBackColor = True
        '
        'ctlTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cbDuyet)
        Me.Controls.Add(Me.lab)
        Me.Name = "ctlTest"
        Me.Size = New System.Drawing.Size(79, 25)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lab As System.Windows.Forms.Label
    Friend WithEvents cbDuyet As System.Windows.Forms.CheckBox

End Class
