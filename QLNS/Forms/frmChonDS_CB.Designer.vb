<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChonDS_CB
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChonDS_CB))
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.bntChoiseCB = New System.Windows.Forms.Button()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.treeCocau = New System.Windows.Forms.TreeView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rdCC = New System.Windows.Forms.RadioButton()
        Me.rdVH = New System.Windows.Forms.RadioButton()
        Me.rdTS = New System.Windows.Forms.RadioButton()
        Me.rdNH = New System.Windows.Forms.RadioButton()
        Me.rdCB = New System.Windows.Forms.RadioButton()
        Me.Panel11.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.DarkGray
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel11.Controls.Add(Me.bntChoiseCB)
        Me.Panel11.Controls.Add(Me.bntClose)
        Me.Panel11.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel11.Location = New System.Drawing.Point(0, 451)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(274, 26)
        Me.Panel11.TabIndex = 11
        '
        'bntChoiseCB
        '
        Me.bntChoiseCB.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntChoiseCB.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntChoiseCB.Location = New System.Drawing.Point(148, 0)
        Me.bntChoiseCB.Name = "bntChoiseCB"
        Me.bntChoiseCB.Size = New System.Drawing.Size(61, 22)
        Me.bntChoiseCB.TabIndex = 1
        Me.bntChoiseCB.Text = "&Chọn"
        Me.bntChoiseCB.UseVisualStyleBackColor = True
        '
        'bntClose
        '
        Me.bntClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.bntClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bntClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bntClose.Location = New System.Drawing.Point(209, 0)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(61, 22)
        Me.bntClose.TabIndex = 5
        Me.bntClose.Text = "&Quay ra"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.treeCocau)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(274, 451)
        Me.Panel1.TabIndex = 136
        '
        'treeCocau
        '
        Me.treeCocau.CheckBoxes = True
        Me.treeCocau.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeCocau.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.treeCocau.Location = New System.Drawing.Point(0, 0)
        Me.treeCocau.Name = "treeCocau"
        Me.treeCocau.Size = New System.Drawing.Size(274, 358)
        Me.treeCocau.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.rdCC)
        Me.Panel2.Controls.Add(Me.rdVH)
        Me.Panel2.Controls.Add(Me.rdTS)
        Me.Panel2.Controls.Add(Me.rdNH)
        Me.Panel2.Controls.Add(Me.rdCB)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 358)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(274, 93)
        Me.Panel2.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(0, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(142, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Hiển thị danh sách theo:"
        '
        'rdCC
        '
        Me.rdCC.AutoSize = True
        Me.rdCC.Location = New System.Drawing.Point(3, 73)
        Me.rdCC.Name = "rdCC"
        Me.rdCC.Size = New System.Drawing.Size(252, 18)
        Me.rdCC.TabIndex = 4
        Me.rdCC.Text = "CB đã chuyển tới đơn vị khác trong NHCS"
        Me.rdCC.UseVisualStyleBackColor = True
        '
        'rdVH
        '
        Me.rdVH.AutoSize = True
        Me.rdVH.Location = New System.Drawing.Point(172, 49)
        Me.rdVH.Name = "rdVH"
        Me.rdVH.Size = New System.Drawing.Size(99, 18)
        Me.rdVH.TabIndex = 3
        Me.rdVH.Text = "CB đã về hưu"
        Me.rdVH.UseVisualStyleBackColor = True
        '
        'rdTS
        '
        Me.rdTS.AutoSize = True
        Me.rdTS.Location = New System.Drawing.Point(3, 49)
        Me.rdTS.Name = "rdTS"
        Me.rdTS.Size = New System.Drawing.Size(88, 18)
        Me.rdTS.TabIndex = 2
        Me.rdTS.Text = "CB thử việc"
        Me.rdTS.UseVisualStyleBackColor = True
        '
        'rdNH
        '
        Me.rdNH.AutoSize = True
        Me.rdNH.Location = New System.Drawing.Point(172, 25)
        Me.rdNH.Name = "rdNH"
        Me.rdNH.Size = New System.Drawing.Size(94, 18)
        Me.rdNH.TabIndex = 1
        Me.rdNH.Text = "CB ngắn hạn"
        Me.rdNH.UseVisualStyleBackColor = True
        '
        'rdCB
        '
        Me.rdCB.AutoSize = True
        Me.rdCB.Checked = True
        Me.rdCB.Location = New System.Drawing.Point(3, 25)
        Me.rdCB.Name = "rdCB"
        Me.rdCB.Size = New System.Drawing.Size(159, 18)
        Me.rdCB.TabIndex = 0
        Me.rdCB.TabStop = True
        Me.rdCB.Text = "CB đang công tác tại ĐV"
        Me.rdCB.UseVisualStyleBackColor = True
        '
        'frmChonDS_CB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(274, 477)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel11)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmChonDS_CB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chọn danh sách cán bộ"
        Me.Panel11.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents bntChoiseCB As System.Windows.Forms.Button
    Friend WithEvents bntClose As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents treeCocau As System.Windows.Forms.TreeView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rdCC As System.Windows.Forms.RadioButton
    Friend WithEvents rdVH As System.Windows.Forms.RadioButton
    Friend WithEvents rdTS As System.Windows.Forms.RadioButton
    Friend WithEvents rdNH As System.Windows.Forms.RadioButton
    Friend WithEvents rdCB As System.Windows.Forms.RadioButton
End Class
