<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHT_ChotDuLieu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DGV = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboChonKy = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnChotDuLieu = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblKyGanNhat = New System.Windows.Forms.Label()
        Me.POS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DonVi = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ThoiGian = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DGV
        '
        Me.DGV.AllowUserToAddRows = False
        Me.DGV.AllowUserToDeleteRows = False
        Me.DGV.AllowUserToResizeColumns = False
        Me.DGV.AllowUserToResizeRows = False
        Me.DGV.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.DGV.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.DGV.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.DGV.ColumnHeadersHeight = 34
        Me.DGV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.POS, Me.DonVi, Me.ThoiGian})
        Me.DGV.Location = New System.Drawing.Point(12, 140)
        Me.DGV.MultiSelect = False
        Me.DGV.Name = "DGV"
        Me.DGV.ReadOnly = True
        Me.DGV.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.DGV.RowHeadersVisible = False
        Me.DGV.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV.Size = New System.Drawing.Size(500, 430)
        Me.DGV.TabIndex = 76
        Me.DGV.Tag = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(502, 19)
        Me.Label3.TabIndex = 77
        Me.Label3.Text = "Danh sách đơn vị đã thực hiện chốt dữ liệu theo kỳ"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboChonKy
        '
        Me.cboChonKy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChonKy.FormattingEnabled = True
        Me.cboChonKy.Location = New System.Drawing.Point(96, 24)
        Me.cboChonKy.Name = "cboChonKy"
        Me.cboChonKy.Size = New System.Drawing.Size(180, 21)
        Me.cboChonKy.TabIndex = 78
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(44, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 79
        Me.Label1.Text = "Chọn kỳ"
        '
        'btnChotDuLieu
        '
        Me.btnChotDuLieu.Location = New System.Drawing.Point(334, 22)
        Me.btnChotDuLieu.Name = "btnChotDuLieu"
        Me.btnChotDuLieu.Size = New System.Drawing.Size(124, 23)
        Me.btnChotDuLieu.TabIndex = 80
        Me.btnChotDuLieu.Text = "Chốt dữ liệu"
        Me.btnChotDuLieu.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnChotDuLieu)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cboChonKy)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 38)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(500, 74)
        Me.GroupBox1.TabIndex = 81
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(318, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(163, 19)
        Me.Label2.TabIndex = 81
        Me.Label2.Text = "(thời gian được lấy theo server)"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(181, 19)
        Me.Label4.TabIndex = 82
        Me.Label4.Text = "Kỳ chốt dữ liệu gần nhất của đơn vị:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblKyGanNhat
        '
        Me.lblKyGanNhat.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKyGanNhat.Location = New System.Drawing.Point(199, 16)
        Me.lblKyGanNhat.Name = "lblKyGanNhat"
        Me.lblKyGanNhat.Size = New System.Drawing.Size(152, 19)
        Me.lblKyGanNhat.TabIndex = 83
        Me.lblKyGanNhat.Text = "..."
        Me.lblKyGanNhat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'POS
        '
        Me.POS.HeaderText = "POS"
        Me.POS.Name = "POS"
        Me.POS.ReadOnly = True
        Me.POS.Width = 80
        '
        'DonVi
        '
        Me.DonVi.HeaderText = "Tên đơn vị"
        Me.DonVi.Name = "DonVi"
        Me.DonVi.ReadOnly = True
        Me.DonVi.Width = 280
        '
        'ThoiGian
        '
        Me.ThoiGian.HeaderText = "Thời điểm chốt"
        Me.ThoiGian.Name = "ThoiGian"
        Me.ThoiGian.ReadOnly = True
        Me.ThoiGian.Width = 150
        '
        'frmHT_ChotDuLieu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(524, 582)
        Me.Controls.Add(Me.lblKyGanNhat)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DGV)
        Me.Name = "frmHT_ChotDuLieu"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chốt dữ liệu"
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DGV As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboChonKy As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnChotDuLieu As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblKyGanNhat As System.Windows.Forms.Label
    Friend WithEvents POS As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DonVi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ThoiGian As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
