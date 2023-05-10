<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReport_KhenThuong
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
        Me.rptView = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.rpt_KhenThuongTT = New QLNS.KhenThuong_TapThe()
        Me.rpt_KhenThuongCN = New QLNS.KhenThuong_CaNhan()
        Me.rpt_KhenThuongCN_TQ = New QLNS.KhenThuong_CaNhan_CN()
        Me.rpt_KhenThuongTT_TQ = New QLNS.KhenThuong_TapThe_CN()
        Me.SuspendLayout()
        '
        'rptView
        '
        Me.rptView.ActiveViewIndex = -1
        Me.rptView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rptView.DisplayGroupTree = False
        Me.rptView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rptView.Location = New System.Drawing.Point(0, 0)
        Me.rptView.Name = "rptView"
        Me.rptView.SelectionFormula = ""
        Me.rptView.Size = New System.Drawing.Size(785, 561)
        Me.rptView.TabIndex = 0
        Me.rptView.UseWaitCursor = True
        Me.rptView.ViewTimeSelectionFormula = ""
        '
        'rpt_KhenThuongTT
        '
        Me.rpt_KhenThuongTT.FullResourceName = "QLNS.KhenThuong_TapThe.rpt"
        Me.rpt_KhenThuongTT.NewGenerator = True
        '
        'rpt_KhenThuongCN
        '
        Me.rpt_KhenThuongCN.FullResourceName = "QLNS.KhenThuong_CaNhan.rpt"
        Me.rpt_KhenThuongCN.NewGenerator = True
        '
        'rpt_KhenThuongCN_TQ
        '
        Me.rpt_KhenThuongCN_TQ.FullResourceName = "QLNS.KhenThuong_CaNhan_CN.rpt"
        Me.rpt_KhenThuongCN_TQ.NewGenerator = True
        '
        'rpt_KhenThuongTT_TQ
        '
        Me.rpt_KhenThuongTT_TQ.FullResourceName = "QLNS.KhenThuong_TapThe_CN.rpt"
        Me.rpt_KhenThuongTT_TQ.NewGenerator = True
        '
        'frmReport_KhenThuong
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(785, 561)
        Me.Controls.Add(Me.rptView)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Name = "frmReport_KhenThuong"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: DANH SÁCH KHEN THƯỞNG"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rpt_KhenThuongCN As QLNS.KhenThuong_CaNhan
    Friend WithEvents rpt_KhenThuongTT As QLNS.KhenThuong_TapThe
    Friend WithEvents rpt_KhenThuongCN_TQ As QLNS.KhenThuong_CaNhan_CN
    Friend WithEvents rpt_KhenThuongTT_TQ As QLNS.KhenThuong_TapThe_CN
    Private WithEvents rptView As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
