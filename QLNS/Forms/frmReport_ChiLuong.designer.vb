<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReport_ChiLuong
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReport_ChiLuong))
        Me.rptView = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.rpt_ChiLuong1KY = New QLNS.ChiLuong1KY()
        Me.rpt_ChiLuongTS = New QLNS.ChiLuongThuViec()
        Me.rpt_ChiLuongKY1 = New QLNS.ChiLuongKY1()
        Me.rpt_ChiLuongHDNH = New QLNS.ChiLuongHDNganHan()
        Me.rpt_ChiLuongKY2 = New QLNS.ChiLuongKY2()
        Me.rpt_ChiThemGio = New QLNS.ChiThemGio()
        Me.rpt_ChiBoSung = New QLNS.ChiBoSung()
        Me.rpt_ChiBoSung_NH = New QLNS.ChiBoSung_NH()
        Me.rpt_ChiBoSung_TS = New QLNS.ChiBoSung_TS()
        Me.rpt_ChiLuong1KY_NKL = New QLNS.ChiLuong1KY_NKL()
        Me.rpt_ChiLuongKY2_NKL = New QLNS.ChiLuongKY2_NKL()
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
        Me.rptView.ShowRefreshButton = False
        Me.rptView.Size = New System.Drawing.Size(787, 563)
        Me.rptView.TabIndex = 5
        Me.rptView.UseWaitCursor = True
        Me.rptView.ViewTimeSelectionFormula = ""
        '
        'rpt_ChiLuong1KY
        '
        Me.rpt_ChiLuong1KY.FullResourceName = "QLNS.ChiLuong1KY.rpt"
        Me.rpt_ChiLuong1KY.NewGenerator = True
        '
        'rpt_ChiLuongTS
        '
        Me.rpt_ChiLuongTS.FullResourceName = "QLNS.ChiLuongThuViec.rpt"
        Me.rpt_ChiLuongTS.NewGenerator = True
        '
        'rpt_ChiLuongKY1
        '
        Me.rpt_ChiLuongKY1.FullResourceName = "QLNS.ChiLuongKY1.rpt"
        Me.rpt_ChiLuongKY1.NewGenerator = True
        '
        'rpt_ChiLuongHDNH
        '
        Me.rpt_ChiLuongHDNH.FullResourceName = "QLNS.ChiLuongHDNganHan.rpt"
        Me.rpt_ChiLuongHDNH.NewGenerator = True
        '
        'rpt_ChiLuongKY2
        '
        Me.rpt_ChiLuongKY2.FullResourceName = "QLNS.ChiLuongKY2.rpt"
        Me.rpt_ChiLuongKY2.NewGenerator = True
        '
        'rpt_ChiThemGio
        '
        Me.rpt_ChiThemGio.FullResourceName = "QLNS.ChiThemGio.rpt"
        Me.rpt_ChiThemGio.NewGenerator = True
        '
        'rpt_ChiBoSung
        '
        Me.rpt_ChiBoSung.FullResourceName = "QLNS.ChiBoSung.rpt"
        Me.rpt_ChiBoSung.NewGenerator = True
        '
        'rpt_ChiBoSung_NH
        '
        Me.rpt_ChiBoSung_NH.FullResourceName = "QLNS.ChiBoSung_NH.rpt"
        Me.rpt_ChiBoSung_NH.NewGenerator = True
        '
        'rpt_ChiBoSung_TS
        '
        Me.rpt_ChiBoSung_TS.FullResourceName = "QLNS.ChiBoSung_TS.rpt"
        Me.rpt_ChiBoSung_TS.NewGenerator = True
        '
        'rpt_ChiLuong1KY_NKL
        '
        Me.rpt_ChiLuong1KY_NKL.FullResourceName = "QLNS.ChiLuong1KY_NKL.rpt"
        Me.rpt_ChiLuong1KY_NKL.NewGenerator = True
        '
        'rpt_ChiLuongKY2_NKL
        '
        Me.rpt_ChiLuongKY2_NKL.FullResourceName = "QLNS.ChiLuongKY2_NKL.rpt"
        Me.rpt_ChiLuongKY2_NKL.NewGenerator = True
        '
        'frmReport_ChiLuong
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(787, 563)
        Me.Controls.Add(Me.rptView)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmReport_ChiLuong"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản lý nhân sự: BẢNG KÊ CHI LƯƠNG THEO THÁNG"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rpt_ChiLuong1KY As QLNS.ChiLuong1KY
    Friend WithEvents rpt_ChiLuongKY1 As QLNS.ChiLuongKY1
    Friend WithEvents rpt_ChiLuongKY2 As QLNS.ChiLuongKY2
    Friend WithEvents rpt_ChiLuongHDNH As QLNS.ChiLuongHDNganHan
    Friend WithEvents rpt_ChiLuongTS As QLNS.ChiLuongThuViec
    Friend WithEvents rpt_ChiThemGio As QLNS.ChiThemGio
    Friend WithEvents rpt_ChiBoSung As QLNS.ChiBoSung
    Friend WithEvents rpt_ChiBoSung_NH As QLNS.ChiBoSung_NH
    Friend WithEvents rpt_ChiBoSung_TS As QLNS.ChiBoSung_TS
    Friend WithEvents rpt_ChiLuong1KY_NKL As QLNS.ChiLuong1KY_NKL
    Friend WithEvents rpt_ChiLuongKY2_NKL As QLNS.ChiLuongKY2_NKL
    Private WithEvents rptView As CrystalDecisions.Windows.Forms.CrystalReportViewer

End Class
