Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms

Public Class frmReport_ChiLuong
    Private _IdChiLuong As String
    Private _IDLoaiThuChi As Integer
    Private _MaLoaiThuChi As String
    Private _Ky As Int16
    Private _TK As Boolean  '= true: Report hien thi TT Tai khoan'
    Private _IdDonVi As Integer
    Private _Thang As Integer
    Private _Nam As Integer
    Private _LoaiCB As Int16
    Private _LapBieu As String
    Private _KiemSoat As String
    Private _GiamDoc As String
    Private _NoiDung As String
    Private _GhiChu As String

    Public Property IdChiLuong() As String
        Get
            Return _IdChiLuong
        End Get
        Set(ByVal value As String)
            _IdChiLuong = value
        End Set
    End Property

    Public Property IDLoaiThuChi() As Integer
        Get
            Return _IDLoaiThuChi
        End Get
        Set(ByVal value As Integer)
            _IDLoaiThuChi = value
        End Set
    End Property

    Public Property MaLoaiThuChi() As String
        Get
            Return _MaLoaiThuChi
        End Get
        Set(ByVal value As String)
            _MaLoaiThuChi = value
        End Set
    End Property

    Public Property Ky() As Int16
        Get
            Return _Ky
        End Get
        Set(ByVal value As Int16)
            _Ky = value
        End Set
    End Property

    Public Property TK() As Boolean
        Get
            Return _TK
        End Get
        Set(ByVal value As Boolean)
            _TK = value
        End Set
    End Property

    Public Property ID_DonVi() As Integer
        Get
            Return _IdDonVi
        End Get
        Set(ByVal value As Integer)
            _IdDonVi = value
        End Set
    End Property

    Public Property Thang() As Integer
        Get
            Return _Thang
        End Get
        Set(ByVal value As Integer)
            _Thang = value
        End Set
    End Property

    Public Property Nam() As Integer
        Get
            Return _Nam
        End Get
        Set(ByVal value As Integer)
            _Nam = value
        End Set
    End Property

    Public Property LoaiCB() As Int16
        Get
            Return _LoaiCB
        End Get
        Set(ByVal value As Int16)
            _LoaiCB = value
        End Set
    End Property

    Public Property LapBieu() As String
        Get
            Return _LapBieu
        End Get
        Set(ByVal value As String)
            _LapBieu = value
        End Set
    End Property

    Public Property KiemSoat() As String
        Get
            Return _KiemSoat
        End Get
        Set(ByVal value As String)
            _KiemSoat = value
        End Set
    End Property

    Public Property GiamDoc() As String
        Get
            Return _GiamDoc
        End Get
        Set(ByVal value As String)
            _GiamDoc = value
        End Set
    End Property

    Public Property Noidung() As String
        Get
            Return _NoiDung
        End Get
        Set(ByVal value As String)
            _NoiDung = value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal value As String)
            _GhiChu = value
        End Set
    End Property

    Private Sub frmReport_ChiLuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Select Case MaLoaiThuChi

                Case "4201"
                    showReportChiLuong(IdChiLuong, LoaiCB, Ky, ID_DonVi, Thang, Nam, TK, LapBieu, KiemSoat, GiamDoc, Noidung, GhiChu)
                    'If LoaiCB = 0 Then
                    '    showReportChiLuong(IdChiLuong, 0, Ky, ID_DonVi, Thang, Nam, TK, LapBieu, KiemSoat, GiamDoc, Noidung, GhiChu)
                    'Else
                    '    If LoaiCB = 1 Then
                    '        showReportChiLuong(IdChiLuong, 1, Ky, ID_DonVi, Thang, Nam, TK, LapBieu, KiemSoat, GiamDoc, Noidung, GhiChu)
                    '        'showReportChiLuong_HDNH(IdChiLuong, ID_DonVi, Thang, Nam, LapBieu, KiemSoat, GiamDoc, GhiChu)
                    '    Else
                    '        showReportChiLuong_TS(IdChiLuong, ID_DonVi, Thang, Nam, LapBieu, KiemSoat, GiamDoc, GhiChu)
                    '    End If
                    'End If
                Case "4204"
                    showReportChiThemGio(IdChiLuong, TK, ID_DonVi, Thang, Nam, LapBieu, KiemSoat, GiamDoc, GhiChu)
                Case Else
                    showReportChiBoSung(IdChiLuong, LoaiCB, ID_DonVi, Thang, Nam, LapBieu, KiemSoat, GiamDoc, GhiChu)
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    Private Sub frmBangKeChiLuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub showReportChiLuong(ByVal IdChiLuong As String, ByVal vNH_TS As Int16, ByVal vKy As Int16, _
                                 ByVal vIdDonVi As Integer, ByVal vThang As Integer, ByVal vNam As Integer, ByVal vTK As Boolean, _
                                 ByVal vLapBieu As String, ByVal vKiemSoat As String, ByVal vGiamDoc As String, ByVal vNoidung As String, ByVal vGhiChu As String)

        Dim tCMNV As Double = 0
        Dim tPC_CV As Double = 0
        Dim tPC_TN As Double = 0
        Dim tPC_KV As Double = 0
        Dim tPC_DH As Double = 0
        Dim tPC_TH As Double = 0
        Dim vIdPhong As String = ""
        Dim vkindPB As String = ""
        Dim vIDPB As Integer = 0
        Dim tTong100, tT_Nghi, tTongChi, tT_TNCNtmp, tT_DFCD, tT_BHXH, tT_BHYT, tT_BHTN, tT_TamUng, tT_CacKhoanThuKhac, tTongTru, tT_SotienTH, tThucLinh As Long
        Dim dt As DataTable = New DataTable
        dt = printBangChiLuong(IdChiLuong, vNH_TS, vKy, vTK, vIdPhong, tCMNV, tPC_CV, tPC_TN, tPC_KV, tPC_DH, tTong100, tT_Nghi, tTongChi, tT_TNCNtmp, tT_DFCD, tT_BHXH, tT_BHYT, tT_BHTN, tT_TamUng, tT_CacKhoanThuKhac, tTongTru, tPC_TH, tT_SotienTH, tThucLinh)
        If vIdPhong <> "" Then
            vkindPB = vIdPhong.Substring(0, 2)
            vIDPB = vIdPhong.Substring(3)
        Else
            vkindPB = "PB"
            vIDPB = 0
        End If

        If vNH_TS = 0 Then
            'Hiển thị báo cáo cho lao động chính thức
            If vKy = 0 Then
                If tT_Nghi = 0 Then
                    rpt_ChiLuong1KY.SetDataSource(dt)
                    rpt_ChiLuong1KY.SetParameterValue("TC4", tCMNV)
                    rpt_ChiLuong1KY.SetParameterValue("TC5", tPC_CV)
                    rpt_ChiLuong1KY.SetParameterValue("TC6", tPC_KV)
                    rpt_ChiLuong1KY.SetParameterValue("TC8", tPC_DH)
                    rpt_ChiLuong1KY.SetParameterValue("TC9", tPC_TN)
                    rpt_ChiLuong1KY.SetParameterValue("TN1", tTong100)
                    rpt_ChiLuong1KY.SetParameterValue("TN2", tTongChi)
                    rpt_ChiLuong1KY.SetParameterValue("TN3", tT_TNCNtmp)
                    rpt_ChiLuong1KY.SetParameterValue("TN4", tT_DFCD)
                    rpt_ChiLuong1KY.SetParameterValue("TN5", tT_BHXH)
                    rpt_ChiLuong1KY.SetParameterValue("TN6", tT_BHYT)
                    rpt_ChiLuong1KY.SetParameterValue("TN7", tT_TamUng)
                    rpt_ChiLuong1KY.SetParameterValue("TN9", tTongTru)
                    rpt_ChiLuong1KY.SetParameterValue("TN10", tThucLinh)
                    rpt_ChiLuong1KY.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                    rpt_ChiLuong1KY.SetParameterValue("TC15", tPC_TH)
                    rpt_ChiLuong1KY.SetParameterValue("TN14", tT_SotienTH)
                    rpt_ChiLuong1KY.SetParameterValue("TN13", tT_BHTN)
                    rpt_ChiLuong1KY.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                    If vkindPB = "PB" Then
                        If (vIDPB = 0 And vIdDonVi > 5) Then
                            rpt_ChiLuong1KY.SetParameterValue("Phong", "Hội sở tỉnh")
                        Else
                            rpt_ChiLuong1KY.SetParameterValue("Phong", getPhong(vIDPB))
                        End If
                    Else
                        rpt_ChiLuong1KY.SetParameterValue("Phong", getDonvi(vIDPB))
                    End If
                    rpt_ChiLuong1KY.SetParameterValue("tinh", DIABAN)
                    rpt_ChiLuong1KY.SetParameterValue("ng", Now.Day)
                    rpt_ChiLuong1KY.SetParameterValue("th", Now.Month)
                    rpt_ChiLuong1KY.SetParameterValue("nm", Now.Year)
                    rpt_ChiLuong1KY.SetParameterValue("Ky", vKy)
                    rpt_ChiLuong1KY.SetParameterValue("thang", vThang)
                    rpt_ChiLuong1KY.SetParameterValue("nam", vNam)
                    If DONVI = gMaDonViTW Then
                        rpt_ChiLuong1KY.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                    Else
                        rpt_ChiLuong1KY.SetParameterValue("labGD", "GIÁM ĐỐC")
                    End If
                    rpt_ChiLuong1KY.SetParameterValue("LAPBANG", vLapBieu.Trim)
                    rpt_ChiLuong1KY.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                    rpt_ChiLuong1KY.SetParameterValue("GD", vGiamDoc.Trim)
                    rpt_ChiLuong1KY.SetParameterValue("title", vNoidung.Trim)
                    rpt_ChiLuong1KY.SetParameterValue("GhiChu", vGhiChu.Trim)
                    If TK Then
                        rpt_ChiLuong1KY.SetParameterValue("TK", "Tài khoản")
                    Else
                        rpt_ChiLuong1KY.SetParameterValue("TK", "Ký nhận")
                    End If
                    'frm.rptView.ReportSource = rptDoc
                    'frm.rptView.Refresh()
                    'frm.rptView.Zoom(100)
                    'frm.Show()
                    rptView.ReportSource = rpt_ChiLuong1KY
                    'rptView.Refresh()
                    'rptView.Zoom(100)
                Else
                    rpt_ChiLuong1KY_NKL.SetDataSource(dt)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TC4", tCMNV)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TC5", tPC_CV)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TC6", tPC_KV)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TC8", tPC_DH)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TC9", tPC_TN)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN1", tTong100)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN2", tTongChi)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN3", tT_TNCNtmp)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN4", tT_DFCD)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN5", tT_BHXH)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN6", tT_BHYT)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN7", tT_TamUng)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN8", tT_Nghi)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN9", tTongTru)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN10", tThucLinh)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TC15", tPC_TH)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN14", tT_SotienTH)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("TN13", tT_BHTN)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                    If vkindPB = "PB" Then
                        If (vIDPB = 0 And vIdDonVi > 5) Then
                            rpt_ChiLuong1KY_NKL.SetParameterValue("Phong", "Hội sở tỉnh")
                        Else
                            rpt_ChiLuong1KY_NKL.SetParameterValue("Phong", getPhong(vIDPB))
                        End If
                    Else
                        rpt_ChiLuong1KY_NKL.SetParameterValue("Phong", getDonvi(vIDPB))
                    End If
                    rpt_ChiLuong1KY_NKL.SetParameterValue("tinh", DIABAN)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("ng", Now.Day)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("th", Now.Month)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("nm", Now.Year)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("Ky", vKy)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("thang", vThang)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("nam", vNam)
                    If DONVI = gMaDonViTW Then
                        rpt_ChiLuong1KY_NKL.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                    Else
                        rpt_ChiLuong1KY_NKL.SetParameterValue("labGD", "GIÁM ĐỐC")
                    End If
                    rpt_ChiLuong1KY_NKL.SetParameterValue("LAPBANG", vLapBieu.Trim)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("GD", vGiamDoc.Trim)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("title", vNoidung.Trim)
                    rpt_ChiLuong1KY_NKL.SetParameterValue("GhiChu", vGhiChu.Trim)
                    If TK Then
                        rpt_ChiLuong1KY_NKL.SetParameterValue("TK", "Tài khoản")
                    Else
                        rpt_ChiLuong1KY_NKL.SetParameterValue("TK", "Ký nhận")
                    End If
                    rptView.ReportSource = rpt_ChiLuong1KY_NKL
                End If

            Else
                If vKy = 1 Then
                    rpt_ChiLuongKY1.SetDataSource(dt)
                    rpt_ChiLuongKY1.SetParameterValue("TC4", tCMNV)
                    rpt_ChiLuongKY1.SetParameterValue("TC5", tPC_CV)
                    rpt_ChiLuongKY1.SetParameterValue("TC6", tPC_KV)
                    rpt_ChiLuongKY1.SetParameterValue("TC8", tPC_DH)
                    rpt_ChiLuongKY1.SetParameterValue("TC9", tPC_TN)
                    rpt_ChiLuongKY1.SetParameterValue("TN1", tTong100)
                    rpt_ChiLuongKY1.SetParameterValue("TN2", tTongChi)
                    rpt_ChiLuongKY1.SetParameterValue("TN3", tT_TNCNtmp)
                    rpt_ChiLuongKY1.SetParameterValue("TN4", tT_DFCD)
                    rpt_ChiLuongKY1.SetParameterValue("TN5", tT_BHXH)
                    rpt_ChiLuongKY1.SetParameterValue("TN6", tT_BHYT)
                    'rpt_ChiLuongKY1.SetParameterValue("TN7", tT_TamUng)
                    'rpt_ChiLuongKY1.SetParameterValue("TN8", tT_Nghi)
                    rpt_ChiLuongKY1.SetParameterValue("TN9", tTongTru)
                    rpt_ChiLuongKY1.SetParameterValue("TN10", tThucLinh)
                    'rpt_ChiLuongKY1.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                    'rpt_ChiLuongKY1.SetParameterValue("TC15", tPC_TH)
                    'rpt_ChiLuongKY1.SetParameterValue("TN14", tT_SotienTH)
                    rpt_ChiLuongKY1.SetParameterValue("TN13", tT_BHTN)
                    rpt_ChiLuongKY1.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                    If vkindPB = "PB" Then
                        If (vIDPB = 0 And vIdDonVi > 5) Then
                            rpt_ChiLuongKY1.SetParameterValue("Phong", "Hội sở tỉnh")
                        Else
                            rpt_ChiLuongKY1.SetParameterValue("Phong", getPhong(vIDPB))
                        End If
                    Else
                        rpt_ChiLuongKY1.SetParameterValue("Phong", getDonvi(vIDPB))
                    End If
                    rpt_ChiLuongKY1.SetParameterValue("tinh", DIABAN)
                    rpt_ChiLuongKY1.SetParameterValue("ng", Now.Day)
                    rpt_ChiLuongKY1.SetParameterValue("th", Now.Month)
                    rpt_ChiLuongKY1.SetParameterValue("nm", Now.Year)
                    rpt_ChiLuongKY1.SetParameterValue("Ky", vKy)
                    rpt_ChiLuongKY1.SetParameterValue("thang", vThang)
                    rpt_ChiLuongKY1.SetParameterValue("nam", vNam)
                    If DONVI = gMaDonViTW Then
                        rpt_ChiLuongKY1.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                    Else
                        rpt_ChiLuongKY1.SetParameterValue("labGD", "GIÁM ĐỐC")
                    End If
                    rpt_ChiLuongKY1.SetParameterValue("LAPBANG", vLapBieu.Trim)
                    rpt_ChiLuongKY1.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                    rpt_ChiLuongKY1.SetParameterValue("GD", vGiamDoc.Trim)
                    rpt_ChiLuongKY1.SetParameterValue("title", vNoidung.Trim)
                    rpt_ChiLuongKY1.SetParameterValue("GhiChu", vGhiChu.Trim)
                    If TK Then
                        rpt_ChiLuongKY1.SetParameterValue("TK", "Tài khoản")
                    Else
                        rpt_ChiLuongKY1.SetParameterValue("TK", "Ký nhận")
                    End If
                    rptView.ReportSource = rpt_ChiLuongKY1
                Else
                    If tT_Nghi = 0 Then
                        rpt_ChiLuongKY2.SetDataSource(dt)
                        rpt_ChiLuongKY2.SetParameterValue("TC4", tCMNV)
                        rpt_ChiLuongKY2.SetParameterValue("TC5", tPC_CV)
                        rpt_ChiLuongKY2.SetParameterValue("TC6", tPC_KV)
                        rpt_ChiLuongKY2.SetParameterValue("TC8", tPC_DH)
                        rpt_ChiLuongKY2.SetParameterValue("TC9", tPC_TN)
                        rpt_ChiLuongKY2.SetParameterValue("TN1", tTong100)
                        rpt_ChiLuongKY2.SetParameterValue("TN2", tTongChi)
                        rpt_ChiLuongKY2.SetParameterValue("TN7", tT_TamUng)
                        rpt_ChiLuongKY2.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                        rpt_ChiLuongKY2.SetParameterValue("TN9", tTongTru)
                        rpt_ChiLuongKY2.SetParameterValue("TN10", tThucLinh)
                        rpt_ChiLuongKY2.SetParameterValue("TC15", tPC_TH)
                        rpt_ChiLuongKY2.SetParameterValue("TN14", tT_SotienTH)
                        rpt_ChiLuongKY2.SetParameterValue("TN10", tThucLinh)
                        rpt_ChiLuongKY2.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                        If vkindPB = "PB" Then
                            If (vIDPB = 0 And vIdDonVi > 5) Then
                                rpt_ChiLuongKY2.SetParameterValue("Phong", "Hội sở tỉnh")
                            Else
                                rpt_ChiLuongKY2.SetParameterValue("Phong", getPhong(vIDPB))
                            End If
                        Else
                            rpt_ChiLuongKY2.SetParameterValue("Phong", getDonvi(vIDPB))
                        End If
                        rpt_ChiLuongKY2.SetParameterValue("tinh", DIABAN)
                        rpt_ChiLuongKY2.SetParameterValue("ng", Now.Day)
                        rpt_ChiLuongKY2.SetParameterValue("th", Now.Month)
                        rpt_ChiLuongKY2.SetParameterValue("nm", Now.Year)
                        rpt_ChiLuongKY2.SetParameterValue("Ky", vKy)
                        rpt_ChiLuongKY2.SetParameterValue("thang", vThang)
                        rpt_ChiLuongKY2.SetParameterValue("nam", vNam)
                        If DONVI = gMaDonViTW Then
                            rpt_ChiLuongKY2.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                        Else
                            rpt_ChiLuongKY2.SetParameterValue("labGD", "GIÁM ĐỐC")
                        End If
                        rpt_ChiLuongKY2.SetParameterValue("LAPBANG", vLapBieu.Trim)
                        rpt_ChiLuongKY2.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                        rpt_ChiLuongKY2.SetParameterValue("GD", vGiamDoc.Trim)
                        rpt_ChiLuongKY2.SetParameterValue("title", vNoidung.Trim)
                        rpt_ChiLuongKY2.SetParameterValue("GhiChu", vGhiChu.Trim)
                        If TK Then
                            rpt_ChiLuongKY2.SetParameterValue("TK", "Tài khoản")
                        Else
                            rpt_ChiLuongKY2.SetParameterValue("TK", "Ký nhận")
                        End If
                        rptView.ReportSource = rpt_ChiLuongKY2
                    Else
                        rpt_ChiLuongKY2_NKL.SetDataSource(dt)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TC4", tCMNV)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TC5", tPC_CV)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TC6", tPC_KV)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TC8", tPC_DH)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TC9", tPC_TN)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN1", tTong100)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN8", tT_Nghi)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN2", tTongChi)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN7", tT_TamUng)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN9", tTongTru)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN10", tThucLinh)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TC15", tPC_TH)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("TN14", tT_SotienTH)

                        rpt_ChiLuongKY2_NKL.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                        If vkindPB = "PB" Then
                            If (vIDPB = 0 And vIdDonVi > 5) Then
                                rpt_ChiLuongKY2_NKL.SetParameterValue("Phong", "Hội sở tỉnh")
                            Else
                                rpt_ChiLuongKY2_NKL.SetParameterValue("Phong", getPhong(vIDPB))
                            End If
                        Else
                            rpt_ChiLuongKY2_NKL.SetParameterValue("Phong", getDonvi(vIDPB))
                        End If
                        rpt_ChiLuongKY2_NKL.SetParameterValue("tinh", DIABAN)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("ng", Now.Day)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("th", Now.Month)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("nm", Now.Year)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("Ky", vKy)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("thang", vThang)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("nam", vNam)
                        If DONVI = gMaDonViTW Then
                            rpt_ChiLuongKY2_NKL.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                        Else
                            rpt_ChiLuongKY2_NKL.SetParameterValue("labGD", "GIÁM ĐỐC")
                        End If
                        rpt_ChiLuongKY2_NKL.SetParameterValue("LAPBANG", vLapBieu.Trim)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("GD", vGiamDoc.Trim)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("title", vNoidung.Trim)
                        rpt_ChiLuongKY2_NKL.SetParameterValue("GhiChu", vGhiChu.Trim)
                        If TK Then
                            rpt_ChiLuongKY2_NKL.SetParameterValue("TK", "Tài khoản")
                        Else
                            rpt_ChiLuongKY2_NKL.SetParameterValue("TK", "Ký nhận")
                        End If
                        rptView.ReportSource = rpt_ChiLuongKY2_NKL
                    End If

                End If
            End If
        Else
            If vNH_TS = 1 Then
                'Hiển thị báo cáo cho lao động ngắn hạn
                rpt_ChiLuongHDNH.SetDataSource(dt)

                rpt_ChiLuongHDNH.SetParameterValue("TN1", tTong100)
                rpt_ChiLuongHDNH.SetParameterValue("TN8", tT_Nghi)
                rpt_ChiLuongHDNH.SetParameterValue("TN2", tTongChi)
                rpt_ChiLuongHDNH.SetParameterValue("TN4", tT_DFCD)
                rpt_ChiLuongHDNH.SetParameterValue("TN5", tT_BHXH)
                rpt_ChiLuongHDNH.SetParameterValue("TN6", tT_BHYT)
                rpt_ChiLuongHDNH.SetParameterValue("TN9", tTongTru)
                rpt_ChiLuongHDNH.SetParameterValue("TN10", tThucLinh)
                rpt_ChiLuongHDNH.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                rpt_ChiLuongHDNH.SetParameterValue("TN13", tT_BHTN)
                rpt_ChiLuongHDNH.SetParameterValue("DonVi", getDonvi(IdDONVI))
                If vkindPB = "PB" Then
                    If (vIDPB = 0 And vIdDonVi > 5) Then
                        rpt_ChiLuongHDNH.SetParameterValue("Phong", "Hội sở tỉnh")
                    Else
                        rpt_ChiLuongHDNH.SetParameterValue("Phong", getPhong(vIDPB))
                    End If
                Else
                    If vIDPB = IdDONVI Then
                        rpt_ChiLuongHDNH.SetParameterValue("Phong", "Hội sở tỉnh")
                    Else
                        rpt_ChiLuongHDNH.SetParameterValue("Phong", getDonvi(vIDPB))
                    End If
                End If
                'rpt_ChiLuongHDNH.SetParameterValue("Phong", getDonvi(vIdDonVi))
                rpt_ChiLuongHDNH.SetParameterValue("tinh", DIABAN)
                rpt_ChiLuongHDNH.SetParameterValue("ng", Now.Day)
                rpt_ChiLuongHDNH.SetParameterValue("th", Now.Month)
                rpt_ChiLuongHDNH.SetParameterValue("nm", Now.Year)
                rpt_ChiLuongHDNH.SetParameterValue("thang", vThang)
                rpt_ChiLuongHDNH.SetParameterValue("nam", vNam)
                If DONVI = gMaDonViTW Then
                    rpt_ChiLuongHDNH.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                Else
                    rpt_ChiLuongHDNH.SetParameterValue("labGD", "GIÁM ĐỐC")
                End If
                rpt_ChiLuongHDNH.SetParameterValue("LAPBANG", vLapBieu.Trim)
                rpt_ChiLuongHDNH.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                rpt_ChiLuongHDNH.SetParameterValue("GD", vGiamDoc.Trim)
                rpt_ChiLuongHDNH.SetParameterValue("title", vNoidung.Trim)
                rpt_ChiLuongHDNH.SetParameterValue("GhiChu", vGhiChu.Trim)
                If TK Then
                    rpt_ChiLuongHDNH.SetParameterValue("TK", "Tài khoản")
                Else
                    rpt_ChiLuongHDNH.SetParameterValue("TK", "Ký nhận")
                End If
                rptView.ReportSource = rpt_ChiLuongHDNH
            Else
                'Hiển thị báo cáo cho lao động tập sự
                rpt_ChiLuongTS.SetDataSource(dt)
                rpt_ChiLuongTS.SetParameterValue("TC4", tCMNV)
                rpt_ChiLuongTS.SetParameterValue("TN3", tT_TNCNtmp)  'tong100 theo he so
                rpt_ChiLuongTS.SetParameterValue("TN1", tTong100)    'theo ty le huong
                rpt_ChiLuongTS.SetParameterValue("TN8", tT_Nghi)
                rpt_ChiLuongTS.SetParameterValue("TN2", tTongChi)
                rpt_ChiLuongTS.SetParameterValue("TN9", tTongTru)
                rpt_ChiLuongTS.SetParameterValue("TN10", tThucLinh)
                rpt_ChiLuongTS.SetParameterValue("TN11", tT_CacKhoanThuKhac)
                rpt_ChiLuongTS.SetParameterValue("DonVi", getDonvi(IdDONVI))
                If vkindPB = "PB" Then
                    If (vIDPB = 0 And vIdDonVi > 5) Then
                        rpt_ChiLuongTS.SetParameterValue("Phong", "Hội sở tỉnh")
                    Else
                        rpt_ChiLuongTS.SetParameterValue("Phong", getPhong(vIDPB))
                    End If
                Else
                    If vIDPB = IdDONVI Then
                        rpt_ChiLuongTS.SetParameterValue("Phong", "Hội sở tỉnh")
                    Else
                        rpt_ChiLuongTS.SetParameterValue("Phong", getDonvi(vIDPB))
                    End If
                End If
                'rpt_ChiLuongTS.SetParameterValue("Phong", getDonvi(vIdDonVi))
                rpt_ChiLuongTS.SetParameterValue("tinh", DIABAN)
                rpt_ChiLuongTS.SetParameterValue("ng", Now.Day)
                rpt_ChiLuongTS.SetParameterValue("th", Now.Month)
                rpt_ChiLuongTS.SetParameterValue("nm", Now.Year)
                rpt_ChiLuongTS.SetParameterValue("thang", vThang)
                rpt_ChiLuongTS.SetParameterValue("nam", vNam)
                If DONVI = gMaDonViTW Then
                    rpt_ChiLuongTS.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                Else
                    rpt_ChiLuongTS.SetParameterValue("labGD", "GIÁM ĐỐC")
                End If
                rpt_ChiLuongTS.SetParameterValue("LAPBANG", vLapBieu.Trim)
                rpt_ChiLuongTS.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                rpt_ChiLuongTS.SetParameterValue("GD", vGiamDoc.Trim)
                rpt_ChiLuongTS.SetParameterValue("title", vNoidung.Trim)
                rpt_ChiLuongTS.SetParameterValue("GhiChu", vGhiChu.Trim)
                If TK Then
                    rpt_ChiLuongTS.SetParameterValue("TK", "Tài khoản")
                Else
                    rpt_ChiLuongTS.SetParameterValue("TK", "Ký nhận")
                End If
                rptView.ReportSource = rpt_ChiLuongTS
            End If
        End If


    End Sub

    'Private Sub showReportChiLuong_HDNH(ByVal IdChiLuong As String, _
    '                             ByVal vIdDonVi As Integer, ByVal vThang As Integer, ByVal vNam As Integer, _
    '                             ByVal vLapBieu As String, ByVal vKiemSoat As String, ByVal vGiamDoc As String, ByVal vGhiChu As String)

    '    Dim tTong100, tTongChi, tT_DFCD, tT_BHXH, tT_BHYT, tT_BHTN, tT_Nghi, tT_TruyThu, tTongTru, tTruyLinh, tThucLinh As Long
    '    Dim dt As DataTable = New DataTable
    '    'dt = getReportBangChiLuong_HDNH(IdChiLuong, tTong100, tTongChi, tT_DFCD, tT_BHXH, tT_BHYT, tT_BHTN, tT_Nghi, tT_TruyThu, tTongTru, tTruyLinh, tThucLinh)
    '    rpt_ChiLuongHDNH.SetDataSource(dt)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN1", tTong100)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN2", tTongChi)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN3", tT_BHXH)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN4", tT_BHYT)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN10", tT_BHTN)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN5", tT_DFCD)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN11", tT_Nghi)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN8", tT_TruyThu)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN6", tTongTru)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN9", tTruyLinh)
    '    rpt_ChiLuongHDNH.SetParameterValue("TN7", tThucLinh)
    '    rpt_ChiLuongHDNH.SetParameterValue("DonVi", getDonvi(vIdDonVi))
    '    rpt_ChiLuongHDNH.SetParameterValue("tinh", DIABAN)
    '    rpt_ChiLuongHDNH.SetParameterValue("ng", Now.Day)
    '    rpt_ChiLuongHDNH.SetParameterValue("th", Now.Month)
    '    rpt_ChiLuongHDNH.SetParameterValue("nm", Now.Year)
    '    rpt_ChiLuongHDNH.SetParameterValue("thang", vThang)
    '    rpt_ChiLuongHDNH.SetParameterValue("nam", vNam)
    '    If DONVI = gMaDonViTW Then
    '        rpt_ChiLuongHDNH.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
    '    Else
    '        rpt_ChiLuongHDNH.SetParameterValue("labGD", "GIÁM ĐỐC")
    '    End If
    '    rpt_ChiLuongHDNH.SetParameterValue("LAPBANG", vLapBieu.Trim)
    '    rpt_ChiLuongHDNH.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
    '    rpt_ChiLuongHDNH.SetParameterValue("GD", vGiamDoc.Trim)
    '    rpt_ChiLuongHDNH.SetParameterValue("GhiChu", vGhiChu.Trim)
    '    rptView.ReportSource = rpt_ChiLuongHDNH
    'End Sub

    Private Sub showReportChiLuong_TS(ByVal IdChiLuong As String, _
                                   ByVal vIdDonVi As Integer, ByVal vThang As Integer, ByVal vNam As Integer, _
                                   ByVal vLapBieu As String, ByVal vKiemSoat As String, ByVal vGiamDoc As String, ByVal vGhiChu As String)

        Dim tCMNV As String = "0"
        Dim tTong, tTong100, tTongChi, tT_TruyThu, tTruyLinh, tThucLinh As Long
        Dim dt As DataTable = New DataTable
        dt = getReportBangChiLuong_TS(IdChiLuong, tCMNV, tTong, tTong100, tTongChi, tT_TruyThu, tTruyLinh, tThucLinh)
        rpt_ChiLuongTS.SetDataSource(dt)
        rpt_ChiLuongTS.SetParameterValue("TC3", tCMNV)
        rpt_ChiLuongTS.SetParameterValue("TN1", tTong)
        rpt_ChiLuongTS.SetParameterValue("TN2", tTong100)
        rpt_ChiLuongTS.SetParameterValue("TN3", tTongChi)
        rpt_ChiLuongTS.SetParameterValue("TN5", tT_TruyThu)
        rpt_ChiLuongTS.SetParameterValue("TN6", tTruyLinh)
        rpt_ChiLuongTS.SetParameterValue("TN4", tThucLinh)
        rpt_ChiLuongTS.SetParameterValue("DonVi", getDonvi(vIdDonVi))
        rpt_ChiLuongTS.SetParameterValue("tinh", DIABAN)
        rpt_ChiLuongTS.SetParameterValue("ng", Now.Day)
        rpt_ChiLuongTS.SetParameterValue("th", Now.Month)
        rpt_ChiLuongTS.SetParameterValue("nm", Now.Year)
        rpt_ChiLuongTS.SetParameterValue("thang", vThang)
        rpt_ChiLuongTS.SetParameterValue("nam", vNam)
        If DONVI = gMaDonViTW Then
            rpt_ChiLuongTS.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
        Else
            rpt_ChiLuongTS.SetParameterValue("labGD", "GIÁM ĐỐC")
        End If
        rpt_ChiLuongTS.SetParameterValue("LAPBANG", vLapBieu.Trim)
        rpt_ChiLuongTS.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
        rpt_ChiLuongTS.SetParameterValue("GD", vGiamDoc.Trim)
        rpt_ChiLuongTS.SetParameterValue("GhiChu", vGhiChu.Trim)
        rptView.ReportSource = rpt_ChiLuongTS
    End Sub

    Private Sub showReportChiThemGio(ByVal IdChiLuong As String, ByVal TK As Boolean, ByVal vIdDonVi As Integer, ByVal vThang As Integer, ByVal vNam As Integer, _
                                     ByVal vLapBieu As String, ByVal vKiemSoat As String, ByVal vGiamDoc As String, ByVal vGhiChu As String)

        Dim tCMNV As String = "0"
        Dim tPC As String = "0"
        Dim tHeSo As String = "0"
        Dim tGio50 As String = "0"
        Dim tGio100 As String = "0"
        Dim tGio150 As String = "0"
        Dim tGio200 As String = "0"
        Dim tGio300 As String = "0"
        Dim tGioNghiBu As String = "0"
        Dim tGioQD As String = "0"
        Dim vIdPhong As String = ""
        Dim vkindPB As String = ""
        Dim vIDPB As Integer = 0
        Dim tTraThemLamDem, tThucLinh As Long
        Dim dt As DataTable = New DataTable
        dt = printBangKeChiThemGio(IdChiLuong, TK, vIdPhong, tCMNV, tPC, tHeSo, tGio50, tGio100, tGio150, tGio200, tGio300, tGioQD, tGioNghiBu, tTraThemLamDem, tThucLinh)
        If vIdPhong <> "" Then
            vkindPB = vIdPhong.Substring(0, 2)
            vIDPB = vIdPhong.Substring(3)
        Else
            vkindPB = "PB"
            vIDPB = 0
        End If
        rpt_ChiThemGio.SetDataSource(dt)
        rpt_ChiThemGio.SetParameterValue("TC4", tCMNV)
        rpt_ChiThemGio.SetParameterValue("TC5", tPC)
        rpt_ChiThemGio.SetParameterValue("TC15", tHeSo)
        rpt_ChiThemGio.SetParameterValue("TC6", tGio50)
        rpt_ChiThemGio.SetParameterValue("TC7", tGio100)
        rpt_ChiThemGio.SetParameterValue("TC8", tGio150)
        rpt_ChiThemGio.SetParameterValue("TC9", tGio200)
        rpt_ChiThemGio.SetParameterValue("TC10", tGio300)
        rpt_ChiThemGio.SetParameterValue("TC11", tGioQD)
        rpt_ChiThemGio.SetParameterValue("TC12", tGioNghiBu)
        rpt_ChiThemGio.SetParameterValue("TC13", tTraThemLamDem)
        rpt_ChiThemGio.SetParameterValue("TC14", tThucLinh)
        rpt_ChiThemGio.SetParameterValue("DonVi", getDonvi(vIdDonVi))
        If vkindPB = "PB" Then
            If (vIDPB = 0 And vIdDonVi > 5) Then
                rpt_ChiThemGio.SetParameterValue("Phong", "Hội sở tỉnh")
            Else
                'rpt_ChiThemGio.SetParameterValue("Phong", getPhong(vIDPB))
                rpt_ChiThemGio.SetParameterValue("Phong", "")
            End If
        Else
            rpt_ChiThemGio.SetParameterValue("Phong", getDonvi(vIDPB))
        End If
        rpt_ChiThemGio.SetParameterValue("tinh", DIABAN)
        rpt_ChiThemGio.SetParameterValue("ng", Now.Day)
        rpt_ChiThemGio.SetParameterValue("th", Now.Month)
        rpt_ChiThemGio.SetParameterValue("nm", Now.Year)
        If TK Then
            rpt_ChiThemGio.SetParameterValue("loai", "Tài khoản")
        Else
            rpt_ChiThemGio.SetParameterValue("loai", "Ký nhận")
        End If
        rpt_ChiThemGio.SetParameterValue("thang", vThang)
        rpt_ChiThemGio.SetParameterValue("nam", vNam)
        If DONVI = gMaDonViTW Then
            rpt_ChiThemGio.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
        Else
            rpt_ChiThemGio.SetParameterValue("labGD", "GIÁM ĐỐC")
        End If
        rpt_ChiThemGio.SetParameterValue("LAPBIEU", vLapBieu.Trim)
        rpt_ChiThemGio.SetParameterValue("HCTC", vKiemSoat.Trim)
        rpt_ChiThemGio.SetParameterValue("GD", vGiamDoc.Trim)
        rpt_ChiThemGio.SetParameterValue("GhiChu", vGhiChu.Trim)
        rptView.ReportSource = rpt_ChiThemGio

    End Sub

    Private Sub showReportChiBoSung(ByVal IdChiLuong As String, ByVal vLoaiCB As Int16, _
                                 ByVal vIdDonVi As Integer, ByVal vThang As Integer, ByVal vNam As Integer, _
                                 ByVal vLapBieu As String, ByVal vKiemSoat As String, ByVal vGiamDoc As String, ByVal vGhiChu As String)

        Dim tCMNV As String = "0"
        Dim tPC_CV As String = "0"
        Dim tPC_KV As String = "0"
        Dim tPC_KN As String = "0"
        Dim tPC_DH As String = "0"
        Dim tPC_TN As String = "0"
        Dim tPC_TNVK As String = "0"
        Dim tPC_TNN As String = "0"
        Dim tPC_Khac As String = "0"
        Dim vIdPhong As String = ""
        Dim vkindPB As String = ""
        Dim vNoidungChi As String = ""
        Dim vIDPB As Integer = 0
        Dim tTong100, tTongChi As Long
        Dim dt As DataTable = New DataTable
        If vLoaiCB = 1 Or LoaiCB = 2 Then
            dt = getReportBangChiBoSung_NHTS(IdChiLuong, vIdPhong, tCMNV, tTong100, tTongChi, vNoidungChi)
            If vIdPhong <> "" Then
                vkindPB = vIdPhong.Substring(0, 2)
                vIDPB = vIdPhong.Substring(3)
            Else
                vkindPB = "PB"
                vIDPB = 0
            End If
            If vLoaiCB = 1 Then
                rpt_ChiBoSung_NH.SetDataSource(dt)
                rpt_ChiBoSung_NH.SetParameterValue("GhiChu1", vNoidungChi)
                rpt_ChiBoSung_NH.SetParameterValue("TN1", tTong100)
                rpt_ChiBoSung_NH.SetParameterValue("TN2", tTongChi)
                rpt_ChiBoSung_NH.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                If vkindPB = "PB" Then
                    If (vIDPB = 0 And vIdDonVi > 5) Then
                        rpt_ChiBoSung_NH.SetParameterValue("Phong", "Hội sở tỉnh")
                    Else
                        rpt_ChiBoSung_NH.SetParameterValue("Phong", getPhong(vIDPB))
                    End If
                Else
                    rpt_ChiBoSung_NH.SetParameterValue("Phong", getDonvi(vIDPB))
                End If
                rpt_ChiBoSung_NH.SetParameterValue("tinh", DIABAN)
                rpt_ChiBoSung_NH.SetParameterValue("ng", Now.Day)
                rpt_ChiBoSung_NH.SetParameterValue("th", Now.Month)
                rpt_ChiBoSung_NH.SetParameterValue("nm", Now.Year)
                rpt_ChiBoSung_NH.SetParameterValue("thang", vThang)
                rpt_ChiBoSung_NH.SetParameterValue("nam", vNam)
                If DONVI = gMaDonViTW Then
                    rpt_ChiBoSung_NH.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                Else
                    rpt_ChiBoSung_NH.SetParameterValue("labGD", "GIÁM ĐỐC")
                End If
                rpt_ChiBoSung_NH.SetParameterValue("LAPBANG", vLapBieu.Trim)
                rpt_ChiBoSung_NH.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                rpt_ChiBoSung_NH.SetParameterValue("GD", vGiamDoc.Trim)
                rpt_ChiBoSung_NH.SetParameterValue("GhiChu", vGhiChu.Trim)
                rptView.ReportSource = rpt_ChiBoSung_NH
            Else
                rpt_ChiBoSung_TS.SetDataSource(dt)
                rpt_ChiBoSung_TS.SetParameterValue("GhiChu1", vNoidungChi)
                rpt_ChiBoSung_TS.SetParameterValue("TC4", tCMNV)
                rpt_ChiBoSung_TS.SetParameterValue("TN1", tTong100)
                rpt_ChiBoSung_TS.SetParameterValue("TN2", tTongChi)
                rpt_ChiBoSung_TS.SetParameterValue("DonVi", getDonvi(vIdDonVi))
                If vkindPB = "PB" Then
                    If (vIDPB = 0 And vIdDonVi > 5) Then
                        rpt_ChiBoSung_TS.SetParameterValue("Phong", "Hội sở tỉnh")
                    Else
                        rpt_ChiBoSung_TS.SetParameterValue("Phong", getPhong(vIDPB))
                    End If
                Else
                    rpt_ChiBoSung_TS.SetParameterValue("Phong", getDonvi(vIDPB))
                End If
                rpt_ChiBoSung_TS.SetParameterValue("tinh", DIABAN)
                rpt_ChiBoSung_TS.SetParameterValue("ng", Now.Day)
                rpt_ChiBoSung_TS.SetParameterValue("th", Now.Month)
                rpt_ChiBoSung_TS.SetParameterValue("nm", Now.Year)
                rpt_ChiBoSung_TS.SetParameterValue("thang", vThang)
                rpt_ChiBoSung_TS.SetParameterValue("nam", vNam)
                If DONVI = gMaDonViTW Then
                    rpt_ChiBoSung_TS.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                Else
                    rpt_ChiBoSung_TS.SetParameterValue("labGD", "GIÁM ĐỐC")
                End If
                rpt_ChiBoSung_TS.SetParameterValue("LAPBANG", vLapBieu.Trim)
                rpt_ChiBoSung_TS.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
                rpt_ChiBoSung_TS.SetParameterValue("GD", vGiamDoc.Trim)
                rpt_ChiBoSung_TS.SetParameterValue("GhiChu", vGhiChu.Trim)
                rptView.ReportSource = rpt_ChiBoSung_TS
            End If
        Else
            dt = getReportBangChiBoSung(IdChiLuong, vLoaiCB, vIdPhong, tCMNV, tPC_CV, tPC_KV, tPC_KN, tPC_DH, tPC_TN, tPC_TNVK, tPC_TNN, tPC_Khac, tTong100, tTongChi, vNoidungChi)
            If vIdPhong <> "" Then
                vkindPB = vIdPhong.Substring(0, 2)
                vIDPB = vIdPhong.Substring(3)
            Else
                vkindPB = "PB"
                vIDPB = 0
            End If
            rpt_ChiBoSung.SetDataSource(dt)
            rpt_ChiBoSung.SetParameterValue("GhiChu1", vNoidungChi)
            rpt_ChiBoSung.SetParameterValue("TC4", tCMNV)
            rpt_ChiBoSung.SetParameterValue("TC5", tPC_CV)
            rpt_ChiBoSung.SetParameterValue("TC6", tPC_KV)
            rpt_ChiBoSung.SetParameterValue("TC7", tPC_KN)
            rpt_ChiBoSung.SetParameterValue("TC8", tPC_DH)
            rpt_ChiBoSung.SetParameterValue("TC9", tPC_TN)
            rpt_ChiBoSung.SetParameterValue("TC10", tPC_TNVK)
            rpt_ChiBoSung.SetParameterValue("TC11", tPC_TNN)
            rpt_ChiBoSung.SetParameterValue("TC12", tPC_Khac)
            rpt_ChiBoSung.SetParameterValue("TN1", tTong100)
            rpt_ChiBoSung.SetParameterValue("TN2", tTongChi)
            rpt_ChiBoSung.SetParameterValue("DonVi", getDonvi(vIdDonVi))
            If vkindPB = "PB" Then
                If (vIDPB = 0 And vIdDonVi > 5) Then
                    rpt_ChiBoSung.SetParameterValue("Phong", "Hội sở tỉnh")
                Else
                    rpt_ChiBoSung.SetParameterValue("Phong", getPhong(vIDPB))
                End If
            Else
                rpt_ChiBoSung.SetParameterValue("Phong", getDonvi(vIDPB))
            End If
            rpt_ChiBoSung.SetParameterValue("tinh", DIABAN)
            rpt_ChiBoSung.SetParameterValue("ng", Now.Day)
            rpt_ChiBoSung.SetParameterValue("th", Now.Month)
            rpt_ChiBoSung.SetParameterValue("nm", Now.Year)
            rpt_ChiBoSung.SetParameterValue("thang", vThang)
            rpt_ChiBoSung.SetParameterValue("nam", vNam)
            If DONVI = gMaDonViTW Then
                rpt_ChiBoSung.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
            Else
                rpt_ChiBoSung.SetParameterValue("labGD", "GIÁM ĐỐC")
            End If
            rpt_ChiBoSung.SetParameterValue("LAPBANG", vLapBieu.Trim)
            rpt_ChiBoSung.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
            rpt_ChiBoSung.SetParameterValue("GD", vGiamDoc.Trim)
            rpt_ChiBoSung.SetParameterValue("GhiChu", vGhiChu.Trim)
            rptView.ReportSource = rpt_ChiBoSung
        End If
    End Sub

End Class