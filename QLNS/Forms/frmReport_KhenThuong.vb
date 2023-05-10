Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms

Public Class frmReport_KhenThuong
    Private _IdDonVi As Integer
    Private _Nam As Integer
    Private _KT_ChuyenMon As Int16
    Private _KhenThuong As Int16
    Private _DinhKy As Int16
    Private _CaNhan As Int16

    Public Property IdDonVi() As Integer
        Get
            Return _IdDonVi
        End Get
        Set(ByVal value As Integer)
            _IdDonVi = value
        End Set
    End Property

    Public Property KT_ChuyenMon() As Int16
        Get
            Return _KT_ChuyenMon
        End Get
        Set(ByVal value As Int16)
            _KT_ChuyenMon = value
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

    Public Property KhenThuong() As Int16
        Get
            Return _KhenThuong
        End Get
        Set(ByVal value As Int16)
            _KhenThuong = value
        End Set
    End Property

    Public Property DinhKy() As Int16
        Get
            Return _DinhKy
        End Get
        Set(ByVal value As Int16)
            _DinhKy = value
        End Set
    End Property

    Public Property CaNhan() As Int16
        Get
            Return _CaNhan
        End Get
        Set(ByVal value As Int16)
            _CaNhan = value
        End Set
    End Property

    Private Sub frmReport_KhenThuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmReport_KhenThuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        showReportKhenThuong(IdDonVi, KhenThuong, DinhKy, Nam, KT_ChuyenMon, CaNhan)

    End Sub

    Private Sub showReportKhenThuong(ByVal vIdDonVi As Integer, ByVal vKhenThuong As Int16, ByVal vDinhKy As Int16, ByVal vNam As Integer, ByVal vKT_ChuyenMon As Int16, ByVal vCaNhan As Int16)

        Dim dt As DataTable = New DataTable
        Dim tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27 As Integer
        dt = getDSKhenThuong(vIdDonVi, vKhenThuong, vDinhKy, vNam, vKT_ChuyenMon, vCaNhan, tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27)
        If vCaNhan = 1 Then
            If vIdDonVi = 0 Then
                rpt_KhenThuongCN_TQ.SetDataSource(dt)
                rpt_KhenThuongCN_TQ.SetParameterValue("ChiNhanh", "")
                rpt_KhenThuongCN_TQ.SetParameterValue("TN6", tC6)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN7", tC7)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN8", tC8)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN9", tC9)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN10", tC10)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN11", tC11)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN12", tC12)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN13", tC13)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN14", tC14)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN15", tC15)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN16", tC16)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN17", tC17)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN18", tC18)
                rpt_KhenThuongCN_TQ.SetParameterValue("TN19", tC19)
                If vKhenThuong = 1 Then
                    rpt_KhenThuongCN_TQ.SetParameterValue("NameKT", "DANH SÁCH CÁ NHÂN KHEN THƯỞNG")
                Else
                    rpt_KhenThuongCN_TQ.SetParameterValue("NameKT", "DANH SÁCH CÁ NHÂN ĐỀ NGHỊ KHEN THƯỞNG")
                End If
                rpt_KhenThuongCN_TQ.SetParameterValue("tinh", DIABAN)
                rpt_KhenThuongCN_TQ.SetParameterValue("ng", Now.Day)
                rpt_KhenThuongCN_TQ.SetParameterValue("th", Now.Month)
                rpt_KhenThuongCN_TQ.SetParameterValue("nm", Now.Year)
                rpt_KhenThuongCN_TQ.SetParameterValue("nam", vNam)
                rpt_KhenThuongCN_TQ.SetParameterValue("NguoiLapBieu", "Nguyen Thi Nhan")
                rptView.ReportSource = rpt_KhenThuongCN_TQ
            Else
                rpt_KhenThuongCN.SetDataSource(dt)
                rpt_KhenThuongCN.SetParameterValue("ChiNhanh", getDonvi(vIdDonVi))
                rpt_KhenThuongCN.SetParameterValue("TN6", tC6)
                rpt_KhenThuongCN.SetParameterValue("TN7", tC7)
                rpt_KhenThuongCN.SetParameterValue("TN8", tC8)
                rpt_KhenThuongCN.SetParameterValue("TN9", tC9)
                rpt_KhenThuongCN.SetParameterValue("TN10", tC10)
                rpt_KhenThuongCN.SetParameterValue("TN11", tC11)
                rpt_KhenThuongCN.SetParameterValue("TN12", tC12)
                rpt_KhenThuongCN.SetParameterValue("TN13", tC13)
                rpt_KhenThuongCN.SetParameterValue("TN14", tC14)
                rpt_KhenThuongCN.SetParameterValue("TN15", tC15)
                rpt_KhenThuongCN.SetParameterValue("TN16", tC16)
                rpt_KhenThuongCN.SetParameterValue("TN17", tC17)
                rpt_KhenThuongCN.SetParameterValue("TN18", tC18)
                rpt_KhenThuongCN.SetParameterValue("TN19", tC19)
                If vKhenThuong = 1 Then
                    rpt_KhenThuongCN.SetParameterValue("NameKT", "DANH SÁCH CÁ NHÂN KHEN THƯỞNG")
                Else
                    rpt_KhenThuongCN.SetParameterValue("NameKT", "DANH SÁCH CÁ NHÂN ĐỀ NGHỊ KHEN THƯỞNG")
                End If
                rpt_KhenThuongCN.SetParameterValue("tinh", DIABAN)
                rpt_KhenThuongCN.SetParameterValue("ng", Now.Day)
                rpt_KhenThuongCN.SetParameterValue("th", Now.Month)
                rpt_KhenThuongCN.SetParameterValue("nm", Now.Year)
                rpt_KhenThuongCN.SetParameterValue("nam", vNam)
                'If DONVI = gMaDonViTW Then
                '    rpt_KhenThuongCN.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                'Else
                '    rpt_KhenThuongCN.SetParameterValue("labGD", "GIÁM ĐỐC")
                'End If
                rpt_KhenThuongCN.SetParameterValue("NguoiLapBieu", "Nguyen Thi Nhan")
                'rpt_KhenThuongCN.SetParameterValue("HCTC", vKiemSoat.Trim)
                'rpt_KhenThuongCN.SetParameterValue("GD", vGiamDoc.Trim)
                'rpt_KhenThuongCN.SetParameterValue("GhiChu", vGhiChu.Trim)
                rptView.ReportSource = rpt_KhenThuongCN
            End If
        Else
            If vIdDonVi = 0 Then
                rpt_KhenThuongTT_TQ.SetDataSource(dt)
                rpt_KhenThuongTT_TQ.SetParameterValue("ChiNhanh", "")
                rpt_KhenThuongTT_TQ.SetParameterValue("TN6", tC6)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN7", tC7)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN8", tC8)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN9", tC9)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN10", tC10)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN11", tC11)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN12", tC12)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN13", tC13)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN14", tC14)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN15", tC15)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN16", tC16)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN17", tC17)
                rpt_KhenThuongTT_TQ.SetParameterValue("TN18", tC18)
                If vKhenThuong = 1 Then
                    rpt_KhenThuongTT_TQ.SetParameterValue("NameKT", "DANH SÁCH TẬP THỂ KHEN THƯỞNG")
                Else
                    rpt_KhenThuongTT_TQ.SetParameterValue("NameKT", "DANH SÁCH TẬP THỂ ĐỀ NGHỊ KHEN THƯỞNG")
                End If
                rpt_KhenThuongTT_TQ.SetParameterValue("tinh", DIABAN)
                rpt_KhenThuongTT_TQ.SetParameterValue("ng", Now.Day)
                rpt_KhenThuongTT_TQ.SetParameterValue("th", Now.Month)
                rpt_KhenThuongTT_TQ.SetParameterValue("nm", Now.Year)
                rpt_KhenThuongTT_TQ.SetParameterValue("nam", vNam)
                rpt_KhenThuongTT_TQ.SetParameterValue("NguoiLapBieu", "Nguyen Thi Nhan")
                rptView.ReportSource = rpt_KhenThuongTT_TQ
            Else
                rpt_KhenThuongTT.SetDataSource(dt)
                rpt_KhenThuongTT.SetParameterValue("ChiNhanh", getDonvi(vIdDonVi))
                rpt_KhenThuongTT.SetParameterValue("TN6", tC6)
                rpt_KhenThuongTT.SetParameterValue("TN7", tC7)
                rpt_KhenThuongTT.SetParameterValue("TN8", tC8)
                rpt_KhenThuongTT.SetParameterValue("TN9", tC9)
                rpt_KhenThuongTT.SetParameterValue("TN10", tC10)
                rpt_KhenThuongTT.SetParameterValue("TN11", tC11)
                rpt_KhenThuongTT.SetParameterValue("TN12", tC12)
                rpt_KhenThuongTT.SetParameterValue("TN13", tC13)
                rpt_KhenThuongTT.SetParameterValue("TN14", tC14)
                rpt_KhenThuongTT.SetParameterValue("TN15", tC15)
                rpt_KhenThuongTT.SetParameterValue("TN16", tC16)
                rpt_KhenThuongTT.SetParameterValue("TN17", tC17)
                rpt_KhenThuongTT.SetParameterValue("TN18", tC18)
                If vKhenThuong = 1 Then
                    rpt_KhenThuongTT.SetParameterValue("NameKT", "DANH SÁCH TẬP THỂ KHEN THƯỞNG")
                Else
                    rpt_KhenThuongTT.SetParameterValue("NameKT", "DANH SÁCH TẬP THỂ ĐỀ NGHỊ KHEN THƯỞNG")
                End If
                rpt_KhenThuongTT.SetParameterValue("tinh", DIABAN)
                rpt_KhenThuongTT.SetParameterValue("ng", Now.Day)
                rpt_KhenThuongTT.SetParameterValue("th", Now.Month)
                rpt_KhenThuongTT.SetParameterValue("nm", Now.Year)
                rpt_KhenThuongTT.SetParameterValue("nam", vNam)
                rpt_KhenThuongTT.SetParameterValue("NguoiLapBieu", "Nguyen Thi Nhan")
                rptView.ReportSource = rpt_KhenThuongTT
            End If
        End If
    End Sub

    'Private Sub showReportKhenThuong(ByVal vIdDonVi As Integer, ByVal vKhenThuong As Int16, ByVal vNam As Integer, ByVal vKT_ChuyenMon As Int16, ByVal vCaNhan As Int16)

    '    Dim dt As DataTable = New DataTable
    '    Dim tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27 As Integer
    '    dt = getDSKhenThuong(vIdDonVi, vKhenThuong, vNam, vKT_ChuyenMon, vCaNhan, tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27)
    '    If vCaNhan = 1 Then
    '        rpt_KhenThuongCN.SetDataSource(dt)
    '        rpt_KhenThuongCN.SetParameterValue("TN6", tC6)
    '        rpt_KhenThuongCN.SetParameterValue("TN7", tC7)
    '        rpt_KhenThuongCN.SetParameterValue("TN8", tC8)
    '        rpt_KhenThuongCN.SetParameterValue("TN9", tC9)
    '        rpt_KhenThuongCN.SetParameterValue("TN10", tC10)
    '        rpt_KhenThuongCN.SetParameterValue("TN11", tC11)
    '        rpt_KhenThuongCN.SetParameterValue("TN12", tC12)
    '        rpt_KhenThuongCN.SetParameterValue("TN13", tC13)
    '        rpt_KhenThuongCN.SetParameterValue("TN14", tC14)
    '        rpt_KhenThuongCN.SetParameterValue("TN15", tC15)
    '        rpt_KhenThuongCN.SetParameterValue("TN16", tC16)
    '        rpt_KhenThuongCN.SetParameterValue("TN17", tC17)
    '        rpt_KhenThuongCN.SetParameterValue("TN18", tC18)
    '        rpt_KhenThuongCN.SetParameterValue("TN19", tC19)
    '        If vIdDonVi = 0 Then
    '            rpt_KhenThuongCN.SetParameterValue("ChiNhanh", "")
    '        Else
    '            rpt_KhenThuongCN.SetParameterValue("ChiNhanh", getDonvi(vIdDonVi))
    '        End If
    '        If vKhenThuong = 1 Then
    '            rpt_KhenThuongCN.SetParameterValue("NameKT", "DANH SÁCH CÁ NHÂN KHEN THƯỞNG")
    '        Else
    '            rpt_KhenThuongCN.SetParameterValue("NameKT", "DANH SÁCH CÁ NHÂN ĐỀ NGHỊ KHEN THƯỞNG")
    '        End If
    '        rpt_KhenThuongCN.SetParameterValue("tinh", DIABAN)
    '        rpt_KhenThuongCN.SetParameterValue("ng", Now.Day)
    '        rpt_KhenThuongCN.SetParameterValue("th", Now.Month)
    '        rpt_KhenThuongCN.SetParameterValue("nm", Now.Year)
    '        rpt_KhenThuongCN.SetParameterValue("nam", vNam)
    '        'If DONVI = gMaDonViTW Then
    '        '    rpt_KhenThuongCN.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
    '        'Else
    '        '    rpt_KhenThuongCN.SetParameterValue("labGD", "GIÁM ĐỐC")
    '        'End If
    '        rpt_KhenThuongCN.SetParameterValue("NguoiLapBieu", "Nguyen Thi Nhan")
    '        'rpt_KhenThuongCN.SetParameterValue("HCTC", vKiemSoat.Trim)
    '        'rpt_KhenThuongCN.SetParameterValue("GD", vGiamDoc.Trim)
    '        'rpt_KhenThuongCN.SetParameterValue("GhiChu", vGhiChu.Trim)
    '        rptView.ReportSource = rpt_KhenThuongCN
    '    Else
    '        rpt_KhenThuongTT.SetDataSource(dt)
    '        rpt_KhenThuongTT.SetParameterValue("TN6", tC6)
    '        rpt_KhenThuongTT.SetParameterValue("TN7", tC7)
    '        rpt_KhenThuongTT.SetParameterValue("TN8", tC8)
    '        rpt_KhenThuongTT.SetParameterValue("TN9", tC9)
    '        rpt_KhenThuongTT.SetParameterValue("TN10", tC10)
    '        rpt_KhenThuongTT.SetParameterValue("TN11", tC11)
    '        rpt_KhenThuongTT.SetParameterValue("TN12", tC12)
    '        rpt_KhenThuongTT.SetParameterValue("TN13", tC13)
    '        rpt_KhenThuongTT.SetParameterValue("TN14", tC14)
    '        rpt_KhenThuongTT.SetParameterValue("TN15", tC15)
    '        rpt_KhenThuongTT.SetParameterValue("TN16", tC16)
    '        rpt_KhenThuongTT.SetParameterValue("TN17", tC17)
    '        rpt_KhenThuongTT.SetParameterValue("TN18", tC18)
    '        If vIdDonVi = 0 Then
    '            rpt_KhenThuongTT.SetParameterValue("ChiNhanh", "")
    '        Else
    '            rpt_KhenThuongTT.SetParameterValue("ChiNhanh", getDonvi(vIdDonVi))
    '        End If
    '        If vKhenThuong = 1 Then
    '            rpt_KhenThuongTT.SetParameterValue("NameKT", "DANH SÁCH TẬP THỂ KHEN THƯỞNG")
    '        Else
    '            rpt_KhenThuongTT.SetParameterValue("NameKT", "DANH SÁCH TẬP THỂ ĐỀ NGHỊ KHEN THƯỞNG")
    '        End If
    '        rpt_KhenThuongTT.SetParameterValue("tinh", DIABAN)
    '        rpt_KhenThuongTT.SetParameterValue("ng", Now.Day)
    '        rpt_KhenThuongTT.SetParameterValue("th", Now.Month)
    '        rpt_KhenThuongTT.SetParameterValue("nm", Now.Year)
    '        rpt_KhenThuongTT.SetParameterValue("nam", vNam)
    '        rpt_KhenThuongTT.SetParameterValue("NguoiLapBieu", "Nguyen Thi Nhan")
    '        rptView.ReportSource = rpt_KhenThuongTT

    '    End If


    'End Sub
End Class