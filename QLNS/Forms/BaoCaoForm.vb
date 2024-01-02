Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Syncfusion.XlsIO
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Collections
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class BaoCaoForm
    Private _ChiNhanhBLL As ChiNhanhBLL = New ChiNhanhBLL()
    Private _KHLD_MangLuoiBLL As KHLD_MangLuoiBLL = New KHLD_MangLuoiBLL()
    Private _HeThongBLL As HeThongBLL = New HeThongBLL()
    Private _HsCanBoBLL As HsCanBoBLL = New HsCanBoBLL()
    Private _BaoCaoBLL As BaoCaoBLL = New BaoCaoBLL()
    Private _dateTime As Date
    Private ARL_BaoCao As ArrayList = New ArrayList


    Private Property DateTime(p1 As Integer, p2 As Integer, p3 As Integer, p4 As Integer, p5 As Integer, p6 As Integer, p7 As Integer) As Date
        Get
            Return _dateTime
        End Get
        Set(value As Date)
            _dateTime = value
        End Set
    End Property

#Region "---> Events <---"
    Private Sub BaoCaoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Bind danh sách các sao kê cần lấy theo yêu cầu
        cb_baocao.Items.Clear()
        ARL_BaoCao.Clear()
        cb_baocao.Items.Add("--- Chọn báo cáo ---")
        'cb_baocao.Items.Add("Tổng hợp tình hình thực hiện Lao động - Mạng lưới (01/BC-TCCB)")                    '0001 - Gồm cả chi tiết và Tổng hợp
        cb_baocao.Items.Add("Mẫu số 01/BC-TCCB. Tổng hợp Tình hình thực hiện lao động")                                    '0001A - Gồm cả chi tiết và Tổng hợp
        cb_baocao.Items.Add("Mẫu số 01A/BC-TCCB. Tổng hợp Tình hình thực hiện lao động")                                    '0001A - Gồm cả chi tiết và Tổng hợp
        cb_baocao.Items.Add("Mẫu số 05/BC-TCCB hoặc 06/BC-TCCB. Tổng hợp Tình hình thực hiện lao động bình quân năm")       '0002 - Gồm cả chi tiết và Tổng hợp
        'cb_baocao.Items.Add("Mẫu số 03/BC-TCCB. Thống kê số lượng và chất lượng cán bộ (03/BC-TCCB)")                       '0003
        cb_baocao.Items.Add("Mẫu số 09/BCLĐ hoặc 10/BCLĐ. Báo cáo lao động - Mạng lưới")                             '0004   'Index = 3
        cb_baocao.Items.Add("Mẫu số 08/BC-TCCB. Thống kê số lượng và chất lượng cán bộ (08/BC-TCCB)")   '0005                     'Index = 3
        cb_baocao.Items.Add("Mẫu số 02/BC-TCCB. Chi tiết tình hình thực hiện Lao động - Mạng lưới (02/BC-TCCB)")                    '0001_02
        cb_baocao.Items.Add("Mẫu số 03/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới theo vùng (03/BC-TCCB)")          '0001_03
        cb_baocao.Items.Add("Mẫu số 04/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới (04/BC-TCCB)")                    '0001_04
        cb_baocao.Items.Add("Mẫu số 07/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới chi tiết đến từng PGD (07/BC-TCCB)")                    '0001_07
        cb_baocao.Items.Add("Mẫu số 08/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới chi tiết đến từng PGD (08/BC-TCCB)")                    '0001_08

        ARL_BaoCao.Add("")
        ARL_BaoCao.Add("0001")
        ARL_BaoCao.Add("0001A")
        ARL_BaoCao.Add("0002")
        'ARL_BaoCao.Add("0003")
        ARL_BaoCao.Add("0004")
        ARL_BaoCao.Add("0005")
        ARL_BaoCao.Add("0001_02")
        ARL_BaoCao.Add("0001_03")
        ARL_BaoCao.Add("0001_04")
        ARL_BaoCao.Add("0001_07")
        ARL_BaoCao.Add("0001_08")


        'cb_baocao.Items.Add("Tiền lương. Mẫu 01A/TUL - Bảng kê chi lương - Kỳ I (Đối với Lao động Chuyên môn nghiệp vụ)")
        'cb_baocao.Items.Add("Tiền lương. Mẫu 01B/TUL - Bảng kê chi lương - Kỳ II (Đối với Lao động Chuyên môn nghiệp vụ)")
        'cb_baocao.Items.Add("Tiền lương. Mẫu 01B/TULNH - Bảng kê thanh toán tiền công, trực đêm của bảo vệ, lao công, tạp vụ")
        'cb_baocao.Items.Add("Tiền lương. Mẫu 01/TL - Báo cáo Tình hình thực hiện Lao động - Tiền lương (Tháng)")
        'cb_baocao.Items.Add("Tiền lương. Mẫu 02/TL - Báo cáo Tình hình thực hiện Lao động - Tiền lương (Năm)")
        'cb_baocao.Items.Add("Tiền lương. Mẫu 01/TU - Báo cáo Tạm ứng chi lương năm")
        'cb_baocao.Items.Add("Tiền lương. Mẫu 02/TU - Báo cáo Về việc tạm ứng chi lương năm")

        If Not (cb_baocao Is Nothing) Then
            If cb_baocao.Items.Count > 0 Then
                cb_baocao.SelectedIndex = 0
            End If
        End If

        'Danh sách đơn vị trực thuộc NHCSXH cần lấy sao kê
        cb_donvi.DataSource = _ChiNhanhBLL.GetListComBo_ChiNhanh(3, IIf(TRUCTHUOC = 0, "", "--- Chọn tất cả ---"))
        cb_donvi.DisplayMember = "Display"
        cb_donvi.ValueMember = "Value"
        If Not (cb_donvi Is Nothing) Then
            If cb_donvi.Items.Count > 0 Then
                cb_donvi.SelectedIndex = 0
            End If
        End If
        rb_tonghop.Checked = True
        cb_baocao_SelectedIndexChanged(sender, Nothing)
        Dim _DateTMP As DateTime = Globals.GetDateTime_ForServerDB
        dtpk_tungay.Value = DateTimeUtil.StringToDateTime("01-01-" + _DateTMP.Year.ToString(), "dd-MM-yyyy")
        dtpk_ngaybc.Value = DateTimeUtil.GetEndDateOfMonthOfDate(Date.Now)
    End Sub

    Private Sub cb_baocao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_baocao.SelectedIndexChanged
        lbl_tungay.Visible = True
        dtpk_tungay.Visible = True
        pnl_opt_tonghop_chitiet.Visible = False
        pnl_tonghop_cn_hay_pgd.Visible = False
        Dim _MaHieuBaoCao As String = ARL_BaoCao(IIf(cb_baocao.SelectedIndex > 0, cb_baocao.SelectedIndex, "0"))

        If _MaHieuBaoCao = "0002" Then
            lbl_ngaybc.Text = "Đến ngày "
            pnl_opt_tonghop_chitiet.Visible = True
        ElseIf _MaHieuBaoCao = "0005" Then
            pnl_tonghop_cn_hay_pgd.Visible = True
            lbl_ngaybc.Text = "Ngày báo cáo "
            lbl_tungay.Visible = False
            dtpk_tungay.Visible = False
        Else
            lbl_ngaybc.Text = "Ngày báo cáo "
            lbl_tungay.Visible = False
            dtpk_tungay.Visible = False
            If _MaHieuBaoCao = "0002" Then
                pnl_opt_tonghop_chitiet.Visible = True
            End If
            'If (_MaHieuBaoCao = "0001" And Cap_Nd = 1) Or _MaHieuBaoCao = "0002" Then
            '    pnl_opt_tonghop_chitiet.Visible = True
            'End If
        End If
    End Sub
#End Region

    Private Sub btn_xuatexcel_Click(sender As Object, e As EventArgs) Handles btn_xuatexcel.Click
        Dim _DateReportTMP As DateTime = Globals.GetDateTime_ForServerDB
        If cb_baocao.SelectedIndex = 1 Or cb_baocao.SelectedIndex = 2 Then
            If dtpk_ngaybc.Value.ToString("dd/MM/yyyy") <> DateTimeUtil.GetEndOfMonth(dtpk_ngaybc.Value.Month, dtpk_ngaybc.Value.Year).ToString("dd/MM/yyyy") Then
                MessageBox.Show("Bạn phải chọn ngày/thời điểm báo cáo là ngày cuối tháng. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return
            End If
        End If
        Dim _LapBieu_1 = "", _LapBieu_2 = "", _KiemSoat_1 = "", _KiemSoat_2 = "", _GiamDoc_1 = "", _GiamDoc_2 = ""
        Dim sTitleBC01 As String = "", sTitleBC02 As String = "", sTitleBC03 As String = "", sMaHieuBC As String = ""
        Dim iRowStart As Integer = 0, iThangBC As Integer, iNamBC As Integer, iRowTMP As Integer = 0, iTuThang As Integer = 0, iTuNam As Integer = 0, iSoThang As Integer = 0
        Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
        Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
        Dim sPath_Export As String = "C:\Temp\"
        Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
        sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")
        iThangBC = CType(dtpk_ngaybc.Value.Month, Integer)
        iNamBC = CType(dtpk_ngaybc.Value.Year, Integer)

        Dim sProvinceName As String = "", sBranchName As String = "", sThang01 As String = "", sThang02 As String = "", sThang03 As String = ""
        Dim iFlagCall As Integer = 0
        Dim db_report As DataTable = New DataTable()
        Dim _IdDonVi As Integer = CInt(cb_donvi.SelectedValue)
        Dim _MaHieuBaoCao As String = ARL_BaoCao(IIf(cb_baocao.SelectedIndex > 0, cb_baocao.SelectedIndex, "0"))


        Select Case _MaHieuBaoCao
            'Case "0001" ' Tình hình thực hiện lao động - Mạng lưới (01/BC-TCCB)
            '    iRowStart = 10
            '    If rb_tonghop.Checked = True Then
            '        db_report = _KHLD_MangLuoiBLL.KHLD_MangLuoi_GetSearchBC(Cap_Nd, DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), 0, "", 4, _IdDonVi, "")
            '    Else
            '        db_report = _KHLD_MangLuoiBLL.KHLD_MangLuoi_GetSearchBC(Convert.ToByte("6"), DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), 0, "", 4, _IdDonVi, "")
            '    End If

            '    If _IdDonVi > 0 Then
            '        sMaHieuBC = "Mẫu số 01A/BC-TCCB"
            '        sTitleBC01 = "BÁO CÁO CHI TIẾT TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI"
            '    Else
            '        sMaHieuBC = "Mẫu số 01B/BC-TCCB"
            '        sTitleBC01 = "BÁO CÁO TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI"
            '    End If
            '    sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
            '    sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_01_TCCB.xlsx"
            '    If rb_tonghop.Checked = True Then
            '        sFileNameEx = "BC_MangLuoiLD_01_TCCB_" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
            '    Else
            '        sFileNameEx = "BC_MangLuoiLD_01_TCCB_" + "Thang_Vung_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
            '    End If

            '    sColNameEnd = "R"
            '    _LapBieu_1 = "A"
            '    _LapBieu_2 = "B"
            '    _KiemSoat_1 = "H"
            '    _KiemSoat_2 = "J"
            '    _GiamDoc_1 = "O"
            '    _GiamDoc_2 = "R"
            '    If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
            '        sBranchName = db_report.Rows(0)("ChiNhanh_HT").ToString().ToUpper()
            '    End If
            Case "0001" ' Mẫu số 01/BC-TCCB - Tổng hợp Tình hình thực hiện lao động  Áp dụng từ 11/2022
                iRowStart = 12
                sMaHieuBC = "Mẫu số 01/BC-TCCB"
                iTuThang = 1
                iTuNam = CType(dtpk_tungay.Value.Year, Integer)
                iSoThang = iThangBC - iTuThang + 1
                sTitleBC01 = "BÁO CÁO TÌNH HÌNH THỰC HIỆN LAO ĐỘNG"
                Dim dNgayDauNam As DateTime = New DateTime(dtpk_ngaybc.Value.Year, 1, 1, 0, 0, 0, 0)
                db_report = _BaoCaoBLL.GetDLBaoCao_BaoCao_TinhHinhLD_01TCCB(DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, 1)
                If Not (db_report Is Nothing) Then
                    sTitleBC02 = db_report.Rows(0)("ThoiDiem_HT").ToString().ToUpper()
                Else : sTitleBC02 = ""
                End If
                If _IdDonVi = 0 Then
                    sFileExcelTemplate = sPathExcelTemplate + "BC_TinhHinhLaoDong_01_TCCB_TQ.xlsx"
                Else
                    sFileExcelTemplate = sPathExcelTemplate + "BC_TinhHinhLaoDong_01_TCCB_CN.xlsx"
                End If
                sFileNameEx = "BC_TinhHinhLaoDong_01_TCCB_ThoiDiem" + DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "ddMMyyyy") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "X"
                _LapBieu_1 = "A"
                _LapBieu_2 = "C"
                _KiemSoat_1 = "J"
                _KiemSoat_2 = "O"
                _GiamDoc_1 = "U"
                _GiamDoc_2 = "X"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    If _IdDonVi <> 1 Then
                        sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                    Else
                        sBranchName = db_report.Rows(0)("DonVi_HT").ToString().ToUpper()
                    End If
                End If
                sThang03 = dtpk_ngaybc.Value.Month.ToString("D2")
                sThang02 = dtpk_ngaybc.Value.AddMonths(-1).Month.ToString("D2")
                sThang01 = dtpk_ngaybc.Value.AddMonths(-2).Month.ToString("D2")
            Case "0002" ' 05/BC-TCCB và 06/BC-TCCB. Tổng hợp Tình hình thực hiện lao động bình quân năm - Áp dụng từ 11/2022
                iRowStart = 9
                sMaHieuBC = "Mẫu số 01C/BC-TCCB"
                iTuThang = CType(dtpk_tungay.Value.Month, Integer)
                iTuNam = CType(dtpk_tungay.Value.Year, Integer)
                iSoThang = iThangBC - iTuThang + 1
                If iTuThang = 1 Then
                    sTitleBC01 = "BÁO CÁO TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG BÌNH QUÂN NĂM " + iNamBC.ToString("D4")
                Else
                    sTitleBC01 = "BÁO CÁO TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG BÌNH QUÂN"
                End If
                If rb_tonghop.Checked = True Then
                    db_report = _BaoCaoBLL.GetDuLieu_BaoCao(_IdDonVi, DateTimeUtil.DateTimeToString(dtpk_tungay.Value, "yyyy-MM-dd"), DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), 1)
                    sMaHieuBC = "Mẫu số 05/BC-TCCB"
                Else
                    db_report = _BaoCaoBLL.GetDuLieu_BaoCao(_IdDonVi, DateTimeUtil.DateTimeToString(dtpk_tungay.Value, "yyyy-MM-dd"), DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), 2)
                    sMaHieuBC = "Mẫu số 06/BC-TCCB"
                End If
                If Not (db_report Is Nothing) Then
                    sTitleBC02 = db_report.Rows(0)("ThoiDiem_HT").ToString().ToUpper()
                Else : sTitleBC02 = ""
                End If

                sFileExcelTemplate = sPathExcelTemplate + "BC_LaoDongBinhQuan_01.xlsx"
                sFileNameEx = "BC_LaoDongBinhQuan_" + iTuThang.ToString("D2") + iTuNam.ToString("D4") + "_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "O"
                _LapBieu_1 = "A"
                _LapBieu_2 = "B"
                _KiemSoat_1 = "G"
                _KiemSoat_2 = "I"
                _GiamDoc_1 = "M"
                _GiamDoc_2 = "O"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    If _IdDonVi <> 1 Then
                        sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                    Else
                        sBranchName = db_report.Rows(0)("DonVi_HT").ToString().ToUpper()
                    End If
                End If
            Case "0001A" ' Mẫu số 01A/BC-TCCB - Tổng hợp Tình hình thực hiện lao động  Áp dụng từ 11/2022
                iRowStart = 10
                sMaHieuBC = "Mẫu số 01A/BC-TCCB"
                iTuThang = 1
                iTuNam = CType(dtpk_tungay.Value.Year, Integer)
                iSoThang = iThangBC - iTuThang + 1
                sTitleBC01 = "BÁO CÁO TÌNH HÌNH THỰC HIỆN LAO ĐỘNG"
                Dim dNgayDauNam As DateTime = New DateTime(dtpk_ngaybc.Value.Year, 1, 1, 0, 0, 0, 0)
                db_report = _BaoCaoBLL.GetDuLieu_BaoCao(_IdDonVi, DateTimeUtil.DateTimeToString(dNgayDauNam, "yyyy-MM-dd"), DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), 3)
                If Not (db_report Is Nothing) Then
                    sTitleBC02 = db_report.Rows(0)("ThoiDiem_HT").ToString().ToUpper()
                Else : sTitleBC02 = ""
                End If

                sFileExcelTemplate = sPathExcelTemplate + "BC_LaoDongBinhQuan_01A.xlsx"
                sFileNameEx = "BC_LaoDongBinhQuan_01A_" + iTuThang.ToString("D2") + iTuNam.ToString("D4") + "_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "S"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"
                _KiemSoat_1 = "H"
                _KiemSoat_2 = "K"
                _GiamDoc_1 = "O"
                _GiamDoc_2 = "S"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    If _IdDonVi <> 1 Then
                        sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                    Else
                        sBranchName = db_report.Rows(0)("DonVi_HT").ToString().ToUpper()
                    End If
                End If
            Case "0004" ' Mẫu số 09/BCLĐ. Báo cáo lao động - Mạng lưới
                iRowStart = 10
                iFlagCall = _IdDonVi
                If _IdDonVi = 0 Then
                    iFlagCall = 1
                Else : iFlagCall = 2
                End If
                db_report = _KHLD_MangLuoiBLL.BaoCao_KHLDMangLuoi09(Cap_Nd, DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, iFlagCall)
                sMaHieuBC = "Mẫu số 09/BCLĐ"
                sTitleBC01 = "BÁO CÁO LAO ĐỘNG - MẠNG LƯỚI "
                sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
                sTitleBC03 = "Số liệu tính đến " + DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "dd/MM/yyyy")

                sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_09_TCCB.xlsx"
                sFileNameEx = "BC_MangLuoiLD_09_TCCB_" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"

                sColNameEnd = "J"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"

                _KiemSoat_1 = "E"
                _KiemSoat_2 = "G"

                _GiamDoc_1 = "H"
                _GiamDoc_2 = "J"

                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                End If

            Case "0005" ' 08/BC-TCCB. Thống kê số lượng và chất lượng cán bộ (08/BC-TCCB) => Áp dụng từ 11/2022
                If (rb_tonghop_cn.Checked = True) Then
                    iFlagCall = 1
                ElseIf rb_tonghop_pgd.Checked = True Then
                    iFlagCall = 2
                Else : iFlagCall = 9
                End If
                db_report = _HsCanBoBLL.GetData_SaoKe(5, _IdDonVi, DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), iFlagCall)
                iRowStart = 10
                If iFlagCall = 1 Or iFlagCall = 2 Then
                    sMaHieuBC = "Mẫu số 08/BC-TCCB"
                    sTitleBC01 = "BÁO CÁO THỐNG KÊ SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ"
                    sTitleBC02 = "Thời điểm " + DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "dd/MM/yyyy")
                    sFileExcelTemplate = sPathExcelTemplate + "BC_SoLuong_ChatLuong_CanBo_TH_08_TCCB_2022.xlsx"
                    sFileNameEx = "BC_SoLuong_ChatLuong_08_TCCB" + "_ThoiDiem_" + DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "ddMMyyyy") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                    sColNameEnd = "AV"
                    _LapBieu_1 = "B"
                    _LapBieu_2 = "F"
                    _KiemSoat_1 = "Q"
                    _KiemSoat_2 = "Z"
                    _GiamDoc_1 = "AL"
                    _GiamDoc_2 = "AV"
                Else
                    sMaHieuBC = "Mẫu số 08/BC-TCCB"
                    sTitleBC01 = "BÁO CÁO THỐNG KÊ SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ"
                    sTitleBC02 = "Thời điểm " + DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "dd/MM/yyyy")
                    sFileExcelTemplate = sPathExcelTemplate + "SK_SoLuong_ChatLuong_CanBo_TH_08_TCCB_CT.xlsx"
                    sFileNameEx = "BC_SoLuong_ChatLuong_08_TCCB_ChiTiet_" + "ThoiDiem_" + DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "ddMMyyyy") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                    sColNameEnd = "R"
                    _LapBieu_1 = "A"
                    _LapBieu_2 = "C"
                    _KiemSoat_1 = "E"
                    _KiemSoat_2 = "I"
                    _GiamDoc_1 = "L"
                    _GiamDoc_2 = "R"
                End If
                If _IdDonVi <> 0 And _IdDonVi <> 1 Then
                    sBranchName = cb_donvi.SelectedText
                End If

            Case "0001_02" ' Mẫu số 02/BC-TCCB. Chi tiết tình hình thực hiện Lao động - Mạng lưới (02/BC-TCCB) - Áp dụng từ 11/2022
                iRowStart = 10
                db_report = _BaoCaoBLL.GetDLBaoCao_LaoDongMangLuoi(DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, 2)
                sMaHieuBC = "Mẫu số 02/BC-TCCB"
                sTitleBC01 = "BÁO CÁO CHI TIẾT TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI"
                sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
                sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_02_03_04_07_TCCB.xlsx"
                sFileNameEx = "BC_MangLuoiLD_02_TCCB" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "U"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"
                _KiemSoat_1 = "H"
                _KiemSoat_2 = "M"
                _GiamDoc_1 = "R"
                _GiamDoc_2 = "U"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                End If
            Case "0001_03" ' Mẫu số 03/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới theo vùng (03/BC-TCCB) - Áp dụng từ 11/2022
                iRowStart = 10
                db_report = _BaoCaoBLL.GetDLBaoCao_LaoDongMangLuoi(DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, 3)
                sMaHieuBC = "Mẫu số 03/BC-TCCB"
                sTitleBC01 = "BÁO CÁO TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI THEO VÙNG"
                sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
                sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_02_03_04_07_TCCB.xlsx"
                sFileNameEx = "BC_MangLuoiLD_03_TCCB" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "U"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"
                _KiemSoat_1 = "H"
                _KiemSoat_2 = "M"
                _GiamDoc_1 = "R"
                _GiamDoc_2 = "U"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    sBranchName = db_report.Rows(1)("TenChiNhanh_HT").ToString().ToUpper()
                End If
            Case "0001_04" ' Mẫu số 04/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới (04/BC-TCCB) - Áp dụng từ 11/2022
                iRowStart = 10
                db_report = _BaoCaoBLL.GetDLBaoCao_LaoDongMangLuoi(DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, 4)
                sMaHieuBC = "Mẫu số 04/BC-TCCB"
                sTitleBC01 = "BÁO CÁO TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI"
                sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
                sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_02_03_04_07_TCCB.xlsx"
                sFileNameEx = "BC_MangLuoiLD_04_TCCB" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "U"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"
                _KiemSoat_1 = "H"
                _KiemSoat_2 = "M"
                _GiamDoc_1 = "R"
                _GiamDoc_2 = "U"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                End If
            Case "0001_07" ' Mẫu số 07/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới chi tiết đến từng Phòng giao dịch (07/BC-TCCB) - Áp dụng từ 11/2022
                iRowStart = 10
                db_report = _BaoCaoBLL.GetDLBaoCao_LaoDongMangLuoi(DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, 7)
                sMaHieuBC = "Mẫu số 07/BC-TCCB"
                sTitleBC01 = "BÁO CÁO  TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI CHI TIẾT ĐẾN TỪNG PHÒNG GIAO DỊCH"
                sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
                sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_02_03_04_07_TCCB.xlsx"
                sFileNameEx = "BC_MangLuoiLD_07_TCCB" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "U"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"
                _KiemSoat_1 = "H"
                _KiemSoat_2 = "M"
                _GiamDoc_1 = "R"
                _GiamDoc_2 = "U"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                End If
            Case "0001_08" ' Mẫu số 07A/BC-TCCB. Tổng hợp tình hình thực hiện Lao động - Mạng lưới chi tiết đến từng Phòng giao dịch (07A/BC-TCCB) - Áp dụng từ 11/2023
                iRowStart = 10
                db_report = _BaoCaoBLL.GetDLBaoCao_LaoDongMangLuoi(DateTimeUtil.DateTimeToString(dtpk_ngaybc.Value, "yyyy-MM-dd"), _IdDonVi, 7)
                sMaHieuBC = "Mẫu số 08/BC-TCCB"
                sTitleBC01 = "BÁO CÁO  TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI CHI TIẾT ĐẾN TỪNG PHÒNG GIAO DỊCH"
                sTitleBC02 = "THÁNG " + iThangBC.ToString("D2") + " NĂM " + iNamBC.ToString("D4")
                sFileExcelTemplate = sPathExcelTemplate + "BC_MangLuoiLD_08_TCCB.xlsx"
                sFileNameEx = "BC_MangLuoiLD_07_TCCB" + "Thang_" + iThangBC.ToString("D2") + iNamBC.ToString("D4") + "_" + _DateReportTMP.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                sColNameEnd = "S"
                _LapBieu_1 = "A"
                _LapBieu_2 = "D"
                _KiemSoat_1 = "F"
                _KiemSoat_2 = "L"
                _GiamDoc_1 = "M"
                _GiamDoc_2 = "S"
                If (_IdDonVi <> 0) And Not (db_report Is Nothing) Then
                    sBranchName = db_report.Rows(0)("TenChiNhanh_HT").ToString().ToUpper()
                End If
        End Select

        If (db_report Is Nothing) Then
            MessageBox.Show("Không có dữ liệu để in báo cáo. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        ElseIf (db_report.Rows.Count <= 0) Then
            MessageBox.Show("Không có dữ liệu để in báo cáo. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If

        Try
            Cursor = Cursors.WaitCursor
            lbl_waitting.Text = "Waitting for data export..."
            'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
            sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
            
            If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
                Using excelEngine As ExcelEngine = New ExcelEngine()
                    Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                    Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                    Dim worksheet As IWorksheet = workbook.Worksheets(0)
                    worksheet.Range("A2").Text = sProvinceName.ToUpper()
                    worksheet.Range("A3").Text = sBranchName.ToUpper()
                    worksheet.Range("A4").Text = sTitleBC01
                    worksheet.Range("A5").Text = sTitleBC02
                    'If cb_baocao.SelectedIndex = 4 Then
                    '    worksheet.Range("F7").Text = sTitleBC03
                    '    worksheet.Range("F7:J7").CellStyle.Font.Bold = True
                    'End If

                    worksheet.Range(sColNameEnd + "1").Text = sMaHieuBC
                    If _MaHieuBaoCao = "0002" Then
                        worksheet.Range("O7").Text = "Lao động bình quân " + iSoThang.ToString() + " tháng"
                    End If
                    If _MaHieuBaoCao = "0001" Then
                        worksheet.Range("C8").Text = "Tháng " + sThang01
                        worksheet.Range("F8").Text = "Tháng " + sThang02
                        worksheet.Range("I8").Text = "Tháng " + sThang03
                        worksheet.Range("L8").Text = "Tháng " + sThang01
                        worksheet.Range("O8").Text = "Tháng " + sThang02
                        worksheet.Range("R8").Text = "Tháng " + sThang03
                    End If
                    'Create Template Marker Processor
                    Dim marker As ITemplateMarkersProcessor = workbook.CreateTemplateMarkersProcessor()
                    Dim columnNames(db_report.Columns.Count) As String
                    Dim i As Integer = 0
                    For Each column As DataColumn In db_report.Columns
                        columnNames(i) = column.ColumnName
                        i += 1
                    Next

                    For Each c As DataColumn In db_report.Columns
                        Dim columnName As String = c.ColumnName
                        Dim columnData As EnumerableRowCollection(Of Object)
                        If Globals.IsNumeric(c) Then
                            columnData = db_report.AsEnumerable().[Select](Function(r) If(r.Field(Of Object)(columnName), 0))
                        ElseIf c.DataType Is GetType(Date) Then
                            columnData = db_report.AsEnumerable.[Select](Function(r) If(r.Field(Of Object)(columnName), ""))
                        Else
                            columnData = db_report.AsEnumerable.[Select](Function(r) If(("'" & r.Field(Of Object)(columnName)), ""))
                        End If

                        Dim paramDataArray As Object() = columnData.ToArray
                        marker.AddVariable(columnName, paramDataArray)
                    Next
                    marker.ApplyMarkers()
                    workbook.Version = ExcelVersion.Excel2013

                    'Căn chỉnh bôi đậm, nghiêng các dòng bản ghi và điền dữ liệu vào một số CELL
                    iRowTMP = 0
                    If Not (db_report Is Nothing) Then
                        If (db_report.Rows.Count > 0) Then
                            For iTT As Integer = 0 To db_report.Rows.Count - 1
                                iRowTMP = (iRowStart + iTT)
                                Dim iKieuIn As Integer = CType(db_report.Rows(iTT)("KieuIn"), Integer)
                                If (iKieuIn = 0 Or iKieuIn = 1) Then
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Color = ExcelKnownColors.Red
                                ElseIf (iKieuIn = 2) Then
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Italic = True
                                ElseIf (iKieuIn = 4) Then
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Bold = True
                                ElseIf (iKieuIn = 5) Then
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Underline = ExcelUnderline.Single
                                End If
                            Next
                            'Xuất ghi chú nếu là mẫu 01A/BC-TCCB và 01B/BC-TCCB
                            If _MaHieuBaoCao = "0001ZZZ" Then
                                'SoLD_GiamTN với DonVi_Cd = '999999'
                                If _IdDonVi > 0 Then
                                    iRowTMP = iRowTMP + 1
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Text = " Ghi chú: Chi nhánh báo cáo cụ thể chi tiết số liệu sau"
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).RowHeight = 21

                                    '- Số lao động giảm tự nhiên trong tháng: ………. Người. Trong đó: Nghỉ hưu: ……… người; chấm dứt HĐLĐ: ……. người; chuyển đi chi nhánh khác: ……….. người
                                    Dim iSlgChamDutHD As Integer = 0
                                    iSlgChamDutHD = CType(db_report.Rows(0)("GiamTN_3702_SaThai").ToString(), Integer) + CType(db_report.Rows(0)("GiamTN_3705_ChuyenNganh").ToString(), Integer) + CType(db_report.Rows(0)("GiamTN_3706_Khac").ToString(), Integer)

                                    Dim iSlgChuyenCN As Integer = 0
                                    iSlgChuyenCN = CType(db_report.Rows(0)("GiamTN_1501_DieuDong").ToString(), Integer) + CType(db_report.Rows(0)("GiamTN_1502_DieuDongBoNhiem").ToString(), Integer) + CType(db_report.Rows(0)("GiamTN_1509_TiepNhan").ToString(), Integer) + CType(db_report.Rows(0)("GiamTN_1512_ThuyenChuyen").ToString(), Integer) + CType(db_report.Rows(0)("GiamTN_1532_LuanChuyenBN").ToString(), Integer)

                                    iRowTMP = iRowTMP + 1
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Text = " - Số lao động giảm tự nhiên trong tháng: " + db_report.Rows(0)("SoLD_GiamTuNhien_KhacCN").ToString().Trim() + " người. Trong đó: Nghỉ hưu: " + db_report.Rows(0)("GiamTN_3703_NghiHuu").ToString().Trim() + " người; chấm dứt HĐLĐ: " + iSlgChamDutHD.ToString() + " người; chuyển đi chi nhánh khác: " + iSlgChuyenCN.ToString() + " người."
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).RowHeight = 18

                                    Dim iSlgBoSung As Integer = 0
                                    iSlgChuyenCN = CType(db_report.Rows(0)("BoSung_TuyenDung").ToString(), Integer) + CType(db_report.Rows(0)("BoSung_ChuyenDenTuCNKhac").ToString(), Integer)
                                    iRowTMP = iRowTMP + 1
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Text = " - Số lao động được bổ sung trong tháng: " + iSlgChuyenCN.ToString() + " người. Trong đó: Tuyển dụng mới: " + db_report.Rows(0)("BoSung_TuyenDung").ToString() + " người; chuyển đến từ chi nhánh khác: " + db_report.Rows(0)("BoSung_ChuyenDenTuCNKhac").ToString() + " người."
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).RowHeight = 18

                                    iRowTMP = iRowTMP + 1
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Text = " - Số Phòng giao dịch/số quận, huyện thị xã, thành phố: (không bao gồm thị xã, thành phố do Hội sở tỉnh trực tiếp quản lý); ghi rõ tên quận, huyện, thị xã, thành phố chưa được chia tách, thành lập Phòng giao dịch (nếu có)."
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).WrapText = True
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).RowHeight = 30

                                End If
                            End If

                            'Xuất đoạn cuối: Lập biểu/Kiểm soát/Giám đốc
                            Dim db_sysvar As DataTable = New DataTable()
                            db_sysvar = _HeThongBLL.GetListSysVarSearch(0, _IdDonVi, "", "", "", "", 1)
                            If Not (db_sysvar Is Nothing) Then
                                If (db_sysvar.Rows.Count > 0) Then
                                    iRowTMP = iRowTMP + 1
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).Text = db_sysvar.Rows(0)("DiaBan").ToString().Trim() + ", ngày " + _DateReportTMP.Day.ToString("D2") + " tháng " + _DateReportTMP.Month.ToString("D2") + " năm " + _DateReportTMP.Year.ToString("D4") + "        "
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
                                    worksheet.Range(sColNameEnd + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).RowHeight = 23
                                    iRowTMP = iRowTMP + 1
                                    worksheet.Range(_LapBieu_1 + iRowTMP.ToString() + ":" + _LapBieu_2 + iRowTMP.ToString()).Text = "LẬP BIỂU"
                                    worksheet.Range(_LapBieu_1 + iRowTMP.ToString() + ":" + _LapBieu_2 + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range(_KiemSoat_1 + iRowTMP.ToString() + ":" + _KiemSoat_2 + iRowTMP.ToString()).Text = "KIỂM SOÁT"
                                    worksheet.Range(_KiemSoat_1 + iRowTMP.ToString() + ":" + _KiemSoat_2 + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range(_GiamDoc_1 + iRowTMP.ToString() + ":" + _GiamDoc_2 + iRowTMP.ToString()).Text = "GIÁM ĐỐC"
                                    worksheet.Range(_GiamDoc_1 + iRowTMP.ToString() + ":" + _GiamDoc_2 + iRowTMP.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
                                    worksheet.Range("A" + iRowTMP.ToString() + ":" + sColNameEnd + iRowTMP.ToString()).RowHeight = 20
                                End If
                            End If

                            If _MaHieuBaoCao = "0004" Then
                                worksheet.Range("F7:J7").Text = sTitleBC03
                                worksheet.Range("F7:J7").Merge(True)
                                worksheet.Range("F7:J7").CellStyle.Font.FontName = "Times New Roman"
                                worksheet.Range("F7:J7").CellStyle.Font.Size = 12
                                worksheet.Range("F7:J7").CellStyle.Font.Bold = True
                                worksheet.Range("F7:J7").CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                worksheet.Range("F7:J7").CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
                                worksheet.Range("F7:J7").RowHeight = 18
                            End If

                        End If
                    End If

                    If (Not System.IO.Directory.Exists(sPath_Export)) Then
                        Directory.CreateDirectory(sPath_Export)
                    End If

                    Dim file_name As String = sPath_Export + sFileNameEx
                    If Globals.FileDao.IsFile(file_name) Then
                        Globals.FileDao.DeleteFile(file_name)
                    End If
                    Dim fs As FileStream = File.Create(file_name)
                    workbook.SaveAs(fs)
                    workbook.Close()
                    streamRead.Close()
                    fs.Close()

                    MessageBox.Show("Xuất file báo cáo thành công!" + vbNewLine + "Đường dẫn chứa file báo cáo: [" + file_name + "]", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

                End Using
            Else
                MessageBox.Show("Không lấy được file excel mẫu [" + sFileExcelTemplate + "]. Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
            
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
            lbl_waitting.Text = ""
            Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show("Lỗi xuất báo cáo ra excel báo cáo: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi xuất báo cáo ra excel báo cáo BaoCaoForm\btn_xuatexcel_Click: " + ex.Message)
        End Try
    End Sub

    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        Close()
    End Sub
End Class