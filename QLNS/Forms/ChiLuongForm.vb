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
Public Class ChiLuongForm

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Defined parametter and properties <---"
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
    Public bCoLuuDL As Boolean = False

    Private vNFInfo As System.Globalization.NumberFormatInfo
    Private _Globals As Globals = New Globals
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
    Private _ChiLuongBLL As ChiLuongBLL = New ChiLuongBLL
    Private _HeThongBLL As HeThongBLL = New HeThongBLL

    Private ARL_PosCode As ArrayList = New ArrayList
    Private ARL_PhongBan As ArrayList = New ArrayList
    Private ARL_PhanLoai As ArrayList = New ArrayList

    Private _EventCall As Byte        'Biến lưu thông tin cờ: Thêm = 1 hay Sửa đổi = 2
    Private _ValueFrame As Byte        'Biến gọi định nghĩa lưới dữ liệu. Giá trị: 1 - Lương tháng kỳ 1; 2 - Lương tháng kỳ 2; 3 - Lương hợp đồng ngắn hạn; 4 - Lương thưởng (Lương bổ sung); 5 - Lương trọn gói tập nghề; 6 - Truy lĩnh lương
    Dim _ClassNumber As Integer = 5
    Public ValUpdate As Byte                                        '1 - Thêm mới; 2 - Sửa đổi
    Public CLTongHopId As Long
    Private ckb_ChoiceAll As CheckBox = Nothing                     'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0                              'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0                       'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False                  'Cờ báo việc Checkall
    Dim IsCalculate As Boolean = False                              'Cờ báo được phép tính toán trên lưới dữ liệu
    Dim sThongBaoCL As String = ""
    Dim ARL_FieldDecimal As New List(Of String)(New String() {"LCB_HeSo", "PhCap_ChVu_HeSo", "PhCap_TrNhiem_HeSo", "PhCap_DocHai_HeSo", "PhCap_ThuHut_HeSo", "PhCap_KhuVuc_HeSo", "BHXH_CB", "BHXH_DV", "BHYT_CB", "BHYT_DV", "BHTN_CB", "BHTN_DV", "DPCD_CB", "DPCD_LD_NN", "TNTT_TyLe", "MucTamUng_V2", "HeSoLuong_V2", "DonGia_LamDem_Gio", "LamDem_SoGio", "LamDem_SoTienPC", "MucHuong"})
    Dim ARL_FieldInteger As New List(Of String)(New String() {"PhCap_ThuHut_SoTien", "PhCap_KhuVuc_SoTien", "PhCap_KhuVuc_ST_Goc", "PhCap_ThuHut_ST_Goc", "SoNgayLViec_Thang", "SoNgay_KhongLV", "SoNgayNghi_TruLuong", "SoTienNghi_TruLuong", "Luong_TTV", "Luong_CSo", "Luong_V1", "Luong_V1_TamUng_K1", "Luong_V1_TamUng_K2", "BHXH_CB_SoTien", "BHXH_DV_SoTien", "BHYT_CB_SoTien", "BHYT_DV_SoTien", "BHTN_CB_SoTien", "BHTN_DV_SoTien", "DPCD_CB_SoTien", "DPCD_LD_NN_SoTien", "Tru_Khoan_Khac", "Tru_UngHo_Khac", "Tru_CacKhoan_Tong", "TNTT_Muc_BanThan", "TNTT_Muc_PhuThuoc", "TNTT_SoTien", "TNTT_SoTien_NopThue", "SoNguoi_PhuThuoc", "Luong_V1_ConLai", "Luong_V1_ThucTra", "Luong_V2_TamUng", "Luong_ThucLinh"})
    Dim ARL_FieldInputs As New List(Of String)(New String() {"LamDem_SoGio", "MucHuong", "Tru_Khoan_Khac", "Tru_UngHo_Khac", "SoNgayLViec_Thang", "SoNgay_KhongLV", "SoNgayNghi_TruLuong", "BHXH_CB_SoTien", "BHXH_DV_SoTien", "BHYT_CB_SoTien", "BHYT_DV_SoTien", "BHTN_CB_SoTien", "BHTN_DV_SoTien", "DPCD_CB_SoTien", "DPCD_LD_NN_SoTien"})

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        'vNFInfo.NumberDecimalDigits = 2
        'vNFInfo.NumberGroupSeparator = " "
    End Sub

#End Region

#Region "---> Hàm, thủ tục chính được sử dụng <---"
    Private Sub ResetAll_Controls()
        cb_loaichiluong.SelectedIndex = 0
        If Not (cb_loaichiluong Is Nothing) Then
            If cb_loaichiluong.Items.Count <> 0 Then
                cb_loaichiluong.SelectedIndex = 1
            End If
        End If

        cb_donvi.SelectedIndex = 0
        If Not (cb_donvi Is Nothing) Then
            If cb_donvi.Items.Count <> 0 Then
                cb_donvi.SelectedIndex = 1
            End If
        End If
        cb_donvi_SelectedIndexChanged(Nothing, Nothing)
        If cb_phongban.Items.Count <> 0 Then
            cb_phongban.SelectedIndex = 0
        End If
        Dim _NgayTMP As DateTime
        _NgayTMP = Globals.GetDateTime_ForServerDB
        dtpk_ngaybc.Text = _NgayTMP
        num_kybc.Value = IIf(_NgayTMP.Day > 20, 2, 1)
        num_kybc_ValueChanged(Nothing, Nothing)
        rb_daihan.Checked = True
        rb_daihan_CheckedChanged(Nothing, Nothing)
        rb_nganhan.Checked = False
        rb_tapsu.Checked = False
        edt_ghichu.Text = ""
        edt_bosung_sotien_khac.Text = "0"
        edt_bosung_ghichu_khac.Text = ""
        cb_bosung_loai.SelectedIndex = 0
    End Sub

    Private Sub AddHeaderCheckBox()
        ckb_ChoiceAll = New CheckBox()
        ckb_ChoiceAll.Size = New Size(15, 15)
        'Add the CheckBox into the DataGridView
        ckb_ChoiceAll.Checked = False
        dgv_main.Controls.Add(ckb_ChoiceAll)
    End Sub

    Private Sub Bind_DataGrids()
        Try
            Cursor = Cursors.WaitCursor
            lbl_infors.Text = "Waiting ..."

            Dim _LoaiChi_Cd As String = IIf(ARL_PhanLoai.Count > 0 And cb_loaichiluong.SelectedIndex >= 0, ARL_PhanLoai(cb_loaichiluong.SelectedIndex), "")
            Dim _PosCode As String = ""
            If cb_donvi.Items.Count <> 0 And ARL_PosCode.Count > 0 And cb_donvi.SelectedIndex >= 0 Then
                _PosCode = IIf(ARL_PosCode.Count > 0 And cb_donvi.SelectedIndex >= 0, ARL_PosCode(cb_donvi.SelectedIndex), "")
            End If
            Dim _PhBanCode As String = ""
            If (cb_phongban.Items.Count <> 0 And ARL_PhongBan IsNot Nothing) Then
                _PhBanCode = IIf(cb_phongban.Items.Count <> 0, ARL_PhongBan(cb_phongban.SelectedIndex), "")
            End If
            Dim _NgayBC As String = DateTimeUtil.DateTimeToString(GetDate_ForDieuKienCL(), "dd/MM/yyyy")
            If Not String.IsNullOrEmpty(_NgayBC) And _NgayBC <> "" Then
                _NgayBC = _NgayBC.Substring(0, 10)
            End If
            Dim _ChonCa As Byte = 0
            Dim _DonVi_TK_Cd As String = ""
            If (_PosCode = "" Or _PosCode = "0") Then
                _ChonCa = 1
                _DonVi_TK_Cd = DONVI
            Else : _DonVi_TK_Cd = _PosCode
            End If

            IsCalculate = False
            Dim db_luong As DataTable = New DataTable()
            Dim _ColumnVisible As String = ""

            If _ValueFrame = 1 Then
                _ColumnVisible = "Luong_TTV;Luong_TTV_Ma;MaCB;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;Loai_CB;LCB_NgHuong;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_NgHuong;PhCap_ThuHut_HeSo;PhCap_KhuVuc_HeSo;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_V1_TamUng_K2;BHXH_DV_SoTien;BHYT_DV_SoTien;BHTN_DV_SoTien;SoNguoi_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;DPCD_CB_SoTien;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;SoNgayLViec_Thang;Luong_V1_ConLai;Luong_V1_ThucTra;Luong_V2_TamUng;Tinh_Thue_TNCN;MucTamUng_V2;HeSoLuong_V2;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc;LCB_Tong_HeSo"
            ElseIf _ValueFrame = 2 Then
                _ColumnVisible = "MaCB;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;Loai_CB;LCB_BacLuong_Id;LCB_NgHuong;KieuIn;Luong_TTV_Ma;Luong_TTV;Luong_CSo;BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TrangThai;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_NgHuong;PhCap_ThuHut_HeSo;PhCap_KhuVuc_HeSo;BHXH_DV_SoTien;BHYT_DV_SoTien;BHTN_DV_SoTien;TNTT_TyLe;DPCD_LD_NN_SoTien;SoNgayLViec_Thang;Luong_V1;Luong_V1_TamUng_K1;Tru_CacKhoan_Tong;Tru_CacKhoan_GhiChu;BHXH_CB_SoTien;BHYT_CB_SoTien;BHTN_CB_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tinh_Thue_TNCN;HeSoLuong_V2;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc;SoNguoi_PhuThuoc;LCB_Tong_HeSo"
            ElseIf _ValueFrame = 3 Then
                _ColumnVisible = "Luong_TTV;Luong_TTV_Ma;MaCB;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;Loai_CB;LCB_NgHuong;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_NgHuong;PhCap_ThuHut_HeSo;PhCap_KhuVuc_HeSo;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_V1_TamUng_K2;BHXH_DV_SoTien;BHYT_DV_SoTien;BHTN_DV_SoTien;SoNguoi_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;DPCD_CB_SoTien;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;SoNgayLViec_Thang;Luong_V1_ConLai;Luong_V1_ThucTra;Luong_V2_TamUng;Tinh_Thue_TNCN;MucTamUng_V2;HeSoLuong_V2;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1_TamUng_K1;DonGia_LamDem_Gio;Loai_HDNH;SoNgay_KhongLV;MucHuong;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc"
            ElseIf _ValueFrame = 4 Then 'Chi lương thưởng
                _ColumnVisible = "KieuIn;TongHopId;ChiTietId;MaCB;NgaySinh;GioiTinh;ChiNhanh_Cd;ChiNhanh_HT;DonVi_Cd;DonVi_HT;PhongBan_Cd;PhongBan_HT;ChucVu_Cd;ChucVu_HT;ChuyenMon_Cd;ChuyenMon_HT;LoaiQD_Cd;LoaiQD_HT;LoaiHDLD_Cd;LoaiHDLD_HT;LCB_BacLuong_Id;LCB_BacLuong_HT;LCB_NgHuong;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_HeSo;PhCap_ThuHut_ST_Goc;PhCap_ThuHut_SoTien;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_HeSo;PhCap_KhuVuc_ST_Goc;PhCap_KhuVuc_SoTien;PhCap_KhuVuc_NgHuong;SoNgay_KhongLV;SoNgayLViec_Thang;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_TTV_Ma;Luong_TTV;Luong_CSo;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;BHXH_CB;BHXH_CB_SoTien;BHXH_DV;BHXH_DV_SoTien;BHYT_CB;BHYT_CB_SoTien;BHYT_DV;BHYT_DV_SoTien;BHTN_CB;BHTN_CB_SoTien;BHTN_DV;BHTN_DV_SoTien;DPCD_CB;DPCD_CB_SoTien;DPCD_LD_NN;DPCD_LD_NN_SoTien;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tru_CacKhoan_Tong;Tru_CacKhoan_GhiChu;Tinh_Thue_TNCN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;SoNguoi_PhuThuoc;Luong_V1_ConLai;Luong_V1_ThucTra;MucTamUng_V2;HeSoLuong_V2;MucHuong;Loai_CB;TrangThai;TrangThai_HT;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;Loai_HDNH_DB;Loai_HDNH_PT"
            ElseIf _ValueFrame = 5 Then 'Chi lương Trọn gói (tập nghề)
                _ColumnVisible = "KieuIn;MaCB;NgaySinh;GioiTinh;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;LCB_BacLuong_Id;LCB_NgHuong;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_HeSo;PhCap_ThuHut_ST_Goc;PhCap_ThuHut_SoTien;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_HeSo;PhCap_KhuVuc_ST_Goc;PhCap_KhuVuc_SoTien;PhCap_KhuVuc_NgHuong;SoNgay_KhongLV;SoNgayLViec_Thang;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_TTV_Ma;Luong_TTV;Luong_CSo;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;BHXH_CB;BHXH_CB_SoTien;BHXH_DV;BHXH_DV_SoTien;BHYT_CB;BHYT_CB_SoTien;BHYT_DV;BHYT_DV_SoTien;BHTN_CB;BHTN_CB_SoTien;BHTN_DV;BHTN_DV_SoTien;DPCD_CB;DPCD_CB_SoTien;DPCD_LD_NN;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tru_CacKhoan_Tong;Tru_CacKhoan_GhiChu;Tinh_Thue_TNCN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;SoNguoi_PhuThuoc;Luong_V1_ConLai;Luong_V1_ThucTra;MucTamUng_V2;HeSoLuong_V2;Luong_V2_TamUng;MucHuong;Loai_CB;Loai_CB_HT;TrangThai;TrangThai_HT;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;Loai_HDNH_DB;Loai_HDNH_PT"
            ElseIf _ValueFrame = 6 Then 'Truy lĩnh lương
                _ColumnVisible = "KieuIn;TongHopId;ChiTietId;MaCB;NgaySinh;GioiTinh;ChiNhanh_Cd;ChiNhanh_HT;DonVi_Cd;DonVi_HT;PhongBan_Cd;PhongBan_HT;ChucVu_Cd;ChucVu_HT;ChuyenMon_Cd;ChuyenMon_HT;LoaiQD_Cd;LoaiQD_HT;LoaiHDLD_Cd;LoaiHDLD_HT;LCB_Tong_HeSo;LCB_BacLuong_Id;LCB_BacLuong_HT;LCB_NgHuong;PhCap_ChVu_HeSo;PhCap_ChVu_NgHuong;PhCap_TrNhiem_HeSo;PhCap_TrNhiem_NgHuong;PhCap_DocHai_HeSo;PhCap_DocHai_NgHuong;PhCap_ThuHut_HeSo;PhCap_ThuHut_SoTien;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_HeSo;PhCap_KhuVuc_SoTien;PhCap_KhuVuc_NgHuong;SoNgay_KhongLV;SoNgayLViec_Thang;SoTienNghi_TruLuong;Luong_TTV_Ma;Luong_TTV;Luong_CSo;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;BHXH_CB;BHXH_DV;BHXH_DV_SoTien;BHYT_CB;BHYT_DV;BHYT_DV_SoTien;BHTN_CB;BHTN_DV;BHTN_DV_SoTien;DPCD_CB;DPCD_CB_SoTien;DPCD_LD_NN;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tru_CacKhoan_GhiChu;Tinh_Thue_TNCN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;SoNguoi_PhuThuoc;Luong_V1_ConLai;MucTamUng_V2;HeSoLuong_V2;MucHuong;Loai_CB;Loai_CB_HT;TrangThai;TrangThai_HT;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;Loai_HDNH_DB;Loai_HDNH_PT"
            End If

            If ValUpdate = 1 Then           'Trường hợp thêm mới
                If _ValueFrame = 4 Then 'Chi lương thưởng
                    db_luong = _ChiLuongBLL.GetChiLuongSearch(_LoaiChi_Cd, _NgayBC, dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _DonVi_TK_Cd, _PhBanCode, _ChonCa, 4)
                ElseIf _ValueFrame = 6 Then 'Truy lĩnh lương 
                    db_luong = _ChiLuongBLL.GetChiLuongSearch(_LoaiChi_Cd, _NgayBC, dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _DonVi_TK_Cd, _PhBanCode, _ChonCa, 6)
                Else                    'Chi lương tháng
                    If rb_daihan.Checked Then
                        db_luong = _ChiLuongBLL.GetChiLuongSearch(_LoaiChi_Cd, _NgayBC, dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _DonVi_TK_Cd, _PhBanCode, _ChonCa, _EventCall)
                    ElseIf rb_nganhan.Checked Then
                        db_luong = _ChiLuongBLL.GetChiLuongSearch(_LoaiChi_Cd, _NgayBC, dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _DonVi_TK_Cd, _PhBanCode, _ChonCa, 3)
                    ElseIf rb_tapsu.Checked Then
                        db_luong = _ChiLuongBLL.GetChiLuongSearch(_LoaiChi_Cd, _NgayBC, dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _DonVi_TK_Cd, _PhBanCode, _ChonCa, 5)
                    End If
                End If
            ElseIf ValUpdate = 2 Then       'Trường hợp sửa đổi
                db_luong = _ChiLuongBLL.ChiLuong_BangKeCT_GetSearch(CLTongHopId, 0, "", 0, 0, 99, "", "", 0, 1)
            End If

            'Nếu là lương thưởng sửa tiêu đề cột   "Tiền lương V2 (Hệ số 0,8; mức 100%)"     HeSoLuong_V2/MucTamUng_V2 
            'Hai số thập phân
            Dim sColumnCellNullN2 As String = ""
            If _ValueFrame = 4 Then 'Chi lương thưởng
                sColumnCellNullN2 = "LCB_Tong_HeSo;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;MucTamUng_V2;HeSoLuong_V2;MucHuong"
            ElseIf _ValueFrame = 6 Then 'Truy lĩnh lương
                sColumnCellNullN2 = "LCB_Tong_HeSo;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;MucTamUng_V2;HeSoLuong_V2;MucHuong;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc"
            Else
                sColumnCellNullN2 = "LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;MucTamUng_V2;HeSoLuong_V2;MucHuong"
            End If

            Dim ARL_ColNullN2() As String = Globals.Splip_Strings(sColumnCellNullN2, ";")

            Dim sColumnCellNullN5 As String = "BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN"
            Dim ARL_ColNullN5() As String = Globals.Splip_Strings(sColumnCellNullN5, ";")

            Dim sColumnCellNull As String = "PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_V1_ConLai;Luong_V1_ThucTra;Luong_V2_TamUng;TNTT_SoTien_NopThue;DPCD_CB_SoTien;BHXH_CB_SoTien;BHYT_CB_SoTien;BHTN_CB_SoTien;Tru_Khoan_Khac;Tru_CacKhoan_Tong;Luong_ThucLinh;Luong_CSo;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien"
            Dim ARL_ColNull() As String = Globals.Splip_Strings(sColumnCellNull, ";")
            Dim sColumnCellZero As String = ""
            If _ValueFrame = 6 Then
                sColumnCellZero = "Luong_V1;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;SoNgay_KhongLV"
            Else
                sColumnCellZero = "Luong_V1;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;SoNgay_KhongLV;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc"
            End If
            Dim ARL_ColZero() As String = Globals.Splip_Strings(sColumnCellZero, ";")

            dgv_main.Rows.Clear()
            Dim _KieuIn As Byte = 0
            Dim _ValueNote As Byte = 0
            Dim iCountRows As Integer = 0
            If Not (db_luong Is Nothing) Then
                If (db_luong.Rows.Count > 0) Then
                    For i As Integer = 0 To db_luong.Rows.Count - 1
                        dgv_main.Rows.Add()
                        _ValueNote = 0
                        dgv_main.Rows(i).Cells("cln_IdCanBo").Value = db_luong.Rows(i)("IdCanBo").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_KieuIn").Value = db_luong.Rows(i)("KieuIn").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_LCB_BacLuong_Id").Value = db_luong.Rows(i)("LCB_BacLuong_Id").ToString().Trim()
                        Dim ARL_Cols() As String = Globals.Splip_Strings(_ColumnVisible, ";")
                        For Each _Value As String In ARL_Cols
                            If Not String.IsNullOrEmpty(_Value) Then
                                dgv_main.Rows(i).Cells("cln_" + _Value).Value = db_luong.Rows(i)(_Value).ToString().Trim()
                            End If
                        Next
                        dgv_main.Rows(i).Cells("cln_STT").Value = db_luong.Rows(i)("STT").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_HoTen").Value = db_luong.Rows(i)("HoTen").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_ThongTin_HT").Value = db_luong.Rows(i)("ThongTin_HT").ToString().Trim()

                        'Fill dữ liệu nếu NULL hoặc rỗng cho thành [Rỗng]
                        For Each _ValueNull As String In ARL_ColNull
                            If Not String.IsNullOrEmpty(_ValueNull) Then
                                If IsNothing(db_luong.Rows(i)(_ValueNull)) Or String.IsNullOrEmpty(db_luong.Rows(i)(_ValueNull).ToString()) Or db_luong.Rows(i)(_ValueNull).ToString().Trim() = "" Or db_luong.Rows(i)(_ValueNull).ToString().Trim() = "0" Then
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull).Value = ""
                                Else
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull).Value = Double.Parse(db_luong.Rows(i)(_ValueNull).ToString(), Globals.cultureNum).ToString("N0", vNFInfo)
                                End If
                            End If
                        Next

                        For Each _ValueNull2 As String In ARL_ColNullN2
                            If Not String.IsNullOrEmpty(_ValueNull2) Then
                                If IsNothing(db_luong.Rows(i)(_ValueNull2)) Or String.IsNullOrEmpty(db_luong.Rows(i)(_ValueNull2).ToString()) Or db_luong.Rows(i)(_ValueNull2).ToString().Trim() = "" Or db_luong.Rows(i)(_ValueNull2).ToString().Trim() = "0" Then
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull2).Value = ""
                                Else
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull2).Value = Double.Parse(db_luong.Rows(i)(_ValueNull2).ToString(), Globals.cultureNum).ToString("N2", vNFInfo)
                                End If
                            End If
                        Next

                        For Each _ValueNull5 As String In ARL_ColNullN5
                            If Not String.IsNullOrEmpty(_ValueNull5) Then
                                If IsNothing(db_luong.Rows(i)(_ValueNull5)) Or String.IsNullOrEmpty(db_luong.Rows(i)(_ValueNull5).ToString()) Or db_luong.Rows(i)(_ValueNull5).ToString().Trim() = "" Or db_luong.Rows(i)(_ValueNull5).ToString().Trim() = "0" Then
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull5).Value = ""
                                Else
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull5).Value = Double.Parse(db_luong.Rows(i)(_ValueNull5).ToString(), Globals.cultureNum).ToString("N5", vNFInfo)
                                End If
                            End If
                        Next

                        'Fill dữ liệu nếu NULL hoặc rỗng cho thành 0
                        For Each _Value As String In ARL_ColZero
                            If Not String.IsNullOrEmpty(_Value) Then
                                If IsNothing(db_luong.Rows(i)(_Value)) Or String.IsNullOrEmpty(db_luong.Rows(i)(_Value).ToString()) Or db_luong.Rows(i)(_Value).ToString().Trim() = "" Then
                                    dgv_main.Rows(i).Cells("cln_" + _Value).Value = "0"
                                Else
                                    dgv_main.Rows(i).Cells("cln_" + _Value).Value = Double.Parse(db_luong.Rows(i)(_Value).ToString(), Globals.cultureNum).ToString("N0", vNFInfo)
                                End If
                            End If
                        Next
                        dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac_GhiChu").Value = db_luong.Rows(i)("Tru_Khoan_Khac_GhiChu").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_GhiChu").Value = db_luong.Rows(i)("Tru_CacKhoan_GhiChu").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_GhiChu_ChiLuong").Value = db_luong.Rows(i)("GhiChu_ChiLuong").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_Loai_CB_HT").Value = db_luong.Rows(i)("Loai_CB_HT").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_SoTK_NHCS").Value = db_luong.Rows(i)("SoTK_NHCS").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_TrangThai").Value = db_luong.Rows(i)("TrangThai").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_IsBHXH").Value = db_luong.Rows(i)("IsBHXH").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_IsBHYT").Value = db_luong.Rows(i)("IsBHYT").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_IsBHTN").Value = db_luong.Rows(i)("IsBHTN").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_IsDPCD").Value = db_luong.Rows(i)("IsDPCD").ToString().Trim()
                        If _ValueFrame = 3 Then
                            dgv_main.Rows(i).Cells("cln_Loai_HDNH_DB").Value = db_luong.Rows(i)("Loai_HDNH_DB").ToString().Trim()
                            dgv_main.Rows(i).Cells("cln_Loai_HDNH_PT").Value = db_luong.Rows(i)("Loai_HDNH_PT").ToString().Trim()
                        End If
                        _KieuIn = CType(db_luong.Rows(i)("KieuIn"), Byte)
                        If _ValueFrame = 1 Or _ValueFrame = 2 Then
                            If String.IsNullOrEmpty(db_luong.Rows(i)("LCB_BacLuong_Id").ToString()) Or db_luong.Rows(i)("LCB_BacLuong_Id").ToString() = "0" Or String.IsNullOrEmpty(db_luong.Rows(i)("Luong_TTV").ToString()) Or db_luong.Rows(i)("Luong_TTV").ToString() = "0" Then
                                _ValueNote = 1
                            Else
                                _ValueNote = 0
                            End If
                        ElseIf _ValueFrame = 3 Then
                            If String.IsNullOrEmpty(db_luong.Rows(i)("Luong_V1").ToString()) Or db_luong.Rows(i)("Luong_V1").ToString() = "0" Or String.IsNullOrEmpty(db_luong.Rows(i)("Luong_TTV").ToString()) Or db_luong.Rows(i)("Luong_TTV").ToString() = "0" Then
                                _ValueNote = 1
                            Else
                                _ValueNote = 0
                            End If
                        ElseIf _ValueFrame = 4 And Not IsNothing(db_luong.Rows(i)("Loai_CB")) Then
                            If db_luong.Rows(i)("Loai_CB").ToString().Trim() <> "1" Then
                                _ValueNote = 2    'Mầu xanh lá cây
                            End If
                        End If
                        _ChiLuongBLL.SetStyleRowGrid(i, dgv_main, _KieuIn, _ValueNote)
                        iCountRows = iCountRows + 1
                    Next
                End If
            End If
            If iCountRows > 0 Then
                IsCalculate = True
            End If
            TotalCheckBoxes = dgv_main.RowCount
            pnl_bosung_sotienkhac.Enabled = False
            If TotalCheckBoxes > 0 Then
                pnl_bosung_sotienkhac.Enabled = True
            End If
            TotalCheckedCheckBoxes = 0
            lbl_infors.Text = ""
            Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show("Chi lương Bind_DataGrids() lỗi: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Chi lương Bind_DataGrids() lỗi: " + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện tính tổng Giá trị theo cột dữ liệu truyền vào trên lưới dữ liệu.
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="pColumnName">Tên cột dữ liệu cần SUM trên lưới. Không truyền chuỗi "cln_". Ví dụ tên cột là "cln_Luong_ThucLinh" thì truyền vào "Luong_ThucLinh"</param>
    ''' <param name="pChiNhanh_Cd">Mã chi nhánh cần tính tổng</param>
    ''' <param name="pDonVi_Cd">Mã đơn vị cần tính tổng</param>
    ''' <param name="pPhongBan_Cd">Mã phòng ban cần tính tổng</param>
    ''' <returns>Tổng giá trị trả về</returns>
    ''' <remarks></remarks>
    Private Function GetTotalValue_ForPosCode(ByVal dgv_name As DataGridView, ByVal pColumnName As String, ByVal pChiNhanh_Cd As String, ByVal pDonVi_Cd As String, ByVal pPhongBan_Cd As String)
        Dim _RetVal As Double = 0
        Dim _SoTMP As Double = 0

        Dim _KieuIn As Byte = 0
        Dim _DonVi_Cd As String = ""
        Dim _ChiNhanh_Cd As String = ""
        Dim _PhongBan_Cd As String = ""
        Dim _ValueTMP As String = ""

        If dgv_name.Rows.Count > 1 Then
            For i As Integer = 0 To dgv_name.Rows.Count - 1
                _KieuIn = CType(dgv_name.Rows(i).Cells("cln_KieuIn").Value, Byte)
                _DonVi_Cd = dgv_name.Rows(i).Cells("cln_DonVi_Cd").Value.ToString()
                _ChiNhanh_Cd = dgv_name.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString()
                _PhongBan_Cd = dgv_name.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString()
                _SoTMP = 0
                _ValueTMP = dgv_name.Rows(i).Cells("cln_" & pColumnName).Value.ToString().Replace(" ", "").Replace(",", "")
                If String.IsNullOrEmpty(_ValueTMP) Then
                    _ValueTMP = "0"
                End If
                _SoTMP = CType(_ValueTMP, Double)

                If _KieuIn <> 0 And _KieuIn <> 1 Then
                    If String.IsNullOrEmpty(pPhongBan_Cd) And pChiNhanh_Cd = _ChiNhanh_Cd And pDonVi_Cd = _DonVi_Cd Then
                        _RetVal = _RetVal + _SoTMP
                    ElseIf (Not String.IsNullOrEmpty(pPhongBan_Cd)) And (pChiNhanh_Cd = _ChiNhanh_Cd And pDonVi_Cd = _DonVi_Cd And pPhongBan_Cd = _PhongBan_Cd) Then
                        _RetVal = _RetVal + _SoTMP
                    End If
                End If
            Next
        End If
        Return _RetVal
    End Function

    ''' <summary>
    ''' Hàm kiểm tra đã tồn tại Chi lương hay chưa theo các tham số truyền vào
    ''' </summary>
    ''' <param name="pNgayCL">Ngày báo cáo, chỉ tính với Chi lương không phải lương tháng</param>
    ''' <param name="pNamCL">Năm chi lương</param>
    ''' <param name="pThangCL">Tháng chi lương</param>
    ''' <param name="pKyCL">Kỳ chi lương</param>
    ''' <param name="pLoaiCL">Loại chi lương</param>
    ''' <param name="pDoiTuongCL">Đối tượng chi lương</param>
    ''' <param name="pDonViCdCL">Đơn vị chi lương</param>
    ''' <param name="pPhongBanCdCL">Phòng ban chi lương</param>
    ''' <returns>True - Đã tồn tại; False - Chưa tồn tại</returns>
    ''' <remarks></remarks>
    Private Function IsExistSalaries(ByVal pNgayCL As DateTime, ByVal pNamCL As Integer, ByVal pThangCL As Integer, ByVal pKyCL As Byte, ByVal pLoaiCL As String, ByVal pDoiTuongCL As String, ByVal pDonViCdCL As String, ByVal pPhongBanCdCL As String)
        Dim _Result As Byte = 0
        Dim sSQL As String = ""
        If pLoaiCL = "01" Then      'Nếu chi lương tháng không tính điều kiện ngày báo cáo
            sSQL = String.Format("Select IsNull(Count(TongHopId),0) From ChiLuong_TongHop Where TrangThai <> 0 And DonVi_CL_Cd='{0}' And PhongBan_CL_Cd='{1}' And NamBC={2} And ThangBC={3} And KyBC={4} And PhanLoai_Cd='{5}' And LaoDong_Cd='{6}'", pDonViCdCL, pPhongBanCdCL, pNamCL, pThangCL, pKyCL, pLoaiCL, pDoiTuongCL)
        Else
            sSQL = String.Format("Select IsNull(Count(TongHopId),0) From ChiLuong_TongHop Where TrangThai <> 0 And DonVi_CL_Cd='{0}' And PhongBan_CL_Cd='{1}' And NamBC={2} And ThangBC={3} And KyBC={4} And PhanLoai_Cd='{5}' And LaoDong_Cd='{6}' And Cast(NgayBC As Date)='{7}'", pDonViCdCL, pPhongBanCdCL, pNamCL, pThangCL, pKyCL, pLoaiCL, pDoiTuongCL, DateTimeUtil.DateTimeToString(pNgayCL, "yyyy-MM-dd"))
        End If
        Dim iCount As Integer = SoftSqlHelper.GetNumber(sSQL, 0)
        If iCount > 0 Then
            _Result = 1
        End If
        If pLoaiCL = "01" And pKyCL = 2 And pDoiTuongCL = "1" Then      'Nếu chi lương tháng - Kỳ 2 => Kiểm tra xem đã chi lương kỳ 1 chưa
            sSQL = String.Format("Select IsNull(Count(TongHopId),0) From ChiLuong_TongHop Where TrangThai <> 0 And DonVi_CL_Cd='{0}' And PhongBan_CL_Cd='{1}' And NamBC={2} And ThangBC={3} And KyBC={4} And PhanLoai_Cd='{5}' And LaoDong_Cd='{6}'", pDonViCdCL, pPhongBanCdCL, pNamCL, pThangCL, 1, pLoaiCL, pDoiTuongCL)
            iCount = SoftSqlHelper.GetNumber(sSQL, 0)
            If iCount <= 0 Then
                _Result = 2
            End If
        End If
        'Kiểm tra xem đã khai báo Tham số lương chưa
        sSQL = String.Format("Select IsNull(Count(*),0) From SysVar Where (Case When DonVi_Cd='000100' Then N'000199' Else DonVi_Cd End)='{0}' And Cast(NgayHL As Date) <= '{1}' And TrangThai=1", pDonViCdCL, DateTimeUtil.DateTimeToString(pNgayCL, "yyyy-MM-dd"))
        iCount = SoftSqlHelper.GetNumber(sSQL, 0)
        If iCount <= 0 Then
            _Result = 3
        End If

        Return _Result
    End Function

    Private Function IsValid_ChiLuong(ByVal pEventCall As Byte) As Boolean
        If (cb_loaichiluong.SelectedIndex <= 0 And cb_loaichiluong.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn Loại chi lương cho cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaichiluong
            Return False
        End If
        Dim _LoaiChi_Cd As String = IIf(ARL_PhanLoai.Count > 0, ARL_PhanLoai(cb_loaichiluong.SelectedIndex), "")
        If (_LoaiChi_Cd = "01" And CType(num_kybc.Value, Integer) <= 0 And CType(num_kybc.Value, Integer) > 2) Then
            MessageBox.Show("Kỳ chi lương không hợp lệ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaichiluong
            Return False
        End If

        If (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn Đơn vị chi lương cho cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_donvi
            Return False
        End If
        Dim _PosCode As String = IIf(ARL_PosCode.Count > 0, ARL_PosCode(cb_donvi.SelectedIndex), "")

        If _PosCode = "000100" Or _PosCode = "000199" Then
            If (cb_phongban.SelectedIndex <= 0 And cb_phongban.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn Ban CMNV để chi lương cho cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_phongban
                Return False
            End If
        End If
        If dgv_main Is Nothing Then
            MessageBox.Show("Bạn lập bảng kê chi lương của Đơn vị!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = btn_lapbang
            Return False
        End If
        If dgv_main.Rows.Count <= 1 Then
            MessageBox.Show("Bạn lập bảng kê chi lương của Đơn vị!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = btn_lapbang
            Return False
        End If
        Return True
    End Function

    Private Sub SetDateReport()
        Dim _NgayTMP As DateTime
        _NgayTMP = dtpk_ngaybc.Value
        _ValueFrame = 0
        dtpk_ngaybc.CustomFormat = "MM/yyyy"
        lbl_ngaybc.Text = "Tháng báo cáo "
        pnl_laodong.Visible = True
        lbl_laodong.Visible = True
        lbl_kybc.Text = "Kỳ BC "
        lbl_kybc.Width = 55
        num_kybc.Maximum = 2
        num_kybc.Minimum = 0
        If (cb_loaichiluong.SelectedIndex > 0 And cb_loaichiluong.Items.Count <> 0) Then
            Dim _LoaiChi_Cd As String = ARL_PhanLoai(cb_loaichiluong.SelectedIndex)
            If (_LoaiChi_Cd = "01") And (CType(num_kybc.Value, Integer) = 1 Or CType(num_kybc.Value, Integer) = 2) Then  'Luong tháng
                If rb_daihan.Checked = True Then
                    If CType(num_kybc.Value, Integer) = 1 Then
                        _ValueFrame = 1
                        dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("15/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy")
                    ElseIf CType(num_kybc.Value, Integer) = 2 Then
                        _ValueFrame = 2
                        dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("01/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy") 'DateTimeUtil.GetDateEndOfMonth(_NgayTMP.Month, _NgayTMP.Year)
                    End If
                    num_kybc.Enabled = True
                ElseIf rb_nganhan.Checked = True Then
                    _ValueFrame = 3
                    num_kybc.Enabled = False
                    dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("01/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy") 'DateTimeUtil.GetDateEndOfMonth(_NgayTMP.Month, _NgayTMP.Year)
                ElseIf rb_tapsu.Checked = True Then
                    _ValueFrame = 5
                    num_kybc.Enabled = False
                    dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("01/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy") 'DateTimeUtil.GetDateEndOfMonth(_NgayTMP.Month, _NgayTMP.Year)
                End If
            ElseIf _LoaiChi_Cd = "02" Then
                _ValueFrame = 4
                dtpk_ngaybc.Text = _NgayTMP
                dtpk_ngaybc.CustomFormat = "dd/MM/yyyy"
                lbl_ngaybc.Text = "Ngày lấy danh sách "
                num_kybc.Value = 2
                num_kybc.Enabled = False
                pnl_laodong.Visible = False
                lbl_laodong.Visible = False
            ElseIf _LoaiChi_Cd = "03" Then
                _ValueFrame = 6
                dtpk_ngaybc.Text = _NgayTMP
                dtpk_ngaybc.CustomFormat = "dd/MM/yyyy"
                lbl_ngaybc.Text = "Ngày báo cáo "
                lbl_kybc.Text = "Quý truy lĩnh "
                lbl_kybc.Width = 110
                num_kybc.Maximum = 4
                num_kybc.Minimum = 1
                num_kybc.Enabled = True
                'If CType(_NgayTMP.Month, Integer) = 1 Or CType(_NgayTMP.Month, Integer) = 2 Or CType(_NgayTMP.Month, Integer) = 3 Then
                '    num_kybc.Value = 1
                'ElseIf CType(_NgayTMP.Month, Integer) = 4 Or CType(_NgayTMP.Month, Integer) = 5 Or CType(_NgayTMP.Month, Integer) = 6 Then
                '    num_kybc.Value = 2
                'ElseIf CType(_NgayTMP.Month, Integer) = 7 Or CType(_NgayTMP.Month, Integer) = 8 Or CType(_NgayTMP.Month, Integer) = 9 Then
                '    num_kybc.Value = 3
                'Else
                '    num_kybc.Value = 4
                'End If
                pnl_laodong.Visible = False
                lbl_laodong.Visible = False
            End If
        End If
    End Sub

    Private Function GetDate_ForDieuKienCL() As DateTime
        Dim _DateReport As DateTime
        _DateReport = dtpk_ngaybc.Value
        If (cb_loaichiluong.SelectedIndex > 0 And cb_loaichiluong.Items.Count <> 0) Then
            Dim _LoaiChi_Cd As String = ARL_PhanLoai(cb_loaichiluong.SelectedIndex)
            If (_LoaiChi_Cd = "01") And (CType(num_kybc.Value, Integer) = 1 Or CType(num_kybc.Value, Integer) = 2) Then  'Luong tháng
                If rb_daihan.Checked = True Then
                    If CType(num_kybc.Value, Integer) = 1 Then
                        _DateReport = New DateTime(dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, 15)
                    ElseIf CType(num_kybc.Value, Integer) = 2 Then
                        _DateReport = DateTimeUtil.GetDateEndOfMonth(dtpk_ngaybc.Value.Month, dtpk_ngaybc.Value.Year)
                    End If
                ElseIf rb_nganhan.Checked = True Then
                    _DateReport = DateTimeUtil.GetDateEndOfMonth(dtpk_ngaybc.Value.Month, dtpk_ngaybc.Value.Year)
                End If
            Else
                dtpk_ngaybc.Text = _DateReport
            End If
        End If
        Return _DateReport
    End Function
#End Region

#Region "---> Các sự kiện Chính <---"
    Private Sub ChiLuongForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Bind danh sách Dữ liệu khi load forms
        Dim sSQL As String = ""
        If DONVI = "000100" Or DONVI = "000199" Or DONVI = "000101" Or DONVI = "000196" Or DONVI = "000197" Then
            sSQL = String.Format("Select (Case When Ma_So In ('000199','000100') Then '000199' Else Ma_So End) As Code,Ten_Goi Name,Id,Id_Goc,Ten_VT,Dia_Chi VungKT,Status,TT_SapXep,Ma_SapXep From ChiNhanh Where Status=1 And Ma_So Like '{0}%' Order By Ma_SapXep,Ma_So", DONVI)
        Else
            sSQL = String.Format("Select Ma_So As Code,Ten_Goi Name,Id,Id_Goc,Ten_VT,Dia_Chi VungKT,Status,TT_SapXep,Ma_SapXep From ChiNhanh Where Status=1 And Ma_So Like '{0}%' Order By Ma_SapXep,Ma_So", DONVI.Substring(0, 4))
        End If

        ARL_PosCode.Clear()
        cb_donvi.Items.Clear()
        ARL_PosCode = _Globals.BindList_ComBoBox_ListBox(cb_donvi, Nothing, sSQL, "--- Chọn đơn vị ---", 1)

        ARL_PhanLoai.Clear()
        cb_loaichiluong.Items.Clear()
        ARL_PhanLoai = _Globals.BindList_ComBoBox_ListBox(cb_loaichiluong, Nothing, IIf(DONVI = "000196", clsHT_DanhMuc.Sql_LoaiChiLuongALL, clsHT_DanhMuc.Sql_LoaiChiLuongBoEOD), "--- Chọn Phân loại ---", 1)

        ResetAll_Controls()
        pnl_bosung_sotienkhac.Enabled = False
        cb_phongban.Enabled = IIf(DONVI = "000100" Or DONVI = "000199", True, False)
        If ValUpdate = 1 Then               'Trường hợp thêm mới
            If _ValueFrame = 0 Then
                _ValueFrame = 1
            End If
            _ChiLuongBLL.Create_Frame(dgv_main, _ValueFrame)
        ElseIf ValUpdate = 2 Then           'Trường hợp sửa đổi
            _EventCall = 1
            cb_loaichiluong.Enabled = False
            num_kybc.Enabled = False
            dtpk_ngaybc.Enabled = False
            pnl_laodong.Enabled = False
            cb_donvi.Enabled = False
            cb_phongban.Enabled = False
            edt_ghichu.ReadOnly = True
            btn_lapbang.Enabled = False
            Dim db_cluong As DataTable = New DataTable()
            db_cluong = _ChiLuongBLL.GetChiLuong_TongHop_GetSearch("", "", "", 0, 0, 99, "", "", CLTongHopId)
            If Not (db_cluong Is Nothing) Then
                If (db_cluong.Rows.Count > 0) Then
                    cb_donvi.SelectedIndex = IIf(db_cluong.Rows(0)("DonVi_CL_Cd").ToString() <> "", ARL_PosCode.IndexOf(db_cluong.Rows(0)("DonVi_CL_Cd").ToString()), 0)
                    If Not String.IsNullOrEmpty(db_cluong.Rows(0)("PhongBan_CL_Cd")) Then
                        cb_phongban.SelectedIndex = IIf(db_cluong.Rows(0)("PhongBan_CL_Cd").ToString() <> "", ARL_PhongBan.IndexOf(db_cluong.Rows(0)("PhongBan_CL_Cd").ToString()), 0)
                    End If
                    cb_loaichiluong.SelectedIndex = IIf(db_cluong.Rows(0)("PhanLoai_Cd").ToString() <> "", ARL_PhanLoai.IndexOf(db_cluong.Rows(0)("PhanLoai_Cd").ToString()), 0)
                    num_kybc.Value = IIf(db_cluong.Rows(0)("KyBC").ToString() <> "", CType(db_cluong.Rows(0)("KyBC").ToString(), Integer), 0)
                    If (db_cluong.Rows(0)("LaoDong_Cd").ToString().Trim() <> "") Then
                        If db_cluong.Rows(0)("LaoDong_Cd").ToString().Trim() = "1" Then
                            rb_daihan.Checked = True
                            rb_daihan_CheckedChanged(sender, Nothing)
                        ElseIf db_cluong.Rows(0)("LaoDong_Cd").ToString().Trim() = "2" Then
                            rb_nganhan.Checked = True
                            rb_nganhan_CheckedChanged(sender, Nothing)
                        ElseIf db_cluong.Rows(0)("LaoDong_Cd").ToString().Trim() = "3" Then
                            rb_tapsu.Checked = True
                            rb_tapsu_CheckedChanged(sender, Nothing)
                        End If
                    End If
                    dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime(db_cluong.Rows(0)("NgayBC_HT").ToString(), "dd/MM/yyyy")
                    edt_ghichu.Text = db_cluong.Rows(0)("GhiChu_CL").ToString()
                End If
            End If
            'Fill Dữ liệu chi lương
            _ChiLuongBLL.Create_Frame(dgv_main, _ValueFrame)
            Bind_DataGrids()
        End If
        bCoLuuDL = False
        AddHeaderCheckBox()
        Me.dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        Me.dgv_main.ColumnHeadersHeight = Me.dgv_main.ColumnHeadersHeight * _ClassNumber
        Me.dgv_main.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        AddHandler Me.dgv_main.CellPainting, AddressOf dgv_main_CellPainting

        AddHandler Me.dgv_main.Paint, AddressOf dgv_main_Paint
        AddHandler Me.dgv_main.Scroll, AddressOf dgv_main_Scroll
        AddHandler Me.dgv_main.ColumnWidthChanged, AddressOf dgv_main_ColumnWidthChanged

        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged

        AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged
    End Sub

    Private Sub cb_donvi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_donvi.SelectedIndexChanged
        Try
            ARL_PhongBan.Clear()
            cb_phongban.Items.Clear()
            cb_phongban.Enabled = False
            'lbl_title.Text = "CHI TIẾT NỘI DUNG CHI " + IIf(String.IsNullOrEmpty(sThongBaoCL), "LƯƠNG", sThongBaoCL.ToUpper())
            If (cb_donvi.SelectedIndex > 0 And cb_donvi.Items.Count <> 0) Then
                Dim _PosCode As String = ARL_PosCode(cb_donvi.SelectedIndex)
                If _PosCode <> "" Then
                    Dim sTrucThuoc As String = _HS_CanBo.GetTrucThuoc(_PosCode)
                    Dim sSQL As String = ""
                    sSQL = String.Format("Select '60'+Ma_So Code,Ten_Phong Name,Id,Status,Truc_Thuoc,Ten_VT From PhongBan Where Status = 1 And Charindex('{0}',Truc_Thuoc) > 0 Order By Ma_So", sTrucThuoc)
                    ARL_PhongBan = _Globals.BindList_ComBoBox_ListBox(cb_phongban, Nothing, sSQL, "--- Phòng ban trực thuộc ---", 1)
                    If (_PosCode = "000100" Or _PosCode = "000199") And rb_daihan.Checked = True Then
                        cb_phongban.Enabled = True
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Điền dữ liệu danh sách Phòng ban: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub cb_loaichiluong_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_loaichiluong.SelectedIndexChanged
        Try
            sThongBaoCL = ""
            lbl_title.Text = "CHI TIẾT NỘI DUNG CHI LƯƠNG"
            If ValUpdate = 1 Then
                edt_ghichu.Text = sThongBaoCL
            End If
            Dim _NgayTMP As DateTime
            _NgayTMP = Globals.GetDateTime_ForServerDB
            _ValueFrame = 0
            dtpk_ngaybc.CustomFormat = "MM/yyyy"
            lbl_ngaybc.Text = "Tháng báo cáo "
            pnl_laodong.Visible = True
            lbl_laodong.Visible = True
            lbl_kybc.Text = "Kỳ BC "
            lbl_kybc.Width = 55
            num_kybc.Maximum = 2
            num_kybc.Minimum = 0
            If (cb_loaichiluong.SelectedIndex > 0 And cb_loaichiluong.Items.Count <> 0) Then
                Dim _LoaiChi_Cd As String = ARL_PhanLoai(cb_loaichiluong.SelectedIndex)
                Dim iKyBC As Integer = CType(num_kybc.Value, Integer)
                If _LoaiChi_Cd = "01" Then
                    sThongBaoCL = cb_loaichiluong.SelectedItem + " " + IIf(_LoaiChi_Cd = "01", "của kỳ " + iKyBC.ToString("D2"), "") + " đối với lao động " + IIf(rb_daihan.Checked = True, "Dài hạn", IIf(rb_nganhan.Checked = True, "Ngắn hạn", "Tập nghề/Tư vấn")) + " tháng " + dtpk_ngaybc.Value.Month.ToString("D2") + " năm " + dtpk_ngaybc.Value.Year.ToString("D4")
                ElseIf _LoaiChi_Cd = "02" Then
                    sThongBaoCL = cb_loaichiluong.SelectedItem + " ngày " + dtpk_ngaybc.Value.ToString("dd/MM/yyyy")
                ElseIf _LoaiChi_Cd = "03" Then
                    sThongBaoCL = cb_loaichiluong.SelectedItem + " quý " + CType(num_kybc.Value, Integer).ToString("D2") + " năm " + dtpk_ngaybc.Value.Year.ToString("D4")
                End If
                lbl_title.Text = "CHI TIẾT NỘI DUNG CHI " + sThongBaoCL.ToUpper()
                If ValUpdate = 1 Then
                    edt_ghichu.Text = sThongBaoCL.Replace(" năm ", "/")
                End If
                If (_LoaiChi_Cd = "01") And (CType(num_kybc.Value, Integer) = 1 Or CType(num_kybc.Value, Integer) = 2) Then  'Luong tháng
                    If rb_daihan.Checked = True Then
                        If CType(num_kybc.Value, Integer) = 1 Then
                            _ValueFrame = 1
                            dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("15/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy")
                        ElseIf CType(num_kybc.Value, Integer) = 2 Then
                            _ValueFrame = 2
                            dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("01/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy")
                        End If
                        num_kybc.Enabled = True
                    ElseIf rb_nganhan.Checked = True Then
                        _ValueFrame = 3
                        dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("01/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy")
                        num_kybc.Enabled = False
                    ElseIf rb_tapsu.Checked = True Then
                        _ValueFrame = 5
                        dtpk_ngaybc.Text = DateTimeUtil.StringToDateTime("01/" + _NgayTMP.Month.ToString("D2") + "/" + _NgayTMP.Year.ToString("D4"), "dd/MM/yyyy")
                        num_kybc.Enabled = False
                    End If
                ElseIf (_LoaiChi_Cd = "02") Then
                    _ValueFrame = 4
                    dtpk_ngaybc.Text = _NgayTMP
                    dtpk_ngaybc.CustomFormat = "dd/MM/yyyy"
                    lbl_ngaybc.Text = "Ngày lấy danh sách "
                    num_kybc.Value = 2
                    num_kybc.Enabled = False
                    pnl_laodong.Visible = False
                    lbl_laodong.Visible = False
                ElseIf (_LoaiChi_Cd = "03") Then
                    _ValueFrame = 6
                    dtpk_ngaybc.Text = _NgayTMP
                    dtpk_ngaybc.CustomFormat = "dd/MM/yyyy"
                    lbl_ngaybc.Text = "Ngày báo cáo "
                    lbl_kybc.Text = "Quý truy lĩnh "
                    lbl_kybc.Width = 110
                    num_kybc.Maximum = 4
                    num_kybc.Minimum = 1
                    num_kybc.Enabled = True
                    'If CType(_NgayTMP.Month, Integer) = 1 Or CType(_NgayTMP.Month, Integer) = 2 Or CType(_NgayTMP.Month, Integer) = 3 Then
                    '    num_kybc.Value = 1
                    'ElseIf CType(_NgayTMP.Month, Integer) = 4 Or CType(_NgayTMP.Month, Integer) = 5 Or CType(_NgayTMP.Month, Integer) = 6 Then
                    '    num_kybc.Value = 2
                    'ElseIf CType(_NgayTMP.Month, Integer) = 7 Or CType(_NgayTMP.Month, Integer) = 8 Or CType(_NgayTMP.Month, Integer) = 9 Then
                    '    num_kybc.Value = 3
                    'Else
                    '    num_kybc.Value = 4
                    'End If
                    pnl_laodong.Visible = False
                    lbl_laodong.Visible = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Lấy tự động ngày báo cáo: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lấy tự động ngày báo cáo [event cb_loaichiluong_SelectedIndexChanged] lỗi: " + ex.Message)
        End Try
    End Sub

    Private Sub num_kybc_ValueChanged(sender As Object, e As EventArgs) Handles num_kybc.ValueChanged
        cb_loaichiluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub rb_daihan_CheckedChanged(sender As Object, e As EventArgs) Handles rb_daihan.CheckedChanged
        If ValUpdate = 1 Then
            num_kybc.ReadOnly = False
            num_kybc.Enabled = True
        End If
        cb_loaichiluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub rb_nganhan_CheckedChanged(sender As Object, e As EventArgs) Handles rb_nganhan.CheckedChanged
        If ValUpdate = 1 Then
            num_kybc.Value = 2
            num_kybc.ReadOnly = True
            num_kybc.Enabled = False
        End If
        cb_loaichiluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub rb_tapsu_CheckedChanged(sender As Object, e As EventArgs) Handles rb_tapsu.CheckedChanged
        If ValUpdate = 1 Then
            num_kybc.Value = 2
            num_kybc.ReadOnly = True
            num_kybc.Enabled = False
        End If
        cb_loaichiluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub btn_lapbang_Click(sender As Object, e As EventArgs) Handles btn_lapbang.Click
        'System.Reflection.MethodBase.GetCurrentMethod()
        Dim _DoiTuongCd As String = IIf(rb_daihan.Checked = True, "1", IIf(rb_nganhan.Checked = True, "2", "3"))
        Dim _DoiTuongHT As String = IIf(rb_daihan.Checked = True, "Dài hạn", IIf(rb_nganhan.Checked = True, "Ngắn hạn", "Tập nghề/Tư vấn"))
        Dim _LoaiCLCd As String = IIf(ARL_PhanLoai.Count > 0, ARL_PhanLoai(cb_loaichiluong.SelectedIndex), "")
        Dim _PosCode As String = IIf(ARL_PosCode.Count > 0, ARL_PosCode(cb_donvi.SelectedIndex), "")
        Dim _PhBanCode As String = ""
        If (cb_phongban.Items.Count <> 0 And ARL_PhongBan IsNot Nothing) Then
            _PhBanCode = IIf(cb_phongban.Items.Count <> 0, ARL_PhongBan(cb_phongban.SelectedIndex), "")
        End If
        Dim _NgayBC As String = DateTimeUtil.DateTimeToString(GetDate_ForDieuKienCL(), "dd/MM/yyyy")
        _NgayBC = IIf(String.IsNullOrEmpty(_NgayBC), _NgayBC, _NgayBC.Substring(0, 10))
        Dim _DonVi_TK_Cd As String = ""
        If (_PosCode = "" Or _PosCode = "0") Then
            _DonVi_TK_Cd = DONVI
        Else : _DonVi_TK_Cd = _PosCode
        End If

        If String.IsNullOrEmpty(_LoaiCLCd) Then
            MessageBox.Show("Bạn chưa chọn Phân loại chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaichiluong
        ElseIf ((_LoaiCLCd = "01") And (CType(num_kybc.Value, Integer) <> 1 And CType(num_kybc.Value, Integer) <> 2)) Then
            MessageBox.Show("Bạn chưa chọn Kỳ báo cáo cho chi lương tháng. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = num_kybc
        ElseIf (String.IsNullOrEmpty(_DoiTuongCd)) Then
            MessageBox.Show("Bạn chưa chọn lao động cần chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = rb_daihan
        ElseIf (String.IsNullOrEmpty(_PosCode)) Then
            MessageBox.Show("Bạn chưa chọn đơn vị cần chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_donvi
        ElseIf ((_PosCode = "000100" Or _PosCode = "000199") And rb_daihan.Checked = True And String.IsNullOrEmpty(_PhBanCode)) Then
            MessageBox.Show("Bạn chưa chọn Ban CMNV cần chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_phongban
        ElseIf (IsExistSalaries(GetDate_ForDieuKienCL(), dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _LoaiCLCd, _DoiTuongCd, _PosCode, _PhBanCode)) = 1 Then
            MessageBox.Show("Đã tồn tại Chi [" + sThongBaoCL + "]!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaichiluong
        ElseIf (IsExistSalaries(GetDate_ForDieuKienCL(), dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _LoaiCLCd, _DoiTuongCd, _PosCode, _PhBanCode)) = 2 Then
            MessageBox.Show("Chưa thực hiện chi lương kỳ 1, nên không thể chi lương kỳ 2. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaichiluong
        ElseIf (IsExistSalaries(GetDate_ForDieuKienCL(), dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, CType(num_kybc.Value, Integer), _LoaiCLCd, _DoiTuongCd, _PosCode, _PhBanCode)) = 3 Then
            MessageBox.Show("Đơn vị " + cb_donvi.SelectedItem + " chưa khai báo tham số về lương" + vbNewLine + "Chú ý: Vui lòng đăng nhập bằng người dùng có quyền kiểm soát tại đơn vị và khai báo tham số theo Menu: Hệ thống\Tham số hệ thống!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaichiluong
        Else
            _EventCall = 1
            _ChiLuongBLL.Create_Frame(dgv_main, _ValueFrame)
            SetDateReport()
            Bind_DataGrids()
            dgv_main_CellBeginEdit(sender, Nothing)
        End If
    End Sub

    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        'Try
        '    'logger.Info("Form1_Load() - Start")
        '    'logger.Debug("Form1_Load() - Code Implementation goes here......")
        '    Globals.Logger.Error("Lỗi chương trình")
        '    Globals.Logger.Info("Info-Lỗi chương trình")
        '    Globals.Logger.Debug("Debug-Lỗi chương trình")
        '    Globals.WriteToLog("C:\Temp\QLNS_Info.log", "Lỗi chương trình chức năng thử thêm file log!")
        'Catch ex As Exception
        '    Globals.Logger.Error("Lỗi chương trình " + ex.Message)
        'End Try

        Close()
    End Sub

    Private Sub btn_luulai_Click(sender As Object, e As EventArgs) Handles btn_luulai.Click
        Try
            'Kiểm tra đã tồn tại - Cập nhật dữ liệu
            SetDateReport()
            If (IsValid_ChiLuong(_EventCall)) Then
                Dim _TongHopIdOut As Long = 0
                'Cập nhật dữ liệu vào Chi lương - Tổng hợp
                Dim _ChiLuong_TongHop As ChiLuong_TongHop = New ChiLuong_TongHop()
                _ChiLuong_TongHop.TongHopId = CLTongHopId
                Dim _PosCode As String = IIf(ARL_PosCode.Count > 0, ARL_PosCode(cb_donvi.SelectedIndex), "")
                _ChiLuong_TongHop.DonVi_CL_Cd = IIf(_PosCode = "000100", "000199", _PosCode)
                _ChiLuong_TongHop.PhongBan_CL_Cd = IIf(ARL_PhongBan.Count > 0, ARL_PhongBan(cb_phongban.SelectedIndex), "")
                _ChiLuong_TongHop.NgayBC = GetDate_ForDieuKienCL()
                _ChiLuong_TongHop.NamBC = dtpk_ngaybc.Value.Year
                _ChiLuong_TongHop.ThangBC = dtpk_ngaybc.Value.Month
                _ChiLuong_TongHop.KyBC = CType(num_kybc.Value, Integer)
                _ChiLuong_TongHop.PhanLoai_Cd = IIf(ARL_PhanLoai.Count > 0, ARL_PhanLoai(cb_loaichiluong.SelectedIndex), "")
                _ChiLuong_TongHop.LaoDong_Cd = IIf(rb_daihan.Checked = True, "1", IIf(rb_nganhan.Checked = True, "2", "3"))

                For i As Integer = 0 To dgv_main.Rows.Count - 1
                    If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 3 And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = _ChiLuong_TongHop.DonVi_CL_Cd Then
                        _ChiLuong_TongHop.Luong_TTV_Ma = dgv_main.Rows(i).Cells("cln_Luong_TTV_Ma").Value.ToString()
                        _ChiLuong_TongHop.Luong_TTV = CType(dgv_main.Rows(i).Cells("cln_Luong_TTV").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.Luong_CSo = CType(dgv_main.Rows(i).Cells("cln_Luong_CSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.BHXH_CB = CType(dgv_main.Rows(i).Cells("cln_BHXH_CB").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.BHXH_DV = CType(dgv_main.Rows(i).Cells("cln_BHXH_DV").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.BHYT_CB = CType(dgv_main.Rows(i).Cells("cln_BHYT_CB").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.BHYT_DV = CType(dgv_main.Rows(i).Cells("cln_BHYT_DV").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.BHTN_CB = CType(dgv_main.Rows(i).Cells("cln_BHTN_CB").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.BHTN_DV = CType(dgv_main.Rows(i).Cells("cln_BHTN_DV").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.DPCD_CB = CType(dgv_main.Rows(i).Cells("cln_DPCD_CB").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.DPCD_LD_NN = CType(dgv_main.Rows(i).Cells("cln_DPCD_LD_NN").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.TNTT_Muc_BanThan = CType(dgv_main.Rows(i).Cells("cln_TNTT_Muc_BanThan").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.TNTT_Muc_PhuThuoc = CType(dgv_main.Rows(i).Cells("cln_TNTT_Muc_PhuThuoc").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        _ChiLuong_TongHop.SoNgayLViec_Thang = CType(dgv_main.Rows(i).Cells("cln_SoNgayLViec_Thang").Value.ToString().Replace(",", "").Replace(" ", ""), Integer)
                        _ChiLuong_TongHop.TrangThai = CType(dgv_main.Rows(i).Cells("cln_TrangThai").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                        _ChiLuong_TongHop.Tinh_Thue_TNCN = CType(dgv_main.Rows(i).Cells("cln_Tinh_Thue_TNCN").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                        If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_MucTamUng_V2").Value.ToString()) Then
                            _ChiLuong_TongHop.MucTamUng_V2 = 0
                        Else
                            _ChiLuong_TongHop.MucTamUng_V2 = CType(dgv_main.Rows(i).Cells("cln_MucTamUng_V2").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        End If
                        If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_HeSoLuong_V2").Value.ToString()) Then
                            _ChiLuong_TongHop.HeSoLuong_V2 = 0
                        Else
                            _ChiLuong_TongHop.HeSoLuong_V2 = CType(dgv_main.Rows(i).Cells("cln_HeSoLuong_V2").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                        End If
                        Exit For
                    End If
                Next
                _ChiLuong_TongHop.GhiChu_CL = Globals.Find_Replace(edt_ghichu.Text.Trim.ToString())
                _ChiLuong_TongHop.CreatedBy = Globals.UserVal
                _ChiLuong_TongHop.ModifiedBy = Globals.UserVal

                'Kiểm tra nếu tồn tại thực hiện xóa trước khi cập nhật mới (Xóa cả Tổng hợp và Chi tiết luôn)
                If ValUpdate = 1 And CLTongHopId <= 0 Then
                    _ChiLuongBLL.Delete_UpdateStatus_ChiLuong_TongHop(0, 0, _ChiLuong_TongHop.ModifiedBy, _ChiLuong_TongHop.PhanLoai_Cd, _ChiLuong_TongHop.LaoDong_Cd, _ChiLuong_TongHop.DonVi_CL_Cd, _ChiLuong_TongHop.PhongBan_CL_Cd, _ChiLuong_TongHop.NamBC, _ChiLuong_TongHop.ThangBC, _ChiLuong_TongHop.KyBC, DateTimeUtil.DateTimeToString(_ChiLuong_TongHop.NgayBC, "dd/MM/yyyy"), 4)
                End If

                'Cập nhật dữ liệu
                _TongHopIdOut = _ChiLuongBLL.Insert_Update_ChiLuong_TongHop(_ChiLuong_TongHop)

                'Cập nhật dữ liệu vào Chi lương - Chi tiết
                Dim _ChiLuong_BangKeCT As ChiLuong_BangKeCT = New ChiLuong_BangKeCT()
                Dim _ChiTietIdOut As Long = 0
                Dim _CountChiTiet As Integer = 0
                If (_TongHopIdOut > 0) Then
                    If ValUpdate = 2 Then
                        _ChiLuongBLL.Delete_ChiLuong_BangKeCT(0, _TongHopIdOut, Globals.UserVal, 3)
                    End If
                    _ChiLuong_BangKeCT.ChiTietId = 0
                    _ChiLuong_BangKeCT.TongHopId = _TongHopIdOut

                    For i As Integer = 0 To dgv_main.Rows.Count - 1
                        If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 3 And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = _ChiLuong_TongHop.DonVi_CL_Cd Then
                            _ChiLuong_BangKeCT.IdCanBo = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_IdCanBo").Value.ToString())
                            _ChiLuong_BangKeCT.MaCB = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_MaCB").Value.ToString())
                            _ChiLuong_BangKeCT.HoTen = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_HoTen").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.ThongTin_HT = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_ThongTin_HT").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.SoTK_NHCS = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_SoTK_NHCS").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.ChiNhanh_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.DonVi_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.PhongBan_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.ChucVu_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_ChucVu_Cd").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.ChuyenMon_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_ChuyenMon_Cd").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.LoaiQD_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_LoaiQD_Cd").Value.ToString().Trim())
                            _ChiLuong_BangKeCT.LoaiHDLD_Cd = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_LoaiHDLD_Cd").Value.ToString().Trim())
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_LCB_BacLuong_Id").Value.ToString()) Then
                                _ChiLuong_BangKeCT.LCB_BacLuong_Id = 0
                            Else
                                _ChiLuong_BangKeCT.LCB_BacLuong_Id = CType(dgv_main.Rows(i).Cells("cln_LCB_BacLuong_Id").Value.ToString(), Integer)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_LCB_HeSo").Value.ToString()) Then
                                _ChiLuong_BangKeCT.LCB_HeSo = 0
                            Else
                                _ChiLuong_BangKeCT.LCB_HeSo = CType(dgv_main.Rows(i).Cells("cln_LCB_HeSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_LCB_NgHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.LCB_NgHuong = "1900-01-01"
                            Else
                                _ChiLuong_BangKeCT.LCB_NgHuong = DateTimeUtil.StringToDateTime(dgv_main.Rows(i).Cells("cln_LCB_NgHuong").Value.ToString(), "dd-MM-yyyy")
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_ChVu_HeSo").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_ChVu_HeSo = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_ChVu_HeSo = CType(dgv_main.Rows(i).Cells("cln_PhCap_ChVu_HeSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_ChVu_NgHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_ChVu_NgHuong = "1900-01-01"
                            Else
                                _ChiLuong_BangKeCT.PhCap_ChVu_NgHuong = DateTimeUtil.StringToDateTime(dgv_main.Rows(i).Cells("cln_PhCap_ChVu_NgHuong").Value.ToString(), "dd-MM-yyyy")
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_TrNhiem_HeSo").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_TrNhiem_HeSo = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_TrNhiem_HeSo = CType(dgv_main.Rows(i).Cells("cln_PhCap_TrNhiem_HeSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_TrNhiem_NgHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_TrNhiem_NgHuong = "1900-01-01"
                            Else
                                _ChiLuong_BangKeCT.PhCap_TrNhiem_NgHuong = DateTimeUtil.StringToDateTime(dgv_main.Rows(i).Cells("cln_PhCap_TrNhiem_NgHuong").Value.ToString(), "dd-MM-yyyy")
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_DocHai_HeSo").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_DocHai_HeSo = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_DocHai_HeSo = CType(dgv_main.Rows(i).Cells("cln_PhCap_DocHai_HeSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_DocHai_NgHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_DocHai_NgHuong = "1900-01-01"
                            Else
                                _ChiLuong_BangKeCT.PhCap_DocHai_NgHuong = DateTimeUtil.StringToDateTime(dgv_main.Rows(i).Cells("cln_PhCap_DocHai_NgHuong").Value.ToString(), "dd-MM-yyyy")
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_HeSo").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_ThuHut_HeSo = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_ThuHut_HeSo = CType(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_HeSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_ST_Goc").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_ThuHut_ST_Goc = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_ThuHut_ST_Goc = CType(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_ST_Goc").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_ThuHut_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_ThuHut_SoTien = CType(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_NgHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_ThuHut_NgHuong = "1900-01-01"
                            Else
                                _ChiLuong_BangKeCT.PhCap_ThuHut_NgHuong = DateTimeUtil.StringToDateTime(dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_NgHuong").Value.ToString(), "dd-MM-yyyy")
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_HeSo").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_HeSo = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_HeSo = CType(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_HeSo").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_ST_Goc").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_ST_Goc = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_ST_Goc = CType(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_ST_Goc").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_SoTien = CType(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_NgHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_NgHuong = "1900-01-01"
                            Else
                                _ChiLuong_BangKeCT.PhCap_KhuVuc_NgHuong = DateTimeUtil.StringToDateTime(dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_NgHuong").Value.ToString(), "dd-MM-yyyy")
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_SoNgayNghi_TruLuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.SoNgayNghi_TruLuong = 0
                            Else
                                _ChiLuong_BangKeCT.SoNgayNghi_TruLuong = CType(dgv_main.Rows(i).Cells("cln_SoNgayNghi_TruLuong").Value.ToString().Replace(",", "").Replace(" ", ""), Integer)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_SoNgay_KhongLV").Value.ToString()) Then
                                _ChiLuong_BangKeCT.SoNgay_KhongLV = 0
                            Else
                                _ChiLuong_BangKeCT.SoNgay_KhongLV = CType(dgv_main.Rows(i).Cells("cln_SoNgay_KhongLV").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_SoTienNghi_TruLuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.SoTienNghi_TruLuong = 0
                            Else
                                _ChiLuong_BangKeCT.SoTienNghi_TruLuong = CType(dgv_main.Rows(i).Cells("cln_SoTienNghi_TruLuong").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_V1").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_V1 = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_V1 = CType(dgv_main.Rows(i).Cells("cln_Luong_V1").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K1").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_V1_TamUng_K1 = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_V1_TamUng_K1 = CType(dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K1").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K2").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_V1_TamUng_K2 = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_V1_TamUng_K2 = CType(dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K2").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_BHXH_CB_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.BHXH_CB_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.BHXH_CB_SoTien = CType(dgv_main.Rows(i).Cells("cln_BHXH_CB_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_BHXH_DV_SoTien").Value.ToString()) Or _ChiLuong_BangKeCT.BHXH_CB_SoTien = 0 Then
                                _ChiLuong_BangKeCT.BHXH_DV_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.BHXH_DV_SoTien = CType(dgv_main.Rows(i).Cells("cln_BHXH_DV_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_BHYT_CB_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.BHYT_CB_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.BHYT_CB_SoTien = CType(dgv_main.Rows(i).Cells("cln_BHYT_CB_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_BHYT_DV_SoTien").Value.ToString()) Or _ChiLuong_BangKeCT.BHYT_CB_SoTien = 0 Then
                                _ChiLuong_BangKeCT.BHYT_DV_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.BHYT_DV_SoTien = CType(dgv_main.Rows(i).Cells("cln_BHYT_DV_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_BHTN_CB_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.BHTN_CB_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.BHTN_CB_SoTien = CType(dgv_main.Rows(i).Cells("cln_BHTN_CB_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_BHTN_DV_SoTien").Value.ToString()) Or _ChiLuong_BangKeCT.BHTN_CB_SoTien = 0 Then
                                _ChiLuong_BangKeCT.BHTN_DV_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.BHTN_DV_SoTien = CType(dgv_main.Rows(i).Cells("cln_BHTN_DV_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_DPCD_CB_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.DPCD_CB_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.DPCD_CB_SoTien = CType(dgv_main.Rows(i).Cells("cln_DPCD_CB_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_DPCD_LD_NN").Value.ToString()) Then
                                _ChiLuong_BangKeCT.DPCD_LD_NN = 0
                            Else
                                _ChiLuong_BangKeCT.DPCD_LD_NN = CType(dgv_main.Rows(i).Cells("cln_DPCD_LD_NN").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_DPCD_LD_NN_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.DPCD_LD_NN_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.DPCD_LD_NN_SoTien = CType(dgv_main.Rows(i).Cells("cln_DPCD_LD_NN_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Tru_Khoan_Khac = 0
                            Else
                                _ChiLuong_BangKeCT.Tru_Khoan_Khac = CType(dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            _ChiLuong_BangKeCT.Tru_Khoan_Khac_GhiChu = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac_GhiChu").Value.ToString())
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Tru_UngHo_Khac").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Tru_UngHo_Khac = 0
                            Else
                                _ChiLuong_BangKeCT.Tru_UngHo_Khac = CType(dgv_main.Rows(i).Cells("cln_Tru_UngHo_Khac").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            _ChiLuong_BangKeCT.Tru_UngHo_Khac_GhiChu = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_Tru_UngHo_Khac_GhiChu").Value.ToString())

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Tru_CacKhoan_Tong = 0
                            Else
                                _ChiLuong_BangKeCT.Tru_CacKhoan_Tong = CType(dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            _ChiLuong_BangKeCT.Tru_CacKhoan_GhiChu = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_GhiChu").Value.ToString())


                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_SoNguoi_PhuThuoc").Value.ToString()) Then
                                _ChiLuong_BangKeCT.SoNguoi_PhuThuoc = 0
                            Else
                                _ChiLuong_BangKeCT.SoNguoi_PhuThuoc = CType(dgv_main.Rows(i).Cells("cln_SoNguoi_PhuThuoc").Value.ToString().Replace(",", "").Replace(" ", ""), Integer)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_TNTT_SoTien").Value.ToString()) Then
                                _ChiLuong_BangKeCT.TNTT_SoTien = 0
                            Else
                                _ChiLuong_BangKeCT.TNTT_SoTien = CType(dgv_main.Rows(i).Cells("cln_TNTT_SoTien").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_TNTT_TyLe").Value.ToString()) Then
                                _ChiLuong_BangKeCT.TNTT_TyLe = 0
                            Else
                                _ChiLuong_BangKeCT.TNTT_TyLe = CType(dgv_main.Rows(i).Cells("cln_TNTT_TyLe").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_TNTT_SoTien_NopThue").Value.ToString()) Then
                                _ChiLuong_BangKeCT.TNTT_SoTien_NopThue = 0
                            Else
                                _ChiLuong_BangKeCT.TNTT_SoTien_NopThue = CType(dgv_main.Rows(i).Cells("cln_TNTT_SoTien_NopThue").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_V1_ConLai").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_V1_ConLai = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_V1_ConLai = CType(dgv_main.Rows(i).Cells("cln_Luong_V1_ConLai").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_V1_ThucTra").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_V1_ThucTra = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_V1_ThucTra = CType(dgv_main.Rows(i).Cells("cln_Luong_V1_ThucTra").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_MucHuong").Value.ToString()) Then
                                _ChiLuong_BangKeCT.MucHuong = 0
                            Else
                                _ChiLuong_BangKeCT.MucHuong = CType(dgv_main.Rows(i).Cells("cln_MucHuong").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_MucTamUng_V2").Value.ToString()) Then
                                _ChiLuong_BangKeCT.MucTamUng_V2 = 0
                            Else
                                _ChiLuong_BangKeCT.MucTamUng_V2 = CType(dgv_main.Rows(i).Cells("cln_MucTamUng_V2").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_V2_TamUng").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_V2_TamUng = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_V2_TamUng = CType(dgv_main.Rows(i).Cells("cln_Luong_V2_TamUng").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If

                            If String.IsNullOrEmpty(dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value.ToString()) Then
                                _ChiLuong_BangKeCT.Luong_ThucLinh = 0
                            Else
                                _ChiLuong_BangKeCT.Luong_ThucLinh = CType(dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            End If
                            _ChiLuong_BangKeCT.IsBHXH = CType(dgv_main.Rows(i).Cells("cln_IsBHXH").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                            _ChiLuong_BangKeCT.IsBHYT = CType(dgv_main.Rows(i).Cells("cln_IsBHYT").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                            _ChiLuong_BangKeCT.IsBHTN = CType(dgv_main.Rows(i).Cells("cln_IsBHTN").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                            _ChiLuong_BangKeCT.IsDPCD = CType(dgv_main.Rows(i).Cells("cln_IsDPCD").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                            _ChiLuong_BangKeCT.DonGia_LamDem_Gio = CType(dgv_main.Rows(i).Cells("cln_DonGia_LamDem_Gio").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            _ChiLuong_BangKeCT.LamDem_SoGio = CType(dgv_main.Rows(i).Cells("cln_LamDem_SoGio").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            _ChiLuong_BangKeCT.LamDem_SoTienPC = CType(dgv_main.Rows(i).Cells("cln_LamDem_SoTienPC").Value.ToString().Replace(",", "").Replace(" ", ""), Double)
                            _ChiLuong_BangKeCT.GhiChu_ChiLuong = Globals.Find_Replace(dgv_main.Rows(i).Cells("cln_GhiChu_ChiLuong").Value.ToString())
                            _ChiLuong_BangKeCT.Loai_CB = CType(dgv_main.Rows(i).Cells("cln_Loai_CB").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                            _ChiLuong_BangKeCT.TrangThai = CType(dgv_main.Rows(i).Cells("cln_TrangThai").Value.ToString().Replace(",", "").Replace(" ", ""), Byte)
                            _ChiLuong_BangKeCT.Loai_HDNH = CType(dgv_main.Rows(i).Cells("cln_Loai_HDNH").Value.ToString().Replace(",", "").Replace(" ", ""), Integer)
                            _ChiLuong_BangKeCT.CreatedBy = Globals.UserVal
                            _ChiLuong_BangKeCT.ModifiedBy = Globals.UserVal

                            _ChiTietIdOut = _ChiLuongBLL.Insert_Update_ChiLuong_BangKeCT(_ChiLuong_BangKeCT)
                            If _ChiTietIdOut > 0 Then
                                _CountChiTiet = _CountChiTiet + 1
                            End If
                        End If
                    Next
                    If _CountChiTiet > 0 And _TongHopIdOut > 0 Then
                        bCoLuuDL = True
                        Dim _DoiTuongHT As String = IIf(rb_daihan.Checked = True, "Dài hạn", IIf(rb_nganhan.Checked = True, "Ngắn hạn", "Tập nghề/Tư vấn"))
                        MessageBox.Show("Cập nhật thành công Chi [" + sThongBaoCL + "]!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi cập nhật chi lương: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            'System.Reflection.MethodBase.GetCurrentMethod()
            Globals.Logger.Error("Chi lương Lưu dữ liệu lỗi: " + ex.Message)
        End Try
    End Sub

    Private Sub btn_xoabo_Click(sender As Object, e As EventArgs) Handles btn_xoabo.Click
        'Xoa bỏ những bản ghi Tich CheckBox
        If (dgv_main.Rows.Count <= 0) Then Return
        Try
            Dim ARL_Deletes As ArrayList = New ArrayList()
            Dim _count As Int32 = 0
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_IdCanBo").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_IdCanBo").Value.ToString() <> "")) Then
                    If dgv_main.CurrentRow.Cells("cln_KieuIn").Value IsNot Nothing And CType(dgv_main.CurrentRow.Cells("cln_KieuIn").Value, Byte) = 3 Then
                        For i As Integer = dgv_main.RowCount - 1 To 0 Step -1
                            If (dgv_main.Rows(i).Cells("cln_IdCanBo").Value IsNot Nothing) Then
                                If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                    _count += 1
                                    ARL_Deletes.Add(dgv_main.Rows(i).Cells("cln_IdCanBo").Value.ToString())
                                    dgv_main.Rows.Remove(dgv_main.Rows(i))
                                    dgv_main.Refresh()
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            If (_count > 0) Then
                'Kiểm tra có trong CSDL xem có xóa luôn không
                ARL_Deletes.Clear()
                TotalCheckBoxes = dgv_main.RowCount
                ckb_ChoiceAll.Checked = False
                HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi Xoá bản ghi trên lưới dữ liệu chi lương: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi Xoá bản ghi trên lưới dữ liệu chi lương (ChiLuongForm\btn_xoabo_Click): " + ex.Message)
        End Try
    End Sub

    Private Sub ChiLuongForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub dtpk_ngaybc_ValueChanged(sender As Object, e As EventArgs) Handles dtpk_ngaybc.ValueChanged
        Dim _LoaiChi_Cd As String = ARL_PhanLoai(cb_loaichiluong.SelectedIndex)
        Dim iKyBC As Integer = CType(num_kybc.Value, Integer)
        If _LoaiChi_Cd = "03" Then
            sThongBaoCL = cb_loaichiluong.SelectedItem + " quý " + CType(num_kybc.Value, Integer).ToString("D2") + " năm " + dtpk_ngaybc.Value.Year.ToString("D4") + " đối với lao động " + IIf(rb_daihan.Checked = True, "Dài hạn", IIf(rb_nganhan.Checked = True, "Ngắn hạn", "Tập nghề/Tư vấn"))
        ElseIf _LoaiChi_Cd = "02" Then
            sThongBaoCL = cb_loaichiluong.SelectedItem + " ngày " + dtpk_ngaybc.Value.ToString("dd/MM/yyyy")
        Else
            sThongBaoCL = cb_loaichiluong.SelectedItem + " " + IIf(_LoaiChi_Cd = "01", "của kỳ " + iKyBC.ToString("D2"), "") + " đối với lao động " + IIf(rb_daihan.Checked = True, "Dài hạn", IIf(rb_nganhan.Checked = True, "Ngắn hạn", "Tập nghề/Tư vấn")) + " tháng " + dtpk_ngaybc.Value.Month.ToString("D2") + " năm " + dtpk_ngaybc.Value.Year.ToString("D4")
        End If

        lbl_title.Text = "CHI TIẾT NỘI DUNG CHI " + IIf(String.IsNullOrEmpty(sThongBaoCL), "LƯƠNG", sThongBaoCL.ToUpper())
        If ValUpdate = 1 Then
            edt_ghichu.Text = sThongBaoCL.Replace(" năm ", "/")
        End If
    End Sub
#End Region

#Region "---> Hàm, sự kiện liên quan tới Merge lưới và CheckAll lưới dữ liệu <---"
    Private Sub dgv_main_ColumnWidthChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_main.DisplayRectangle
        rtHeader.Height = Me.dgv_main.ColumnHeadersHeight / 2
        Me.dgv_main.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_main_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_main.DisplayRectangle
        rtHeader.Height = Me.dgv_main.ColumnHeadersHeight / 2
        Me.dgv_main.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_main_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs)
        If e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then
            Dim r2 As Rectangle = e.CellBounds
            r2.Y += e.CellBounds.Height / 3
            r2.Height = e.CellBounds.Height / 3
            e.PaintBackground(r2, True)
            e.PaintContent(r2)
            e.Handled = True
        End If

        'Căn chỉnh cho Checkbox nằm vào giữa tiêu đề của cột Chọn xoá
        If (e.RowIndex = -1 And e.ColumnIndex = 1) Then
            ResetHeaderCheckBoxLocation(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    Private Sub dgv_main_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim format As StringFormat = New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center
        Dim rHeSoLuongHeight As Integer = 0

        'Chỉ lặp những cột nào cần merge
        If (_ValueFrame = 1) Then               'Lương tháng kỳ 1
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 5 Then       'Hệ số lương và phụ cáp lương
                    Dim rHeSoLuong As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rHeSoLuong.X += 1
                    rHeSoLuong.Y += 1
                    rHeSoLuong.Width = rHeSoLuong.Width + w1 + w2 + w3 - 1
                    rHeSoLuong.Height = rHeSoLuong.Height / 3 - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rHeSoLuong.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rHeSoLuong)
                    e.Graphics.DrawRectangle(Pens.Silver, rHeSoLuong)
                    e.Graphics.DrawString("Hệ số lương và phụ cấp lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rHeSoLuong, format)
                End If

                If j = 6 Then       'Phụ cấp lương
                    Dim rPhuCap As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += rHeSoLuongHeight + 1
                    rPhuCap.Width = rPhuCap.Width + w1 + w2 - 1
                    rPhuCap.Height = rPhuCap.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Phụ cấp lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If

                If j = 11 Then       'Phụ cấp khác
                    Dim rPhuCap As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += 1
                    rPhuCap.Width = rPhuCap.Width + w1 - 1
                    rPhuCap.Height = rPhuCap.Height / 3 - 2 ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Phụ cấp khác", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                                         New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If

                If j = 15 Then       'Các khoản phải trừ
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                    Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width

                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 + w2 + w3 + w4 + w5 + w6 - 1
                    rPhaiTru.Height = rPhaiTru.Height / 3 - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("Các khoản phải trừ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If

            Next
        ElseIf (_ValueFrame = 2) Then       'Lương tháng kỳ 2
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 5 Then       'Hệ số lương và phụ cấp lương
                    Dim rHeSoLuong As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rHeSoLuong.X += 1
                    rHeSoLuong.Y += 1
                    rHeSoLuong.Width = rHeSoLuong.Width + w1 + w2 + w3 - 1
                    rHeSoLuong.Height = rHeSoLuong.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rHeSoLuong.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rHeSoLuong)
                    e.Graphics.DrawRectangle(Pens.Silver, rHeSoLuong)
                    e.Graphics.DrawString("Hệ số lương và phụ cấp lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rHeSoLuong, format)
                End If

                If j = 6 Then       'Phụ cấp lương
                    Dim rPhuCap As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += rHeSoLuongHeight + 1
                    rPhuCap.Width = rPhuCap.Width + w1 + w2 - 1
                    rPhuCap.Height = rPhuCap.Height / _ClassNumber - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Phụ cấp lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If

                If j = 11 Then       'Phụ cấp khác
                    Dim rPhuCap As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += 1
                    rPhuCap.Width = rPhuCap.Width + w1 - 1
                    rPhuCap.Height = rPhuCap.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Phụ cấp khác", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                                         New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If

                If j = 13 Then       'Tiền lương V1 còn lại (kỳ này)
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 + w2 + w3 + w4 - 1
                    rPhaiTru.Height = rPhaiTru.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("Tiền lương V1 còn lại (kỳ này)", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If

                If j = 17 Then       'Tạm ứng tiền lương V2
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 + w2 + w3 - 1
                    rPhaiTru.Height = rPhaiTru.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("Tạm ứng tiền lương V2", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If
                If j = 20 Then       'Các khoản phải trừ
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 + w2 + w3 - 1
                    rPhaiTru.Height = rPhaiTru.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("Các khoản phải trừ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If
            Next

        ElseIf (_ValueFrame = 3) Then   'Lao động ngăn hạn
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 4 Then       'Các khoản phải trừ
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 + w2 - 1
                    rPhaiTru.Height = rPhaiTru.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("Vị trí công việc", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If

                If j = 8 Then       'Các khoản phải trừ
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                    Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width
                    Dim w7 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 7, -1, True).Width

                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 + w2 + w3 + w4 + w5 + w6 + w7 - 1
                    rPhaiTru.Height = rPhaiTru.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("CÁC KHOẢN PHẢI TRỪ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If
                If j = 16 Then       'Các khoản phải trừ
                    Dim rPhaiTru As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rPhaiTru.X += 1
                    rPhaiTru.Y += 1
                    rPhaiTru.Width = rPhaiTru.Width + w1 - 1
                    rPhaiTru.Height = rPhaiTru.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rPhaiTru.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhaiTru)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhaiTru)
                    e.Graphics.DrawString("PHỤ CẤP LÀM ĐÊM", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhaiTru, format)
                End If
            Next
        ElseIf (_ValueFrame = 5) Then   'Lương lao động Tập nghề (Trọn gói)
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 5 Then       'Hệ số lương và phụ cấp lương
                    Dim rHeSoLuong As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rHeSoLuong.X += 1
                    rHeSoLuong.Y += 1
                    rHeSoLuong.Width = rHeSoLuong.Width + w1 + w2 + w3 - 1
                    rHeSoLuong.Height = rHeSoLuong.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    rHeSoLuongHeight = rHeSoLuong.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rHeSoLuong)
                    e.Graphics.DrawRectangle(Pens.Silver, rHeSoLuong)
                    e.Graphics.DrawString("Hệ số lương và phụ cấp lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rHeSoLuong, format)
                End If

                If j = 6 Then       'Phụ cấp lương
                    Dim rPhuCap As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += rHeSoLuongHeight + 1
                    rPhuCap.Width = rPhuCap.Width + w1 + w2 - 1
                    rPhuCap.Height = rPhuCap.Height / _ClassNumber - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Phụ cấp lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If

                If j = 10 Then       'Các khoản phải trừ
                    Dim rPhuCap As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += 1
                    rPhuCap.Width = rPhuCap.Width + w1 - 1
                    rPhuCap.Height = rPhuCap.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Các khoản phải trừ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                                         New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If
            Next

        ElseIf (_ValueFrame = 4) Then   'Lương thưởng
            Dim rHeSoLuong As Rectangle = New Rectangle()
            Dim rTrongDo As Rectangle = New Rectangle()
            Dim rPhuCap As Rectangle = New Rectangle()
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 5 Then       'Hệ số lương
                    rHeSoLuong = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    rHeSoLuong.X += 1
                    rHeSoLuong.Y += 1
                    rHeSoLuong.Width = rHeSoLuong.Width + w1 + w2 + w3 + w4 - 1
                    rHeSoLuong.Height = (rHeSoLuong.Height / _ClassNumber - 2) '* 2
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rHeSoLuong)
                    e.Graphics.DrawRectangle(Pens.Silver, rHeSoLuong)
                    e.Graphics.DrawString("Hệ số lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rHeSoLuong, format)
                End If
                If j = 6 Then       'Trong đó
                    rTrongDo = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rTrongDo.X += 1
                    rTrongDo.Y += rHeSoLuong.Height + 1

                    rTrongDo.Width = rTrongDo.Width + w1 + w2 + w3 - 1
                    rTrongDo.Height = rTrongDo.Height / _ClassNumber - 2         '3 la vi 3 lop

                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTrongDo)
                    e.Graphics.DrawRectangle(Pens.Silver, rTrongDo)
                    e.Graphics.DrawString("Trong đó", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTrongDo, format)
                End If
                If j = 7 Then       'Phụ cấp lương
                    rPhuCap = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rPhuCap.X += 1
                    rPhuCap.Y += rHeSoLuong.Height + rTrongDo.Height + 1
                    rPhuCap.Width = rPhuCap.Width + w1 + w2 - 1
                    rPhuCap.Height = (rPhuCap.Height / _ClassNumber - 2)
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                    e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                    e.Graphics.DrawString("Phụ cấp", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
                End If

            Next
        ElseIf (_ValueFrame = 6) Then   'Truy lĩnh lương
            Dim rHeSoLuong As Rectangle = New Rectangle()
            Dim rTrongDo As Rectangle = New Rectangle()
            Dim rTruCacKhoan As Rectangle = New Rectangle()
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 5 Then       'Hệ số lương
                    rHeSoLuong = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                    Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width
                    rHeSoLuong.X += 1
                    rHeSoLuong.Y += 1
                    rHeSoLuong.Width = rHeSoLuong.Width + w1 + w2 + w3 + w4 + w5 + w6 - 1
                    rHeSoLuong.Height = (rHeSoLuong.Height / _ClassNumber - 2) '* 2
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rHeSoLuong)
                    e.Graphics.DrawRectangle(Pens.Silver, rHeSoLuong)
                    e.Graphics.DrawString("Truy lĩnh lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rHeSoLuong, format)
                End If
                If j = 5 Then       'HSL
                    rTrongDo = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rTrongDo.X += 1
                    rTrongDo.Y += rHeSoLuong.Height + 1
                    rTrongDo.Width = rTrongDo.Width + w1 + w2 - 1
                    rTrongDo.Height = rTrongDo.Height / _ClassNumber - 2         '3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTrongDo)
                    e.Graphics.DrawRectangle(Pens.Silver, rTrongDo)
                    e.Graphics.DrawString("Hệ số lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTrongDo, format)
                End If

                If j = 12 Then       'Các khoản phải trừ
                    rTruCacKhoan = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width

                    rTruCacKhoan.X += 1
                    rTruCacKhoan.Y += 1
                    rTruCacKhoan.Width = rTruCacKhoan.Width + w1 + w2 + w3 + w4 + w5 - 1
                    rTruCacKhoan.Height = rTruCacKhoan.Height / _ClassNumber - 2 ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTruCacKhoan)
                    e.Graphics.DrawRectangle(Pens.Silver, rTruCacKhoan)
                    e.Graphics.DrawString("Các khoản phải trừ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                                         New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTruCacKhoan, format)
                End If
                If j = 12 Then       'Bảo hiểm
                    rTrongDo = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    rTrongDo.X += 1
                    rTrongDo.Y += rTruCacKhoan.Height + 1

                    rTrongDo.Width = rTrongDo.Width + w1 + w2 - 1
                    rTrongDo.Height = rTrongDo.Height / _ClassNumber - 2         '3 la vi 3 lop

                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTrongDo)
                    e.Graphics.DrawRectangle(Pens.Silver, rTrongDo)
                    e.Graphics.DrawString("Bảo hiểm", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTrongDo, format)
                End If
               

            Next
        End If
    End Sub

    Private Sub ckb_ChoiceAll_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Space) Then
            HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
        End If
    End Sub

    Private Sub ckb_ChoiceAll_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
    End Sub

    Private Sub dgv_main_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellValueChanged
        'Sự kiện này thực hiện khi click các checked items trên lưới dữ liệu thì sẽ Checked hoặc UnChecked Checkall
        If CType(sender, DataGridView).Columns(e.ColumnIndex).Name = "cln_Choice" And e.RowIndex >= 0 Then
            If Not IsHeaderCheckBoxClicked Then
                Dim _vCellCheck As DataGridViewCheckBoxCell
                _vCellCheck = CType(dgv_main("cln_Choice", e.RowIndex), DataGridViewCheckBoxCell)
                RowCheckBoxClick(_vCellCheck)
            End If
        End If

        Try
            If dgv_main.Rows.Count > 1 Then
                Dim colName As String = dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name.Trim()
                If ARL_FieldInteger.Contains(colName.Replace("cln_", "")) Or ARL_FieldDecimal.Contains(colName.Replace("cln_", "")) Then
                    Dim _ValueCell As String = dgv_main.Rows(e.RowIndex).Cells(colName).Value.ToString().Trim().Replace(" ", "").Replace(",", "")
                    If Not String.IsNullOrEmpty(_ValueCell) Then
                        dgv_main.Rows(e.RowIndex).Cells(colName).Value = formatMoney(_ValueCell)
                    End If
                End If
            End If
        Catch ex As Exception
            Globals.Logger.Error("Lỗi format dữ liệu (ChiLuongForm\dgv_main_CellValueChanged): " + ex.Message)
        End Try
    End Sub

    Private Sub dgv_main_CurrentCellDirtyStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_IdCanBo").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_IdCanBo").Value.ToString() <> "")) Then
                Dim vCellCheck As DataGridViewCheckBoxCell
                'Checking whether the Datagridview Checkbox column is the first column
                If dgv_main.CurrentCellAddress.X = 1 Then
                    vCellCheck = dgv_main.CurrentRow.Cells("cln_Choice")
                    If (dgv_main.IsCurrentCellDirty) Then 'Checking for dirty cell
                        dgv_main.CommitEdit(DataGridViewDataErrorContexts.Commit) 'If it is dirty, making them to commit
                    End If
                End If
            End If
        End If
    End Sub

    ''' <summary>
    '''  Hàm thực hiện căn chỉnh cho CheckBox nằm giữa cột tiêu đề của lưới dữ liệu
    ''' </summary>
    ''' <param name="_ColumnIndex"></param>
    ''' <param name="_RowIndex"></param>
    ''' <remarks></remarks>
    Private Sub ResetHeaderCheckBoxLocation(ByVal _ColumnIndex As Integer, ByVal _RowIndex As Integer)
        'Get the column header cell bounds
        Dim oRectangle As Rectangle = dgv_main.GetCellDisplayRectangle(_ColumnIndex, _RowIndex, True)
        Dim oPoint As Point = New Point()
        oPoint.X = oRectangle.Location.X + (oRectangle.Width - ckb_ChoiceAll.Width) / 2 + 1
        oPoint.Y = oRectangle.Location.Y + (oRectangle.Height - ckb_ChoiceAll.Height) / 2 + 1
        'Change the location of the CheckBox to make it stay on the header
        ckb_ChoiceAll.Location = oPoint
    End Sub

    ''' <summary>
    '''  Hàm thực hiện các thao tác khi Checkbox (Chọn cả) được Check click
    ''' </summary>
    ''' <param name="ckb_CheckAll"></param>
    ''' <param name="dgv_name"></param>
    ''' <remarks></remarks>
    Private Sub HeaderCheckBoxClick(ByVal ckb_CheckAll As CheckBox, ByVal dgv_name As DataGridView)
        If (dgv_name.Rows.Count <= 0) Then
            ckb_CheckAll.Checked = False
            Return
        End If
        IsHeaderCheckBoxClicked = True
        For i As Integer = 0 To dgv_name.Rows.Count - 1
            dgv_name.Rows(i).Cells("cln_Choice").Value = ckb_CheckAll.Checked
        Next
        dgv_name.RefreshEdit()
        TotalCheckedCheckBoxes = IIf(ckb_CheckAll.Checked, TotalCheckBoxes, 0)
        IsHeaderCheckBoxClicked = False
    End Sub

    ''' <summary>
    ''' Hàm thực hiện các công việc khi item trên lưới dữ liệu được Checked
    ''' </summary>
    ''' <param name="ckb_CellCheck"></param>
    ''' <remarks></remarks>
    Private Sub RowCheckBoxClick(ByVal ckb_CellCheck As DataGridViewCheckBoxCell)
        If Not (ckb_CellCheck Is Nothing) Then
            'Modifiy Counter
            If ((CType(ckb_CellCheck.Value, Boolean) = True) And (TotalCheckedCheckBoxes < TotalCheckBoxes)) Then
                TotalCheckedCheckBoxes += 1
            ElseIf (TotalCheckedCheckBoxes > 0) Then
                TotalCheckedCheckBoxes -= 1
            End If
            'Change state of the header CheckBox
            If (TotalCheckedCheckBoxes < TotalCheckBoxes) Then
                ckb_ChoiceAll.Checked = False
            ElseIf (TotalCheckedCheckBoxes = TotalCheckBoxes) Then
                ckb_ChoiceAll.Checked = True
            End If
        End If
    End Sub
#End Region

#Region "---> Các sự kiện Ngoại lệ bắt khi cho phép nhập dữ liệu trên lưới <---"
    Private Sub TextBox_KeyPress_Chars(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        e.Handled = True
    End Sub

    Private Sub TextBox_KeyPress_Integer(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)

        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If
        If (e.KeyChar = "-" And CType(sender, TextBox).SelectionStart = 0) Then
            KeyAscii = 1
        End If
        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If

    End Sub

    Private Sub TextBox_KeyPress_Decimal(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Or ((e.KeyChar = "-" And CType(sender, TextBox).SelectionStart = 0)) Then
            KeyAscii = 0
        End If
        If (e.KeyChar = "-" And CType(sender, TextBox).SelectionStart = 0) Then
            KeyAscii = 1
        End If
        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub dgv_main_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgv_main.EditingControlShowing
        Dim _ColName As String = dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name.Trim()
        Try
            RemoveHandler e.Control.KeyPress, AddressOf TextBox_KeyPress_Decimal
            RemoveHandler e.Control.KeyPress, AddressOf TextBox_KeyPress_Integer
            If ARL_FieldDecimal.Contains(_ColName.Replace("cln_", "")) Then
                AddHandler e.Control.KeyPress, AddressOf TextBox_KeyPress_Decimal
            ElseIf ARL_FieldInteger.Contains(_ColName.Replace("cln_", "")) Then
                AddHandler e.Control.KeyPress, AddressOf TextBox_KeyPress_Integer
            End If
        Catch ex As Exception
            Globals.Logger.Error("Lỗi bắt cờ nhập dữ liệu (ChiLuongForm\dgv_main_EditingControlShowing): " + ex.Message)
        End Try
    End Sub

    Private Function GetValueCell(ByVal pColumnName As String, pRowIndex As Integer) As Double
        Dim _Result As Double = 0, sValTMP As String = ""
        Try
            If Not IsNothing(dgv_main.Rows(pRowIndex).Cells(pColumnName).Value) Then
                If String.IsNullOrEmpty(dgv_main.Rows(pRowIndex).Cells(pColumnName).Value.ToString().Replace(" ", "").Replace(",", "")) Then
                    sValTMP = "0"
                Else
                    sValTMP = dgv_main.Rows(pRowIndex).Cells(pColumnName).Value.ToString().Replace(" ", "").Replace(",", "")
                End If
                _Result = CType(sValTMP, Double)
            End If
            Return _Result
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Sub dgv_main_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_main.CellEndEdit
        Dim rowIndex As Integer = e.RowIndex
        Dim columnIndex As Integer = e.ColumnIndex
        Dim _ColName As String = dgv_main.Columns(columnIndex).Name.Trim()
        If ARL_FieldDecimal.Contains(_ColName.Replace("cln_", "")) Or ARL_FieldInteger.Contains(_ColName.Replace("cln_", "")) Then
            Dim validation As Boolean = True
            If (Not IsNothing(dgv_main.Rows(rowIndex).Cells(columnIndex).Value)) Then
                If (Not IsNothing(dgv_main.Rows(rowIndex).Cells(columnIndex).Value) And dgv_main.Rows(rowIndex).Cells(columnIndex).Value.ToString().Trim() <> "") Then
                    Dim DataToValidate As String = dgv_main.Rows(rowIndex).Cells(columnIndex).Value.ToString()
                    If Not Globals.IsNumber(DataToValidate.ToString().Replace(" ", "").Replace(",", "")) Then
                        validation = False
                    End If
                    If (validation = False) Then
                        dgv_main.Rows(rowIndex).Cells(columnIndex).ErrorText = "Phải nhập dữ liệu kiểu số"
                    Else
                        dgv_main.Rows(rowIndex).Cells(columnIndex).ErrorText = ""
                    End If
                End If
            End If
        End If

        If IsCalculate And dgv_main.Rows.Count > 1 And (Not dgv_main.Rows(rowIndex).Cells("cln_ChiNhanh_Cd").Value Is Nothing) And (ARL_FieldDecimal.Contains(_ColName.Replace("cln_", "")) Or ARL_FieldInteger.Contains(_ColName.Replace("cln_", ""))) Then
            If Not (dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString() Is Nothing) Then
                CallCalculateCelInGrid(rowIndex, _ColName)
            End If

        End If
    End Sub

    Private Sub CallCalculateCelInGrid(ByVal rowIndex As Integer, ByVal _ColName As String)
        If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() <> "1" Then
            Dim sValueCurrCell As String = "0"
            If Not IsNothing(dgv_main.Rows(rowIndex).Cells(_ColName).Value) Then
                sValueCurrCell = dgv_main.Rows(rowIndex).Cells(_ColName).Value.ToString().Replace(" ", "").Replace(",", "")
            End If
            If String.IsNullOrEmpty(sValueCurrCell) Then
                sValueCurrCell = "0"
            End If
            Dim sChiNhanh As String = dgv_main.Rows(rowIndex).Cells("cln_ChiNhanh_Cd").Value.ToString()
            Dim sDonVi As String = dgv_main.Rows(rowIndex).Cells("cln_DonVi_Cd").Value.ToString()
            Dim sPhongBan As String = "" 'dgv_main.Rows(rowIndex).Cells("cln_PhongBan_Cd").Value.ToString()         CHUDV
            Dim _PhBanCode As String = ""
            If (cb_phongban.Items.Count <> 0 And ARL_PhongBan IsNot Nothing) Then
                _PhBanCode = IIf(cb_phongban.Items.Count <> 0, ARL_PhongBan(cb_phongban.SelectedIndex), "")
            End If
            If (Not String.IsNullOrEmpty(_PhBanCode.Trim())) Then
                sPhongBan = dgv_main.Rows(rowIndex).Cells("cln_PhongBan_Cd").Value.ToString()
            End If
            Dim dBHXH_CB As Double = GetValueCell("cln_BHXH_CB", rowIndex)
            Dim dBHXH_DV As Double = GetValueCell("cln_BHXH_DV", rowIndex)
            Dim dBHYT_CB As Double = GetValueCell("cln_BHYT_CB", rowIndex)
            Dim dBHYT_DV As Double = GetValueCell("cln_BHYT_DV", rowIndex)
            Dim dBHTN_CB As Double = GetValueCell("cln_BHTN_CB", rowIndex)
            Dim dBHTN_DV As Double = GetValueCell("cln_BHTN_DV", rowIndex)

            Dim dBHXH_CB_SoTien As Double = GetValueCell("cln_BHXH_CB_SoTien", rowIndex)
            Dim dBHYT_CB_SoTien As Double = GetValueCell("cln_BHYT_CB_SoTien", rowIndex)
            Dim dBHTN_CB_SoTien As Double = GetValueCell("cln_BHTN_CB_SoTien", rowIndex)
            Dim dBHXH_DV_SoTien As Double = GetValueCell("cln_BHXH_DV_SoTien", rowIndex)
            Dim dBHYT_DV_SoTien As Double = GetValueCell("cln_BHYT_DV_SoTien", rowIndex)
            Dim dBHTN_DV_SoTien As Double = GetValueCell("cln_BHTN_DV_SoTien", rowIndex)

            Dim dTru_Khoan_Khac As Double = GetValueCell("cln_Tru_Khoan_Khac", rowIndex)
            Dim dLCB_HeSo As Double = GetValueCell("cln_LCB_HeSo", rowIndex)
            Dim dPhCap_ChVu_HeSo As Double = GetValueCell("cln_PhCap_ChVu_HeSo", rowIndex)
            Dim dPhCap_TrNhiem_HeSo As Double = GetValueCell("cln_PhCap_TrNhiem_HeSo", rowIndex)
            Dim dPhCap_DocHai_HeSo As Double = GetValueCell("cln_PhCap_DocHai_HeSo", rowIndex)
            Dim dPhCap_ThuHut_SoTien As Double = GetValueCell("cln_PhCap_ThuHut_SoTien", rowIndex)
            Dim dPhCap_KhuVuc_SoTien As Double = GetValueCell("cln_PhCap_KhuVuc_SoTien", rowIndex)
            Dim dPhCap_ThuHut_ST_Goc As Double = GetValueCell("cln_PhCap_ThuHut_ST_Goc", rowIndex)
            Dim dPhCap_KhuVuc_ST_Goc As Double = GetValueCell("cln_PhCap_KhuVuc_ST_Goc", rowIndex)
            Dim dLuong_TTV As Double = GetValueCell("cln_Luong_TTV", rowIndex)
            Dim dLuong_CSo As Double = GetValueCell("cln_Luong_CSo", rowIndex)
            Dim dSoNgayNghi_TruLuong As Double = GetValueCell("cln_SoNgayNghi_TruLuong", rowIndex)
            Dim dLuong_V1 As Double = GetValueCell("cln_Luong_V1", rowIndex)
            Dim dLuong_V1_TamUng_K1 As Double = GetValueCell("cln_Luong_V1_TamUng_K1", rowIndex)
            Dim dLuong_V1_TamUng_K2 As Double = GetValueCell("cln_Luong_V1_TamUng_K2", rowIndex)
            Dim dMucTamUng_V2 As Double = GetValueCell("cln_MucTamUng_V2", rowIndex)
            Dim dHeSoLuong_V2 As Double = GetValueCell("cln_HeSoLuong_V2", rowIndex)
            Dim dDPCD_CB_SoTien As Double = GetValueCell("cln_DPCD_CB_SoTien", rowIndex)
            Dim sTinh_Thue_TNCN As String = IIf(String.IsNullOrEmpty(dgv_main.Rows(rowIndex).Cells("cln_Tinh_Thue_TNCN").Value.ToString().Replace(" ", "").Replace(",", "")), "0", dgv_main.Rows(rowIndex).Cells("cln_Tinh_Thue_TNCN").Value.ToString().Replace(" ", "").Replace(",", ""))
            Dim bLoai_CB As Byte = CType(dgv_main.Rows(rowIndex).Cells("cln_Tinh_Thue_TNCN").Value.ToString(), Byte)
            Dim dSoNgayLViec_Thang As Double = GetValueCell("cln_SoNgayLViec_Thang", rowIndex)
            Dim dTNTT_Muc_BanThan As Double = GetValueCell("cln_TNTT_Muc_BanThan", rowIndex)
            Dim dTNTT_Muc_PhuThuoc As Double = GetValueCell("cln_TNTT_Muc_PhuThuoc", rowIndex)
            Dim dSoNguoi_PhuThuoc As Double = GetValueCell("cln_SoNguoi_PhuThuoc", rowIndex)
            Dim dTru_UngHo_Khac As Double = GetValueCell("cln_Tru_UngHo_Khac", rowIndex)
            Dim dMucHuong As Double = GetValueCell("cln_MucHuong", rowIndex)
            Dim dSoNgay_KhongLV As Double = GetValueCell("cln_SoNgay_KhongLV", rowIndex)

            Dim dTongHSL As Double = dLCB_HeSo + dPhCap_ChVu_HeSo + dPhCap_TrNhiem_HeSo + dPhCap_DocHai_HeSo
            Dim dTru_CacKhoan_Tong As Double = 0
            Dim dLuong_ThucLinh As Double = 0
            Dim dTC_Tru_Khoan_Khac As Double = 0, dTC_Luong_ThucLinh As Double = 0, dTC_Tru_CacKhoan_Tong As Double = 0
            Dim dTC_BHXH_CB_SoTien As Double = 0, dTC_BHXH_DV_SoTien As Double = 0, dTC_BHYT_CB_SoTien As Double = 0, dTC_BHYT_DV_SoTien As Double = 0
            Dim dTC_BHTN_CB_SoTien As Double = 0, dTC_BHTN_DV_SoTien As Double = 0, dTC_PhCap_KhuVuc_SoTien As Double = 0, dTC_PhCap_ThuHut_SoTien As Double = 0
            Dim dTC_Luong_V1 As Double = 0, dTC_Luong_V1_TamUng_K1 As Double = 0, dTC_Luong_V1_TamUng_K2 As Double = 0, dTC_SoNgay_KhongLV As Double = 0
            Dim dTC_Luong_V2_TamUng As Double = 0, dTC_Luong_V1_ThucTra As Double = 0, dTC_SoNgayNghi_TruLuong As Double = 0, dTC_DPCD_CB_SoTien As Double = 0
            Dim dLuong_V1_GocK1 As Double = 0
            'Tính toán lại các cột khác theo giá trị vừa thay đổi
            If _ValueFrame = 1 Or _ValueFrame = 2 Then
                dPhCap_KhuVuc_SoTien = Math.Round((dPhCap_KhuVuc_ST_Goc - ((dSoNgay_KhongLV + dSoNgayNghi_TruLuong) / dSoNgayLViec_Thang) * dPhCap_KhuVuc_ST_Goc) * dMucHuong / 100, 0)
                dPhCap_ThuHut_SoTien = Math.Round((dPhCap_ThuHut_ST_Goc - ((dSoNgay_KhongLV + dSoNgayNghi_TruLuong) / dSoNgayLViec_Thang) * dPhCap_ThuHut_ST_Goc) * dMucHuong / 100, 0)
                dLuong_V1 = Math.Round(((((dTongHSL * dLuong_TTV) / dSoNgayLViec_Thang) * (dSoNgayLViec_Thang - dSoNgay_KhongLV - dSoNgayNghi_TruLuong)) * dMucHuong / 100 + (dPhCap_ThuHut_SoTien + dPhCap_KhuVuc_SoTien)), 0)
                dLuong_V1_GocK1 = Math.Round(((((dTongHSL * dLuong_TTV) / dSoNgayLViec_Thang) * (dSoNgayLViec_Thang - 0 - 0)) * dMucHuong / 100 + (dPhCap_ThuHut_SoTien + dPhCap_KhuVuc_SoTien)), 0)
                dLuong_V1_TamUng_K1 = Math.Round(dLuong_V1 * 80 / 100, 0)
                dLuong_V1_TamUng_K2 = dLuong_V1 - dLuong_V1_TamUng_K1
            End If
            If (_ColName = "cln_SoNgay_KhongLV" Or _ColName = "cln_MucHuong") And (dBHXH_CB_SoTien <> 0) And (_ValueFrame <> 6) Then
                dBHXH_CB_SoTien = Math.Round(dBHXH_CB * dLuong_V1_GocK1, 0)
                dBHXH_DV_SoTien = Math.Round(dBHXH_DV * dLuong_V1_GocK1, 0)
            End If
            If (_ColName = "cln_SoNgay_KhongLV" Or _ColName = "cln_MucHuong") And (dBHYT_CB_SoTien <> 0) And (_ValueFrame <> 6) Then
                dBHYT_CB_SoTien = Math.Round(dBHYT_CB * dLuong_V1_GocK1, 0)
                dBHYT_DV_SoTien = Math.Round(dBHYT_DV * dLuong_V1_GocK1, 0)
                dBHTN_CB_SoTien = Math.Round(dBHTN_CB * dLuong_V1_GocK1, 0)
                dBHTN_DV_SoTien = Math.Round(dBHTN_DV * dLuong_V1_GocK1, 0)
            End If
            If (_ColName = "cln_SoNgay_KhongLV" Or _ColName = "cln_MucHuong") And (dBHTN_CB_SoTien <> 0) And (_ValueFrame <> 6) Then
                dBHTN_CB_SoTien = Math.Round(dBHTN_CB * dLuong_V1_GocK1, 0)
                dBHTN_DV_SoTien = Math.Round(dBHTN_DV * dLuong_V1_GocK1, 0)
            End If

            Dim dBaoHiem As Double = dBHXH_CB_SoTien + dBHYT_CB_SoTien + dBHTN_CB_SoTien
            Select Case _ValueFrame
                Case 1 ' Chi lương tháng - Kỳ 1
                    If _ColName = "cln_Tru_Khoan_Khac" Or _ColName = "cln_SoNgay_KhongLV" Or _ColName = "cln_MucHuong" Or _ColName = "cln_BHXH_CB_SoTien" Or _ColName = "cln_BHYT_CB_SoTien" Or _ColName = "cln_BHTN_CB_SoTien" Then
                        dgv_main.Rows(rowIndex).Cells(_ColName).Style.BackColor = Color.LavenderBlush
                        If (dSoNgay_KhongLV + dSoNgayNghi_TruLuong <> 0) And (bLoai_CB <> 1 And bLoai_CB <> 4) Then
                            dBHXH_CB_SoTien = 0
                            dBHXH_DV_SoTien = 0
                            dBHYT_CB_SoTien = 0
                            dBHYT_DV_SoTien = 0
                            dBHTN_CB_SoTien = 0
                            dBHTN_DV_SoTien = 0
                        End If
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_CB_SoTien").Value = formatMoney(dBHXH_CB_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_CB_SoTien").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_DV_SoTien").Value = formatMoney(dBHXH_DV_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_DV_SoTien").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_CB_SoTien").Value = formatMoney(dBHYT_CB_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_CB_SoTien").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_DV_SoTien").Value = formatMoney(dBHYT_DV_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_DV_SoTien").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_CB_SoTien").Value = formatMoney(dBHTN_CB_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_CB_SoTien").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_DV_SoTien").Value = formatMoney(dBHTN_DV_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_DV_SoTien").Style.BackColor = Color.LavenderBlush

                        If dPhCap_KhuVuc_SoTien <> 0 Then
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_KhuVuc_SoTien").Value = formatMoney(dPhCap_KhuVuc_SoTien.ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_KhuVuc_SoTien").Style.BackColor = Color.LavenderBlush
                        End If
                        If dPhCap_ThuHut_SoTien <> 0 Then
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_ThuHut_SoTien").Value = formatMoney(dPhCap_ThuHut_SoTien.ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_ThuHut_SoTien").Style.BackColor = Color.LavenderBlush
                        End If

                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1").Value = formatMoney(dLuong_V1.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K1").Value = formatMoney(dLuong_V1_TamUng_K1.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K1").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K2").Value = formatMoney(dLuong_V1_TamUng_K2.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K2").Style.BackColor = Color.LavenderBlush

                        dTru_CacKhoan_Tong = dBHXH_CB_SoTien + dBHYT_CB_SoTien + dBHTN_CB_SoTien + dTru_Khoan_Khac  ' dTru_UngHo_Khac + dDPCD_CB_SoTien + dDPCD_LD_NN_SoTien
                        dLuong_ThucLinh = dLuong_V1_TamUng_K1 - dTru_CacKhoan_Tong
                        '1 - Tổng các khoản phải thu
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTru_CacKhoan_Tong.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Style.BackColor = Color.LavenderBlush

                        '2 - Tiền lương thực lĩnh kỳ này.
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Value = formatMoney(dLuong_ThucLinh.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Style.BackColor = Color.LavenderBlush

                        '3 - Tính lại Tổng cộng các cột
                        dTC_Tru_Khoan_Khac = GetTotalValue_ForPosCode(dgv_main, "Tru_Khoan_Khac", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_ThucLinh = GetTotalValue_ForPosCode(dgv_main, "Luong_ThucLinh", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_CacKhoan_Tong = GetTotalValue_ForPosCode(dgv_main, "Tru_CacKhoan_Tong", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHXH_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHXH_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHXH_DV_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHXH_DV_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHYT_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHYT_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHYT_DV_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHYT_DV_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHTN_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHTN_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHTN_DV_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHTN_DV_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_PhCap_KhuVuc_SoTien = GetTotalValue_ForPosCode(dgv_main, "PhCap_KhuVuc_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_PhCap_ThuHut_SoTien = GetTotalValue_ForPosCode(dgv_main, "PhCap_ThuHut_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1_TamUng_K1 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_TamUng_K1", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1_TamUng_K2 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_TamUng_K2", sChiNhanh, sDonVi, sPhongBan)
                        dTC_SoNgay_KhongLV = GetTotalValue_ForPosCode(dgv_main, "SoNgay_KhongLV", sChiNhanh, sDonVi, sPhongBan)
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 1 And dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString() = sChiNhanh And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = sDonVi And dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString() = sPhongBan Then
                                dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value = formatMoney(dTC_Tru_Khoan_Khac.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value = formatMoney(dTC_Luong_ThucLinh.ToString())
                                dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTC_Tru_CacKhoan_Tong.ToString())
                                dgv_main.Rows(i).Cells("cln_BHXH_CB_SoTien").Value = formatMoney(dTC_BHXH_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHXH_DV_SoTien").Value = formatMoney(dTC_BHXH_DV_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHYT_CB_SoTien").Value = formatMoney(dTC_BHYT_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHYT_DV_SoTien").Value = formatMoney(dTC_BHYT_DV_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHTN_CB_SoTien").Value = formatMoney(dTC_BHTN_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHTN_DV_SoTien").Value = formatMoney(dTC_BHTN_DV_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_SoTien").Value = formatMoney(dTC_PhCap_KhuVuc_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_SoTien").Value = formatMoney(dTC_PhCap_ThuHut_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1").Value = formatMoney(dTC_Luong_V1.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K1").Value = formatMoney(dTC_Luong_V1_TamUng_K1.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K2").Value = formatMoney(dTC_Luong_V1_TamUng_K2.ToString())
                                dgv_main.Rows(i).Cells("cln_SoNgay_KhongLV").Value = formatMoney(dTC_SoNgay_KhongLV.ToString())
                                Exit For
                            End If
                        Next
                    End If

                Case 2 ' Chi lương tháng - Kỳ 2
                    If _ColName = "cln_SoNgayNghi_TruLuong" Or _ColName = "cln_SoTienNghi_TruLuong" Or _ColName = "cln_Tru_Khoan_Khac" Or _ColName = "cln_TNTT_SoTien_NopThue" Or _ColName = "TNTT_SoTien" Or _ColName = "cln_MucTamUng_V2" Or _ColName = "cln_Luong_V1_ConLai" Or _ColName = "cln_Luong_V1_ThucTra" Or _ColName = "cln_Luong_V2_TamUng" Or _ColName = "cln_HeSoLuong_V2" Or _ColName = "cln_DPCD_CB_SoTien" Or _ColName = "cln_SoNgay_KhongLV" Or _ColName = "cln_MucHuong" Then
                        dgv_main.Rows(rowIndex).Cells(_ColName).Style.BackColor = Color.LavenderBlush

                        If dPhCap_KhuVuc_SoTien <> 0 Then
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_KhuVuc_SoTien").Value = formatMoney(dPhCap_KhuVuc_SoTien.ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_KhuVuc_SoTien").Style.BackColor = Color.LavenderBlush
                        End If
                        If dPhCap_ThuHut_SoTien <> 0 Then
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_ThuHut_SoTien").Value = formatMoney(dPhCap_ThuHut_SoTien.ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_PhCap_ThuHut_SoTien").Style.BackColor = Color.LavenderBlush
                        End If
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1").Value = formatMoney(dLuong_V1.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K1").Value = formatMoney(dLuong_V1_TamUng_K1.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K1").Style.BackColor = Color.LavenderBlush
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K2").Value = formatMoney(dLuong_V1_TamUng_K2.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_TamUng_K2").Style.BackColor = Color.LavenderBlush

                        Dim dSoTienNghi_TruLuong As Double = 0, dLuong_V1_ThucTra As Double = 0, dLuong_V1_ConLai As Double = 0, dLuong_V2_TamUng As Double = 0
                        dSoTienNghi_TruLuong = Math.Round(((dTongHSL * dLuong_TTV) * (dSoNgayNghi_TruLuong / dSoNgayLViec_Thang) + dPhCap_ThuHut_ST_Goc * (dSoNgayNghi_TruLuong / dSoNgayLViec_Thang) + dPhCap_KhuVuc_ST_Goc * (dSoNgayNghi_TruLuong / dSoNgayLViec_Thang)) * dMucHuong / 100, 0)
                        If dSoTienNghi_TruLuong <> 0 Then
                            dgv_main.Rows(rowIndex).Cells("cln_SoTienNghi_TruLuong").Value = formatMoney(Math.Round(dSoTienNghi_TruLuong).ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_SoTienNghi_TruLuong").Style.BackColor = Color.LavenderBlush
                        End If
                        dLuong_V1_ConLai = dLuong_V1_TamUng_K2 - dSoTienNghi_TruLuong
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_ConLai").Value = formatMoney(Math.Round(dLuong_V1_ConLai).ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_ConLai").Style.BackColor = Color.LavenderBlush
                        dLuong_V1_ThucTra = Math.Round((((dTongHSL * dLuong_TTV) - ((dTongHSL * dLuong_TTV) / dSoNgayLViec_Thang) * dSoNgayNghi_TruLuong) * (dMucHuong / 100)), 0)
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_ThucTra").Value = formatMoney(Math.Round(dLuong_V1_ThucTra).ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_ThucTra").Style.BackColor = Color.LavenderBlush
                        Dim dDPCD_CB_SoTienTMP As Double = 0 ' ((dTongHSL * dLuong_TTV) * (dMucHuong / 100) + dPhCap_KhuVuc_ST_Goc * dMucHuong / 100 + dPhCap_ThuHut_ST_Goc * dMucHuong / 100) / 100
                        
                        dLuong_V2_TamUng = dLuong_V1_ThucTra * dHeSoLuong_V2 * (dMucTamUng_V2 / 100)
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V2_TamUng").Value = formatMoney(Math.Round(dLuong_V2_TamUng).ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V2_TamUng").Style.BackColor = Color.LavenderBlush

                        'sTinh_Thue_TNCN
                        Dim dTNTT_SoTien As Double = 0
                        Dim dTNTT_SoTien_NopThue As Double = 0
                        Dim dTC_TNTT_SoTien As Double = 0
                        Dim dTC_TNTT_SoTien_NopThue As Double = 0
                        If sTinh_Thue_TNCN = "1" Then
                            Dim dLuong_V1_Goc As Double = ((dTongHSL * dLuong_TTV) + dPhCap_ThuHut_ST_Goc + dPhCap_KhuVuc_ST_Goc) * dMucHuong / 100
                            dBHXH_CB_SoTien = Math.Round(dBHXH_CB * dLuong_V1_Goc, 0)
                            dBHYT_CB_SoTien = Math.Round(dBHYT_CB * dLuong_V1_Goc, 0)
                            dBHTN_CB_SoTien = Math.Round(dBHTN_CB * dLuong_V1_Goc, 0)
                            dTNTT_SoTien = (dLuong_V1_ThucTra + dLuong_V2_TamUng) - (dBHXH_CB_SoTien + dBHYT_CB_SoTien + dBHTN_CB_SoTien) - dTNTT_Muc_BanThan - dTNTT_Muc_PhuThuoc * dSoNguoi_PhuThuoc
                            dTNTT_SoTien_NopThue = _ChiLuongBLL.GetTinhThue_TNCN(dTNTT_SoTien, GetDate_ForDieuKienCL())
                            dgv_main.Rows(rowIndex).Cells("cln_TNTT_SoTien").Value = formatMoney(Math.Round(dTNTT_SoTien).ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_TNTT_SoTien").Style.BackColor = Color.LavenderBlush
                            dgv_main.Rows(rowIndex).Cells("cln_TNTT_SoTien_NopThue").Value = formatMoney(Math.Round(dTNTT_SoTien_NopThue).ToString())
                            dgv_main.Rows(rowIndex).Cells("cln_TNTT_SoTien_NopThue").Style.BackColor = Color.LavenderBlush
                            dTC_TNTT_SoTien = GetTotalValue_ForPosCode(dgv_main, "TNTT_SoTien", sChiNhanh, sDonVi, sPhongBan)
                            dTC_TNTT_SoTien_NopThue = GetTotalValue_ForPosCode(dgv_main, "TNTT_SoTien_NopThue", sChiNhanh, sDonVi, sPhongBan)
                        End If
                        If (sDonVi = "000196") Then
                            dDPCD_CB_SoTien = dLuong_CSo * 10 / 100
                        Else
                            dDPCD_CB_SoTienTMP = Math.Round((dLuong_V1 + dLuong_V2_TamUng - dTNTT_SoTien_NopThue - dBHXH_CB_SoTien - dBHYT_CB_SoTien - dBHTN_CB_SoTien) * 1 / 100)
                            If dDPCD_CB_SoTienTMP <= dLuong_CSo * 10 / 100 Then
                                dDPCD_CB_SoTien = dDPCD_CB_SoTienTMP
                            Else : dDPCD_CB_SoTien = dLuong_CSo * 10 / 100
                            End If
                        End If

                        dgv_main.Rows(rowIndex).Cells("cln_DPCD_CB_SoTien").Value = formatMoney(Math.Round(dDPCD_CB_SoTien).ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_DPCD_CB_SoTien").Style.BackColor = Color.LavenderBlush

                        dTru_CacKhoan_Tong = dTNTT_SoTien_NopThue + dDPCD_CB_SoTien + dTru_Khoan_Khac + dTru_UngHo_Khac
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(Math.Round(dTru_CacKhoan_Tong).ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Style.BackColor = Color.LavenderBlush

                        dLuong_ThucLinh = dLuong_V1_ConLai + dLuong_V2_TamUng - dTru_CacKhoan_Tong
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Value = formatMoney(Math.Round(dLuong_ThucLinh).ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Style.BackColor = Color.LavenderBlush

                        dTC_Tru_Khoan_Khac = GetTotalValue_ForPosCode(dgv_main, "Tru_Khoan_Khac", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_CacKhoan_Tong = GetTotalValue_ForPosCode(dgv_main, "Tru_CacKhoan_Tong", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_ThucLinh = GetTotalValue_ForPosCode(dgv_main, "Luong_ThucLinh", sChiNhanh, sDonVi, sPhongBan)
                        dTC_SoNgay_KhongLV = GetTotalValue_ForPosCode(dgv_main, "SoNgay_KhongLV", sChiNhanh, sDonVi, sPhongBan)
                        dTC_SoNgayNghi_TruLuong = GetTotalValue_ForPosCode(dgv_main, "SoNgayNghi_TruLuong", sChiNhanh, sDonVi, sPhongBan)
                        Dim dTC_SoTienNghi_TruLuong As Double = GetTotalValue_ForPosCode(dgv_main, "SoTienNghi_TruLuong", sChiNhanh, sDonVi, sPhongBan)
                        Dim dTC_Luong_V1_ConLai As Double = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_ConLai", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1_ThucTra = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_ThucTra", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V2_TamUng = GetTotalValue_ForPosCode(dgv_main, "Luong_V2_TamUng", sChiNhanh, sDonVi, sPhongBan)
                        dTC_DPCD_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "DPCD_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)

                        dTC_PhCap_KhuVuc_SoTien = GetTotalValue_ForPosCode(dgv_main, "PhCap_KhuVuc_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_PhCap_ThuHut_SoTien = GetTotalValue_ForPosCode(dgv_main, "PhCap_ThuHut_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1_TamUng_K1 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_TamUng_K1", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1_TamUng_K2 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_TamUng_K2", sChiNhanh, sDonVi, sPhongBan)
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 1 And dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString() = sChiNhanh And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = sDonVi And dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString() = sPhongBan Then
                                dgv_main.Rows(i).Cells("cln_SoNgayNghi_TruLuong").Value = formatMoney(dTC_SoNgayNghi_TruLuong.ToString())
                                dgv_main.Rows(i).Cells("cln_SoTienNghi_TruLuong").Value = formatMoney(dTC_SoTienNghi_TruLuong.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_ConLai").Value = formatMoney(dTC_Luong_V1_ConLai.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_ThucTra").Value = formatMoney(dTC_Luong_V1_ThucTra.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V2_TamUng").Value = formatMoney(dTC_Luong_V2_TamUng.ToString())
                                If sTinh_Thue_TNCN = "1" Then
                                    dgv_main.Rows(i).Cells("cln_TNTT_SoTien_NopThue").Value = formatMoney(dTC_TNTT_SoTien_NopThue.ToString())
                                    dgv_main.Rows(i).Cells("cln_TNTT_SoTien").Value = formatMoney(dTC_TNTT_SoTien.ToString())
                                End If

                                dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value = formatMoney(dTC_Tru_Khoan_Khac.ToString())
                                dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTC_Tru_CacKhoan_Tong.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value = formatMoney(dTC_Luong_ThucLinh.ToString())
                                dgv_main.Rows(i).Cells("cln_DPCD_CB_SoTien").Value = formatMoney(dTC_DPCD_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_SoNgay_KhongLV").Value = formatMoney(dTC_SoNgay_KhongLV.ToString())

                                dgv_main.Rows(i).Cells("cln_PhCap_KhuVuc_SoTien").Value = formatMoney(dTC_PhCap_KhuVuc_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_PhCap_ThuHut_SoTien").Value = formatMoney(dTC_PhCap_ThuHut_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1").Value = formatMoney(dTC_Luong_V1.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K1").Value = formatMoney(dTC_Luong_V1_TamUng_K1.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_TamUng_K2").Value = formatMoney(dTC_Luong_V1_TamUng_K2.ToString())
                                Exit For
                            End If
                        Next
                    End If

                Case 3 ' Chi lương Cho Lao động Ngắn hạn
                    If (_ColName = "cln_Tru_Khoan_Khac" Or _ColName = "cln_LamDem_SoGio") Then
                        dgv_main.Rows(rowIndex).Cells(_ColName).Style.BackColor = Color.LavenderBlush
                        'Tính toán lại các cột khác theo giá trị vừa thay đổi
                        Dim dDonGia_LamDem_Gio As Double = GetValueCell("cln_DonGia_LamDem_Gio", rowIndex)
                        Dim dLamDem_SoGio As Double = GetValueCell("cln_LamDem_SoGio", rowIndex)
                        Dim dLamDem_SoTienPC As Double = GetValueCell("cln_LamDem_SoTienPC", rowIndex)

                        '1 - Tổng các khoản phải thu
                        dTru_CacKhoan_Tong = dBHXH_CB_SoTien + dBHYT_CB_SoTien + dBHTN_CB_SoTien + dTru_Khoan_Khac + dDPCD_CB_SoTien    ' CType(sTru_UngHo_Khac, Double)+ CType(sDPCD_LD_NN_SoTien, Double) + 
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTru_CacKhoan_Tong.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Style.BackColor = Color.LavenderBlush

                        'Tiền phụ cấp làm đêm
                        dLamDem_SoTienPC = Math.Round(dLamDem_SoGio * dDonGia_LamDem_Gio, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_LamDem_SoTienPC").Value = formatMoney(dLamDem_SoTienPC.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_LamDem_SoTienPC").Style.BackColor = Color.LavenderBlush

                        '2 - Tiền lương thực lĩnh kỳ này.
                        dLuong_ThucLinh = dLuong_V1 + dLamDem_SoTienPC - dTru_CacKhoan_Tong
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Value = formatMoney(dLuong_ThucLinh.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Style.BackColor = Color.LavenderBlush
                        '3 - Tính lại Tổng cộng các cột
                        Dim dTC_LamDem_SoGio As Double = GetTotalValue_ForPosCode(dgv_main, "LamDem_SoGio", sChiNhanh, sDonVi, sPhongBan)
                        Dim dTC_LamDem_SoTienPC As Double = GetTotalValue_ForPosCode(dgv_main, "LamDem_SoTienPC", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_Khoan_Khac = GetTotalValue_ForPosCode(dgv_main, "Tru_Khoan_Khac", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_CacKhoan_Tong = GetTotalValue_ForPosCode(dgv_main, "Tru_CacKhoan_Tong", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_ThucLinh = GetTotalValue_ForPosCode(dgv_main, "Luong_ThucLinh", sChiNhanh, sDonVi, sPhongBan)
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 1 And dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString() = sChiNhanh And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = sDonVi And dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString() = sPhongBan Then
                                dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value = formatMoney(dTC_Tru_Khoan_Khac.ToString())
                                dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTC_Tru_CacKhoan_Tong.ToString())
                                dgv_main.Rows(i).Cells("cln_LamDem_SoGio").Value = formatMoney(dTC_LamDem_SoGio.ToString())
                                dgv_main.Rows(i).Cells("cln_LamDem_SoTienPC").Value = formatMoney(dTC_LamDem_SoTienPC.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value = formatMoney(dTC_Luong_ThucLinh.ToString())
                                Exit For
                            End If
                        Next
                    End If
                Case 5 ' Chi lương Lao động tập Nghề (Trọn gói)
                    If (_ColName = "cln_Tru_Khoan_Khac") Then
                        dgv_main.Rows(rowIndex).Cells(_ColName).Style.BackColor = Color.LavenderBlush
                        '1 - Tổng các khoản phải thu
                        dTru_CacKhoan_Tong = dTru_Khoan_Khac
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTru_CacKhoan_Tong.ToString())

                        '2 - Tiền lương thực lĩnh kỳ này.
                        dLuong_ThucLinh = dLuong_V1 - dTru_CacKhoan_Tong
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Value = formatMoney(dLuong_ThucLinh.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Style.BackColor = Color.LavenderBlush

                        '3 - Tính lại Tổng cộng các cột
                        dTC_Tru_Khoan_Khac = GetTotalValue_ForPosCode(dgv_main, "Tru_Khoan_Khac", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_CacKhoan_Tong = GetTotalValue_ForPosCode(dgv_main, "Tru_CacKhoan_Tong", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_ThucLinh = GetTotalValue_ForPosCode(dgv_main, "Luong_ThucLinh", sChiNhanh, sDonVi, sPhongBan)
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 1 And dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString() = sChiNhanh And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = sDonVi And dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString() = sPhongBan Then
                                dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value = formatMoney(dTC_Tru_Khoan_Khac.ToString())
                                dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTC_Tru_CacKhoan_Tong.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value = formatMoney(dTC_Luong_ThucLinh.ToString())
                                Exit For
                            End If
                        Next
                    End If
                Case 4 ' Chi lương Thưởng
                    If (_ColName = "cln_Luong_V1") And dgv_main.Rows(rowIndex).Cells("cln_Loai_CB").Value <> "1" Then
                        dgv_main.Rows(rowIndex).Cells(_ColName).Style.BackColor = Color.LavenderBlush

                        '2 - Tiền lương thực lĩnh kỳ này.
                        dLuong_ThucLinh = dLuong_V1
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Value = formatMoney(dLuong_ThucLinh.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Style.BackColor = Color.LavenderBlush

                        '3 - Tính lại Tổng cộng các cột
                        dTC_Luong_V1 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_ThucLinh = GetTotalValue_ForPosCode(dgv_main, "Luong_ThucLinh", sChiNhanh, sDonVi, sPhongBan)
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 1 And dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString() = sChiNhanh And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = sDonVi And dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString() = sPhongBan Then
                                dgv_main.Rows(i).Cells("cln_Luong_V1").Value = formatMoney(dTC_Luong_V1.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value = formatMoney(dTC_Luong_ThucLinh.ToString())
                                Exit For
                            End If
                        Next
                    End If
                Case 6 ' Truy lĩnh lương
                    If (_ColName = "cln_SoNgayNghi_TruLuong" Or _ColName = "cln_Tru_Khoan_Khac") Then
                        dgv_main.Rows(rowIndex).Cells(_ColName).Style.BackColor = Color.LavenderBlush
                        dLuong_V1 = Math.Round((((dPhCap_KhuVuc_ST_Goc) * dLuong_TTV) / (dSoNgayLViec_Thang)) * 3 * (dSoNgayNghi_TruLuong), 0)  'dLuong_V1 = (8) = (((6)*LTTV)/(Số ngày thực tế trong quý))*3*(7)
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1").Value = formatMoney(dLuong_V1.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1").Style.BackColor = Color.LavenderBlush

                        Dim dLuong_V2_TamUng As Double = 0, dLuong_V1_ThucTra As Double = 0, dDPCD_LD_NN_SoTien As Double = 0
                        dLuong_V2_TamUng = Math.Round(dLuong_V1 * dHeSoLuong_V2 * (dMucTamUng_V2 / 100), 0)
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V2_TamUng").Value = formatMoney(dLuong_V2_TamUng.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V2_TamUng").Style.BackColor = Color.LavenderBlush

                        dBHXH_CB_SoTien = Math.Round(dBHXH_CB * dLuong_V1, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_CB_SoTien").Value = formatMoney(dBHXH_CB_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_CB_SoTien").Style.BackColor = Color.LavenderBlush

                        dBHXH_DV_SoTien = Math.Round(dBHXH_DV * dLuong_V1, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_DV_SoTien").Value = formatMoney(dBHXH_DV_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHXH_DV_SoTien").Style.BackColor = Color.LavenderBlush

                        dBHYT_CB_SoTien = Math.Round(dBHYT_CB * dLuong_V1, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_CB_SoTien").Value = formatMoney(dBHYT_CB_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_CB_SoTien").Style.BackColor = Color.LavenderBlush

                        dBHYT_DV_SoTien = Math.Round(dBHYT_DV * dLuong_V1, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_DV_SoTien").Value = formatMoney(dBHYT_DV_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHYT_DV_SoTien").Style.BackColor = Color.LavenderBlush

                        dBHTN_CB_SoTien = Math.Round(dBHTN_CB * dLuong_V1, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_CB_SoTien").Value = formatMoney(dBHTN_CB_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_CB_SoTien").Style.BackColor = Color.LavenderBlush

                        dBHTN_DV_SoTien = Math.Round(dBHTN_DV * dLuong_V1, 0)
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_DV_SoTien").Value = formatMoney(dBHTN_DV_SoTien.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_BHTN_DV_SoTien").Style.BackColor = Color.LavenderBlush

                        dBaoHiem = dBHXH_CB_SoTien + dBHYT_CB_SoTien + dBHTN_CB_SoTien
                        dLuong_V1_ThucTra = dLuong_V1 + dLuong_V2_TamUng
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_ThucTra").Value = formatMoney(dLuong_V1_ThucTra.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_V1_ThucTra").Style.BackColor = Color.LavenderBlush

                        dTru_CacKhoan_Tong = dBHXH_CB_SoTien + dBHYT_CB_SoTien + dBHTN_CB_SoTien + dDPCD_CB_SoTien + dDPCD_LD_NN_SoTien + dTru_Khoan_Khac + dTru_UngHo_Khac
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTru_CacKhoan_Tong.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Tru_CacKhoan_Tong").Style.BackColor = Color.LavenderBlush

                        dLuong_ThucLinh = (dLuong_V1 + dLuong_V2_TamUng) - dTru_CacKhoan_Tong
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Value = formatMoney(dLuong_ThucLinh.ToString())
                        dgv_main.Rows(rowIndex).Cells("cln_Luong_ThucLinh").Style.BackColor = Color.LavenderBlush

                        dgv_main.Rows(rowIndex).Cells("cln_GhiChu_ChiLuong").Value = "NL " + dgv_main.Rows(rowIndex).Cells("cln_Loai_HDNH_DB").Value.ToString().Trim() + " -> " + dgv_main.Rows(rowIndex).Cells("cln_LCB_BacLuong_HT").Value.ToString().Trim() + " từ ngày " + dgv_main.Rows(rowIndex).Cells("cln_LCB_NgHuong").Value.ToString().Trim() + ". Được hưởng: " + dgv_main.Rows(rowIndex).Cells("cln_SoNgayNghi_TruLuong").Value.ToString().Trim() + "/" + dgv_main.Rows(rowIndex).Cells("cln_SoNgayLViec_Thang").Value.ToString().Trim() + " ngày trong quý"

                        '3 - Tính lại Tổng cộng các cột
                        dTC_SoNgayNghi_TruLuong = GetTotalValue_ForPosCode(dgv_main, "SoNgayNghi_TruLuong", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1 = GetTotalValue_ForPosCode(dgv_main, "Luong_V1", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V2_TamUng = GetTotalValue_ForPosCode(dgv_main, "Luong_V2_TamUng", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHXH_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHXH_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHXH_DV_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHXH_DV_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHYT_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHYT_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHYT_DV_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHYT_DV_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHTN_CB_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHTN_CB_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_BHTN_DV_SoTien = GetTotalValue_ForPosCode(dgv_main, "BHTN_DV_SoTien", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_V1_ThucTra = GetTotalValue_ForPosCode(dgv_main, "Luong_V1_ThucTra", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_Khoan_Khac = GetTotalValue_ForPosCode(dgv_main, "Tru_Khoan_Khac", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Tru_CacKhoan_Tong = GetTotalValue_ForPosCode(dgv_main, "Tru_CacKhoan_Tong", sChiNhanh, sDonVi, sPhongBan)
                        dTC_Luong_ThucLinh = GetTotalValue_ForPosCode(dgv_main, "Luong_ThucLinh", sChiNhanh, sDonVi, sPhongBan)
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If CType(dgv_main.Rows(i).Cells("cln_KieuIn").Value, Byte) = 1 And dgv_main.Rows(i).Cells("cln_ChiNhanh_Cd").Value.ToString() = sChiNhanh And dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value.ToString() = sDonVi And dgv_main.Rows(i).Cells("cln_PhongBan_Cd").Value.ToString() = sPhongBan Then
                                dgv_main.Rows(i).Cells("cln_SoNgayNghi_TruLuong").Value = formatMoney(dTC_SoNgayNghi_TruLuong.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1").Value = formatMoney(dTC_Luong_V1.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V2_TamUng").Value = formatMoney(dTC_Luong_V2_TamUng.ToString())
                                dgv_main.Rows(i).Cells("cln_BHXH_CB_SoTien").Value = formatMoney(dTC_BHXH_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHXH_DV_SoTien").Value = formatMoney(dTC_BHXH_DV_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHYT_CB_SoTien").Value = formatMoney(dTC_BHYT_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHYT_DV_SoTien").Value = formatMoney(dTC_BHYT_DV_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHTN_CB_SoTien").Value = formatMoney(dTC_BHTN_CB_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_BHTN_DV_SoTien").Value = formatMoney(dTC_BHTN_DV_SoTien.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_V1_ThucTra").Value = formatMoney(dTC_Luong_V1_ThucTra.ToString())
                                dgv_main.Rows(i).Cells("cln_Tru_Khoan_Khac").Value = formatMoney(dTC_Tru_Khoan_Khac.ToString())
                                dgv_main.Rows(i).Cells("cln_Tru_CacKhoan_Tong").Value = formatMoney(dTC_Tru_CacKhoan_Tong.ToString())
                                dgv_main.Rows(i).Cells("cln_Luong_ThucLinh").Value = formatMoney(dTC_Luong_ThucLinh.ToString())
                                Exit For
                            End If
                        Next
                    End If
            End Select
        End If

    End Sub

    Private Sub dgv_main_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv_main.KeyDown
        dgv_main_CellContentClick(sender, Nothing)
    End Sub

    Private Sub dgv_main_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgv_main.CellBeginEdit
        'dgv_main_CellContentClick(sender, Nothing)
        If dgv_main.Rows.Count > 1 Then
            Dim colName As String = dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name.Trim()
            'If (dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name = "cln_Tru_Khoan_Khac") Then
            If ARL_FieldInteger.Contains(colName.Replace("cln_", "")) Or ARL_FieldDecimal.Contains(colName.Replace("cln_", "")) And colName <> "cln_LCB_HeSo" Then
                If Not (dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString() Is Nothing) Then
                    If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() = "1" Then
                        dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = True
                        Return
                    Else
                        dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = False
                    End If
                Else
                    dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = False
                End If
            End If
        End If
    End Sub

    Private Sub dgv_main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_main.CellContentClick
        If dgv_main.Rows.Count > 1 Then
            Dim colName As String = dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name.Trim()
            If _ValueFrame = 4 Then
                If colName = "cln_Luong_V1" Then
                    If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() = "1" Or dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Loai_CB").Value.ToString().Trim() = "1" Then
                        dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = True
                        Return
                    Else
                        dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = False
                    End If
                End If
            Else
                'If (dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name = "cln_Tru_Khoan_Khac") Then
                If ARL_FieldInputs.Contains(colName.Replace("cln_", "")) And colName <> "cln_LCB_HeSo" Then 'Or ARL_FieldDecimal.Contains(colName.Replace("cln_", "")) 
                    If Not (dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString() Is Nothing) Then
                        If colName = "cln_BHXH_CB_SoTien" Or colName = "cln_BHYT_CB_SoTien" Or colName = "cln_BHTN_CB_SoTien" Or colName = "cln_SoNgay_KhongLV" Then
                            Dim dSoNgayKhongLV As Double = GetValueCell("cln_SoNgay_KhongLV", dgv_main.CurrentCell.RowIndex)
                            If dSoNgayKhongLV <> 0 And dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() <> "1" Then
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_BHXH_CB_SoTien").ReadOnly = False
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_BHYT_CB_SoTien").ReadOnly = False
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_BHTN_CB_SoTien").ReadOnly = False
                            Else
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_BHXH_CB_SoTien").ReadOnly = True
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_BHYT_CB_SoTien").ReadOnly = True
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_BHTN_CB_SoTien").ReadOnly = True
                            End If
                        Else
                            If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() = "1" Then
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = True
                                Return
                            Else
                                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells(colName).ReadOnly = False
                            End If
                        End If
                    End If
                End If
            End If

           
        End If
    End Sub

    Private Sub dgv_main_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgv_main.CellFormatting
        If dgv_main.Rows.Count > 1 Then
            If _ValueFrame = 1 Or _ValueFrame = 2 Then
                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_SoNgay_KhongLV").ToolTipText = "Lưu ý: Cột [Số ngày không làm việc trong tháng tại đơn vị (Không lương, PhC các loại)] => Chỉ cập nhật cho những trường hợp cán bộ làm việc tại đơn vị chi lương không đủ số ngày công thực tế chi lương (Ví dụ: Vào NHCSXH từ ngày 10/10/2020; Đến ngày 12/10/2020 được điều động/tăng cường đến đơn vị khác)!"
                dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_MucHuong").ToolTipText = "Tỷ lệ hưởng lương theo HĐLĐ => Mặc định là 100%. Trường hợp Loại QĐ là [Tuyển dụng] và Loại HĐLĐ là [Thử việc] thì hưởng 85%"
                If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_LCB_BacLuong_Id").Value Is Nothing And dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Loai_CB").Value.ToString() = "1" Then
                    dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_HoTen").ToolTipText = "Hệ số lương của cán bộ chưa cập nhật"
                ElseIf String.IsNullOrEmpty(dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_LCB_BacLuong_Id").Value.ToString()) Or dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_LCB_BacLuong_Id").Value.ToString() = "0" And dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Loai_CB").Value.ToString() = "1" Then
                    dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_HoTen").ToolTipText = "Hệ số lương của cán bộ chưa cập nhật"
                ElseIf dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Luong_TTV").Value Is Nothing Then
                    dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_HoTen").ToolTipText = "Chưa cập nhật tham số Lương tối thiểu vùng cho đơn vị"
                ElseIf String.IsNullOrEmpty(dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Luong_TTV").Value.ToString()) Or dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Luong_TTV").Value.ToString() = "0" Then
                    dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_HoTen").ToolTipText = "Chưa cập nhật tham số Lương tối thiểu vùng cho đơn vị"
                End If
            ElseIf _ValueFrame = 3 Then
                If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_Luong_V1").Value Is Nothing Then
                    dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_HoTen").ToolTipText = "Chưa nhập số tiền công/tháng trong HĐLĐ Ngắn hạn"
                End If
            End If
            Dim colName As String = dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name.Trim()
            If ARL_FieldInteger.Contains(colName.Replace("cln_", "")) Or ARL_FieldDecimal.Contains(colName.Replace("cln_", "")) Then
                If Not IsNothing(e.Value) Then
                    If Globals.IsNumber(e.Value.ToString().Replace(" ", "").Replace(",", "")) Then
                        e.Value = formatMoney(e.Value)
                    End If
                End If
            End If
        End If

    End Sub
#End Region

#Region "---> Hàm, Sự kiện xuất dữ liệu Báo cáo ra Excel <---"
    Private Sub btn_export_excel_Click(sender As Object, e As EventArgs) Handles btn_export_excel.Click
        Try
            Dim sSQL As String = ""
            Dim _NgayHeThong As DateTime = Globals.GetDateTime_ForServerDB
            Dim _DoiTuongCd As String = IIf(rb_daihan.Checked = True, "1", IIf(rb_nganhan.Checked = True, "2", "3"))
            Dim _LoaiCLCd As String = IIf(ARL_PhanLoai.Count > 0, ARL_PhanLoai(cb_loaichiluong.SelectedIndex), "")
            Dim _PosCode As String = IIf(ARL_PosCode.Count > 0, ARL_PosCode(cb_donvi.SelectedIndex), "")
            Dim _PhBanCode As String = ""
            If (cb_phongban.Items.Count <> 0 And ARL_PhongBan IsNot Nothing) Then
                _PhBanCode = IIf(cb_phongban.Items.Count <> 0, ARL_PhongBan(cb_phongban.SelectedIndex), "")
            End If
            Dim _NgayBC As String = DateTimeUtil.DateTimeToString(GetDate_ForDieuKienCL(), "dd/MM/yyyy")
            _NgayBC = IIf(String.IsNullOrEmpty(_NgayBC), _NgayBC, _NgayBC.Substring(0, 10))
            Dim _DonVi_TK_Cd As String = ""
            If (_PosCode = "" Or _PosCode = "0") Then
                _DonVi_TK_Cd = DONVI
            Else : _DonVi_TK_Cd = _PosCode
            End If

            If String.IsNullOrEmpty(_LoaiCLCd) Then
                MessageBox.Show("Bạn chưa chọn Phân loại chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_loaichiluong
            ElseIf ((_LoaiCLCd = "01") And (CType(num_kybc.Value, Integer) <> 1 And CType(num_kybc.Value, Integer) <> 2)) Then
                MessageBox.Show("Bạn chưa chọn Kỳ báo cáo cho chi lương tháng. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = num_kybc
            ElseIf (String.IsNullOrEmpty(_DoiTuongCd)) Then
                MessageBox.Show("Bạn chưa chọn lao động (Dài hạn/Ngắn hạn/Tập nghề và Tư vấn) chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = rb_daihan
            ElseIf (String.IsNullOrEmpty(_PosCode)) Then
                MessageBox.Show("Bạn chưa chọn đơn vị cần chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_donvi
            ElseIf ((_PosCode = "000100" Or _PosCode = "000199") And rb_nganhan.Checked = False And String.IsNullOrEmpty(_PhBanCode)) Then
                MessageBox.Show("Bạn chưa chọn Ban CMNV cần chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_phongban
            Else
                Dim iKyBC As Byte = CType(num_kybc.Value, Byte)
                sSQL = String.Format("Select IsNull(TongHopId,0) From ChiLuong_TongHop Where TrangThai <> 0 And DonVi_CL_Cd='{0}' And PhongBan_CL_Cd='{1}' And NamBC={2} And ThangBC={3} And KyBC={4} And PhanLoai_Cd='{5}' And LaoDong_Cd='{6}'", _PosCode, _PhBanCode, dtpk_ngaybc.Value.Year, dtpk_ngaybc.Value.Month, iKyBC, _LoaiCLCd, _DoiTuongCd)
                Dim TongHopIdEx As Long = SoftSqlHelper.GetNumberLong(sSQL, 0)
                If TongHopIdEx <= 0 Then
                    MessageBox.Show("Bạn cập nhật dữ liệu chi lương [" + sThongBaoCL + "]. Không thể xuất Excel!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_loaichiluong
                Else
                    'Bắt đầu xuất dữ liệu Chi lương ra excel
                    Dim db_chiluong As DataTable = New DataTable()
                    db_chiluong = _ChiLuongBLL.ChiLuong_BangKeCT_GetSearch(TongHopIdEx, 0, "", 0, 0, 99, "", "", 0, 9)
                    If (db_chiluong Is Nothing Or db_chiluong.Rows.Count <= 0) Then
                        MessageBox.Show("Không có dữ liệu Chi lương [" + sThongBaoCL + "] để xuất file excel. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Return
                    End If
                    Cursor = Cursors.WaitCursor
                    lbl_waitting.Text = "Waitting for data export..."
                    'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
                    Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
                    System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
                    Dim _IdDonVi As Integer = CInt(cb_donvi.SelectedValue)
                    Dim _ProvinceName As String = "", _BranchName = ""
                    If _IdDonVi = 1 Then
                        _ProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                        _BranchName = ""
                    Else
                        _ProvinceName = db_chiluong.Rows(0)("ChiNhanh_HT").ToString().ToUpper()
                        _BranchName = db_chiluong.Rows(0)("DonVi_HT").ToString().Replace("Chi nhánh NHCSXH", "Hội sở").ToUpper()
                    End If
                    Dim _ResultEx As String = _ChiLuongBLL.ExportExcel_CT_LuongThang(_DonVi_TK_Cd, _ProvinceName, _BranchName, dtpk_ngaybc.Value, _LoaiCLCd, iKyBC, _DoiTuongCd, db_chiluong)
                    MessageBox.Show("Dữ liệu được xuất ra file excel thành công theo đường dẫn: [" + _ResultEx + "]", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

                    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
                    lbl_waitting.Text = ""
                    Cursor = Cursors.Default
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi xuất báo cáo ra excel: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi xuất báo cáo ra excel (ChiLuongForm\btn_export_excel_Click): " + ex.Message)
        End Try
    End Sub
#End Region

    Private Sub edt_bosung_sotien_khac_TextChanged(sender As Object, e As EventArgs) Handles edt_bosung_sotien_khac.TextChanged
        edt_bosung_sotien_khac = _ChiLuongBLL.FormatMoneyInTextbox(edt_bosung_sotien_khac)
    End Sub

    Private Sub edt_bosung_sotien_khac_Leave(sender As Object, e As EventArgs) Handles edt_bosung_sotien_khac.Leave
        If (edt_bosung_sotien_khac.Text = "") Then
            edt_bosung_sotien_khac.Text = "0"
        End If
    End Sub

    Private Sub edt_bosung_sotien_khac_KeyDown(sender As Object, e As KeyEventArgs) Handles edt_bosung_sotien_khac.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_bosung_sotien_khac.Text = "") Then
                edt_bosung_sotien_khac.Text = "0"
            End If
        End If
    End Sub

    Private Sub btn_bosung_trichnopkhac_Click(sender As Object, e As EventArgs) Handles btn_bosung_trichnopkhac.Click
        Try
            If (dgv_main Is Nothing Or dgv_main.Rows.Count <= 0) Then
                MessageBox.Show("Bạn chưa Có danh sách chi lương. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = btn_lapbang
                Return
            End If
            If (cb_bosung_loai.SelectedIndex <= 0 And cb_bosung_loai.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn Loại phải thu khác muốn cộng thêm vào cột [Các khoản trích nộp khác]. Vui lòng chọn Loại phải thu khác!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_bosung_loai
                Return
            End If
            If (edt_bosung_sotien_khac.Text.Trim() = "" Or edt_bosung_sotien_khac.Text.Trim() = "0") Then
                If (cb_bosung_loai.SelectedIndex = 1) Then
                    MessageBox.Show("Bạn chưa nhập số tiền cụ thể phải thu khác muốn cộng thêm vào cột [Các khoản trích nộp khác]!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_bosung_sotien_khac
                    Return
                ElseIf (cb_bosung_loai.SelectedIndex = 2) Then
                    MessageBox.Show("Bạn chưa nhập số ngày lương thực lĩnh muốn cộng thêm vào cột [Các khoản trích nộp khác]!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_bosung_sotien_khac
                    Return
                ElseIf (cb_bosung_loai.SelectedIndex = 3) Then
                    MessageBox.Show("Bạn chưa nhập số ngày lương bảo hiểm muốn cộng thêm vào cột [Các khoản trích nộp khác]!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_bosung_sotien_khac
                    Return
                End If
            End If
            If (edt_bosung_ghichu_khac.Text.Trim() = "") Then
                MessageBox.Show("Bạn chưa nhập ghi chú nội dung cho Bổ sung [Các khoản trích nộp khác]!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_bosung_ghichu_khac
                Return
            End If
            If dgv_main.Rows.Count > 0 Then
                '      dgSound.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = file
                Dim dTru_Khoan_Khac_OLD As Double = 0
                Dim dBoSung_SoTien_Ngay_Khac As Double = 0
                Dim dSoTienBoSungKhac As Double = 0
                Dim sGhiChu_OLD As String = ""
                Dim dLuong_V1 As Double = 0
                Dim dLCB_Tong_HeSo As Double = 0
                Dim dLuong_TTV As Double = 0
                Dim dMucHuong As Double = 0
                Dim dPhCap_ThuHut_ST_Goc As Double = 0
                Dim dPhCap_KhuVuc_ST_Goc As Double = 0

                Dim dHeSoLuong_V2 As Double = 0
                Dim dMucTamUng_V2 As Double = 0
                Dim dLuong_V2_TamUng As Double = 0
                Dim dLuong_V1_ThucTra As Double = 0
                Dim dLuong_ThucLinh As Double = 0

                dBoSung_SoTien_Ngay_Khac = CType(edt_bosung_sotien_khac.Text.Replace(" ", "").Replace(",", ""), Double)
                Dim dSoNgayLViec_Thang As Double = 0
                For iRow As Integer = 0 To dgv_main.RowCount - 1
                    If dgv_main.Rows(iRow).Cells("cln_KieuIn").Value.ToString().Trim() <> "1" Then
                        dTru_Khoan_Khac_OLD = GetValueCell("cln_Tru_Khoan_Khac", iRow)
                        sGhiChu_OLD = dgv_main.Rows(iRow).Cells("cln_Tru_Khoan_Khac_GhiChu").Value
                        dSoTienBoSungKhac = 0

                        dLCB_Tong_HeSo = GetValueCell("cln_LCB_Tong_HeSo", iRow)
                        dLuong_TTV = GetValueCell("cln_Luong_TTV", iRow)
                        dMucHuong = GetValueCell("cln_MucHuong", iRow)
                        dPhCap_ThuHut_ST_Goc = GetValueCell("cln_PhCap_ThuHut_ST_Goc", iRow)
                        dPhCap_KhuVuc_ST_Goc = GetValueCell("cln_PhCap_KhuVuc_ST_Goc", iRow)


                        If cb_bosung_loai.SelectedIndex = 1 Then
                            dSoTienBoSungKhac = dBoSung_SoTien_Ngay_Khac + dTru_Khoan_Khac_OLD
                        ElseIf cb_bosung_loai.SelectedIndex = 2 Then    'Số ngày lương thực lĩnh
                            dSoNgayLViec_Thang = GetValueCell("cln_SoNgayLViec_Thang", iRow)
                            dHeSoLuong_V2 = GetValueCell("cln_HeSoLuong_V2", iRow)
                            dMucTamUng_V2 = GetValueCell("cln_MucTamUng_V2", iRow)

                            dLuong_V1 = Math.Round(((dLCB_Tong_HeSo * dLuong_TTV) * dMucHuong) / 100) + dPhCap_ThuHut_ST_Goc + dPhCap_KhuVuc_ST_Goc
                            dLuong_V1_ThucTra = Math.Round(((dLCB_Tong_HeSo * dLuong_TTV) * dMucHuong) / 100)
                            dLuong_V2_TamUng = Math.Round((dLuong_V1_ThucTra * dHeSoLuong_V2) * (dMucTamUng_V2 / 100))
                            dLuong_ThucLinh = dLuong_V1 + dLuong_V2_TamUng
                            dSoTienBoSungKhac = Math.Round((dLuong_ThucLinh * dBoSung_SoTien_Ngay_Khac) / dSoNgayLViec_Thang) + dTru_Khoan_Khac_OLD

                        ElseIf cb_bosung_loai.SelectedIndex = 3 Then    'Số ngày lương theo lương bảo hiểm (Lương V1 làm căn cứ đóng bảo hiểm)
                            'dLuong_V1 = GetValueCell("cln_Luong_V1", iRow)
                            dSoNgayLViec_Thang = GetValueCell("cln_SoNgayLViec_Thang", iRow)
                            dLuong_V1 = Math.Round(((dLCB_Tong_HeSo * dLuong_TTV) * dMucHuong) / 100) + dPhCap_ThuHut_ST_Goc + dPhCap_KhuVuc_ST_Goc

                            dSoTienBoSungKhac = Math.Round((dLuong_V1 * dBoSung_SoTien_Ngay_Khac) / dSoNgayLViec_Thang) + dTru_Khoan_Khac_OLD
                        End If

                        dgv_main.Rows(iRow).Cells("cln_Tru_Khoan_Khac").Value = dSoTienBoSungKhac.ToString()
                        If (sGhiChu_OLD = "" Or sGhiChu_OLD Is Nothing) Then
                            dgv_main.Rows(iRow).Cells("cln_Tru_Khoan_Khac_GhiChu").Value = edt_bosung_ghichu_khac.Text
                        Else
                            dgv_main.Rows(iRow).Cells("cln_Tru_Khoan_Khac_GhiChu").Value = sGhiChu_OLD + "; " + edt_bosung_ghichu_khac.Text
                        End If
                        CallCalculateCelInGrid(iRow, "cln_Tru_Khoan_Khac")

                    End If
                Next

            End If
            'Lấy Lương V1 làm căn cứ đóng bảo hiểm nên em gọi tắt là lương bảo hiểm.
            'lương thực lĩnh bao gồm cả lương V1 + V2 ; Lương đóng bảo hiểm: Lương V1
            'Lương V1 = (Hệ số lương cấp bậc + hệ số phụ cấp lương (nếu có) )* mức lương cơ sở NHCS (hiện tại là :4.180.000) + phụ cấp khu vực thu hút (nếu có)
            'Lương V2 = Lương V1 * hệ sô tăng thêm * mức tạm ứng

            '            --- Chọn loại phải thu khác muốn cộng thêm vào cột [Các khoản trích nộp khác] ---
            'Số tiền cụ thể
            'Số ngày lương thực lĩnh
            'Số ngày lương theo lương bảo hiểm
        Catch ex As Exception
            MessageBox.Show("Lỗi khi nhấn nút lệnh bổ sung số tiền vào cột (Các khoản trích nộp khác): " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub cb_bosung_loai_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_bosung_loai.SelectedIndexChanged
        Try
            edt_bosung_sotien_khac.Enabled = False
            edt_bosung_sotien_khac.Text = "0"
            edt_bosung_ghichu_khac.Text = ""
            If (cb_bosung_loai.SelectedIndex > 0 And cb_bosung_loai.Items.Count <> 0) Then
                If cb_bosung_loai.SelectedIndex <> 0 Then
                    edt_bosung_sotien_khac.Enabled = True
                End If
            End If
            If cb_bosung_loai.SelectedIndex = 1 Then
                lbl_donvitinh.Text = "(đồng)"
            ElseIf cb_bosung_loai.SelectedIndex = 2 Or cb_bosung_loai.SelectedIndex = 3 Then
                lbl_donvitinh.Text = "(ngày)"
            End If

        Catch ex As Exception
            MessageBox.Show("Lỗi khi thay đổi loại bổ sung số tiền vào cột (Các khoản trích nộp khác): " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
End Class

