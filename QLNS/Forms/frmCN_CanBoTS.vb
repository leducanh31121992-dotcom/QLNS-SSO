Public Class frmCN_CanBoTS

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    'Mảng lưu id danh sách Đơn vị (Chi nhánh trong hệ thống ngân hàng)
    Private arr_Donvi As ArrayList = New ArrayList()
    'Mảng lưu id danh sách Phòng ban (Phòng ban này được load theo chi nhánh truyền vào)
    Private arr_Phongban As ArrayList = New ArrayList()
    'Mảng lưu id danh sách các Quốc gia --> Lấy dữ liệu từ bảng Quốc gia
    Private arr_Quoctich As ArrayList = New ArrayList()
    'Mảng lưu id danh sách - danh mục dân tộc --> Lấy dữ liệu từ bảng danh mục với Id_goc = 21
    Private arr_Dantoc As ArrayList = New ArrayList()
    'Mảng lưu id danh sách - danh mục Tôn giáo --> Lấy dữ liệu từ bảng danh mục với Id_goc = 24
    Private arr_Tongiao As ArrayList = New ArrayList()

    'Mảng lưu id danh sách tỉnh thành phố - Cho biết thông tin nơi sinh - tỉnh
    Private arr_Ns_Tinh As ArrayList = New ArrayList
    'Mảng lưu id danh sách quận huyện - Cho biết thông tin nơi sinh - quận huyện
    Private arr_Ns_Huyen As ArrayList = New ArrayList
    Private arr_Ns_Xa As ArrayList = New ArrayList
    Private arr_Ns_Thon As ArrayList = New ArrayList

    'Thông tin nguyên quán của cán bộ
    Private arr_Nq_Tinh As ArrayList = New ArrayList
    Private arr_Nq_Huyen As ArrayList = New ArrayList
    Private arr_Nq_Xa As ArrayList = New ArrayList
    Private arr_Nq_Thon As ArrayList = New ArrayList

    'Thông tin Thường trú của cán bộ
    Private arr_TT_Tinh As ArrayList = New ArrayList
    Private arr_TT_Huyen As ArrayList = New ArrayList
    Private arr_TT_Xa As ArrayList = New ArrayList
    Private arr_TT_Thon As ArrayList = New ArrayList

    'Thông tin Tạm trú của cán bộ
    Private arr_TTr_Tinh As ArrayList = New ArrayList()
    Private arr_TTr_Huyen As ArrayList = New ArrayList()
    Private arr_TTr_Xa As ArrayList = New ArrayList()
    Private arr_TTr_Thon As ArrayList = New ArrayList()

    'Mảng lưu id danh mục - Thành phần gia đình --> Truy suất dl đến danh mục với id gốc là 13
    Private arr_Tp_Giadinh As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Trình độ văn hoá --> Truy suất dl đến danh mục với id gốc là 6
    Private arr_Td_Vanhoa As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Trình độ Chính trị --> Truy suất dl đến danh mục với id gốc là 36
    Private arr_Td_Chinhtri As ArrayList = New ArrayList()

    'Mảng lưu id danh mục - Học hàm --> Truy suất dl đến danh mục với id gốc là 26
    Private arr_Hocham As ArrayList = New ArrayList()

    'Mảng lưu id danh mục - Học vị --> Truy suất dl đến danh mục với id gốc là 20
    Private arr_Hocvi As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Trình độ Chuyên môn  --> Truy suất dl đến danh mục với id gốc là 38
    Private arr_TrdChMon As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Trình độ Chuyên môn Ngoại ngữ  --> Truy suất dl đến danh mục với id gốc là 38 nhưng giới hạn 4 loại (ĐH, C, B, A)
    Private arr_TrdNgoaiNgu As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Trình độ Chuyên môn Tin học --> Truy suất dl đến danh mục với id gốc là 38 nhưng giới hạn 4 loại (ĐH, C, B, A)
    Private arr_TrdTinHoc As ArrayList = New ArrayList()
    'Mảng lưu id danh mục Chuyên ngành đào tạo --> Truy suất dl đến danh mục với id gốc là 11
    Private arr_ChNganh As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Ưu tiên gia đình --> Truy suất dl đến danh mục với id gốc là 19
    Private arr_Ut_Giadinh As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Ưu tiên bản thân --> Truy suất dl đến danh mục với id gốc là 18
    Private arr_Ut_Banthan As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Hình thức trả lương --> Truy suất dl đến danh mục với id gốc là 40
    Private arr_Ht_Traluong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Nghị định lương
    Private arr_NghiDinhLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Bảng lương
    Private arr_BangLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Ngạch lương
    Private arr_NgachLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Bậc lương
    Private arr_BacLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Chức vụ
    Private arr_Chucvu As ArrayList = New ArrayList()
    'Mảng lưu id danh mục - Chuyên môn được phân công  --> Truy suất dl đến danh mục với id gốc là 12
    Private arr_ChuyenMon As ArrayList = New ArrayList()
    Private ARL_TT_HonNhan As ArrayList = New ArrayList
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _Labour As clsHS_Hdld = New clsHS_Hdld()
    'Khai báo và khởi tạo thuộc tính của các đối tượng
    Private _IdCanBo As String
    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal value As String)
            _IdCanBo = value
        End Set
    End Property

    ''' <summary>
    ''' Get - Set lại biến lưu Node hiện thời khi click vào Treeview 
    ''' </summary>
    ''' <remarks></remarks>
    Private _Node As String
    Public Property Node() As String
        Get
            Return _Node
        End Get
        Set(ByVal value As String)
            _Node = value
        End Set
    End Property

    ''' <summary>
    ''' Biến lưu sự kiện người dùng Hồ sơ khác: 1 - Thêm mới  2 - Sửa đổi   3 - Xem chi tiết
    ''' </summary>
    ''' <remarks></remarks>
    Public _FlagEvent As Byte
    ''' <summary>
    ''' Biến lưu thông tin đường dẫn của ảnh khi Chọn mới
    ''' </summary>
    ''' <remarks></remarks>
    Private strPath As String = ""

    ''' <summary>
    ''' Biến lưu thông tin đường dẫn của ảnh khi Load ra sửa đổi
    ''' </summary>
    ''' <remarks></remarks>
    Private strPath_Old As String = ""

    'Biến lưu Id hợp đồng lao động của cán bộ tập sự
    Private _IdHdld As String = ""

    ''' <summary>
    ''' Biến lưu cờ báo --> Khi update lại dữ liệu không chọn ảnh nữa
    ''' </summary>
    ''' <remarks></remarks>
    Private flagChoice As Boolean = False
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

    'Id của cán bộ Chính thức muốn chuyển sang lao động ngắn hạn
    Private _IdCB As String
    Public Property IdCB() As String
        Get
            Return _IdCB
        End Get
        Set(ByVal value As String)
            _IdCB = value
        End Set
    End Property
#End Region

#Region "---> Functions: Các hàm chính <---"
    ''' <summary>
    ''' Hàm thực hiện - Thiết lập lại trạng thái ban đầu cho các Controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Control()
        strPath = ""
        strPath_Old = ""
        edt_macb.Text = ""
        edt_hoten.Text = ""
        edt_ten_tg.Text = ""
        edt_bidanh.Text = ""
        rb_nam.Checked = False
        rb_nu.Checked = False
        dtpk_ngaysinh.Text = DateTime.Now.ToShortDateString()

        If cb_donvi.Items.Count <> 0 And _IdCB = "" And _FlagEvent = 1 Then
            If (_Node <> "") Then
                cb_donvi.SelectedIndex = IIf(_Node.ToString.Trim().Substring(4) <> "", CType(arr_Donvi.IndexOf(_Node.ToString.Trim().Substring(4)), Integer), 0)
                cb_donvi_SelectedIndexChanged(Nothing, Nothing)
            Else
                cb_donvi.SelectedIndex = 1
            End If
        Else
            cb_donvi.SelectedIndex = 0
        End If
        cb_donvi_SelectedIndexChanged(Nothing, Nothing)
        If cb_quoctich.Items.Count <> 0 And _IdCB = "" And _FlagEvent = 1 Then
            cb_quoctich.SelectedIndex = 1
        Else
            cb_quoctich.SelectedIndex = 0
        End If
        If cb_dantoc.Items.Count <> 0 And _IdCB = "" And _FlagEvent = 1 Then
            cb_dantoc.SelectedIndex = 1
        Else
            cb_dantoc.SelectedIndex = 0
        End If
        If cb_tongiao.Items.Count <> 0 And _IdCB = "" And _FlagEvent = 1 Then
            cb_tongiao.SelectedIndex = 1
        Else
            cb_tongiao.SelectedIndex = 0
        End If
        edt_cmt_so.Text = ""
        dtpk_cmt_ngaycap.Text = DateTime.Now.ToShortDateString()
        edt_cmt_noicap.Text = ""
        cb_ns_tinh.SelectedIndex = 0
        cb_ns_tinh_SelectedIndexChanged(Nothing, Nothing)
        If (cb_ns_xa.Items.Count <> 0) Then
            cb_ns_xa.SelectedIndex = 0
        End If
        cb_ns_xa_SelectedIndexChanged(Nothing, Nothing)
        If cb_ns_thon.Items.Count <> 0 Then
            cb_ns_thon.SelectedIndex = 0
        End If
        cb_ns_thon_SelectedIndexChanged(Nothing, Nothing)
        edt_ns_diachi.Text = ""

        cb_nq_tinh.SelectedIndex = 0
        cb_nq_tinh_SelectedIndexChanged(Nothing, Nothing)
        If cb_nq_xa.Items.Count <> 0 Then
            cb_nq_xa.SelectedIndex = 0
        End If
        cb_nq_xa_SelectedIndexChanged(Nothing, Nothing)
        If cb_nq_thon.Items.Count <> 0 Then
            cb_nq_thon.SelectedIndex = 0
        End If
        cb_nq_thon_SelectedIndexChanged(Nothing, Nothing)
        edt_nq_diachi.Text = ""

        cb_tt_tinh.SelectedIndex = 0
        cb_tt_tinh_SelectedIndexChanged(Nothing, Nothing)
        If cb_tt_xa.Items.Count <> 0 Then
            cb_tt_xa.SelectedIndex = 0
        End If
        cb_tt_xa_SelectedIndexChanged(Nothing, Nothing)
        If cb_tt_thon.Items.Count <> 0 Then
            cb_tt_thon.SelectedIndex = 0
        End If
        cb_tt_thon_SelectedIndexChanged(Nothing, Nothing)
        edt_tt_diachi.Text = ""
        edt_tt_dienthoai.Text = ""

        cb_ttr_tinh.SelectedIndex = 0
        cb_ttr_tinh_SelectedIndexChanged(Nothing, Nothing)
        If cb_ttr_xa.Items.Count <> 0 Then
            cb_ttr_xa.SelectedIndex = 0
        End If
        cb_ttr_xa_SelectedIndexChanged(Nothing, Nothing)
        If cb_ttr_thon.Items.Count <> 0 Then
            cb_ttr_thon.SelectedIndex = 0
        End If
        edt_ttr_diachi.Text = ""
        edt_ttr_dienthoai.Text = ""

        edt_dt_didong.Text = ""
        edt_dt_coquan.Text = ""
        edt_dt_nharieng.Text = ""

        edt_email.Text = ""
        cb_nhommau.SelectedIndex = 0
        cb_tp_giadinh.SelectedIndex = 0
        cb_td_vanhoa.SelectedIndex = 0
        cb_td_chinhtri.SelectedIndex = 0
        cb_hocham.SelectedIndex = 0
        If (_FlagEvent = 1) Then
            cb_hocham.SelectedIndex = IIf(cb_hocham.FindString("Không") > 0, cb_hocham.FindString("Không"), 0)
        Else
            cb_hocham.SelectedIndex = 0
        End If
        cb_hocvi.SelectedIndex = 0
        cb_trd_chmon.SelectedIndex = 0
        cb_trd_ngngu.SelectedIndex = 0
        cb_trd_tinhoc.SelectedIndex = 0
        cb_chnganh.SelectedIndex = 0
        _Globals.ResetItems_CheckedListBox(clb_ut_banthan)
        If (_FlagEvent = 1) Then
            cb_ut_giadinh.SelectedIndex = IIf(cb_ut_giadinh.FindString("Không") > 0, cb_ut_giadinh.FindString("Không"), 0)
        Else
            cb_ut_giadinh.SelectedIndex = 0
        End If
        'Reset controls - Tab thông tin khác
        dtpk_cm_ngay_tg.Text = DateTime.Now.ToShortDateString()
        dtpk_cm_ngay_tg.Checked = False
        edt_cm_tochuc.Text = ""

        dtpk_dv_ngayvao.Text = DateTime.Now.ToShortDateString()
        dtpk_dv_ngayvao.Checked = False
        dtpk_dv_ngay_chinhthuc.Text = DateTime.Now.ToShortDateString()
        dtpk_dv_ngay_chinhthuc.Checked = False
        dtpk_dv_ngayra.Text = DateTime.Now.ToShortDateString()
        dtpk_dv_ngayra.Checked = False

        edt_dv_noikn.Text = ""
        edt_dv_sothe.Text = ""
        edt_dv_nguoi_gt.Text = ""
        edt_dv_lydo.Text = ""

        edt_tk_makh.Text = ""
        edt_tk_sotk.Text = ""
        edt_tk_tennh.Text = ""
        edt_ms_thue.Text = ""

        edt_bhxh_soso.Text = ""
        dtpk_bhxh_ngaylam.Text = DateTime.Now.ToShortDateString()
        dtpk_bhxh_ngaylam.Checked = False
        dtpk_bhxh_ngaydong.Text = DateTime.Now.ToShortDateString()
        dtpk_bhxh_ngaydong.Checked = False
        edt_bhxh_noilam.Text = ""

        'Thông tin Hợp đồng lao động
        rb_DinhBien.Checked = True
        rb_hdld_CheckedChanged(Nothing, Nothing)
        edt_hdld_sohd.Text = ""
        dtpk_hdld_ngayky.Text = DateTime.Now.ToShortDateString()
        edt_hdld_nguoiky.Text = ""
        cb_hdld_chucvu.SelectedIndex = 0
        If (_FlagEvent = 1 And _IdCB = "") Then
            edt_hdld_nguoiky.Text = _HS_CanBo.GetVarNam("GIAMDOC")
            Dim _IdIndex As Integer = CType(arr_Donvi(IIf(cb_donvi.SelectedIndex > 0, cb_donvi.SelectedIndex, "0")), Integer)
            Dim strSQL As String = String.Format("Select * from ChiNhanh Where Status = 1 and id = {0}", _IdIndex)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If db.Rows.Count > 0 Then
                        If db.Rows(0)("ma_so").ToString().Trim() <> "" And db.Rows(0)("ma_so").ToString().Trim() = gMaDonViTW Then
                            strSQL = "SELECT * FROM danhmuc WHERE id_goc = 14 And ma_so = '1402' And Status = 1"
                        Else
                            strSQL = "SELECT * FROM danhmuc WHERE id_goc = 14 And ma_so = '1410' And Status = 1"
                        End If
                        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db_child Is Nothing) Then
                                If db_child.Rows.Count > 0 Then
                                    If (db_child.Rows(0)("id").ToString().Trim() <> "") Then
                                        cb_hdld_chucvu.SelectedIndex = CType(arr_Chucvu.IndexOf(db_child.Rows(0)("id").ToString().Trim()), Integer)
                                    End If
                                End If
                            End If
                        End Using
                    End If
                End If
            End Using
        End If
        dtpk_tg_tungay.Text = DateTime.Now.ToShortDateString()
        dtpk_tg_denngay.Text = DateTime.Now.ToShortDateString()
        edt_gio_batdau.Text = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 7, 30, 0)
        edt_gio_ketthuc.Text = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 16, 30, 0)
        cb_hinhthuc_cv.SelectedIndex = 0
        cb_ht_traluong.SelectedIndex = 0
        cb_cdld_nghidinh.SelectedIndex = 0
        cb_cdld_nghidinh_SelectedIndexChanged(Nothing, Nothing)
        edt_cdld_tylehuong.Text = "0"
        edt_cdld_cv_damnhan.Text = ""
        cb_cdld_chuyenmon.SelectedIndex = 0
        edt_st_congtac.Text = ""
        edt_cv_launhat.Text = ""
        ckb_choice_bhxh.Checked = False
        ckb_choice_bhyt.Checked = False
        edt_ghichu.Text = ""
        edt_tienluong.Text = "0"
        'UnChecked các control - CheckBox
        ckb_cachmang.Checked = False
        ckb_cachmang_CheckedChanged(Nothing, Nothing)
        ckb_dangvien.Checked = False
        ckb_dangvien_CheckedChanged(Nothing, Nothing)
        ckb_taikhoan.Checked = False
        ckb_taikhoan_CheckedChanged(Nothing, Nothing)
        ckb_bhxh.Checked = False
        ckb_bhxh_CheckedChanged(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Load thông tin Cập nhật hồ sơ Cán bộ tập sự
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Load_Infors()
        ResetAll_Control()
        Select Case _FlagEvent
            Case 1 ' Load dữ liệu trong trường hợp thêm mới
                'If (picbx_main.Image IsNot Nothing) Then
                '    picbx_main.Image.Dispose()
                '    picbx_main.Image = Nothing
                'End If
                _IdHdld = ""
                If (_IdCB <> "") Then   'Trường hợp chuyển từ lao động dài hạn sang lao động ngắn hạn
                    Dim strSQL As String = ""
                    Dim dr As DataRow
                    dr = _HS_CanBo.GetHuman_ForCode(_IdCB)
                    If Not (dr Is Nothing) Then
                        If (dr.Table.Rows.Count > 0) Then
                            edt_macb.Text = dr("MaCB").ToString()
                            edt_hoten.Text = dr("HoTen").ToString()
                            edt_ten_tg.Text = dr("TenThuongGoi").ToString()
                            edt_bidanh.Text = dr("BiDanh").ToString()
                            edt_ten_tg.Text = dr("TenThuongGoi").ToString()
                            If (dr("GioiTinh").ToString() <> "") Then
                                If dr("GioiTinh").ToString() <> "True" Then
                                    rb_nam.Checked = True
                                Else
                                    rb_nu.Checked = True
                                End If
                            End If
                            dtpk_ngaysinh.Value = CType(dr("NgaySinh").ToString(), DateTime)
                            Dim _Iddv As Integer = IIf(dr("IdDonVi").ToString().Trim() <> "", CType(dr("IdDonVi").ToString().Trim(), Integer), 0)
                            cb_donvi.SelectedIndex = IIf(dr("IdDonVi").ToString() <> "", CType(arr_Donvi.IndexOf(dr("IdDonVi").ToString()), Integer), 0)
                            cb_donvi_SelectedIndexChanged(Nothing, Nothing)
                            strSQL = String.Format("Select Top 1 * From QDNhanSu as a, HS_CanBo as b Where a.IdCanBo = b.IdCanBo And a.IdCanBo = '{0}' And (a.IdDonvi_moi IN (Select id From ChiNhanh Where id = {1} or id_goc = {1})) And a.IsKiemNhiem = 0 Order by NgayKy_QD Desc", _IdCB, _Iddv)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        Dim _IddvNew As Integer = IIf(db.Rows(0)("IdDonvi_moi").ToString().Trim() <> "", CType(db.Rows(0)("IdDonvi_moi").ToString().Trim(), Integer), 0)
                                        If (_Iddv <> _IddvNew) Then
                                            'Phòng ban chính là IdDonvi_moi trong bảng Quyết định nhân sự (Phòng ban tương đương PGD)
                                            cb_phongban.SelectedIndex = IIf(_IddvNew > 0, CType(arr_Phongban.IndexOf("DV_" + _IddvNew.ToString()), Integer), 0)
                                        Else    'Trường hợp này chắc chắn cán bộ làm tại chi nhánh tỉnh
                                            'Lấy phòng ban mới của cán bộ tại tỉnh
                                            If (db.Rows(0)("IdPhong_Moi").ToString().Trim() <> "" And CType(db.Rows(0)("IdPhong_Moi").ToString().Trim(), Integer) > 0) Then
                                                Dim _PBNew As String = "PB_" + db.Rows(0)("IdPhong_Moi").ToString().Trim()
                                                cb_phongban.SelectedIndex = CType(arr_Phongban.IndexOf(_PBNew), Integer)
                                            End If
                                        End If
                                    End If
                                End If
                            End Using
                            cb_quoctich.SelectedIndex = IIf(dr("IdQuocTich").ToString() <> "", CType(arr_Quoctich.IndexOf(dr("IdQuocTich").ToString()), Integer), 0)
                            cb_dantoc.SelectedIndex = IIf(dr("IdDanToc").ToString() <> "", CType(arr_Dantoc.IndexOf(dr("IdDanToc").ToString()), Integer), 0)
                            cb_tongiao.SelectedIndex = IIf(dr("IdTonGiao").ToString() <> "", CType(arr_Tongiao.IndexOf(dr("IdTonGiao").ToString()), Integer), 0)
                            edt_cmt_so.Text = dr("CMT_So").ToString().Trim()
                            dtpk_cmt_ngaycap.Value = CType(dr("CMT_NgayCap").ToString(), DateTime)
                            edt_cmt_noicap.Text = dr("CMT_NoiCap").ToString().Trim()

                            '--> Fill ra control trong Tab sơ yếu lý lịch
                            cb_ns_tinh.SelectedIndex = IIf(dr("IdNS_Tinh").ToString() <> "", CType(arr_Ns_Tinh.IndexOf(dr("IdNS_Tinh").ToString()), Integer), 0)
                            cb_ns_tinh_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdNS_Huyen").ToString() <> "") Then
                                cb_ns_huyen.SelectedIndex = CType(arr_Ns_Huyen.IndexOf(dr("IdNS_Huyen").ToString()), Integer)
                            End If
                            If (dr("IdNS_Xa").ToString() <> "") Then
                                cb_ns_xa.SelectedIndex = CType(arr_Ns_Xa.IndexOf(dr("IdNS_Xa").ToString()), Int32)
                            End If
                            cb_ns_xa_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdNS_Thon").ToString() <> "") Then
                                cb_ns_thon.SelectedIndex = CType(arr_Ns_Thon.IndexOf(dr("IdNS_Thon").ToString()), Int32)
                            End If
                            edt_ns_diachi.Text = dr("NS_DChi").ToString().Trim()

                            cb_nq_tinh.SelectedIndex = IIf(dr("IdNQ_Tinh").ToString() <> "", CType(arr_Nq_Tinh.IndexOf(dr("IdNQ_Tinh").ToString()), Integer), 0)
                            cb_nq_tinh_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdNQ_Huyen").ToString() <> "") Then
                                cb_nq_huyen.SelectedIndex = CType(arr_Nq_Huyen.IndexOf(dr("IdNQ_Huyen").ToString()), Integer)
                            End If
                            If (dr("IdNQ_Xa").ToString() <> "") Then
                                cb_nq_xa.SelectedIndex = CType(arr_Nq_Xa.IndexOf(dr("IdNQ_Xa").ToString()), Int32)
                            End If
                            cb_nq_xa_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdNQ_Thon").ToString() <> "") Then
                                cb_nq_thon.SelectedIndex = CType(arr_Nq_Thon.IndexOf(dr("IdNQ_Thon").ToString()), Int32)
                            End If
                            edt_nq_diachi.Text = dr("NQ_DChi").ToString().Trim()

                            cb_tt_tinh.SelectedIndex = IIf(dr("IdThT_Tinh").ToString() <> "", CType(arr_TT_Tinh.IndexOf(dr("IdThT_Tinh").ToString()), Integer), 0)
                            cb_tt_tinh_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdThT_Huyen").ToString() <> "") Then
                                cb_tt_huyen.SelectedIndex = CType(arr_TT_Huyen.IndexOf(dr("IdThT_Huyen").ToString()), Integer)
                            End If
                            If (dr("IdThT_Xa").ToString() <> "") Then
                                cb_tt_xa.SelectedIndex = CType(arr_TT_Xa.IndexOf(dr("IdThT_Xa").ToString()), Int32)
                            End If
                            cb_tt_xa_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdThT_Thon").ToString() <> "") Then
                                cb_tt_thon.SelectedIndex = CType(arr_TT_Thon.IndexOf(dr("IdThT_Thon").ToString()), Int32)
                            End If
                            edt_tt_diachi.Text = dr("ThT_Diachi").ToString().Trim()
                            edt_tt_dienthoai.Text = dr("ThT_Dienthoai").ToString().Trim()

                            cb_ttr_tinh.SelectedIndex = IIf(dr("IdTTr_Tinh").ToString() <> "", CType(arr_TTr_Tinh.IndexOf(dr("IdTTr_Tinh").ToString()), Integer), 0)
                            cb_ttr_tinh_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdTTr_Huyen").ToString() <> "") Then
                                cb_ttr_huyen.SelectedIndex = CType(arr_TTr_Huyen.IndexOf(dr("IdTTr_Huyen").ToString()), Integer)
                            End If
                            If (dr("IdTTr_Xa").ToString() <> "") Then
                                cb_ttr_xa.SelectedIndex = CType(arr_TTr_Xa.IndexOf(dr("IdTTr_Xa").ToString()), Int32)
                            End If
                            cb_ttr_xa_SelectedIndexChanged(Nothing, Nothing)
                            If (dr("IdTTr_Thon").ToString() <> "") Then
                                cb_ttr_thon.SelectedIndex = CType(arr_TTr_Thon.IndexOf(dr("IdTTr_Thon").ToString()), Int32)
                            End If
                            edt_ttr_diachi.Text = dr("TTr_Diachi").ToString().Trim()
                            edt_ttr_dienthoai.Text = dr("TTr_Dienthoai").ToString().Trim()

                            edt_dt_coquan.Text = dr("DienThoai_CQ").ToString().Trim()
                            edt_dt_nharieng.Text = dr("DienThoai_NR").ToString().Trim()
                            edt_dt_didong.Text = dr("DienThoai_DD").ToString().Trim()

                            edt_email.Text = dr("Emai").ToString().Trim()
                            If (dr("NhomMau").ToString().Trim() = "A") Then
                                cb_nhommau.SelectedIndex = 1
                            ElseIf (dr("NhomMau").ToString().Trim() = "B") Then
                                cb_nhommau.SelectedIndex = 2
                            ElseIf (dr("NhomMau").ToString().Trim() = "AB") Then
                                cb_nhommau.SelectedIndex = 3
                            ElseIf (dr("NhomMau").ToString().Trim() = "O") Then
                                cb_nhommau.SelectedIndex = 4
                            Else : cb_nhommau.SelectedIndex = 0
                            End If

                            cb_tp_giadinh.SelectedIndex = IIf(dr("IdThanhPhanGD").ToString() <> "", CType(arr_Tp_Giadinh.IndexOf(dr("IdThanhPhanGD").ToString()), Integer), 0)
                            cb_td_vanhoa.SelectedIndex = IIf(dr("IdTrinhDoVH").ToString() <> "", CType(arr_Td_Vanhoa.IndexOf(dr("IdTrinhDoVH").ToString()), Integer), 0)
                            cb_td_chinhtri.SelectedIndex = IIf(dr("IdTrinhDoCT").ToString() <> "", CType(arr_Td_Chinhtri.IndexOf(dr("IdTrinhDoCT").ToString()), Integer), 0)

                            strSQL = String.Format("Select a.IdCB_HocHam,b.ten_goi as Hoc_Ham From CB_HocHam a, DanhMuc b Where (a.IdHocHam = b.id and b.id_goc = 26 and b.Status = 1) and IdCanBo = '{0}' Order by b.ma_so", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        cb_hocham.SelectedIndex = IIf(db.Rows(0)("IdCB_HocHam").ToString() <> "", CType(arr_Hocham.IndexOf(db.Rows(0)("IdCB_HocHam").ToString()), Integer), 0)
                                    End If
                                End If
                            End Using
                            strSQL = String.Format("Select a.IdCB_HocVi,b.ten_goi as Hoc_Vi From CB_HocVi a, DanhMuc b Where (a.IdHocVi = b.id and b.id_goc = 20 and b.Status = 1) and IdCanBo = '{0}' Order by b.ma_so", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        cb_hocvi.SelectedIndex = IIf(db.Rows(0)("IdCB_HocVi").ToString() <> "", CType(arr_Hocvi.IndexOf(db.Rows(0)("IdCB_HocVi").ToString()), Integer), 0)
                                    End If
                                End If
                            End Using

                            '--> Fill ra control trong Tab Thông tin hồ sơ khác
                            edt_ms_thue.Text = dr("MaSoThue").ToString().Trim()
                            cb_ut_giadinh.SelectedIndex = IIf(dr("IdUT_GDinh").ToString() <> "", CType(arr_Ut_Giadinh.IndexOf(dr("IdUT_GDinh").ToString()), Integer), 0)
                            If (dr("IdUT_BThan").ToString() <> "") Then
                                _Globals.SetItems_CheckedListBox(clb_ut_banthan, dr("IdUT_BThan").ToString(), arr_Ut_Banthan)
                            End If
                            If dr("CM_Ngay").ToString().Trim() <> "" Then
                                If (CType(dr("CM_Ngay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                    dtpk_cm_ngay_tg.Checked = False
                                Else
                                    dtpk_cm_ngay_tg.Checked = True
                                    dtpk_cm_ngay_tg.Value = CType(dr("CM_Ngay").ToString(), DateTime)
                                    ckb_cachmang.Checked = True
                                    ckb_cachmang_CheckedChanged(Nothing, Nothing)
                                End If
                            Else
                                dtpk_cm_ngay_tg.Checked = False
                            End If
                            edt_cm_tochuc.Text = dr("CM_ToChuc").ToString().Trim()
                            '           Lấy thông tin hồ sơ đảng
                            strSQL = String.Format("Select * from HS_DangVien Where IdCanBo = '{0}' Order by NgayVao Desc", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        If db.Rows(0)("NgayKN").ToString().Trim() <> "" Then
                                            If (CType(db.Rows(0)("NgayKN"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                dtpk_dv_ngayvao.Checked = False
                                            Else
                                                dtpk_dv_ngayvao.Value = CType(db.Rows(0)("NgayKN").ToString(), DateTime)
                                                dtpk_dv_ngayvao.Checked = True
                                            End If
                                        Else
                                            dtpk_dv_ngayvao.Checked = False
                                        End If
                                        If db.Rows(0)("NgayVao").ToString().Trim() <> "" Then
                                            If (CType(db.Rows(0)("NgayVao"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                dtpk_dv_ngay_chinhthuc.Checked = False
                                            Else
                                                dtpk_dv_ngay_chinhthuc.Value = CType(db.Rows(0)("NgayVao").ToString(), DateTime)
                                                dtpk_dv_ngay_chinhthuc.Checked = True
                                            End If
                                        Else
                                            dtpk_dv_ngay_chinhthuc.Checked = False
                                        End If

                                        If db.Rows(0)("NgayRa").ToString().Trim() <> "" Then
                                            If (CType(db.Rows(0)("NgayRa"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                dtpk_dv_ngayra.Checked = False
                                            Else
                                                dtpk_dv_ngayra.Value = CType(db.Rows(0)("NgayRa").ToString(), DateTime)
                                                dtpk_dv_ngayra.Checked = True
                                            End If
                                        Else
                                            dtpk_dv_ngayra.Checked = False
                                        End If
                                        dtpk_dv_ngayra_Leave(Nothing, Nothing)
                                        edt_dv_noikn.Text = db.Rows(0)("Noi_KetNap").ToString().Trim()
                                        edt_dv_sothe.Text = db.Rows(0)("SoThe").ToString().Trim()
                                        edt_dv_nguoi_gt.Text = db.Rows(0)("Nguoi_GT").ToString().Trim()
                                        edt_dv_lydo.Text = db.Rows(0)("LyDo").ToString().Trim()
                                        If (db.Rows(0)("Noi_KetNap").ToString().Trim() <> "") Then
                                            ckb_dangvien.Checked = True
                                        Else
                                            ckb_dangvien.Checked = False
                                        End If
                                        ckb_dangvien_CheckedChanged(Nothing, Nothing)
                                    End If
                                End If
                            End Using
                            'Tài khoản ngân hàng
                            edt_tk_makh.Text = dr("NH_MaKH").ToString().Trim()
                            edt_tk_sotk.Text = dr("NH_SoTK").ToString().Trim()
                            edt_tk_tennh.Text = dr("NH_TenNH").ToString().Trim()
                            If (dr("NH_SoTK").ToString().Trim() <> "") Then
                                ckb_taikhoan.Checked = True
                            Else
                                ckb_taikhoan.Checked = False
                            End If
                            ckb_taikhoan_CheckedChanged(Nothing, Nothing)

                            'Thông tin sổ bảo hiểm xã hội
                            edt_bhxh_soso.Text = dr("BHXH_SoSo").ToString().Trim()
                            If dr("BHXH_NgaySo").ToString().Trim() <> "" Then
                                dtpk_bhxh_ngaylam.Value = CType(dr("BHXH_NgaySo").ToString(), DateTime)
                                dtpk_bhxh_ngaylam.Checked = True
                            Else
                                dtpk_bhxh_ngaylam.Checked = False
                            End If
                            If dr("BHXH_NgayBatDau").ToString().Trim() <> "" Then
                                dtpk_bhxh_ngaydong.Value = CType(dr("BHXH_NgayBatDau").ToString(), DateTime)
                                dtpk_bhxh_ngaydong.Checked = True
                            Else
                                dtpk_bhxh_ngaydong.Checked = False
                            End If
                            edt_bhxh_noilam.Text = dr("BHXH_NoiLam").ToString().Trim()
                            If (dr("BHXH_SoSo").ToString().Trim() <> "") Then
                                ckb_bhxh.Checked = True
                            Else
                                ckb_bhxh.Checked = False
                            End If
                            ckb_bhxh_CheckedChanged(Nothing, Nothing)
                            '   Chuyên ngành đào tạo - lấy thằng mới nhất mà có thể liên quan đến học vị cao nhất
                            strSQL = String.Format("Select * from HS_DTVBCC Where IdCanBo = '{0}' And VBCC = 1 And HoanThanh = 1 order by NgayHL desc", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        cb_chnganh.SelectedIndex = IIf(db.Rows(0)("IdChuyenNganhDT").ToString() <> "", CType(arr_ChNganh.IndexOf(db.Rows(0)("IdChuyenNganhDT").ToString()), Integer), 0)
                                    End If
                                End If
                            End Using
                            'Các thông tin còn lại - của tab thông tin khác
                            edt_st_congtac.Text = dr("SoTruong_CT").ToString().Trim()
                            edt_cv_launhat.Text = dr("CV_Lau").ToString().Trim()
                            edt_ghichu.Text = dr("GhiChu").ToString().Trim()
                            cb_tt_honnhan.SelectedIndex = IIf(My_CStr(dr("HonNhan_Cd").ToString(), "") <> "", ARL_TT_HonNhan.IndexOf(My_CStr(dr("HonNhan_Cd").ToString(), "")), 0)
                            '--> Fill ra control trong Tab Hợp đồng lao động
                            strSQL = String.Format("Select * from HS_HDLD Where IdCanBo = '{0}' Order by NgayKy_HD desc", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        rb_DinhBien.Checked = True
                                        rb_hdld_CheckedChanged(Nothing, Nothing)
                                        edt_hdld_sohd.Text = db.Rows(0)("SoHD").ToString().Trim()
                                        If (db.Rows(0)("NgayKy_HD").ToString() <> "") Then
                                            dtpk_hdld_ngayky.Value = CType(db.Rows(0)("NgayKy_HD").ToString(), DateTime)
                                        End If
                                        edt_hdld_nguoiky.Text = db.Rows(0)("NguoKy_QD").ToString().Trim()
                                        cb_hdld_chucvu.SelectedIndex = IIf(db.Rows(0)("IdCV_Nguoiky_QD").ToString() <> "", CType(arr_Chucvu.IndexOf(db.Rows(0)("IdCV_Nguoiky_QD").ToString()), Integer), 0)
                                        If db.Rows(0)("Tungay").ToString().Trim() <> "" Then
                                            dtpk_tg_tungay.Value = CType(db.Rows(0)("Tungay").ToString(), DateTime)
                                        End If
                                        If db.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                            dtpk_tg_denngay.Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime)
                                        End If
                                        cb_hinhthuc_cv.SelectedIndex = IIf(cb_hinhthuc_cv.FindString("Hợp đồng ngắn hạn") > 0, cb_hinhthuc_cv.FindString("Hợp đồng ngắn hạn"), 0)
                                        cb_ht_traluong.SelectedIndex = IIf(db.Rows(0)("IdHT_TraLuong").ToString() <> "", CType(arr_Ht_Traluong.IndexOf(db.Rows(0)("IdHT_TraLuong").ToString()), Integer), 0)
                                        Dim _BangLuongId As Integer = 0
                                        Dim _NgachLuongId As Integer = 0
                                        Dim _BacLuongId As Integer = 0
                                        Dim _NghiDinhLuongId As Integer = 0

                                        _BacLuongId = IIf(db.Rows(0)("IdBacLuong").ToString() <> "0", CType(db.Rows(0)("IdBacLuong").ToString(), Integer), 0) ' Id bậc lương
                                        If (_BacLuongId > 0) Then
                                            rb_thangbang.Checked = True
                                            rb_thangbang_CheckedChanged(Nothing, Nothing)
                                            _NgachLuongId = _Labour.GetIdNgachLuong(_BacLuongId)
                                            If (_NgachLuongId > 0) Then
                                                _BangLuongId = _Labour.GetIdBangLuong(_NgachLuongId)
                                            End If
                                            If (_BangLuongId > 0) Then
                                                _NghiDinhLuongId = _Labour.GetIdNghiDinhLuong(_BangLuongId)
                                            End If
                                            cb_cdld_nghidinh.SelectedIndex = CType(arr_NghiDinhLuong.IndexOf(_NghiDinhLuongId.ToString()), Integer)
                                            cb_cdld_nghidinh_SelectedIndexChanged(Nothing, Nothing)
                                            cb_cdld_bangluong.SelectedIndex = CType(arr_BangLuong.IndexOf(_BangLuongId.ToString()), Integer)
                                            cb_cdld_bangluong_SelectedIndexChanged(Nothing, Nothing)
                                            cb_cdld_ngachluong.SelectedIndex = CType(arr_NgachLuong.IndexOf(_NgachLuongId.ToString()), Integer)
                                            cb_cdld_ngachluong_SelectedIndexChanged(Nothing, Nothing)
                                            cb_cdld_bacluong.SelectedIndex = CType(arr_BacLuong.IndexOf(db.Rows(0)("IdBacLuong").ToString()), Integer)
                                            cb_cdld_bacluong_SelectedIndexChanged(Nothing, Nothing)
                                            edt_cdld_tylehuong.Text = IIf(db.Rows(0)("TyleHuong").ToString().Trim() <> "", CType(db.Rows(0)("TyleHuong"), Double).ToString("N2"), "0")
                                        Else
                                            rb_trongoi.Checked = True
                                            rb_trongoi_CheckedChanged(Nothing, Nothing)
                                            edt_tienluong.Text = "0"
                                        End If
                                    End If
                                End If
                            End Using
                            strSQL = String.Format("Select * from HS_BHXH Where IdCanBo = '{0}' Order by TuNgay Desc", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        ckb_choice_bhxh.Checked = True
                                    End If
                                End If
                            End Using
                            strSQL = String.Format("Select * from HS_BHYT Where IdCanBo = '{0}' Order by TuNgay Desc", _IdCB)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        ckb_choice_bhyt.Checked = True
                                    End If
                                End If
                            End Using
                            'Gọi hàm Show picture của cán bộ công nhân viên
                            My_LoadImage(picbx_main, dr("AnhThe"))
                        End If
                    End If
                End If
                ActiveControl = edt_macb
            Case 2 ' Load dữ liệu trong trường hợp sửa đổi
                strPath = ""
                strPath_Old = ""
                flagChoice = False
                _IdHdld = ""
                'Load thông tin cần sửa đổi - Về hồ sơ cán bộ tập sự
                'Dim dr As DataRow
                Dim db_rowedit As DataTable = New DataTable()
                db_rowedit = _HS_CanBo.GetAll_Trainee_New(0, _IdCanBo, "", "", CType("1", Byte), CType("1", Byte))
                'dr = db_new.Rows(0)
                If Not (db_rowedit Is Nothing) Then
                    If (db_rowedit.Rows.Count > 0) Then
                        edt_macb.Text = db_rowedit.Rows(0)("MaCB").ToString()
                        edt_hoten.Text = db_rowedit.Rows(0)("HoTen").ToString()
                        edt_ten_tg.Text = db_rowedit.Rows(0)("TenThuongGoi").ToString()
                        edt_bidanh.Text = db_rowedit.Rows(0)("BiDanh").ToString()
                        edt_ten_tg.Text = db_rowedit.Rows(0)("TenThuongGoi").ToString()
                        If (db_rowedit.Rows(0)("GioiTinh").ToString() <> "") Then
                            If db_rowedit.Rows(0)("GioiTinh").ToString() <> "True" Then
                                rb_nam.Checked = True
                            Else
                                rb_nu.Checked = True
                            End If
                        End If
                        dtpk_ngaysinh.Value = CType(db_rowedit.Rows(0)("NgaySinh").ToString(), DateTime)
                        cb_donvi.SelectedIndex = IIf(db_rowedit.Rows(0)("IdChiNhanh").ToString() <> "", CType(arr_Donvi.IndexOf(db_rowedit.Rows(0)("IdChiNhanh").ToString()), Integer), 0)
                        cb_donvi_SelectedIndexChanged(Nothing, Nothing)
                        cb_phongban.SelectedIndex = IIf(db_rowedit.Rows(0)("IdPhongBan").ToString() <> "", CType(arr_Phongban.IndexOf(db_rowedit.Rows(0)("IdPhongBan").ToString()), Integer), 0)
                        cb_quoctich.SelectedIndex = IIf(db_rowedit.Rows(0)("IdQuocTich").ToString() <> "", CType(arr_Quoctich.IndexOf(db_rowedit.Rows(0)("IdQuocTich").ToString()), Integer), 0)
                        cb_dantoc.SelectedIndex = IIf(db_rowedit.Rows(0)("IdDanToc").ToString() <> "", CType(arr_Dantoc.IndexOf(db_rowedit.Rows(0)("IdDanToc").ToString()), Integer), 0)
                        cb_tongiao.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTonGiao").ToString() <> "", CType(arr_Tongiao.IndexOf(db_rowedit.Rows(0)("IdTonGiao").ToString()), Integer), 0)
                        edt_cmt_so.Text = db_rowedit.Rows(0)("CMT_So").ToString().Trim()
                        dtpk_cmt_ngaycap.Value = CType(db_rowedit.Rows(0)("CMT_NgayCap").ToString(), DateTime)
                        edt_cmt_noicap.Text = db_rowedit.Rows(0)("CMT_NoiCap").ToString().Trim()

                        'Show ảnh thẻ và lấy đường dẫn của ảnh
                        ''Dim obj As Object = dr("AnhThe")
                        If Not db_rowedit.Rows(0)("AnhThe") Is Nothing AndAlso Len(db_rowedit.Rows(0)("AnhThe").ToString()) > 0 Then
                            My_LoadImage(picbx_main, db_rowedit.Rows(0)("AnhThe"))
                        End If

                        'Fill data lên tab - Sơ yếu lý lịch cán bộ
                        cb_ns_tinh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdNS_Tinh").ToString() <> "", CType(arr_Ns_Tinh.IndexOf(db_rowedit.Rows(0)("IdNS_Tinh").ToString()), Integer), 0)
                        cb_ns_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdNS_Huyen").ToString() <> "") Then
                            cb_ns_huyen.SelectedIndex = CType(arr_Ns_Huyen.IndexOf(db_rowedit.Rows(0)("IdNS_Huyen").ToString()), Integer)
                        End If
                        If (db_rowedit.Rows(0)("IdNS_Xa").ToString() <> "") Then
                            cb_ns_xa.SelectedIndex = CType(arr_Ns_Xa.IndexOf(db_rowedit.Rows(0)("IdNS_Xa").ToString()), Int32)
                        End If
                        cb_ns_xa_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdNS_Thon").ToString() <> "") Then
                            cb_ns_thon.SelectedIndex = CType(arr_Ns_Thon.IndexOf(db_rowedit.Rows(0)("IdNS_Thon").ToString()), Int32)
                        End If
                        edt_ns_diachi.Text = db_rowedit.Rows(0)("NS_DiaChi").ToString().Trim()

                        cb_nq_tinh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdNQ_Tinh").ToString() <> "", CType(arr_Nq_Tinh.IndexOf(db_rowedit.Rows(0)("IdNQ_Tinh").ToString()), Integer), 0)
                        cb_nq_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdNQ_Huyen").ToString() <> "") Then
                            cb_nq_huyen.SelectedIndex = CType(arr_Nq_Huyen.IndexOf(db_rowedit.Rows(0)("IdNQ_Huyen").ToString()), Integer)
                        End If
                        If (db_rowedit.Rows(0)("IdNQ_Xa").ToString() <> "") Then
                            cb_nq_xa.SelectedIndex = CType(arr_Nq_Xa.IndexOf(db_rowedit.Rows(0)("IdNQ_Xa").ToString()), Int32)
                        End If
                        cb_nq_xa_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdNQ_Thon").ToString() <> "") Then
                            cb_nq_thon.SelectedIndex = CType(arr_Nq_Thon.IndexOf(db_rowedit.Rows(0)("IdNQ_Thon").ToString()), Int32)
                        End If
                        edt_nq_diachi.Text = db_rowedit.Rows(0)("NQ_DiaChi").ToString().Trim()

                        cb_tt_tinh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdThT_Tinh").ToString() <> "", CType(arr_TT_Tinh.IndexOf(db_rowedit.Rows(0)("IdThT_Tinh").ToString()), Integer), 0)
                        cb_tt_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdThT_Huyen").ToString() <> "") Then
                            cb_tt_huyen.SelectedIndex = CType(arr_TT_Huyen.IndexOf(db_rowedit.Rows(0)("IdThT_Huyen").ToString()), Integer)
                        End If
                        If (db_rowedit.Rows(0)("IdThT_Xa").ToString() <> "") Then
                            cb_tt_xa.SelectedIndex = CType(arr_TT_Xa.IndexOf(db_rowedit.Rows(0)("IdThT_Xa").ToString()), Int32)
                        End If
                        cb_tt_xa_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdThT_Thon").ToString() <> "") Then
                            cb_tt_thon.SelectedIndex = CType(arr_TT_Thon.IndexOf(db_rowedit.Rows(0)("IdThT_Thon").ToString()), Int32)
                        End If
                        edt_tt_diachi.Text = db_rowedit.Rows(0)("ThT_Diachi").ToString().Trim()
                        edt_tt_dienthoai.Text = db_rowedit.Rows(0)("ThT_Dienthoai").ToString().Trim()

                        cb_ttr_tinh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTTr_Tinh").ToString() <> "", CType(arr_TTr_Tinh.IndexOf(db_rowedit.Rows(0)("IdTTr_Tinh").ToString()), Integer), 0)
                        cb_ttr_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdTTr_Huyen").ToString() <> "") Then
                            cb_ttr_huyen.SelectedIndex = CType(arr_TTr_Huyen.IndexOf(db_rowedit.Rows(0)("IdTTr_Huyen").ToString()), Integer)
                        End If
                        If (db_rowedit.Rows(0)("IdTTr_Xa").ToString() <> "") Then
                            cb_ttr_xa.SelectedIndex = CType(arr_TTr_Xa.IndexOf(db_rowedit.Rows(0)("IdTTr_Xa").ToString()), Int32)
                        End If
                        cb_ttr_xa_SelectedIndexChanged(Nothing, Nothing)
                        If (db_rowedit.Rows(0)("IdTTr_Thon").ToString() <> "") Then
                            cb_ttr_thon.SelectedIndex = CType(arr_TTr_Thon.IndexOf(db_rowedit.Rows(0)("IdTTr_Thon").ToString()), Int32)
                        End If
                        edt_ttr_diachi.Text = db_rowedit.Rows(0)("TTr_Diachi").ToString().Trim()
                        edt_ttr_dienthoai.Text = db_rowedit.Rows(0)("TTr_Dienthoai").ToString().Trim()
                        edt_dt_coquan.Text = db_rowedit.Rows(0)("DienThoai_CQ").ToString().Trim()
                        edt_dt_nharieng.Text = db_rowedit.Rows(0)("DienThoai_NR").ToString().Trim()
                        edt_dt_didong.Text = db_rowedit.Rows(0)("DienThoai_DD").ToString().Trim()
                        edt_sofax.Text = db_rowedit.Rows(0)("SoFax").ToString().Trim()
                        edt_email.Text = db_rowedit.Rows(0)("Email").ToString().Trim()
                        'Fill thông tin của nhóm máu
                        cb_nhommau.SelectedIndex = GetIndex_NhomMau(db_rowedit.Rows(0)("NhomMau").ToString().Trim())
                        cb_tp_giadinh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdThanhPhan_GD").ToString() <> "", CType(arr_Tp_Giadinh.IndexOf(db_rowedit.Rows(0)("IdThanhPhan_GD").ToString()), Integer), 0)
                        cb_ut_giadinh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdUT_GD").ToString() <> "", CType(arr_Ut_Giadinh.IndexOf(db_rowedit.Rows(0)("IdUT_GD").ToString()), Integer), 0)
                        If (db_rowedit.Rows(0)("IdUT_BT").ToString() <> "") Then
                            _Globals.SetItems_CheckedListBox(clb_ut_banthan, db_rowedit.Rows(0)("IdUT_BT").ToString(), arr_Ut_Banthan)
                        End If
                        cb_td_vanhoa.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTrinhDoVH").ToString() <> "", CType(arr_Td_Vanhoa.IndexOf(db_rowedit.Rows(0)("IdTrinhDoVH").ToString()), Integer), 0)
                        cb_td_chinhtri.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTrinhDoCT").ToString() <> "", CType(arr_Td_Chinhtri.IndexOf(db_rowedit.Rows(0)("IdTrinhDoCT").ToString()), Integer), 0)
                        cb_hocham.SelectedIndex = IIf(db_rowedit.Rows(0)("IdHocHam").ToString() <> "", CType(arr_Hocham.IndexOf(db_rowedit.Rows(0)("IdHocHam").ToString()), Integer), 0)
                        cb_hocvi.SelectedIndex = IIf(db_rowedit.Rows(0)("IdHocVi").ToString() <> "", CType(arr_Hocvi.IndexOf(db_rowedit.Rows(0)("IdHocVi").ToString()), Integer), 0)
                        cb_trd_chmon.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTrdChMon").ToString() <> "", CType(arr_TrdChMon.IndexOf(db_rowedit.Rows(0)("IdTrdChMon").ToString()), Integer), 0)
                        cb_trd_ngngu.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTrdNgoaiNgu").ToString() <> "", CType(arr_TrdNgoaiNgu.IndexOf(db_rowedit.Rows(0)("IdTrdNgoaiNgu").ToString()), Integer), 0)
                        cb_trd_tinhoc.SelectedIndex = IIf(db_rowedit.Rows(0)("IdTrdTinHoc").ToString() <> "", CType(arr_TrdTinHoc.IndexOf(db_rowedit.Rows(0)("IdTrdTinHoc").ToString()), Integer), 0)
                        cb_chnganh.SelectedIndex = IIf(db_rowedit.Rows(0)("IdChuyenNganhDT").ToString() <> "", CType(arr_ChNganh.IndexOf(db_rowedit.Rows(0)("IdChuyenNganhDT").ToString()), Integer), 0)
                        'Fill data lên tab - Thông tin khác về Hồ sơ cán hộ tập sự
                        'Thông tin về việc Tham gia cách mạng của Cán bộ
                        If (db_rowedit.Rows(0)("CM_ToChuc").ToString().Trim() <> "") Then
                            ckb_cachmang.Checked = True
                        Else
                            ckb_cachmang.Checked = False
                        End If
                        ckb_cachmang_CheckedChanged(Nothing, Nothing)
                        If db_rowedit.Rows(0)("CM_Ngay").ToString().Trim() <> "" Then
                            If (CType(db_rowedit.Rows(0)("CM_Ngay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_cm_ngay_tg.Checked = False
                            Else
                                dtpk_cm_ngay_tg.Checked = True
                                dtpk_cm_ngay_tg.Value = CType(db_rowedit.Rows(0)("CM_Ngay").ToString(), DateTime)
                                ckb_cachmang.Checked = True
                                ckb_cachmang_CheckedChanged(Nothing, Nothing)
                            End If
                        Else
                            dtpk_cm_ngay_tg.Checked = False
                        End If
                        edt_cm_tochuc.Text = db_rowedit.Rows(0)("CM_ToChuc").ToString().Trim()
                        'Thông tin về việc Tham gia vào Đảng Cán bộ
                        If db_rowedit.Rows(0)("Dang_NgayVao").ToString().Trim() <> "" Then
                            If (CType(db_rowedit.Rows(0)("Dang_NgayVao"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_dv_ngayvao.Checked = False
                            Else
                                dtpk_dv_ngayvao.Value = CType(db_rowedit.Rows(0)("Dang_NgayVao").ToString(), DateTime)
                                dtpk_dv_ngayvao.Checked = True
                            End If
                        Else
                            dtpk_dv_ngayvao.Checked = False
                        End If
                        If db_rowedit.Rows(0)("Dang_NgayChTh").ToString().Trim() <> "" Then
                            If (CType(db_rowedit.Rows(0)("Dang_NgayChTh"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_dv_ngay_chinhthuc.Checked = False
                            Else
                                dtpk_dv_ngay_chinhthuc.Value = CType(db_rowedit.Rows(0)("Dang_NgayChTh").ToString(), DateTime)
                                dtpk_dv_ngay_chinhthuc.Checked = True
                            End If
                        Else
                            dtpk_dv_ngay_chinhthuc.Checked = False
                        End If
                        If db_rowedit.Rows(0)("Dang_NgayRa").ToString().Trim() <> "" Then
                            If (CType(db_rowedit.Rows(0)("Dang_NgayRa"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_dv_ngayra.Checked = False
                            Else
                                dtpk_dv_ngayra.Value = CType(db_rowedit.Rows(0)("Dang_NgayRa").ToString(), DateTime)
                                dtpk_dv_ngayra.Checked = True
                            End If
                        Else
                            dtpk_dv_ngayra.Checked = False
                        End If
                        dtpk_dv_ngayra_Leave(Nothing, Nothing)
                        edt_dv_noikn.Text = db_rowedit.Rows(0)("Dang_NoiKN").ToString().Trim()
                        edt_dv_sothe.Text = db_rowedit.Rows(0)("Dang_SoThe").ToString().Trim()
                        edt_dv_nguoi_gt.Text = db_rowedit.Rows(0)("Dang_NGT").ToString().Trim()
                        edt_dv_lydo.Text = db_rowedit.Rows(0)("Dang_LyDoRa").ToString().Trim()
                        If (db_rowedit.Rows(0)("Dang_NoiKN").ToString().Trim() <> "") Then
                            ckb_dangvien.Checked = True
                        Else
                            ckb_dangvien.Checked = False
                        End If
                        ckb_dangvien_CheckedChanged(Nothing, Nothing)
                        'Thông tin tài khoản ngân hàng
                        edt_tk_makh.Text = db_rowedit.Rows(0)("NH_MSKH").ToString().Trim()
                        edt_tk_sotk.Text = db_rowedit.Rows(0)("NH_SHTK").ToString().Trim()
                        edt_tk_tennh.Text = db_rowedit.Rows(0)("NH_Ten_NH").ToString().Trim()
                        If (db_rowedit.Rows(0)("NH_SHTK").ToString().Trim() <> "") Then
                            ckb_taikhoan.Checked = True
                        Else
                            ckb_taikhoan.Checked = False
                        End If
                        ckb_taikhoan_CheckedChanged(Nothing, Nothing)
                        edt_ms_thue.Text = db_rowedit.Rows(0)("MaSoThue").ToString().Trim()

                        'Thông tin sổ bảo hiểm xã hội
                        edt_bhxh_soso.Text = db_rowedit.Rows(0)("BHXH_SoSo").ToString().Trim()
                        If db_rowedit.Rows(0)("BHXH_NgayLam").ToString().Trim() <> "" Then
                            If (CType(db_rowedit.Rows(0)("BHXH_NgayLam"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_bhxh_ngaylam.Checked = False
                            Else
                                dtpk_bhxh_ngaylam.Value = CType(db_rowedit.Rows(0)("BHXH_NgayLam").ToString(), DateTime)
                                dtpk_bhxh_ngaylam.Checked = True
                            End If
                        Else
                            dtpk_bhxh_ngaylam.Checked = False
                        End If
                        If db_rowedit.Rows(0)("BHXH_NgayDong").ToString().Trim() <> "" Then
                            If (CType(db_rowedit.Rows(0)("BHXH_NgayDong"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_bhxh_ngaydong.Checked = False
                            Else
                                dtpk_bhxh_ngaydong.Value = CType(db_rowedit.Rows(0)("BHXH_NgayDong").ToString(), DateTime)
                                dtpk_bhxh_ngaydong.Checked = True
                            End If
                        Else
                            dtpk_bhxh_ngaydong.Checked = False
                        End If
                        edt_bhxh_noilam.Text = db_rowedit.Rows(0)("BHXH_NoiLam").ToString().Trim()
                        If (db_rowedit.Rows(0)("BHXH_SoSo").ToString().Trim() <> "") Then
                            ckb_bhxh.Checked = True
                        Else
                            ckb_bhxh.Checked = False
                        End If
                        ckb_bhxh_CheckedChanged(Nothing, Nothing)

                        'Các thông tin còn lại - của tab thông tin khác
                        edt_st_congtac.Text = db_rowedit.Rows(0)("SoTruong_CT").ToString().Trim()
                        edt_cv_launhat.Text = db_rowedit.Rows(0)("CV_Lau").ToString().Trim()
                        edt_ghichu.Text = db_rowedit.Rows(0)("GhiChu").ToString().Trim()
                        cb_tt_honnhan.SelectedIndex = IIf(My_CStr(db_rowedit.Rows(0)("HonNhan_Cd").ToString(), "") <> "", ARL_TT_HonNhan.IndexOf(My_CStr(db_rowedit.Rows(0)("HonNhan_Cd").ToString(), "")), 0)

                        'Các thông tin Hợp đồng lao động của Hồ sơ cán bộ Tập sự
                        'Lấy dữ liệu Hợp đồng lao động của cán bộ tập sự
                        Dim strSQL As String = String.Format("Select * from HSCB_TS_HDLD Where IdCanbo = '{0}' Order by DenNgay Desc", _IdCanBo)
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    _IdHdld = db.Rows(0)("IdCBTS_HDLD").ToString().Trim()
                                    If (db.Rows(0)("Status").ToString() <> "") Then
                                        If db.Rows(0)("Status").ToString() <> "True" Then
                                            rb_DinhBien.Checked = True
                                            rb_hdld_CheckedChanged(Nothing, Nothing)
                                        Else
                                            rb_PhuTro.Checked = True
                                            rb_qdtamtuyen_CheckedChanged(Nothing, Nothing)
                                        End If
                                    End If

                                    edt_hdld_sohd.Text = db.Rows(0)("SoQD").ToString().Trim()
                                    If (db.Rows(0)("NgayQD").ToString() <> "") Then
                                        dtpk_hdld_ngayky.Value = CType(db.Rows(0)("NgayQD").ToString(), DateTime)
                                    End If
                                    edt_hdld_nguoiky.Text = db.Rows(0)("NguoiQD").ToString().Trim()
                                    cb_hdld_chucvu.SelectedIndex = IIf(db.Rows(0)("IdCV_Nguoi_QD").ToString() <> "", CType(arr_Chucvu.IndexOf(db.Rows(0)("IdCV_Nguoi_QD").ToString()), Integer), 0)

                                    If db.Rows(0)("Tungay").ToString().Trim() <> "" Then
                                        dtpk_tg_tungay.Value = CType(db.Rows(0)("Tungay").ToString(), DateTime)
                                    End If
                                    If db.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                        dtpk_tg_denngay.Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime)
                                    End If
                                    'Load thông tin - Loại hình công việc
                                    cb_hinhthuc_cv.SelectedIndex = CType(db.Rows(0)("Loai").ToString().Trim(), Byte)
                                    'Load giờ trong điều khoản công việc
                                    If (db.Rows(0)("TuGio").ToString().Trim() <> "") Then
                                        edt_gio_batdau.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, CInt(db.Rows(0)("TuGio").ToString().Trim().Substring(0, 2)), CInt(db.Rows(0)("TuGio").ToString().Trim().Substring(3, 2)), 0)
                                    Else
                                        edt_gio_batdau.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
                                    End If
                                    If (db.Rows(0)("DenGio").ToString().Trim() <> "") Then
                                        edt_gio_ketthuc.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, CInt(db.Rows(0)("DenGio").ToString().Trim().Substring(0, 2)), CInt(db.Rows(0)("DenGio").ToString().Trim().Substring(3, 2)), 0)
                                    Else
                                        edt_gio_ketthuc.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
                                    End If
                                    'If (rb_DinhBien.Checked = True) Then
                                    '    ckb_choice_bhxh.Checked = IIf(db.Rows(0)("BHXH").ToString() = True, True, False)
                                    '    ckb_choice_bhyt.Checked = IIf(db.Rows(0)("BHYT").ToString() = True, True, False)
                                    'End If
                                    ckb_choice_bhxh.Checked = IIf(db.Rows(0)("BHXH").ToString() = True, True, False)
                                    ckb_choice_bhyt.Checked = IIf(db.Rows(0)("BHYT").ToString() = True, True, False)
                                    cb_ht_traluong.SelectedIndex = IIf(db.Rows(0)("IdHT_TraLuong").ToString() <> "", CType(arr_Ht_Traluong.IndexOf(db.Rows(0)("IdHT_TraLuong").ToString()), Integer), 0)
                                    'Load thông tin về Lương -> Từ bậc lương
                                    Dim _BangLuongId As Integer = 0
                                    Dim _NgachLuongId As Integer = 0
                                    Dim _BacLuongId As Integer = 0
                                    Dim _NghiDinhLuongId As Integer = 0
                                    _BacLuongId = IIf(db.Rows(0)("IdBacLuong").ToString() <> "0", CType(db.Rows(0)("IdBacLuong").ToString(), Integer), 0) ' Id bậc lương
                                    If (_BacLuongId > 0) Then
                                        rb_thangbang.Checked = True
                                        rb_thangbang_CheckedChanged(Nothing, Nothing)
                                        _NgachLuongId = _Labour.GetIdNgachLuong(_BacLuongId)
                                        If (_NgachLuongId > 0) Then
                                            _BangLuongId = _Labour.GetIdBangLuong(_NgachLuongId)
                                        End If
                                        If (_BangLuongId > 0) Then
                                            _NghiDinhLuongId = _Labour.GetIdNghiDinhLuong(_BangLuongId)
                                        End If
                                        cb_cdld_nghidinh.SelectedIndex = CType(arr_NghiDinhLuong.IndexOf(_NghiDinhLuongId.ToString()), Integer)
                                        cb_cdld_nghidinh_SelectedIndexChanged(Nothing, Nothing)
                                        cb_cdld_bangluong.SelectedIndex = CType(arr_BangLuong.IndexOf(_BangLuongId.ToString()), Integer)
                                        cb_cdld_bangluong_SelectedIndexChanged(Nothing, Nothing)
                                        cb_cdld_ngachluong.SelectedIndex = CType(arr_NgachLuong.IndexOf(_NgachLuongId.ToString()), Integer)
                                        cb_cdld_ngachluong_SelectedIndexChanged(Nothing, Nothing)
                                        cb_cdld_bacluong.SelectedIndex = CType(arr_BacLuong.IndexOf(db.Rows(0)("IdBacLuong").ToString()), Integer)
                                        cb_cdld_bacluong_SelectedIndexChanged(Nothing, Nothing)
                                        edt_cdld_tylehuong.Text = IIf(db.Rows(0)("TyleHuong").ToString().Trim() <> "", CType(db.Rows(0)("TyleHuong"), Double).ToString("N2"), "0")
                                    Else
                                        rb_trongoi.Checked = True
                                        rb_trongoi_CheckedChanged(Nothing, Nothing)
                                        edt_tienluong.Text = db.Rows(0)("Tien_Luong").ToString().Trim()
                                    End If
                                    cb_cdld_chuyenmon.SelectedIndex = IIf(db.Rows(0)("IdChuyenMon").ToString() <> "", CType(arr_ChuyenMon.IndexOf(db.Rows(0)("IdChuyenMon").ToString()), Integer), 0)
                                    edt_cdld_cv_damnhan.Text = db.Rows(0)("CongViec").ToString().Trim()
                                End If
                            End If
                        End Using
                    End If
                End If
                ActiveControl = edt_macb
            Case Else
        End Select
    End Sub

    ''' <summary>
    ''' Hàm trả về đường dẫn của thư mục lưu ảnh thẻ
    ''' </summary>
    ''' <param name="branchCode">Mã hiệu chi nhánh</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPath(ByVal branchCode As String) As String
        Dim _Result As String = ""
        Dim strRoot As String = ""
        strRoot = System.IO.Path.Combine(Application.StartupPath, "ANH_HOSO")
        If Not (Globals.FileDao.IsDirectory(strRoot)) Then 'Kiểm tra xem đã có thư mục gốc ANH_HOSO chưa
            Globals.FileDao.MakeFolder(Application.StartupPath, "ANH_HOSO")
        End If
        strRoot = System.IO.Path.Combine(strRoot, "HSCB_TAPSU")
        If Not (Globals.FileDao.IsDirectory(strRoot)) Then 'Kiểm tra xem đã có thư mục gốc ANH_HOSO chưa
            Globals.FileDao.MakeFolder(System.IO.Path.Combine(Application.StartupPath, "ANH_HOSO"), "HSCB_TAPSU")
        End If
        _Result = strRoot

        strRoot = System.IO.Path.Combine(strRoot, branchCode)
        If Not (Globals.FileDao.IsDirectory(strRoot)) Then 'Kiểm tra xem đã có thư mục gốc - tên chi nhánh lưu ảnh chưa
            Globals.FileDao.MakeFolder(_Result, branchCode)
        End If
        'Trả về đường dẫn thư mục cần lưu ảnh thẻ chuẩn bị được tạo
        _Result = System.IO.Path.Combine(_Result, branchCode)
        Return _Result + "\"
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra dữ liệu hợp lệ trước khi cập nhật
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function Valid_Data() As Boolean
        Dim strSQL As String = ""
        If (edt_macb.Text.Trim() = "") Then
            MessageBox.Show("Mã cán bộ trong hồ sơ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_macb
            Return False
        End If

        'Băt điều kiện - Trùng mã cán bộ tập sự => Từ 2024 bổ sung kiểm tra trùng với mã CB cả HS_CanBo
        If (edt_macb.Text.Trim() <> "") Then
            If (_IdCanBo <> "") Then
                strSQL = String.Format("Select Id IdCanBo,MaCB,HoTen,IdChiNhanh,NgaySinh,CMT_So From HSCB_TS Where MaCB = '{0}' And Id <> '{1}' Union Select IdCanBo,MaCB,HoTen,IdDonVi IdChiNhanh,NgaySinh,CMT_So From HS_CanBo Where MaCB = '{0}' ", Globals.Find_Replace(edt_macb.Text.Trim().ToString()), _IdCanBo)
            Else
                strSQL = String.Format("Select Id IdCanBo,MaCB,HoTen,IdChiNhanh,NgaySinh,CMT_So From HSCB_TS Where MaCB = '{0}' Union Select IdCanBo,MaCB,HoTen,IdDonVi IdChiNhanh,NgaySinh,CMT_So From HS_CanBo Where MaCB = '{0}' ", Globals.Find_Replace(edt_macb.Text.Trim().ToString()))
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        'Lấy dẫy số mã cán bộ theo quy định để cho người dùng cập nhật
                        Dim sCodeByBranch As String = ""
                        If DONVI = "000100" Or DONVI = "000101" Or DONVI = "000196" Or DONVI = "000197" Then
                            strSQL = String.Format("Select Top 1 * From ChiNhanh Where Ma_So = '{0}' Order By Id", DONVI)
                        Else
                            strSQL = String.Format("Select Top 1 * From ChiNhanh Where Ma_So Like '{0}%' And Id_Goc In (0,1) Order By Status,Id Desc", DONVI.Substring(0, 4))
                        End If
                        Dim db_cn As DataTable = Nothing
                        db_cn = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db_cn Is Nothing) Then
                            If db_cn.Rows.Count > 0 Then
                                sCodeByBranch = db_cn.Rows(0)("MaCB_Begin").ToString().Trim() + " => " + db_cn.Rows(0)("MaCB_End").ToString().Trim()
                            End If
                        End If
                        MessageBox.Show(String.Format("Mã cán bộ: '{0}' đã tồn tại, vui lòng nhập mã cán bộ bằng một số khác trong dẫy số '{1}' theo đơn vị '{2}'. Vui lòng kiểm tra lại!", edt_macb.Text.Trim().ToString(), sCodeByBranch, BrandNameByUserLogin), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_macb
                        Return False
                    End If
                End If
            End Using
        End If

        If (edt_hoten.Text.Trim() = "") Then
            MessageBox.Show("Họ tên cán bộ trong hồ sơ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_hoten
            Return False
        End If
        If (rb_nam.Checked = False And rb_nu.Checked = False) Then
            MessageBox.Show("Bạn chưa lựa chọn giới tính của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            rb_nam.Focus()
            Return False
        End If

        If (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn đơn vị công tác của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_donvi
            Return False
        End If
        If (cb_quoctich.SelectedIndex <= 0 And cb_quoctich.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn quốc tịch của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_quoctich
            Return False
        End If
        If (cb_dantoc.SelectedIndex <= 0 And cb_dantoc.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn dân tộc của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_dantoc
            Return False
        End If
        If (cb_tongiao.SelectedIndex <= 0 And cb_tongiao.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn tôn giáo của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_tongiao
            Return False
        End If
        If (edt_cmt_so.Text.Trim() = "") Then
            MessageBox.Show("Số chứng minh thư của cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_cmt_so
            Return False
        End If
        If (edt_cmt_noicap.Text.Trim() = "") Then
            MessageBox.Show("Nơi cấp chứng minh thư của cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_cmt_noicap
            Return False
        End If
        'Bắt điều kiện nhập trùng số - Chứng minh thư nhân dân
        If (edt_cmt_so.Text.Trim() <> "") Then
            If (_IdCanBo <> "") Then
                strSQL = String.Format("Select * From HSCB_TS Where CMT_So = '{0}' and Id <> '{1}'", Globals.Find_Replace(edt_cmt_so.Text.Trim().ToString()), _IdCanBo)
            Else
                strSQL = String.Format("Select * From HSCB_TS Where CMT_So = '{0}'", Globals.Find_Replace(edt_cmt_so.Text.Trim().ToString()))
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        MessageBox.Show(String.Format("Số chứng minh thư: '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_cmt_so.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_cmt_so
                        Return False
                    End If
                End If
            End Using
        End If

        'Thực hiện kiểm tra tuổi làm chứng minh thư có phải là 16 tuổi không ?
        Dim ngay_dk As DateTime = dtpk_ngaysinh.Value.AddYears(14)
        If (dtpk_cmt_ngaycap.Value <= ngay_dk) Then
            MessageBox.Show("Ngày cấp chứng minh thư không hợp lệ!" + vbCrLf + "Lưu ý: Theo quy định tại Điều 3, Nghị Định Số: 05/1999/NĐ-CP thì Công dân Việt Nam từ đủ 14 tuổi trở lên mới được làm chứng minh thư nhân dân", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = dtpk_cmt_ngaycap
            Return False
        End If

        'Kiểm tra điều kiện thông tin về địa chỉ Nơi sinh
        If (cb_ns_tinh.SelectedIndex <= 0 And cb_ns_tinh.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nơi sinh tỉnh (thành phố) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_ns_tinh
            Return False
        End If
        If (cb_ns_xa.SelectedIndex <= 0 And cb_ns_xa.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nơi sinh xã (phường) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_ns_xa
            Return False
        End If
        If (cb_ns_thon.SelectedIndex <= 0 And cb_ns_thon.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nơi sinh xã (phường) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_ns_xa
            Return False
        End If
        'If (edt_ns_diachi.Text.Trim() = "") Then
        '    MessageBox.Show("Địa chỉ nơi sinh không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '    tctrl_main.SelectedIndex = 0
        '    ActiveControl = edt_ns_diachi
        '    Return False
        'End If


        'Kiểm tra điều kiện thông tin về địa chỉ Nguyên quán
        If (cb_nq_tinh.SelectedIndex <= 0 And cb_nq_tinh.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nguyên quán tỉnh (thành phố) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_nq_tinh
            Return False
        End If
        If (cb_nq_xa.SelectedIndex <= 0 And cb_nq_xa.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nguyên quán xã (phường) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_nq_xa
            Return False
        End If
        If (cb_nq_thon.SelectedIndex <= 0 And cb_nq_thon.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nguyên quán thôn của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_nq_thon
            Return False
        End If
        'If (edt_nq_diachi.Text.Trim() = "") Then
        '    MessageBox.Show("Địa chỉ nguyên quán không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '    tctrl_main.SelectedIndex = 0
        '    ActiveControl = edt_nq_diachi
        '    Return False
        'End If

        'Kiểm tra điều kiện thông tin về địa chỉ Thường trú
        If (cb_tt_tinh.SelectedIndex <= 0 And cb_tt_tinh.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn thường trú tỉnh (thành phố) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_tt_tinh
            Return False
        End If
        If (cb_tt_xa.SelectedIndex <= 0 And cb_tt_xa.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn thường trú xã (phường) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_tt_xa
            Return False
        End If
        If (cb_tt_thon.SelectedIndex <= 0 And cb_tt_thon.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn thường trú thôn của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_tt_thon
            Return False
        End If
        If (edt_tt_diachi.Text.Trim() = "") Then
            MessageBox.Show("Địa chỉ thường trú không được để trống. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = edt_tt_diachi
            Return False
        End If

        ' Nếu địa chỉ e-mail của cán bộ mà nhập thì kiểm tra có hợp lệ không
        If (edt_email.Text.Trim() <> "") Then
            If (Not Globals.IsEmail(edt_email.Text.Trim().ToString())) Then
                MessageBox.Show("Địa chỉ e-mail của cán bộ không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 0
                ActiveControl = edt_email
                Return False
            End If
        End If
        If (cb_phongban.SelectedIndex <= 0 And cb_phongban.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn phòng ban công tác của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_phongban
            Return False
        End If
        'Nếu chọn cập nhật cách mạng -> Bắt các thông tin cần thiết về việc tham gia Cách mạng
        If (ckb_cachmang.Checked = True) Then
            If (dtpk_cm_ngay_tg.Checked = False) Then
                MessageBox.Show("Bạn chưa chọn ngày tham gia cách mạng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = dtpk_cm_ngay_tg
                Return False
            End If
            If (edt_cm_tochuc.Text.Trim() = "") Then
                MessageBox.Show("Tổ chức cách mạng mà cán bộ tham gia không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = edt_cm_tochuc
                Return False
            End If
        End If
        'Nếu chọn cập nhật Đảng viên -> Bắt các thông tin cần thiết về Đảng của cán bộ
        If (ckb_dangvien.Checked = True) Then
            If (dtpk_dv_ngayvao.Checked = False) Then
                MessageBox.Show("Bạn chưa chọn ngày vào đảng của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = dtpk_dv_ngayvao
                Return False
            End If
            If (dtpk_dv_ngay_chinhthuc.Checked = False) Then
                MessageBox.Show("Bạn chưa chọn ngày vào đảng chính thức của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = dtpk_dv_ngay_chinhthuc
                Return False
            End If
            If (edt_dv_noikn.Text.Trim() = "") Then
                MessageBox.Show("Nơi kết nạp đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = edt_dv_noikn
                Return False
            End If
            If (edt_dv_sothe.Text.Trim() = "") Then
                MessageBox.Show("Số thẻ đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = edt_dv_sothe
                Return False
            End If
            If (edt_dv_nguoi_gt.Text.Trim() = "") Then
                MessageBox.Show("Người giới thiệu vào đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = edt_dv_nguoi_gt
                Return False
            End If
        End If

        'Băt điều kiện - Trùng mã số thuế cá nhân
        If (edt_ms_thue.Text.Trim() <> "") Then
            If (_IdCanBo <> "") Then
                strSQL = String.Format("Select * From HSCB_TS Where MaSoThue = '{0}' and Id <> '{1}'", Globals.Find_Replace(edt_ms_thue.Text.Trim().ToString()), _IdCanBo)
            Else
                strSQL = String.Format("Select * From HSCB_TS Where MaSoThue = '{0}'", Globals.Find_Replace(edt_ms_thue.Text.Trim().ToString()))
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        MessageBox.Show(String.Format("Mã số thuế: {0} của cán bộ đã tồn tại. Vui lòng kiểm tra lại!", edt_ms_thue.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_main.SelectedIndex = 1
                        ActiveControl = edt_ms_thue
                        Return False
                    End If
                End If
            End Using
        End If

        'Nếu chọn cập nhật Thông tin BHXH -> Bắt các thông tin cần thiết về BHXH
        If (ckb_bhxh.Checked = True) Then
            If (edt_bhxh_soso.Text.Trim() = "") Then
                MessageBox.Show("Số sổ bảo hiểm xã hội không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = edt_bhxh_soso
                Return False
            End If
            If (dtpk_bhxh_ngaylam.Checked = False) Then
                MessageBox.Show("Bạn chưa chọn ngày làm sổ bảo hiểm xã hội!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = dtpk_bhxh_ngaylam
                Return False
            End If
            If (dtpk_bhxh_ngaydong.Checked = False) Then
                MessageBox.Show("Bạn chưa chọn ngày đóng bảo hiểm xã hội!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = dtpk_bhxh_ngaydong
                Return False
            End If
            If (edt_bhxh_noilam.Text.Trim() = "") Then
                MessageBox.Show("Nơi làm sổ bảo hiểm xã hội không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 1
                ActiveControl = edt_bhxh_noilam
                Return False
            End If
        End If

        'Bắt các thông tin còn lại - Tab hợp đồng lao động
        If (edt_hdld_sohd.Text.Trim() = "") Then
            MessageBox.Show("Số hợp đồng lao động không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = edt_hdld_sohd
            Return False
        End If
        If (edt_hdld_sohd.Text.Trim() <> "") Then
            If (_IdHdld = "") Then
                strSQL = String.Format("Select * From HSCB_TS_HDLD Where SoQD = '{0}'", Globals.Find_Replace(edt_hdld_sohd.Text.Trim().ToString()))
            Else
                strSQL = String.Format("Select * From HSCB_TS_HDLD Where SoQD = '{0}' And IdCBTS_HDLD <> '{1}'", Globals.Find_Replace(edt_hdld_sohd.Text.Trim().ToString()), _IdHdld)
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        MessageBox.Show(String.Format("Số hợp đồng: '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_hdld_sohd.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_main.SelectedIndex = 2
                        ActiveControl = edt_hdld_sohd
                        Return False
                    End If
                End If
            End Using
        End If
        If (edt_hdld_nguoiky.Text.Trim() = "") Then
            MessageBox.Show("Người ký hợp đồng lao động không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = edt_hdld_nguoiky
            Return False
        End If
        If (cb_hdld_chucvu.SelectedIndex <= 0 And cb_hdld_chucvu.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn chức vụ người ký hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = cb_hdld_chucvu
            Return False
        End If

        If dtpk_tg_tungay.Value >= dtpk_tg_denngay.Value Then
            MessageBox.Show("Ngày bắt đầu không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = dtpk_tg_tungay
            Return False
        End If

        If (edt_gio_batdau.Value.Hour >= edt_gio_ketthuc.Value.Hour) Then
            MessageBox.Show("Giờ làm việc trong hợp đồng không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = edt_gio_batdau
            Return False
        End If
        If (cb_hinhthuc_cv.SelectedIndex <= 0 And cb_hinhthuc_cv.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn loại hình công việc mà cán bộ tập sự tham gia!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = cb_hinhthuc_cv
            Return False
        End If
        If (cb_ht_traluong.SelectedIndex <= 0 And cb_ht_traluong.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn hình thức trả lương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = cb_ht_traluong
            Return False
        End If

        If (rb_thangbang.Checked = True) Then
            If (cb_cdld_bacluong.SelectedIndex <= 0 And cb_cdld_bacluong.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn bậc lương!" + vbCrLf + "Lưu ý: Muốn chọn dữ liệu bậc lương thì trước hết bạn phải chọn dữ liệu Bảng lương rồi đến Ngạch lương.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 2
                ActiveControl = cb_cdld_bacluong
                Return False
            End If
            If (edt_cdld_tylehuong.Text.Trim() = "" Or CType(edt_cdld_tylehuong.Text.Trim(), Double) <= 0) Then
                MessageBox.Show("Tỷ lệ hưởng lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 2
                ActiveControl = edt_cdld_tylehuong
                Return False
            End If
            If (edt_cdld_tylehuong.Text.Trim() <> "") Then
                If (CType(edt_cdld_tylehuong.Text.Trim(), Double) > 300) Then
                    MessageBox.Show("Tỷ lệ hưởng lương không hợp lệ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_main.SelectedIndex = 2
                    ActiveControl = edt_cdld_tylehuong
                    Return False
                End If
            End If
        End If
        If (rb_trongoi.Checked = True) Then
            If (edt_tienluong.Text.Trim() = "" Or CType(MoneyValue(edt_tienluong.Text.Trim()), Double) <= 0) Then
                MessageBox.Show("Số tiền lương được hưởng của cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                tctrl_main.SelectedIndex = 2
                ActiveControl = edt_tienluong
                Return False
            End If
        End If
        If (cb_cdld_chuyenmon.SelectedIndex <= 0 And cb_cdld_chuyenmon.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn chuyên môn đảm nhiệm của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            tctrl_main.SelectedIndex = 2
            ActiveControl = cb_cdld_chuyenmon
            Return False
        End If


        'Xét khoảng thời gian trong hợp đồng hoặc quyết định Tạm tuyển trùng hoặc đan xen lẫn nhau.
        If (_FlagEvent = 1) Then
            strSQL = String.Format("Select * From HSCB_TS_HDLD Where IdCanbo = '{0}' Order by TuNgay Asc", _IdCanBo, _IdHdld)
        ElseIf (_FlagEvent = 2) Then
            strSQL = String.Format("Select * From HSCB_TS_HDLD Where IdCanbo = '{0}' And IdCBTS_HDLD <> '{1}' Order by TuNgay Asc", _IdCanBo, _IdHdld)
        End If
        Dim Exist As Boolean = False
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                        Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                        If ((_TuNgay <= dtpk_tg_tungay.Value And dtpk_tg_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_tg_denngay.Value And dtpk_tg_denngay.Value <= _DenNgay) Or (dtpk_tg_tungay.Value <= _TuNgay And dtpk_tg_denngay.Value >= _DenNgay)) Then
                            Exist = True
                            Exit For
                        End If
                    Next
                    If (Exist = True) Then
                        MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hợp đồng không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian của hợp đồng lao động hoặc quyết định tạm tuyển không thể trùng hoặc đan xen lẫn nhau.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_main.SelectedIndex = 2
                        ActiveControl = dtpk_tg_tungay
                        Return False
                    End If
                End If
            End If
        End Using
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Trả về chỉ số index của nhóm máu khi biết chuỗi nhóm máu
    ''' </summary>
    ''' <returns>1: A. 2: B. 3: AB. 4: O</returns>
    ''' <remarks></remarks>
    Private Function GetIndex_NhomMau(ByVal _Nhommau As String) As Byte
        Select Case _Nhommau
            Case "A"
                Return 1
            Case "B"
                Return 2
            Case "AB"
                Return 3
            Case "O"
                Return 4
            Case Else
                Return 0
        End Select
    End Function
#End Region

#Region "---> Events: Các sự kiện liên quan <---"
    Private Sub frmCN_CanBoTS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Thực hiện fill dữ liệu vào Các ComBoBox
        Dim strSQL As String = ""
        If TRUCTHUOC = 1 Then
            strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And id_goc IN (0,1)"
        Else
            strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And ma_so='" & DONVI.Trim & "'"
        End If
        'Fill dữ liệu combobox đơn vị công tác

        arr_Donvi.Clear()
        cb_donvi.Items.Clear()
        arr_Donvi = _Globals.Bind_ComBoBox(cb_donvi, strSQL, "---Đơn vị (Chi nhánh)---")
        'Bind dữ liệu vào combobox - Quốc tịch
        arr_Quoctich.Clear()
        cb_quoctich.Items.Clear()
        arr_Quoctich = _Globals.Bind_ComBoBox(cb_quoctich, clsHT_DanhMuc.Sql_Quocgia, "---Quốc tịch---")
        'Bind dữ liệu vào combobox - Dân tộc
        arr_Dantoc.Clear()
        cb_dantoc.Items.Clear()
        arr_Dantoc = _Globals.Bind_ComBoBox(cb_dantoc, clsHT_DanhMuc.Sql_Dantoc, "---Dân tộc---")
        'Bind dữ liệu vào combobox - Tôn giáo
        arr_Tongiao.Clear()
        cb_tongiao.Items.Clear()
        arr_Tongiao = _Globals.Bind_ComBoBox(cb_tongiao, clsHT_DanhMuc.Sql_Tongiao, "---Tôn giáo---")

        'Fill nơi sinh - Tỉnh thành phố
        arr_Ns_Tinh.Clear()
        cb_ns_tinh.Items.Clear()
        arr_Ns_Tinh = _Globals.Bind_ComBoBox(cb_ns_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

        'Fill nguyên quán - Tỉnh thành phố
        arr_Nq_Tinh.Clear()
        cb_nq_tinh.Items.Clear()
        arr_Nq_Tinh = _Globals.Bind_ComBoBox(cb_nq_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

        'Fill Thường trú - Tỉnh thành phố
        arr_TT_Tinh.Clear()
        cb_tt_tinh.Items.Clear()
        arr_TT_Tinh = _Globals.Bind_ComBoBox(cb_tt_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

        'Fill Tạm trú - Tỉnh thành phố
        arr_TTr_Tinh.Clear()
        cb_ttr_tinh.Items.Clear()
        arr_TTr_Tinh = _Globals.Bind_ComBoBox(cb_ttr_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

        'Fill dữ liệu Thành phần gia đình cán bộ
        arr_Tp_Giadinh.Clear()
        cb_tp_giadinh.Items.Clear()
        arr_Tp_Giadinh = _Globals.Bind_ComBoBox(cb_tp_giadinh, clsHT_DanhMuc.Sql_Tpgd, "---Thành phần gia đình---")

        'Fill dữ liệu Trình độ Văn hoá cán bộ
        arr_Td_Vanhoa.Clear()
        cb_td_vanhoa.Items.Clear()
        arr_Td_Vanhoa = _Globals.Bind_ComBoBox(cb_td_vanhoa, clsHT_DanhMuc.Sql_Tdvh, "---Trình độ văn hoá---")

        'Fill dữ liệu Trình độ Chính trị cán bộ
        arr_Td_Chinhtri.Clear()
        cb_td_chinhtri.Items.Clear()
        arr_Td_Chinhtri = _Globals.Bind_ComBoBox(cb_td_chinhtri, clsHT_DanhMuc.Sql_Tdct, "---Trình độ trính trị---")

        'Fill dữ liệu Học hàm cán bộ
        arr_Hocham.Clear()
        cb_hocham.Items.Clear()
        arr_Hocham = _Globals.Bind_ComBoBox(cb_hocham, clsHT_DanhMuc.Sql_Hocham, "---Học hàm---")

        'Fill dữ liệu Học vị cán bộ
        arr_Hocvi.Clear()
        cb_hocvi.Items.Clear()
        arr_Hocvi = _Globals.Bind_ComBoBox(cb_hocvi, clsHT_DanhMuc.Sql_Hocvi, "---Học vị---")

        arr_TrdChMon.Clear()
        cb_trd_chmon.Items.Clear()
        arr_TrdChMon = _Globals.Bind_ComBoBox(cb_trd_chmon, clsHT_DanhMuc.Sql_TrdChuyenMon, "---Trình độ chuyên môn---")

        strSQL = "Select id,ten_goi From DanhMuc Where id_goc = 38 And Status = 1 And ma_so IN ('3803','3812','3813','3814') Order by Ma_so Asc"
        arr_TrdNgoaiNgu.Clear()
        cb_trd_ngngu.Items.Clear()
        arr_TrdNgoaiNgu = _Globals.Bind_ComBoBox(cb_trd_ngngu, strSQL, "---Trình độ ngoại ngữ---")

        arr_TrdTinHoc.Clear()
        cb_trd_tinhoc.Items.Clear()
        arr_TrdTinHoc = _Globals.Bind_ComBoBox(cb_trd_tinhoc, strSQL, "---Trình độ tin học---")

        'Fill dữ liệu Chuyên ngành đào tạo của Cán bộ
        arr_ChNganh.Clear()
        cb_chnganh.Items.Clear()
        arr_ChNganh = _Globals.Bind_ComBoBox(cb_chnganh, clsHT_DanhMuc.Sql_ChNganh, "---Chuyên ngành đào tạo---")

        'Fill dữ liệu Ưu tiên gia đình cán bộ
        arr_Ut_Giadinh.Clear()
        cb_ut_giadinh.Items.Clear()
        arr_Ut_Giadinh = _Globals.Bind_ComBoBox(cb_ut_giadinh, clsHT_DanhMuc.Sql_Ut_giadinh, "---Ưu tiên gia đình---")

        'Fill dữ liệu Ưu tiên bản thân cán bộ
        clb_ut_banthan.Items.Clear()
        arr_Ut_Banthan.Clear()
        arr_Ut_Banthan = _Globals.FillData_CheckedListBox(clb_ut_banthan, clsHT_DanhMuc.Sql_Ut_banthan)

        'Fill data combobox - Tình trạng hôn nhân
        ARL_TT_HonNhan.Clear()
        cb_tt_honnhan.Items.Clear()
        ARL_TT_HonNhan = _Globals.Bind_ComBoBox(cb_tt_honnhan, clsHT_DanhMuc.Sql_TTrangHonNhan, "---Tình trạng hôn nhân---")

        'Fill dữ liệu Hình thức trả lương cán bộ
        arr_Ht_Traluong.Clear()
        cb_ht_traluong.Items.Clear()
        arr_Ht_Traluong = _Globals.Bind_ComBoBox(cb_ht_traluong, clsHT_DanhMuc.Sql_HtTrLuong, "---Hình thức trả lương---")

        'Fill dữ liệu Nghị định lương cán bộ
        strSQL = "Select IdNDLuong, TenND From NghiDinhLuong Where Status = 1"
        arr_NghiDinhLuong.Clear()
        cb_cdld_nghidinh.Items.Clear()
        arr_NghiDinhLuong = _Globals.Bind_ComBoBox(cb_cdld_nghidinh, strSQL, "---Nghị định lương---")

        'Chức vụ người ký quyết định trong hợp đồng lao động
        arr_Chucvu.Clear()
        cb_hdld_chucvu.Items.Clear()
        arr_Chucvu = _Globals.Bind_ComBoBox(cb_hdld_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ---")

        arr_ChuyenMon.Clear()
        cb_cdld_chuyenmon.Items.Clear()
        arr_ChuyenMon = _Globals.Bind_ComBoBox(cb_cdld_chuyenmon, clsHT_DanhMuc.Sql_Chuyenmon, "---Chuyên môn---")
        'Thực hiện reset các controls còn lại
        Load_Infors()

        If _IdCanBo <> "" Then
            If checkRight_CreateRecord(_IdCanBo) Then
                btn_save.Enabled = True
            Else
                btn_save.Enabled = False
            End If
        End If
        'picbx_main.Image.Tag


    End Sub

    Private Sub btn_chonanh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_chonanh.Click
        Dim ofd_file As OpenFileDialog = New OpenFileDialog
        ofd_file.Filter = "All File (*.*)|*.*|Document File(PDF File(*.pdf)|*.pdf|Document File (*.doc)|*.doc|Web File(*.html)|*.html|Web File(*.htm)|*.htm"
        ofd_file.FilterIndex = 1
        ofd_file.Multiselect = False
        If (ofd_file.ShowDialog() <> Windows.Forms.DialogResult.OK) Then
            Return
        End If
        'Giải phóng hoàn toàn dữ liệu ảnh cũ nếu có
        If (picbx_main.Image IsNot Nothing) Then
            picbx_main.Image.Dispose()
            picbx_main.Image = Nothing
        End If

        strPath = ""
        strPath = ofd_file.FileName.ToString()
        If (System.IO.File.Exists(strPath)) Then
            If (strPath.IndexOf(".") > 0) Then
                Dim k As Integer = strPath.LastIndexOf(".") + 1
                Dim file_end As String = Globals.FileDao.GetFormatSize(strPath).ToLower()
                If (file_end = ".jpg" Or file_end = ".jpeg" Or file_end = ".bmp" Or file_end = ".gif" Or file_end = ".dib" Or file_end = ".png") Then
                    picbx_main.Image = Image.FromFile(strPath)
                    picbx_main.SizeMode = PictureBoxSizeMode.StretchImage
                Else
                    MessageBox.Show("Bạn phải chọn đúng file ảnh!", "Chọn ảnh hồ sơ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End If
            End If
        End If
    End Sub

    Private Sub btn_xoaanh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_xoaanh.Click
        If (picbx_main.Image Is Nothing) Then
            Return
        End If
        strPath = ""
        If (_FlagEvent = 2) Then
            flagChoice = True
        End If
        picbx_main.Image.Dispose()
        picbx_main.Image = Nothing
        Dim strdefault As String = System.IO.Path.Combine(Application.StartupPath, "ANH_HOSO") + "//"
        strdefault = String.Format("{0}{1}", strdefault, "default.jpg")
        If (System.IO.File.Exists(strdefault)) Then
            picbx_main.Image = Image.FromFile(strdefault)
            picbx_main.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub

    Private Sub frmCN_CanBoTS_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        'Giải phóng dữ liệu ảnh nếu có
        If (picbx_main.Image IsNot Nothing) Then
            picbx_main.Image.Dispose()
            picbx_main.Image = Nothing
            Dim strdefault As String = System.IO.Path.Combine(Application.StartupPath, "ANH_HOSO") + "//"
            strdefault = String.Format("{0}{1}", strdefault, "default.jpg")
            If (System.IO.File.Exists(strdefault)) Then
                picbx_main.Image = Image.FromFile(strdefault)
                picbx_main.SizeMode = PictureBoxSizeMode.StretchImage
            End If
        End If

        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        Load_Infors()
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Try
            Select Case _FlagEvent
                Case 1
                    If (Valid_Data()) Then
                        Dim obj_document As clsHS_CanBo.HSCB_TS = New clsHS_CanBo.HSCB_TS()
                        obj_document.MaCB = Globals.Find_Replace(edt_macb.Text.ToString().Trim())
                        obj_document.HoTen = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hoten.Text.ToString().Trim()))
                        obj_document.TenThuongGoi = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_ten_tg.Text.ToString().Trim()))
                        obj_document.BiDanh = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_bidanh.Text.ToString().Trim()))
                        obj_document.GioiTinh = CType(IIf(rb_nam.Checked = True, 0, 1), Byte)
                        obj_document.NgaySinh = dtpk_ngaysinh.Value
                        obj_document.IdChiNhanh = CType(IIf(arr_Donvi.Count > 0, arr_Donvi(cb_donvi.SelectedIndex), "0"), Integer)
                        obj_document.IdPhongBan = IIf(arr_Phongban.Count > 0, arr_Phongban(cb_phongban.SelectedIndex), "0").ToString()
                        obj_document.IdQuocTich = CType(IIf(arr_Quoctich.Count > 0, arr_Quoctich(cb_quoctich.SelectedIndex), "0"), Integer)
                        obj_document.IdDanToc = CType(IIf(arr_Dantoc.Count > 0, arr_Dantoc(cb_dantoc.SelectedIndex), "0"), Integer)
                        obj_document.IdTonGiao = CType(IIf(arr_Tongiao.Count > 0, arr_Tongiao(cb_tongiao.SelectedIndex), "0"), Integer)
                        obj_document.CMT_So = Globals.Find_Replace(edt_cmt_so.Text.ToString().Trim())
                        obj_document.CMT_NgayCap = dtpk_cmt_ngaycap.Value
                        obj_document.CMT_NoiCap = Globals.Find_Replace(edt_cmt_noicap.Text.Trim.ToString())
                        obj_document.IdNS_Tinh = CType(IIf(arr_Ns_Tinh.Count > 0, arr_Ns_Tinh(cb_ns_tinh.SelectedIndex), "0"), Integer)
                        If (cb_ns_huyen.Items.Count <> 0) Then
                            obj_document.IdNS_Huyen = CType(IIf(arr_Ns_Huyen.Count > 0, arr_Ns_Huyen(cb_ns_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdNS_Huyen = 0
                        End If
                        If (cb_ns_xa.Items.Count <> 0) Then
                            obj_document.IdNS_xa = CType(IIf(arr_Ns_Xa.Count > 0, arr_Ns_Xa(cb_ns_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNS_Xa = 0
                        End If
                        If (cb_ns_thon.Items.Count <> 0) Then
                            obj_document.IdNS_Thon = CType(IIf(arr_Ns_Thon.Count > 0, arr_Ns_Thon(cb_ns_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNS_Thon = 0
                        End If
                        obj_document.NS_DiaChi = Globals.Find_Replace(edt_ns_diachi.Text.Trim.ToString())

                        obj_document.IdNQ_Tinh = CType(IIf(arr_Nq_Tinh.Count > 0, arr_Nq_Tinh(cb_nq_tinh.SelectedIndex), "0"), Integer)
                        If cb_nq_huyen.Items.Count <> 0 Then
                            obj_document.IdNQ_Huyen = CType(IIf(arr_Nq_Huyen.Count > 0, arr_Nq_Huyen(cb_nq_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdNQ_Huyen = 0
                        End If
                        If cb_nq_xa.Items.Count <> 0 Then
                            obj_document.IdNQ_Xa = CType(IIf(arr_Nq_Xa.Count > 0, arr_Nq_Xa(cb_nq_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNQ_Xa = 0
                        End If
                        If cb_nq_thon.Items.Count <> 0 Then
                            obj_document.IdNQ_Thon = CType(IIf(arr_Nq_Thon.Count > 0, arr_Nq_Thon(cb_nq_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNQ_Thon = 0
                        End If
                        obj_document.NQ_DiaChi = Globals.Find_Replace(edt_nq_diachi.Text.Trim.ToString())

                        obj_document.IdThT_Tinh = CType(IIf(arr_TT_Tinh.Count > 0, arr_TT_Tinh(cb_tt_tinh.SelectedIndex), "0"), Integer)
                        If (cb_tt_huyen.Items.Count <> 0) Then
                            obj_document.IdThT_Huyen = CType(IIf(arr_TT_Huyen.Count > 0, arr_TT_Huyen(cb_tt_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdThT_Huyen = 0
                        End If
                        If (cb_tt_xa.Items.Count <> 0) Then
                            obj_document.IdThT_Xa = CType(IIf(arr_TT_Xa.Count > 0, arr_TT_Xa(cb_tt_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdThT_Xa = 0
                        End If
                        If (cb_tt_thon.Items.Count <> 0) Then
                            obj_document.IdThT_Thon = CType(IIf(arr_TT_Thon.Count > 0, arr_TT_Thon(cb_tt_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdThT_Thon = 0
                        End If
                        obj_document.ThT_DiaChi = Globals.Find_Replace(edt_tt_diachi.Text.Trim.ToString())
                        obj_document.ThT_DienThoai = Globals.Find_Replace(edt_tt_dienthoai.Text.Trim.ToString())

                        obj_document.IdTTr_Tinh = CType(IIf(arr_TTr_Tinh.Count > 0, arr_TTr_Tinh(cb_ttr_tinh.SelectedIndex), "0"), Integer)
                        If cb_ttr_huyen.Items.Count <> 0 Then
                            obj_document.IdTTr_Huyen = CType(IIf(arr_TTr_Huyen.Count > 0, arr_TTr_Huyen(cb_ttr_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdTTr_Huyen = 0
                        End If
                        If cb_ttr_xa.Items.Count <> 0 Then
                            obj_document.IdTTr_Xa = CType(IIf(arr_TTr_Xa.Count > 0, arr_TTr_Xa(cb_ttr_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdTTr_Xa = 0
                        End If
                        If cb_ttr_thon.Items.Count <> 0 Then
                            obj_document.IdTTr_Thon = CType(IIf(arr_TTr_Thon.Count > 0, arr_TTr_Thon(cb_ttr_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdTTr_Thon = 0
                        End If
                        obj_document.TTr_DiaChi = Globals.Find_Replace(edt_ttr_diachi.Text.Trim.ToString())
                        obj_document.TTr_DienThoai = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_ttr_dienthoai.Text.Trim.ToString()))
                        obj_document.DienThoai_CQ = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_dt_coquan.Text.Trim.ToString()))
                        obj_document.DienThoai_DD = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_dt_didong.Text.Trim.ToString()))
                        obj_document.DienThoai_NR = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_dt_nharieng.Text.Trim.ToString()))
                        obj_document.SoFax = Globals.Find_Replace(edt_sofax.Text.Trim.ToString())
                        obj_document.Email = Globals.Find_Replace(edt_email.Text.Trim.ToString())
                        obj_document.NhomMau = IIf(cb_nhommau.SelectedIndex > 0, cb_nhommau.SelectedItem.ToString(), "")
                        obj_document.IdThanhPhan_GD = CType(IIf(arr_Tp_Giadinh.Count > 0, arr_Tp_Giadinh(cb_tp_giadinh.SelectedIndex), "0"), Integer)
                        obj_document.IdUT_BT = _Globals.GetItems_CheckedListBox(clb_ut_banthan, arr_Ut_Banthan)
                        obj_document.IdUT_GD = CType(IIf(arr_Ut_Giadinh.Count > 0, arr_Ut_Giadinh(cb_ut_giadinh.SelectedIndex), "0"), Integer)
                        obj_document.IdTrinhDoVH = CType(IIf(arr_Td_Vanhoa.Count > 0, arr_Td_Vanhoa(cb_td_vanhoa.SelectedIndex), "0"), Integer)
                        obj_document.IdTrinhDoCT = CType(IIf(arr_Td_Chinhtri.Count > 0, arr_Td_Chinhtri(cb_td_chinhtri.SelectedIndex), "0"), Integer)
                        obj_document.IdHocHam = CType(IIf(arr_Hocham.Count > 0, arr_Hocham(cb_hocham.SelectedIndex), "0"), Integer)
                        obj_document.IdHocVi = CType(IIf(arr_Hocvi.Count > 0, arr_Hocvi(cb_hocvi.SelectedIndex), "0"), Integer)
                        obj_document.IdTrdChMon = CType(IIf(arr_TrdChMon.Count > 0, arr_TrdChMon(cb_trd_chmon.SelectedIndex), "0"), Integer)
                        obj_document.IdTrdNgoaiNgu = CType(IIf(arr_TrdNgoaiNgu.Count > 0, arr_TrdNgoaiNgu(cb_trd_ngngu.SelectedIndex), "0"), Integer)
                        obj_document.IdTrdTinHoc = CType(IIf(arr_TrdTinHoc.Count > 0, arr_TrdTinHoc(cb_trd_tinhoc.SelectedIndex), "0"), Integer)
                        obj_document.IdChuyenNganhDT = CType(IIf(arr_ChNganh.Count > 0, arr_ChNganh(cb_chnganh.SelectedIndex), "0"), Integer)
                        If (ckb_cachmang.Checked) Then
                            obj_document.CM_Ngay = IIf(dtpk_cm_ngay_tg.Checked, dtpk_cm_ngay_tg.Value, DateTime.Parse("01/01/1900"))
                            obj_document.CM_ToChuc = Globals.Find_Replace(edt_cm_tochuc.Text.Trim.ToString())
                        Else
                            obj_document.CM_Ngay = DateTime.Parse("01/01/1900")
                            obj_document.CM_ToChuc = ""
                        End If
                        If (ckb_dangvien.Checked) Then
                            obj_document.Dang_NgayVao = IIf(dtpk_dv_ngayvao.Checked, dtpk_dv_ngayvao.Value, DateTime.Parse("01/01/1900"))
                            obj_document.Dang_NgayChTh = IIf(dtpk_dv_ngay_chinhthuc.Checked, dtpk_dv_ngay_chinhthuc.Value, DateTime.Parse("01/01/1900"))
                            obj_document.Dang_NgayRa = IIf(dtpk_dv_ngayra.Checked, dtpk_dv_ngayra.Value, DateTime.Parse("01/01/1900"))
                            obj_document.Dang_NoiKN = Globals.Find_Replace(edt_dv_noikn.Text.Trim.ToString())
                            obj_document.Dang_SoThe = Globals.Find_Replace(edt_dv_sothe.Text.Trim.ToString())
                            obj_document.Dang_NGT = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_dv_nguoi_gt.Text.Trim.ToString()))
                            obj_document.Dang_LyDoRa = Globals.Find_Replace(edt_dv_lydo.Text.Trim.ToString())
                        Else
                            obj_document.Dang_NgayVao = DateTime.Parse("01/01/1900")
                            obj_document.Dang_NgayChTh = DateTime.Parse("01/01/1900")
                            obj_document.Dang_NgayRa = DateTime.Parse("01/01/1900")
                            obj_document.Dang_NoiKN = ""
                            obj_document.Dang_SoThe = ""
                            obj_document.Dang_NGT = ""
                            obj_document.Dang_LyDoRa = ""
                        End If
                        If (ckb_taikhoan.Checked = True) Then
                            obj_document.NH_MSKH = Globals.Find_Replace(edt_tk_makh.Text.Trim.ToString())
                            obj_document.NH_SHTK = Globals.Find_Replace(edt_tk_sotk.Text.Trim.ToString())
                            obj_document.NH_Ten_NH = Globals.Find_Replace(edt_tk_tennh.Text.Trim.ToString())
                        Else
                            obj_document.NH_MSKH = ""
                            obj_document.NH_SHTK = ""
                            obj_document.NH_Ten_NH = ""
                        End If
                        obj_document.MaSoThue = Globals.Find_Replace(edt_ms_thue.Text.Trim.ToString())
                        If (ckb_bhxh.Checked = True) Then
                            obj_document.BHXH_SoSo = Globals.Find_Replace(edt_bhxh_soso.Text.Trim.ToString())
                            obj_document.BHXH_NgayLam = IIf(dtpk_bhxh_ngaylam.Checked, dtpk_bhxh_ngaylam.Value, DateTime.Parse("01/01/1900"))
                            obj_document.BHXH_NgayDong = IIf(dtpk_bhxh_ngaydong.Checked, dtpk_bhxh_ngaydong.Value, DateTime.Parse("01/01/1900"))
                            obj_document.BHXH_NoiLam = Globals.Find_Replace(edt_bhxh_noilam.Text.Trim.ToString())
                        Else
                            obj_document.BHXH_SoSo = ""
                            obj_document.BHXH_NgayLam = DateTime.Parse("01/01/1900")
                            obj_document.BHXH_NgayDong = DateTime.Parse("01/01/1900")
                            obj_document.BHXH_NoiLam = ""
                        End If

                        obj_document.idNew = ""
                        obj_document.SoTruong_CT = Globals.Find_Replace(edt_st_congtac.Text.Trim.ToString())
                        obj_document.CV_Lau = Globals.Find_Replace(edt_cv_launhat.Text.Trim.ToString())
                        obj_document.GhiChu = Globals.Find_Replace(edt_ghichu.Text.Trim.ToString())
                        obj_document.HonNhan_Cd = IIf(ARL_TT_HonNhan.Count > 0, ARL_TT_HonNhan(cb_tt_honnhan.SelectedIndex), "")
                        'Ảnh thẻ
                        If Not picbx_main.Image Is Nothing Then
                            obj_document.AnhThe = My_CImgToByte(picbx_main.Image)
                        End If

                        _IdCanBo = _HS_CanBo.Insert_HSCB_TS(obj_document)
                        'Cập nhật lại trạng thái cho hồ sơ cán bộ tập sự nếu lấy dữ liệu từ Hồ sơ cán bộ tập sự

                        If (_IdCB <> "") Then
                            'Thực hiện cập nhật idnew (id của cán bộ tập sự) vào hồ sơ cán bộ
                            _HS_CanBo.UpdateIdNew_HS_CanBo(_IdCanBo, _IdCB, 1)
                        End If

                        'Thực hiện Thêm mới dữ liệu - Hợp đồng lao động của cán bộ tập sự
                        Dim obj_hdld As clsHS_CanBo.HSCB_TS_HDLD = New clsHS_CanBo.HSCB_TS_HDLD()
                        If (rb_DinhBien.Checked = True) Then
                            obj_hdld.Status = 0
                        ElseIf (rb_PhuTro.Checked = True) Then
                            obj_hdld.Status = 1
                        End If
                        obj_hdld.IdCanbo = _IdCanBo
                        obj_hdld.SoQD = Globals.Find_Replace(edt_hdld_sohd.Text.Trim.ToString())
                        obj_hdld.NgayQD = dtpk_hdld_ngayky.Value
                        obj_hdld.NguoiQD = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hdld_nguoiky.Text.Trim.ToString()))
                        obj_hdld.IdCV_Nguoi_QD = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_hdld_chucvu.SelectedIndex), "0"), Integer)

                        obj_hdld.TuNgay = dtpk_tg_tungay.Value
                        obj_hdld.DenNgay = dtpk_tg_denngay.Value
                        'obj_hdld.TuGio = Globals.Find_Replace(edt_gio_batdau.Text.Trim.ToString())
                        'obj_hdld.DenGio = Globals.Find_Replace(edt_gio_ketthuc.Text.Trim.ToString())
                        obj_hdld.TuGio = CType(edt_gio_batdau.Value, DateTime).ToString("HH:mm")
                        obj_hdld.DenGio = CType(edt_gio_ketthuc.Value, DateTime).ToString("HH:mm")
                        obj_hdld.TV_Ngay_HL = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy")
                        obj_hdld.TV_NgayKy_QD = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy")
                        obj_hdld.TV_So_QD = ""
                        obj_hdld.Loai = CType(cb_hinhthuc_cv.SelectedIndex, Byte)
                        obj_hdld.IdHT_TraLuong = CType(IIf(arr_Ht_Traluong.Count > 0, arr_Ht_Traluong(cb_ht_traluong.SelectedIndex), "0"), Integer)

                        If (rb_thangbang.Checked = True) Then
                            If cb_cdld_bacluong.Items.Count <> 0 Then
                                obj_hdld.IdBacLuong = CType(IIf(arr_BacLuong.Count > 0, arr_BacLuong(cb_cdld_bacluong.SelectedIndex), "0"), Integer)
                            Else : obj_hdld.IdBacLuong = 0
                            End If
                            obj_hdld.HeSo = IIf(edt_cdld_heso.Text.ToString() <> "", CType(edt_cdld_heso.Text.ToString().Trim(), Double), 0)
                            obj_hdld.TyleHuong = IIf(edt_cdld_tylehuong.Text.ToString() <> "", CType(edt_cdld_tylehuong.Text.ToString().Trim(), Double), 0)
                        Else
                            obj_hdld.IdBacLuong = 0
                            obj_hdld.HeSo = 0
                            obj_hdld.TyleHuong = 0
                        End If
                        If (rb_trongoi.Checked = True) Then
                            obj_hdld.Tien_Luong = IIf(edt_tienluong.Text.Trim() <> "", CType(MoneyValue(edt_tienluong.Text.Trim()), Double), 0)
                        Else : obj_hdld.Tien_Luong = 0
                        End If
                        obj_hdld.IdChuyenMon = CType(IIf(arr_ChuyenMon.Count > 0, arr_ChuyenMon(cb_cdld_chuyenmon.SelectedIndex), "0"), Integer)
                        obj_hdld.CongViec = Globals.Find_Replace(edt_cdld_cv_damnhan.Text.Trim.ToString())
                        'If (rb_DinhBien.Checked = True) Then
                        '    obj_hdld.BHXH = IIf(ckb_choice_bhxh.Checked = True, 1, 0)
                        '    obj_hdld.BHYT = IIf(ckb_choice_bhyt.Checked = True, 1, 0)
                        'End If
                        obj_hdld.BHXH = IIf(ckb_choice_bhxh.Checked = True, 1, 0)
                        obj_hdld.BHYT = IIf(ckb_choice_bhyt.Checked = True, 1, 0)
                        obj_hdld.IdCanBo = _IdCanBo
                        obj_hdld.IdCBTS_HDLD = ""
                        obj_hdld.PhongBanDonViId = obj_document.IdPhongBan
                        obj_hdld.ChiNhanhId = obj_document.IdChiNhanh
                        _HS_CanBo.Insert_Update_HSCB_TS_HDLD(obj_hdld)

                        If (Progress_Changed IsNot Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    End If
                Case 2
                    If (Valid_Data()) Then
                        Dim obj_document As clsHS_CanBo.HSCB_TS = New clsHS_CanBo.HSCB_TS()
                        obj_document.Id = _IdCanBo
                        obj_document.MaCB = Globals.Find_Replace(edt_macb.Text.ToString().Trim())
                        obj_document.HoTen = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hoten.Text.ToString().Trim()))
                        obj_document.TenThuongGoi = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_ten_tg.Text.ToString().Trim()))
                        obj_document.BiDanh = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_bidanh.Text.ToString().Trim()))
                        obj_document.GioiTinh = CType(IIf(rb_nam.Checked = True, 0, 1), Byte)
                        obj_document.NgaySinh = dtpk_ngaysinh.Value
                        obj_document.IdChiNhanh = CType(IIf(arr_Donvi.Count > 0, arr_Donvi(cb_donvi.SelectedIndex), "0"), Integer)
                        obj_document.IdPhongBan = IIf(arr_Phongban.Count > 0, arr_Phongban(cb_phongban.SelectedIndex), "0").ToString()
                        obj_document.IdQuocTich = CType(IIf(arr_Quoctich.Count > 0, arr_Quoctich(cb_quoctich.SelectedIndex), "0"), Integer)
                        obj_document.IdDanToc = CType(IIf(arr_Dantoc.Count > 0, arr_Dantoc(cb_dantoc.SelectedIndex), "0"), Integer)
                        obj_document.IdTonGiao = CType(IIf(arr_Tongiao.Count > 0, arr_Tongiao(cb_tongiao.SelectedIndex), "0"), Integer)
                        obj_document.CMT_So = Globals.Find_Replace(edt_cmt_so.Text.ToString().Trim())
                        obj_document.CMT_NgayCap = dtpk_cmt_ngaycap.Value
                        obj_document.CMT_NoiCap = Globals.Find_Replace(edt_cmt_noicap.Text.Trim.ToString())

                        obj_document.IdNS_Tinh = CType(IIf(arr_Ns_Tinh.Count > 0, arr_Ns_Tinh(cb_ns_tinh.SelectedIndex), "0"), Integer)
                        If (cb_ns_huyen.Items.Count <> 0) Then
                            obj_document.IdNS_Huyen = CType(IIf(arr_Ns_Huyen.Count > 0, arr_Ns_Huyen(cb_ns_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdNS_Huyen = 0
                        End If
                        If (cb_ns_xa.Items.Count <> 0) Then
                            obj_document.IdNS_Xa = CType(IIf(arr_Ns_Xa.Count > 0, arr_Ns_Xa(cb_ns_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNS_Xa = 0
                        End If
                        If (cb_ns_thon.Items.Count <> 0) Then
                            obj_document.IdNS_Thon = CType(IIf(arr_Ns_Thon.Count > 0, arr_Ns_Thon(cb_ns_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNS_Thon = 0
                        End If
                        obj_document.NS_DiaChi = Globals.Find_Replace(edt_ns_diachi.Text.Trim.ToString())

                        obj_document.IdNQ_Tinh = CType(IIf(arr_Nq_Tinh.Count > 0, arr_Nq_Tinh(cb_nq_tinh.SelectedIndex), "0"), Integer)
                        If cb_nq_huyen.Items.Count <> 0 Then
                            obj_document.IdNQ_Huyen = CType(IIf(arr_Nq_Huyen.Count > 0, arr_Nq_Huyen(cb_nq_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdNQ_Huyen = 0
                        End If
                        If cb_nq_xa.Items.Count <> 0 Then
                            obj_document.IdNQ_Xa = CType(IIf(arr_Nq_Xa.Count > 0, arr_Nq_Xa(cb_nq_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNQ_Xa = 0
                        End If
                        If cb_nq_thon.Items.Count <> 0 Then
                            obj_document.IdNQ_Thon = CType(IIf(arr_Nq_Thon.Count > 0, arr_Nq_Thon(cb_nq_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdNQ_Thon = 0
                        End If
                        obj_document.NQ_DiaChi = Globals.Find_Replace(edt_nq_diachi.Text.Trim.ToString())

                        obj_document.IdThT_Tinh = CType(IIf(arr_TT_Tinh.Count > 0, arr_TT_Tinh(cb_tt_tinh.SelectedIndex), "0"), Integer)
                        If (cb_tt_huyen.Items.Count <> 0) Then
                            obj_document.IdThT_Huyen = CType(IIf(arr_TT_Huyen.Count > 0, arr_TT_Huyen(cb_tt_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdThT_Huyen = 0
                        End If
                        If (cb_tt_xa.Items.Count <> 0) Then
                            obj_document.IdThT_Xa = CType(IIf(arr_TT_Xa.Count > 0, arr_TT_Xa(cb_tt_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdThT_Xa = 0
                        End If
                        If (cb_tt_thon.Items.Count <> 0) Then
                            obj_document.IdThT_Thon = CType(IIf(arr_TT_Thon.Count > 0, arr_TT_Thon(cb_tt_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdThT_Thon = 0
                        End If
                        obj_document.ThT_DiaChi = Globals.Find_Replace(edt_tt_diachi.Text.Trim.ToString())
                        obj_document.ThT_DienThoai = Globals.Find_Replace(edt_tt_dienthoai.Text.Trim.ToString())

                        obj_document.IdTTr_Tinh = CType(IIf(arr_TTr_Tinh.Count > 0, arr_TTr_Tinh(cb_ttr_tinh.SelectedIndex), "0"), Integer)
                        If cb_ttr_huyen.Items.Count <> 0 Then
                            obj_document.IdTTr_Huyen = CType(IIf(arr_TTr_Huyen.Count > 0, arr_TTr_Huyen(cb_ttr_huyen.SelectedIndex), "0"), Integer)
                        Else : obj_document.IdTTr_Huyen = 0
                        End If
                        If cb_ttr_xa.Items.Count <> 0 Then
                            obj_document.IdTTr_Xa = CType(IIf(arr_TTr_Xa.Count > 0, arr_TTr_Xa(cb_ttr_xa.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdTTr_Xa = 0
                        End If
                        If cb_ttr_thon.Items.Count <> 0 Then
                            obj_document.IdTTr_Thon = CType(IIf(arr_TTr_Thon.Count > 0, arr_TTr_Thon(cb_ttr_thon.SelectedIndex), "0"), Int32)
                        Else : obj_document.IdTTr_Thon = 0
                        End If
                        obj_document.TTr_DiaChi = Globals.Find_Replace(edt_ttr_diachi.Text.Trim.ToString())
                        obj_document.TTr_DienThoai = Globals.Find_Replace(edt_ttr_dienthoai.Text.Trim.ToString())

                        obj_document.DienThoai_CQ = Globals.Find_Replace(edt_dt_coquan.Text.Trim.ToString())
                        obj_document.DienThoai_DD = Globals.Find_Replace(edt_dt_didong.Text.Trim.ToString())
                        obj_document.DienThoai_NR = Globals.Find_Replace(edt_dt_nharieng.Text.Trim.ToString())
                        obj_document.SoFax = Globals.Find_Replace(edt_sofax.Text.Trim.ToString())
                        obj_document.Email = Globals.Find_Replace(edt_email.Text.Trim.ToString())
                        obj_document.NhomMau = IIf(cb_nhommau.SelectedIndex > 0, cb_nhommau.SelectedItem.ToString(), "")

                        obj_document.IdThanhPhan_GD = CType(IIf(arr_Tp_Giadinh.Count > 0, arr_Tp_Giadinh(cb_tp_giadinh.SelectedIndex), "0"), Integer)
                        obj_document.IdUT_BT = _Globals.GetItems_CheckedListBox(clb_ut_banthan, arr_Ut_Banthan)
                        obj_document.IdUT_GD = CType(IIf(arr_Ut_Giadinh.Count > 0, arr_Ut_Giadinh(cb_ut_giadinh.SelectedIndex), "0"), Integer)
                        obj_document.IdTrinhDoVH = CType(IIf(arr_Td_Vanhoa.Count > 0, arr_Td_Vanhoa(cb_td_vanhoa.SelectedIndex), "0"), Integer)
                        obj_document.IdTrinhDoCT = CType(IIf(arr_Td_Chinhtri.Count > 0, arr_Td_Chinhtri(cb_td_chinhtri.SelectedIndex), "0"), Integer)

                        obj_document.IdHocHam = CType(IIf(arr_Hocham.Count > 0, arr_Hocham(cb_hocham.SelectedIndex), "0"), Integer)
                        obj_document.IdHocVi = CType(IIf(arr_Hocvi.Count > 0, arr_Hocvi(cb_hocvi.SelectedIndex), "0"), Integer)
                        obj_document.IdTrdChMon = CType(IIf(arr_TrdChMon.Count > 0, arr_TrdChMon(cb_trd_chmon.SelectedIndex), "0"), Integer)
                        obj_document.IdTrdNgoaiNgu = CType(IIf(arr_TrdNgoaiNgu.Count > 0, arr_TrdNgoaiNgu(cb_trd_ngngu.SelectedIndex), "0"), Integer)
                        obj_document.IdTrdTinHoc = CType(IIf(arr_TrdTinHoc.Count > 0, arr_TrdTinHoc(cb_trd_tinhoc.SelectedIndex), "0"), Integer)
                        obj_document.IdChuyenNganhDT = CType(IIf(arr_ChNganh.Count > 0, arr_ChNganh(cb_chnganh.SelectedIndex), "0"), Integer)
                        If (ckb_cachmang.Checked = True) Then
                            obj_document.CM_Ngay = IIf(dtpk_cm_ngay_tg.Checked = True, dtpk_cm_ngay_tg.Value, DateTime.Parse("01/01/1900"))
                            obj_document.CM_ToChuc = Globals.Find_Replace(edt_cm_tochuc.Text.Trim.ToString())
                        Else
                            obj_document.CM_Ngay = DateTime.Parse("01/01/1900")
                            obj_document.CM_ToChuc = ""
                        End If
                        If (ckb_dangvien.Checked = True) Then
                            obj_document.Dang_NgayVao = IIf(dtpk_dv_ngayvao.Checked = True, dtpk_dv_ngayvao.Value, DateTime.Parse("01/01/1900"))
                            obj_document.Dang_NgayChTh = IIf(dtpk_dv_ngay_chinhthuc.Checked = True, dtpk_dv_ngay_chinhthuc.Value, DateTime.Parse("01/01/1900"))
                            obj_document.Dang_NgayRa = IIf(dtpk_dv_ngayra.Checked = True, dtpk_dv_ngayra.Value, DateTime.Parse("01/01/1900"))
                            obj_document.Dang_NoiKN = Globals.Find_Replace(edt_dv_noikn.Text.Trim.ToString())
                            obj_document.Dang_SoThe = Globals.Find_Replace(edt_dv_sothe.Text.Trim.ToString())
                            obj_document.Dang_NGT = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_dv_nguoi_gt.Text.Trim.ToString()))
                            obj_document.Dang_LyDoRa = Globals.Find_Replace(edt_dv_lydo.Text.Trim.ToString())
                        Else
                            obj_document.Dang_NgayVao = DateTime.Parse("01/01/1900")
                            obj_document.Dang_NgayChTh = DateTime.Parse("01/01/1900")
                            obj_document.Dang_NgayRa = DateTime.Parse("01/01/1900")
                            obj_document.Dang_NoiKN = ""
                            obj_document.Dang_SoThe = ""
                            obj_document.Dang_NGT = ""
                            obj_document.Dang_LyDoRa = ""
                        End If
                        If (ckb_taikhoan.Checked = True) Then
                            obj_document.NH_MSKH = Globals.Find_Replace(edt_tk_makh.Text.Trim.ToString())
                            obj_document.NH_SHTK = Globals.Find_Replace(edt_tk_sotk.Text.Trim.ToString())
                            obj_document.NH_Ten_NH = Globals.Find_Replace(edt_tk_tennh.Text.Trim.ToString())
                        Else
                            obj_document.NH_MSKH = ""
                            obj_document.NH_SHTK = ""
                            obj_document.NH_Ten_NH = ""
                        End If
                        obj_document.MaSoThue = Globals.Find_Replace(edt_ms_thue.Text.Trim.ToString())
                        If (ckb_bhxh.Checked = True) Then
                            obj_document.BHXH_SoSo = Globals.Find_Replace(edt_bhxh_soso.Text.Trim.ToString())
                            obj_document.BHXH_NgayLam = IIf(dtpk_bhxh_ngaylam.Checked = True, dtpk_bhxh_ngaylam.Value, DateTime.Parse("01/01/1900"))
                            obj_document.BHXH_NgayDong = IIf(dtpk_bhxh_ngaydong.Checked = True, dtpk_bhxh_ngaydong.Value, DateTime.Parse("01/01/1900"))
                            obj_document.BHXH_NoiLam = Globals.Find_Replace(edt_bhxh_noilam.Text.Trim.ToString())
                        Else
                            obj_document.BHXH_SoSo = ""
                            obj_document.BHXH_NgayLam = DateTime.Parse("01/01/1900")
                            obj_document.BHXH_NgayDong = DateTime.Parse("01/01/1900")
                            obj_document.BHXH_NoiLam = ""
                        End If

                        'Lấy thông tin - Hợp đồng lao động của Tập sự
                        Dim obj_hdld_ts As clsHS_CanBo.HSCB_TS_HDLD = New clsHS_CanBo.HSCB_TS_HDLD()
                        obj_hdld_ts.IdCBTS_HDLD = _IdHdld
                        If (rb_DinhBien.Checked = True) Then
                            obj_hdld_ts.Status = 0
                        ElseIf (rb_PhuTro.Checked = True) Then
                            obj_hdld_ts.Status = 1
                        End If
                        obj_hdld_ts.IdCanbo = _IdCanBo
                        obj_hdld_ts.SoQD = Globals.Find_Replace(edt_hdld_sohd.Text.Trim.ToString())
                        obj_hdld_ts.NgayQD = dtpk_hdld_ngayky.Value
                        obj_hdld_ts.NguoiQD = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hdld_nguoiky.Text.Trim.ToString()))
                        obj_hdld_ts.IdCV_Nguoi_QD = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_hdld_chucvu.SelectedIndex), "0"), Integer)

                        obj_hdld_ts.TuNgay = dtpk_tg_tungay.Value
                        obj_hdld_ts.DenNgay = dtpk_tg_denngay.Value
                        obj_hdld_ts.TuGio = CType(edt_gio_batdau.Value, DateTime).ToString("HH:mm")
                        obj_hdld_ts.DenGio = CType(edt_gio_ketthuc.Value, DateTime).ToString("HH:mm")
                        obj_hdld_ts.Loai = CType(cb_hinhthuc_cv.SelectedIndex, Byte)
                        obj_hdld_ts.IdHT_TraLuong = CType(IIf(arr_Ht_Traluong.Count > 0, arr_Ht_Traluong(cb_ht_traluong.SelectedIndex), "0"), Integer)

                        If (rb_thangbang.Checked = True) Then
                            If cb_cdld_bacluong.Items.Count <> 0 Then
                                obj_hdld_ts.IdBacLuong = CType(IIf(arr_BacLuong.Count > 0, arr_BacLuong(cb_cdld_bacluong.SelectedIndex), "0"), Integer)
                            Else : obj_hdld_ts.IdBacLuong = 0
                            End If
                            obj_hdld_ts.HeSo = IIf(edt_cdld_heso.Text.ToString() <> "", CType(edt_cdld_heso.Text.ToString().Trim(), Double), 0)
                            obj_hdld_ts.TyleHuong = IIf(edt_cdld_tylehuong.Text.ToString() <> "", CType(edt_cdld_tylehuong.Text.ToString().Trim(), Double), 0)
                        Else
                            obj_hdld_ts.IdBacLuong = 0
                            obj_hdld_ts.HeSo = 0
                            obj_hdld_ts.TyleHuong = 100
                        End If

                        If (rb_trongoi.Checked = True) Then
                            obj_hdld_ts.Tien_Luong = IIf(edt_tienluong.Text.Trim() <> "", CType(MoneyValue(edt_tienluong.Text.Trim()), Double), 0)
                        Else : obj_hdld_ts.Tien_Luong = 0
                        End If
                        obj_hdld_ts.IdChuyenMon = CType(IIf(arr_ChuyenMon.Count > 0, arr_ChuyenMon(cb_cdld_chuyenmon.SelectedIndex), "0"), Integer)
                        obj_hdld_ts.CongViec = Globals.Find_Replace(edt_cdld_cv_damnhan.Text.Trim.ToString())
                        'If (rb_DinhBien.Checked = True) Then
                        '    obj_hdld_ts.BHXH = IIf(ckb_choice_bhxh.Checked = True, 1, 0)
                        '    obj_hdld_ts.BHYT = IIf(ckb_choice_bhyt.Checked = True, 1, 0)
                        'End If
                        obj_hdld_ts.BHXH = IIf(ckb_choice_bhxh.Checked = True, 1, 0)
                        obj_hdld_ts.BHYT = IIf(ckb_choice_bhyt.Checked = True, 1, 0)

                        Dim sSQL As String = ""
                        sSQL = String.Format("Select IsNull(TV_So_QD,'') From HSCB_TS_HDLD Where IdCanBo='{0}'", _IdCanBo)
                        Dim _TV_SoQD = SoftSqlHelper.GetString(String.Format("Select IsNull(TV_So_QD,'') From HSCB_TS_HDLD Where IdCanBo='{0}'", _IdCanBo), "")
                        If String.IsNullOrEmpty(_TV_SoQD) Or IsNothing(_TV_SoQD) Then
                            obj_hdld_ts.TV_Ngay_HL = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy")
                            obj_hdld_ts.TV_NgayKy_QD = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy")
                            obj_hdld_ts.TV_So_QD = ""
                        Else
                            obj_hdld_ts.TV_So_QD = _TV_SoQD
                        End If
                        obj_hdld_ts.PhongBanDonViId = obj_document.IdPhongBan
                        obj_hdld_ts.ChiNhanhId = obj_document.IdChiNhanh

                        obj_document.SoTruong_CT = Globals.Find_Replace(edt_st_congtac.Text.Trim.ToString())
                        obj_document.CV_Lau = Globals.Find_Replace(edt_cv_launhat.Text.Trim.ToString())
                        obj_document.GhiChu = Globals.Find_Replace(edt_ghichu.Text.Trim.ToString())
                        obj_document.HonNhan_Cd = IIf(ARL_TT_HonNhan.Count > 0, ARL_TT_HonNhan(cb_tt_honnhan.SelectedIndex), "")
                        'Ảnh thẻ

                        obj_document.AnhThe = My_CImgToByte(picbx_main.Image)

                        'Thực hiện gọi hàm cập nhật dữ liệu - Hồ sơ cán bộ tập sự

                        _HS_CanBo.Update_HSCB_TS(obj_document)

                        'Thực hiện cập nhật dữ liệu - Hợp đồng lao động của cán bộ tập sự
                        _HS_CanBo.Insert_Update_HSCB_TS_HDLD(obj_hdld_ts)
                        If (Progress_Changed IsNot Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    End If
            End Select
        Catch ex As Exception
            MessageBox.Show("Cập nhật hồ sơ cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

#Region "---> Events: Bắt Các sự kiện ngoại lệ <---"
    Private Sub cb_donvi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_donvi.SelectedIndexChanged
        cb_phongban.Items.Clear()
        arr_Phongban.Clear()
        If (cb_donvi.SelectedIndex > 0 And cb_donvi.Items.Count <> 0) Then
            Dim _valId As Integer = CType(arr_Donvi(IIf(cb_donvi.SelectedIndex > 0, cb_donvi.SelectedIndex, "0")), Integer)
            If _valId > 0 Then
                'Lấy mã chi nhánh từ Id chi nhánh
                Dim strSQL As String = ""
                strSQL = String.Format("SELECT * FROM ChiNhanh WHERE id = {0} AND Status = 1", _valId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim tructhuocCurrent As String = _HS_CanBo.GetTrucThuoc(db.Rows(0)("ma_so").ToString().Trim())
                            If db.Rows(0)("ma_so").ToString().Trim() = gMaDonViTW Then
                                strSQL = String.Format("Select ('PB_'+Ltrim(Str(id))) as id,ten_phong as ten_goi From PhongBan Where Charindex('{0}',truc_thuoc) > 0 and Status = 1 order by ten_phong ", tructhuocCurrent.ToString)
                            Else
                                strSQL = String.Format("Select ('PB_'+Ltrim(Str(id))) as id,ten_phong as ten_goi From PhongBan Where Charindex('{0}',truc_thuoc) > 0 and Status = 1 Union Select ('DV_'+Ltrim(Str(id))) as id, ten_goi from ChiNhanh Where id_goc = {1} and Status = 1 Order by ten_goi", tructhuocCurrent.ToString, _valId)
                            End If
                            arr_Phongban = _Globals.Bind_ComBoBox(cb_phongban, strSQL, "---Phòng ban---")
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    Private Sub ckb_cachmang_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_cachmang.CheckedChanged
        If (ckb_cachmang.Checked = True) Then
            dtpk_cm_ngay_tg.Enabled = True
            edt_cm_tochuc.ReadOnly = False
        Else
            dtpk_cm_ngay_tg.Enabled = False
            edt_cm_tochuc.ReadOnly = True
        End If
    End Sub

    Private Sub ckb_dangvien_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_dangvien.CheckedChanged
        If (ckb_dangvien.Checked = True And _FlagEvent <> 3) Then
            dtpk_dv_ngayvao.Enabled = True
            dtpk_dv_ngay_chinhthuc.Enabled = True
            dtpk_dv_ngayra.Enabled = True
            edt_dv_noikn.ReadOnly = False
            edt_dv_sothe.ReadOnly = False
            edt_dv_nguoi_gt.ReadOnly = False
            edt_dv_lydo.ReadOnly = False
        Else
            dtpk_dv_ngayvao.Enabled = False
            dtpk_dv_ngay_chinhthuc.Enabled = False
            dtpk_dv_ngayra.Enabled = False
            edt_dv_noikn.ReadOnly = True
            edt_dv_sothe.ReadOnly = True
            edt_dv_nguoi_gt.ReadOnly = True
            edt_dv_lydo.ReadOnly = True
        End If
    End Sub

    Private Sub ckb_taikhoan_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_taikhoan.CheckedChanged
        If (ckb_taikhoan.Checked = True And _FlagEvent <> 3) Then
            edt_tk_makh.ReadOnly = False
            edt_tk_sotk.ReadOnly = False
            edt_tk_tennh.ReadOnly = False
        Else
            edt_tk_makh.ReadOnly = True
            edt_tk_sotk.ReadOnly = True
            edt_tk_tennh.ReadOnly = True
        End If
    End Sub

    Private Sub ckb_bhxh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_bhxh.CheckedChanged
        If (ckb_bhxh.Checked = True And _FlagEvent <> 3) Then
            edt_bhxh_soso.ReadOnly = False
            dtpk_bhxh_ngaylam.Enabled = True
            dtpk_bhxh_ngaydong.Enabled = True
            edt_bhxh_noilam.ReadOnly = False
        Else
            edt_bhxh_soso.ReadOnly = True
            dtpk_bhxh_ngaylam.Enabled = False
            dtpk_bhxh_ngaydong.Enabled = False
            edt_bhxh_noilam.ReadOnly = True
        End If
    End Sub

    Private Sub rb_thangbang_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_thangbang.CheckedChanged
        If (rb_thangbang.Checked = True) Then
            edt_tienluong.Visible = False
            lbl_ngdinhluong.Text = "Nghị định lương  "
            edt_tienluong.Width = 1

            cb_cdld_nghidinh.Visible = True
            pnl_tt_0.Visible = True
            pnl_tt_1.Visible = True
            pnl_tt_2.Visible = True

            pnl_div_0.Visible = True
            pnl_div_1.Visible = True
            pnl_div_2.Visible = True

            If (cb_cdld_nghidinh.Items.Count <> 0) Then
                cb_cdld_nghidinh.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub rb_trongoi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_trongoi.CheckedChanged
        If (rb_trongoi.Checked = True) Then
            edt_tienluong.Visible = True
            lbl_ngdinhluong.Text = "Tiền lương  "
            edt_tienluong.Width = 180
            cb_cdld_nghidinh.Visible = False
            pnl_tt_0.Visible = False
            pnl_tt_1.Visible = False
            pnl_div_0.Visible = False
            pnl_div_1.Visible = False
            pnl_tt_2.Visible = False
            pnl_div_2.Visible = False
            If (edt_tienluong.Text.Trim() = "") Then
                edt_tienluong.Text = "0"
            End If
        End If
    End Sub

    Private Sub cb_ns_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ns_tinh.SelectedIndexChanged
        'cb_ns_huyen.Items.Clear()
        'arr_Ns_Huyen.Clear()
        'If (cb_ns_tinh.SelectedIndex > 0 And cb_ns_tinh.Items.Count <> 0) Then
        '    Dim _Id_TTP As Integer = CType(arr_Ns_Tinh(IIf(cb_ns_tinh.SelectedIndex > 0, cb_ns_tinh.SelectedIndex, "0")), Integer)
        '    If _Id_TTP > 0 Then
        '        Dim strSQL As String = String.Format("Select Id,Ten_Goi From DiaDanh Where id_goc != 0 and id_goc = {0} And Status = 1", _Id_TTP)
        '        arr_Ns_Huyen = _Globals.Bind_ComBoBox(cb_ns_huyen, strSQL, "---Quận - huyện---")
        '        If (_FlagEvent = 1 And cb_nq_tinh.Items.Count <> 0) Then
        '            cb_nq_tinh.SelectedIndex = cb_ns_tinh.SelectedIndex
        '            cb_nq_tinh_SelectedIndexChanged(sender, e)
        '        End If
        '    End If
        'End If

        cb_ns_xa.Items.Clear()
        arr_Ns_Xa.Clear()
        If (cb_ns_tinh.SelectedIndex > 0 And cb_ns_tinh.Items.Count <> 0) Then
            Dim _IdTinhTP As Int32 = CType(arr_Ns_Tinh(IIf(cb_ns_tinh.SelectedIndex > 0, cb_ns_tinh.SelectedIndex, "0")), Int32)
            If _IdTinhTP > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
                arr_Ns_Xa = _Globals.Bind_ComBoBox(cb_ns_xa, strSQL, "--- Chọn Xã/Phường ---")
                If (_FlagEvent = 1 And cb_nq_tinh.Items.Count <> 0) Then
                    cb_nq_tinh.SelectedIndex = cb_ns_tinh.SelectedIndex
                    cb_nq_tinh_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_ns_huyen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ns_huyen.SelectedIndexChanged
        If (cb_nq_huyen.Items.Count <> 0 And cb_ns_huyen.Items.Count <> 0) Then
            If (_FlagEvent = 1) Then
                cb_nq_huyen.SelectedIndex = cb_ns_huyen.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_ns_xa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ns_xa.SelectedIndexChanged
        cb_ns_thon.Items.Clear()
        arr_Ns_Thon.Clear()
        If (cb_ns_xa.SelectedIndex > 0 And cb_ns_xa.Items.Count <> 0) Then
            Dim _IdXaPhuong As Int32 = CType(arr_Ns_Xa(IIf(cb_ns_xa.SelectedIndex > 0, cb_ns_xa.SelectedIndex, "0")), Int32)
            If _IdXaPhuong > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon <> '00' And TrangThai = 'A' And Ma_Tinh+Ma_Xa In (Select Top 1 X.Ma_Tinh+X.Ma_Xa From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdXaPhuong)
                arr_Ns_Thon = _Globals.Bind_ComBoBox(cb_ns_thon, strSQL, "--- Chọn Thôn/Xóm ---")
                If (_FlagEvent = 1 And cb_nq_xa.Items.Count <> 0) Then
                    cb_nq_xa.SelectedIndex = cb_ns_xa.SelectedIndex
                    cb_nq_xa_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_ns_thon_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ns_thon.SelectedIndexChanged
        If (cb_nq_thon.Items.Count <> 0 And cb_ns_thon.Items.Count <> 0) Then
            If (_FlagEvent = 1 And _IdCanBo = "") Then
                cb_nq_thon.SelectedIndex = cb_ns_thon.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_nq_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_nq_tinh.SelectedIndexChanged
        'cb_nq_huyen.Items.Clear()
        'arr_Nq_Huyen.Clear()
        'If (cb_nq_tinh.SelectedIndex > 0 And cb_nq_tinh.Items.Count <> 0) Then
        '    Dim _Id_TTP As Integer = CType(arr_Ns_Tinh(IIf(cb_nq_tinh.SelectedIndex > 0, cb_nq_tinh.SelectedIndex, "0")), Integer)
        '    If _Id_TTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi from DiaDanh where id_goc != 0 and id_goc = {0} and Status = 1", _Id_TTP)
        '        arr_Nq_Huyen = _Globals.Bind_ComBoBox(cb_nq_huyen, strSQL, "---Quận - huyện---")
        '    End If
        'End If
        cb_nq_xa.Items.Clear()
        arr_Nq_Xa.Clear()
        If (cb_nq_tinh.SelectedIndex > 0 And cb_nq_tinh.Items.Count <> 0) Then
            Dim _IdTinhTP As Int32 = CType(arr_Nq_Tinh(IIf(cb_nq_tinh.SelectedIndex > 0, cb_nq_tinh.SelectedIndex, "0")), Int32)
            If _IdTinhTP > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
                arr_Nq_Xa = _Globals.Bind_ComBoBox(cb_nq_xa, strSQL, "--- Chọn Xã/Phường ---")
            End If
        End If
    End Sub

    Private Sub cb_nq_xa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_nq_xa.SelectedIndexChanged
        cb_nq_thon.Items.Clear()
        arr_Nq_Thon.Clear()
        If (cb_nq_xa.SelectedIndex > 0 And cb_nq_xa.Items.Count <> 0) Then
            Dim _IdXaPhuong As Int32 = CType(arr_Nq_Xa(IIf(cb_nq_xa.SelectedIndex > 0, cb_nq_xa.SelectedIndex, "0")), Int32)
            If _IdXaPhuong > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon <> '00' And TrangThai = 'A' And Ma_Tinh+Ma_Xa In (Select Top 1 X.Ma_Tinh+X.Ma_Xa From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdXaPhuong)
                arr_Nq_Thon = _Globals.Bind_ComBoBox(cb_nq_thon, strSQL, "--- Chọn Thôn/Xóm ---")
                If (_FlagEvent = 1 And cb_nq_xa.Items.Count <> 0 And cb_tt_xa.Items.Count <> 0) Then
                    cb_tt_xa.SelectedIndex = cb_nq_xa.SelectedIndex
                    cb_tt_xa_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_nq_thon_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_nq_thon.SelectedIndexChanged
        If (cb_nq_thon.Items.Count <> 0 And cb_tt_thon.Items.Count <> 0) Then
            If (_FlagEvent = 1 And _IdCanBo = "") Then
                cb_tt_thon.SelectedIndex = cb_nq_thon.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_tt_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tt_tinh.SelectedIndexChanged
        'cb_tt_huyen.Items.Clear()
        'arr_TT_Huyen.Clear()
        'If (cb_tt_tinh.SelectedIndex > 0 And cb_tt_tinh.Items.Count <> 0) Then
        '    Dim _Id_TTP As Integer = CType(arr_TT_Tinh(IIf(cb_tt_tinh.SelectedIndex > 0, cb_tt_tinh.SelectedIndex, "0")), Integer)
        '    If _Id_TTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi from DiaDanh Where id_goc != 0 and id_goc = {0} and Status = 1", _Id_TTP)
        '        arr_TT_Huyen = _Globals.Bind_ComBoBox(cb_tt_huyen, strSQL, "---Quận - huyện---")
        '        If (_FlagEvent = 1 And cb_ttr_tinh.Items.Count <> 0) Then
        '            cb_ttr_tinh.SelectedIndex = cb_tt_tinh.SelectedIndex
        '            cb_ttr_tinh_SelectedIndexChanged(sender, e)
        '        End If
        '    End If
        'End If
        cb_tt_xa.Items.Clear()
        arr_TT_Xa.Clear()
        If (cb_tt_tinh.SelectedIndex > 0 And cb_tt_tinh.Items.Count <> 0) Then
            Dim _IdTinhTP As Int32 = CType(arr_TT_Tinh(IIf(cb_tt_tinh.SelectedIndex > 0, cb_tt_tinh.SelectedIndex, "0")), Int32)
            If _IdTinhTP > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
                arr_TT_Xa = _Globals.Bind_ComBoBox(cb_tt_xa, strSQL, "--- Chọn Xã/Phường ---")
                If (_FlagEvent = 1 And cb_ttr_tinh.Items.Count <> 0 And cb_tt_tinh.Items.Count <> 0) Then
                    cb_ttr_tinh.SelectedIndex = cb_tt_tinh.SelectedIndex
                    cb_ttr_tinh_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_tt_huyen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tt_huyen.SelectedIndexChanged
        If (cb_ttr_huyen.Items.Count <> 0) Then
            If (_FlagEvent = 1) Then
                cb_ttr_huyen.SelectedIndex = cb_tt_huyen.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_tt_xa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_tt_xa.SelectedIndexChanged
        cb_tt_thon.Items.Clear()
        arr_TT_Thon.Clear()
        If (cb_tt_xa.SelectedIndex > 0 And cb_tt_xa.Items.Count <> 0) Then
            Dim _IdXaPhuong As Int32 = CType(arr_TT_Xa(IIf(cb_tt_xa.SelectedIndex > 0, cb_tt_xa.SelectedIndex, "0")), Int32)
            If _IdXaPhuong > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon <> '00' And TrangThai = 'A' And Ma_Tinh+Ma_Xa In (Select Top 1 X.Ma_Tinh+X.Ma_Xa From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdXaPhuong)
                arr_TT_Thon = _Globals.Bind_ComBoBox(cb_tt_thon, strSQL, "--- Chọn Thôn/Xóm ---")
                If (_FlagEvent = 1 And cb_tt_xa.Items.Count <> 0 And cb_ttr_xa.Items.Count <> 0) Then
                    cb_ttr_xa.SelectedIndex = cb_tt_xa.SelectedIndex
                    cb_ttr_xa_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_tt_thon_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_tt_thon.SelectedIndexChanged
        If (cb_tt_thon.Items.Count <> 0 And cb_ttr_thon.Items.Count <> 0) Then
            If (_FlagEvent = 1) Then
                cb_ttr_thon.SelectedIndex = cb_tt_thon.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_ttr_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ttr_tinh.SelectedIndexChanged
        'cb_ttr_huyen.Items.Clear()
        'arr_TTr_Huyen.Clear()
        'If (cb_ttr_tinh.SelectedIndex > 0 And cb_ttr_tinh.Items.Count <> 0) Then
        '    Dim _Id_TTP As Integer = CType(arr_TT_Tinh(IIf(cb_ttr_tinh.SelectedIndex > 0, cb_ttr_tinh.SelectedIndex, "0")), Integer)
        '    If _Id_TTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi from DiaDanh Where id_goc != 0 and id_goc = {0} and Status = 1", _Id_TTP)
        '        arr_TTr_Huyen = _Globals.Bind_ComBoBox(cb_ttr_huyen, strSQL, "---Quận - huyện---")
        '    End If
        'End If
        cb_ttr_xa.Items.Clear()
        arr_TTr_Xa.Clear()
        If (cb_ttr_tinh.SelectedIndex > 0 And cb_ttr_tinh.Items.Count <> 0) Then
            Dim _IdTinhTP As Int32 = CType(arr_TTr_Tinh(IIf(cb_ttr_tinh.SelectedIndex > 0, cb_ttr_tinh.SelectedIndex, "0")), Int32)
            If _IdTinhTP > 0 Then
                'Dim strSQL As String = String.Format("Select id,ten_goi From DiaDanh Where id_goc != 0 And id_goc = {0}", _IdTinhTP)
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
                arr_TTr_Xa = _Globals.Bind_ComBoBox(cb_ttr_xa, strSQL, "--- Chọn Xã/Phường ---")
            End If
        End If
    End Sub

    Private Sub cb_ttr_xa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ttr_xa.SelectedIndexChanged
        cb_ttr_thon.Items.Clear()
        arr_TTr_Thon.Clear()
        If (cb_ttr_xa.SelectedIndex > 0 And cb_ttr_xa.Items.Count <> 0) Then
            Dim _IdXaPhuong As Int32 = CType(arr_TTr_Xa(IIf(cb_ttr_xa.SelectedIndex > 0, cb_ttr_xa.SelectedIndex, "0")), Int32)
            If _IdXaPhuong > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon <> '00' And TrangThai = 'A' And Ma_Tinh+Ma_Xa In (Select Top 1 X.Ma_Tinh+X.Ma_Xa From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdXaPhuong)
                arr_TTr_Thon = _Globals.Bind_ComBoBox(cb_ttr_thon, strSQL, "--- Chọn Thôn/Xóm ---")
            End If
        End If
    End Sub

    Private Sub cb_cdld_nghidinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_cdld_nghidinh.SelectedIndexChanged
        cb_cdld_bangluong.Items.Clear()
        arr_NgachLuong.Clear()
        If (cb_cdld_nghidinh.SelectedIndex > 0 And cb_cdld_nghidinh.Items.Count <> 0) Then
            Dim NghiDinhId As Integer = CType(arr_NghiDinhLuong(IIf(cb_cdld_nghidinh.SelectedIndex > 0, cb_cdld_nghidinh.SelectedIndex, "0")), Integer)
            If (NghiDinhId > 0) Then
                Dim strSQL As String = String.Format("Select IdBangLuong, Mota From BangLuong Where IdND_Luong = {0}", NghiDinhId)
                arr_BangLuong = _Globals.Bind_ComBoBox(cb_cdld_bangluong, strSQL, "---Bảng lương---")
            End If
        End If
        cb_cdld_bangluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_cdld_bangluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_cdld_bangluong.SelectedIndexChanged
        cb_cdld_ngachluong.Items.Clear()
        arr_NgachLuong.Clear()
        lbl_bangluong_mota.Text = ""
        If (cb_cdld_bangluong.SelectedIndex > 0 And cb_cdld_bangluong.Items.Count <> 0) Then
            Dim BangLuongId As Integer = CType(arr_BangLuong(IIf(cb_cdld_bangluong.SelectedIndex > 0, cb_cdld_bangluong.SelectedIndex, "0")), Integer)
            If (BangLuongId > 0) Then
                Dim strSQL As String = String.Format("Select IdNgachLuong, Mota From NgachLuong Where IdBangLuong = {0} and Status = 1", BangLuongId)
                arr_NgachLuong = _Globals.Bind_ComBoBox(cb_cdld_ngachluong, strSQL, "---Ngạch lương---")
                'Load thông tin tên bảng lương ra theo dõi
                strSQL = String.Format("Select * From BangLuong Where IdBangLuong = {0}", BangLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            lbl_bangluong_mota.Text = db.Rows(0)("BangLuong").ToString().Trim()
                        End If
                    End If
                End Using
            End If
        End If
        cb_cdld_ngachluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_cdld_ngachluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_cdld_ngachluong.SelectedIndexChanged
        cb_cdld_bacluong.Items.Clear()
        arr_BacLuong.Clear()
        lbl_ngachluong_mota.Text = ""
        If (cb_cdld_ngachluong.SelectedIndex > 0 And cb_cdld_ngachluong.Items.Count <> 0) Then
            Dim NgachLuongId As Integer = CType(arr_NgachLuong(IIf(cb_cdld_ngachluong.SelectedIndex > 0, cb_cdld_ngachluong.SelectedIndex, "0")), Integer)
            If (NgachLuongId > 0) Then
                Dim strSQL As String = String.Format("Select IdBacLuong, Mota From BacLuong Where IdNgachLuong = {0} and Status = 1", NgachLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            cb_cdld_bacluong.Items.Add("---Bậc lương---")
                            arr_BacLuong.Add(0)
                            For i As Integer = 0 To db.Rows.Count - 1
                                cb_cdld_bacluong.Items.Add(db.Rows(i)("Mota").ToString())
                                arr_BacLuong.Add(IIf(db.Rows(i)("IdBacLuong").ToString() <> "", db.Rows(i)("IdBacLuong").ToString(), ""))
                            Next
                            If (cb_cdld_bacluong.Items.Count <> 0) Then
                                cb_cdld_bacluong.SelectedIndex = 0
                            End If
                        End If
                    End If
                End Using

                'Load thông tin tên bảng lương ra theo dõi
                strSQL = String.Format("Select * From NgachLuong Where IdNgachLuong = {0}", NgachLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            lbl_ngachluong_mota.Text = db.Rows(0)("NgachLuong").ToString().Trim()
                        End If
                    End If
                End Using
            End If
        End If
        cb_cdld_bacluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_cdld_bacluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_cdld_bacluong.SelectedIndexChanged
        edt_cdld_heso.Text = "0"
        If (cb_cdld_bacluong.SelectedIndex > 0 And cb_cdld_bacluong.Items.Count <> 0) Then
            Dim BacLuongId As Integer = CType(arr_BacLuong(IIf(cb_cdld_bacluong.SelectedIndex > 0, cb_cdld_bacluong.SelectedIndex, "0")), Integer)
            If (BacLuongId > 0) Then
                Dim strSQL As String = String.Format("Select * From BacLuong Where Status = 1 and IdBacLuong = {0}", BacLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            edt_cdld_heso.Text = IIf(db.Rows(0)("Heso").ToString().Trim() <> "", CType(db.Rows(0)("Heso"), Double).ToString("N2"), "0")
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    Private Sub edt_cdld_tylehuong_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_cdld_tylehuong.KeyPress
        Dim decimalString As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
        Dim decimalChar As Char = Convert.ToChar(decimalString)
        If (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
        ElseIf (e.KeyChar = decimalString And edt_cdld_tylehuong.Text.IndexOf(decimalString) = -1) Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub edt_cmt_so_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_cmt_so.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_tk_sotk_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_tk_sotk.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_ttr_dienthoai_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_ttr_dienthoai.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_tt_dienthoai_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_tt_dienthoai.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_dt_didong_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dt_didong.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_dt_coquan_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dt_coquan.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_dt_nharieng_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dt_nharieng.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_cdld_tylehuong_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_cdld_tylehuong.Leave
        If (edt_cdld_tylehuong.Text.Trim() = "") Then
            edt_cdld_tylehuong.Text = "0"
        End If
    End Sub

    Private Sub edt_tt_dienthoai_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tt_dienthoai.TextChanged
        If (_FlagEvent = 1) Then
            edt_ttr_dienthoai.Text = edt_tt_dienthoai.Text.Trim()
            edt_dt_nharieng.Text = edt_tt_dienthoai.Text.Trim()
        End If
    End Sub

    Private Sub edt_ns_diachi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_ns_diachi.TextChanged
        If (_FlagEvent = 1) Then
            edt_nq_diachi.Text = edt_ns_diachi.Text.Trim()
        End If
    End Sub

    Private Sub edt_tt_diachi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tt_diachi.TextChanged
        If (_FlagEvent = 1) Then
            edt_ttr_diachi.Text = edt_tt_diachi.Text.Trim()
        End If
    End Sub

    Private Sub rb_hdld_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_DinhBien.CheckedChanged
        If rb_DinhBien.Checked = True Then
            'pnl_div_hdld_baohiem.Visible = True
            'pnl_hdld_baohiem.Visible = True
            lbl_hdld_sohd_soqd.Text = "Số hợp đồng  "
        End If
    End Sub

    Private Sub rb_qdtamtuyen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_PhuTro.CheckedChanged
        If rb_PhuTro.Checked = True Then
            'pnl_div_hdld_baohiem.Visible = False
            'pnl_hdld_baohiem.Visible = False
            lbl_hdld_sohd_soqd.Text = "Số quyết định  "
        End If
    End Sub
#End Region

#Region "---> Events: Bắt Các sự kiện ngoại lệ di chuyển bằng bàn phím <---"
    Private Sub edt_macb_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_macb.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_macb.Text.Trim() <> "") Then
                edt_hoten.Focus()
            Else
                edt_macb.Focus()
            End If
        End If
    End Sub

    Private Sub edt_hoten_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hoten.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hoten.Text.Trim() <> "") Then
                edt_ten_tg.Focus()
                If _FlagEvent = 1 Then
                    edt_ten_tg.Text = edt_hoten.Text
                    edt_bidanh.Text = "Không"
                End If
            Else
                edt_hoten.Focus()
            End If
        End If
    End Sub

    Private Sub edt_hoten_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hoten.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_macb.Focus()
        End If
    End Sub

    Private Sub edt_ten_tg_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ten_tg.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_bidanh.Focus()
        End If
    End Sub

    Private Sub edt_ten_tg_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ten_tg.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hoten.Focus()
        End If
    End Sub

    Private Sub edt_bidanh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bidanh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            rb_nam.Focus()
        End If
    End Sub

    Private Sub edt_bidanh_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bidanh.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_ten_tg.Focus()
        End If
    End Sub

    Private Sub dtpk_ngaysinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngaysinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_donvi.Focus()
        End If
    End Sub

    Private Sub cb_donvi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_donvi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_quoctich.Focus()
        End If
    End Sub

    Private Sub cb_quoctich_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_quoctich.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_dantoc.Focus()
        End If
    End Sub

    Private Sub cb_dantoc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_dantoc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_tongiao.Focus()
        End If
    End Sub

    Private Sub cb_tongiao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tongiao.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_cmt_so.Focus()
        End If
    End Sub

    Private Sub edt_cmt_so_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cmt_so.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_cmt_so.Text.Trim() <> "") Then
                dtpk_cmt_ngaycap.Focus()
            Else
                edt_cmt_so.Focus()
            End If
        End If
    End Sub

    Private Sub edt_cmt_so_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cmt_so.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_tongiao.Focus()
        End If
    End Sub

    Private Sub dtpk_cmt_ngaycap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_cmt_ngaycap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_cmt_noicap.Focus()
        End If
    End Sub

    Private Sub edt_cmt_noicap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cmt_noicap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_ns_tinh
        End If
    End Sub

    Private Sub edt_cmt_noicap_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cmt_noicap.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_cmt_ngaycap.Focus()
        End If
    End Sub

    Private Sub cb_ns_tinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ns_tinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_ns_huyen.Focus()
        End If
    End Sub

    Private Sub cb_ns_huyen_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ns_huyen.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_ns_diachi.Focus()
        End If
    End Sub

    Private Sub edt_ns_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ns_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_nq_tinh.Focus()
        End If
    End Sub

    Private Sub edt_ns_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ns_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_ns_huyen.Focus()
        End If
    End Sub

    Private Sub cb_nq_tinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_nq_tinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_nq_huyen.Focus()
        End If
    End Sub

    Private Sub cb_nq_huyen_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_nq_huyen.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_nq_diachi.Focus()
        End If
    End Sub

    Private Sub edt_nq_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nq_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_tt_tinh.Focus()
        End If
    End Sub

    Private Sub edt_nq_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nq_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_nq_huyen.Focus()
        End If
    End Sub

    Private Sub cb_tt_tinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tt_tinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_tt_huyen.Focus()
        End If
    End Sub

    Private Sub cb_tt_huyen_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tt_huyen.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_tt_diachi.Focus()
        End If
    End Sub

    Private Sub edt_tt_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tt_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_tt_dienthoai.Focus()
        End If
    End Sub

    Private Sub edt_tt_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tt_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_tt_huyen.Focus()
        End If
    End Sub

    Private Sub edt_tt_dienthoai_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tt_dienthoai.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_ttr_tinh.Focus()
        End If
    End Sub

    Private Sub edt_tt_dienthoai_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tt_dienthoai.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_tt_diachi.Focus()
        End If
    End Sub

    Private Sub cb_ttr_tinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ttr_tinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_ttr_huyen.Focus()
        End If
    End Sub

    Private Sub cb_ttr_huyen_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ttr_huyen.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_ttr_diachi.Focus()
        End If
    End Sub

    Private Sub edt_ttr_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ttr_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_ttr_dienthoai.Focus()
        End If
    End Sub

    Private Sub edt_ttr_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ttr_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_ttr_huyen.Focus()
        End If
    End Sub

    Private Sub edt_ttr_dienthoai_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ttr_dienthoai.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dt_didong.Focus()
        End If
    End Sub

    Private Sub edt_ttr_dienthoai_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ttr_dienthoai.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_ttr_diachi.Focus()
        End If
    End Sub

    Private Sub edt_dt_didong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_didong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dt_coquan.Focus()
        End If
    End Sub

    Private Sub edt_dt_didong_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_didong.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_ttr_dienthoai.Focus()
        End If
    End Sub

    Private Sub edt_dt_coquan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_coquan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dt_nharieng.Focus()
        End If
    End Sub

    Private Sub edt_dt_coquan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_coquan.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dt_didong.Focus()
        End If
    End Sub

    Private Sub edt_dt_nharieng_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_nharieng.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_email.Focus()
        End If
    End Sub

    Private Sub edt_dt_nharieng_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_nharieng.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_sofax.Focus()
        End If
    End Sub

    Private Sub edt_sofax_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sofax.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_email.Focus()
        End If
    End Sub

    Private Sub edt_sofax_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sofax.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dt_nharieng.Focus()
        End If
    End Sub

    Private Sub edt_sofax_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_sofax.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 32 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_email_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_email.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_nhommau.Focus()
        End If
    End Sub

    Private Sub edt_email_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_email.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_sofax.Focus()
        End If
    End Sub

    Private Sub cb_nhommau_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_nhommau.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_tp_giadinh.Focus()
        End If
    End Sub

    Private Sub cb_tp_giadinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tp_giadinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_td_vanhoa.Focus()
        End If
    End Sub

    Private Sub cb_td_vanhoa_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_td_vanhoa.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_td_chinhtri.Focus()
        End If
    End Sub

    Private Sub cb_td_chinhtri_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_td_chinhtri.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hocham.Focus()
        End If
    End Sub

    Private Sub cb_hocham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hocham.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hocvi.Focus()
        End If
    End Sub

    Private Sub cb_hocvi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hocvi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_phongban
        End If
    End Sub

    Private Sub cb_phongban_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_phongban.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            tctrl_main.SelectedIndex = 1
            ActiveControl = edt_ms_thue
        End If
    End Sub

    Private Sub edt_ms_thue_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ms_thue.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_ut_giadinh.Focus()
        End If
    End Sub

    Private Sub edt_ms_thue_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ms_thue.KeyUp
        If (e.KeyCode = Keys.Up) Then
            tctrl_main.SelectedIndex = 0
            ActiveControl = cb_hocvi
        End If
    End Sub

    Private Sub cb_ut_giadinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ut_giadinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            clb_ut_banthan.Focus()
        End If
    End Sub

    Private Sub clb_ut_banthan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles clb_ut_banthan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ckb_cachmang.Focus()
        End If
    End Sub

    '----------------------------------------------'

    Private Sub dtpk_cm_ngay_tg_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_cm_ngay_tg.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_cm_tochuc.Focus()
        End If
    End Sub

    Private Sub edt_cm_tochuc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cm_tochuc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ckb_dangvien.Focus()
        End If
    End Sub

    Private Sub edt_cm_tochuc_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cm_tochuc.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_cm_ngay_tg.Focus()
        End If
    End Sub

    Private Sub dtpk_dv_ngayvao_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_dv_ngayvao.ValueChanged
        If (_FlagEvent = 1 And dtpk_dv_ngayvao.Checked = True) Then
            dtpk_dv_ngay_chinhthuc.Value = dtpk_dv_ngayvao.Value.AddYears(1)
        End If
    End Sub

    Private Sub dtpk_dv_ngayvao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_dv_ngayvao.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_dv_ngay_chinhthuc.Focus()
        End If
    End Sub

    Private Sub dtpk_dv_ngay_chinhthuc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_dv_ngay_chinhthuc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_dv_ngayra.Focus()
        End If
    End Sub

    Private Sub dtpk_dv_ngayra_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_dv_ngayra.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_dv_noikn.Focus()
        End If
    End Sub

    Private Sub dtpk_dv_ngayra_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_dv_ngayra.Leave
        If (dtpk_dv_ngayra.Checked = False) Then
            edt_dv_lydo.Text = ""
            edt_dv_lydo.ReadOnly = True
        Else
            edt_dv_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dtpk_dv_ngayra_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtpk_dv_ngayra.MouseUp
        If (dtpk_dv_ngayra.Checked = False) Then
            edt_dv_lydo.Text = ""
            edt_dv_lydo.ReadOnly = True
        Else
            edt_dv_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub edt_dv_noikn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_noikn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dv_sothe.Focus()
        End If
    End Sub

    Private Sub edt_dv_noikn_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_noikn.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_dv_ngayra.Focus()
        End If
    End Sub

    Private Sub edt_dv_sothe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_sothe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dv_nguoi_gt.Focus()
        End If
    End Sub

    Private Sub edt_dv_sothe_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_sothe.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dv_noikn.Focus()
        End If
    End Sub

    Private Sub edt_dv_nguoi_gt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_nguoi_gt.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dv_lydo.Focus()
        End If
    End Sub

    Private Sub edt_dv_nguoi_gt_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_nguoi_gt.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dv_sothe.Focus()
        End If
    End Sub

    Private Sub edt_dv_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ckb_taikhoan.Focus()
        End If
    End Sub

    Private Sub edt_dv_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dv_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dv_nguoi_gt.Focus()
        End If
    End Sub

    Private Sub edt_tk_makh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tk_makh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_tk_sotk.Focus()
        End If
    End Sub

    Private Sub edt_tk_makh_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tk_makh.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ckb_taikhoan.Focus()
        End If
    End Sub

    Private Sub edt_tk_sotk_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tk_sotk.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_tk_tennh.Focus()
        End If
    End Sub

    Private Sub edt_tk_sotk_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tk_sotk.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_tk_makh.Focus()
        End If
    End Sub

    Private Sub edt_tk_tennh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tk_tennh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ckb_bhxh.Focus()
        End If
    End Sub

    Private Sub edt_tk_tennh_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tk_tennh.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_tk_sotk.Focus()
        End If
    End Sub

    Private Sub edt_bhxh_soso_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_soso.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_bhxh_ngaylam.Focus()
        End If
    End Sub

    Private Sub edt_bhxh_soso_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_soso.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ckb_bhxh.Focus()
        End If
    End Sub

    Private Sub dtpk_bhxh_ngaylam_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_bhxh_ngaylam.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_bhxh_ngaydong.Focus()
        End If
    End Sub

    Private Sub dtpk_bhxh_ngaydong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_bhxh_ngaydong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_bhxh_noilam.Focus()
        End If
    End Sub

    Private Sub edt_bhxh_noilam_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_noilam.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_trd_chmon.Focus()
        End If
    End Sub

    Private Sub edt_bhxh_noilam_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_noilam.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_bhxh_ngaydong.Focus()
        End If
    End Sub

    Private Sub cb_trd_chmon_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_trd_chmon.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_trd_ngngu.Focus()
        End If
    End Sub

    Private Sub cb_trd_ngngu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_trd_ngngu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_trd_tinhoc.Focus()
        End If
    End Sub

    Private Sub cb_trd_tinhoc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_trd_tinhoc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_chnganh.Focus()
        End If
    End Sub

    Private Sub cb_chnganh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_chnganh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_st_congtac.Focus()
        End If
    End Sub

    Private Sub edt_st_congtac_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_st_congtac.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_cv_launhat.Focus()
        End If
    End Sub

    Private Sub edt_st_congtac_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_st_congtac.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_chnganh.Focus()
        End If
    End Sub

    Private Sub edt_cv_launhat_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cv_launhat.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_cv_launhat_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cv_launhat.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_st_congtac.Focus()
        End If
    End Sub

    Private Sub edt_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ghichu.KeyDown
        If (e.KeyCode = Keys.Tab) Then
            tctrl_main.SelectedIndex = 2
            ActiveControl = edt_hdld_sohd
        End If
    End Sub

    Private Sub edt_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cv_launhat.Focus()
        End If
    End Sub

    '----------------------------------------------'

    Private Sub edt_hdld_sohd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hdld_sohd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hdld_sohd.Text.Trim() <> "") Then
                dtpk_hdld_ngayky.Focus()
            Else
                edt_hdld_sohd.Focus()
            End If
        End If
    End Sub

    Private Sub dtpk_hdld_ngayky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hdld_ngayky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hdld_nguoiky.Focus()

            If (_IdHdld = "") Then
                dtpk_tg_tungay.Value = dtpk_hdld_ngayky.Value
                dtpk_tg_denngay.Value = dtpk_hdld_ngayky.Value.AddYears(1)
            End If
        End If
    End Sub

    Private Sub dtpk_hdld_ngayky_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_hdld_ngayky.Leave
        If (_IdHdld = "") Then
            dtpk_tg_tungay.Value = dtpk_hdld_ngayky.Value
            dtpk_tg_denngay.Value = dtpk_hdld_ngayky.Value.AddYears(1)
        End If
    End Sub

    Private Sub edt_hdld_nguoiky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hdld_nguoiky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_tg_tungay.Focus()
        End If
    End Sub

    Private Sub edt_hdld_nguoiky_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hdld_nguoiky.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_hdld_ngayky.Focus()
        End If
    End Sub

    Private Sub dtpk_tg_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_tg_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_tg_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_tg_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_tg_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hdld_chucvu.Focus()
        End If
    End Sub

    Private Sub cb_hdld_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hdld_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_gio_batdau.Focus()
        End If
    End Sub

    Private Sub edt_gio_batdau_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gio_batdau.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_gio_ketthuc.Focus()
        End If
    End Sub

    Private Sub edt_gio_ketthuc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gio_ketthuc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hinhthuc_cv.Focus()
        End If
    End Sub

    Private Sub cb_hinhthuc_cv_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hinhthuc_cv.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_ht_traluong.Focus()
        End If
    End Sub

    Private Sub cb_ht_traluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ht_traluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_cdld_nghidinh.Focus()
        End If
    End Sub

    Private Sub cb_cdld_nghidinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cdld_nghidinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_cdld_bangluong.Focus()
        End If
    End Sub

    Private Sub cb_cdld_bangluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cdld_bangluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_cdld_ngachluong.Focus()
        End If
    End Sub

    Private Sub cb_cdld_ngachluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cdld_ngachluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_cdld_bacluong.Focus()
        End If
    End Sub

    Private Sub cb_cdld_bacluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cdld_bacluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_cdld_heso.Focus()
        End If
    End Sub

    Private Sub edt_cdld_heso_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cdld_heso.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_cdld_heso.Text.Trim() = "") Then
                edt_cdld_heso.Text = "0"
            End If
            edt_cdld_tylehuong.Focus()
        End If
    End Sub

    Private Sub edt_cdld_heso_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_cdld_heso.Leave
        If (edt_cdld_heso.Text.Trim() = "") Then
            edt_cdld_heso.Text = "0"
        End If
    End Sub

    Private Sub edt_cdld_heso_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cdld_heso.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_cdld_bacluong.Focus()
        End If
    End Sub

    Private Sub edt_cdld_tylehuong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cdld_tylehuong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_cdld_tylehuong.Text.Trim() = "") Then
                edt_cdld_tylehuong.Text = "0"
            End If
            cb_cdld_chuyenmon.Focus()
        End If
    End Sub

    Private Sub edt_cdld_tylehuong_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cdld_tylehuong.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cdld_heso.Focus()
        End If
    End Sub

    Private Sub cb_cdld_chuyenmon_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cdld_chuyenmon.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_cdld_cv_damnhan.Focus()
        End If
    End Sub

    Private Sub edt_cdld_cv_damnhan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cdld_cv_damnhan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (rb_DinhBien.Checked = True) Then
                pnl_baohiem.Focus()
            Else
                btn_save.Focus()
            End If
        End If
    End Sub

    Private Sub edt_cdld_cv_damnhan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cdld_cv_damnhan.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_cdld_chuyenmon.Focus()
        End If
    End Sub

    Private Sub edt_tienluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tienluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_tienluong.Text.Trim() = "") Then
                edt_tienluong.Text = "0"
            End If
            edt_cdld_cv_damnhan.Focus()
        End If
    End Sub

    Private Sub edt_tienluong_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tienluong.KeyUp
        If (e.KeyCode = Keys.Up) Then
            rb_trongoi.Focus()
        End If
    End Sub

    Private Sub edt_tienluong_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tienluong.Leave
        If (edt_tienluong.Text.Trim() = "") Then
            edt_tienluong.Text = "0"
        End If
    End Sub

    Private Sub edt_tienluong_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tienluong.TextChanged
        Try
            edt_tienluong = formatMoneyinTextbox(edt_tienluong)
        Catch ex As Exception
        End Try
    End Sub





#End Region

End Class