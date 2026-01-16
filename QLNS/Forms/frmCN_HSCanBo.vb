Public Class frmCN_HSCanBo

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Defined parametter and properties <---"
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()

    Private arr_Donvi As ArrayList = New ArrayList
    Private arr_Quoctich As ArrayList = New ArrayList
    Private arr_Dantoc As ArrayList = New ArrayList
    Private arr_Tongiao As ArrayList = New ArrayList

    Private arr_Ns_Tinh As ArrayList = New ArrayList
    Private arr_Ns_Huyen As ArrayList = New ArrayList
    Private arr_Ns_Xa As ArrayList = New ArrayList
    Private arr_Ns_Thon As ArrayList = New ArrayList

    Private arr_Nq_Tinh As ArrayList = New ArrayList
    Private arr_Nq_Huyen As ArrayList = New ArrayList
    Private arr_Nq_Xa As ArrayList = New ArrayList
    Private arr_Nq_Thon As ArrayList = New ArrayList

    Private arr_TT_Tinh As ArrayList = New ArrayList
    Private arr_TT_Huyen As ArrayList = New ArrayList
    Private arr_TT_Xa As ArrayList = New ArrayList
    Private arr_TT_Thon As ArrayList = New ArrayList

    Private arr_TTr_Tinh As ArrayList = New ArrayList
    Private arr_TTr_Huyen As ArrayList = New ArrayList
    Private arr_TTr_Xa As ArrayList = New ArrayList
    Private arr_TTr_Thon As ArrayList = New ArrayList

    Private arr_Tdvh As ArrayList = New ArrayList
    Private arr_Tdct As ArrayList = New ArrayList

    Private arr_Tpgd As ArrayList = New ArrayList
    Private arr_Ut_Banthan As ArrayList = New ArrayList
    Private arr_Ut_Giadinh As ArrayList = New ArrayList

    Private arr_Quocgiaden As ArrayList = New ArrayList
    Private arr_Chucvu As ArrayList = New ArrayList

    Private arr_PL_LLVT As ArrayList = New ArrayList
    Private arr_Quanham As ArrayList = New ArrayList

    Private arr_Hhhv_CHocham As ArrayList = New ArrayList
    Private arr_Hhhv_CHocvi As ArrayList = New ArrayList
    Private ARL_TT_HonNhan As ArrayList = New ArrayList

    'Mảng lưu danh sách các bản ghi đã load ra lưới dữ liệu
    Private arr_Hhhv_LHocham As ArrayList = New ArrayList
    Private arr_Hhhv_LHocvi As ArrayList = New ArrayList

    'Mảng danh sách cần cho tab - Quyết định
    Private arrQd_LoaiQd As ArrayList = New ArrayList
    Private arrQd_Chucvu As ArrayList = New ArrayList
    Private arrQd_Chucvu_cu As ArrayList = New ArrayList
    Private arrQd_ChMon_cu As ArrayList = New ArrayList
    Private arrQd_Donvi_cu As ArrayList = New ArrayList
    Private arrQd_Phongban_cu As ArrayList = New ArrayList
    Private arrQd_Chucvu_moi As ArrayList = New ArrayList
    Private arrQd_ChMon_moi As ArrayList = New ArrayList
    Private arrQd_Donvi_moi As ArrayList = New ArrayList
    Private arrQd_Phongban_moi As ArrayList = New ArrayList

    'Lưu đường dẫn của ảnh thẻ khi chọn ảnh --> Để cập nhật ảnh thẻ
    Private strPath As String = ""
    'Lưu thông tin đường dẫn load ra của ảnh khi sửa đổi
    Private strPathedit As String = ""

    'Biến lưu thông báo việc cập nhật. True: Hồ sơ khác. False: Hồ sơ cán bộ
    Public FlagHS As Boolean

    'Biến lưu thông tin việc người dùng muốn thực hiện Thêm = 1 hay Sửa đổi = 2
    Public FlagEvent As Byte

    'Biến lưu cờ báo --> Khi update lại dữ liệu không chọn ảnh nữa
    Private flagChoice As Boolean = False
    'Biến lưu trạng thái thêm mới của hồ sơ học hàm học vị. True - Thêm mới
    Private flagAdd As Boolean = False

    Private _IdEdit As String = ""
    Private _Code_HocHam As String = ""
    Private _Code_HocVi As String = ""
    Private _IdCanBo As String
    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal value As String)
            _IdCanBo = value
        End Set
    End Property

    Private _Node As String
    Public Property Node() As String
        Get
            Return _Node
        End Get
        Set(ByVal value As String)
            _Node = value
        End Set
    End Property
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    Private _IdCB_TS As String
    Public Property IdCB_TS() As String
        Get
            Return _IdCB_TS
        End Get
        Set(ByVal value As String)
            _IdCB_TS = value
        End Set
    End Property

    ''' <summary>
    ''' Biến lưu giá trị xác định Tab nào được gọi khi right click menu trên lưới dữ liệu
    ''' 0: Công tác. 1: Xuất ngoại. 2: Hộ chiếu. 3: Lực lượng vũ trang. 4: Hồ sơ cũ. 5: Học hàm học vị
    ''' </summary>
    ''' <remarks></remarks>
    Public TabVal As Byte = 0
    'Biến lưu đơn vị Hiện tại của Cán bộ (Load thông tin quyết định nhân sự mới đây nhất của cán bộ)
    Public _IdDonviHT As Integer = 0
#End Region

#Region "---> Functions: Các hàm chính <---"
    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào combobox khi thực hiện load lần đầu tiên
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillData_ComboBox()
        'Thực hiện fill data vào các combobox - Thông tin chung của cán bô
        Dim strSQL As String = ""
        If TRUCTHUOC = 1 Then
            strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And id_goc IN (0,1)"
        Else
            strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And ma_so='" & DONVI.Trim & "' and id_goc <= 1"
        End If
        arr_Donvi.Clear()
        cb_donvi.Items.Clear()
        arr_Donvi = _Globals.Bind_ComBoBox(cb_donvi, strSQL, "---Đơn vị---")
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

        'Lấy dữ liệu đơn vị - Tab Quyết định nhân sự
        strSQL = "Select id,ten_goi from ChiNhanh Where Status = 1 And (ma_so='" & DONVI.Trim & "' or id_goc in (Select id from ChiNhanh where ma_so='" & DONVI.Trim & "'))"
        If (FlagHS = False) Then    'Hồ sơ cán bộ
            'Fill data combobox - Nơi sinh - tỉnh thành
            arr_Ns_Tinh.Clear()
            cb_ns_tinh.Items.Clear()
            arr_Ns_Tinh = _Globals.Bind_ComBoBox(cb_ns_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

            'Fill data combobox - Nguyên quán - tỉnh thành
            arr_Nq_Tinh.Clear()
            cb_nq_tinh.Items.Clear()
            arr_Nq_Tinh = _Globals.Bind_ComBoBox(cb_nq_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

            'Fill data combobox - Thường trú - tỉnh thành
            arr_TT_Tinh.Clear()
            cb_tt_tinh.Items.Clear()
            arr_TT_Tinh = _Globals.Bind_ComBoBox(cb_tt_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

            'Fill data combobox - Tạm trú - tỉnh thành
            arr_TTr_Tinh.Clear()
            cb_ttr_tinh.Items.Clear()
            arr_TTr_Tinh = _Globals.Bind_ComBoBox(cb_ttr_tinh, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - thành phố---")

            'Fill data combobox - Trình độ văn hoá
            arr_Tdvh.Clear()
            cb_tdvh.Items.Clear()
            arr_Tdvh = _Globals.Bind_ComBoBox(cb_tdvh, clsHT_DanhMuc.Sql_Tdvh, "---Trình độ văn hoá---")

            'Fill data combobox - Trình độ chính trị
            arr_Tdct.Clear()
            cb_tdct.Items.Clear()
            arr_Tdct = _Globals.Bind_ComBoBox(cb_tdct, clsHT_DanhMuc.Sql_Tdct, "---Trình độ trính trị---")

            'Fill data combobox - Tình trạng hôn nhân
            ARL_TT_HonNhan.Clear()
            cb_tt_honnhan.Items.Clear()
            ARL_TT_HonNhan = _Globals.Bind_ComBoBox(cb_tt_honnhan, clsHT_DanhMuc.Sql_TTrangHonNhan, "---Tình trạng hôn nhân---")

            'Fill data combobox - Thành phần gia đình
            arr_Tpgd.Clear()
            cb_tpgd.Items.Clear()
            arr_Tpgd = _Globals.Bind_ComBoBox(cb_tpgd, clsHT_DanhMuc.Sql_Tpgd, "---Thành phần gia đình---")

            'Fill data combobox - Ưu tiên gia đình
            arr_Ut_Giadinh.Clear()
            cb_ut_giadinh.Items.Clear()
            arr_Ut_Giadinh = _Globals.Bind_ComBoBox(cb_ut_giadinh, clsHT_DanhMuc.Sql_Ut_giadinh, "---Ưu tiên gia đình---")

            'Fill data combobox - Ưu tiên bản thân
            clb_ut_banthan.Items.Clear()
            arr_Ut_Banthan.Clear()
            arr_Ut_Banthan = _Globals.FillData_CheckedListBox(clb_ut_banthan, clsHT_DanhMuc.Sql_Ut_banthan)

            arrQd_LoaiQd.Clear()
            cb_qd_loaiqd.Items.Clear()
            arrQd_LoaiQd = _Globals.Bind_ComBoBox(cb_qd_loaiqd, clsHT_DanhMuc.Sql_LoaiQd_Ns, "---Loại quyết định---")

            arrQd_Chucvu.Clear()
            cb_qd_chucvu.Items.Clear()
            arrQd_Chucvu = _Globals.Bind_ComBoBox(cb_qd_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ người ký---")

            'Load danh sách bản ghi chức vụ với
            arrQd_Chucvu_moi.Clear()
            cb_qd_chucvu_moi.Items.Clear()
            arrQd_Chucvu_moi = _Globals.Bind_ComBoBox(cb_qd_chucvu_moi, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ---")

            'Load danh sách chuyên môn Nghiệp vụ
            arrQd_ChMon_moi.Clear()
            cb_qd_cm_moi.Items.Clear()
            arrQd_ChMon_moi = _Globals.Bind_ComBoBox(cb_qd_cm_moi, clsHT_DanhMuc.Sql_Chuyenmon, "---Chuyên môn nghiệp vụ---")

        Else                        'Hồ sơ khác
            'Bind dữ liệu combobox - Quốc gia đến
            arr_Quocgiaden.Clear()
            cb_xn_nuocden.Items.Clear()
            arr_Quocgiaden = _Globals.Bind_ComBoBox(cb_xn_nuocden, clsHT_DanhMuc.Sql_Quocgia, "---Quốc gia đến---")

            arr_Chucvu.Clear()
            cb_xn_chucvu.Items.Clear()
            arr_Chucvu = _Globals.Bind_ComBoBox(cb_xn_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ---")

            arr_PL_LLVT.Clear()
            cb_llvt_phanloai.Items.Clear()
            arr_PL_LLVT = _Globals.Bind_ComBoBox(cb_llvt_phanloai, clsHT_DanhMuc.Sql_PL_LLVT, "---Phân loại---")

            arr_Quanham.Clear()
            cb_llvt_quanham.Items.Clear()
            arr_Quanham = _Globals.Bind_ComBoBox(cb_llvt_quanham, clsHT_DanhMuc.Sql_Quanham, "---Quân hàm---")

            'Bind dữ liệu học hàm
            arr_Hhhv_CHocham.Clear()
            cb_hhhv_hocham.Items.Clear()
            arr_Hhhv_CHocham = _Globals.Bind_ComBoBox(cb_hhhv_hocham, clsHT_DanhMuc.Sql_Hocham, "---Học hàm---")
            'Bind dữ liệu học vị
            arr_Hhhv_CHocvi.Clear()
            cb_hhhv_hocvi.Items.Clear()
            arr_Hhhv_CHocvi = _Globals.Bind_ComBoBox(cb_hhhv_hocvi, clsHT_DanhMuc.Sql_Hocvi, "---Học vị---")
        End If

    End Sub

    ''' <summary>
    ''' Hàm thực hiện Reset toàn bộ controls về trạng thái ban đầu
    ''' </summary>
    ''' <param name="state">False: Reset ->Hồ sơ cán bộ. True: Reset ->Hồ sơ khác</param>
    ''' <param name="flag">True: Thực hiện reset thông tin chung của hồ sơ cán bộ. False: Bỏ qua phần này</param>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls(ByVal state As Boolean, ByVal flag As Boolean)
        'Reset các controls được dùng chung trong 2 trường hợp
        If (flag = True) Then
            'edt_macb.Text = setMaCanBo(IdDONVI)
            edt_macb.Text = setMaCanBo(_IdDonviHT)
            edt_hoten.Text = ""
            edt_ten_tg.Text = ""
            edt_bidanh.Text = ""
            rb_nam.Checked = False
            rb_nu.Checked = False
            dtpk_ngaysinh.Text = DateTime.Now.ToShortDateString()
            If cb_donvi.Items.Count <> 0 And _IdCB_TS = "" And FlagEvent = 1 Then
                If _IdDonviHT > 0 Then
                    cb_donvi.SelectedIndex = IIf(_IdDonviHT.ToString() <> "", CType(arr_Donvi.IndexOf(_IdDonviHT.ToString()), Integer), 0)
                    cb_donvi_SelectedIndexChanged(Nothing, Nothing)
                Else
                    cb_donvi.SelectedIndex = 1
                End If
            Else
                cb_donvi.SelectedIndex = 0
            End If
            cb_donvi_SelectedIndexChanged(Nothing, Nothing)
            If FlagEvent = 2 Then cb_donvi.Enabled = False
            If cb_quoctich.Items.Count <> 0 And _IdCB_TS = "" And FlagEvent = 1 Then
                cb_quoctich.SelectedIndex = 1
            Else
                cb_quoctich.SelectedIndex = 0
            End If

            If cb_dantoc.Items.Count <> 0 And _IdCB_TS = "" And FlagEvent = 1 Then
                cb_dantoc.SelectedIndex = 1
            Else
                cb_dantoc.SelectedIndex = 0
            End If

            If cb_tongiao.Items.Count <> 0 And _IdCB_TS = "" And FlagEvent = 1 Then
                cb_tongiao.SelectedIndex = 1
            Else
                cb_tongiao.SelectedIndex = 0
            End If

            edt_cmt_so.Text = ""
            dtpk_cmt_ngaycap.Text = DateTime.Now.ToShortDateString()
            edt_cmt_noicap.Text = ""
        End If

        Select Case state
            Case False  'Reset hồ sơ nhân sự
                strPath = ""
                strPathedit = ""
                flagChoice = False
                'Thông tin về nơi sinh - cb_ns_tinh
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

                'Thông tin về Nguyên quán của Cán bộ
                If cb_nq_tinh.Items.Count <> 0 Then
                    cb_nq_tinh.SelectedIndex = 0
                End If
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

                'Thông tin về Thường trú của cán bộ
                If cb_tt_tinh.Items.Count <> 0 Then
                    cb_tt_tinh.SelectedIndex = 0
                End If
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
                'Thông tin về Tạm trú của cán bộ
                If cb_ttr_tinh.Items.Count <> 0 Then
                    cb_ttr_tinh.SelectedIndex = 0
                End If
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

                edt_dt_coquan.Text = ""
                edt_dt_didong.Text = ""
                edt_dt_nharieng.Text = ""

                edt_email.Text = ""
                cb_nhommau.SelectedIndex = 0

                cb_tdvh.SelectedIndex = 0
                cb_tdct.SelectedIndex = 0

                edt_ghichu.Text = ""
                dtpk_ngay_thamnien.Text = DateTime.Now.ToShortDateString()
                dtpk_ngay_thamnien.Checked = False
                dtpk_ngay_vaonganh.Text = DateTime.Now.ToShortDateString()
                dtpk_ngay_vaonganh.Checked = False
                dtpk_ngay_vaonhcsxh.Text = DateTime.Now.ToShortDateString()
                dtpk_ngay_vaonhcsxh.Checked = False
                dtpk_ngay_bc_nhcsxh.Text = DateTime.Now.ToShortDateString()
                dtpk_ngay_bc_nhcsxh.Checked = False
                dtpk_ngay_cachmang.Text = DateTime.Now.ToShortDateString()
                dtpk_ngay_cachmang.Checked = False
                edt_cm_tochuc.Text = ""
                edt_sotruong_ct.Text = ""
                edt_congviec_lamlau.Text = ""

                edt_bhxh_so.Text = ""
                dtpk_bhxh_ngaylam.Text = DateTime.Now.ToShortDateString()
                dtpk_bhxh_ngaylam.Checked = False
                dtpk_bhxh_ngaydong.Text = DateTime.Now.ToShortDateString()
                dtpk_bhxh_ngaydong.Checked = False
                edt_bhxh_noilam.Text = ""

                'Tab - Bản thân và mối quan hệ
                cb_tpgd.SelectedIndex = 0
                cb_ut_giadinh.SelectedIndex = 0
                _Globals.ResetItems_CheckedListBox(clb_ut_banthan)

                edt_dacdiem_banthan.Text = ""
                edt_quanhe_ng_nn.Text = ""

                'Tab - Quyết định nhân sự
                edt_qd_soqd.Text = ""
                If (_IdCB_TS <> "") Then
                    For i As Integer = 0 To cb_qd_loaiqd.Items.Count - 1
                        If (cb_qd_loaiqd.Items(i).ToString().Trim() = "Tuyển dụng cán bộ mới") Then
                            cb_qd_loaiqd.SelectedIndex = i
                            Exit For
                        End If
                    Next
                Else
                    cb_qd_loaiqd.SelectedIndex = 0
                End If
                edt_qd_nguoiky.Text = ""
                cb_qd_chucvu.SelectedIndex = 0
                If (FlagEvent = 1 And _IdCB_TS = "") Then
                    'Lấy thông tin có sẵn về Người ký và Chức vụ người ký
                    edt_qd_nguoiky.Text = _HS_CanBo.GetVarNam("GIAMDOC")
                    Dim _IdIndex As Int32 = CType(arr_Donvi(IIf(cb_donvi.SelectedIndex > 0, cb_donvi.SelectedIndex, "0")), Integer)
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
                                                cb_qd_chucvu.SelectedIndex = CType(arrQd_Chucvu.IndexOf(db_child.Rows(0)("id").ToString().Trim()), Integer)
                                            End If
                                        End If
                                    End If
                                End Using
                            End If
                        End If
                    End Using
                End If
                dtpk_qd_ngayky.Text = DateTime.Now.ToShortDateString()
                dtpk_qd_ngayhl.Text = DateTime.Now.ToShortDateString()
                cb_qd_chucvu_moi.SelectedIndex = 0
                cb_qd_cm_moi.SelectedIndex = 0

            Case True   'Reset hồ sơ khác
                _IdEdit = ""
                _Code_HocHam = ""
                _Code_HocVi = ""
                flagAdd = False
                'Reset Tab - Công Tác
                dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
                dtpk_ct_denngay.Checked = False
                edt_ct_chucvu.Text = ""
                edt_ct_diachi.Text = ""
                edt_ct_lydo.Text = ""
                edt_ct_ghichu.Text = ""
                dgv_congtac.Rows.Clear()

                'Reset Tab Xuất Ngoại
                dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()

                cb_xn_nuocden.SelectedIndex = 0

                edt_xn_mucdich.Text = ""
                edt_xn_soqd.Text = ""
                dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
                edt_xn_nguoiky.Text = ""
                edt_xn_ghichu.Text = ""

                cb_xn_chucvu.SelectedIndex = 0
                dgv_xuatngoai.Rows.Clear()

                'Reset Tab Hộ Chiếu
                edt_hc_sohc.Text = ""
                cb_hc_loaihc.SelectedIndex = 0
                dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
                edt_hc_noicap.Text = ""
                dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
                cb_hc_tinhtrang.SelectedIndex = 0
                edt_hc_ghichu.Text = ""
                dgv_hochieu.Rows.Clear()

                'Reset Tab - Lực lượng vũ trang
                dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()

                cb_llvt_phanloai.SelectedIndex = 0
                cb_llvt_quanham.SelectedIndex = 0

                edt_llvt_chucvu.Text = ""
                edt_llvt_donvi.Text = ""
                edt_llvt_ghichu.Text = ""
                dgv_llvt.Rows.Clear()

                'Reset Tab Hồ sơ cũ
                dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
                dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
                edt_hscu_diachi.Text = ""
                edt_hscu_nghenghiep.Text = ""
                edt_hscu_ghichu.Text = ""
                dgv_hosocu.Rows.Clear()

                'Reset Tab - Học hàm học vị
                cb_hhhv_hocham.SelectedIndex = 0
                cb_hhhv_hocvi.SelectedIndex = 0

                'Clear các CheckListBox
                arr_Hhhv_LHocham.Clear()
                arr_Hhhv_LHocvi.Clear()
                clb_hhhv_hocham.Items.Clear()
                clb_hhhv_hocvi.Items.Clear()
                ckb_hocham.Checked = False
                ckb_hocham_CheckedChanged(Nothing, Nothing)
                ckb_hocvi.Checked = False
                ckb_hocvi_CheckedChanged(Nothing, Nothing)
                ckb_hocham.Enabled = False
                ckb_hocvi.Enabled = False
            Case Else
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện thiết lập trạng thái các điều khiển - Lock and UnLock controls
    ''' </summary>
    ''' <param name="status">True: Khoá controls. False: Mở khoá controls</param>
    ''' <param name="indexTab">0: Phần thông tin chung. 1 - Quyết định nhân sự</param>
    ''' <remarks></remarks>
    Private Sub SetStatus_Controls(ByVal status As Boolean, ByVal indexTab As Byte)
        If (indexTab = 0) Then
            edt_macb.ReadOnly = status
            edt_hoten.ReadOnly = status
            edt_ten_tg.ReadOnly = status
            edt_bidanh.ReadOnly = status
            pnl_sex.Enabled = Not (status)
            dtpk_ngaysinh.Enabled = Not (status)
            cb_donvi.Enabled = Not (status)
            cb_quoctich.Enabled = Not (status)
            cb_dantoc.Enabled = Not (status)
            cb_tongiao.Enabled = Not (status)
            edt_cmt_so.ReadOnly = status
            dtpk_cmt_ngaycap.Enabled = Not (status)
            edt_cmt_noicap.ReadOnly = status
            btn_chonanh.Enabled = Not (status)
            btn_xoaanh.Enabled = Not (status)

            edt_qd_soqd.Enabled = Not (status)
            edt_qd_nguoiky.Enabled = Not (status)
            edt_qd_ghichu.Enabled = Not (status)

        ElseIf (indexTab = 1) Then
            If status Then
                If tctrl_hs_cb.TabPages.Contains(tp_qd_nhansu) Then tctrl_hs_cb.TabPages.Remove(tp_qd_nhansu)
            End If
            edt_qd_soqd.ReadOnly = status
            cb_qd_loaiqd.Enabled = Not (status)
            dtpk_qd_ngayky.Enabled = Not (status)
            dtpk_qd_ngayhl.Enabled = Not (status)
            edt_qd_nguoiky.ReadOnly = status

            cb_qd_chucvu.Enabled = Not (status)
            cb_qd_donvi_moi.Enabled = Not (status)
            cb_qd_phongban_moi.Enabled = Not (status)
            cb_qd_chucvu_moi.Enabled = Not (status)
            cb_qd_cm_moi.Enabled = Not (status)
            edt_qd_ghichu.ReadOnly = status

            edt_qd_soqd.Enabled = Not (status)
            edt_qd_nguoiky.Enabled = Not (status)
            edt_qd_ghichu.Enabled = Not (status)
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - fill dữ liệu các hồ sơ theo id cán bộ truyền vào
    ''' </summary>
    ''' <param name="_IndexDocument">Chỉ số xác định hồ sơ cần fill data
    '''              0: Hồ sơ cán bộ
    '''              1: Quyết định nhân sự
    '''              2: Hồ sơ Công tác
    '''              3: Hồ sơ Xuất ngoại
    '''              4: Hồ sơ Hộ chiếu
    '''              5: Hồ sơ Tham gia Lực lượng vũ trang
    '''              6: Hồ sơ cũ (Thông tin của cán bộ khi chưa thoát ly)
    '''              7: Hồ sơ học hàm học vị của cán bộ
    ''' <param name="status">True-fill thông tin chung. False-Không fill phần thông tin chung</param>
    ''' </param>
    ''' <remarks></remarks>
    Private Sub Fill_Data(ByVal _IndexDocument As Byte, ByVal status As Boolean)
        'Fill data - Thông tin chung của cán bộ
        If (status = True) Then
            Dim dr As DataRow
            dr = _HS_CanBo.GetHuman_ForCode(_IdCanBo)
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    edt_macb.Text = dr("MaCB").ToString()
                    If dr("IdOld").ToString <> "" Then
                        edt_macb.Enabled = False
                    Else
                        edt_macb.Enabled = True
                    End If
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
                    cb_donvi.SelectedIndex = IIf(dr("IdDonVi").ToString() <> "", CType(arr_Donvi.IndexOf(dr("IdDonVi").ToString()), Int32), 0)
                    cb_donvi_SelectedIndexChanged(Nothing, Nothing)
                    cb_quoctich.SelectedIndex = IIf(dr("IdQuocTich").ToString() <> "", CType(arr_Quoctich.IndexOf(dr("IdQuocTich").ToString()), Int32), 0)
                    cb_dantoc.SelectedIndex = IIf(dr("IdDanToc").ToString() <> "", CType(arr_Dantoc.IndexOf(dr("IdDanToc").ToString()), Int32), 0)
                    cb_tongiao.SelectedIndex = IIf(dr("IdTonGiao").ToString() <> "", CType(arr_Tongiao.IndexOf(dr("IdTonGiao").ToString()), Int32), 0)
                    edt_cmt_so.Text = dr("CMT_So").ToString().Trim()
                    dtpk_cmt_ngaycap.Value = CType(dr("CMT_NgayCap").ToString(), DateTime)
                    edt_cmt_noicap.Text = dr("CMT_NoiCap").ToString().Trim()
                    If (FlagHS = True) Then    'Load thông tin Hồ sơ khác
                        If (dr("IdNew").ToString().Trim() <> "") Then
                            'Thực hiện ẩn hai Chức năng liên quan đến cập nhật
                            btn_add.Visible = False
                            btn_save.Visible = False
                        End If
                    End If
                    'Gọi hàm Show picture của cán bộ công nhân viên
                    If strPathedit = "" Then strPathedit = dr("AnhThe").ToString()
                    '_HS_CanBo.Show_Picture(dr("AnhThe").ToString(), picbx_main)
                    My_LoadImage(picbx_main, dr("AnhThe"))
                End If
            End If
        End If

        'Fill data - các hồ sơ theo chỉ số truyền vào
        Select Case _IndexDocument
            Case 0  'Fill - Hồ sơ cán bộ
                Dim dr As DataRow
                dr = _HS_CanBo.GetHuman_ForCode(_IdCanBo)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        cb_ns_tinh.SelectedIndex = IIf(dr("IdNS_Tinh").ToString() <> "", CType(arr_Ns_Tinh.IndexOf(dr("IdNS_Tinh").ToString()), Int32), 0)
                        cb_ns_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdNS_Huyen").ToString() <> "") Then
                            cb_ns_huyen.SelectedIndex = CType(arr_Ns_Huyen.IndexOf(dr("IdNS_Huyen").ToString()), Int32)
                        End If
                        If (dr("IdNS_Xa").ToString() <> "") Then
                            cb_ns_xa.SelectedIndex = CType(arr_Ns_Xa.IndexOf(dr("IdNS_Xa").ToString()), Int32)
                        End If
                        cb_ns_xa_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdNS_Thon").ToString() <> "") Then
                            cb_ns_thon.SelectedIndex = CType(arr_Ns_Thon.IndexOf(dr("IdNS_Thon").ToString()), Int32)
                        End If
                        edt_ns_diachi.Text = dr("NS_DChi").ToString().Trim()

                        cb_nq_tinh.SelectedIndex = IIf(dr("IdNQ_Tinh").ToString() <> "", CType(arr_Nq_Tinh.IndexOf(dr("IdNQ_Tinh").ToString()), Int32), 0)
                        cb_nq_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdNQ_Huyen").ToString() <> "") Then
                            cb_nq_huyen.SelectedIndex = CType(arr_Nq_Huyen.IndexOf(dr("IdNQ_Huyen").ToString()), Int32)
                        End If
                        If (dr("IdNQ_Xa").ToString() <> "") Then
                            cb_nq_xa.SelectedIndex = CType(arr_Nq_Xa.IndexOf(dr("IdNQ_Xa").ToString()), Int32)
                        End If
                        cb_nq_xa_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdNQ_Thon").ToString() <> "") Then
                            cb_nq_thon.SelectedIndex = CType(arr_Nq_Thon.IndexOf(dr("IdNQ_Thon").ToString()), Int32)
                        End If
                        edt_nq_diachi.Text = dr("NQ_DChi").ToString().Trim()

                        cb_tt_tinh.SelectedIndex = IIf(dr("IdThT_Tinh").ToString() <> "", CType(arr_TT_Tinh.IndexOf(dr("IdThT_Tinh").ToString()), Int32), 0)
                        cb_tt_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdThT_Huyen").ToString() <> "") Then
                            cb_tt_huyen.SelectedIndex = CType(arr_TT_Huyen.IndexOf(dr("IdThT_Huyen").ToString()), Int32)
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
                        cb_ttr_tinh.SelectedIndex = IIf(dr("IdTTr_Tinh").ToString() <> "", CType(arr_TTr_Tinh.IndexOf(dr("IdTTr_Tinh").ToString()), Int32), 0)
                        cb_ttr_tinh_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdTTr_Huyen").ToString() <> "") Then
                            cb_ttr_huyen.SelectedIndex = CType(arr_TTr_Huyen.IndexOf(dr("IdTTr_Huyen").ToString()), Int32)
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
                        edt_sofax.Text = dr("SoFax").ToString().Trim()
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
                        'Tab - Bản thân - mối quan hệ
                        cb_tpgd.SelectedIndex = IIf(dr("IdThanhPhanGD").ToString() <> "", CType(arr_Tpgd.IndexOf(dr("IdThanhPhanGD").ToString()), Int32), 0)
                        cb_ut_giadinh.SelectedIndex = IIf(dr("IdUT_GDinh").ToString() <> "", CType(arr_Ut_Giadinh.IndexOf(dr("IdUT_GDinh").ToString()), Int32), 0)
                        'Load dữ liệu Ưu tiên bản thân
                        If (dr("IdUT_BThan").ToString() <> "") Then
                            _Globals.SetItems_CheckedListBox(clb_ut_banthan, dr("IdUT_BThan").ToString(), arr_Ut_Banthan)
                        End If

                        edt_dacdiem_banthan.Text = dr("DacDiem_BT").ToString().Trim()
                        edt_quanhe_ng_nn.Text = dr("QuanHe_Nguoi_NN").ToString().Trim()
                        'Tab - Thông tin hồ sơ khác
                        cb_tdvh.SelectedIndex = IIf(dr("IdTrinhDoVH").ToString() <> "", CType(arr_Tdvh.IndexOf(dr("IdTrinhDoVH").ToString()), Int32), 0)
                        cb_tdct.SelectedIndex = IIf(dr("IdTrinhDoCT").ToString() <> "", CType(arr_Tdct.IndexOf(dr("IdTrinhDoCT").ToString()), Int32), 0)
                        cb_tt_honnhan.SelectedIndex = IIf(My_CStr(dr("HonNhan_Cd").ToString(), "") <> "", ARL_TT_HonNhan.IndexOf(My_CStr(dr("HonNhan_Cd").ToString(), "")), 0)

                        If dr("Ngay_ThamNien").ToString().Trim() <> "" Then
                            If (CType(dr("Ngay_ThamNien"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_ngay_thamnien.Checked = False
                            Else
                                dtpk_ngay_thamnien.Checked = True
                                dtpk_ngay_thamnien.Value = CType(dr("Ngay_ThamNien").ToString(), DateTime)
                            End If
                        Else
                            dtpk_ngay_thamnien.Checked = False
                        End If
                        If dr("Ngay_NH").ToString().Trim() <> "" Then
                            If (CType(dr("Ngay_NH"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_ngay_vaonganh.Checked = False
                            Else
                                dtpk_ngay_vaonganh.Checked = True
                                dtpk_ngay_vaonganh.Value = CType(dr("Ngay_NH").ToString(), DateTime)
                            End If
                        Else
                            dtpk_ngay_vaonganh.Checked = False
                        End If
                        If dr("Ngay_VBSP").ToString().Trim() <> "" Then
                            If (CType(dr("Ngay_VBSP"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_ngay_vaonhcsxh.Checked = False
                            Else
                                dtpk_ngay_vaonhcsxh.Checked = True
                                dtpk_ngay_vaonhcsxh.Value = CType(dr("Ngay_VBSP").ToString(), DateTime)
                            End If
                        Else
                            dtpk_ngay_vaonhcsxh.Checked = False
                        End If
                        If dr("Ngay_BienChe").ToString().Trim() <> "" Then
                            If (CType(dr("Ngay_BienChe"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_ngay_bc_nhcsxh.Checked = False
                            Else
                                dtpk_ngay_bc_nhcsxh.Checked = True
                                dtpk_ngay_bc_nhcsxh.Value = CType(dr("Ngay_BienChe").ToString(), DateTime)
                            End If
                        Else
                            dtpk_ngay_bc_nhcsxh.Checked = False
                        End If
                        If dr("CM_Ngay").ToString().Trim() <> "" Then
                            If (CType(dr("CM_Ngay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_ngay_cachmang.Checked = False
                            Else
                                dtpk_ngay_cachmang.Checked = True
                                dtpk_ngay_cachmang.Value = CType(dr("CM_Ngay").ToString(), DateTime)
                            End If
                        Else
                            dtpk_ngay_cachmang.Checked = False
                        End If
                        edt_cm_tochuc.Text = dr("CM_ToChuc").ToString().Trim()
                        edt_sotruong_ct.Text = dr("SoTruong_CT").ToString().Trim()
                        edt_congviec_lamlau.Text = dr("CV_Lau").ToString().Trim()
                        edt_nh_makh.Text = dr("NH_MaKH").ToString().Trim()
                        edt_nh_sotk.Text = dr("NH_SoTK").ToString().Trim()
                        edt_nh_tennh.Text = dr("NH_TenNH").ToString().Trim()
                        edt_masothue.Text = dr("MaSoThue").ToString().Trim()

                        edt_bhxh_so.Text = dr("BHXH_SoSo").ToString().Trim()
                        If dr("BHXH_NgaySo").ToString().Trim() <> "" Then
                            If (CType(dr("BHXH_NgaySo"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_bhxh_ngaylam.Checked = False
                            Else
                                dtpk_bhxh_ngaylam.Checked = True
                                dtpk_bhxh_ngaylam.Value = CType(dr("BHXH_NgaySo").ToString(), DateTime)
                            End If
                        Else
                            dtpk_bhxh_ngaylam.Checked = False
                        End If
                        If dr("BHXH_NgayBatDau").ToString().Trim() <> "" Then
                            If (CType(dr("BHXH_NgayBatDau"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_bhxh_ngaydong.Checked = False
                            Else
                                dtpk_bhxh_ngaydong.Checked = True
                                dtpk_bhxh_ngaydong.Value = CType(dr("BHXH_NgayBatDau").ToString(), DateTime)
                            End If
                        Else
                            dtpk_bhxh_ngaydong.Checked = False
                        End If
                        edt_bhxh_noilam.Text = dr("BHXH_NoiLam").ToString().Trim()
                        edt_ghichu.Text = dr("GhiChu").ToString().Trim()
                    End If
                End If

            Case 1  'Fill - Quyết định nhân sự
                Dim m_QDNhansu As QDNhanSu = New QDNhanSu
                m_QDNhansu = m_QDNhansu.getFinalRecord(_IdCanBo, Nothing, True)
                edt_qd_soqd.Text = m_QDNhansu.So_QD
                cb_qd_loaiqd.SelectedIndex = IIf(m_QDNhansu.IdLoaiQD.ToString() <> "", CType(arrQd_LoaiQd.IndexOf(m_QDNhansu.IdLoaiQD.ToString()), Int32), 0)
                dtpk_qd_ngayky.Value = m_QDNhansu.NgayKy_QD
                dtpk_qd_ngayhl.Value = m_QDNhansu.NgayHL
                edt_qd_nguoiky.Text = m_QDNhansu.NguoiKy_QD
                cb_qd_chucvu.SelectedIndex = IIf(m_QDNhansu.idCV_Nguoiky_QD.ToString() <> "", CType(arrQd_Chucvu.IndexOf(m_QDNhansu.idCV_Nguoiky_QD.ToString()), Int32), 0)
                cb_qd_donvi_moi.SelectedIndex = IIf(m_QDNhansu.IdDonvi_Moi.ToString() <> "", CType(arrQd_Donvi_moi.IndexOf(m_QDNhansu.IdDonvi_Moi.ToString()), Int32), 0)
                cb_qd_donvi_moi_SelectedIndexChanged(Nothing, Nothing)
                cb_qd_phongban_moi.SelectedIndex = IIf(m_QDNhansu.IdPhong_Moi.ToString() <> "", CType(arrQd_Phongban_moi.IndexOf(m_QDNhansu.IdPhong_Moi.ToString()), Int32), 0)
                cb_qd_chucvu_moi.SelectedIndex = IIf(m_QDNhansu.IdChucvu_Moi.ToString() <> "", CType(arrQd_Chucvu_moi.IndexOf(m_QDNhansu.IdChucvu_Moi.ToString()), Int32), 0)
                cb_qd_cm_moi.SelectedIndex = IIf(m_QDNhansu.IdChuyenMon_Moi.ToString() <> "", CType(arrQd_ChMon_moi.IndexOf(m_QDNhansu.IdChuyenMon_Moi.ToString()), Int32), 0)
                edt_qd_ghichu.Text = m_QDNhansu.GhiChu

                'Using db As DataTable = _HS_CanBo.GetAll_Decision(_IdCanBo, _IdDonviHT, 0)
                '    If Not (db Is Nothing) Then
                '        If (db.Rows.Count > 0) Then
                '            edt_qd_soqd.Text = db.Rows(0)("So_QD").ToString()
                '            cb_qd_loaiqd.SelectedIndex = IIf(db.Rows(0)("IdLoaiQD").ToString() <> "", CType(arrQd_LoaiQd.IndexOf(db.Rows(0)("IdLoaiQD").ToString()), Int32), 0)
                '            If db.Rows(0)("NgayKy_QD").ToString().Trim() <> "" Then
                '                dtpk_qd_ngayky.Value = CType(db.Rows(0)("NgayKy_QD").ToString(), DateTime)
                '            End If

                '            If db.Rows(0)("NgayHL").ToString().Trim() <> "" Then
                '                dtpk_qd_ngayhl.Value = CType(db.Rows(0)("NgayHL").ToString(), DateTime)
                '            End If

                '            edt_qd_nguoiky.Text = db.Rows(0)("NguoiKy_QD").ToString()
                '            cb_qd_chucvu.SelectedIndex = IIf(db.Rows(0)("IdCV_Nguoiky_QD").ToString() <> "", CType(arrQd_Chucvu.IndexOf(db.Rows(0)("IdCV_Nguoiky_QD").ToString()), Int32), 0)
                '            cb_qd_donvi_moi.SelectedIndex = IIf(db.Rows(0)("IdDonVi_Moi").ToString() <> "", CType(arrQd_Donvi_moi.IndexOf(db.Rows(0)("IdDonVi_Moi").ToString()), Int32), 0)
                '            cb_qd_donvi_moi_SelectedIndexChanged(Nothing, Nothing)
                '            cb_qd_phongban_moi.SelectedIndex = IIf(db.Rows(0)("IdPhong_Moi").ToString() <> "", CType(arrQd_Phongban_moi.IndexOf(db.Rows(0)("IdPhong_Moi").ToString()), Int32), 0)
                '            cb_qd_chucvu_moi.SelectedIndex = IIf(db.Rows(0)("IdChucVu_Moi").ToString() <> "", CType(arrQd_Chucvu_moi.IndexOf(db.Rows(0)("IdChucVu_Moi").ToString()), Int32), 0)
                '            cb_qd_cm_moi.SelectedIndex = IIf(db.Rows(0)("IdChuyenMon_Moi").ToString() <> "", CType(arrQd_ChMon_moi.IndexOf(db.Rows(0)("IdChuyenMon_Moi").ToString()), Int32), 0)
                '            edt_qd_ghichu.Text = db.Rows(0)("GhiChu").ToString()
                '        End If
                '    End If
                'End Using

            Case 2  'Fill - Hồ sơ công tác
                dgv_congtac.Rows.Clear()
                Using db As DataTable = listQDNhansu(_IdCanBo, 2)
                    If db.Rows.Count > 0 Then
                        Dim row As Integer = 0
                        For i As Integer = 0 To db.Rows.Count - 1
                            Dim ThoiGian As String = ""
                            If CInt(db.Rows(i)("IsQD_NHCS")) = 0 Then
                                If i = 0 Then
                                    dgv_congtac.Rows.Add()
                                    dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                    ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                    Dim j As Integer = i + 1
                                    While j < db.Rows.Count
                                        Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536"
                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                i = j - 1
                                                Exit While
                                            Case Else
                                                j = j + 1
                                        End Select
                                    End While
                                    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                    row = row + 1
                                Else
                                    If Not (db.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
                                            Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
                                            Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
                                        dgv_congtac.Rows.Add()
                                        dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                        ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        If Not (db.Rows(i)("DenNgay") Is DBNull.Value) Then
                                            If Not (CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                                dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                            Else
                                                If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                                Else
                                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                End If
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                Dim j As Integer = i + 1
                                                While j < db.Rows.Count
                                                    Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536"
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            i = j - 1
                                                            Exit While
                                                        Case Else
                                                            j = j + 1
                                                    End Select
                                                End While
                                                dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                            End If
                                        Else
                                            If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                            Else
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                            End If
                                            dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                            Dim j As Integer = i + 1
                                            While j < db.Rows.Count
                                                Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                                    Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536"
                                                        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                        i = j - 1
                                                        Exit While
                                                    Case Else
                                                        j = j + 1
                                                End Select
                                            End While
                                            dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                        End If
                                        row = row + 1
                                    End If
                                End If
                            Else
                                If db.Rows(i)("MaLoaiQD").ToString().Trim() <> "1519" Then
                                    dgv_congtac.Rows.Add()
                                    dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                    ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                    Dim j As Integer = i + 1
                                    While j < db.Rows.Count
                                        Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                            Case "1534", "1535"
                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                Exit While
                                            Case "1518"
                                                'Bỏ qua nếu QĐ tiếp theo là QĐ kiêm nhiệm
                                                Exit While
                                            Case "1519"
                                                'Lấy mốc hết thời gian kiêm nhiệm
                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                j = j + 1
                                                Exit While
                                            Case Else
                                                'TH không đọc tiếp nếu: Nếu đơn vị, phòng ban, chức vụ cũ; hoặc QĐ tiếp theo là QĐ kiêm nhiệm
                                                If (CInt(db.Rows(i)("IDDonVi_Moi")) = CInt(db.Rows(j)("IDDonVi_Moi")) And CInt(db.Rows(i)("IdPhong_Moi")) = CInt(db.Rows(j)("IdPhong_Moi")) And CInt(db.Rows(i)("IDChucVu_Moi")) = CInt(db.Rows(j)("IDChucVu_Moi"))) Then
                                                    If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1515" Then 'Thử việc
                                                        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                        Exit While
                                                    Else
                                                        j = j + 1
                                                    End If
                                                Else
                                                    'Nếu QĐ trước đó la QD Kiêm nhiệm mà QĐ tiếp theo ko phải là QĐ Thôi kiêm nhiệm thì bỏ qua
                                                    If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                        Exit While
                                                    Else
                                                        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                        i = j - 1
                                                        Exit While
                                                    End If
                                                End If
                                        End Select
                                    End While
                                    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                    If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1534" Then
                                        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("LoaiQD").ToString() <> "", db.Rows(i)("LoaiQD").ToString(), "")
                                    Else
                                        If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                            dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = db.Rows(i)("LoaiQD").ToString() & " " & IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString().Replace("(CN cấp I & tương đương)", ""), "")
                                        Else
                                            If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1515" Then 'Thử việc
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", "") & " (" & db.Rows(i)("LoaiQD").ToString() & ")", "")
                                            Else
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", ""), "")
                                            End If
                                        End If
                                    End If
                                    row = row + 1
                                    If j = db.Rows.Count Then Exit For
                                End If

                                'dgv_congtac.Rows.Add()
                                'dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                'ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                'Dim j As Integer = i + 1
                                'While j < db.Rows.Count
                                '    If (CInt(db.Rows(i)("IDDonVi_Moi")) = CInt(db.Rows(j)("IDDonVi_Moi")) And CInt(db.Rows(i)("IdPhong_Moi")) = CInt(db.Rows(j)("IdPhong_Moi")) And CInt(db.Rows(i)("IDChucVu_Moi")) = CInt(db.Rows(j)("IDChucVu_Moi"))) Then
                                '        j = j + 1
                                '    Else
                                '        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                '        i = j - 1
                                '        Exit While
                                '    End If
                                'End While
                                'dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                'dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                'row = row + 1
                                'If j = db.Rows.Count Then Exit For
                            End If
                        Next
                        'For i As Integer = 0 To db.Rows.Count - 1
                        '    Dim ThoiGian As String = ""
                        '    If i = 0 Then
                        '        dgv_congtac.Rows.Add()
                        '        dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                        '        ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        '        Dim j As Integer = i + 1
                        '        While j < db.Rows.Count
                        '            Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                        '                Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516"
                        '                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                        '                    i = j - 1
                        '                    Exit While
                        '                Case Else
                        '                    j = j + 1
                        '            End Select
                        '        End While
                        '        dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                        '        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        '        row = row + 1
                        '    Else
                        '        If Not (db.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
                        '                Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
                        '                Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
                        '            dgv_congtac.Rows.Add()
                        '            dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                        '            ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        '            If Not (db.Rows(i)("DenNgay") Is DBNull.Value) Then
                        '                If Not (CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                        '                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                        '                    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                        '                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        '                Else
                        '                    If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                        '                        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                        '                    Else
                        '                        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        '                    End If
                        '                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        '                    Dim j As Integer = i + 1
                        '                    While j < db.Rows.Count
                        '                        Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                        '                            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516"
                        '                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                        '                                i = j - 1
                        '                                Exit While
                        '                            Case Else
                        '                                j = j + 1
                        '                        End Select
                        '                    End While
                        '                    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                        '                End If
                        '            Else
                        '                If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                        '                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                        '                Else
                        '                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        '                End If
                        '                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        '                Dim j As Integer = i + 1
                        '                While j < db.Rows.Count
                        '                    Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                        '                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516"
                        '                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                        '                            i = j - 1
                        '                            Exit While
                        '                        Case Else
                        '                            j = j + 1
                        '                    End Select
                        '                End While
                        '                dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                        '            End If
                        '            row = row + 1
                        '        End If
                        '    End If
                        'Next
                    End If
                End Using

                'Using db As DataTable = _HS_CanBo.GetAll_Document(_IdCanBo, 2)
                '    If Not (db Is Nothing) Then
                '        If (db.Rows.Count > 0) Then
                '            For i As Integer = 0 To db.Rows.Count - 1
                '                dgv_congtac.Rows.Add()
                '                dgv_congtac.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdHSCongTac").ToString() <> "", db.Rows(i)("IdHSCongTac").ToString(), "")
                '                dgv_congtac.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                '                dgv_congtac.Rows(i).Cells("cln_Tungay").Value = IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                '                If (db.Rows(i)("DenNgay").ToString().Trim() <> "") Then
                '                    dgv_congtac.Rows(i).Cells("cln_Denngay").Value = Format(CType(db.Rows(i)("DenNgay"), DateTime), "dd-MM-yyyy")
                '                Else
                '                    dgv_congtac.Rows(i).Cells("cln_Denngay").Value = ""
                '                End If
                '                dgv_congtac.Rows(i).Cells("cln_Chucvu").Value = IIf(db.Rows(i)("ChucVu").ToString() <> "", db.Rows(i)("ChucVu").ToString(), "")
                '                dgv_congtac.Rows(i).Cells("cln_Diachi").Value = IIf(db.Rows(i)("DiaChi").ToString() <> "", db.Rows(i)("DiaChi").ToString(), "")
                '                dgv_congtac.Rows(i).Cells("cln_Lydo").Value = IIf(db.Rows(i)("LyDo").ToString() <> "", db.Rows(i)("LyDo").ToString(), "")
                '                dgv_congtac.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                '            Next
                '        End If
                '    End If
                'End Using
                'dgv_congtac_CellClick(Nothing, Nothing)

            Case 3  'Fill - Hồ sơ Xuất ngoại
                dgv_xuatngoai.Rows.Clear()
                Using db As DataTable = _HS_CanBo.GetAll_Document(_IdCanBo, 3)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_xuatngoai.Rows.Add()
                                dgv_xuatngoai.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdXuatNgoai").ToString() <> "", db.Rows(i)("IdXuatNgoai").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                                dgv_xuatngoai.Rows(i).Cells("cln_Tungay").Value = IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Denngay").Value = IIf(db.Rows(i)("DenNgay").ToString().Trim() <> "", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Quocgia").Value = IIf(db.Rows(i)("Nuoc_Den").ToString() <> "", db.Rows(i)("Nuoc_Den").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Mucdich").Value = IIf(db.Rows(i)("MucDich").ToString() <> "", db.Rows(i)("MucDich").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Soqd").Value = IIf(db.Rows(i)("SoQD").ToString() <> "", db.Rows(i)("SoQD").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Ngayky").Value = IIf(db.Rows(i)("NgayKy_QD").ToString().Trim() <> "", CType(db.Rows(i)("NgayKy_QD").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Nguoiky").Value = IIf(db.Rows(i)("NguoiKy_QD").ToString() <> "", db.Rows(i)("NguoiKy_QD").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Chucvu").Value = IIf(db.Rows(i)("Chuc_Vu").ToString() <> "", db.Rows(i)("Chuc_Vu").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

            Case 4  'Fill - Hồ sơ Hộ chiếu
                dgv_hochieu.Rows.Clear()
                Using db As DataTable = _HS_CanBo.GetAll_Document(_IdCanBo, 4)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_hochieu.Rows.Add()
                                dgv_hochieu.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdHoChieu").ToString() <> "", db.Rows(i)("IdHoChieu").ToString(), "")
                                dgv_hochieu.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                                dgv_hochieu.Rows(i).Cells("cln_Sohc").Value = IIf(db.Rows(i)("So_HoChieu").ToString() <> "", db.Rows(i)("So_HoChieu").ToString(), "")
                                If db.Rows(i)("Loai_HC").ToString() = "1" Then
                                    dgv_hochieu.Rows(i).Cells("cln_Loaihc").Value = "Phổ thông"
                                ElseIf db.Rows(i)("Loai_HC").ToString() = "2" Then
                                    dgv_hochieu.Rows(i).Cells("cln_Loaihc").Value = "Công vụ"
                                ElseIf db.Rows(i)("Loai_HC").ToString() = "3" Then
                                    dgv_hochieu.Rows(i).Cells("cln_Loaihc").Value = "Ngoại giao"
                                Else : dgv_hochieu.Rows(i).Cells("cln_Loaihc").Value = ""
                                End If

                                dgv_hochieu.Rows(i).Cells("cln_Ngaycap").Value = IIf(db.Rows(i)("NgayCap").ToString().Trim() <> "", CType(db.Rows(i)("NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_hochieu.Rows(i).Cells("cln_Noicap").Value = IIf(db.Rows(i)("NoiCap").ToString() <> "", db.Rows(i)("NoiCap").ToString(), "")
                                dgv_hochieu.Rows(i).Cells("cln_Ngayhh").Value = IIf(db.Rows(i)("NgayHH").ToString().Trim() <> "", CType(db.Rows(i)("NgayHH").ToString(), DateTime).ToString("dd-MM-yyyy"), "")

                                If db.Rows(i)("TinhTrang").ToString() = "1" Then
                                    dgv_hochieu.Rows(i).Cells("cln_Tinhtrang").Value = "Còn hiệu lực"
                                ElseIf db.Rows(i)("TinhTrang").ToString() = "2" Then
                                    dgv_hochieu.Rows(i).Cells("cln_Tinhtrang").Value = "Hết hạn"
                                ElseIf db.Rows(i)("TinhTrang").ToString() = "3" Then
                                    dgv_hochieu.Rows(i).Cells("cln_Tinhtrang").Value = "Mất"
                                Else : dgv_hochieu.Rows(i).Cells("cln_Tinhtrang").Value = ""
                                End If
                                dgv_hochieu.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

            Case 5  'Fill - Hồ sơ Tham gia lực lượng vũ trang
                dgv_llvt.Rows.Clear()
                Using db As DataTable = _HS_CanBo.GetAll_Document(_IdCanBo, 5)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_llvt.Rows.Add()
                                dgv_llvt.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdHSLLVT").ToString() <> "", db.Rows(i)("IdHSLLVT").ToString(), "")
                                dgv_llvt.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                                dgv_llvt.Rows(i).Cells("cln_Tungay").Value = IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_llvt.Rows(i).Cells("cln_Denngay").Value = IIf(db.Rows(i)("DenNgay").ToString().Trim() <> "", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_llvt.Rows(i).Cells("cln_Phanloai").Value = IIf(db.Rows(i)("Phan_Loai").ToString() <> "", db.Rows(i)("Phan_Loai").ToString(), "")
                                'Fill loại quân hàm
                                Dim strName As String = String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 2 And Status = 1", CType(db.Rows(i)("IdQuanHam").ToString(), Int32))
                                dgv_llvt.Rows(i).Cells("cln_Quanham").Value = _HS_CanBo.GetNameByCode(strName)
                                dgv_llvt.Rows(i).Cells("cln_Chucvu").Value = IIf(db.Rows(i)("ChucVu").ToString() <> "", db.Rows(i)("ChucVu").ToString(), "")
                                dgv_llvt.Rows(i).Cells("cln_Donvi").Value = IIf(db.Rows(i)("DonVi").ToString() <> "", db.Rows(i)("DonVi").ToString(), "")
                                dgv_llvt.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

            Case 6  'Fill - Hồ sơ cũ (Thông tin của cán bộ khi chưa thoát ly)
                dgv_hosocu.Rows.Clear()
                Using db As DataTable = _HS_CanBo.GetAll_Document(_IdCanBo, 6)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Int32 = 0 To db.Rows.Count - 1
                                dgv_hosocu.Rows.Add()
                                dgv_hosocu.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdHSCu").ToString() <> "", db.Rows(i)("IdHSCu").ToString(), "")
                                dgv_hosocu.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                dgv_hosocu.Rows(i).Cells("cln_Tuthang").Value = IIf(db.Rows(i)("TuThang").ToString().Trim() <> "", CType(db.Rows(i)("TuThang").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_hosocu.Rows(i).Cells("cln_Denthang").Value = IIf(db.Rows(i)("DenThang").ToString().Trim() <> "", CType(db.Rows(i)("DenThang").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_hosocu.Rows(i).Cells("cln_Diachi").Value = IIf(db.Rows(i)("DiaChi").ToString() <> "", db.Rows(i)("DiaChi").ToString(), "")
                                dgv_hosocu.Rows(i).Cells("cln_Nghenghiep").Value = IIf(db.Rows(i)("NgheNghiep").ToString() <> "", db.Rows(i)("NgheNghiep").ToString(), "")
                                dgv_hosocu.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

            Case 7  'Fill - Hồ sơ học hàm học vị của cán bộ
                'Thực hiện fill dữ liệu ra - CheckListBox Học hàm của cán bộ nếu có
                Dim strSQL As String = ""
                strSQL = String.Format("Select a.IdCB_HocHam,b.ten_goi as Hoc_Ham From CB_HocHam a, DanhMuc b Where (a.IdHocHam = b.id and b.id_goc = 26 and b.Status = 1) and IdCanBo = '{0}' order by b.ma_so", _IdCanBo)
                arr_Hhhv_LHocham.Clear()
                clb_hhhv_hocham.Items.Clear()
                arr_Hhhv_LHocham = _Globals.FillData_CheckedListBox(clb_hhhv_hocham, strSQL)
                _Globals.ResetItems_CheckedListBox(clb_hhhv_hocham)
                clb_hhhv_hocham_SelectedIndexChanged(Nothing, Nothing)
            Case 8  'Fill - Hồ sơ học hàm học vị của cán bộ
                Dim strSQL As String = ""
                strSQL = String.Format("Select a.IdCB_HocVi,b.ten_goi as Hoc_Vi From CB_HocVi a, DanhMuc b Where (a.IdHocVi = b.id and b.id_goc = 20 and b.Status = 1) and IdCanBo = '{0}' order by b.ma_so", _IdCanBo)
                arr_Hhhv_LHocvi.Clear()
                clb_hhhv_hocvi.Items.Clear()
                arr_Hhhv_LHocvi = _Globals.FillData_CheckedListBox(clb_hhhv_hocvi, strSQL)
                _Globals.ResetItems_CheckedListBox(clb_hhhv_hocvi)
                clb_hhhv_hocvi_SelectedIndexChanged(Nothing, Nothing)
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra dữ liệu Hợp lệ khi cập nhật
    ''' </summary>
    ''' <param name="state">Chỉ số xác định hồ sơ cần kiểm tra
    '''                           0: Hồ sơ Cán bộ
    '''                           1: Hồ sơ Công tác
    '''                           2: Hồ sơ Xuất ngoại
    '''                           3: Hồ sơ Hộ chiếu
    '''                           4: Hồ sơ Tham gia Lực lượng vũ trang
    '''                           5: Hồ sơ Cũ (Thông tin cán bộ khi chưa thoát ly)
    '''                           6: Hồ sơ Học hàm học vị
    ''' </param>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function Valid(ByVal state As Byte) As Boolean
        Dim strSQL As String = ""
        Select Case state
            Case 0      'Kiểm tra hồ sơ Cán bộ
                Dim dbconn As DBAccess = New DBAccess
                If Not (_IdCanBo <> "" And dbconn.getString("SELECT IdOld From HS_Canbo WHERE IdCanBo='" & _IdCanBo & "'") <> "") Then
                    If (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn đơn vị công tác của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_donvi
                        Return False
                    End If

                    If (edt_macb.Text.Trim() = "") Then
                        MessageBox.Show("Mã cán bộ không được để trống. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_macb
                        Return False
                    End If
                    'Băt điều kiện - Trùng mã cán bộ
                    If (edt_macb.Text.Trim() <> "") Then
                        'If (_IdCanBo <> "") Then
                        '    strSQL = String.Format("Select * From HS_CanBo Where MaCB = '{0}' and IdCanBo <> '{1}'", Globals.Find_Replace(edt_macb.Text.Trim().ToString()), _IdCanBo)
                        'Else
                        '    strSQL = String.Format("Select * From HS_CanBo Where MaCB = '{0}'", Globals.Find_Replace(edt_macb.Text.Trim().ToString()))
                        'End If
                        'Using db = _SqlHelper.SelectDBRows(strSQL)
                        '    If Not (db Is Nothing) Then
                        '        If (db.Rows.Count > 0) Then
                        '            MessageBox.Show(String.Format("Mã '{0}' cán bộ đã tồn tại. Vui lòng kiểm tra lại!", edt_macb.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        '            ActiveControl = edt_macb
                        '            Return False
                        '        End If
                        '    End If
                        'End Using
                        edt_macb.Text = formatLenString(5, edt_macb.Text.Trim)
                        If checkMaCanBo(arr_Donvi(cb_donvi.SelectedIndex), _IdCanBo, edt_macb.Text) = 0 Then
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
                            erpd_main.SetError(edt_macb, String.Format("Mã cán bộ '{0}' đã tồn tại, hãy nhập mã cán bộ khác thuộc một số trong dẫy số '{1}' theo đơn vị '{2}'. Vui lòng kiểm tra lại!", edt_macb.Text, sCodeByBranch, BrandNameByUserLogin))
                            ActiveControl = edt_macb
                            Return False
                        End If
                    End If
                End If

                If (edt_hoten.Text.Trim() = "") Then
                    MessageBox.Show("Họ tên cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_hoten
                    Return False
                End If
                If (rb_nam.Checked = False And rb_nu.Checked = False) Then
                    MessageBox.Show("Bạn chưa lựa chọn giới tính của cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    rb_nam.Focus()
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
                        strSQL = String.Format("Select * From HS_CanBo Where CMT_So = '{0}' and IdCanBo <> '{1}'", Globals.Find_Replace(edt_cmt_so.Text.Trim().ToString()), _IdCanBo)
                    Else
                        strSQL = String.Format("Select * From HS_CanBo Where CMT_So = '{0}'", Globals.Find_Replace(edt_cmt_so.Text.Trim().ToString()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Số chứng minh thư: '{0}' của cán bộ đã tồn tại. Vui lòng kiểm tra lại!", edt_cmt_so.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_cmt_so
                                Return False
                            End If
                        End If
                    End Using
                End If
                'Thực hiện kiểm tra tuổi làm chứng minh thư có phải là 16 tuổi không ?
                Dim ngay_dk As DateTime = dtpk_ngaysinh.Value.AddYears(14)
                If (dtpk_cmt_ngaycap.Value <= ngay_dk) Then
                    MessageBox.Show("Ngày cấp chứng minh thư không hợp lệ. Vui lòng kiểm tra lại!" + vbCrLf + "Lưu ý: Theo quy định tại Điều 3, Nghị Định Số: 05/1999/NĐ-CP thì Công dân Việt Nam từ đủ 14 tuổi trở lên mới được làm chứng minh thư nhân dân", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_cmt_ngaycap
                    Return False
                End If

                'Kiểm tra thông tin: Nới sinh - Nguyên quán - Thường trú - Tạm trú
                If (cb_ns_tinh.SelectedIndex <= 0 And cb_ns_tinh.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn nơi sinh tỉnh (thành phố) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_ns_tinh
                    Return False
                End If
                If (cb_ns_xa.SelectedIndex <= 0 And cb_ns_xa.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn nơi sinh xã (phường) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_ns_xa
                    Return False
                End If
                'If (cb_ns_thon.SelectedIndex <= 0 And cb_ns_thon.Items.Count <> 0) Then
                '    MessageBox.Show("Bạn chưa chọn nơi sinh thôn của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = cb_ns_xa
                '    Return False
                'End If
                'If (edt_ns_diachi.Text.Trim() = "") Then
                '    MessageBox.Show("Địa chỉ nơi sinh không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = edt_ns_diachi
                '    Return False
                'End If

                If (cb_nq_tinh.SelectedIndex <= 0 And cb_nq_tinh.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn nguyên quán tỉnh (thành phố) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_nq_tinh
                    Return False
                End If
                If (cb_nq_xa.SelectedIndex <= 0 And cb_nq_xa.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn nguyên quán xã (phường) của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_nq_xa
                    Return False
                End If
                'If (cb_nq_thon.SelectedIndex <= 0 And cb_nq_thon.Items.Count <> 0) Then
                '    MessageBox.Show("Bạn chưa chọn nguyên quán thôn của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = cb_nq_thon
                '    Return False
                'End If
                'If (edt_nq_diachi.Text.Trim() = "") Then
                '    MessageBox.Show("Địa chỉ nguyên quán không được để trống. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = edt_nq_diachi
                '    Return False
                'End If

                If (cb_tt_tinh.SelectedIndex <= 0 And cb_tt_tinh.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn địa chỉ tỉnh/thành phố nơi đăng ký hộ khẩu của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_tt_tinh
                    Return False
                End If
                If (cb_tt_xa.SelectedIndex <= 0 And cb_tt_xa.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn địa chỉ xã/phường nơi đăng ký hộ khẩu của cán bộ. Vui lòng kiểm tra lại!!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_tt_xa
                    Return False
                End If
                'If (cb_tt_thon.SelectedIndex <= 0 And cb_tt_thon.Items.Count <> 0) Then
                '    MessageBox.Show("Bạn chưa chọn địa chỉ thôn nơi đăng ký hộ khẩu của cán bộ. Vui lòng kiểm tra lại!!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = cb_tt_thon
                '    Return False
                'End If
                'If (edt_tt_diachi.Text.Trim() = "") Then
                '    MessageBox.Show("Địa chỉ chi tiết nơi đăng ký hộ khẩu không được để trống. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = edt_tt_diachi
                '    Return False
                'End If


                If (cb_ttr_tinh.SelectedIndex <= 0 And cb_ttr_tinh.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn địa chỉ tỉnh/thành phố thường trú/tạm trú của cán bộ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_ttr_tinh
                    Return False
                End If
                If (cb_ttr_xa.SelectedIndex <= 0 And cb_ttr_xa.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn địa chỉ xã/phường thường trú/tạm trú của cán bộ. Vui lòng kiểm tra lại!!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_ttr_xa
                    Return False
                End If
                If (cb_ttr_thon.SelectedIndex <= 0 And cb_ttr_thon.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn địa chỉ thôn thường trú/tạm trú của cán bộ. Vui lòng kiểm tra lại!!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    tctrl_hs_cb.SelectedIndex = 0
                    ActiveControl = cb_ttr_thon
                    Return False
                End If
                'If (edt_ttr_diachi.Text.Trim() = "") Then
                '    MessageBox.Show("Địa chỉ chi tiết thường trú/tạm trú của cán bộ không được để trống. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    tctrl_hs_cb.SelectedIndex = 0
                '    ActiveControl = edt_ttr_diachi
                '    Return False
                'End If

                If (edt_email.Text.Trim() <> "") Then ' Nếu địa chỉ e-mail của cán bộ mà nhập thì kiểm tra có hợp lệ không
                    If (Not Globals.IsEmail(edt_email.Text.Trim().ToString())) Then
                        MessageBox.Show("Địa chỉ e-mail của cán bộ không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 0
                        ActiveControl = edt_email
                        Return False
                    End If
                End If
                If cb_tpgd.SelectedIndex <= 0 Then cb_tpgd.SelectedIndex = 0
                If cb_ut_giadinh.SelectedIndex <= 0 Then cb_ut_giadinh.SelectedIndex = 0
                If cb_tdvh.SelectedIndex <= 0 Then cb_tdvh.SelectedIndex = 0
                If cb_tdct.SelectedIndex <= 0 Then cb_tdct.SelectedIndex = 0
                If FlagEvent = 1 Then
                    'Bắt dữ liệu trống phần quyết định nhân sự
                    If (edt_qd_soqd.Text.Trim() = "") Then
                        MessageBox.Show("Số quyết định nhân sự của cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = edt_qd_soqd
                        Return False
                    End If
                    If (cb_qd_loaiqd.SelectedIndex <= 0 And cb_qd_loaiqd.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn loại quyết định nhân sự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = cb_qd_loaiqd
                        Return False
                    End If
                    If (edt_qd_nguoiky.Text.Trim() = "") Then
                        MessageBox.Show("Người ký quyết định nhân sự không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = edt_qd_nguoiky
                        Return False
                    End If

                    If (cb_qd_chucvu.SelectedIndex <= 0 And cb_qd_chucvu.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn chức vụ người ký quyết định!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = cb_qd_chucvu
                        Return False
                    End If
                    If (cb_qd_donvi_moi.SelectedIndex <= 0 And cb_qd_donvi_moi.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn đơn vị mới của quyết định nhân sự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = cb_qd_donvi_moi
                        Return False
                    End If
                    If (cb_qd_phongban_moi.SelectedIndex <= 0 And cb_qd_phongban_moi.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn phòng ban mới của quyết định nhân sự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = cb_qd_phongban_moi
                        Return False
                    End If
                    If (cb_qd_chucvu_moi.SelectedIndex <= 0 And cb_qd_chucvu_moi.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn chức vụ mới của quyết định nhân sự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = cb_qd_chucvu_moi
                        Return False
                    End If
                    If (cb_qd_cm_moi.SelectedIndex <= 0 And cb_qd_cm_moi.Items.Count <> 0) Then
                        MessageBox.Show("Bạn chưa chọn chuyên môn mới của quyết định nhân sự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        tctrl_hs_cb.SelectedIndex = 3
                        ActiveControl = cb_qd_cm_moi
                        Return False
                    End If
                End If

                If (strPath <> "") Then
                    If Not System.IO.File.Exists(strPath) Then
                        MessageBox.Show("Đường dẫn chứa ảnh hồ sơ nhân sự mà bạn chọn không tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = btn_chonanh
                        Return False
                    End If
                End If

            Case 1      'Kiểm tra hồ sơ Công tác
                If dtpk_ct_denngay.Checked = True Then
                    If dtpk_ct_tungay.Value >= dtpk_ct_denngay.Value Then
                        MessageBox.Show("Ngày bắt đầu quá trình công tác không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_ct_denngay
                        Return False
                    End If
                End If
                If (edt_ct_chucvu.Text.Trim() = "") Then
                    MessageBox.Show("Chức vụ công tác không được trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_ct_chucvu
                    Return False
                End If
                If (edt_ct_diachi.Text.Trim() = "") Then
                    MessageBox.Show("Địa chỉ nơi công tác không được trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_ct_diachi
                    Return False
                End If

                'Bắt điều kiện Không cho --> Các khoảng thời gian đan xen lẫn nhau
                If (_IdEdit = "") Then
                    strSQL = String.Format("Select * From HS_CongTac Where IdCanBo = '{0}' Order by TuNgay Asc", _IdCanBo)
                ElseIf (_IdEdit <> "") Then
                    strSQL = String.Format("Select * From HS_CongTac Where IdCanBo = '{0}' And IdHSCongTac <> '{1}' Order by TuNgay Asc", _IdCanBo, _IdEdit)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = DateTime.Now
                                If (db.Rows(i)("DenNgay").ToString() <> "") Then
                                    If CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900" Then
                                        _DenNgay = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                    End If
                                End If
                                If (dtpk_ct_denngay.Checked = True) Then
                                    If ((_TuNgay <= dtpk_ct_tungay.Value And dtpk_ct_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_ct_denngay.Value And dtpk_ct_denngay.Value <= _DenNgay) Or (dtpk_ct_tungay.Value <= _TuNgay And dtpk_ct_denngay.Value >= _DenNgay)) Then
                                        Exist = True
                                        Exit For
                                    End If
                                Else
                                    If ((_TuNgay <= dtpk_ct_tungay.Value And dtpk_ct_tungay.Value <= _DenNgay) Or (dtpk_ct_tungay.Value <= _TuNgay)) Then
                                        Exist = True
                                        Exit For
                                    End If
                                End If
                                If (db.Rows(i)("DenNgay").ToString() = "") Then
                                    If (dtpk_ct_tungay.Value > DateTime.Now) Then
                                        Exist = True
                                    End If
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ công tác không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian trong hồ sơ công tác không thể trùng hoặc đan xen lẫn nhau.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_ct_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using

            Case 2      'Kiểm tra hồ sơ Xuất ngoại
                If dtpk_xn_tungay.Value >= dtpk_xn_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu hồ sơ xuất ngoại không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_xn_denngay
                    Return False
                End If
                If (cb_xn_nuocden.SelectedIndex <= 0 And cb_xn_nuocden.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn quốc gia đến trong hồ sơ xuất ngoại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_xn_nuocden
                    Return False
                End If
                If (cb_xn_nuocden.SelectedIndex > 0) Then
                    If (cb_xn_nuocden.SelectedItem.ToString().Trim() = "Việt Nam") Then
                        MessageBox.Show("Quốc gia xuất ngoại " + cb_xn_nuocden.SelectedItem.ToString().Trim() + " mà bạn chọn không hợp lý!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_xn_nuocden
                        Return False
                    End If
                End If
                'If (edt_xn_soqd.Text.Trim() = "") Then
                '    MessageBox.Show("Số quyết định xuất ngoại không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_xn_soqd
                '    Return False
                'End If
                If dtpk_xn_tungay.Value < dtpk_xn_ngayky.Value Then
                    MessageBox.Show("Ngày ký quyết định không hợp lệ so với khoảng thời gian xuất ngoại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_xn_ngayky
                    Return False
                End If

                If (edt_xn_nguoiky.Text.Trim() = "") Then
                    MessageBox.Show("Người ký quyết định xuất ngoại không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_xn_nguoiky
                    Return False
                End If
                If (cb_xn_chucvu.SelectedIndex <= 0 And cb_xn_chucvu.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn chức vụ người ký quyết định!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_xn_chucvu
                    Return False
                End If
                'Bắt điều kiện Không cho --> các khoảng thời gian đan xen lẫn nhau
                If (_IdEdit = "") Then
                    strSQL = String.Format("Select * from HS_XuatNgoai Where IdCanBo = '{0}' Order by TuNgay Asc", _IdCanBo)
                ElseIf (_IdEdit <> "") Then
                    strSQL = String.Format("Select * From HS_XuatNgoai Where IdCanBo = '{0}' And IdXuatNgoai <> '{1}' Order by TuNgay Asc", _IdCanBo, _IdEdit)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_xn_tungay.Value And dtpk_xn_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_xn_denngay.Value And dtpk_xn_denngay.Value <= _DenNgay) Or (dtpk_xn_tungay.Value <= _TuNgay And dtpk_xn_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ xuất ngoại không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_xn_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using

            Case 3      'Kiểm tra hồ sơ Hộ chiếu
                If (edt_hc_sohc.Text.Trim() = "") Then
                    MessageBox.Show("Số hộ chiếu không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_hc_sohc
                    Return False
                End If
                If (cb_hc_loaihc.SelectedIndex <= 0) Then
                    MessageBox.Show("Bạn chưa chọn loại hộ chiếu!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_hc_loaihc
                    Return False
                End If
                If (edt_hc_noicap.Text.Trim() = "") Then
                    MessageBox.Show("Nơi cấp hộ chiếu không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_hc_noicap
                    Return False
                End If
                If (cb_hc_tinhtrang.SelectedIndex <= 0) Then
                    MessageBox.Show("Bạn chưa tình trạng hộ chiếu!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_hc_tinhtrang
                    Return False
                End If
                If dtpk_hc_ngaycap.Value >= dtpk_hc_ngayhh.Value Then
                    MessageBox.Show("Ngày cấp hộ chiếu không thể lớn hơn hoặc bằng ngày hết hạn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_hc_ngayhh
                    Return False
                End If

            Case 4      'Kiểm tra hồ sơ Tham gia Lực lượng vũ trang
                If dtpk_llvt_tungay.Value >= dtpk_llvt_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_llvt_denngay
                    Return False
                End If
                If (cb_llvt_phanloai.SelectedIndex <= 0 And cb_llvt_phanloai.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn phân loại lực lượng vũ trang!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_llvt_phanloai
                    Return False
                End If
                If (cb_llvt_quanham.SelectedIndex <= 0 And cb_llvt_quanham.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn quân hàm trong hồ sơ lực lượng vũ trang!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_llvt_quanham
                    Return False
                End If
                If (edt_llvt_chucvu.Text.Trim() = "") Then
                    MessageBox.Show("Chức vụ trong hồ sơ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_llvt_chucvu
                    Return False
                End If
                If (edt_llvt_donvi.Text.Trim() = "") Then
                    MessageBox.Show("Đơn vị tham gia không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_llvt_donvi
                    Return False
                End If

                'Bắt điều kiện Không cho --> các khoảng thời gian đan xen lẫn nhau
                If (_IdEdit = "") Then
                    strSQL = String.Format("Select * from HS_LLVT Where IdCanBo = '{0}' Order by TuNgay ASC", _IdCanBo)
                ElseIf (_IdEdit <> "") Then
                    strSQL = String.Format("Select * From HS_LLVT Where IdCanBo = '{0}' And IdHSLLVT <> '{1}' Order by TuNgay ASC", _IdCanBo, _IdEdit)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_llvt_tungay.Value And dtpk_llvt_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_llvt_denngay.Value And dtpk_llvt_denngay.Value <= _DenNgay) Or (dtpk_llvt_tungay.Value <= _TuNgay And dtpk_llvt_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ tham gia lực lượng vũ trang không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_llvt_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using

            Case 5      'Kiểm tra hồ sơ Cũ (Thông tin cán bộ khi chưa thoát ly)
                If dtpk_hscu_tungay.Value >= dtpk_hscu_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_llvt_denngay
                    Return False
                End If

                'Bắt điều kiện Không cho --> các khoảng thời gian đan xen lẫn nhau
                If (_IdEdit = "") Then
                    strSQL = String.Format("Select * from HS_Cu Where IdCanBo = '{0}' Order by TuThang Asc", _IdCanBo)
                ElseIf (_IdEdit <> "") Then
                    strSQL = String.Format("Select * From HS_Cu Where IdCanBo = '{0}' And IdHSCu <> '{1}' Order by TuThang Asc", _IdCanBo, _IdEdit)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuThang").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenThang").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_hscu_tungay.Value And dtpk_hscu_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_hscu_denngay.Value And dtpk_hscu_denngay.Value <= _DenNgay) Or (dtpk_hscu_tungay.Value <= _TuNgay And dtpk_hscu_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ cũ không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_hscu_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using

            Case 6      'Kiểm tra hồ sơ Học hàm học vị
                If (cb_hhhv_hocham.SelectedIndex <= 0) And (cb_hhhv_hocvi.SelectedIndex <= 0) Then
                    MessageBox.Show("Bạn muốn thêm mới dữ liệu cho học hàm hay học vị!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_hhhv_hocham
                    Return False
                End If
                'Kiểm tra điều kiện trùng - Học hàm học vị
                If (ckb_hocham.Checked = True) Then
                    If (cb_hhhv_hocham.SelectedIndex <= 0) Then
                        MessageBox.Show("Bạn chưa chọn học hàm cần cập nhật!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_hhhv_hocham
                    End If
                    If (cb_hhhv_hocham.SelectedIndex > 0) Then
                        If (flagAdd = True And _Code_HocHam = "") Then
                            strSQL = String.Format("Select * From CB_HocHam Where IdCanBo = '{0}' And IdHocHam = {1}", _IdCanBo, CType(IIf(arr_Hhhv_CHocham.Count > 0, arr_Hhhv_CHocham(cb_hhhv_hocham.SelectedIndex), "0"), Int32))
                        Else
                            If (_Code_HocHam <> "") Then
                                strSQL = String.Format("Select * From CB_HocHam Where IdCanBo = '{0}' And IdHocHam = {1} And IdCB_HocHam <> '{2}'", _IdCanBo, CType(IIf(arr_Hhhv_CHocham.Count > 0, arr_Hhhv_CHocham(cb_hhhv_hocham.SelectedIndex), "0"), Int32), _Code_HocHam)
                            End If
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        MessageBox.Show("Học hàm của cán bộ này đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                        ActiveControl = cb_hhhv_hocham
                                        Return False
                                    End If
                                End If
                            End Using
                        End If
                    End If
                End If

                If (ckb_hocvi.Checked = True) Then
                    If (cb_hhhv_hocvi.SelectedIndex <= 0) Then
                        MessageBox.Show("Bạn chưa chọn học vị cần cập nhật!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_hhhv_hocvi
                    End If
                    If (cb_hhhv_hocvi.SelectedIndex > 0) Then
                        If (flagAdd = True And _Code_HocVi = "") Then
                            strSQL = String.Format("Select * From CB_HocVi Where IdCanBo = '{0}' And IdHocVi = {1}", _IdCanBo, CType(IIf(arr_Hhhv_CHocvi.Count > 0, arr_Hhhv_CHocvi(cb_hhhv_hocvi.SelectedIndex), "0"), Int32))
                        Else
                            If (_Code_HocVi <> "") Then
                                strSQL = String.Format("Select * From CB_HocVi Where IdCanBo = '{0}' And IdHocVi = {1} And IdCB_HocVi <> '2'", _IdCanBo, CType(IIf(arr_Hhhv_CHocvi.Count > 0, arr_Hhhv_CHocvi(cb_hhhv_hocvi.SelectedIndex), "0"), Int32), _Code_HocVi)
                            End If
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show("Học vị của cán bộ này đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = cb_hhhv_hocvi
                                    Return False
                                End If
                            End If
                        End Using
                    End If
                End If
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Fill dữ liệu hồ sơ cán bộ tập sự (Nếu khi thêm mới chuyển một cán bộ tập sự sang hồ sơ cán bộ)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Fill_Data()
        Dim dr As DataRow
        dr = _HS_CanBo.GetTrainee(0, _IdCB_TS, "", "", CType("1", Byte), CType("1", Byte))
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
                cb_donvi.SelectedIndex = IIf(dr("IdChiNhanh").ToString() <> "", CType(arr_Donvi.IndexOf(dr("IdChiNhanh").ToString()), Integer), 0)
                cb_quoctich.SelectedIndex = IIf(dr("IdQuocTich").ToString() <> "", CType(arr_Quoctich.IndexOf(dr("IdQuocTich").ToString()), Integer), 0)
                cb_donvi_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdPhongBan").ToString().Trim() <> "") Then
                    If (dr("IdPhongBan").ToString().Trim().Substring(0, 3) = "DV_") Then
                        cb_qd_donvi_moi.SelectedIndex = IIf(dr("IdPhongBan").ToString() <> "", CType(arrQd_Donvi_moi.IndexOf(dr("IdPhongBan").ToString().Trim().Substring(3)), Integer), 0)
                        cb_qd_donvi_moi_SelectedIndexChanged(Nothing, Nothing)
                    ElseIf (dr("IdPhongBan").ToString().Trim().Substring(0, 3) = "PB_") Then
                        cb_qd_donvi_moi.SelectedIndex = IIf(dr("IdChiNhanh").ToString() <> "", CType(arrQd_Donvi_moi.IndexOf(dr("IdChiNhanh").ToString()), Integer), 0)
                        cb_qd_donvi_moi_SelectedIndexChanged(Nothing, Nothing)
                        cb_qd_phongban_moi.SelectedIndex = IIf(dr("IdPhongBan").ToString().Trim().Substring(3) <> "", CType(arrQd_Phongban_moi.IndexOf(dr("IdPhongBan").ToString().Trim().Substring(3)), Integer), 0)
                    End If
                End If
                cb_qd_chucvu_moi.SelectedIndex = IIf(cb_qd_chucvu_moi.FindString("Nhân viên") > 0, cb_qd_chucvu_moi.FindString("Nhân viên"), 0)
                cb_dantoc.SelectedIndex = IIf(dr("IdDanToc").ToString() <> "", CType(arr_Dantoc.IndexOf(dr("IdDanToc").ToString()), Integer), 0)
                cb_tongiao.SelectedIndex = IIf(dr("IdTonGiao").ToString() <> "", CType(arr_Tongiao.IndexOf(dr("IdTonGiao").ToString()), Integer), 0)
                edt_cmt_so.Text = dr("CMT_So").ToString().Trim()
                dtpk_cmt_ngaycap.Value = CType(dr("CMT_NgayCap").ToString(), DateTime)
                edt_cmt_noicap.Text = dr("CMT_NoiCap").ToString().Trim()

                'Ảnh thẻ
                My_LoadImage(picbx_main, dr("AnhThe"))

                cb_ns_tinh.SelectedIndex = IIf(dr("IdNS_Tinh").ToString() <> "", CType(arr_Ns_Tinh.IndexOf(dr("IdNS_Tinh").ToString()), Integer), 0)
                cb_ns_tinh_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdNS_Huyen").ToString() <> "") Then
                    cb_ns_huyen.SelectedIndex = CType(arr_Ns_Huyen.IndexOf(dr("IdNS_Huyen").ToString()), Integer)
                End If
                If (dr("IdNS_Xa").ToString() <> "") Then
                    cb_ns_xa.SelectedIndex = CType(arr_Ns_Xa.IndexOf(dr("IdNS_Xa").ToString()), Integer)
                End If
                cb_ns_xa_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdNS_Thon").ToString() <> "") Then
                    cb_ns_thon.SelectedIndex = CType(arr_Ns_Thon.IndexOf(dr("IdNS_Thon").ToString()), Integer)
                End If
                cb_ns_thon_SelectedIndexChanged(Nothing, Nothing)
                edt_ns_diachi.Text = dr("NS_DiaChi").ToString().Trim()

                cb_nq_tinh.SelectedIndex = IIf(dr("IdNQ_Tinh").ToString() <> "", CType(arr_Nq_Tinh.IndexOf(dr("IdNQ_Tinh").ToString()), Int32), 0)
                cb_nq_tinh_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdNQ_Huyen").ToString() <> "") Then
                    cb_nq_huyen.SelectedIndex = CType(arr_Nq_Huyen.IndexOf(dr("IdNQ_Huyen").ToString()), Int32)
                End If
                If (dr("IdNQ_Xa").ToString() <> "") Then
                    cb_nq_xa.SelectedIndex = CType(arr_Nq_Xa.IndexOf(dr("IdNQ_Xa").ToString()), Int32)
                End If
                cb_nq_xa_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdNQ_Thon").ToString() <> "") Then
                    cb_nq_thon.SelectedIndex = CType(arr_Nq_Thon.IndexOf(dr("IdNQ_Thon").ToString()), Int32)
                End If
                cb_nq_thon_SelectedIndexChanged(Nothing, Nothing)
                edt_nq_diachi.Text = dr("NQ_DiaChi").ToString().Trim()

                cb_tt_tinh.SelectedIndex = IIf(dr("IdThT_Tinh").ToString() <> "", CType(arr_TT_Tinh.IndexOf(dr("IdThT_Tinh").ToString()), Int32), 0)
                cb_tt_tinh_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdThT_Huyen").ToString() <> "") Then
                    cb_tt_huyen.SelectedIndex = CType(arr_TT_Huyen.IndexOf(dr("IdThT_Huyen").ToString()), Int32)
                End If
                If (dr("IdThT_Xa").ToString() <> "") Then
                    cb_tt_xa.SelectedIndex = CType(arr_TT_Xa.IndexOf(dr("IdThT_Xa").ToString()), Int32)
                End If
                cb_tt_xa_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdThT_Thon").ToString() <> "") Then
                    cb_tt_thon.SelectedIndex = CType(arr_TT_Thon.IndexOf(dr("IdThT_Thon").ToString()), Int32)
                End If
                cb_tt_thon_SelectedIndexChanged(Nothing, Nothing)
                edt_tt_diachi.Text = dr("ThT_DiaChi").ToString().Trim()
                edt_tt_dienthoai.Text = dr("ThT_Dienthoai").ToString().Trim()

                cb_ttr_tinh.SelectedIndex = IIf(dr("IdTTr_Tinh").ToString() <> "", CType(arr_TTr_Tinh.IndexOf(dr("IdTTr_Tinh").ToString()), Int32), 0)
                cb_ttr_tinh_SelectedIndexChanged(Nothing, Nothing)
                If (dr("IdTTr_Huyen").ToString() <> "") Then
                    cb_ttr_huyen.SelectedIndex = CType(arr_TTr_Huyen.IndexOf(dr("IdTTr_Huyen").ToString()), Int32)
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
                edt_sofax.Text = dr("SoFax").ToString().Trim()
                edt_email.Text = dr("Email").ToString().Trim()
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

                cb_tpgd.SelectedIndex = IIf(dr("IdThanhPhan_GD").ToString() <> "", CType(arr_Tpgd.IndexOf(dr("IdThanhPhan_GD").ToString()), Int32), 0)
                cb_ut_giadinh.SelectedIndex = IIf(dr("IdUT_GD").ToString() <> "", CType(arr_Ut_Giadinh.IndexOf(dr("IdUT_GD").ToString()), Int32), 0)
                'Load dữ liệu Ưu tiên bản thân
                If (dr("IdUT_BT").ToString() <> "") Then
                    _Globals.SetItems_CheckedListBox(clb_ut_banthan, dr("IdUT_BT").ToString(), arr_Ut_Banthan)
                End If

                'Tab - Thông tin hồ sơ khác
                cb_tdvh.SelectedIndex = IIf(dr("IdTrinhDoVH").ToString() <> "", CType(arr_Tdvh.IndexOf(dr("IdTrinhDoVH").ToString()), Int32), 0)
                cb_tdct.SelectedIndex = IIf(dr("IdTrinhDoCT").ToString() <> "", CType(arr_Tdct.IndexOf(dr("IdTrinhDoCT").ToString()), Int32), 0)
                cb_tt_honnhan.SelectedIndex = IIf(dr("HonNhan_Cd").ToString() <> "", ARL_TT_HonNhan.IndexOf(dr("HonNhan_Cd").ToString()), 0)
                If dr("CM_Ngay").ToString().Trim() <> "" Then
                    If (CType(dr("CM_Ngay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                        dtpk_ngay_cachmang.Checked = False
                    Else
                        dtpk_ngay_cachmang.Checked = True
                        dtpk_ngay_cachmang.Value = CType(dr("CM_Ngay").ToString(), DateTime)
                    End If
                Else
                    dtpk_ngay_cachmang.Checked = False
                End If
                edt_cm_tochuc.Text = dr("CM_ToChuc").ToString().Trim()
                edt_sotruong_ct.Text = dr("SoTruong_CT").ToString().Trim()
                edt_congviec_lamlau.Text = dr("CV_Lau").ToString().Trim()
                edt_nh_makh.Text = dr("NH_MaKH").ToString().Trim()
                edt_nh_sotk.Text = dr("NH_SoTK").ToString().Trim()
                edt_nh_tennh.Text = dr("NH_TenNH").ToString().Trim()
                edt_masothue.Text = dr("MaSoThue").ToString().Trim()
                edt_bhxh_so.Text = dr("BHXH_SoSo").ToString().Trim()
                If dr("BHXH_NgayLam").ToString().Trim() <> "" Then
                    If (CType(dr("BHXH_NgayLam"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                        dtpk_bhxh_ngaylam.Checked = False
                    Else
                        dtpk_bhxh_ngaylam.Checked = True
                        dtpk_bhxh_ngaylam.Value = CType(dr("BHXH_NgayLam").ToString(), DateTime)
                    End If
                Else
                    dtpk_bhxh_ngaylam.Checked = False
                End If

                If dr("BHXH_NgayDong").ToString().Trim() <> "" Then
                    If (CType(dr("BHXH_NgayDong"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                        dtpk_bhxh_ngaydong.Checked = False
                    Else
                        dtpk_bhxh_ngaydong.Checked = True
                        dtpk_bhxh_ngaydong.Value = CType(dr("BHXH_NgayDong").ToString(), DateTime)
                    End If
                Else
                    dtpk_bhxh_ngaydong.Checked = False
                End If
                edt_bhxh_noilam.Text = dr("BHXH_NoiLam").ToString().Trim()
                edt_ghichu.Text = dr("GhiChu").ToString().Trim()
                Dim strSQL As String = String.Format("Select * from HSCB_TS_HDLD Where IdCanbo = '{0}' Order by DenNgay Desc", _IdCB_TS)
                Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (_db Is Nothing) Then
                        If (_db.Rows.Count > 0) Then
                            cb_qd_cm_moi.SelectedIndex = IIf(_db.Rows(0)("IdChuyenMon").ToString() <> "", CType(arrQd_ChMon_moi.IndexOf(_db.Rows(0)("IdChuyenMon").ToString()), Integer), 0)
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện thiết lập điều khiển theo quyền được phân
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        'Tại đây chỉ cần check quyền với - Hồ sơ khác
        If (FlagHS = True) Then
            Dim _roles As String = Globals.Roles
            'Trường hợp không có quyền nào của Hồ sơ công tác
            If Not (Globals.IsIntersect(";75;76;77;78;", _roles)) Then
                If tctrl_main.TabPages.Contains(tp_hs_congtac) Then
                    tctrl_main.TabPages.Remove(tp_hs_congtac)
                    If TabVal > 0 Then TabVal = TabVal - 1
                End If
            End If
            'Trường hợp không có quyền nào của Hồ sơ Xuất ngoại
            If Not (Globals.IsIntersect(";79;80;81;82;", _roles)) Then
                If tctrl_main.TabPages.Contains(tp_xuatngoai) Then
                    tctrl_main.TabPages.Remove(tp_xuatngoai)
                    If TabVal > 0 Then TabVal -= 1
                End If
            End If
            'Trường hợp không có quyền nào của Hồ sơ Hộ chiếu
            If Not (Globals.IsIntersect(";83;84;85;86;", _roles)) Then
                If tctrl_main.TabPages.Contains(tp_hochieu) Then
                    tctrl_main.TabPages.Remove(tp_hochieu)
                    If TabVal > 0 Then TabVal -= 1
                End If
            End If
            'Trường hợp không có quyền nào của Hồ sơ Lực lượng vũ trang
            If Not (Globals.IsIntersect(";87;88;89;90;", _roles)) Then
                If tctrl_main.TabPages.Contains(tp_llvt) Then
                    tctrl_main.TabPages.Remove(tp_llvt)
                    If TabVal > 0 Then TabVal -= 1
                End If
            End If
            'Trường hợp không có quyền nào của Hồ sơ cũ
            If Not (Globals.IsIntersect(";91;92;93;94;", _roles)) Then
                If tctrl_main.TabPages.Contains(tp_hs_cu) Then
                    tctrl_main.TabPages.Remove(tp_hs_cu)
                    If TabVal > 0 Then TabVal -= 1
                End If
            End If
            'Trường hợp không có quyền nào của Hồ sơ Học hàm học vị
            If Not (Globals.IsIntersect(";95;96;97;98;", _roles)) Then
                If tctrl_main.TabPages.Contains(tp_hh_hv) Then
                    tctrl_main.TabPages.Remove(tp_hh_hv)
                    If TabVal > 0 Then TabVal -= 1
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện select row trên lưới dữ liệu --> Fill dữ liệu vào các controls
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi (Id record) đang select</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String)
        _IdEdit = _Idnew
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_congtac"
                Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, _IdEdit, 1)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If db.Rows(0)("TuNgay").ToString().Trim() <> "" Then
                                dtpk_ct_tungay.Value = CType(db.Rows(0)("TuNgay").ToString(), DateTime)
                            End If
                            If (db.Rows(0)("DenNgay").ToString().Trim() <> "") Then
                                If (CType(db.Rows(0)("DenNgay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                    dtpk_ct_denngay.Checked = False
                                Else
                                    dtpk_ct_denngay.Checked = True
                                    dtpk_ct_denngay.Value = CType(db.Rows(0)("DenNgay").ToString().Trim(), DateTime)
                                End If
                            Else
                                dtpk_ct_denngay.Checked = False
                            End If
                            edt_ct_chucvu.Text = db.Rows(0)("ChucVu").ToString()
                            edt_ct_diachi.Text = db.Rows(0)("DiaChi").ToString()
                            edt_ct_lydo.Text = db.Rows(0)("LyDo").ToString()
                            edt_ct_ghichu.Text = db.Rows(0)("GhiChu").ToString()
                        End If
                    End If
                End Using
            Case "tp_xuatngoai"
                Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, _IdEdit, 2)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If db.Rows(0)("TuNgay").ToString().Trim() <> "" Then
                                dtpk_xn_tungay.Value = CType(db.Rows(0)("TuNgay").ToString(), DateTime)
                            End If
                            If db.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                dtpk_xn_denngay.Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime)
                            End If
                            cb_xn_nuocden.SelectedIndex = CType(arr_Quocgiaden.IndexOf(db.Rows(0)("IdNuocDen").ToString()), Int32)
                            cb_xn_chucvu.SelectedIndex = CType(arr_Chucvu.IndexOf(db.Rows(0)("IdCV_NguoiKy_QD").ToString()), Int32)
                            edt_xn_mucdich.Text = db.Rows(0)("MucDich").ToString()
                            edt_xn_soqd.Text = db.Rows(0)("SoQD").ToString()
                            If db.Rows(0)("NgayKy_QD").ToString().Trim() <> "" Then
                                dtpk_xn_ngayky.Value = CType(db.Rows(0)("NgayKy_QD").ToString(), DateTime)
                            End If
                            edt_xn_nguoiky.Text = db.Rows(0)("NguoiKy_QD").ToString()
                            edt_xn_ghichu.Text = db.Rows(0)("GhiChu").ToString()
                        End If
                    End If
                End Using
            Case "tp_hochieu"
                Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, _IdEdit, 3)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            edt_hc_sohc.Text = db.Rows(0)("So_HoChieu").ToString()
                            cb_hc_loaihc.SelectedIndex = CType(db.Rows(0)("Loai_HC").ToString(), Int32)
                            If db.Rows(0)("NgayCap").ToString().Trim() <> "" Then
                                dtpk_hc_ngaycap.Value = CType(db.Rows(0)("NgayCap").ToString(), DateTime)
                            End If
                            If db.Rows(0)("NgayHH").ToString().Trim() <> "" Then
                                dtpk_hc_ngayhh.Value = CType(db.Rows(0)("NgayHH").ToString(), DateTime)
                            End If
                            edt_hc_noicap.Text = db.Rows(0)("NoiCap").ToString()
                            cb_hc_tinhtrang.SelectedIndex = CType(db.Rows(0)("TinhTrang").ToString(), Int32)
                            edt_hc_ghichu.Text = db.Rows(0)("GhiChu").ToString()
                        End If
                    End If
                End Using
            Case "tp_llvt"
                'Lấy Id công tác tại bản ghi đang Select
                Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, _IdEdit, 4)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If db.Rows(0)("TuNgay").ToString().Trim() <> "" Then
                                dtpk_llvt_tungay.Value = CType(db.Rows(0)("TuNgay").ToString(), DateTime)
                            End If
                            If db.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                dtpk_llvt_denngay.Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime)
                            End If
                            cb_llvt_phanloai.SelectedIndex = CType(arr_PL_LLVT.IndexOf(db.Rows(0)("IdLoaiLLVT").ToString()), Int32)
                            cb_llvt_quanham.SelectedIndex = CType(arr_Quanham.IndexOf(db.Rows(0)("IdQuanHam").ToString()), Int32)
                            edt_llvt_chucvu.Text = db.Rows(0)("ChucVu").ToString()
                            edt_llvt_donvi.Text = db.Rows(0)("DonVi").ToString()
                            edt_llvt_ghichu.Text = db.Rows(0)("GhiChu").ToString()
                        End If
                    End If
                End Using
            Case "tp_hs_cu"
                Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, _IdEdit, 5)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If db.Rows(0)("TuThang").ToString().Trim() <> "" Then
                                dtpk_hscu_tungay.Value = CType(db.Rows(0)("TuThang").ToString(), DateTime)
                            End If
                            If db.Rows(0)("DenThang").ToString().Trim() <> "" Then
                                dtpk_hscu_denngay.Value = CType(db.Rows(0)("DenThang").ToString(), DateTime)
                            End If
                            edt_hscu_diachi.Text = db.Rows(0)("DiaChi").ToString()
                            edt_hscu_nghenghiep.Text = db.Rows(0)("NgheNghiep").ToString()
                            edt_hscu_ghichu.Text = db.Rows(0)("GhiChu").ToString()
                        End If
                    End If
                End Using
        End Select
    End Sub
#End Region

#Region "---> Functions Update data - Hàm cập nhật dữ liệu <---"
    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - hồ sơ nhân sự
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Update_Personals()
        If (FlagEvent = 1) Then         'Cập nhật thêm mới hồ sơ cán bộ
            If (Valid(0)) Then
                Try
                    Dim obj_personal As clsHS_CanBo.HS_CanBo = New clsHS_CanBo.HS_CanBo()
                    obj_personal.MaCB = Globals.Find_Replace(edt_macb.Text.ToString().Trim())
                    obj_personal.HoTen = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hoten.Text.ToString().Trim()))
                    obj_personal.TenThuongGoi = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_ten_tg.Text.ToString().Trim()))
                    obj_personal.BiDanh = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_bidanh.Text.ToString().Trim()))
                    obj_personal.GioiTinh = CType(IIf(rb_nam.Checked = True, 0, 1), Byte)
                    obj_personal.NgaySinh = dtpk_ngaysinh.Value
                    obj_personal.IdDonVi = CType(IIf(arr_Donvi.Count > 0, arr_Donvi(cb_donvi.SelectedIndex), "0"), Integer)
                    obj_personal.IdQuocTich = CType(IIf(arr_Quoctich.Count > 0, arr_Quoctich(cb_quoctich.SelectedIndex), "0"), Int32)
                    obj_personal.IdDanToc = CType(IIf(arr_Dantoc.Count > 0, arr_Dantoc(cb_dantoc.SelectedIndex), "0"), Int32)
                    obj_personal.IdTonGiao = CType(IIf(arr_Tongiao.Count > 0, arr_Tongiao(cb_tongiao.SelectedIndex), "0"), Int32)
                    obj_personal.CMT_So = Globals.Find_Replace(edt_cmt_so.Text.ToString().Trim())
                    obj_personal.CMT_NgayCap = dtpk_cmt_ngaycap.Value
                    obj_personal.CMT_NoiCap = Globals.Find_Replace(edt_cmt_noicap.Text.Trim.ToString())
                    obj_personal.IdNS_Tinh = CType(IIf(arr_Ns_Tinh.Count > 0, arr_Ns_Tinh(cb_ns_tinh.SelectedIndex), "0"), Int32)
                    If (cb_ns_huyen.Items.Count <> 0) Then
                        obj_personal.IdNS_Huyen = CType(IIf(arr_Ns_Huyen.Count > 0, arr_Ns_Huyen(cb_ns_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNS_Huyen = 0
                    End If
                    If (cb_ns_xa.Items.Count <> 0) Then
                        obj_personal.IdNS_xa = CType(IIf(arr_Ns_Xa.Count > 0, arr_Ns_Xa(cb_ns_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNS_Xa = 0
                    End If
                    If (cb_ns_thon.Items.Count <> 0) Then
                        obj_personal.IdNS_Thon = CType(IIf(arr_Ns_Thon.Count > 0, arr_Ns_Thon(cb_ns_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNS_Thon = 0
                    End If
                    obj_personal.NS_DChi = Globals.Find_Replace(edt_ns_diachi.Text.Trim.ToString())

                    obj_personal.IdNQ_Tinh = CType(IIf(arr_Nq_Tinh.Count > 0, arr_Nq_Tinh(cb_nq_tinh.SelectedIndex), "0"), Int32)
                    If cb_nq_huyen.Items.Count <> 0 Then
                        obj_personal.IdNQ_Huyen = CType(IIf(arr_Nq_Huyen.Count > 0, arr_Nq_Huyen(cb_nq_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNQ_Huyen = 0
                    End If
                    If cb_nq_xa.Items.Count <> 0 Then
                        obj_personal.IdNQ_Xa = CType(IIf(arr_Nq_Xa.Count > 0, arr_Nq_Xa(cb_nq_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNQ_Xa = 0
                    End If
                    If cb_nq_thon.Items.Count <> 0 Then
                        obj_personal.IdNQ_Thon = CType(IIf(arr_Nq_Thon.Count > 0, arr_Nq_Thon(cb_nq_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNQ_Thon = 0
                    End If
                    obj_personal.NQ_DChi = Globals.Find_Replace(edt_nq_diachi.Text.Trim.ToString())

                    obj_personal.IdThT_Tinh = CType(IIf(arr_TT_Tinh.Count > 0, arr_TT_Tinh(cb_tt_tinh.SelectedIndex), "0"), Int32)
                    If (cb_tt_huyen.Items.Count <> 0) Then
                        obj_personal.IdThT_Huyen = CType(IIf(arr_TT_Huyen.Count > 0, arr_TT_Huyen(cb_tt_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdThT_Huyen = 0
                    End If
                    If (cb_tt_xa.Items.Count <> 0) Then
                        obj_personal.IdThT_Xa = CType(IIf(arr_TT_Xa.Count > 0, arr_TT_Xa(cb_tt_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdThT_Xa = 0
                    End If
                    If (cb_tt_thon.Items.Count <> 0) Then
                        obj_personal.IdThT_Thon = CType(IIf(arr_TT_Thon.Count > 0, arr_TT_Thon(cb_tt_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdThT_Thon = 0
                    End If
                    obj_personal.ThT_Diachi = Globals.Find_Replace(edt_tt_diachi.Text.Trim.ToString())
                    obj_personal.ThT_Dienthoai = Globals.Find_Replace(edt_tt_dienthoai.Text.Trim.ToString())

                    obj_personal.IdTTr_Tinh = CType(IIf(arr_TTr_Tinh.Count > 0, arr_TTr_Tinh(cb_ttr_tinh.SelectedIndex), "0"), Int32)
                    If cb_ttr_huyen.Items.Count <> 0 Then
                        obj_personal.IdTTr_Huyen = CType(IIf(arr_TTr_Huyen.Count > 0, arr_TTr_Huyen(cb_ttr_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdTTr_Huyen = 0
                    End If
                    If cb_ttr_xa.Items.Count <> 0 Then
                        obj_personal.IdTTr_Xa = CType(IIf(arr_TTr_Xa.Count > 0, arr_TTr_Xa(cb_ttr_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdTTr_Xa = 0
                    End If
                    If cb_ttr_thon.Items.Count <> 0 Then
                        obj_personal.IdTTr_Thon = CType(IIf(arr_TTr_Thon.Count > 0, arr_TTr_Thon(cb_ttr_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdTTr_Thon = 0
                    End If
                    obj_personal.TTr_Diachi = Globals.Find_Replace(edt_ttr_diachi.Text.Trim.ToString())
                    obj_personal.TTr_Dienthoai = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_ttr_dienthoai.Text.Trim.ToString()))
                    obj_personal.DienThoai_CQ = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_dt_coquan.Text.Trim.ToString()))
                    obj_personal.DienThoai_DD = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_dt_didong.Text.Trim.ToString()))
                    obj_personal.DienThoai_NR = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_dt_nharieng.Text.Trim.ToString()))
                    obj_personal.SoFax = Globals.RemoveCharacterNonNumber(Globals.Find_Replace(edt_sofax.Text.Trim.ToString()))
                    obj_personal.Email = Globals.Find_Replace(edt_email.Text.Trim.ToString())
                    obj_personal.NhomMau = IIf(cb_nhommau.SelectedIndex > 0, cb_nhommau.SelectedItem.ToString(), "")
                    obj_personal.IdThanhPhanGD = CType(IIf(arr_Tpgd.Count > 0, arr_Tpgd(cb_tpgd.SelectedIndex), "0"), Int32)
                    obj_personal.IdUT_BThan = _Globals.GetItems_CheckedListBox(clb_ut_banthan, arr_Ut_Banthan)
                    obj_personal.IdUT_GDinh = CType(IIf(arr_Ut_Giadinh.Count > 0, arr_Ut_Giadinh(cb_ut_giadinh.SelectedIndex), "0"), Int32)
                    obj_personal.DacDiem_BT = Globals.Find_Replace(edt_dacdiem_banthan.Text.Trim.ToString())
                    obj_personal.QuanHe_Nguoi_NN = Globals.Find_Replace(edt_quanhe_ng_nn.Text.Trim.ToString())
                    obj_personal.IdTrinhDoVH = CType(IIf(arr_Tdvh.Count > 0, arr_Tdvh(cb_tdvh.SelectedIndex), "0"), Int32)
                    obj_personal.IdTrinhDoCT = CType(IIf(arr_Tdct.Count > 0, arr_Tdct(cb_tdct.SelectedIndex), "0"), Int32)
                    obj_personal.HonNhan_Cd = IIf(ARL_TT_HonNhan.Count > 0, ARL_TT_HonNhan(cb_tt_honnhan.SelectedIndex), "")

                    obj_personal.Ngay_ThamNien = IIf(dtpk_ngay_thamnien.Checked, dtpk_ngay_thamnien.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_NH = IIf(dtpk_ngay_vaonganh.Checked, dtpk_ngay_vaonganh.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_VBSP = IIf(dtpk_ngay_vaonhcsxh.Checked, dtpk_ngay_vaonhcsxh.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_BienChe = IIf(dtpk_ngay_bc_nhcsxh.Checked, dtpk_ngay_bc_nhcsxh.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_CQ = dtpk_qd_ngayhl.Value
                    obj_personal.CM_Ngay = IIf(dtpk_ngay_cachmang.Checked, dtpk_ngay_cachmang.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.CM_ToChuc = Globals.Find_Replace(edt_cm_tochuc.Text.Trim.ToString())
                    obj_personal.SoTruong_CT = Globals.Find_Replace(edt_sotruong_ct.Text.Trim.ToString())
                    obj_personal.CV_Lau = Globals.Find_Replace(edt_congviec_lamlau.Text.Trim.ToString())
                    obj_personal.NH_MaKH = Globals.Find_Replace(edt_nh_makh.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", ""))
                    obj_personal.NH_SoTK = Globals.Find_Replace(edt_nh_sotk.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", ""))
                    obj_personal.NH_TenNH = Globals.Find_Replace(edt_nh_tennh.Text.Trim.ToString())
                    obj_personal.MaSoThue = Globals.Find_Replace(edt_masothue.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", ""))
                    obj_personal.BHXH_SoSo = Globals.Find_Replace(edt_bhxh_so.Text.Trim.ToString())
                    obj_personal.BHXH_NgaySo = IIf(dtpk_bhxh_ngaylam.Checked, dtpk_bhxh_ngaylam.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.BHXH_NgayBatDau = IIf(dtpk_bhxh_ngaydong.Checked, dtpk_bhxh_ngaydong.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.BHXH_NoiLam = Globals.Find_Replace(edt_bhxh_noilam.Text.Trim.ToString())
                    obj_personal.GhiChu = Globals.Find_Replace(edt_ghichu.Text.Trim.ToString())
                    obj_personal.IdNew = ""

                    'Ảnh thẻ
                    obj_personal.AnhThe = My_CImgToByte(picbx_main.Image)

                    Dim strCode As String = ""
                    strCode = _HS_CanBo.Insert_Human(obj_personal)

                    'Lấy dữ liệu - Quyết định nhân sự để thêm mới 1 quyết định nhân sự của cán bộ
                    'Dim obj_qdnhansu As clsHS_CanBo.QD_NhanSu = New clsHS_CanBo.QD_NhanSu
                    Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                    m_QDNhanSu.IdCanBo = strCode
                    m_QDNhanSu.So_QD = Globals.Find_Replace(edt_qd_soqd.Text.Trim.ToString())
                    m_QDNhanSu.NgayKy_QD = dtpk_qd_ngayky.Value
                    m_QDNhanSu.IdLoaiQD = CType(IIf(arrQd_LoaiQd.Count > 0, arrQd_LoaiQd(cb_qd_loaiqd.SelectedIndex), "0"), Int32)
                    m_QDNhanSu.NgayHL = dtpk_qd_ngayhl.Value
                    m_QDNhanSu.NgayBoNhiem_TT = DateTime.Parse("01/01/1900")
                    m_QDNhanSu.NgayThoiLuong = DateTime.Parse("01/01/1900")
                    m_QDNhanSu.NguoiKy_QD = standardizeName(edt_qd_nguoiky.Text.Trim.ToString())
                    m_QDNhanSu.idCV_Nguoiky_QD = CType(IIf(arrQd_Chucvu.Count > 0, arrQd_Chucvu(cb_qd_chucvu.SelectedIndex), "0"), Int32)
                    m_QDNhanSu.IdDonvi_Cu = 0
                    m_QDNhanSu.IdPhong_Cu = 0
                    m_QDNhanSu.IdChucvu_Cu = 0
                    m_QDNhanSu.IdChuyenMon_Cu = 0
                    m_QDNhanSu.IdDonvi_Moi = CType(IIf(arrQd_Donvi_moi.Count > 0, arrQd_Donvi_moi(cb_qd_donvi_moi.SelectedIndex), "0"), Int32)
                    If (cb_qd_phongban_moi.Items.Count <> 0) Then
                        m_QDNhanSu.IdPhong_Moi = CType(IIf(arrQd_Phongban_moi.Count > 0, arrQd_Phongban_moi(cb_qd_phongban_moi.SelectedIndex), "0"), Int32)
                    Else : m_QDNhanSu.IdPhong_Moi = 0
                    End If
                    m_QDNhanSu.IdChucvu_Moi = CType(IIf(arrQd_Chucvu_moi.Count > 0, arrQd_Chucvu_moi(cb_qd_chucvu_moi.SelectedIndex), "0"), Int32)
                    m_QDNhanSu.IdChuyenMon_Moi = CType(IIf(arrQd_ChMon_moi.Count > 0, arrQd_ChMon_moi(cb_qd_cm_moi.SelectedIndex), "0"), Int32)
                    m_QDNhanSu.GhiChu = standardizeString(edt_qd_ghichu.Text.Trim.ToString())
                    m_QDNhanSu.Active = 1
                    m_QDNhanSu.IsKiemNhiem = 0
                    m_QDNhanSu.IsQD_NHCS = 1
                    m_QDNhanSu.DVraQD = getDonvi(IdDONVI)
                    m_QDNhanSu.CV_NguoiKy_QD = ""
                    m_QDNhanSu.LoaiQD = ""
                    m_QDNhanSu.DenNgay = DateTime.MinValue
                    m_QDNhanSu.NoiDung = ""
                    m_QDNhanSu.Add()
                    '_HS_CanBo.Insert_Decision(obj_qdnhansu)
                    'Cập nhật lại trạng thái cho hồ sơ cán bộ tập sự nếu lấy dữ liệu từ Hồ sơ cán bộ tập sự
                    If (_IdCB_TS <> "") Then
                        _HS_CanBo.Update_HSCB_TS_WhenMove(_IdCB_TS, strCode, 1, 0)
                        _HS_CanBo.Update_HSCB_TS_HDLD_WhenMove(_IdCB_TS, m_QDNhanSu.NgayHL)
                    End If
                    lbl_thongbao.Text = " Bạn đã thêm mới thành công thông tin hồ sơ nhân sự!"
                    _IdCanBo = strCode
                    FlagEvent = 2   'Sửa đổi hồ sơ cán bộ
                    'Thực hiện Lock controls của quyết định nhân sự -->Không cho sửa
                    SetStatus_Controls(True, 1)

                    'Thực hiện việc hỏi xem có muốn thêm mới dữ liệu Hồ sơ khác cho cán bộ vừa thêm không ?
                    If (strCode <> "") Then
                        If (MessageBox.Show("Bạn có muốn cập nhật luôn thông tin hồ sơ khác của cán bộ vừa thêm mới không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện load thông tin hồ sơ khác
                            FlagHS = True
                            'Thực hiện gọi hàm định nghĩa các lưới dữ liệu
                            '_HS_CanBo.Create_Frame(dgv_congtac, 2)
                            _HS_CanBo.Create_Frame(dgv_xuatngoai, 3)
                            _HS_CanBo.Create_Frame(dgv_hochieu, 4)
                            _HS_CanBo.Create_Frame(dgv_llvt, 5)
                            '_HS_CanBo.Create_Frame(dgv_hosocu, 6)
                            'Thực hiện fill dữ liệu vào các ComboBox
                            FillData_ComboBox()

                            'Thực hiện add tab - Hồ sơ khác
                            'If Not (tctrl_main.TabPages.Contains(tp_hs_congtac)) Then tctrl_main.TabPages.Add(tp_hs_congtac)
                            If Not (tctrl_main.TabPages.Contains(tp_xuatngoai)) Then tctrl_main.TabPages.Add(tp_xuatngoai)
                            If Not (tctrl_main.TabPages.Contains(tp_hochieu)) Then tctrl_main.TabPages.Add(tp_hochieu)
                            If Not (tctrl_main.TabPages.Contains(tp_llvt)) Then tctrl_main.TabPages.Add(tp_llvt)
                            'If Not (tctrl_main.TabPages.Contains(tp_hs_cu)) Then tctrl_main.TabPages.Add(tp_hs_cu)
                            If Not (tctrl_main.TabPages.Contains(tp_hh_hv)) Then tctrl_main.TabPages.Add(tp_hh_hv)

                            'Thực hiện remove và add tab - cần dùng (Tab - Hồ sơ Khác)
                            If tctrl_main.TabPages.Contains(tp_chitiet_hscb) Then
                                tctrl_main.TabPages.Remove(tp_chitiet_hscb)
                            End If

                            btn_add.Visible = True
                            btn_delete.Visible = True
                            lbl_div_hsk.Visible = True

                            'Reset controls - Thông tin chung để xem về cán bộ
                            ResetAll_Controls(True, True)
                            'Fill data - Thông tin chung để xem
                            Fill_Data(10, True)
                            'Fill dữ liệu các hồ sơ khác ra để thao tác
                            'Fill_Data(2, False)
                            Fill_Data(3, False)
                            Fill_Data(4, False)
                            Fill_Data(5, False)
                            'Fill_Data(6, False)
                            Fill_Data(7, False)
                            Fill_Data(8, False)
                            'Thực hiện khoá các điều khiển - Thông tin chung cán bộ
                            SetStatus_Controls(True, 0)
                        End If
                    End If
                    'Thực hiện thiết lập --> Khoá các controls
                Catch ex As Exception
                    MessageBox.Show("Cập nhật thêm mới hồ sơ nhân sự: " + ex.Message.ToString(), "Cập nhật dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Return
                End Try
            End If
        ElseIf (FlagEvent = 2) Then     'Cập nhật sửa đổi hồ sơ cán bộ
            If (Valid(0)) Then
                Try
                    Dim obj_personal As clsHS_CanBo.HS_CanBo = New clsHS_CanBo.HS_CanBo
                    obj_personal.IdCanBo = _IdCanBo
                    obj_personal.MaCB = Globals.Find_Replace(edt_macb.Text.ToString().Trim())
                    obj_personal.HoTen = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hoten.Text.ToString().Trim()))
                    obj_personal.TenThuongGoi = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_ten_tg.Text.ToString().Trim()))
                    obj_personal.BiDanh = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_bidanh.Text.ToString().Trim()))
                    obj_personal.GioiTinh = CType(IIf(rb_nam.Checked = True, 0, 1), Byte)
                    obj_personal.NgaySinh = dtpk_ngaysinh.Value
                    obj_personal.IdDonVi = CType(IIf(arr_Donvi.Count > 0, arr_Donvi(cb_donvi.SelectedIndex), "0"), Int32)
                    obj_personal.IdQuocTich = CType(IIf(arr_Quoctich.Count > 0, arr_Quoctich(cb_quoctich.SelectedIndex), "0"), Int32)
                    obj_personal.IdDanToc = CType(IIf(arr_Dantoc.Count > 0, arr_Dantoc(cb_dantoc.SelectedIndex), "0"), Int32)
                    obj_personal.IdTonGiao = CType(IIf(arr_Tongiao.Count > 0, arr_Tongiao(cb_tongiao.SelectedIndex), "0"), Int32)
                    obj_personal.CMT_So = Globals.Find_Replace(edt_cmt_so.Text.ToString().Trim())
                    obj_personal.CMT_NgayCap = dtpk_cmt_ngaycap.Value
                    obj_personal.CMT_NoiCap = Globals.Find_Replace(edt_cmt_noicap.Text.Trim.ToString())

                    obj_personal.IdNS_Tinh = CType(IIf(arr_Ns_Tinh.Count > 0, arr_Ns_Tinh(cb_ns_tinh.SelectedIndex), "0"), Int32)
                    If (cb_ns_huyen.Items.Count <> 0) Then
                        If cb_ns_huyen.SelectedIndex = -1 Then cb_ns_huyen.SelectedIndex = 0
                        obj_personal.IdNS_Huyen = CType(IIf(arr_Ns_Huyen.Count > 0, arr_Ns_Huyen(cb_ns_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNS_Huyen = 0
                    End If
                    If (cb_ns_xa.Items.Count <> 0) Then
                        If cb_ns_xa.SelectedIndex = -1 Then cb_ns_xa.SelectedIndex = 0
                        obj_personal.IdNS_Xa = CType(IIf(arr_Ns_Xa.Count > 0, arr_Ns_Xa(cb_ns_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNS_Xa = 0
                    End If
                    If (cb_ns_thon.Items.Count <> 0) Then
                        If cb_ns_thon.SelectedIndex = -1 Then cb_ns_thon.SelectedIndex = 0
                        obj_personal.IdNS_Thon = CType(IIf(arr_Ns_Thon.Count > 0, arr_Ns_Thon(cb_ns_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNS_Thon = 0
                    End If
                    obj_personal.NS_DChi = Globals.Find_Replace(edt_ns_diachi.Text.Trim.ToString())

                    obj_personal.IdNQ_Tinh = CType(IIf(arr_Nq_Tinh.Count > 0, arr_Nq_Tinh(cb_nq_tinh.SelectedIndex), "0"), Int32)
                    If cb_nq_huyen.Items.Count <> 0 Then
                        If cb_nq_huyen.SelectedIndex = -1 Then cb_nq_huyen.SelectedIndex = 0
                        obj_personal.IdNQ_Huyen = CType(IIf(arr_Nq_Huyen.Count > 0, arr_Nq_Huyen(cb_nq_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNQ_Huyen = 0
                    End If
                    If (cb_nq_xa.Items.Count <> 0) Then
                        If cb_nq_xa.SelectedIndex = -1 Then cb_nq_xa.SelectedIndex = 0
                        obj_personal.IdNQ_Xa = CType(IIf(arr_Nq_Xa.Count > 0, arr_Nq_Xa(cb_nq_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNQ_Xa = 0
                    End If
                    If (cb_nq_thon.Items.Count <> 0) Then
                        If cb_nq_thon.SelectedIndex = -1 Then cb_nq_thon.SelectedIndex = 0
                        obj_personal.IdNQ_Thon = CType(IIf(arr_Nq_Thon.Count > 0, arr_Nq_Thon(cb_nq_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdNQ_Thon = 0
                    End If
                    obj_personal.NQ_DChi = Globals.Find_Replace(edt_nq_diachi.Text.Trim.ToString())

                    obj_personal.IdThT_Tinh = CType(IIf(arr_TT_Tinh.Count > 0, arr_TT_Tinh(cb_tt_tinh.SelectedIndex), "0"), Int32)
                    If (cb_tt_huyen.Items.Count <> 0) Then
                        If cb_tt_huyen.SelectedIndex = -1 Then cb_tt_huyen.SelectedIndex = 0
                        obj_personal.IdThT_Huyen = CType(IIf(arr_TT_Huyen.Count > 0, arr_TT_Huyen(cb_tt_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdThT_Huyen = 0
                    End If
                    If (cb_tt_xa.Items.Count <> 0) Then
                        If cb_tt_xa.SelectedIndex = -1 Then cb_tt_xa.SelectedIndex = 0
                        obj_personal.IdThT_Xa = CType(IIf(arr_TT_Xa.Count > 0, arr_TT_Xa(cb_tt_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdThT_Xa = 0
                    End If
                    If (cb_tt_thon.Items.Count <> 0) Then
                        If cb_tt_thon.SelectedIndex = -1 Then cb_tt_thon.SelectedIndex = 0
                        obj_personal.IdThT_Thon = CType(IIf(arr_TT_Thon.Count > 0, arr_TT_Thon(cb_tt_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdThT_Thon = 0
                    End If
                    obj_personal.ThT_Diachi = Globals.Find_Replace(edt_tt_diachi.Text.Trim.ToString())
                    obj_personal.ThT_Dienthoai = Globals.Find_Replace(edt_tt_dienthoai.Text.Trim.ToString())

                    obj_personal.IdTTr_Tinh = CType(IIf(arr_TTr_Tinh.Count > 0, arr_TTr_Tinh(cb_ttr_tinh.SelectedIndex), "0"), Int32)
                    If cb_ttr_huyen.Items.Count <> 0 Then
                        If cb_ttr_huyen.SelectedIndex = -1 Then cb_ttr_huyen.SelectedIndex = 0
                        obj_personal.IdTTr_Huyen = CType(IIf(arr_TTr_Huyen.Count > 0, arr_TTr_Huyen(cb_ttr_huyen.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdTTr_Huyen = 0
                    End If
                    If (cb_ttr_xa.Items.Count <> 0) Then
                        If cb_ttr_xa.SelectedIndex = -1 Then cb_ttr_xa.SelectedIndex = 0
                        obj_personal.IdTTr_Xa = CType(IIf(arr_TTr_Xa.Count > 0, arr_TTr_Xa(cb_ttr_xa.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdTTr_Xa = 0
                    End If
                    If (cb_ttr_thon.Items.Count <> 0) Then
                        If cb_ttr_thon.SelectedIndex = -1 Then cb_ttr_thon.SelectedIndex = 0
                        obj_personal.IdTTr_Thon = CType(IIf(arr_TTr_Thon.Count > 0, arr_TTr_Thon(cb_ttr_thon.SelectedIndex), "0"), Int32)
                    Else : obj_personal.IdTTr_Thon = 0
                    End If
                    obj_personal.TTr_Diachi = Globals.Find_Replace(edt_ttr_diachi.Text.Trim.ToString())
                    obj_personal.TTr_Dienthoai = Globals.Find_Replace(edt_ttr_dienthoai.Text.Trim.ToString())
                    obj_personal.DienThoai_CQ = Globals.Find_Replace(edt_dt_coquan.Text.Trim.ToString())
                    obj_personal.DienThoai_DD = Globals.Find_Replace(edt_dt_didong.Text.Trim.ToString())
                    obj_personal.DienThoai_NR = Globals.Find_Replace(edt_dt_nharieng.Text.Trim.ToString())
                    obj_personal.SoFax = Globals.Find_Replace(edt_sofax.Text.Trim.ToString())
                    obj_personal.Email = Globals.Find_Replace(edt_email.Text.Trim.ToString())
                    obj_personal.NhomMau = IIf(cb_nhommau.SelectedIndex > 0, cb_nhommau.SelectedItem.ToString(), "")
                    obj_personal.IdThanhPhanGD = CType(IIf(arr_Tpgd.Count > 0, arr_Tpgd(cb_tpgd.SelectedIndex), "0"), Int32)
                    obj_personal.IdUT_BThan = _Globals.GetItems_CheckedListBox(clb_ut_banthan, arr_Ut_Banthan)
                    obj_personal.IdUT_GDinh = CType(IIf(arr_Ut_Giadinh.Count > 0, arr_Ut_Giadinh(cb_ut_giadinh.SelectedIndex), "0"), Int32)
                    obj_personal.DacDiem_BT = Globals.Find_Replace(edt_dacdiem_banthan.Text.Trim.ToString())
                    obj_personal.QuanHe_Nguoi_NN = Globals.Find_Replace(edt_quanhe_ng_nn.Text.Trim.ToString())
                    obj_personal.IdTrinhDoVH = CType(IIf(arr_Tdvh.Count > 0, arr_Tdvh(cb_tdvh.SelectedIndex), "0"), Int32)
                    obj_personal.IdTrinhDoCT = CType(IIf(arr_Tdct.Count > 0, arr_Tdct(cb_tdct.SelectedIndex), "0"), Int32)
                    obj_personal.HonNhan_Cd = IIf(ARL_TT_HonNhan.Count > 0, ARL_TT_HonNhan(cb_tt_honnhan.SelectedIndex), "")

                    obj_personal.Ngay_ThamNien = IIf(dtpk_ngay_thamnien.Checked, dtpk_ngay_thamnien.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_NH = IIf(dtpk_ngay_vaonganh.Checked, dtpk_ngay_vaonganh.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_VBSP = IIf(dtpk_ngay_vaonhcsxh.Checked, dtpk_ngay_vaonhcsxh.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_BienChe = IIf(dtpk_ngay_bc_nhcsxh.Checked, dtpk_ngay_bc_nhcsxh.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.Ngay_CQ = dtpk_qd_ngayhl.Value
                    obj_personal.CM_Ngay = IIf(dtpk_ngay_cachmang.Checked, dtpk_ngay_cachmang.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.CM_ToChuc = Globals.Find_Replace(edt_cm_tochuc.Text.Trim.ToString())
                    obj_personal.SoTruong_CT = Globals.Find_Replace(edt_sotruong_ct.Text.Trim.ToString())
                    obj_personal.CV_Lau = Globals.Find_Replace(edt_congviec_lamlau.Text.Trim.ToString())
                    obj_personal.BHXH_SoSo = Globals.Find_Replace(edt_bhxh_so.Text.Trim.ToString())
                    obj_personal.BHXH_NgaySo = IIf(dtpk_bhxh_ngaylam.Checked, dtpk_bhxh_ngaylam.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.BHXH_NgayBatDau = IIf(dtpk_bhxh_ngaydong.Checked, dtpk_bhxh_ngaydong.Value, DateTime.Parse("01/01/1900"))
                    obj_personal.BHXH_NoiLam = Globals.Find_Replace(edt_bhxh_noilam.Text.Trim.ToString())
                    obj_personal.GhiChu = Globals.Find_Replace(edt_ghichu.Text.Trim.ToString())
                    obj_personal.IdNew = ""

                    obj_personal.NH_MaKH = Globals.Find_Replace(edt_nh_makh.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", ""))
                    obj_personal.NH_SoTK = Globals.Find_Replace(edt_nh_sotk.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", ""))
                    obj_personal.NH_TenNH = Globals.Find_Replace(edt_nh_tennh.Text.Trim.ToString())
                    obj_personal.MaSoThue = Globals.Find_Replace(edt_masothue.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", ""))

                    'Ảnh thẻ
                    obj_personal.AnhThe = My_CImgToByte(picbx_main.Image)

                    _HS_CanBo.Update_Human(obj_personal)

                    lbl_thongbao.Text = " Bạn đã sửa đổi thành công thông tin hồ sơ nhân sự!"
                    ' Thực hiện Lock lại các control ---> Nếu có yêu cầu
                Catch ex As Exception
                    MessageBox.Show("Cập nhật sửa đổi hồ sơ nhân sự: " + ex.Message.ToString(), "Cập nhật dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Return
                End Try
            End If
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmCN_HSCanBo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lbl_thongbao.Text = ""
        'Thực hiện gọi hàm định nghĩa các lưới dữ liệu
        _HS_CanBo.Create_Frame(dgv_congtac, 2)
        _HS_CanBo.Create_Frame(dgv_xuatngoai, 3)
        _HS_CanBo.Create_Frame(dgv_hochieu, 4)
        _HS_CanBo.Create_Frame(dgv_llvt, 5)
        '_HS_CanBo.Create_Frame(dgv_hosocu, 6)
        'Thực hiện fill dữ liệu vào các ComboBox
        FillData_ComboBox()
        'Reset controls and fill data

        If (FlagHS = False) Then    'Load thông tin Hồ sơ cán bộ
            'Thực hiện remove và add tab - cần dùng (Tab - Hồ sơ cán bộ)
            If tctrl_main.TabPages.Contains(tp_hs_congtac) Then tctrl_main.TabPages.Remove(tp_hs_congtac)
            If tctrl_main.TabPages.Contains(tp_xuatngoai) Then tctrl_main.TabPages.Remove(tp_xuatngoai)
            If tctrl_main.TabPages.Contains(tp_hochieu) Then tctrl_main.TabPages.Remove(tp_hochieu)
            If tctrl_main.TabPages.Contains(tp_llvt) Then tctrl_main.TabPages.Remove(tp_llvt)
            If tctrl_main.TabPages.Contains(tp_hs_cu) Then tctrl_main.TabPages.Remove(tp_hs_cu)
            If tctrl_main.TabPages.Contains(tp_hh_hv) Then tctrl_main.TabPages.Remove(tp_hh_hv)

            btn_add.Visible = False
            btn_delete.Visible = False
            lbl_div_hsk.Visible = False

            Select Case FlagEvent
                Case 1
                    'Thực hiện thiết lập và Load dữ liệu thêm mới
                    ResetAll_Controls(False, True)
                    'Nếu khi thêm mới cán bộ --> Mà chọn dữ liệu từ cán bộ tập sự
                    If (_IdCB_TS <> "") Then
                        Fill_Data()
                    End If
                Case 2
                    'Thực hiện thiết lập và Load dữ liệu sửa đổi
                    ResetAll_Controls(False, True)
                    'Fill data của bản ghi cần sửa ra
                    Fill_Data(0, True)
                    Fill_Data(1, False) 'Fill dữ liệu Quyết định nhân sự

                    ''QUYETDINH''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Thực hiện Lock controls của quyết định nhân sự -->Không cho sửa
                    SetStatus_Controls(True, 1)
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End Select
        Else                        'Load thông tin Hồ sơ Khác
            'Thực hiện remove và add tab - cần dùng (Tab - Hồ sơ Khác)
            If tctrl_main.TabPages.Contains(tp_chitiet_hscb) Then tctrl_main.TabPages.Remove(tp_chitiet_hscb)
            If tctrl_main.TabPages.Contains(tp_hs_cu) Then tctrl_main.TabPages.Remove(tp_hs_cu)
            If tctrl_main.SelectedTab.Name = "tp_hs_congtac" Then
                btn_add.Visible = False
                btn_delete.Visible = False
                lbl_div_hsk.Visible = False
                pnl_choice.Visible = False
                btn_huybo.Visible = False
                btn_save.Visible = False
            Else
                btn_add.Visible = True
                btn_delete.Visible = True
                lbl_div_hsk.Visible = True
                pnl_choice.Visible = False
            End If
            'If tctrl_main.TabPages.Contains(tp_hs_congtac) Then
            '    btn_add.Visible = False
            '    btn_delete.Visible = False
            '    lbl_div_hsk.Visible = False
            '    pnl_choice.Visible = False
            'Else
            '    btn_add.Visible = True
            '    btn_delete.Visible = True
            '    lbl_div_hsk.Visible = True
            '    pnl_choice.Visible = False
            'End If
            'Reset controls - Thông tin chung để xem về cán bộ
            ResetAll_Controls(True, True)
            'Fill data - Thông tin chung để xem
            Fill_Data(10, True)
            'Fill dữ liệu các hồ sơ khác ra để thao tác
            Fill_Data(2, False)
            Fill_Data(3, False)
            Fill_Data(4, False)
            Fill_Data(5, False)
            'Fill_Data(6, False)
            Fill_Data(7, False)
            Fill_Data(8, False)
            'Thực hiện khoá các điều khiển - Thông tin chung cán bộ
            'If tctrl_main.TabPages.Contains(tp_hs_congtac) Then
            '    btn_huybo.Visible = False
            '    btn_save.Visible = False
            'Else
            '    SetStatus_Controls(True, 0)
            'End If
            'Kiểm tra quyền người sử dụng
            Check_Permits()
            'Nếu được gọi từ menu trên lưới Hồ sơ cán bộ
            If (TabVal > 0) Then
                tctrl_main.SelectedIndex = TabVal - 1
            End If
        End If

        'Kiểm tra đơn vị người sửa và người được sửa có giống nhau không <--- có vẻ thừa nên tạm bỏ
        'If _IdCanBo <> "" Then
        '    If checkRight_CreateRecord(_IdCanBo) Then
        '        btn_save.Enabled = True
        '        btn_delete.Enabled = True
        '    Else
        '        btn_save.Enabled = False
        '        btn_delete.Enabled = False
        '    End If
        'End If

    End Sub

    Private Sub btn_chonanh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_chonanh.Click
        Dim dlg As New OpenFileDialog
        dlg.Title = "Chọn file hình ảnh"
        dlg.Filter = "Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*"
        If dlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Try
                picbx_main.Image = My_CImgToThumbnail(Image.FromFile(dlg.FileName), picbx_main)
            Catch ex As Exception
                picbx_main.Image = Nothing
                My_MessageBox("Lỗi xảy ra khi đọc file ảnh!")
            End Try
        End If
    End Sub

    Private Sub btn_xoaanh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_xoaanh.Click
        picbx_main.Image = Nothing
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Me.Close()
    End Sub

    Private Sub frmCN_HSCanBo_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_huybo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_huybo.Click
        lbl_thongbao.Text = ""
        If (FlagHS = False) Then    'Load thông tin Hồ sơ cán bộ
            Select Case FlagEvent
                Case 1
                    'Thực hiện thiết lập và Load dữ liệu thêm mới
                    ResetAll_Controls(False, True)
                Case 2
                    'Thực hiện thiết lập và Load dữ liệu sửa đổi
                    ResetAll_Controls(False, True)
                    'Fill data của bản ghi cần sửa ra
                    Fill_Data(0, True)
            End Select
        Else                        'Load thông tin Hồ sơ Khác
            Dim _currRow As String = _IdEdit
            'Reset controls - Thông tin chung để xem về cán bộ
            _IdEdit = ""

            'Trả lại dữ liệu vừa load ra
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_hs_congtac"
                    'Reset Tab - Công Tác
                    dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
                    dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
                    dtpk_ct_denngay.Checked = False
                    edt_ct_chucvu.Text = ""
                    edt_ct_diachi.Text = ""
                    edt_ct_lydo.Text = ""
                    edt_ct_ghichu.Text = ""
                    If (dgv_congtac.Rows.Count <= 0) Then Return
                    If (_currRow = "") Then _currRow = dgv_congtac.CurrentRow.Cells("cln_Id").Value.ToString()
                    dgv_congtac.CurrentRow.Selected = False
                    dgv_congtac.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_congtac, "cln_Id")).Selected = True
                    If (dgv_congtac.Rows.Count > 0) Then
                        If ((dgv_congtac.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_congtac.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                Case "tp_xuatngoai"
                    'Reset Tab Xuất Ngoại
                    dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
                    dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()
                    cb_xn_nuocden.SelectedIndex = 0
                    edt_xn_mucdich.Text = ""
                    edt_xn_soqd.Text = ""
                    dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
                    edt_xn_nguoiky.Text = ""
                    edt_xn_ghichu.Text = ""
                    cb_xn_chucvu.SelectedIndex = 0
                    'Select lại row hiện tại trước đó của người dùng
                    If (dgv_xuatngoai.Rows.Count <= 0) Then Return
                    If (_currRow = "") Then _currRow = dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value.ToString()
                    dgv_xuatngoai.CurrentRow.Selected = False
                    dgv_xuatngoai.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_xuatngoai, "cln_Id")).Selected = True
                    If (dgv_xuatngoai.Rows.Count > 0) Then
                        If ((dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                Case "tp_hochieu"
                    'Reset Tab Hộ Chiếu
                    edt_hc_sohc.Text = ""
                    cb_hc_loaihc.SelectedIndex = 0
                    dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
                    edt_hc_noicap.Text = ""
                    dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
                    cb_hc_tinhtrang.SelectedIndex = 0
                    edt_hc_ghichu.Text = ""
                    'Select lại row hiện tại trước đó của người dùng
                    If (dgv_hochieu.Rows.Count <= 0) Then Return
                    If (_currRow = "") Then _currRow = dgv_hochieu.CurrentRow.Cells("cln_Id").Value.ToString()
                    dgv_hochieu.CurrentRow.Selected = False
                    dgv_hochieu.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hochieu, "cln_Id")).Selected = True
                    If (dgv_hochieu.Rows.Count > 0) Then
                        If ((dgv_hochieu.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_hochieu.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                Case "tp_llvt"
                    'Reset Tab - Lực lượng vũ trang
                    dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
                    dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()
                    cb_llvt_phanloai.SelectedIndex = 0
                    cb_llvt_quanham.SelectedIndex = 0
                    edt_llvt_chucvu.Text = ""
                    edt_llvt_donvi.Text = ""
                    edt_llvt_ghichu.Text = ""
                    'Select lại row hiện tại trước đó của người dùng
                    If (dgv_llvt.Rows.Count <= 0) Then Return
                    If (_currRow = "") Then _currRow = dgv_llvt.CurrentRow.Cells("cln_Id").Value.ToString()
                    dgv_llvt.CurrentRow.Selected = False
                    dgv_llvt.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_llvt, "cln_Id")).Selected = True
                    If (dgv_llvt.Rows.Count > 0) Then
                        If ((dgv_llvt.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_llvt.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                Case "tp_hs_cu"
                    'Reset Tab Hồ sơ cũ
                    dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
                    dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
                    edt_hscu_diachi.Text = ""
                    edt_hscu_nghenghiep.Text = ""
                    edt_hscu_ghichu.Text = ""
                    'Select lại row hiện tại trên lưới dữ liệu - trước đó của người dùng
                    If (dgv_hosocu.Rows.Count <= 0) Then Return
                    If (_currRow = "") Then _currRow = dgv_hosocu.CurrentRow.Cells("cln_Id").Value.ToString()
                    dgv_hosocu.CurrentRow.Selected = False
                    dgv_hosocu.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hosocu, "cln_Id")).Selected = True
                    If (dgv_hosocu.Rows.Count > 0) Then
                        If ((dgv_hosocu.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_hosocu.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                Case "tp_hh_hv"
                    _Code_HocHam = ""
                    _Code_HocVi = ""
                    flagAdd = False 'Lưu trạng thái cập nhật của hồ sơ Học hàm - Học vị
                    'Reset Tab - Học hàm học vị
                    cb_hhhv_hocham.SelectedIndex = 0
                    cb_hhhv_hocvi.SelectedIndex = 0
                    'Clear các CheckListBox
                    ckb_hocham.Checked = False
                    ckb_hocham_CheckedChanged(Nothing, Nothing)
                    ckb_hocvi.Checked = False
                    ckb_hocvi_CheckedChanged(Nothing, Nothing)
                    ckb_hocham.Enabled = False
                    ckb_hocvi.Enabled = False
                    Fill_Data(7, False)
                    Fill_Data(8, False)
            End Select
        End If
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        lbl_thongbao.Text = ""
        If (FlagHS = False) Then    'Cập nhật thông tin hồ sơ nhân sự
            Update_Personals()
        Else                        'Cập nhật thông tin hồ sơ Khác
            Dim _currRow As String = ""
            Select Case tctrl_main.SelectedTab.Name
                'Case "tp_hs_congtac"    'Cập nhật hồ sơ khác - Hồ sơ công tác
                '    'Kiểm tra Quyền thêm mới hoặc sửa đổi Hồ sơ công tác
                '    If (_IdEdit = "") Then
                '        If (Globals.Roles.IndexOf(";76;") < 0) Then
                '            MessageBox.Show("Bạn không có quyền thêm mới hồ sơ công tác của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                '            dgv_congtac_CellClick(sender, Nothing)
                '            Return
                '        End If
                '    Else
                '        If (Globals.Roles.IndexOf(";77;") < 0) Then
                '            MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ công tác của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                '            _currRow = _IdEdit
                '            'Reset null controls
                '            _IdEdit = ""
                '            dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
                '            dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
                '            dtpk_ct_denngay.Checked = False
                '            edt_ct_chucvu.Text = ""
                '            edt_ct_diachi.Text = ""
                '            edt_ct_lydo.Text = ""
                '            edt_ct_ghichu.Text = ""
                '            'Select row trên lưới dữ liệu - Fill data theo row đang select ra controls
                '            dgv_congtac.CurrentRow.Selected = False
                '            dgv_congtac.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_congtac, "cln_Id")).Selected = True
                '            If (dgv_congtac.Rows.Count > 0) Then
                '                If ((dgv_congtac.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_congtac.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                '                    SelectRow(_currRow)
                '                End If
                '            End If
                '            Return
                '        End If
                '    End If

                '    'Thực hiện cập nhật dữ liệu - Hồ sơ công tác của cán bộ
                '    If (Valid(1)) Then
                '        Dim obj_congtac As clsHS_CanBo.HS_CongTac = New clsHS_CanBo.HS_CongTac()
                '        obj_congtac.IdCanBo = _IdCanBo
                '        obj_congtac.TuNgay = dtpk_ct_tungay.Value
                '        obj_congtac.DenNgay = IIf(dtpk_ct_denngay.Checked, dtpk_ct_denngay.Value, DateTime.Parse("01/01/1900"))
                '        obj_congtac.ChucVu = Globals.Find_Replace(edt_ct_chucvu.Text.Trim.ToString())
                '        obj_congtac.DiaChi = Globals.Find_Replace(edt_ct_diachi.Text.Trim.ToString())
                '        obj_congtac.LyDo = Globals.Find_Replace(edt_ct_lydo.Text.Trim.ToString())
                '        obj_congtac.GhiChu = Globals.Find_Replace(edt_ct_ghichu.Text.Trim.ToString())
                '        If (_IdEdit = "") Then
                '            _currRow = _HS_CanBo.Insert_Employment_History(obj_congtac)
                '            lbl_thongbao.Text = " Bạn đã thêm mới thành công hồ sơ công tác!"
                '        Else
                '            obj_congtac.IdHSCongTac = _IdEdit
                '            _HS_CanBo.Update_Employment_History(obj_congtac)
                '            _currRow = _IdEdit
                '            lbl_thongbao.Text = " Bạn đã sửa đổi thành công hồ sơ công tác!"
                '        End If
                '        'Fill data - Hồ sơ công tác của cán bộ
                '        Fill_Data(2, False)
                '        'Reset null controls
                '        _IdEdit = ""
                '        dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
                '        dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
                '        edt_ct_chucvu.Text = ""
                '        edt_ct_diachi.Text = ""
                '        edt_ct_lydo.Text = ""
                '        edt_ct_ghichu.Text = ""
                '        'Tìm lại dòng đang select trước đó
                '        dgv_congtac.CurrentRow.Selected = False
                '        dgv_congtac.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_congtac, "cln_Id")).Selected = True
                '        'Select row trên lưới dữ liệu
                '        If (dgv_congtac.Rows.Count > 0) Then
                '            If ((dgv_congtac.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_congtac.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                '                SelectRow(_currRow)
                '            End If
                '        End If
                '    End If

                Case "tp_xuatngoai"
                    'Kiểm tra Quyền thêm mới hoặc sửa đổi Hồ sơ Xuất ngoại
                    If (_IdEdit = "") Then
                        If (Globals.Roles.IndexOf(";80;") < 0) Then
                            MessageBox.Show("Bạn không có quyền thêm mới hồ sơ xuất ngoại của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            dgv_xuatngoai_CellClick(sender, Nothing)
                            Return
                        End If
                    Else
                        If (Globals.Roles.IndexOf(";81;") < 0) Then
                            MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ xuất ngoại của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            _currRow = _IdEdit
                            'Reset null controls
                            _IdEdit = ""
                            dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
                            dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()
                            cb_xn_nuocden.SelectedIndex = 0
                            edt_xn_mucdich.Text = ""
                            edt_xn_soqd.Text = ""
                            dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
                            edt_xn_nguoiky.Text = ""
                            edt_xn_ghichu.Text = ""
                            cb_xn_chucvu.SelectedIndex = 0
                            'Select row trên lưới dữ liệu - Fill data theo row đang select ra controls
                            dgv_xuatngoai.CurrentRow.Selected = False
                            dgv_xuatngoai.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_xuatngoai, "cln_Id")).Selected = True
                            If (dgv_xuatngoai.Rows.Count > 0) Then
                                If ((dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                    SelectRow(_currRow)
                                End If
                            End If
                            Return
                        End If
                    End If

                    'Thực hiện cập nhật dữ liệu - Hồ sơ Xuất ngoại của cán bộ
                    If (Valid(2)) Then
                        Dim obj_xuatngoai As clsHS_CanBo.HS_XuatNgoai = New clsHS_CanBo.HS_XuatNgoai()
                        obj_xuatngoai.IdCanBo = _IdCanBo
                        obj_xuatngoai.TuNgay = dtpk_xn_tungay.Value
                        obj_xuatngoai.DenNgay = dtpk_xn_denngay.Value
                        obj_xuatngoai.IdNuocDen = CType(IIf(arr_Quocgiaden.Count > 0, arr_Quocgiaden(cb_xn_nuocden.SelectedIndex), "0"), Int32)
                        obj_xuatngoai.MucDich = Globals.Find_Replace(edt_xn_mucdich.Text.Trim.ToString())
                        obj_xuatngoai.SoQD = Globals.Find_Replace(edt_xn_soqd.Text.Trim.ToString())
                        obj_xuatngoai.NgayKy_QD = dtpk_xn_ngayky.Value
                        obj_xuatngoai.NguoiKy_QD = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_xn_nguoiky.Text.Trim.ToString()))
                        obj_xuatngoai.IdCV_NguoiKy_QD = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_xn_chucvu.SelectedIndex), "0"), Int32)
                        obj_xuatngoai.GhiChu = Globals.Find_Replace(edt_xn_ghichu.Text.Trim.ToString())
                        If (_IdEdit = "") Then
                            _currRow = _HS_CanBo.Insert_Abroad(obj_xuatngoai)
                            lbl_thongbao.Text = " Bạn đã thêm mới thành công hồ sơ xuất ngoại!"
                        Else
                            obj_xuatngoai.IdXuatNgoai = _IdEdit
                            _HS_CanBo.Update_Abroad(obj_xuatngoai)
                            _currRow = _IdEdit
                            lbl_thongbao.Text = " Bạn đã sửa đổi thành công hồ sơ xuất ngoại!"
                        End If
                        'Fill data - Hồ sơ công tác của cán bộ
                        Fill_Data(3, False)
                        _IdEdit = ""
                        dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
                        dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()
                        cb_xn_nuocden.SelectedIndex = 0
                        edt_xn_mucdich.Text = ""
                        edt_xn_soqd.Text = ""
                        dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
                        edt_xn_nguoiky.Text = ""
                        edt_xn_ghichu.Text = ""
                        cb_xn_chucvu.SelectedIndex = 0
                        'Tìm lại dòng đang select trước đó
                        dgv_xuatngoai.CurrentRow.Selected = False
                        dgv_xuatngoai.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_xuatngoai, "cln_Id")).Selected = True
                        'Select row trên lưới dữ liệu
                        If (dgv_xuatngoai.Rows.Count > 0) Then
                            If ((dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                    End If

                Case "tp_hochieu"
                    'Kiểm tra Quyền thêm mới hoặc sửa đổi Hồ sơ Hộ chiếu
                    If (_IdEdit = "") Then
                        If (Globals.Roles.IndexOf(";84;") < 0) Then
                            MessageBox.Show("Bạn không có quyền thêm mới hồ sơ hộ chiếu của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            dgv_hochieu_CellClick(sender, Nothing)
                            Return
                        End If
                    Else
                        If (Globals.Roles.IndexOf(";85;") < 0) Then
                            MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ hộ chiếu của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            _currRow = _IdEdit
                            'Reset null controls
                            _IdEdit = ""
                            edt_hc_sohc.Text = ""
                            cb_hc_loaihc.SelectedIndex = 0
                            dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
                            edt_hc_noicap.Text = ""
                            dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
                            cb_hc_tinhtrang.SelectedIndex = 0
                            edt_hc_ghichu.Text = ""
                            'Tìm lại dòng đang select trước đó
                            dgv_hochieu.CurrentRow.Selected = False
                            dgv_hochieu.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hochieu, "cln_Id")).Selected = True
                            If (dgv_hochieu.Rows.Count > 0) Then
                                If ((dgv_hochieu.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_hochieu.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                    SelectRow(_currRow)
                                End If
                            End If
                            Return
                        End If
                    End If

                    'Thực hiện cập nhật dữ liệu - Hồ sơ Hộ chiếu của cán bộ
                    If (Valid(3)) Then
                        Dim obj_hochieu As clsHS_CanBo.HS_HoChieu = New clsHS_CanBo.HS_HoChieu()
                        obj_hochieu.IdCanBo = _IdCanBo
                        obj_hochieu.So_HoChieu = Globals.Find_Replace(edt_hc_sohc.Text.Trim.ToString())
                        obj_hochieu.Loai_HC = CType(cb_hc_loaihc.SelectedIndex, Byte)
                        obj_hochieu.NgayCap = dtpk_hc_ngaycap.Value
                        obj_hochieu.NoiCap = Globals.Find_Replace(edt_hc_noicap.Text.Trim.ToString())
                        obj_hochieu.NgayHH = dtpk_hc_ngayhh.Value
                        obj_hochieu.TinhTrang = CType(cb_hc_tinhtrang.SelectedIndex, Byte)
                        obj_hochieu.GhiChu = Globals.Find_Replace(edt_hc_ghichu.Text.Trim.ToString())
                        If (_IdEdit = "") Then
                            _currRow = _HS_CanBo.Insert_Passport(obj_hochieu)
                            lbl_thongbao.Text = " Bạn đã thêm mới thành công hồ sơ hộ chiếu!"
                        Else
                            obj_hochieu.IdHoChieu = _IdEdit
                            _HS_CanBo.Update_Passport(obj_hochieu)
                            _currRow = _IdEdit
                            lbl_thongbao.Text = " Bạn đã sửa đổi thành công hồ sơ hộ chiếu!"
                        End If
                        'Fill data - Hồ sơ công tác của cán bộ
                        Fill_Data(4, False)
                        _IdEdit = ""
                        edt_hc_sohc.Text = ""
                        cb_hc_loaihc.SelectedIndex = 0
                        dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
                        edt_hc_noicap.Text = ""
                        dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
                        cb_hc_tinhtrang.SelectedIndex = 0
                        edt_hc_ghichu.Text = ""

                        'Tìm lại dòng đang select trước đó
                        dgv_hochieu.CurrentRow.Selected = False
                        dgv_hochieu.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hochieu, "cln_Id")).Selected = True
                        'Select row trên lưới dữ liệu
                        If (dgv_hochieu.Rows.Count > 0) Then
                            If ((dgv_hochieu.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_hochieu.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                    End If

                Case "tp_llvt"
                    'Kiểm tra Quyền thêm mới hoặc sửa đổi Hồ sơ Lực lượng vũ trang
                    If (_IdEdit = "") Then
                        If (Globals.Roles.IndexOf(";88;") < 0) Then
                            MessageBox.Show("Bạn không có quyền thêm mới hồ sơ tham gia lực lượng vũ trang của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            dgv_llvt_CellClick(sender, Nothing)
                            Return
                        End If
                    Else
                        If (Globals.Roles.IndexOf(";89;") < 0) Then
                            MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ tham gia lực lượng vũ trang của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            _currRow = _IdEdit
                            'Reset null controls
                            _IdEdit = ""
                            dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
                            dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()
                            cb_llvt_phanloai.SelectedIndex = 0
                            cb_llvt_quanham.SelectedIndex = 0
                            edt_llvt_chucvu.Text = ""
                            edt_llvt_donvi.Text = ""
                            edt_llvt_ghichu.Text = ""
                            'Tìm lại dòng đang select trước đó
                            dgv_llvt.CurrentRow.Selected = False
                            dgv_llvt.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_llvt, "cln_Id")).Selected = True
                            If (dgv_llvt.Rows.Count > 0) Then
                                If ((dgv_llvt.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_llvt.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                    SelectRow(_currRow)
                                End If
                            End If
                            Return
                        End If
                    End If

                    'Thực hiện cập nhật dữ liệu - Hồ sơ Tham gia Lực lượng vũ trang của cán bộ
                    If (Valid(4)) Then
                        Dim obj_llvt As clsHS_CanBo.HS_LLVT = New clsHS_CanBo.HS_LLVT()
                        obj_llvt.IdCanBo = _IdCanBo
                        obj_llvt.TuNgay = dtpk_llvt_tungay.Value
                        obj_llvt.DenNgay = dtpk_llvt_denngay.Value
                        obj_llvt.IdLoaiLLVT = CType(IIf(arr_PL_LLVT.Count > 0, arr_PL_LLVT(cb_llvt_phanloai.SelectedIndex), "0"), Int32)
                        obj_llvt.IdQuanHam = CType(IIf(arr_Quanham.Count > 0, arr_Quanham(cb_llvt_quanham.SelectedIndex), "0"), Int32)
                        obj_llvt.ChucVu = Globals.Find_Replace(edt_llvt_chucvu.Text.Trim.ToString())
                        obj_llvt.DonVi = Globals.Find_Replace(edt_llvt_donvi.Text.Trim.ToString())
                        obj_llvt.GhiChu = Globals.Find_Replace(edt_llvt_ghichu.Text.Trim.ToString())
                        If (_IdEdit = "") Then
                            _currRow = _HS_CanBo.Insert_ArmedForce(obj_llvt)
                            lbl_thongbao.Text = " Bạn đã thêm mới thành công hồ sơ lực lượng vũ trang!"
                        Else
                            obj_llvt.IdHSLLVT = _IdEdit
                            _HS_CanBo.Update_ArmedForce(obj_llvt)
                            _currRow = _IdEdit
                            lbl_thongbao.Text = " Bạn đã sửa đổi thành công hồ sơ lực lượng vũ trang!"
                        End If
                        'Fill data - Hồ sơ công tác của cán bộ
                        Fill_Data(5, False)
                        'Reset controls
                        _IdEdit = ""
                        dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
                        dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()
                        cb_llvt_phanloai.SelectedIndex = 0
                        cb_llvt_quanham.SelectedIndex = 0
                        edt_llvt_chucvu.Text = ""
                        edt_llvt_donvi.Text = ""
                        edt_llvt_ghichu.Text = ""
                        'Tìm lại dòng đang select trước đó
                        dgv_llvt.CurrentRow.Selected = False
                        dgv_llvt.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_llvt, "cln_Id")).Selected = True
                        'Select row trên lưới dữ liệu
                        If (dgv_llvt.Rows.Count > 0) Then
                            If ((dgv_llvt.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_llvt.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                    End If

                Case "tp_hs_cu"
                    'Kiểm tra Quyền thêm mới hoặc sửa đổi Hồ sơ Cũ của cán bộ
                    If (_IdEdit = "") Then
                        If (Globals.Roles.IndexOf(";92;") < 0) Then
                            MessageBox.Show("Bạn không có quyền thêm mới hồ sơ cũ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            dgv_hosocu_CellClick(sender, Nothing)
                            Return
                        End If
                    Else
                        If (Globals.Roles.IndexOf(";93;") < 0) Then
                            MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ cũ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                            _currRow = _IdEdit
                            'Reset null controls
                            _IdEdit = ""
                            dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
                            dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
                            edt_hscu_diachi.Text = ""
                            edt_hscu_nghenghiep.Text = ""
                            edt_hscu_ghichu.Text = ""
                            'Tìm lại dòng đang select trước đó
                            dgv_hosocu.CurrentRow.Selected = False
                            dgv_hosocu.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hosocu, "cln_Id")).Selected = True
                            If (dgv_hosocu.Rows.Count > 0) Then
                                If ((dgv_hosocu.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_hosocu.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                    SelectRow(_currRow)
                                End If
                            End If
                            Return
                        End If
                    End If

                    'Thực hiện cập nhật dữ liệu - Hồ sơ Cũ của cán bộ
                    If (Valid(5)) Then
                        Dim obj_hscu As clsHS_CanBo.HS_Cu = New clsHS_CanBo.HS_Cu()
                        obj_hscu.IdCanBo = _IdCanBo
                        obj_hscu.TuThang = dtpk_hscu_tungay.Value
                        obj_hscu.DenThang = dtpk_hscu_denngay.Value
                        obj_hscu.NgheNghiep = Globals.Find_Replace(edt_hscu_nghenghiep.Text.Trim.ToString())
                        obj_hscu.DiaChi = Globals.Find_Replace(edt_hscu_diachi.Text.Trim.ToString())
                        obj_hscu.GhiChu = Globals.Find_Replace(edt_hscu_ghichu.Text.Trim.ToString())
                        If (_IdEdit = "") Then
                            _currRow = _HS_CanBo.Insert_OldDocumnet(obj_hscu)
                            lbl_thongbao.Text = " Bạn đã thêm mới thành công hồ sơ cũ của cán bộ!"
                        Else
                            obj_hscu.IdHSCu = _IdEdit
                            _HS_CanBo.Update_OldDocumnet(obj_hscu)
                            _currRow = _IdEdit
                            lbl_thongbao.Text = " Bạn đã sửa đổi thành công hồ sơ cũ của cán bộ!"
                        End If
                        'Fill data - Hồ sơ công tác của cán bộ
                        Fill_Data(6, False)
                        'Thực hiện - Reset controls
                        _IdEdit = ""
                        dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
                        dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
                        edt_hscu_diachi.Text = ""
                        edt_hscu_nghenghiep.Text = ""
                        edt_hscu_ghichu.Text = ""

                        'Tìm lại dòng đang select trước đó
                        dgv_hosocu.CurrentRow.Selected = False
                        dgv_hosocu.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hosocu, "cln_Id")).Selected = True
                        'Select row trên lưới dữ liệu
                        If (dgv_hosocu.Rows.Count > 0) Then
                            If ((dgv_hosocu.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_hosocu.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                    End If

                Case "tp_hh_hv"
                    If (ckb_hocham.Checked = False And ckb_hocvi.Checked = False) Then
                        flagAdd = False
                        cb_hhhv_hocham.SelectedIndex = 0
                        cb_hhhv_hocvi.SelectedIndex = 0
                        ckb_hocham.Checked = False
                        ckb_hocham_CheckedChanged(sender, Nothing)
                        ckb_hocvi.Checked = False
                        ckb_hocvi_CheckedChanged(sender, Nothing)
                        ckb_hocham.Enabled = False
                        ckb_hocvi.Enabled = False
                        Return
                    End If

                    If (flagAdd = True) Then
                        If (cb_hhhv_hocham.Items.Count = 0 And cb_hhhv_hocvi.Items.Count = 0) Then Return
                        If (Valid(6)) Then
                            'Thực hiện thêm mới - Hồ sơ học hàm
                            If (ckb_hocham.Checked = True) Then
                                If (cb_hhhv_hocham.SelectedIndex > 0) Then
                                    _HS_CanBo.Insert_HocHam("", _IdCanBo, CType(IIf(arr_Hhhv_CHocham.Count > 0, arr_Hhhv_CHocham(cb_hhhv_hocham.SelectedIndex), "0"), Int32))
                                    flagAdd = False
                                    'Load lại danh sách học hàm của cán bộ
                                    Fill_Data(7, False)
                                    clb_hhhv_hocham_SelectedIndexChanged(sender, Nothing)
                                End If
                            End If

                            'Thực hiện thêm mới - Hồ sơ học vị
                            If (ckb_hocvi.Checked = True) Then
                                If (cb_hhhv_hocvi.SelectedIndex > 0) Then
                                    _HS_CanBo.Insert_HocVi("", _IdCanBo, CType(IIf(arr_Hhhv_CHocvi.Count > 0, arr_Hhhv_CHocvi(cb_hhhv_hocvi.SelectedIndex), "0"), Int32))
                                    'Load lại danh sách học hàm của cán bộ
                                    Fill_Data(8, False)
                                    clb_hhhv_hocvi_SelectedIndexChanged(sender, Nothing)
                                End If
                            End If
                            'Load lại các thông tin về trạng thái ban đầu
                            lbl_thongbao.Text = " Bạn đã thêm mới thành công thông tin hồ sơ học hàm - học vị!"
                            flagAdd = False
                            cb_hhhv_hocham.SelectedIndex = 0
                            cb_hhhv_hocvi.SelectedIndex = 0
                            ckb_hocham.Checked = False
                            ckb_hocham_CheckedChanged(sender, Nothing)
                            ckb_hocvi.Checked = False
                            ckb_hocvi_CheckedChanged(sender, Nothing)
                            ckb_hocham.Enabled = False
                            ckb_hocvi.Enabled = False
                        End If
                    Else
                        If (Globals.Roles.IndexOf(";97;") < 0) Then
                            lbl_thongbao.Text = " Bạn không có quyền sửa đổi thông tin học hàm - học vị của cán bộ!"

                            clb_hhhv_hocham_SelectedIndexChanged(sender, Nothing)
                            clb_hhhv_hocvi_SelectedIndexChanged(sender, Nothing)

                            cb_hhhv_hocham.SelectedIndex = 0
                            cb_hhhv_hocvi.SelectedIndex = 0
                            ckb_hocham.Checked = False
                            ckb_hocham_CheckedChanged(sender, Nothing)
                            ckb_hocvi.Checked = False
                            ckb_hocvi_CheckedChanged(sender, Nothing)
                            _Code_HocVi = ""
                            _Code_HocHam = ""
                            ckb_hocham.Enabled = False
                            ckb_hocvi.Enabled = False
                            Return
                        End If

                        If (_Code_HocHam = "" And _Code_HocVi = "") Then Return
                        If (Valid(6)) Then
                            Dim strSQL As String = ""
                            If (ckb_hocham.Checked = True) Then
                                If (cb_hhhv_hocham.SelectedIndex > 0) Then
                                    _HS_CanBo.Update_HocHam(_Code_HocHam, _IdCanBo, CType(IIf(arr_Hhhv_CHocham.Count > 0, arr_Hhhv_CHocham(cb_hhhv_hocham.SelectedIndex), "0"), Int32))
                                    'Load lại danh sách học hàm của cán bộ
                                    Fill_Data(7, False)
                                    clb_hhhv_hocham_SelectedIndexChanged(sender, Nothing)
                                End If
                            End If
                            If (ckb_hocvi.Checked = True) Then
                                If (cb_hhhv_hocvi.SelectedIndex > 0) Then
                                    _HS_CanBo.Update_HocVi(_Code_HocVi, _IdCanBo, CType(IIf(arr_Hhhv_CHocvi.Count > 0, arr_Hhhv_CHocvi(cb_hhhv_hocvi.SelectedIndex), "0"), Int32))
                                    'Load lại danh sách học hàm của cán bộ
                                    Fill_Data(8, False)
                                    clb_hhhv_hocvi_SelectedIndexChanged(sender, Nothing)
                                End If
                            End If
                            lbl_thongbao.Text = " Bạn đã sửa đổi thành công thông tin học hàm - học vị!"
                            cb_hhhv_hocham.SelectedIndex = 0
                            cb_hhhv_hocvi.SelectedIndex = 0
                            ckb_hocham.Checked = False
                            ckb_hocham_CheckedChanged(sender, Nothing)
                            ckb_hocvi.Checked = False
                            ckb_hocvi_CheckedChanged(sender, Nothing)
                            _Code_HocVi = ""
                            _Code_HocHam = ""
                            ckb_hocham.Enabled = False
                            ckb_hocvi.Enabled = False
                        End If
                    End If
            End Select
        End If
    End Sub
#End Region

#Region "---> Events: Cell click - Sự kiện click vào lưới dữ liệu <---"
    'Private Sub dgv_congtac_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_congtac.CellClick
    '    'Reset null controls
    '    _IdEdit = ""
    '    dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
    '    dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
    '    dtpk_ct_denngay.Checked = False
    '    edt_ct_chucvu.Text = ""
    '    edt_ct_diachi.Text = ""
    '    edt_ct_lydo.Text = ""
    '    edt_ct_ghichu.Text = ""
    '    'Fill data in controls
    '    If (dgv_congtac.Rows.Count > 0) Then
    '        If ((dgv_congtac.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_congtac.CurrentRow.Cells(0).Value.ToString() <> "")) Then
    '            'Lấy Id công tác tại bản ghi đang Select
    '            SelectRow(dgv_congtac.CurrentRow.Cells("cln_Id").Value.ToString())
    '        End If
    '    End If
    'End Sub

    'Private Sub dgv_congtac_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_congtac.KeyUp
    '    dgv_congtac_CellClick(sender, Nothing)
    'End Sub

    Private Sub dgv_xuatngoai_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_xuatngoai.CellClick
        _IdEdit = ""
        dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
        dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()
        cb_xn_nuocden.SelectedIndex = 0
        edt_xn_mucdich.Text = ""
        edt_xn_soqd.Text = ""
        dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
        edt_xn_nguoiky.Text = ""
        edt_xn_ghichu.Text = ""
        cb_xn_chucvu.SelectedIndex = 0
        If (dgv_xuatngoai.Rows.Count > 0) Then
            If ((dgv_xuatngoai.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_xuatngoai.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                'Lấy Id công tác tại bản ghi đang Select
                SelectRow(dgv_xuatngoai.CurrentRow.Cells("cln_Id").Value.ToString())
            End If
        End If
    End Sub

    Private Sub dgv_xuatngoai_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_xuatngoai.KeyUp
        dgv_xuatngoai_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_hochieu_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hochieu.CellClick
        _IdEdit = ""
        edt_hc_sohc.Text = ""
        cb_hc_loaihc.SelectedIndex = 0
        dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
        edt_hc_noicap.Text = ""
        dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
        cb_hc_tinhtrang.SelectedIndex = 0
        edt_hc_ghichu.Text = ""
        If (dgv_hochieu.Rows.Count > 0) Then
            If ((dgv_hochieu.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_hochieu.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                'Lấy Id công tác tại bản ghi đang Select
                SelectRow(dgv_hochieu.CurrentRow.Cells("cln_Id").Value.ToString())
            End If
        End If
    End Sub

    Private Sub dgv_hochieu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hochieu.KeyUp
        dgv_hochieu_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_llvt_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_llvt.CellClick
        _IdEdit = ""
        dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
        dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()
        cb_llvt_phanloai.SelectedIndex = 0
        cb_llvt_quanham.SelectedIndex = 0
        edt_llvt_chucvu.Text = ""
        edt_llvt_donvi.Text = ""
        edt_llvt_ghichu.Text = ""
        If (dgv_llvt.Rows.Count > 0) Then
            If ((dgv_llvt.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_llvt.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                'Lấy Id công tác tại bản ghi đang Select
                SelectRow(dgv_llvt.CurrentRow.Cells("cln_Id").Value.ToString())
            End If
        End If
    End Sub

    Private Sub dgv_llvt_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_llvt.KeyUp
        dgv_llvt_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_hosocu_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hosocu.CellClick
        _IdEdit = ""
        dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
        dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
        edt_hscu_diachi.Text = ""
        edt_hscu_nghenghiep.Text = ""
        edt_hscu_ghichu.Text = ""
        If (dgv_hosocu.Rows.Count > 0) Then
            If ((dgv_hosocu.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_hosocu.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                SelectRow(dgv_hosocu.CurrentRow.Cells("cln_Id").Value.ToString())
            End If
        End If
    End Sub

    Private Sub dgv_hosocu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hosocu.KeyUp
        dgv_hosocu_CellClick(sender, Nothing)
    End Sub

    Private Sub clb_hhhv_hocham_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clb_hhhv_hocham.SelectedIndexChanged
        flagAdd = False
        ckb_hocham.Enabled = False
        _Code_HocHam = ""
        cb_hhhv_hocham.SelectedIndex = 0
        If (clb_hhhv_hocham.Items.Count <> 0 And clb_hhhv_hocham.SelectedIndex >= 0) Then
            'Lấy id của item select
            ckb_hocham.Enabled = True
            _Code_HocHam = arr_Hhhv_LHocham(clb_hhhv_hocham.SelectedIndex).ToString()
            Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, arr_Hhhv_LHocham(clb_hhhv_hocham.SelectedIndex), 6)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        cb_hhhv_hocham.SelectedIndex = CType(arr_Hhhv_CHocham.IndexOf(db.Rows(0)("IdHocHam").ToString()), Int32)
                    End If
                End If
            End Using
        End If
    End Sub

    Private Sub clb_hhhv_hocvi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clb_hhhv_hocvi.SelectedIndexChanged
        flagAdd = False
        ckb_hocvi.Enabled = False
        _Code_HocVi = ""
        cb_hhhv_hocvi.SelectedIndex = 0
        If (clb_hhhv_hocvi.Items.Count <> 0 And clb_hhhv_hocvi.SelectedIndex >= 0) Then
            'Lấy id của item select
            _Code_HocVi = arr_Hhhv_LHocvi(clb_hhhv_hocvi.SelectedIndex).ToString()
            ckb_hocvi.Enabled = True
            Using db As DataTable = _HS_CanBo.GetDocument(_IdCanBo, arr_Hhhv_LHocvi(clb_hhhv_hocvi.SelectedIndex), 7)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        cb_hhhv_hocvi.SelectedIndex = CType(arr_Hhhv_CHocvi.IndexOf(db.Rows(0)("IdHocVi").ToString()), Int32)
                    End If
                End If
            End Using
        End If
    End Sub

    Private Sub tctrl_main_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tctrl_main.SelectedIndexChanged
        Dim _roles As String = Globals.Roles
        'Thiết lập cho phép các chức năng hiển thị
        btn_add.Visible = True
        btn_delete.Visible = True
        btn_huybo.Visible = True
        btn_save.Visible = True
        btn_add.Enabled = True
        btn_delete.Enabled = True
        btn_huybo.Enabled = True
        btn_save.Enabled = True
        lbl_thongbao.Text = ""
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_congtac"
                'dgv_congtac_CellClick(sender, Nothing)
                ''Xét quyền - Hồ sơ công tác
                'If (_roles.IndexOf(";75;") < 0) Then
                '    dgv_congtac.Enabled = False
                'Else
                '    dgv_congtac.Enabled = True
                'End If
                'If (_roles.IndexOf(";76;") < 0) Then
                '    btn_add.Enabled = False
                'End If
                'If (_roles.IndexOf(";78;") < 0) Then
                '    btn_delete.Enabled = False
                'End If
                btn_add.Visible = False
                btn_delete.Visible = False
                btn_huybo.Visible = False
                btn_save.Visible = False
            Case "tp_xuatngoai"
                dgv_xuatngoai_CellClick(sender, Nothing)
                If (_roles.IndexOf(";79;") < 0) Then
                    dgv_xuatngoai.Enabled = False
                Else
                    dgv_xuatngoai.Enabled = True
                End If
                If (_roles.IndexOf(";80;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";82;") < 0) Then
                    btn_delete.Enabled = False
                End If

            Case "tp_hochieu"
                dgv_hochieu_CellClick(sender, Nothing)
                If (_roles.IndexOf(";83;") < 0) Then
                    dgv_hochieu.Enabled = False
                Else
                    dgv_hochieu.Enabled = True
                End If
                If (_roles.IndexOf(";84;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";86;") < 0) Then
                    btn_delete.Enabled = False
                End If

            Case "tp_llvt"
                dgv_llvt_CellClick(sender, Nothing)
                If (_roles.IndexOf(";87;") < 0) Then
                    dgv_llvt.Enabled = False
                Else
                    dgv_llvt.Enabled = True
                End If
                If (_roles.IndexOf(";88;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";90;") < 0) Then
                    btn_delete.Enabled = False
                End If

            Case "tp_hs_cu"
                dgv_hosocu_CellClick(sender, Nothing)
                If (_roles.IndexOf(";91;") < 0) Then
                    dgv_hosocu.Enabled = False
                Else
                    dgv_hosocu.Enabled = True
                End If
                If (_roles.IndexOf(";92;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";94;") < 0) Then
                    btn_delete.Enabled = False
                End If

            Case "tp_hh_hv"
                _Code_HocHam = ""
                _Code_HocVi = ""
                clb_hhhv_hocham_SelectedIndexChanged(sender, Nothing)
                clb_hhhv_hocvi_SelectedIndexChanged(sender, Nothing)
                ckb_hocvi.Checked = False
                ckb_hocham.Checked = False
                ckb_hocham.Enabled = False
                ckb_hocvi.Enabled = False
                If (clb_hhhv_hocham.SelectedIndex >= 0) Then ckb_hocham.Enabled = True
                If (clb_hhhv_hocvi.SelectedIndex >= 0) Then ckb_hocvi.Enabled = True

                If (_roles.IndexOf(";95;") < 0) Then
                    clb_hhhv_hocham.Enabled = False
                    clb_hhhv_hocvi.Enabled = False
                Else
                    clb_hhhv_hocham.Enabled = True
                    clb_hhhv_hocvi.Enabled = True
                End If
                If (_roles.IndexOf(";96;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";98;") < 0) Then
                    btn_delete.Enabled = False
                End If
        End Select
    End Sub
#End Region

#Region "---> Các sự kiện select index changed <---"
    Private Sub cb_ns_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ns_tinh.SelectedIndexChanged
        'cb_ns_huyen.Items.Clear()
        'arr_Ns_Huyen.Clear()
        'If (cb_ns_tinh.SelectedIndex > 0 And cb_ns_tinh.Items.Count <> 0) Then
        '    Dim _IdTinhTP As Int32 = CType(arr_Ns_Tinh(IIf(cb_ns_tinh.SelectedIndex > 0, cb_ns_tinh.SelectedIndex, "0")), Int32)
        '    If _IdTinhTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi From DiaDanh Where id_goc != 0 And id_goc = {0}", _IdTinhTP)
        '        arr_Ns_Huyen = _Globals.Bind_ComBoBox(cb_ns_huyen, strSQL, "---Quận - huyện---")
        '        If (FlagEvent = 1 And _IdCB_TS = "" And cb_nq_tinh.Items.Count <> 0) Then
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
                Dim strSQL As String = String.Format("Select Id,Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N'') As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N'') ,Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
                arr_Ns_Xa = _Globals.Bind_ComBoBox(cb_ns_xa, strSQL, "--- Chọn Xã/Phường ---")
                If (FlagEvent = 1 And _IdCanBo = "" And cb_nq_tinh.Items.Count <> 0) Then
                    cb_nq_tinh.SelectedIndex = cb_ns_tinh.SelectedIndex
                    cb_nq_tinh_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_ns_huyen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ns_huyen.SelectedIndexChanged
        If (cb_nq_huyen.Items.Count <> 0) Then
            If (FlagEvent = 1 And _IdCanBo = "") Then
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
                If (FlagEvent = 1 And _IdCanBo = "" And cb_nq_xa.Items.Count <> 0) Then
                    cb_nq_xa.SelectedIndex = cb_ns_xa.SelectedIndex
                    cb_nq_xa_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_ns_thon_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ns_thon.SelectedIndexChanged
        If (cb_nq_thon.Items.Count <> 0) Then
            If (FlagEvent = 1 And _IdCanBo = "") Then
                cb_nq_thon.SelectedIndex = cb_ns_thon.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_nq_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_nq_tinh.SelectedIndexChanged
        'cb_nq_huyen.Items.Clear()
        'arr_Nq_Huyen.Clear()
        'If (cb_nq_tinh.SelectedIndex > 0 And cb_nq_tinh.Items.Count <> 0) Then
        '    Dim _IdTinhTP As Int32 = CType(arr_Nq_Tinh(IIf(cb_nq_tinh.SelectedIndex > 0, cb_nq_tinh.SelectedIndex, "0")), Int32)
        '    If _IdTinhTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi From DiaDanh Where id_goc != 0 and id_goc = {0}", _IdTinhTP)
        '        arr_Nq_Huyen = _Globals.Bind_ComBoBox(cb_nq_huyen, strSQL, "---Quận - huyện---")
        '    End If
        'End If
        cb_nq_xa.Items.Clear()
        arr_Nq_Xa.Clear()
        If (cb_nq_tinh.SelectedIndex > 0 And cb_nq_tinh.Items.Count <> 0) Then
            Dim _IdTinhTP As Int32 = CType(arr_Nq_Tinh(IIf(cb_nq_tinh.SelectedIndex > 0, cb_nq_tinh.SelectedIndex, "0")), Int32)
            If _IdTinhTP > 0 Then
                Dim strSQL As String = String.Format("Select Id,Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N'') As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N''),Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
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
                arr_nq_Thon = _Globals.Bind_ComBoBox(cb_nq_thon, strSQL, "--- Chọn Thôn/Xóm ---")
                If (FlagEvent = 1 And _IdCanBo = "" And cb_nq_xa.Items.Count <> 0 And cb_tt_xa.Items.Count <> 0) Then
                    cb_tt_xa.SelectedIndex = cb_nq_xa.SelectedIndex
                    cb_tt_xa_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_nq_thon_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_nq_thon.SelectedIndexChanged
        If (cb_nq_thon.Items.Count <> 0 And cb_tt_thon.Items.Count <> 0) Then
            If (FlagEvent = 1 And _IdCanBo = "") Then
                cb_tt_thon.SelectedIndex = cb_nq_thon.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_tt_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tt_tinh.SelectedIndexChanged
        'cb_tt_huyen.Items.Clear()
        'arr_TT_Huyen.Clear()
        'If (cb_tt_tinh.SelectedIndex > 0 And cb_tt_tinh.Items.Count <> 0) Then
        '    Dim _IdTinhTP As Int32 = CType(arr_TT_Tinh(IIf(cb_tt_tinh.SelectedIndex > 0, cb_tt_tinh.SelectedIndex, "0")), Int32)
        '    If _IdTinhTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi from DiaDanh where id_goc != 0 and id_goc = {0} and Status = 1", _IdTinhTP)
        '        arr_TT_Huyen = _Globals.Bind_ComBoBox(cb_tt_huyen, strSQL, "---Quận - huyện---")
        '        If (FlagEvent = 1 And _IdCB_TS = "" And cb_ttr_tinh.Items.Count <> 0) Then
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
                Dim strSQL As String = String.Format("Select id,ten_goi from DiaDanh where id_goc != 0 and id_goc = {0} and Status = 1", _IdTinhTP)
                strSQL = String.Format("Select Id,Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N'') As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N''),Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)

                arr_TT_Xa = _Globals.Bind_ComBoBox(cb_tt_xa, strSQL, "--- Chọn Xã/Phường ---")
                If (FlagEvent = 1 And _IdCanBo = "" And cb_ttr_tinh.Items.Count <> 0) Then
                    cb_ttr_tinh.SelectedIndex = cb_tt_tinh.SelectedIndex
                    cb_ttr_tinh_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_tt_huyen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_tt_huyen.SelectedIndexChanged
        If (cb_ttr_huyen.Items.Count <> 0) Then
            If (FlagEvent = 1 And _IdCanBo = "") Then
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
                If (FlagEvent = 1 And _IdCanBo = "" And cb_tt_xa.Items.Count <> 0 And cb_ttr_xa.Items.Count <> 0) Then
                    cb_ttr_xa.SelectedIndex = cb_tt_xa.SelectedIndex
                    cb_ttr_xa_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub cb_tt_thon_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_tt_thon.SelectedIndexChanged
        If (cb_tt_thon.Items.Count <> 0 And cb_ttr_thon.Items.Count <> 0) Then
            If (FlagEvent = 1 And _IdCanBo = "") Then
                cb_ttr_thon.SelectedIndex = cb_tt_thon.SelectedIndex
            End If
        End If
    End Sub

    Private Sub cb_ttr_tinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ttr_tinh.SelectedIndexChanged
        'cb_ttr_huyen.Items.Clear()
        'arr_TTr_Huyen.Clear()
        'If (cb_ttr_tinh.SelectedIndex > 0 And cb_ttr_tinh.Items.Count <> 0) Then
        '    Dim _IdTinhTP As Int32 = CType(arr_TTr_Tinh(IIf(cb_ttr_tinh.SelectedIndex > 0, cb_ttr_tinh.SelectedIndex, "0")), Int32)
        '    If _IdTinhTP > 0 Then
        '        Dim strSQL As String = String.Format("Select id,ten_goi From DiaDanh Where id_goc != 0 And id_goc = {0}", _IdTinhTP)
        '        arr_TTr_Huyen = _Globals.Bind_ComBoBox(cb_ttr_huyen, strSQL, "---Quận - huyện---")
        '    End If
        'End If

        cb_ttr_xa.Items.Clear()
        arr_TTr_Xa.Clear()
        If (cb_ttr_tinh.SelectedIndex > 0 And cb_ttr_tinh.Items.Count <> 0) Then
            Dim _IdTinhTP As Int32 = CType(arr_TTr_Tinh(IIf(cb_ttr_tinh.SelectedIndex > 0, cb_ttr_tinh.SelectedIndex, "0")), Int32)
            If _IdTinhTP > 0 Then
                'Dim strSQL As String = String.Format("Select id,ten_goi From DiaDanh Where id_goc != 0 And id_goc = {0}", _IdTinhTP)
                Dim strSQL As String = String.Format("Select Id,Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N'')  As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Replace(Replace(Ten_Thon,N'xã ',N''),N'phường ',N''),Ma_Tinh,Ma_Xa,Ma_Thon Asc", _IdTinhTP)
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

    Private Sub cb_qd_donvi_moi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_qd_donvi_moi.SelectedIndexChanged
        Try
            arrQd_Phongban_moi.Clear()
            cb_qd_phongban_moi.Items.Clear()
            If (cb_qd_donvi_moi.SelectedIndex > 0 And cb_qd_donvi_moi.Items.Count <> 0) Then
                Dim _DonviId As Int32 = CType(arrQd_Donvi_moi(IIf(cb_qd_donvi_moi.SelectedIndex > 0, cb_qd_donvi_moi.SelectedIndex, "0")), Int32)
                'Lấy mã chi nhánh từ Id chi nhánh
                Dim strSQL As String = ""
                strSQL = String.Format("Select * from ChiNhanh Where id = {0} and Status = 1", _DonviId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim tructhuocCurrent As String = _HS_CanBo.GetTrucThuoc(db.Rows(0)("ma_so").ToString().Trim(), _DonviId)
                            strSQL = "Select id,ten_phong as ten_goi From PhongBan Where Charindex('" & tructhuocCurrent.ToString & "',truc_thuoc) > 0 and Status = 1 Order by Ma_so"
                            arrQd_Phongban_moi = _Globals.Bind_ComBoBox(cb_qd_phongban_moi, strSQL, "---Phòng ban---")
                        End If
                    End If
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show("Điền dữ liệu danh sách Phòng ban: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub ckb_hocham_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_hocham.CheckedChanged
        If (ckb_hocham.Checked = True) Then
            cb_hhhv_hocham.Enabled = True
        Else
            cb_hhhv_hocham.Enabled = False
        End If
    End Sub

    Private Sub ckb_hocvi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_hocvi.CheckedChanged
        If (ckb_hocvi.Checked = True) Then
            cb_hhhv_hocvi.Enabled = True
        Else
            cb_hhhv_hocvi.Enabled = False
        End If
    End Sub

    Private Sub cb_donvi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_donvi.SelectedIndexChanged
        If (FlagEvent = 1 And arr_Donvi.Count > 0) Then
            edt_macb.Text = setMaCanBo(arr_Donvi(cb_donvi.SelectedIndex))
        End If
        arrQd_Donvi_moi.Clear()
        cb_qd_donvi_moi.Items.Clear()
        If (cb_donvi.SelectedIndex > 0 And cb_donvi.Items.Count <> 0) Then
            Dim _IdIndex As Int32 = CType(arr_Donvi(IIf(cb_donvi.SelectedIndex > 0, cb_donvi.SelectedIndex, "0")), Integer)
            Dim strSQL As String = String.Format("Select * from ChiNhanh Where Status = 1 and id = {0}", _IdIndex)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        If (db.Rows(0)("ma_so").ToString() <> "") Then
                            strSQL = ""
                            If (CType(db.Rows(0)("ma_so").ToString().Substring(0, 4), Integer) <= 1) Then
                                strSQL = String.Format("Select id, ten_goi from ChiNhanh Where Status = 1 And id = {0}", _IdIndex)
                            Else
                                strSQL = String.Format("Select id, ten_goi from ChiNhanh Where (Status = 1) And (id = {0} Or id_goc = {0})", _IdIndex)
                            End If
                            arrQd_Donvi_moi = _Globals.Bind_ComBoBox(cb_qd_donvi_moi, strSQL, "---Đơn vị mới---")
                            If (FlagEvent = 1 And _IdCB_TS = "" And cb_qd_donvi_moi.Items.Count <> 0) Then
                                cb_qd_donvi_moi.SelectedIndex = IIf(cb_qd_donvi_moi.FindString(cb_donvi.SelectedItem.ToString()) > 0, cb_qd_donvi_moi.FindString(cb_donvi.SelectedItem.ToString()), 0)
                                cb_qd_donvi_moi_SelectedIndexChanged(sender, e)
                            End If
                        End If
                    End If
                End If
            End Using
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện liên quan đến Hồ sơ khác <---"
    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        lbl_thongbao.Text = ""
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_congtac"
                _IdEdit = ""
                dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
                edt_ct_chucvu.Text = ""
                edt_ct_diachi.Text = ""
                edt_ct_lydo.Text = ""
                edt_ct_ghichu.Text = ""
                ActiveControl = dtpk_ct_tungay
            Case "tp_xuatngoai"
                _IdEdit = ""
                dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()
                cb_xn_nuocden.SelectedIndex = 0
                edt_xn_mucdich.Text = ""
                edt_xn_soqd.Text = ""
                dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
                edt_xn_nguoiky.Text = ""
                edt_xn_ghichu.Text = ""
                cb_xn_chucvu.SelectedIndex = 0
                ActiveControl = dtpk_xn_tungay
            Case "tp_hochieu"
                _IdEdit = ""
                edt_hc_sohc.Text = ""
                cb_hc_loaihc.SelectedIndex = 0
                dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
                edt_hc_noicap.Text = ""
                dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
                cb_hc_tinhtrang.SelectedIndex = 0
                edt_hc_ghichu.Text = ""
                ActiveControl = edt_hc_sohc
            Case "tp_llvt"
                _IdEdit = ""
                dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()
                cb_llvt_phanloai.SelectedIndex = 0
                cb_llvt_quanham.SelectedIndex = 0
                edt_llvt_chucvu.Text = ""
                edt_llvt_donvi.Text = ""
                edt_llvt_ghichu.Text = ""
                ActiveControl = dtpk_llvt_tungay
            Case "tp_hs_cu"
                _IdEdit = ""
                dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
                dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
                edt_hscu_diachi.Text = ""
                edt_hscu_nghenghiep.Text = ""
                edt_hscu_ghichu.Text = ""
                ActiveControl = dtpk_hscu_tungay
            Case "tp_hh_hv"
                If (cb_hhhv_hocham.Items.Count = 0 And cb_hhhv_hocvi.Items.Count = 0) Then Return
                _Code_HocHam = ""
                _Code_HocVi = ""
                flagAdd = True
                cb_hhhv_hocham.SelectedIndex = 0
                cb_hhhv_hocvi.SelectedIndex = 0
                ckb_hocham.Enabled = True
                ckb_hocvi.Enabled = True
                ActiveControl = cb_hhhv_hocham
        End Select
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        lbl_thongbao.Text = ""
        Select Case tctrl_main.SelectedTab.Name
            'Case "tp_hs_congtac"
            '    If (dgv_congtac.Rows.Count <= 0) Then Return
            '    Try
            '        Dim arr_Del As ArrayList = New ArrayList()
            '        Dim _count As Int16 = 0
            '        If (dgv_congtac.Rows.Count > 0) Then
            '            If ((dgv_congtac.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_congtac.CurrentRow.Cells(0).Value.ToString() <> "")) Then
            '                For i As Int32 = 0 To dgv_congtac.Rows.Count - 1
            '                    If (dgv_congtac.Rows(i).Cells(0).Value IsNot Nothing) Then
            '                        If (CType(dgv_congtac.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
            '                            _count += 1
            '                            arr_Del.Add(dgv_congtac.Rows(i).Cells("cln_Id").Value.ToString())
            '                        End If
            '                    End If
            '                Next
            '            End If
            '        End If
            '        If (_count > 0) Then
            '            If (MessageBox.Show("Bạn có thực sự muốn xoá các bản ghi hồ sơ công tác đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
            '                'Thực hiện xoá dữ liệu khi đã chấp thuận
            '                For i As Int16 = 0 To arr_Del.Count - 1
            '                    _HS_CanBo.Delete_Employment_History(arr_Del(i).ToString())
            '                Next
            '                'Load lại các thông tin cần thiết sau khi xoá dữ liệu
            '                lbl_thongbao.Text = " Bạn đã xoá thành công hồ sơ công tác của cán bộ!"
            '                dtpk_ct_tungay.Text = DateTime.Now.ToShortDateString()
            '                dtpk_ct_denngay.Text = DateTime.Now.ToShortDateString()
            '                edt_ct_chucvu.Text = ""
            '                edt_ct_diachi.Text = ""
            '                edt_ct_lydo.Text = ""
            '                edt_ct_ghichu.Text = ""
            '                'Fill data - Hồ sơ công tác của cán bộ
            '                Fill_Data(2, False)
            '                dgv_congtac_CellClick(sender, Nothing)
            '            Else
            '                arr_Del.Clear()
            '                Globals.Check_All_Items(dgv_congtac, False)
            '            End If
            '        Else
            '            lbl_thongbao.Text = " Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!"
            '        End If
            '    Catch ex As Exception
            '        MessageBox.Show("Cập nhật xoá bỏ hồ sơ công tác của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            '    End Try
            Case "tp_xuatngoai"
                If (dgv_xuatngoai.Rows.Count <= 0) Then Return
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_xuatngoai.Rows.Count > 0) Then
                        If ((dgv_xuatngoai.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_xuatngoai.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                            For i As Int32 = 0 To dgv_xuatngoai.Rows.Count - 1
                                If (dgv_xuatngoai.Rows(i).Cells(0).Value IsNot Nothing) Then
                                    If (CType(dgv_xuatngoai.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_xuatngoai.Rows(i).Cells("cln_Id").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        If (MessageBox.Show("Bạn có thực sự muốn xoá các bản ghi hồ sơ xuất ngoại đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CanBo.Delete_Abroad(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            lbl_thongbao.Text = " Bạn đã xoá thành công hồ sơ xuất ngoại của cán bộ!"
                            dtpk_xn_tungay.Text = DateTime.Now.ToShortDateString()
                            dtpk_xn_denngay.Text = DateTime.Now.ToShortDateString()
                            cb_xn_nuocden.SelectedIndex = 0
                            edt_xn_mucdich.Text = ""
                            edt_xn_soqd.Text = ""
                            dtpk_xn_ngayky.Text = DateTime.Now.ToShortDateString()
                            edt_xn_nguoiky.Text = ""
                            edt_xn_ghichu.Text = ""
                            cb_xn_chucvu.SelectedIndex = 0
                            'Fill data - Hồ sơ công tác của cán bộ
                            Fill_Data(3, False)
                            dgv_xuatngoai_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_xuatngoai, False)
                        End If
                    Else
                        lbl_thongbao.Text = " Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!"
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ xuất ngoại của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
            Case "tp_hochieu"
                If (dgv_hochieu.Rows.Count <= 0) Then Return
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_hochieu.Rows.Count > 0) Then
                        If ((dgv_hochieu.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_hochieu.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                            For i As Int32 = 0 To dgv_hochieu.Rows.Count - 1
                                If (dgv_hochieu.Rows(i).Cells(0).Value IsNot Nothing) Then
                                    If (CType(dgv_hochieu.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_hochieu.Rows(i).Cells("cln_Id").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        If (MessageBox.Show("Bạn có thực sự muốn xoá các bản ghi hồ sơ hộ chiếu đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CanBo.Delete_Passport(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            lbl_thongbao.Text = " Bạn đã xoá thành công hồ sơ hộ chiếu của cán bộ!"

                            edt_hc_sohc.Text = ""
                            cb_hc_loaihc.SelectedIndex = 0
                            dtpk_hc_ngaycap.Text = DateTime.Now.ToShortDateString()
                            edt_hc_noicap.Text = ""
                            dtpk_hc_ngayhh.Text = DateTime.Now.ToShortDateString()
                            cb_hc_tinhtrang.SelectedIndex = 0
                            edt_hc_ghichu.Text = ""
                            'Fill data - Hồ sơ công tác của cán bộ
                            Fill_Data(4, False)
                            dgv_hochieu_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_hochieu, False)
                        End If
                    Else
                        lbl_thongbao.Text = " Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!"
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ hộ chiếu của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
            Case "tp_llvt"
                If (dgv_llvt.Rows.Count <= 0) Then Return
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_llvt.Rows.Count > 0) Then
                        If ((dgv_llvt.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_llvt.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                            For i As Int32 = 0 To dgv_llvt.Rows.Count - 1
                                If (dgv_llvt.Rows(i).Cells(0).Value IsNot Nothing) Then
                                    If (CType(dgv_llvt.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_llvt.Rows(i).Cells("cln_Id").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        If (MessageBox.Show("Bạn có thực sự muốn xoá các bản ghi hồ sơ tham gia lực lượng vũ trang đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CanBo.Delete_ArmedForce(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            lbl_thongbao.Text = " Bạn đã xoá thành công hồ sơ tham gia lực lượng vũ trang của cán bộ!"
                            dtpk_llvt_tungay.Text = DateTime.Now.ToShortDateString()
                            dtpk_llvt_denngay.Text = DateTime.Now.ToShortDateString()
                            cb_llvt_phanloai.SelectedIndex = 0
                            cb_llvt_quanham.SelectedIndex = 0
                            edt_llvt_chucvu.Text = ""
                            edt_llvt_donvi.Text = ""
                            edt_llvt_ghichu.Text = ""
                            'Fill data - Hồ sơ công tác của cán bộ
                            Fill_Data(5, False)
                            dgv_llvt_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_llvt, False)
                        End If
                    Else
                        lbl_thongbao.Text = " Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!"
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ lực lượng vũ trang của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
                'Case "tp_hs_cu"
                '    If (dgv_hosocu.Rows.Count <= 0) Then Return
                '    Try
                '        Dim arr_Del As ArrayList = New ArrayList()
                '        Dim _count As Int16 = 0
                '        If (dgv_hosocu.Rows.Count > 0) Then
                '            If ((dgv_hosocu.CurrentRow.Cells(0).Value IsNot Nothing) And (dgv_hosocu.CurrentRow.Cells(0).Value.ToString() <> "")) Then
                '                For i As Int32 = 0 To dgv_hosocu.Rows.Count - 1
                '                    If (dgv_hosocu.Rows(i).Cells(0).Value IsNot Nothing) Then
                '                        If (CType(dgv_hosocu.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                '                            _count += 1
                '                            arr_Del.Add(dgv_hosocu.Rows(i).Cells("cln_Id").Value.ToString())
                '                        End If
                '                    End If
                '                Next
                '            End If
                '        End If
                '        If (_count > 0) Then
                '            If (MessageBox.Show("Bạn có thực sự muốn xoá các bản ghi hồ sơ cũ đã chọn không?" + vbCrLf + "Chú ý: Hồ sơ cũ là hồ sơ lưu lại các thông tin của cán bộ trước khi đi làm", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                '                'Thực hiện xoá dữ liệu khi đã chấp thuận
                '                For i As Int16 = 0 To arr_Del.Count - 1
                '                    _HS_CanBo.Delete_OldDocumnet(arr_Del(i).ToString())
                '                Next
                '                'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                '                lbl_thongbao.Text = " Bạn đã xoá thành công hồ sơ cũ của cán bộ!"
                '                dtpk_hscu_denngay.Text = DateTime.Now.ToShortDateString()
                '                dtpk_hscu_tungay.Text = DateTime.Now.ToShortDateString()
                '                edt_hscu_diachi.Text = ""
                '                edt_hscu_nghenghiep.Text = ""
                '                edt_hscu_ghichu.Text = ""
                '                'Fill data - Hồ sơ công tác của cán bộ
                '                Fill_Data(6, False)
                '                dgv_hosocu_CellClick(sender, Nothing)
                '            Else
                '                arr_Del.Clear()
                '                Globals.Check_All_Items(dgv_hosocu, False)
                '            End If
                '        Else
                '            lbl_thongbao.Text = " Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!"
                '        End If
                '    Catch ex As Exception
                '        MessageBox.Show("Cập nhật xoá bỏ hồ sơ cũ của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                '    End Try
            Case "tp_hh_hv"
                If (clb_hhhv_hocham.Items.Count = 0 And clb_hhhv_hocvi.Items.Count = 0) Then Return
                Try
                    If (arr_Hhhv_LHocham.Count <= 0 And arr_Hhhv_LHocvi.Count <= 0) Then
                        Return
                    Else
                        Dim strHocham As String = _Globals.GetItems_CheckedListBox(clb_hhhv_hocham, arr_Hhhv_LHocham)
                        Dim strHocvi As String = _Globals.GetItems_CheckedListBox(clb_hhhv_hocvi, arr_Hhhv_LHocvi)
                        If (strHocham = "" And strHocvi = "") Then
                            lbl_thongbao.Text = " Bạn chưa chọn danh sách học hàm hoặc học vị cần xoá!"
                            Return
                        End If
                        If (MessageBox.Show("Bạn có thực sự muốn xoá các bản ghi hồ sơ học hàm - học vị đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            Dim arr_Del_hh As ArrayList = New ArrayList()
                            Dim arr_Del_hv As ArrayList = New ArrayList()
                            arr_Del_hh = Globals.Splip_Strings(strHocham)
                            arr_Del_hv = Globals.Splip_Strings(strHocvi)
                            'Thực hiện xoá các bản ghi đã chọn
                            If (arr_Del_hh.Count > 0) Then
                                For i As Int16 = 0 To arr_Del_hh.Count - 1
                                    _HS_CanBo.Delete_HocHam(arr_Del_hh(i).ToString())
                                Next
                            End If
                            If (arr_Del_hv.Count > 0) Then
                                For i As Int16 = 0 To arr_Del_hv.Count - 1
                                    _HS_CanBo.Delete_HocVi(arr_Del_hv(i).ToString())
                                Next
                            End If
                            lbl_thongbao.Text = " Bạn đã xoá thành công hồ sơ học hàm - học vị của cán bộ!"
                            cb_hhhv_hocham.SelectedIndex = 0
                            cb_hhhv_hocvi.SelectedIndex = 0
                            'Fill data - Hồ sơ công tác của cán bộ
                            Fill_Data(7, False)
                            Fill_Data(8, False)
                            clb_hhhv_hocham_SelectedIndexChanged(sender, Nothing)
                            clb_hhhv_hocvi_SelectedIndexChanged(sender, Nothing)
                        Else
                            _Globals.ResetItems_CheckedListBox(clb_hhhv_hocham)
                            _Globals.ResetItems_CheckedListBox(clb_hhhv_hocvi)
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ học hàm - học vị: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
        End Select
    End Sub
#End Region

#Region "---> Events: Bắt sự kiện ngoại lệ <---"
    Private Sub edt_cmt_so_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_cmt_so.KeyPress
        'Chỉ cho phép nhập số vào TextBox số chứng minh thư
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
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
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
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
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
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
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
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
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
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
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_ms_thue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_tk_so_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub dtpk_ngay_thamnien_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_ngay_thamnien.ValueChanged
        If (FlagEvent = 1 And _IdCB_TS = "") Then
            dtpk_ngay_bc_nhcsxh.Value = dtpk_ngay_thamnien.Value
        End If
    End Sub

    'Private Sub edt_macb_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_macb.Leave
    '    If (edt_macb.Text.Trim() = "" Or Len(edt_macb.Text.Trim()) <= 0) Then
    '        'Assign error Tooltip message To TextBox
    '        erpd_main.SetError(edt_macb, "Mã hiệu của cán bộ không được để trống!")
    '    Else
    '        Dim strSQL As String = ""
    '        If (_IdCanBo <> "") Then
    '            strSQL = String.Format("Select * From HS_CanBo Where MaCB = '{0}' and IdCanBo <> '{1}'", Globals.Find_Replace(edt_macb.Text.Trim().ToString()), _IdCanBo)
    '        Else
    '            strSQL = String.Format("Select * From HS_CanBo Where MaCB = '{0}'", Globals.Find_Replace(edt_macb.Text.Trim().ToString()))
    '        End If
    '        Using db = _SqlHelper.SelectDBRows(strSQL)
    '            If (db.Rows.Count <= 0) Then
    '                'Clear the error Tooltip message for TextBox And Make error icon Invisible
    '                erpd_main.SetError(edt_macb, "")
    '            Else
    '                If Not (db Is Nothing) Then
    '                    If (db.Rows.Count > 0) Then
    '                        erpd_main.SetError(edt_macb, "Mã hiệu của cán bộ không được trùng nhau. Vui lòng kiểm tra lại!")
    '                    End If
    '                End If
    '            End If
    '        End Using
    '    End If
    'End Sub

    Private Sub edt_cmt_so_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_cmt_so.Leave
        If (edt_cmt_so.Text.Trim() = "" Or Len(edt_cmt_so.Text.Trim()) < 9) Then
            erpd_main.SetError(edt_cmt_so, "Độ dài của chuỗi chứng minh thư nhân dân không hợp lệ!")
        Else
            Dim strSQL As String = ""
            If (_IdCanBo <> "") Then
                strSQL = String.Format("Select * From HS_CanBo Where CMT_So = '{0}' and IdCanBo <> '{1}'", Globals.Find_Replace(edt_cmt_so.Text.Trim().ToString()), _IdCanBo)
            Else
                strSQL = String.Format("Select * From HS_CanBo Where CMT_So = '{0}'", Globals.Find_Replace(edt_cmt_so.Text.Trim().ToString()))
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If (db.Rows.Count <= 0) Then
                    'Clear the error Tooltip message for TextBox And Make error icon Invisible
                    erpd_main.SetError(edt_cmt_so, "")
                Else
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            erpd_main.SetError(edt_macb, "Số chứng minh nhân dân của cán bộ không được trùng nhau. Vui lòng kiểm tra lại!")
                        End If
                    End If
                End If
            End Using
        End If
    End Sub

    Private Sub edt_hoten_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hoten.TextChanged
        If (FlagEvent = 1 And _IdCB_TS = "") Then edt_ten_tg.Text = edt_hoten.Text.Trim()
    End Sub
#End Region

#Region "---> Events: Bắt sự kiện ngoại lệ - Di chuyển controls <---"
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
            dtpk_cmt_ngaycap.Focus()
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
            tctrl_hs_cb.SelectedIndex = 0
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
            edt_sofax.Focus()
        End If
    End Sub

    Private Sub edt_dt_nharieng_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dt_nharieng.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dt_coquan.Focus()
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
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
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
            tctrl_hs_cb.SelectedIndex = 1
            ActiveControl = cb_tpgd
        End If
    End Sub

    Private Sub cb_tpgd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tpgd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_ut_giadinh
        End If
    End Sub

    Private Sub cb_ut_giadinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ut_giadinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = clb_ut_banthan
        End If
    End Sub

    Private Sub clb_ut_banthan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles clb_ut_banthan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_dacdiem_banthan
        End If
    End Sub

    Private Sub edt_dacdiem_banthan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dacdiem_banthan.KeyDown
        If (e.KeyCode = Keys.Tab) Then
            edt_quanhe_ng_nn.Focus()
        End If
    End Sub

    Private Sub edt_quanhe_ng_nn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_quanhe_ng_nn.KeyDown
        If (e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_ghichu
        End If
    End Sub

    Private Sub edt_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ghichu.KeyDown
        If (e.KeyCode = Keys.Tab) Then
            tctrl_hs_cb.SelectedIndex = 2
            ActiveControl = cb_tdvh
        End If
    End Sub

    Private Sub cb_tdvh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tdvh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_tdct
        End If
    End Sub

    Private Sub cb_tdct_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_tdct.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = dtpk_ngay_thamnien
        End If
    End Sub

    Private Sub dtpk_ngay_thamnien_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_thamnien.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = dtpk_ngay_vaonganh
        End If
    End Sub

    Private Sub dtpk_ngay_vaonganh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_vaonganh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = dtpk_ngay_vaonhcsxh
        End If
    End Sub

    Private Sub dtpk_ngay_vaonhcsxh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_vaonhcsxh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = dtpk_ngay_bc_nhcsxh
        End If
    End Sub

    Private Sub dtpk_ngay_bc_nhcsxh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_bc_nhcsxh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = dtpk_ngay_cachmang
        End If
    End Sub

    Private Sub dtpk_ngay_cachmang_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_cachmang.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_cm_tochuc
        End If
    End Sub

    Private Sub edt_cm_tochuc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cm_tochuc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_sotruong_ct.Focus()
        End If
    End Sub

    Private Sub edt_cm_tochuc_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cm_tochuc.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_ngay_cachmang.Focus()
        End If
    End Sub

    Private Sub edt_sotruong_ct_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sotruong_ct.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_congviec_lamlau.Focus()
        End If
    End Sub

    Private Sub edt_sotruong_ct_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sotruong_ct.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cm_tochuc.Focus()
        End If
    End Sub

    Private Sub edt_congviec_lamlau_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congviec_lamlau.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_bhxh_so.Focus()
        End If
    End Sub

    Private Sub edt_congviec_lamlau_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congviec_lamlau.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_sotruong_ct.Focus()
        End If
    End Sub

    Private Sub edt_bhxh_so_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_so.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_bhxh_ngaylam.Focus()
        End If
    End Sub

    Private Sub edt_bhxh_so_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_so.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_congviec_lamlau.Focus()
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
            If (FlagEvent = 1) Then
                tctrl_hs_cb.SelectedIndex = 3
                ActiveControl = edt_qd_soqd
            Else
                ActiveControl = btn_save
            End If
        End If
    End Sub

    Private Sub edt_bhxh_noilam_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhxh_noilam.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_bhxh_ngaydong.Focus()
        End If
    End Sub

    Private Sub edt_qd_soqd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qd_soqd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_qd_soqd.Text.Trim() <> "") Then
                cb_qd_loaiqd.Focus()
            Else
                edt_qd_soqd.Focus()
            End If
        End If
    End Sub

    Private Sub edt_qd_soqd_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qd_soqd.KeyUp
        If (e.KeyCode = Keys.Up) Then
            If (FlagEvent = 1) Then
                tctrl_hs_cb.SelectedIndex = 2
                ActiveControl = edt_bhxh_noilam
            End If
        End If
    End Sub

    Private Sub cb_qd_loaiqd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qd_loaiqd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (cb_qd_loaiqd.Items.Count > 0) Then
                If (cb_qd_loaiqd.SelectedIndex > 0) Then
                    dtpk_qd_ngayky.Focus()
                Else
                    cb_qd_loaiqd.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub dtpk_qd_ngayky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_qd_ngayky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_qd_ngayhl.Focus()
        End If
    End Sub

    Private Sub dtpk_qd_ngayhl_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_qd_ngayhl.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_qd_nguoiky.Focus()
        End If
    End Sub

    Private Sub edt_qd_nguoiky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qd_nguoiky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_qd_chucvu.Focus()
        End If
    End Sub

    Private Sub edt_qd_nguoiky_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qd_nguoiky.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = cb_qd_chucvu
        End If
    End Sub

    Private Sub cb_qd_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qd_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_qd_donvi_moi.Focus()
        End If
    End Sub

    Private Sub cb_qd_donvi_moi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qd_donvi_moi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_qd_phongban_moi.Focus()
        End If
    End Sub

    Private Sub cb_qd_phongban_moi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qd_phongban_moi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_qd_chucvu_moi.Focus()
        End If
    End Sub

    Private Sub cb_qd_chucvu_moi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qd_chucvu_moi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_qd_cm_moi.Focus()
        End If
    End Sub

    Private Sub cb_qd_cm_moi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qd_cm_moi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_qd_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_qd_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qd_ghichu.KeyDown
        If (e.KeyCode = Keys.Tab) Then
            ActiveControl = btn_save
        End If
    End Sub

    'Bắt sự kiện di chuyển - Tab - Hồ sơ công tác

    Private Sub dtpk_ct_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ct_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_ct_chucvu.Focus()
        End If
    End Sub

    Private Sub dtpk_ct_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ct_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_ct_chucvu.Focus()
        End If
    End Sub

    Private Sub edt_ct_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_ct_chucvu.Text.Trim() <> "") Then
                edt_ct_diachi.Focus()
            Else
                edt_ct_chucvu.Focus()
            End If
        End If
    End Sub

    Private Sub edt_ct_chucvu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_chucvu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = dtpk_ct_denngay
        End If
    End Sub

    Private Sub edt_ct_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_ct_diachi.Text.Trim() <> "") Then
                edt_ct_lydo.Focus()
            Else
                edt_ct_diachi.Focus()
            End If
        End If
    End Sub

    Private Sub edt_ct_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_ct_chucvu
        End If
    End Sub

    Private Sub edt_ct_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then edt_ct_ghichu.Focus()
    End Sub

    Private Sub edt_ct_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_ct_diachi
        End If
    End Sub

    Private Sub edt_ct_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = btn_save
        End If
    End Sub

    Private Sub edt_ct_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ct_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_ct_lydo
        End If
    End Sub

    'Bắt sự kiện di chuyển - Tab - Hồ sơ Xuất ngoại

    Private Sub dtpk_xn_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_xn_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_xn_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_xn_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_xn_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_xn_nuocden.Focus()
        End If
    End Sub

    Private Sub cb_xn_nuocden_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_xn_nuocden.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_xn_mucdich.Focus()
        End If
    End Sub

    Private Sub edt_xn_mucdich_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_mucdich.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then edt_xn_soqd.Focus()
    End Sub

    Private Sub edt_xn_mucdich_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_mucdich.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = cb_xn_nuocden
        End If
    End Sub

    Private Sub edt_xn_soqd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_soqd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_xn_soqd.Text.Trim() <> "") Then
                dtpk_xn_ngayky.Focus()
            Else
                edt_xn_soqd.Focus()
            End If
        End If
    End Sub

    Private Sub edt_xn_soqd_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_soqd.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_xn_mucdich
        End If
    End Sub

    Private Sub dtpk_xn_ngayky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_xn_ngayky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_xn_nguoiky.Focus()
        End If
    End Sub

    Private Sub edt_xn_nguoiky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_nguoiky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_xn_nguoiky.Text.Trim() <> "") Then
                cb_xn_chucvu.Focus()
            Else
                edt_xn_nguoiky.Focus()
            End If
        End If
    End Sub

    Private Sub edt_xn_nguoiky_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_nguoiky.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = dtpk_xn_ngayky
        End If
    End Sub

    Private Sub cb_xn_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_xn_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_xn_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_xn_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_save.Focus()
        End If
    End Sub

    Private Sub edt_xn_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_xn_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = cb_xn_chucvu
        End If
    End Sub

    'Bắt sự kiện di chuyển - Tab - Hồ sơ Hộ Chiếu

    Private Sub edt_hc_sohc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hc_sohc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hc_sohc.Text.Trim() <> "") Then
                cb_hc_loaihc.Focus()
            Else
                edt_hc_sohc.Focus()
            End If
        End If
    End Sub

    Private Sub cb_hc_loaihc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hc_loaihc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hc_ngaycap.Focus()
        End If
    End Sub

    Private Sub dtpk_hc_ngaycap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hc_ngaycap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hc_noicap.Focus()
        End If
    End Sub

    Private Sub edt_hc_noicap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hc_noicap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then dtpk_hc_ngayhh.Focus()
    End Sub

    Private Sub edt_hc_noicap_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hc_noicap.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = dtpk_hc_ngaycap
        End If
    End Sub

    Private Sub dtpk_hc_ngayhh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hc_ngayhh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hc_tinhtrang.Focus()
        End If
    End Sub

    Private Sub cb_hc_tinhtrang_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hc_tinhtrang.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hc_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_hc_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hc_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_save.Focus()
        End If
    End Sub

    Private Sub edt_hc_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hc_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = cb_hc_tinhtrang
        End If
    End Sub

    'Bắt sự kiện di chuyển - Tab - Hồ sơ Tham gia Lực lượng vũ trang
    Private Sub dtpk_llvt_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_llvt_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_llvt_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_llvt_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_llvt_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_llvt_phanloai.Focus()
        End If
    End Sub

    Private Sub cb_llvt_phanloai_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_llvt_phanloai.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_llvt_quanham.Focus()
        End If
    End Sub

    Private Sub cb_llvt_quanham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_llvt_quanham.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_llvt_chucvu.Focus()
        End If
    End Sub

    Private Sub edt_llvt_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_llvt_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then edt_llvt_donvi.Focus()
    End Sub

    Private Sub edt_llvt_chucvu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_llvt_chucvu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = cb_llvt_quanham
        End If
    End Sub

    Private Sub edt_llvt_donvi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_llvt_donvi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then edt_llvt_ghichu.Focus()
    End Sub

    Private Sub edt_llvt_donvi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_llvt_donvi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_llvt_chucvu
        End If
    End Sub

    Private Sub edt_llvt_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_llvt_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = btn_save
        End If
    End Sub

    Private Sub edt_llvt_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_llvt_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_llvt_donvi
        End If
    End Sub

    'Bắt sự kiện di chuyển - Tab - Hồ sơ Cũ <Lưu lại các thông tin của cán bộ khi chưa thoát ly>

    Private Sub dtpk_hscu_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hscu_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hscu_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_hscu_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hscu_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hscu_nghenghiep.Focus()
        End If
    End Sub

    Private Sub edt_hscu_nghenghiep_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hscu_nghenghiep.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then edt_hscu_diachi.Focus()
    End Sub

    Private Sub edt_hscu_nghenghiep_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hscu_nghenghiep.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = dtpk_hscu_denngay
        End If
    End Sub

    Private Sub edt_hscu_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hscu_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then edt_hscu_ghichu.Focus()
    End Sub

    Private Sub edt_hscu_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hscu_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_hscu_nghenghiep
        End If
    End Sub

    Private Sub edt_hscu_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hscu_ghichu.KeyDown
        If (Keys.Tab = True) Then
            pnl_main.Focus()
            ActiveControl = btn_save
        End If
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            pnl_main.Focus()
            ActiveControl = btn_save
        End If
    End Sub

    Private Sub edt_hscu_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hscu_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_hscu_diachi
        End If
    End Sub

    Private Sub edt_ns_diachi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_ns_diachi.TextChanged
        If (FlagEvent = 1 And _IdCB_TS = "") Then
            edt_nq_diachi.Text = edt_ns_diachi.Text.Trim()
        End If
    End Sub

    Private Sub edt_tt_diachi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tt_diachi.TextChanged
        If (FlagEvent = 1 And _IdCB_TS = "") Then
            edt_ttr_diachi.Text = edt_tt_diachi.Text.Trim()
        End If
    End Sub

    Private Sub edt_tt_dienthoai_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tt_dienthoai.TextChanged
        If (FlagEvent = 1 And _IdCB_TS = "") Then
            edt_ttr_dienthoai.Text = edt_tt_dienthoai.Text.Trim()
            edt_dt_nharieng.Text = edt_tt_dienthoai.Text.Trim()
        End If
    End Sub
#End Region

    Private Sub edt_macb_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles edt_macb.LostFocus
        Dim dbconn As DBAccess = New DBAccess
        If _IdCanBo <> "" And dbconn.getString("SELECT IdOld FROM HS_Canbo WHERE idCanbo='" & _IdCanBo & "'") <> "" Then
            Exit Sub
        End If
        If (edt_macb.Text.Trim() = "" Or Len(edt_macb.Text.Trim()) <= 0) Then
            'Assign error Tooltip message To TextBox
            erpd_main.SetError(edt_macb, "Mã cán bộ không được để trống!")
        Else
            Dim strSQL As String = ""
            'If (_IdCanBo <> "") Then
            '    strSQL = String.Format("Select * From HS_CanBo Where MaCB = '{0}' and IdCanBo <> '{1}'", Globals.Find_Replace(edt_macb.Text.Trim().ToString()), _IdCanBo)
            'Else
            '    strSQL = String.Format("Select * From HS_CanBo Where MaCB = '{0}'", Globals.Find_Replace(edt_macb.Text.Trim().ToString()))
            'End If
            'Using db = _SqlHelper.SelectDBRows(strSQL)
            '    If (db.Rows.Count <= 0) Then
            '        'Clear the error Tooltip message for TextBox And Make error icon Invisible
            '        erpd_main.SetError(edt_macb, "")
            '    Else
            '        If Not (db Is Nothing) Then
            '            If (db.Rows.Count > 0) Then
            '                erpd_main.SetError(edt_macb, "Mã hiệu của cán bộ không được trùng nhau. Vui lòng kiểm tra lại!")
            '            End If
            '        End If
            '    End If
            'End Using
            edt_macb.Text = formatLenString(5, edt_macb.Text.Trim)
            If cb_donvi.SelectedIndex >= 0 AndAlso checkMaCanBo(arr_Donvi(cb_donvi.SelectedIndex), _IdCanBo, edt_macb.Text) = 0 Then
                'Lấy dẫy số mã cán bộ theo quy định để cho người dùng cập nhật
                Dim sCodeByBranch As String = ""
                If DONVI = "000100" Or DONVI = "000101" Or DONVI = "000196" Or DONVI = "000197" Then
                    strSQL = String.Format("Select Top 1 * From ChiNhanh Where Ma_So = '{0}' Order By Id", DONVI)
                Else
                    strSQL = String.Format("Select Top 1 * From ChiNhanh Where Ma_So Like '{0}%' Order By Id", DONVI.Substring(0, 4))
                End If
                Dim db_cn As DataTable = Nothing
                db_cn = _SqlHelper.SelectDBRows(strSQL)
                If Not (db_cn Is Nothing) Then
                    If db_cn.Rows.Count > 0 Then
                        sCodeByBranch = db_cn.Rows(0)("MaCB_Begin").ToString().Trim() + " => " + db_cn.Rows(0)("MaCB_End").ToString().Trim()
                    End If
                End If
                erpd_main.SetError(edt_macb, String.Format("Mã cán bộ '{0}' đã tồn tại, hãy nhập mã cán bộ khác thuộc một số trong dẫy số '{1}' theo đơn vị '{2}'. Vui lòng kiểm tra lại!", edt_macb.Text, sCodeByBranch, BrandNameByUserLogin))
                ActiveControl = edt_macb
            Else
                erpd_main.SetError(edt_macb, "")
            End If
        End If
    End Sub

    Private Sub tctrl_hs_cb_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tctrl_hs_cb.SelectedIndexChanged
        If FlagEvent = 2 Then
            If tctrl_hs_cb.SelectedTab.Name = "tp_qd_nhansu" Then
                btn_save.Visible = False
                btn_huybo.Visible = False
            Else
                btn_save.Visible = True
                btn_huybo.Visible = True
                btn_save.Enabled = True
                btn_huybo.Enabled = True
            End If
        End If
    End Sub

    Private Sub edt_nh_makh_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_nh_makh.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_masothue_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_masothue.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_nh_sotk_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_nh_sotk.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub


End Class