Public Class frmHS_HdldTS

#Region "---> Khai báo thuộc tính và Khởi tạo đối tượng <---"
    Private ARL_ChiNhanh As ArrayList = New ArrayList()
    Private ARL_PhBanDonVi As ArrayList = New ArrayList()

    'Mảng lưu id danh mục - Hình thức trả lương --> Truy suất dl đến danh mục với id gốc là 40
    Private arr_Ht_Traluong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Nghi dinh lương
    Private arr_NghiDinhLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Bảng lương
    Private arr_BangLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Ngạch lương
    Private arr_NgachLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Bậc lương
    Private arr_BacLuong As ArrayList = New ArrayList()
    'Mảng lưu Danh sách id Chức vụ
    Private arr_Chucvu As ArrayList = New ArrayList()
    Private arr_ChuyenMon As ArrayList = New ArrayList()
    Private arrQd_LoaiQd As ArrayList = New ArrayList()
    Private arrQd_LyDoTV As ArrayList = New ArrayList()
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
    Private _Globals As Globals = New Globals
    Private _Labour As clsHS_Hdld = New clsHS_Hdld()
    Private _SqlHelper As DBAccess = New DBAccess()
    Private strSQL As String = ""
    Private _IdChiNhanh As Integer = 0
    Private _IdHd As String = ""

    Private _IdCanBo As String    'Tổng số bản ghi tìm thấy
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

    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    'Khai báo các thông tin add CheckBox vào cột tiêu đề chọn cả các items Check trên lưới
    Private ckb_ChoiceAll As CheckBox = Nothing   'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0          'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0   'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False  'Cờ báo việc Checkall
#End Region

#Region "---> Functions: Các hàm chính <---"
    Private Sub ResetAll_Controls(ByVal status As Boolean)
        If (status = True) Then
            lbl_macb.Text = ""
            lbl_hoten.Text = ""
            lbl_gioitinh.Text = ""
            lbl_ngaysinh.Text = ""
            lbl_donvi.Text = ""
            lbl_socmt.Text = ""
            lbl_ngaycap.Text = ""
            lbl_noicap.Text = ""
        End If
        edt_sohd.Text = ""
        dtpk_ngayky.Text = DateTime.Now.ToShortDateString()
        edt_nguoiky.Text = ""
        cb_chucvu.SelectedIndex = 0
        dtpk_tungay.Text = DateTime.Now.ToShortDateString()
        dtpk_denngay.Text = DateTime.Now.ToShortDateString()

        edt_gio_bd.Text = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 7, 30, 0)
        edt_gio_kt.Text = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 16, 30, 0)
        rb_thangbang.Checked = True

        cb_loaihinh.SelectedIndex = 0
        cb_ht_traluong.SelectedIndex = 0
        rb_DinhBien.Checked = True
        rb_hdld_CheckedChanged(Nothing, Nothing)

        edt_tylehuong.Text = "0"
        edt_cv_damnhan.Text = ""
        cb_chuyenmon.SelectedIndex = 0
        edt_tienluong.Text = "0"
        ckb_bhxh.Checked = False
        ckb_bhyt.Checked = False
        cb_nghidinhlg.SelectedIndex = 0
        cb_nghidinhlg_SelectedIndexChanged(Nothing, Nothing)

        cbIsChamDut.Checked = False
        grpChamDut.Enabled = False
        dpkNgayHL_TV.Value = DateTime.Now.ToShortDateString()
        txtSoQD_TV.Text = ""
        dpkNgayKy_TV.Value = DateTime.Now.ToShortDateString()
        cboLyDo_TV.SelectedIndex = 0
        txtTroCap_TV.Text = 0
        txtTroCapKhac.Text = 0
        txtTienBoiThuong.Text = 0
        txtTienThuHoi.Text = 0
        txtGhiChu.Text = ""

        If cb_chinhanh.Items.Count <> 0 And _IdHd = "" Then
            If (_IdCanBo <> "") Then
                cb_chinhanh.SelectedIndex = IIf(_Node.ToString.Trim().Substring(4) <> "", CType(ARL_ChiNhanh.IndexOf(_Node.ToString.Trim().Substring(4)), Integer), 0)
                cb_chinhanh_SelectedIndexChanged(Nothing, Nothing)
            Else
                cb_chinhanh.SelectedIndex = 1
            End If
        Else
            cb_chinhanh.SelectedIndex = 0
        End If
        cb_chinhanh_SelectedIndexChanged(Nothing, Nothing)

    End Sub

    Private Sub Fill_Data()
        ckb_ChoiceAll.Checked = False
        dgv_main.Rows.Clear()
        If (_IdCanBo <> "") Then
            Using db As DataTable = _HS_CanBo.GetHSCB_TS_HDLD_GetSearch(_IdCanBo, "", 0, "")
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        For i As Integer = 0 To db.Rows.Count - 1
                            dgv_main.Rows.Add()
                            dgv_main.Rows(i).Cells("cln_Id").Value = db.Rows(i)("IdCBTS_HDLD").ToString().Trim()

                            dgv_main.Rows(i).Cells("cln_PhBanDonVi_HT").Value = db.Rows(i)("PhBanDonVi_HT").ToString().Trim()
                            dgv_main.Rows(i).Cells("cln_SoHD").Value = db.Rows(i)("SoQD").ToString().Trim()
                            'If db.Rows(i)("NgayQD").ToString().Trim() <> "" Then
                            '    dgv_main.Rows(i).Cells("cln_NgayKy").Value = CType(db.Rows(i)("NgayQD").ToString(), DateTime).ToString("dd-MM-yyyy")
                            'Else
                            '    dgv_main.Rows(i).Cells("cln_NgayKy").Value = ""
                            'End If
                            'dgv_main.Rows(i).Cells("cln_NguoiKy").Value = db.Rows(i)("NguoiQD").ToString().Trim()
                            'dgv_main.Rows(i).Cells("cln_Chucvu").Value = db.Rows(i)("ChucVu").ToString().Trim()
                            dgv_main.Rows(i).Cells("cln_Loaihinh").Value = clsHS_CanBo.GetLoaiHinh(CType(db.Rows(i)("Loai").ToString().Trim(), Byte))
                            dgv_main.Rows(i).Cells("cln_HT_TraLuong").Value = _HS_CanBo.GetNameByCode(String.Format("Select id,Ten_Goi from DanhMuc Where id = {0} and id_goc = 40 and Status = 1", CType(db.Rows(i)("IdHT_TraLuong").ToString(), Int32)))
                            'Lấy giờ làm việc của cán bộ tập sự trong hợp đồng 
                            'dgv_main.Rows(i).Cells("cln_GioLv").Value = db.Rows(i)("TuGio").ToString().Trim() + " - " + db.Rows(i)("DenGio").ToString().Trim()
                            If db.Rows(i)("TuNgay").ToString().Trim() <> "" Then
                                dgv_main.Rows(i).Cells("cln_Tungay").Value = CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                            Else
                                dgv_main.Rows(i).Cells("cln_Tungay").Value = ""
                            End If

                            If db.Rows(i)("DenNgay").ToString().Trim() <> "" Then
                                dgv_main.Rows(i).Cells("cln_Denngay").Value = CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                            Else
                                dgv_main.Rows(i).Cells("cln_Denngay").Value = ""
                            End If

                            If (db.Rows(i)("IdBacLuong").ToString().Trim() <> "0") Then
                                dgv_main.Rows(i).Cells("cln_HT_HuongLuong").Value = "Thang - bảng lương"
                                dgv_main.Rows(i).Cells("cln_Tienluong").Value = ""
                                dgv_main.Rows(i).Cells("cln_BacLuong").Value = IIf(CType(db.Rows(i)("IdBacLuong"), Integer) > 0, _HS_CanBo.GetNameByCode(String.Format("Select IdBacLuong,BacLuong From BacLuong Where IdBacLuong = {0} and Status = 1", CType(db.Rows(i)("IdBacLuong"), Integer))), "")
                                dgv_main.Rows(i).Cells("cln_Hs_Luong").Value = IIf(db.Rows(i)("Heso").ToString().Trim() <> "", CType(db.Rows(i)("Heso"), Double).ToString("N2"), "0")
                                dgv_main.Rows(i).Cells("cln_Tl_Huong").Value = db.Rows(i)("TyleHuong").ToString().Trim()
                            Else
                                dgv_main.Rows(i).Cells("cln_HT_HuongLuong").Value = "Lương trọn gói"

                                dgv_main.Rows(i).Cells("cln_BacLuong").Value = ""
                                dgv_main.Rows(i).Cells("cln_Hs_Luong").Value = ""
                                dgv_main.Rows(i).Cells("cln_Tl_Huong").Value = ""
                                dgv_main.Rows(i).Cells("cln_TienLuong").Value = IIf(db.Rows(i)("Tien_Luong").ToString().Trim() <> "", Double.Parse(db.Rows(i)("Tien_Luong").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0")
                            End If
                            dgv_main.Rows(i).Cells("cln_Cv_Damnhan").Value = db.Rows(i)("CongViec").ToString().Trim()
                            If db.Rows(i)("TV_Ngay_HL").ToString().Trim() <> "" And CType(db.Rows(i)("TV_Ngay_HL").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01/01/1900" Then
                                dgv_main.Rows(i).Cells("cln_TV_Ngay_HL").Value = CType(db.Rows(i)("TV_Ngay_HL").ToString(), DateTime).ToString("dd-MM-yyyy")
                            Else
                                dgv_main.Rows(i).Cells("cln_TV_Ngay_HL").Value = ""
                            End If

                        Next
                    End If
                End If
            End Using
            dgv_main_CellClick(Nothing, Nothing)
        End If
        TotalCheckBoxes = dgv_main.RowCount
        TotalCheckedCheckBoxes = 0
    End Sub

    Private Function IsValid() As Boolean
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ tập sự cần cập nhật hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            Return False
        End If
        If (edt_sohd.Text.Trim() = "") Then
            MessageBox.Show("Số hợp đồng lao động không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_sohd
            Return False
        End If
        'Bắt điều kiện trùng số Hợp đồng lao động
        If (edt_sohd.Text.Trim() <> "") Then
            If (_IdHd = "") Then
                strSQL = String.Format("Select * From HSCB_TS_HDLD Where SoQD = '{0}'", Globals.Find_Replace(edt_sohd.Text.Trim().ToString()))
            Else
                strSQL = String.Format("Select * From HSCB_TS_HDLD Where SoQD = '{0}' And IdCBTS_HDLD <> '{1}'", Globals.Find_Replace(edt_sohd.Text.Trim().ToString()), _IdHd)
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        MessageBox.Show(String.Format("Số hợp đồng: {0} đã tồn tại. Vui lòng kiểm tra lại!", edt_sohd.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_sohd
                        Return False
                    End If
                End If
            End Using
        End If

        If (edt_nguoiky.Text.Trim() = "") Then
            MessageBox.Show("Người ký hợp đồng lao động không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_nguoiky
            Return False
        End If
        If (cb_chucvu.SelectedIndex <= 0 And cb_chucvu.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn chức vụ người ký hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_chucvu
            Return False
        End If
        If (cb_loaihinh.SelectedIndex <= 0 And cb_loaihinh.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn loại hình công việc mà cán bộ tập sự tham gia!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loaihinh
            Return False
        End If
        If (cb_chuyenmon.SelectedIndex <= 0 And cb_chuyenmon.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn chuyên môn cán bộ đảm nhiệm!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_chuyenmon
            Return False
        End If
        If (cb_ht_traluong.SelectedIndex <= 0 And cb_ht_traluong.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn hình thức trả lương cho cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_ht_traluong
            Return False
        End If
        'Bắt điều kiện hợp lệ của Giờ làm việc
        If (edt_gio_bd.Value.Hour >= edt_gio_kt.Value.Hour) Then
            MessageBox.Show("Giờ làm việc trong hợp đồng không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_gio_bd
            Return False
        End If
        If dtpk_tungay.Value >= dtpk_denngay.Value Then
            MessageBox.Show("Ngày bắt đầu không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = dtpk_tungay
            Return False
        End If

        If (cb_chinhanh.SelectedIndex <= 0 And cb_chinhanh.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn chi nhánh của cán bộ làm việc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_chinhanh
            Return False
        End If
        If (cb_phongban_pgd.SelectedIndex <= 0 And cb_phongban_pgd.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn Phòng ban/đơn vị của cán bộ làm việc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_phongban_pgd
            Return False
        End If
        'Xét trường hợp đan xen của khoảng thời gian trong Hợp đồng lao động
        If (_IdHd = "") Then
            strSQL = String.Format("Select * From HSCB_TS_HDLD Where IdCanbo = '{0}' Order by TuNgay Asc", _IdCanBo)
        Else
            strSQL = String.Format("Select * From HSCB_TS_HDLD Where IdCanbo = '{0}' And IdCBTS_HDLD <> '{1}' Order by TuNgay Asc", _IdCanBo, _IdHd)
        End If
        Dim Exist As Boolean = False
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                        Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                        If ((_TuNgay <= dtpk_tungay.Value And dtpk_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_denngay.Value And dtpk_denngay.Value <= _DenNgay) Or (dtpk_tungay.Value <= _TuNgay And dtpk_denngay.Value >= _DenNgay)) Then
                            Exist = True
                            Exit For
                        End If
                    Next
                    If (Exist = True) Then
                        MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hợp đồng không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_tungay
                        Return False
                    End If
                End If
            End If
        End Using

        If (rb_thangbang.Checked = True) Then
            If (cb_bacluong.SelectedIndex <= 0 And cb_bacluong.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn bậc lương!" + vbCrLf + "Lưu ý: Muốn chọn dữ liệu bậc lương thì trước hết bạn phải chọn dữ liệu Bảng lương rồi đến Ngạch lương.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_bacluong
                Return False
            End If
            If (edt_tylehuong.Text.Trim() = "" Or CType(edt_tylehuong.Text.Trim(), Double) <= 0) Then
                MessageBox.Show("Tỷ lệ hưởng lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_tylehuong
                Return False
            End If
            If (edt_tylehuong.Text.Trim() <> "") Then
                If (CType(edt_tylehuong.Text.Trim(), Double) > 300) Then
                    MessageBox.Show("Tỷ lệ hưởng lương không hợp lệ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_tylehuong
                    Return False
                End If
            End If
        End If
        If (rb_trongoi.Checked = True) Then
            If (edt_tienluong.Text.Trim() = "" Or CType(MoneyValue(edt_tienluong.Text.Trim()), Double) <= 0) Then
                MessageBox.Show("Số tiền lương được hưởng của cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_tienluong
                Return False
            End If
        End If
        If cbIsChamDut.Checked = True Then
            If dpkNgayHL_TV.Value < dtpk_tungay.Value Then
                MessageBox.Show("Ngày hiệu lực của chấm dứt hợp đồng/QĐ tạm tuyển không được trước ngày bắt đầu của Hợp đồng/QĐ tạm tuyển !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = dpkNgayHL_TV
                Return False
            End If
            If (cboLyDo_TV.SelectedIndex <= 0 And cboLyDo_TV.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn lý do thôi việc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cboLyDo_TV
                Return False
            End If
        End If
        Return True
    End Function

    ''' <summary>
    ''' Hàm fill dữ liệu vào controls khi người dùng click vào row truyền vào
    ''' </summary>
    ''' <param name="_CodeId">Chỉ số xác định bản ghi - Bản ghi đang select</param>
    ''' <remarks></remarks>
    Private Sub FillData_SelectRow(ByVal _CodeId)
        Dim dr As DataRow
        _IdHd = _CodeId
        Using dbHopDong As DataTable = _HS_CanBo.GetHSCB_TS_HDLD_GetSearch("", _IdHd, 0, "")
            If Not (dbHopDong Is Nothing) Then
                If (dbHopDong.Rows.Count > 0) Then
                    dr = dbHopDong.Rows(0)
                End If
            End If
        End Using

        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                If (dr("Status").ToString() <> "") Then
                    If dr("Status").ToString() <> "True" Then
                        rb_DinhBien.Checked = True
                        rb_hdld_CheckedChanged(Nothing, Nothing)
                    Else
                        rb_qdtamtuyen.Checked = True
                        rb_qdtamtuyen_CheckedChanged(Nothing, Nothing)
                    End If
                End If
                edt_sohd.Text = dr("SoQD").ToString().Trim()
                If dr("NgayQD").ToString().Trim() <> "" Then
                    dtpk_ngayky.Value = CType(dr("NgayQD").ToString(), DateTime)
                End If
                edt_nguoiky.Text = dr("NguoiQD").ToString().Trim()
                cb_chucvu.SelectedIndex = IIf(dr("IdCV_Nguoi_QD").ToString() <> "", CType(arr_Chucvu.IndexOf(dr("IdCV_Nguoi_QD").ToString()), Integer), 0)
                cb_loaihinh.SelectedIndex = CType(dr("Loai").ToString(), Byte)
                cb_ht_traluong.SelectedIndex = CType(arr_Ht_Traluong.IndexOf(dr("IdHT_TraLuong").ToString()), Integer)
                'Giờ làm việc
                If (dr("TuGio").ToString().Trim() <> "") Then
                    edt_gio_bd.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, CInt(dr("TuGio").ToString().Trim().Substring(0, 2)), CInt(dr("TuGio").ToString().Trim().Substring(3, 2)), 0)
                Else
                    edt_gio_bd.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
                End If
                If (dr("DenGio").ToString().Trim() <> "") Then
                    edt_gio_kt.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, CInt(dr("DenGio").ToString().Trim().Substring(0, 2)), CInt(dr("DenGio").ToString().Trim().Substring(3, 2)), 0)
                Else
                    edt_gio_kt.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
                End If
                If dr("Tungay").ToString().Trim() <> "" Then
                    dtpk_tungay.Value = CType(dr("Tungay").ToString(), DateTime)
                End If
                If dr("DenNgay").ToString().Trim() <> "" Then
                    dtpk_denngay.Value = CType(dr("DenNgay").ToString(), DateTime)
                End If
                'If (rb_DinhBien.Checked = True) Then
                '    ckb_bhxh.Checked = IIf(dr("BHXH").ToString() = True, True, False)
                '    ckb_bhyt.Checked = IIf(dr("BHYT").ToString() = True, True, False)
                'End If
                ckb_bhxh.Checked = IIf(dr("BHXH").ToString() = True, True, False)
                ckb_bhyt.Checked = IIf(dr("BHYT").ToString() = True, True, False)
                'Load thông tin về Lương -> Từ bậc lương
                Dim _NghiDinhLuongId As Int32 = 0
                Dim _BangLuongId As Int32 = 0
                Dim _NgachLuongId As Int32 = 0
                Dim _BacLuongId As Int32 = 0
                _BacLuongId = IIf(dr("IdBacLuong").ToString() <> "0", CType(dr("IdBacLuong").ToString(), Integer), 0) ' Id bậc lương
                If (_BacLuongId > 0) Then
                    rb_thangbang.Checked = True
                    _NgachLuongId = _Labour.GetIdNgachLuong(_BacLuongId)
                    If (_NgachLuongId > 0) Then
                        _BangLuongId = _Labour.GetIdBangLuong(_NgachLuongId)
                    End If
                    If (_BangLuongId > 0) Then
                        _NghiDinhLuongId = _Labour.GetIdNghiDinhLuong(_BangLuongId)
                    End If
                    cb_nghidinhlg.SelectedIndex = CType(arr_NghiDinhLuong.IndexOf(_NghiDinhLuongId.ToString()), Integer)
                    cb_nghidinhlg_SelectedIndexChanged(Nothing, Nothing)
                    cb_bangluong.SelectedIndex = CType(arr_BangLuong.IndexOf(_BangLuongId.ToString()), Integer)
                    cb_bangluong_SelectedIndexChanged(Nothing, Nothing)
                    cb_ngachluong.SelectedIndex = CType(arr_NgachLuong.IndexOf(_NgachLuongId.ToString()), Integer)
                    cb_ngachluong_SelectedIndexChanged(Nothing, Nothing)
                    cb_bacluong.SelectedIndex = CType(arr_BacLuong.IndexOf(dr("IdBacLuong").ToString()), Integer)
                    cb_bacluong_SelectedIndexChanged(Nothing, Nothing)
                    edt_tylehuong.Text = IIf(dr("TyleHuong").ToString().Trim() <> "", CType(dr("TyleHuong"), Double).ToString("N2"), "0")
                    lbl_title_nghidinh.Width = 112
                Else
                    rb_trongoi.Checked = True
                    edt_tienluong.Text = dr("Tien_Luong").ToString().Trim()
                    lbl_title_nghidinh.Width = 127
                End If
                cb_chuyenmon.SelectedIndex = IIf(dr("IdChuyenMon").ToString() <> "", CType(arr_ChuyenMon.IndexOf(dr("IdChuyenMon").ToString()), Integer), 0)
                edt_cv_damnhan.Text = dr("CongViec").ToString().Trim()

                '  obj_hdld.TV_Ngay_HL = DateTime.Parse("01/01/1900")
                If dr("TV_Ngay_HL").ToString().Trim() <> "" And CType(dr("TV_Ngay_HL").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900" Then
                    cbIsChamDut.Checked = True
                    grpChamDut.Enabled = True
                    dpkNgayHL_TV.Value = CType(dr("TV_Ngay_HL").ToString(), DateTime)
                    txtSoQD_TV.Text = dr("TV_So_QD").ToString().Trim()
                    dpkNgayKy_TV.Value = CType(dr("TV_NgayKy_QD").ToString(), DateTime)
                    cboLyDo_TV.SelectedIndex = IIf(dr("TV_IdLyDo").ToString() <> "", CType(arrQd_LyDoTV.IndexOf(dr("TV_IdLyDo").ToString()), Integer), 0)
                    txtTroCap_TV.Text = dr("TV_TroCap_ThoiViec").ToString().Trim()
                    txtTroCapKhac.Text = dr("TV_TroCap_Khac").ToString().Trim()
                    txtTienBoiThuong.Text = dr("TV_SoTien_BoiThuong").ToString().Trim()
                    txtTienThuHoi.Text = dr("TV_SoTien_ThuHoi").ToString().Trim()
                Else
                    cbIsChamDut.Checked = False
                    grpChamDut.Enabled = False
                End If
                txtGhiChu.Text = dr("GhiChu").ToString().Trim()



                cb_chinhanh.SelectedIndex = IIf(dr("ChiNhanhId").ToString() <> "", CType(ARL_ChiNhanh.IndexOf(dr("ChiNhanhId").ToString()), Integer), 0)
                cb_chinhanh_SelectedIndexChanged(Nothing, Nothing)
                cb_phongban_pgd.SelectedIndex = IIf(dr("PhongBanDonViId").ToString() <> "", CType(ARL_PhBanDonVi.IndexOf(dr("PhongBanDonViId").ToString()), Integer), 0)

            End If
        End If
    End Sub
#End Region

#Region "---> Events: Các Sự kiện chính <---"
    Private Sub frmHS_HdldTS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _HS_CanBo.Create_Frame(dgv_main)
        AddHeaderCheckBox()
        'Gọi một số sự kiện liên quan đến Check all items trên lưới dữ liệu
        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged
        AddHandler dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged
        'Thực hiện fill dữ liệu vào các ComboBox
        'Fill dữ liệu Hình thức trả lương cán bộ
        arr_Ht_Traluong.Clear()
        cb_ht_traluong.Items.Clear()
        arr_Ht_Traluong = _Globals.Bind_ComBoBox(cb_ht_traluong, clsHT_DanhMuc.Sql_HtTrLuong, "---Hình thức trả lương---")

        'Fill dữ liệu Nghị định lương cán bộ
        strSQL = "Select IdNDLuong, TenND From NghiDinhLuong Where Status = 1"
        arr_NghiDinhLuong.Clear()
        cb_nghidinhlg.Items.Clear()
        arr_NghiDinhLuong = _Globals.Bind_ComBoBox(cb_nghidinhlg, strSQL, "---Nghị định lương---")

        'Chức vụ người ký quyết định trong hợp đồng lao động
        arr_Chucvu.Clear()
        cb_chucvu.Items.Clear()
        arr_Chucvu = _Globals.Bind_ComBoBox(cb_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ người ký---")

        arr_ChuyenMon.Clear()
        cb_chuyenmon.Items.Clear()
        arr_ChuyenMon = _Globals.Bind_ComBoBox(cb_chuyenmon, clsHT_DanhMuc.Sql_Chuyenmon, "---Chuyên môn---")

        arrQd_LyDoTV.Clear()
        cboLyDo_TV.Items.Clear()
        arrQd_LyDoTV = _Globals.Bind_ComBoBox(cboLyDo_TV, "Select id,Ten_Goi from DanhMuc Where id_goc = 22 and Status = 1 Order by Ten_Goi Asc", "---Lý do thôi việc---")
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Thực hiện fill dữ liệu vào Các ComBoBox
        If TRUCTHUOC = 1 Then
            strSQL = "Select Id,Ten_Goi From ChiNhanh Where Status = 1 And Id_Goc IN (0,1)"
        Else
            'strSQL = "Select Id,Ten_Goi From ChiNhanh Where Status = 1 And Ma_So='" & DONVI.Trim & "'"
            If DONVI = "000196" Or DONVI = "000197" Or DONVI = "000101" Or DONVI = "000100" Or DONVI = "000199" Then
                strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And ma_so='" & DONVI.Trim & "'"
            Else
                strSQL = "Select Id,Ten_Goi From ChiNhanh Where Status = 1 And Id_Goc In (1) And Id In (Select Distinct X.Id_Goc From ChiNhanh X Where X.Ma_So Like '" & DONVI.Substring(0, 4) & "%')"
            End If
        End If
        'Fill dữ liệu combobox đơn vị công tác

        ARL_ChiNhanh.Clear()
        cb_chinhanh.Items.Clear()
        ARL_ChiNhanh = _Globals.Bind_ComBoBox(cb_chinhanh, strSQL, "---Hội sở chính/Chi nhánh---")
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ResetAll_Controls(True)
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
        'Thực hiện load thông tin về hồ sơ chung của cán bộ tập sự
        If (_IdCanBo <> "") Then
            Dim dr As DataRow
            dr = _HS_CanBo.GetTrainee(0, _IdCanBo, "", "", CType("1", Byte), CType("0", Byte))
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    lbl_macb.Text = dr("MaCB").ToString()
                    lbl_hoten.Text = dr("HoTen").ToString()
                    lbl_gioitinh.Text = dr("GioiTinh_HT").ToString()
                    If dr("NgaySinh").ToString() <> "" Then lbl_ngaysinh.Text = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_donvi.Text = dr("ChiNhanh_HT").ToString()
                    _IdChiNhanh = IIf(dr("IdChiNhanh").ToString().Trim() <> "", CType(dr("IdChiNhanh").ToString().Trim(), Integer), 0)
                    lbl_socmt.Text = dr("CMT_So").ToString().Trim()
                    If dr("CMT_NgayCap").ToString() <> "" Then lbl_ngaycap.Text = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_noicap.Text = dr("CMT_NoiCap").ToString().Trim()
                    If (dr("IdNew").ToString().Trim() <> "") Then
                        btn_add.Enabled = False
                        btn_save.Enabled = False
                        lbl_thongbao.Text = "Cán bộ đã chuyển sang lao động chính thức, Không thể thực hiện cập nhật."
                    End If
                End If
            End If
        End If
        'Load dữ liệu danh sách hợp đồng ra lưới
        Fill_Data()

        'Thiết lập quyền được phân theo chức năng của thành viên thao tác chương trình
        '183;184;185;186;
        If (Globals.Roles.IndexOf(";183;") < 0) Then
            dgv_main.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";184;") < 0) Then
            btn_add.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";186;") < 0) Then
            btn_delete.Enabled = False
        End If
    End Sub

    Private Sub cb_nghidinhlg_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_nghidinhlg.SelectedIndexChanged
        cb_bangluong.Items.Clear()
        arr_NgachLuong.Clear()
        If (cb_nghidinhlg.SelectedIndex > 0 And cb_nghidinhlg.Items.Count <> 0) Then
            Dim NghiDinhId As Int32 = CType(arr_NghiDinhLuong(IIf(cb_nghidinhlg.SelectedIndex > 0, cb_nghidinhlg.SelectedIndex, "0")), Int32)
            If (NghiDinhId > 0) Then
                Dim strSQL As String = String.Format("Select IdBangLuong, Mota From BangLuong Where IdND_Luong = {0}", NghiDinhId)
                arr_BangLuong = _Globals.Bind_ComBoBox(cb_bangluong, strSQL, "---Bảng lương---")
            End If
        End If
        cb_bangluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_bangluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_bangluong.SelectedIndexChanged
        cb_ngachluong.Items.Clear()
        arr_NgachLuong.Clear()
        lbl_bangluong.Text = ""
        If (cb_bangluong.SelectedIndex > 0 And cb_bangluong.Items.Count <> 0) Then
            Dim BangLuongId As Int32 = CType(arr_BangLuong(IIf(cb_bangluong.SelectedIndex > 0, cb_bangluong.SelectedIndex, "0")), Int32)
            If (BangLuongId > 0) Then
                Dim strSQL As String = String.Format("Select IdNgachLuong, Mota From NgachLuong Where IdBangLuong = {0}", BangLuongId)
                arr_NgachLuong = _Globals.Bind_ComBoBox(cb_ngachluong, strSQL, "---Ngạch lương---")
                strSQL = String.Format("Select * From BangLuong Where IdBangLuong = {0}", BangLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            lbl_bangluong.Text = db.Rows(0)("BangLuong").ToString().Trim()
                        End If
                    End If
                End Using
            End If
        End If
        cb_ngachluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_ngachluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_ngachluong.SelectedIndexChanged
        cb_bacluong.Items.Clear()
        arr_BacLuong.Clear()
        lbl_ngachluong.Text = ""
        If (cb_ngachluong.SelectedIndex > 0 And cb_ngachluong.Items.Count <> 0) Then
            Dim NgachLuongId As Int32 = CType(arr_NgachLuong(IIf(cb_ngachluong.SelectedIndex > 0, cb_ngachluong.SelectedIndex, "0")), Int32)
            If (NgachLuongId > 0) Then
                Dim strSQL As String = String.Format("Select IdBacLuong, BacLuong From BacLuong Where IdNgachLuong = {0}", NgachLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            cb_bacluong.Items.Add("---Bậc lương---")
                            arr_BacLuong.Add(0)
                            For i As Int32 = 0 To db.Rows.Count - 1
                                cb_bacluong.Items.Add(db.Rows(i)("BacLuong").ToString())
                                arr_BacLuong.Add(IIf(db.Rows(i)("IdBacLuong").ToString() <> "", db.Rows(i)("IdBacLuong").ToString(), ""))
                            Next
                            If (cb_bacluong.Items.Count <> 0) Then
                                cb_bacluong.SelectedIndex = 0
                            End If
                        End If
                    End If
                End Using

                'Load thông tin tên bảng lương ra theo dõi
                strSQL = String.Format("Select * From NgachLuong Where IdNgachLuong = {0}", NgachLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            lbl_ngachluong.Text = db.Rows(0)("NgachLuong").ToString().Trim()
                        End If
                    End If
                End Using

            End If
        End If
        cb_bacluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_bacluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_bacluong.SelectedIndexChanged
        edt_hs_luong.Text = "0"
        If (cb_bacluong.SelectedIndex > 0 And cb_bacluong.Items.Count <> 0) Then
            Dim BacLuongId As Int32 = CType(arr_BacLuong(IIf(cb_bacluong.SelectedIndex > 0, cb_bacluong.SelectedIndex, "0")), Int32)
            If (BacLuongId > 0) Then
                Dim strSQL As String = String.Format("Select * From BacLuong Where Status = 1 and IdBacLuong = {0}", BacLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            edt_hs_luong.Text = IIf(db.Rows(0)("Heso").ToString().Trim() <> "", CType(db.Rows(0)("Heso"), Double).ToString("N2"), "0")
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    Private Sub rb_thangbang_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_thangbang.CheckedChanged
        If (rb_thangbang.Checked = True) Then
            lbl_title_nghidinh.Text = "Nghị định lương"
            lbl_title_nghidinh.Width = 112
            pnl_ttac_1.Visible = True
            cb_nghidinhlg.Visible = True
            cb_bangluong.Visible = True
            lbl_bangluong.Visible = True
            lbl_ttac_div_1.Visible = True
            pnl_ttac_2.Visible = True
            lbl_ttac_div_2.Visible = True
            pnl_ttac_3.Visible = True
            lbl_ttac_div_3.Visible = True
            pnl_ttac_4.Visible = True
            If (cb_nghidinhlg.Items.Count <> 0) Then
                cb_nghidinhlg.SelectedIndex = 0
            End If
            edt_tienluong.Visible = False
            edt_tienluong.Width = 1
        End If
    End Sub

    Private Sub rb_trongoi_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_trongoi.CheckedChanged
        If (rb_trongoi.Checked = True) Then
            edt_tienluong.Visible = True
            lbl_title_nghidinh.Text = "Tiền lương"
            lbl_title_nghidinh.Width = 127
            cb_nghidinhlg.Visible = False
            cb_bangluong.Visible = False
            lbl_bangluong.Visible = False
            edt_tienluong.Width = 180

            lbl_ttac_div_1.Visible = False
            pnl_ttac_2.Visible = False
            lbl_ttac_div_2.Visible = False
            pnl_ttac_3.Visible = False
            lbl_ttac_div_3.Visible = False
            pnl_ttac_4.Visible = False
        End If
    End Sub

    Private Sub rb_hdld_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_DinhBien.CheckedChanged
        If rb_DinhBien.Checked = True Then
            lbl_sohd_soqd.Text = "Số hợp đồng "
            'pnl_baohiem.Visible = True
            cbIsChamDut.Text = "Chấm dứt hợp đồng lao động"
            grpChamDut.Text = "Thông tin chi tiết chấm dứt hợp đồng lao động"
        End If
    End Sub

    Private Sub rb_qdtamtuyen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_qdtamtuyen.CheckedChanged
        If rb_qdtamtuyen.Checked = True Then
            lbl_sohd_soqd.Text = "Số quyết định "
            'pnl_baohiem.Visible = False
            cbIsChamDut.Text = "Chấm dứt quyết định tạm tuyển"
            grpChamDut.Text = "Thông tin chi tiết chấm dứt quyết định tạm tuyển"
        End If
    End Sub

    Private Sub dgv_main_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellClick
        Try
            ResetAll_Controls(False)
            _IdHd = ""
            If (Globals.Roles.IndexOf(";183;") < 0) Then
                Return
            End If
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    FillData_SelectRow(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_main_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub

    'Sự kiện: Thêm mới dữ liệu hợp đồng lao động của cán bộ tập sự
    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        ResetAll_Controls(False)
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
        _IdHd = ""
        strSQL = String.Format("Select * From HSCB_TS_HDLD Where IdCanBo = '{0}' Order By DenNgay Desc", _IdCanBo)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    If (db.Rows(0)("DenNgay").ToString() <> "") Then
                        'Ngày bắt đầu sinh hoạt phải trước ngày vào chính thức ít nhất là 1 năm (1 Năm thủ - thách)
                        Dim dt_ngay_bd As DateTime = CType(db.Rows(0)("DenNgay").ToString(), DateTime).AddDays(1)
                        dtpk_tungay.Value = dt_ngay_bd
                        dtpk_denngay.Value = dt_ngay_bd.AddYears(1)
                    End If
                    'Lấy mặc định chi nhánh/PGD cũ vào
                    cb_chinhanh.SelectedIndex = IIf(db.Rows(0)("ChiNhanhId").ToString() <> "", CType(ARL_ChiNhanh.IndexOf(db.Rows(0)("ChiNhanhId").ToString()), Integer), 0)
                    cb_chinhanh_SelectedIndexChanged(Nothing, Nothing)
                    cb_phongban_pgd.SelectedIndex = IIf(db.Rows(0)("PhongBanDonViId").ToString() <> "", CType(ARL_PhBanDonVi.IndexOf(db.Rows(0)("PhongBanDonViId").ToString()), Integer), 0)
                End If
            End If
        End Using
        'Lấy thông tin có sẵn Người ký và Chức vụ người ký
        edt_nguoiky.Text = "" '_HS_CanBo.GetVarNam("GIAMDOC")
        strSQL = String.Format("Select * from ChiNhanh Where Status = 1 and id = {0}", _IdChiNhanh)
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
                                    cb_chucvu.SelectedIndex = CType(arr_Chucvu.IndexOf(db_child.Rows(0)("id").ToString().Trim()), Integer)
                                End If
                            End If
                        End If
                    End Using
                End If
            End If
        End Using
        ActiveControl = edt_sohd
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Try
            If (IsValid()) Then
                Dim _currRow As String = ""
                If (_IdHd <> "") Then
                    If (Globals.Roles.IndexOf(";185;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hợp đồng lao động của cán bộ tập sự!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        _currRow = _IdHd
                        dgv_main.CurrentRow.Selected = False
                        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Id")).Selected = True
                        ResetAll_Controls(False)
                        ckb_ChoiceAll.Checked = False
                        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                        _IdHd = ""
                        If (Globals.Roles.IndexOf(";183;") < 0) Then
                            Return
                        End If
                        If (dgv_main.Rows.Count > 0) Then
                            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                FillData_SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                End If
                Dim obj_hdld As clsHS_CanBo.HSCB_TS_HDLD
                If (_IdHd = "") Then    'Cập nhật trường hợp Thêm mới
                    obj_hdld = New clsHS_CanBo.HSCB_TS_HDLD()
                    obj_hdld.IdCanbo = _IdCanBo
                    If (rb_DinhBien.Checked = True) Then
                        obj_hdld.Status = 0
                    ElseIf (rb_qdtamtuyen.Checked = True) Then
                        obj_hdld.Status = 1
                    End If
                    obj_hdld.SoQD = Globals.Find_Replace(edt_sohd.Text.ToString().Trim())
                    obj_hdld.NgayQD = dtpk_ngayky.Value
                    obj_hdld.NguoiQD = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_nguoiky.Text.ToString().Trim()))
                    obj_hdld.IdCV_Nguoi_QD = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_chucvu.SelectedIndex), "0"), Int32)
                    obj_hdld.TuNgay = dtpk_tungay.Value
                    obj_hdld.DenNgay = dtpk_denngay.Value
                    obj_hdld.TuGio = CType(edt_gio_bd.Value, DateTime).ToString("HH:mm")
                    obj_hdld.DenGio = CType(edt_gio_kt.Value, DateTime).ToString("HH:mm")
                    obj_hdld.IdHT_TraLuong = CType(IIf(arr_Ht_Traluong.Count > 0, arr_Ht_Traluong(cb_ht_traluong.SelectedIndex), "0"), Int32)

                    If (rb_thangbang.Checked = True) Then
                        If cb_bacluong.Items.Count <> 0 Then
                            obj_hdld.IdBacLuong = CType(IIf(arr_BacLuong.Count > 0, arr_BacLuong(cb_bacluong.SelectedIndex), "0"), Int32)
                        Else : obj_hdld.IdBacLuong = 0
                        End If
                        obj_hdld.HeSo = IIf(edt_hs_luong.Text.ToString() <> "", CType(edt_hs_luong.Text.ToString().Trim(), Double), 0)
                        obj_hdld.TyleHuong = IIf(edt_tylehuong.Text.ToString() <> "", CType(edt_tylehuong.Text.ToString().Trim(), Double), 0)
                    Else
                        obj_hdld.IdBacLuong = 0
                        obj_hdld.HeSo = 0
                        obj_hdld.TyleHuong = 100
                    End If

                    If (rb_trongoi.Checked = True) Then
                        obj_hdld.Tien_Luong = IIf(edt_tienluong.Text.Trim() <> "", CType(MoneyValue(edt_tienluong.Text.Trim()), Double), 0)
                    Else : obj_hdld.Tien_Luong = 0
                    End If
                    obj_hdld.Loai = CType(cb_loaihinh.SelectedIndex, Byte)
                    obj_hdld.CongViec = Globals.Find_Replace(edt_cv_damnhan.Text.Trim.ToString())
                    obj_hdld.IdChuyenMon = CType(IIf(arr_ChuyenMon.Count > 0, arr_ChuyenMon(cb_chuyenmon.SelectedIndex), "0"), Integer)
                    'If (rb_DinhBien.Checked = True) Then
                    '    obj_hdld.BHXH = IIf(ckb_bhxh.Checked = True, 1, 0)
                    '    obj_hdld.BHYT = IIf(ckb_bhyt.Checked = True, 1, 0)
                    'End If
                    obj_hdld.BHXH = IIf(ckb_bhxh.Checked = True, 1, 0)
                    obj_hdld.BHYT = IIf(ckb_bhyt.Checked = True, 1, 0)

                    If cbIsChamDut.Checked Then
                        obj_hdld.TV_Ngay_HL = DateTimeUtil.getDate(dpkNgayHL_TV.Text)
                        obj_hdld.TV_So_QD = txtSoQD_TV.Text.ToString().Trim()
                        obj_hdld.TV_NgayKy_QD = DateTimeUtil.getDate(dpkNgayKy_TV.Text)
                        obj_hdld.TV_IdLyDo = CType(IIf(arrQd_LyDoTV.Count > 0, arrQd_LyDoTV(cboLyDo_TV.SelectedIndex), "0"), Integer)
                        obj_hdld.TV_TroCap_ThoiViec = IIf(txtTroCap_TV.Text.Trim() <> "", CType(MoneyValue(txtTroCap_TV.Text.Trim()), Double), 0)
                        obj_hdld.TV_TroCap_Khac = IIf(txtTroCapKhac.Text.Trim() <> "", CType(MoneyValue(txtTroCapKhac.Text.Trim()), Double), 0)
                        obj_hdld.TV_SoTien_BoiThuong = IIf(txtTienBoiThuong.Text.Trim() <> "", CType(MoneyValue(txtTienBoiThuong.Text.Trim()), Double), 0)
                        obj_hdld.TV_SoTien_ThuHoi = IIf(txtTienThuHoi.Text.Trim() <> "", CType(MoneyValue(txtTienThuHoi.Text.Trim()), Double), 0)
                    Else
                        obj_hdld.TV_Ngay_HL = DateTime.Parse("01/01/1900")
                        obj_hdld.TV_So_QD = ""
                        obj_hdld.TV_NgayKy_QD = DateTime.Parse("01/01/1900")
                        obj_hdld.TV_IdLyDo = 0
                        obj_hdld.TV_TroCap_ThoiViec = 0
                        obj_hdld.TV_TroCap_Khac = 0
                        obj_hdld.TV_SoTien_BoiThuong = 0
                        obj_hdld.TV_SoTien_ThuHoi = 0
                    End If
                    obj_hdld.GhiChu = standardizeString(txtGhiChu.Text.ToString().Trim())
                    obj_hdld.IdCBTS_HDLD = ""
                    obj_hdld.PhongBanDonViId = IIf(ARL_PhBanDonVi.Count > 0, ARL_PhBanDonVi(cb_phongban_pgd.SelectedIndex), "")
                    obj_hdld.ChiNhanhId = CType(IIf(ARL_ChiNhanh.Count > 0, ARL_ChiNhanh(cb_chinhanh.SelectedIndex), "0"), Integer)
                    _currRow = _HS_CanBo.Insert_Update_HSCB_TS_HDLD(obj_hdld)
                Else                    'Cập nhật trường hợp Sửa đổi
                    obj_hdld = New clsHS_CanBo.HSCB_TS_HDLD()
                    obj_hdld.IdCBTS_HDLD = _IdHd
                    obj_hdld.IdCanbo = _IdCanBo
                    If (rb_DinhBien.Checked = True) Then
                        obj_hdld.Status = 0
                    ElseIf (rb_qdtamtuyen.Checked = True) Then
                        obj_hdld.Status = 1
                    End If
                    obj_hdld.SoQD = Globals.Find_Replace(edt_sohd.Text.ToString().Trim())
                    obj_hdld.NgayQD = dtpk_ngayky.Value
                    obj_hdld.NguoiQD = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_nguoiky.Text.ToString().Trim()))
                    obj_hdld.IdCV_Nguoi_QD = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_chucvu.SelectedIndex), "0"), Int32)
                    obj_hdld.TuNgay = dtpk_tungay.Value
                    obj_hdld.DenNgay = dtpk_denngay.Value
                    obj_hdld.TuGio = CType(edt_gio_bd.Value, DateTime).ToString("HH:mm")
                    obj_hdld.DenGio = CType(edt_gio_kt.Value, DateTime).ToString("HH:mm")
                    obj_hdld.IdHT_TraLuong = CType(IIf(arr_Ht_Traluong.Count > 0, arr_Ht_Traluong(cb_ht_traluong.SelectedIndex), "0"), Int32)
                    If (rb_thangbang.Checked = True) Then
                        If cb_bacluong.Items.Count <> 0 Then
                            obj_hdld.IdBacLuong = CType(IIf(arr_BacLuong.Count > 0, arr_BacLuong(cb_bacluong.SelectedIndex), "0"), Int32)
                        Else : obj_hdld.IdBacLuong = 0
                        End If
                        obj_hdld.HeSo = IIf(edt_hs_luong.Text.ToString() <> "", CType(edt_hs_luong.Text.ToString().Trim(), Double), 0)
                        obj_hdld.TyleHuong = IIf(edt_tylehuong.Text.ToString() <> "", CType(edt_tylehuong.Text.ToString().Trim(), Double), 0)
                    Else
                        obj_hdld.IdBacLuong = 0
                        obj_hdld.HeSo = 0
                        obj_hdld.TyleHuong = 0
                    End If

                    If (rb_trongoi.Checked = True) Then
                        obj_hdld.Tien_Luong = IIf(edt_tienluong.Text.Trim() <> "", CType(MoneyValue(edt_tienluong.Text.Trim()), Double), 0)
                    Else : obj_hdld.Tien_Luong = 0
                    End If
                    obj_hdld.Loai = CType(cb_loaihinh.SelectedIndex, Byte)
                    obj_hdld.CongViec = Globals.Find_Replace(edt_cv_damnhan.Text.Trim.ToString())
                    obj_hdld.IdChuyenMon = CType(IIf(arr_ChuyenMon.Count > 0, arr_ChuyenMon(cb_chuyenmon.SelectedIndex), "0"), Integer)
                    'If (rb_DinhBien.Checked = True) Then
                    '    obj_hdld.BHXH = IIf(ckb_bhxh.Checked = True, 1, 0)
                    '    obj_hdld.BHYT = IIf(ckb_bhyt.Checked = True, 1, 0)
                    'End If
                    obj_hdld.BHXH = IIf(ckb_bhxh.Checked = True, 1, 0)
                    obj_hdld.BHYT = IIf(ckb_bhyt.Checked = True, 1, 0)

                    If cbIsChamDut.Checked Then
                        obj_hdld.TV_Ngay_HL = DateTimeUtil.getDate(dpkNgayHL_TV.Text)
                        obj_hdld.TV_So_QD = txtSoQD_TV.Text.ToString().Trim()
                        obj_hdld.TV_NgayKy_QD = DateTimeUtil.getDate(dpkNgayKy_TV.Text)
                        obj_hdld.TV_IdLyDo = CType(IIf(arrQd_LyDoTV.Count > 0, arrQd_LyDoTV(cboLyDo_TV.SelectedIndex), "0"), Integer)
                        obj_hdld.TV_TroCap_ThoiViec = IIf(txtTroCap_TV.Text.Trim() <> "", CType(MoneyValue(txtTroCap_TV.Text.Trim()), Double), 0)
                        obj_hdld.TV_TroCap_Khac = IIf(txtTroCapKhac.Text.Trim() <> "", CType(MoneyValue(txtTroCapKhac.Text.Trim()), Double), 0)
                        obj_hdld.TV_SoTien_BoiThuong = IIf(txtTienBoiThuong.Text.Trim() <> "", CType(MoneyValue(txtTienBoiThuong.Text.Trim()), Double), 0)
                        obj_hdld.TV_SoTien_ThuHoi = IIf(txtTienThuHoi.Text.Trim() <> "", CType(MoneyValue(txtTienThuHoi.Text.Trim()), Double), 0)
                    Else
                        obj_hdld.TV_Ngay_HL = DateTime.Parse("01/01/1900")
                        obj_hdld.TV_So_QD = ""
                        obj_hdld.TV_NgayKy_QD = DateTime.Parse("01/01/1900")
                        obj_hdld.TV_IdLyDo = 0
                        obj_hdld.TV_TroCap_ThoiViec = 0
                        obj_hdld.TV_TroCap_Khac = 0
                        obj_hdld.TV_SoTien_BoiThuong = 0
                        obj_hdld.TV_SoTien_ThuHoi = 0
                    End If
                    obj_hdld.GhiChu = standardizeString(txtGhiChu.Text.ToString().Trim())
                    obj_hdld.PhongBanDonViId = IIf(ARL_PhBanDonVi.Count > 0, ARL_PhBanDonVi(cb_phongban_pgd.SelectedIndex), "")
                    obj_hdld.ChiNhanhId = CType(IIf(ARL_ChiNhanh.Count > 0, ARL_ChiNhanh(cb_chinhanh.SelectedIndex), "0"), Integer)
                    _HS_CanBo.Insert_Update_HSCB_TS_HDLD(obj_hdld)
                    _currRow = _IdHd
                End If
                Fill_Data()
                'Hiển thị lại dòng hiệu thời trước đó đã select hoặc dòng vừa thêm mới
                dgv_main.CurrentRow.Selected = False
                dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Id")).Selected = True
                'Gọi lại sự kiện cell click của lưới dữ liệu
                ResetAll_Controls(False)
                ckb_ChoiceAll.Checked = False
                HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                _IdHd = ""
                If (Globals.Roles.IndexOf(";183;") < 0) Then
                    Return
                End If
                If (dgv_main.Rows.Count > 0) Then
                    If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                        FillData_SelectRow(_currRow)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Cập nhật hợp đồng lao động của cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If (dgv_main.Rows.Count <= 0) Then Return
        Try
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _count As Int16 = 0
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    For i As Int32 = 0 To dgv_main.Rows.Count - 1
                        If (dgv_main.Rows(i).Cells("cln_Id").Value IsNot Nothing) Then
                            If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                _count += 1
                                arr_Del.Add(dgv_main.Rows(i).Cells("cln_Id").Value)
                            End If
                        End If
                    Next
                End If
            End If
            If _count = dgv_main.Rows.Count Then
                MessageBox.Show("Bạn không thể xoá hết hợp đồng lao động hoặc quyết định tạm tuyển của cán bộ!" + vbCrLf + "Lưu ý: Việc xoá hết được chỉ thực hiện khi bạn xoá cán bộ tập sự này.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return
            End If
            If (_count > 0) Then
                Dim _mess As String = IIf(_count = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hợp đồng lao động đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    For i As Int16 = 0 To arr_Del.Count - 1
                        _HS_CanBo.Delete_HSCB_TS_HDLD(arr_Del(i).ToString())
                    Next
                    'Thực hiện Load thông tin sau khi xoá song --> Thành công
                    ResetAll_Controls(False)
                    _IdHd = ""
                    Fill_Data()
                    dgv_main_CellClick(sender, Nothing)
                    MessageBox.Show("Bạn đã xoá thành công thông tin hợp đồng lao động!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Else
                    arr_Del.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            Else
                MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub frmHS_HdldTS_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        'Load dữ liệu danh sách hợp đồng ra lưới
        Dim _currRow As String = _IdHd
        Fill_Data()
        If (dgv_main.Rows.Count <= 0) Then Return
        If (_currRow = "") Then _currRow = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        ResetAll_Controls(False)
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
        dgv_main.CurrentRow.Selected = False
        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Id")).Selected = True
        _IdHd = ""
        If (Globals.Roles.IndexOf(";183;") < 0) Then
            Return
        End If
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                FillData_SelectRow(_currRow)
            End If
        End If
    End Sub
#End Region

#Region "---> Events: Các Sự kiện Ngoại lệ nhập liệu người dùng <---"
    Private Sub edt_tylehuong_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tylehuong.Leave
        If (edt_tylehuong.Text.Trim() = "") Then
            edt_tylehuong.Text = "0"
        End If
    End Sub

    Private Sub edt_tylehuong_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_tylehuong.KeyPress
        Dim decimalString As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
        Dim decimalChar As Char = Convert.ToChar(decimalString)
        If (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
        ElseIf (e.KeyChar = decimalString And edt_tylehuong.Text.IndexOf(decimalString) = -1) Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub edt_sohd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sohd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_sohd.Text.Trim() <> "") Then
                dtpk_ngayky.Focus()
            Else
                edt_sohd.Focus()
            End If
        End If
    End Sub

    Private Sub dtpk_ngayky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngayky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_nguoiky.Focus()
            If (_IdHd = "") Then
                dtpk_tungay.Value = dtpk_ngayky.Value
                dtpk_denngay.Value = dtpk_ngayky.Value.AddYears(1)
            End If
        End If
    End Sub

    Private Sub dtpk_ngayky_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_ngayky.Leave
        If (_IdHd = "") Then
            dtpk_tungay.Value = dtpk_ngayky.Value
            dtpk_denngay.Value = dtpk_ngayky.Value.AddYears(1)
        End If
    End Sub

    Private Sub edt_nguoiky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nguoiky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_chucvu.Focus()
        End If
    End Sub

    Private Sub edt_nguoiky_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nguoiky.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_ngayky.Focus()
        End If
    End Sub

    Private Sub cb_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_loaihinh.Focus()
        End If
    End Sub

    Private Sub cb_loaihinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_loaihinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_ht_traluong.Focus()
        End If
    End Sub

    Private Sub cb_ht_traluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ht_traluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_gio_bd.Focus()
        End If
    End Sub

    Private Sub edt_gio_bd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_gio_kt.Focus()
        End If
    End Sub

    Private Sub edt_gio_kt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_tungay.Focus()
        End If
    End Sub

    Private Sub dtpk_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            rb_thangbang.Focus()
        End If
    End Sub

    Private Sub cb_bangluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_bangluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_ngachluong.Focus()
        End If
    End Sub

    Private Sub cb_ngachluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_ngachluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_bacluong.Focus()
        End If
    End Sub

    Private Sub cb_bacluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_bacluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_tylehuong.Focus()
        End If
    End Sub

    Private Sub edt_hs_luong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hs_luong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hs_luong.Text.Trim() = "") Then
                edt_hs_luong.Text = "0"
            End If
            edt_tylehuong.Focus()
        End If
    End Sub

    Private Sub edt_hs_luong_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hs_luong.Leave
        If (edt_hs_luong.Text.Trim() = "") Then
            edt_hs_luong.Text = "0"
        End If
    End Sub

    Private Sub edt_tylehuong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tylehuong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_tylehuong.Text.Trim() = "") Then
                edt_tylehuong.Text = "0"
            End If
            edt_cv_damnhan.Focus()
        End If
    End Sub

    Private Sub edt_tylehuong_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tylehuong.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_bacluong.Focus()
        End If
    End Sub

    Private Sub edt_cv_damnhan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_save.Focus()
        End If
    End Sub

    Private Sub edt_cv_damnhan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Up) Then
            If (rb_thangbang.Checked = True) Then
                edt_tylehuong.Focus()
            Else
                edt_tienluong.Focus()
            End If
        End If
    End Sub

    Private Sub edt_tienluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tienluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_tienluong.Text.Trim() = "") Then
                edt_tienluong.Text = "0"
            End If
            edt_cv_damnhan.Focus()
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

    Private Sub txtTroCap_TV_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTroCap_TV.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (txtTroCap_TV.Text.Trim() = "") Then
                txtTroCap_TV.Text = "0"
            End If
        End If
    End Sub

    Private Sub txtTroCap_TV_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTroCap_TV.Leave
        If (txtTroCap_TV.Text.Trim() = "") Then
            txtTroCap_TV.Text = "0"
        End If
    End Sub

    Private Sub txtTroCap_TV_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTroCap_TV.TextChanged
        Try
            txtTroCap_TV = formatMoneyinTextbox(txtTroCap_TV)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTroCapKhac_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTroCapKhac.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (txtTroCapKhac.Text.Trim() = "") Then
                txtTroCapKhac.Text = "0"
            End If
        End If
    End Sub

    Private Sub txtTroCapKhac_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTroCapKhac.Leave
        If (txtTroCapKhac.Text.Trim() = "") Then
            txtTroCapKhac.Text = "0"
        End If
    End Sub

    Private Sub txtTroCapKhac_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTroCapKhac.TextChanged
        Try
            txtTroCapKhac = formatMoneyinTextbox(txtTroCapKhac)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTienBoiThuong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTienBoiThuong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (txtTienBoiThuong.Text.Trim() = "") Then
                txtTienBoiThuong.Text = "0"
            End If
        End If
    End Sub

    Private Sub txtTienBoiThuong_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTienBoiThuong.Leave
        If (txtTienBoiThuong.Text.Trim() = "") Then
            txtTienBoiThuong.Text = "0"
        End If
    End Sub

    Private Sub txtTienBoiThuong_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTienBoiThuong.TextChanged
        Try
            txtTienBoiThuong = formatMoneyinTextbox(txtTienBoiThuong)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTienThuHoi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTienThuHoi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (txtTienThuHoi.Text.Trim() = "") Then
                txtTienThuHoi.Text = "0"
            End If
        End If
    End Sub

    Private Sub txtTienThuHoi_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTienThuHoi.Leave
        If (txtTienThuHoi.Text.Trim() = "") Then
            txtTienThuHoi.Text = "0"
        End If
    End Sub

    Private Sub txtTienThuHoi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTienThuHoi.TextChanged
        Try
            txtTienThuHoi = formatMoneyinTextbox(txtTienThuHoi)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cbIsChamDut_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbIsChamDut.CheckedChanged
        If cbIsChamDut.Checked Then
            grpChamDut.Enabled = True
        Else
            grpChamDut.Enabled = False
        End If
    End Sub

#End Region

#Region "---> Events: Các Hàm và sự kiện liên quan đến Check all các items trên lưới dữ liệu <---"
    Private Sub ckb_ChoiceAll_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Space) Then
            HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
        End If
    End Sub

    Private Sub ckb_ChoiceAll_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
    End Sub

    Private Sub dgv_main_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        'Sự kiện này thực hiện khi click các checked items trên lưới dữ liệu thì sẽ Checked hoặc UnChecked Checkall
        If CType(sender, DataGridView).Columns(e.ColumnIndex).Name = "cln_Choice" And e.RowIndex >= 0 Then
            If Not IsHeaderCheckBoxClicked Then
                Dim _vCellCheck As DataGridViewCheckBoxCell
                _vCellCheck = CType(dgv_main("cln_Choice", e.RowIndex), DataGridViewCheckBoxCell)
                RowCheckBoxClick(_vCellCheck)
            End If
        End If
    End Sub

    Private Sub dgv_main_CurrentCellDirtyStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
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

    Private Sub dgv_main_CellPainting(ByVal sender As System.Object, ByVal e As DataGridViewCellPaintingEventArgs)
        'Căn chỉnh cho Checkbox nằm vào giữa tiêu đề của cột Chọn xoá
        If (e.RowIndex = -1 And e.ColumnIndex = 1) Then
            ResetHeaderCheckBoxLocation(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    '-------------- Các hàm liên quan --------------'
    ''' <summary>
    ''' Hàm thực hiện add một CheckBox vào Cột tiêu đề chọn tất cả để xóa các bản ghi đã đánh dấu xóa
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddHeaderCheckBox()
        ckb_ChoiceAll = New CheckBox()
        ckb_ChoiceAll.Size = New Size(15, 15)
        'Add the CheckBox into the DataGridView
        ckb_ChoiceAll.Checked = False
        dgv_main.Controls.Add(ckb_ChoiceAll)
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



    Private Sub cb_chinhanh_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_chinhanh.SelectedIndexChanged
        cb_phongban_pgd.Items.Clear()
        ARL_PhBanDonVi.Clear()
        If (cb_chinhanh.SelectedIndex > 0 And cb_chinhanh.Items.Count <> 0) Then
            Dim _valId As Integer = CType(ARL_ChiNhanh(IIf(cb_chinhanh.SelectedIndex > 0, cb_chinhanh.SelectedIndex, "0")), Integer)
            If _valId > 0 Then
                'Lấy mã chi nhánh từ Id chi nhánh
                Dim strSQL As String = ""
                strSQL = String.Format("SELECT * FROM ChiNhanh WHERE id = {0} AND Status = 1", _valId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim tructhuocCurrent As String = _HS_CanBo.GetTrucThuoc(db.Rows(0)("ma_so").ToString().Trim(), _valId)
                            If db.Rows(0)("ma_so").ToString().Trim() = gMaDonViTW Then
                                strSQL = String.Format("Select ('PB_'+Ltrim(Str(Id))) As Id,Ten_Phong As Ten_Goi From PhongBan Where Charindex('{0}',Truc_Thuoc) > 0 and Status = 1 Order By Ma_So ", tructhuocCurrent.ToString)
                            Else
                                'strSQL = String.Format("Select ZZ.* From (Select ('PB_'+Ltrim(Str(Id))) As Id,Ten_Phong As Ten_Goi From PhongBan Where Charindex('{0}',Truc_Thuoc) > 0 and Status = 1 Union Select ('DV_'+Ltrim(Str(Id))) As Id, Ten_Goi from ChiNhanh Where id_goc = {1} and Status = 1 ) ZZ Order By (Case When Substring(ZZ.Id,1,2) = 'PB' Then 0 Else 1 End),ZZ. Id", tructhuocCurrent.ToString, _valId)
                                strSQL = ""
                                strSQL = strSQL & String.Format("Select Id,Ten_Goi,Ma_So From ")
                                strSQL = strSQL & String.Format("       ( ")
                                strSQL = strSQL & String.Format("        Select ('PB_'+Ltrim(Str(Id))) As Id, Ten_Phong As Ten_Goi,1 STT,Ma_So From PhongBan Where Charindex('{0}',Truc_Thuoc) > 0 And Status = 1 ", tructhuocCurrent.ToString())
                                strSQL = strSQL & String.Format("        Union All ")
                                strSQL = strSQL & String.Format("        Select ('DV_'+Ltrim(Str(Id))) As Id, Ma_So + N' - ' + Ten_Goi,2 STT,Ma_So From ChiNhanh Where Id_Goc = {0} And Status = 1 ", _valId)
                                strSQL = strSQL & String.Format("       ) ZZ Order By ZZ.STT,ZZ.Ma_So,ZZ.Ten_Goi")
                            End If
                            ARL_PhBanDonVi = _Globals.Bind_ComBoBox(cb_phongban_pgd, strSQL, "---Phòng ban/Phòng giao dịch---")
                        End If
                    End If
                End Using
            End If
        End If
    End Sub
End Class