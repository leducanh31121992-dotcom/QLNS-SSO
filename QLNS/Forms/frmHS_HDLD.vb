Public Class frmHS_HDLD

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Globals As Globals = New Globals
    Private _PersonnelFile As clsHS_CanBo = New clsHS_CanBo()
    Private _Labour As clsHS_Hdld = New clsHS_Hdld()
    Private _ListDocument As clsHT_DanhMuc = New clsHT_DanhMuc()
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _IdCanBo As String = ""
    'Mảng lưu id danh sách các Loại hợp đồng
    Private arr_LoaiHd As ArrayList = New ArrayList()
    'Mảng lưu id danh sách các chức vụ ký Hợp đồng
    Private arr_Chucvu_Hd As ArrayList = New ArrayList()
    'Mảng lưu id danh sách các hình thức trả lương
    Private arr_HTTL As ArrayList = New ArrayList()
    'Mảng lưu id danh sách các Nghi dinh lương
    Private arr_NghiDinhLuong As ArrayList = New ArrayList()
    'Mảng lưu id danh sách các Bảng lương
    Private arr_BangLuong As ArrayList = New ArrayList()
    'Mảng lưu id danh sách Ngạch lương theo bảng lương
    Private arr_NgachLuong As ArrayList = New ArrayList()
    'Mảng lưu id danh sách các Bậc lương theo Ngạch lương
    Private arr_BacLuong As ArrayList = New ArrayList()

    'Biến lưu lại Node hiện thời đang Select trên cây dữ liệu
    Private strNode As String = ""
    'Biến lưu lại Id hiện thời của hồ sơ Hợp đồng lao động đang Select trên lưới dữ liệu
    Private _RowId As String = ""
    Private _CodeLoaiHD As String = ""  'Biến lưu lại mã hiệu của loại hợp đồng đã chọn
    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

    'Khai báo biến sử dụng khi được gọi từ Hồ sơ cán bộ
    Public HumanId As String = ""       'Giá trị id của cán bộ
    Public FlagShow As Boolean = False 'Giá trị True là được gọi từ Hồ sơ cán bộ. False: Giá trị mặc định
    Private _IdDonviHT As Integer = IdDONVI
    Public TagNode As String = ""       'Giá trị lưu lại tag đơn vị đang select bên hồ sơ cán bộ
    'Khai báo các thông tin add CheckBox vào cột tiêu đề chọn cả các items Check trên lưới
    Private ckb_ChoiceAll As CheckBox = Nothing   'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0          'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0   'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False  'Cờ báo việc Checkall
#End Region

#Region "---> Functions: Các hàm dùng chung - Quản lý hợp đồng lao động <---"
    ''' <summary>
    ''' Hàm thực hiện - Thiết lập lại trạng thái ban đầu cho tất các điều khiển
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls()
        _RowId = ""
        edt_hd_sohd.Text = ""
        cb_hd_loaihd.SelectedIndex = 0
        dtpk_hd_ngayhl.Text = DateTime.Now.ToShortDateString()
        dtpk_hd_ngayky.Text = DateTime.Now.ToShortDateString()
        edt_hd_nguoiky.Text = ""
        dtpk_hd_tungay.Text = DateTime.Now.ToShortDateString()
        dtpk_hd_denngay.Text = DateTime.Now.ToShortDateString()
        dtpk_hd_denngay.Checked = False
        cb_hd_loaihd_SelectedIndexChanged(Nothing, Nothing)
        cb_hd_chucvu.SelectedIndex = 0
        cb_hd_httraluong.SelectedIndex = 0
        cb_hd_nghidinh.SelectedIndex = 0
        cb_hd_nghidinh_SelectedIndexChanged(Nothing, Nothing)
        edt_hd_heso.Text = ""
        edt_hd_tyle.Text = "0"
        edt_hd_ghichu.Text = ""
    End Sub

    ''' <summary>
    ''' Hàm kiểm tra xem cán bộ cần thêm hợp đồng lao động đã ký hợp đồng không xác định thời hạn --> thì không cho thêm tiếp
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsAdd() As Boolean
        Dim strSQL As String = ""
        strSQL = String.Format("SELECT * FROM HS_HDLD WHERE IdCanBo = '{0}' And IdLoaiHD = (SELECT id FROM DanhMuc WHERE id_goc = 30 and ma_so = '3001')", _IdCanBo)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    MessageBox.Show(String.Format("Cán bộ này đã ký hợp đồng không xác định thời hạn" + vbCrLf + " Không thể thêm tiếp hồ sơ hợp đồng lao động cho cán bộ!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    Return False
                End If
            End If
        End Using
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện - Kiểm tra điều khiện hợp lệ cập nhật dữ liệu
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Dim strSQL As String = ""
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ cần cập nhật hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = tv_main
            Return False
        End If
        If (edt_hd_sohd.Text.Trim() = "") Then
            MessageBox.Show("Số hợp đồng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_hd_sohd
            Return False
        End If
        'Bắt trùng số hợp đồng với một cán bộ
        If (edt_hd_sohd.Text.Trim() <> "") Then
            If (_RowId <> "") Then
                strSQL = String.Format("Select * From HS_HDLD Where IdCanBo = '{0}' And IdCB_HDLD <>'{1}' And SoHD = '{2}'", _IdCanBo, _RowId, Globals.Find_Replace(edt_hd_sohd.Text.Trim().ToString()))
            Else
                strSQL = String.Format("Select * From HS_HDLD Where IdCanBo = '{0}' And SoHD = '{1}'", _IdCanBo, Globals.Find_Replace(edt_hd_sohd.Text.Trim().ToString()))
            End If
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        MessageBox.Show(String.Format("Số hợp đồng '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_hd_sohd.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_hd_sohd
                        Return False
                    End If
                End If
            End Using
        End If
        If (cb_hd_loaihd.SelectedIndex <= 0 And cb_hd_loaihd.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn loại hợp đồng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_hd_loaihd
            Return False
        End If
        If (txtDVKyHD.Text.Trim() = "") Then
            MessageBox.Show("Bạn chưa nhập đơn vị thực hiện kí kết hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_hd_ghichu
            Return False
        End If
        If (txtNoiLamViec.Text.Trim() = "") Then
            MessageBox.Show("Bạn chưa nhập nơi cán bộ làm việc thực hiện kí kết hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_hd_ghichu
            Return False
        End If
        If (edt_hd_nguoiky.Text.Trim() = "") Then
            MessageBox.Show("Người ký quyết định hợp đồng lao động không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_hd_nguoiky
            Return False
        End If
        'If (cb_hd_chucvu.SelectedIndex <= 0 And cb_hd_chucvu.Items.Count <> 0) Then
        '    MessageBox.Show("Bạn chưa chọn chức vụ người ký quyết định!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '    ActiveControl = cb_hd_chucvu
        '    Return False
        'End If
        'If (cb_hd_httraluong.SelectedIndex <= 0 And cb_hd_httraluong.Items.Count <> 0) Then
        '    MessageBox.Show("Bạn chưa chọn hình thức trả lương trong hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '    ActiveControl = cb_hd_httraluong
        '    Return False
        'End If
        'If (cb_hd_bacluong.SelectedIndex <= 0 And cb_hd_bacluong.Items.Count <> 0) Then
        '    MessageBox.Show("Bạn chưa chọn bậc lương trong hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '    ActiveControl = cb_hd_bacluong
        '    Return False
        'End If
        'If (edt_hd_tyle.Text.Trim() = "" Or edt_hd_tyle.Text.Trim() = "0") Then
        '    MessageBox.Show("Tỷ lệ lương được hưởng không thể để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '    ActiveControl = edt_hd_tyle
        '    Return False
        'End If
        If Not (_CodeLoaiHD = "3001" Or _CodeLoaiHD = "3006") Then 'Trường hợp Hợp đồng ngắn hạn
            If (dtpk_hd_denngay.Checked = False) Then
                MessageBox.Show("Bạn chưa chọn ngày kết thúc hợp đồng lao động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = dtpk_hd_denngay
                Return False
            End If
        End If
        If (dtpk_hd_denngay.Checked = True) Then
            If dtpk_hd_tungay.Value >= dtpk_hd_denngay.Value Then
                MessageBox.Show("Ngày bắt đầu không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = dtpk_hd_denngay
                Return False
            End If
            ''Bắt các trường hợp xen kẽ ngày tháng khi loại hợp đồng khác với hợp đồng không xác định thời hạn
            'If (_RowId = "") Then
            '    strSQL = "SELECT * FROM HS_HDLD WHERE IdCanBo = '" + _IdCanBo + "' And Not (DenNgay Is Null Or DenNgay = '' Or DenNgay = '1900/01/01') Order By TuNgay Asc"
            'ElseIf (_RowId <> "") Then
            '    strSQL = "SELECT * FROM HS_HDLD WHERE IdCanBo = '" + _IdCanBo + "' And IdCB_HDLD <> '" + _RowId + "' And Not (DenNgay Is Null Or DenNgay = '' Or DenNgay = '1900/01/01') Order By TuNgay Asc"
            'End If
            'Dim Exist As Boolean = False
            'Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            '    If Not (db Is Nothing) Then
            '        If (db.Rows.Count > 0) Then
            '            For i As Integer = 0 To db.Rows.Count - 1
            '                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
            '                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
            '                If ((_TuNgay <= dtpk_hd_tungay.Value And dtpk_hd_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_hd_denngay.Value And dtpk_hd_denngay.Value <= _DenNgay) Or (dtpk_hd_tungay.Value <= _TuNgay And dtpk_hd_denngay.Value >= _DenNgay)) Then
            '                    Exist = True
            '                    Exit For
            '                End If
            '            Next
            '            If (Exist = True) Then
            '                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hợp đồng lao động không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau giữa các hợp đồng của cán bộ.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            '                ActiveControl = dtpk_hd_tungay
            '                Return False
            '            End If
            '        End If
            '    End If
            'End Using
        End If
        'Bắt trường hợp nếu một cán bộ đã có hợp đồng ngắn hạn trước đó. Thêm tiếp HD không xác định thời hạn mà ngày bắt đầu đan xen Hợp đồng ngắn hạn
        'If (_CodeLoaiHD = "3001") Then 'Trường hợp Hợp đồng ngắn hạn
        '    If (_RowId = "") Then
        '        strSQL = "SELECT * FROM HS_HDLD WHERE IdCanBo = '" + _IdCanBo + "' Order by TuNgay Asc"
        '    Else
        '        strSQL = "SELECT * FROM HS_HDLD WHERE IdCanBo = '" + _IdCanBo + "' And IdCB_HDLD <> '" + _RowId + "' Order by TuNgay Asc"
        '    End If
        '    Dim _ExistTime As Boolean = False
        '    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
        '        If Not (db Is Nothing) Then
        '            If (db.Rows.Count > 0) Then
        '                For i As Integer = 0 To db.Rows.Count - 1
        '                    Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
        '                    Dim _DenNgay As DateTime = DateTime.Now
        '                    If (db.Rows(i)("DenNgay").ToString() <> "") Then
        '                        If CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900" Then
        '                            _DenNgay = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
        '                        End If
        '                    End If
        '                    If ((_TuNgay <= dtpk_hd_tungay.Value And dtpk_hd_tungay.Value <= _DenNgay) Or (dtpk_hd_tungay.Value <= _TuNgay)) Then
        '                        _ExistTime = True
        '                        Exit For
        '                    End If
        '                Next
        '                If (_ExistTime = True) Then
        '                    MessageBox.Show("Thời gian bắt đầu của hợp đồng không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau giữa các hợp đồng của cán bộ.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '                    ActiveControl = dtpk_hd_tungay
        '                    Return False
        '                End If
        '            End If
        '        End If
        '    End Using
        'End If
        Return True
    End Function

    ''' <summary>
    ''' Hàm fill dữ liệu vào controls khi người dùng click vào row truyền vào
    ''' </summary>
    ''' <param name="_CodeId">Chỉ số xác định bản ghi - Bản ghi đang select</param>
    ''' <remarks></remarks>
    Private Sub FillData_SelectRow(ByVal _CodeId As String)
        Dim dr As DataRow
        _RowId = _CodeId
        dr = _Labour.GetLaborContract(_RowId)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                edt_hd_sohd.Text = dr("SoHD").ToString().Trim()
                cb_hd_loaihd.SelectedIndex = CType(arr_LoaiHd.IndexOf(dr("IdLoaiHD").ToString()), Integer)
                If dr("Ngay_HL").ToString().Trim() <> "" Then
                    dtpk_hd_ngayhl.Value = CType(dr("Ngay_HL").ToString(), DateTime)
                End If

                If dr("NgayKy_HD").ToString().Trim() <> "" Then
                    dtpk_hd_ngayky.Value = CType(dr("NgayKy_HD").ToString(), DateTime)
                End If
                edt_hd_nguoiky.Text = dr("NguoKy_QD").ToString().Trim()
                cb_hd_chucvu.SelectedIndex = CType(arr_Chucvu_Hd.IndexOf(dr("IdCV_NguoiKy_QD").ToString()), Integer)

                If dr("TuNgay").ToString().Trim() <> "" Then
                    dtpk_hd_tungay.Value = CType(dr("TuNgay").ToString(), DateTime)
                End If
                If (dr("DenNgay").ToString().Trim() <> "") Then
                    If (CType(dr("DenNgay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                        dtpk_hd_denngay.Checked = False
                    Else
                        dtpk_hd_denngay.Checked = True
                        dtpk_hd_denngay.Value = CType(dr("DenNgay").ToString().Trim(), DateTime)
                    End If
                Else
                    dtpk_hd_denngay.Checked = False
                End If
                cb_hd_httraluong.SelectedIndex = My_CInt(IIf(dr("IdHT_TraLuong").ToString().Trim() <> "", arr_HTTL.IndexOf(dr("IdHT_TraLuong").ToString), 0), -1)
                edt_hd_heso.Text = dr("HeSo").ToString().Trim()
                edt_hd_tyle.Text = IIf(dr("TyleHuong").ToString().Trim() <> "", CType(dr("TyleHuong"), Double).ToString("N2"), "").ToString
                txtDVKyHD.Text = dr("DVKyHDLD").ToString()
                txtNoiLamViec.Text = dr("NoiLamViec").ToString()
                edt_hd_ghichu.Text = dr("GhiChu").ToString()
                'Load bậc lương - Và từ bậc lương fill Ngạch lương và Bảng lương
                Dim _NghiDinhLuongId As Integer = 0
                Dim _BangLuongId As Integer = 0
                Dim _NgachLuongId As Integer = 0
                Dim _BacLuongId As Integer = 0
                _BacLuongId = My_CInt(IIf(dr("IdBacLuong").ToString() <> "0", CType(dr("IdBacLuong").ToString(), Integer), 0), 0) ' Id bậc lương
                If (_BacLuongId > 0) Then
                    _NgachLuongId = _Labour.GetIdNgachLuong(_BacLuongId)
                    If (_NgachLuongId > 0) Then
                        _BangLuongId = _Labour.GetIdBangLuong(_NgachLuongId)
                    End If
                    If (_BangLuongId > 0) Then
                        _NghiDinhLuongId = _Labour.GetIdNghiDinhLuong(_BangLuongId)
                    End If
                End If
                cb_hd_nghidinh.SelectedIndex = CType(arr_NghiDinhLuong.IndexOf(_NghiDinhLuongId.ToString()), Integer)
                cb_hd_nghidinh_SelectedIndexChanged(Nothing, Nothing)
                cb_hd_bangluong.SelectedIndex = CType(arr_BangLuong.IndexOf(_BangLuongId.ToString()), Integer)
                cb_hd_bangluong_SelectedIndexChanged(Nothing, Nothing)
                cb_hd_ngachluong.SelectedIndex = CType(arr_NgachLuong.IndexOf(_NgachLuongId.ToString()), Integer)
                cb_hd_ngachluong_SelectedIndexChanged(Nothing, Nothing)
                cb_hd_bacluong.SelectedIndex = CType(arr_BacLuong.IndexOf(dr("IdBacLuong").ToString()), Integer)
                cb_hd_bacluong_SelectedIndexChanged(Nothing, Nothing)
            End If
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện cần dùng <---"
    Private Sub frmHS_HDLD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Labour.Create_Frame(dgv_hdld, True)
        AddHeaderCheckBox()
        'Gọi một số sự kiện liên quan đến Check all items trên lưới dữ liệu
        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_hdld.CellValueChanged, AddressOf dgv_hdld_CellValueChanged
        AddHandler dgv_hdld.CellPainting, AddressOf dgv_hdld_CellPainting
        AddHandler dgv_hdld.CurrentCellDirtyStateChanged, AddressOf dgv_hdld_CurrentCellDirtyStateChanged

        _PersonnelFile.Fill_Tree(tv_main)

        'Load combobox - Loại hợp đồng
        arr_LoaiHd.Clear()
        cb_hd_loaihd.Items.Clear()
        arr_LoaiHd = _Globals.Bind_ComBoBox(cb_hd_loaihd, clsHT_DanhMuc.Sql_LoaiHd, "---Loại hợp đồng---")

        'Load combobox - Chức vụ người ký hợp đồng
        arr_Chucvu_Hd.Clear()
        cb_hd_chucvu.Items.Clear()
        arr_Chucvu_Hd = _Globals.Bind_ComBoBox(cb_hd_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ người ký QĐ---")

        'Load combobox - Hình thức trả lương cho cán bộ ghi trong Hợp đồng lao động
        Dim strSQL As String = "Select id,ten_goi from DanhMuc Where id_goc != 0 and id_goc = 40 and Status = 1 and ten_goi != N'Khoán'"

        arr_HTTL.Clear()
        cb_hd_httraluong.Items.Clear()
        arr_HTTL = _Globals.Bind_ComBoBox(cb_hd_httraluong, strSQL, "---Hình thức trả lương---")

        'Load combobox - Nghi dinh lương
        strSQL = "Select IdNDLuong, TenND From NghiDinhLuong Where Status = 1"
        arr_NghiDinhLuong.Clear()
        cb_hd_nghidinh.Items.Clear()
        arr_NghiDinhLuong = _Globals.Bind_ComBoBox(cb_hd_nghidinh, strSQL, "---Nghị định lương---")

        ResetAll_Controls()
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_hdld)
        lkl_xemct.Enabled = False
        tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
        strNode = ""

        'Kiểm tra quyền của thành viên được phép thao tác chương trình
        Dim _roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";135;136;137;138;", _roles)) Then
            Close()
        End If
        If (Globals.Roles.IndexOf(";136;") < 0) Then    ' Xét quyền Thêm mới hợp đồng lao động
            btn_addnew.Visible = False
        End If
        If (Globals.Roles.IndexOf(";138;") < 0) Then    ' Xét quyền Xoá bỏ Hợp đồng lao động
            btn_xoadl.Visible = False
            dgv_hdld.Columns("cln_Choice").Visible = False
            ckb_ChoiceAll.Visible = False
        End If
        If (Globals.Roles.IndexOf(";136;") < 0 And Globals.Roles.IndexOf(";137;") < 0) Then
            btn_chapnhan.Visible = False
        End If
        dgv_hdld.Rows.Clear()

        'Nếu được gọi từ Hồ sơ cán bộ
        If (FlagShow = True) Then
            Dim _DonviId As Integer = 0         'Id đơn vị của cán bộ
            Dim _CodeBranch As String = ""      'Mã hiệu của chi nhánh cap tinh hoac tuong duong
            Dim _PhongBanId As Integer = 0      'Id phòng ban của cán bộ
            If (TagNode <> "") Then
                Dim arrElement() As String = TagNode.Split("_")
                If (arrElement.Length > 0) Then
                    If TagNode.Substring(TagNode.ToString.Length - 2, 2) <> "00" Then   'Trường hợp đơn vị là PGD
                        'Tim Id don vi cua can bo dua vao id PGD
                        _DonviId = _ListDocument.GetRootId(CType(arrElement(1).ToString().Trim(), Integer))
                    Else                                    'Trường hợp đơn vị là tỉnh hoặc tương đương
                        _DonviId = CType(arrElement(1).ToString().Trim(), Integer)
                    End If
                End If
            End If
            If (_DonviId > 0) Then
                _CodeBranch = _ListDocument.GetCodeForId(_DonviId)
                _PhongBanId = _PersonnelFile.GetDebtId(HumanId, _DonviId)
            End If
            Dim _NodeFind As String = ""
            Dim _node As TreeNode = Nothing
            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
            'Nếu là Cài đặt tại Hội sở
            If (DONVI = gMaDonViTW) Then
                _NodeFind = "ROOT_1_" & gMaDonViTW
                'Thực hiện tìm Cấp 0
                _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_node Is Nothing) Then
                    tv_main.SelectedNode = _node
                End If
                _NodeFind = "DV_" + _DonviId.ToString() + "_" + _CodeBranch
            Else
                _NodeFind = "ROOT_" + _DonviId.ToString() + "_" + _CodeBranch
            End If

            'Tìm cấp 1 (Gốc root: Ví dụ ROOT_3_10500)
            If (_NodeFind <> "") Then
                _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_node Is Nothing) Then
                    tv_main.SelectedNode = _node
                End If
                'Kiem tra xem Phong Ban co cua Huyen Khong?
                If TagNode.Substring(TagNode.ToString.Length - 2, 2) <> "00" Then
                    _node = Globals.TreeViewFindNode(tv_main.Nodes, TagNode)
                    If Not (_node Is Nothing) Then
                        tv_main.SelectedNode = _node
                    End If
                    'Tim den Phong ban cua PGD
                    _NodeFind = "PB_" + TagNode.Substring(TagNode.IndexOf("_") + 1, TagNode.LastIndexOf("_") - TagNode.IndexOf("_") - 1) + "_" + _PhongBanId.ToString()
                Else
                    'Tìm đến cấp 2 (Ví dụ: PB_3_25)
                    _NodeFind = "PB_" + _DonviId.ToString() + "_" + _PhongBanId.ToString()
                End If
                If (_NodeFind <> "") Then
                    _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                    If Not (_node Is Nothing) Then
                        tv_main.SelectedNode = _node
                    End If
                    'Tìm đến cấp 3 (Ví dụ: CB_CNTT00000000004)
                    _NodeFind = "CB_" + HumanId.ToString()
                    If (_NodeFind <> "") Then
                        _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                        If Not (_node Is Nothing) Then
                            tv_main.SelectedNode = _node
                        End If
                    End If
                End If
            End If
        End If

        If HumanId <> "" Then
            If checkRight_CreateRecord(HumanId) Then
                btn_chapnhan.Enabled = True
                btn_xoadl.Enabled = True
            Else
                btn_chapnhan.Enabled = False
                btn_xoadl.Enabled = False
            End If
        End If

    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        Try
            ckb_ChoiceAll.Checked = False
            dgv_hdld.Rows.Clear()
            lkl_xemct.Enabled = False
            If (tv_main.Nodes.Count > 0) Then
                'Thực hiện Load child node khi click vào parent node 
                Dim arrElement() As String
                If tv_main.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not (tv_main.SelectedNode.IsExpanded) Then
                        arrElement = tv_main.SelectedNode.Tag.ToString().Trim().Split("_")
                        If (arrElement.Length > 0) Then
                            If (arrElement(0) = "DV") Then
                                _PersonnelFile.Fill_Node(tv_main.SelectedNode, arrElement(1), arrElement(2))
                            ElseIf arrElement(0) = "PB" Then
                                _PersonnelFile.Fill_NodeCanbo(tv_main.SelectedNode, arrElement(1), arrElement(2), True)
                            End If
                        End If
                    End If
                End If
                'Thực hiện load dữ liệu cán bộ khi click vào Từng cán bộ
                strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                'Lấy Chỉ số xác định đơn vị Hiện tại của cán bộ (Xét với trường hợp cài đặt đơn vị '000100')
                If (DONVI = gMaDonViTW) Then
                    If (strNode.IndexOf("DV_") >= 0) Then
                        _IdDonviHT = CType(strNode.Substring(3, 1), Integer)
                    End If
                End If
                'Thông tin chung về cán bộ - được Select
                lbl_macb.Text = ""
                lbl_hoten.Text = ""
                lbl_gioitinh.Text = ""
                lbl_ngaysinh.Text = ""
                lbl_so_cmt.Text = ""
                _IdCanBo = ""
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "CB") Then
                    strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                    _IdCanBo = tv_main.SelectedNode.Tag.ToString().Trim().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1)
                    If (_IdCanBo <> "") Then
                        'Fill data - thông tin chung về nhân sự
                        If (Globals.Roles.IndexOf(";71;") < 0) Then
                            lkl_xemct.Enabled = False
                        Else
                            lkl_xemct.Enabled = True
                        End If
                        Dim dr As DataRow
                        dr = _PersonnelFile.GetHuman_ForCode(_IdCanBo)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                lbl_macb.Text = dr("MaCB").ToString().Trim()
                                lbl_hoten.Text = dr("HoTen").ToString().Trim()
                                lbl_gioitinh.Text = IIf(dr("GioiTinh").ToString() = False, "Nam", "Nữ")
                                If dr("NgaySinh").ToString().Trim() <> "" Then
                                    lbl_ngaysinh.Text = IIf(Format(CType(dr("NgaySinh"), DateTime), "dd-MM-yyyy") <> "01-01-1900", Format(CType(dr("NgaySinh"), DateTime), "dd-MM-yyyy"), "")
                                End If
                                lbl_so_cmt.Text = dr("CMT_So").ToString().Trim()
                                If (dr("IdNew").ToString().Trim() <> "") Then  'Nếu cán bộ đã chuyển sang loại Ngắn hạn rồi
                                    btn_addnew.Visible = False
                                    btn_chapnhan.Visible = False
                                Else
                                    btn_chapnhan.Visible = True
                                    btn_addnew.Visible = True
                                End If
                            End If
                        End If
                        'Thực hiện fill data HS Hợp đồng lao động của cán bộ ra lưới dữ liệu

                        Using db As DataTable = _Labour.GetAllLaborContract(_IdCanBo)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_hdld.Rows.Add()
                                        dgv_hdld.Rows(i).Cells("cln_Code").Value = db.Rows(i)("IdCB_HDLD").ToString().Trim()
                                        dgv_hdld.Rows(i).Cells("cln_SoHd").Value = db.Rows(i)("SoHD").ToString().Trim()
                                        Dim strSQL As String = ""
                                        'Load tên gọi - Loại hợp đồng
                                        If (db.Rows(i)("IdLoaiHD").ToString() <> "" And db.Rows(i)("IdLoaiHD").ToString() <> "0") Then
                                            strSQL = String.Format("Select id, ten_goi From DanhMuc Where id = {0} and id_goc = 30 and Status = 1", CType(db.Rows(i)("IdLoaiHD").ToString(), Integer))
                                            dgv_hdld.Rows(i).Cells("cln_LoaiHd").Value = _PersonnelFile.GetNameByCode(strSQL)
                                        Else
                                            dgv_hdld.Rows(i).Cells("cln_LoaiHd").Value = ""
                                        End If
                                        dgv_hdld.Rows(i).Cells("cln_NgayHL").Value = IIf(db.Rows(i)("Ngay_HL").ToString().Trim() <> "", Format(CType(db.Rows(i)("Ngay_HL"), DateTime), "dd-MM-yyyy"), "")
                                        dgv_hdld.Rows(i).Cells("cln_Ngayky").Value = IIf(db.Rows(i)("NgayKy_HD").ToString().Trim() <> "", Format(CType(db.Rows(i)("NgayKy_HD"), DateTime), "dd-MM-yyyy"), "")
                                        dgv_hdld.Rows(i).Cells("cln_Nguoiky").Value = db.Rows(i)("NguoKy_QD").ToString().Trim()
                                        'Load tên gọi - Chức vụ người ký quyết định
                                        If (db.Rows(i)("IdCV_NguoiKy_QD").ToString() <> "" And db.Rows(i)("IdCV_NguoiKy_QD").ToString() <> "0") Then
                                            strSQL = String.Format("Select id, ten_goi From DanhMuc Where id = {0} and id_goc = 14 and Status = 1", CType(db.Rows(i)("IdCV_NguoiKy_QD").ToString(), Integer))
                                            dgv_hdld.Rows(i).Cells("cln_Chucvu").Value = _PersonnelFile.GetNameByCode(strSQL)
                                        Else
                                            dgv_hdld.Rows(i).Cells("cln_Chucvu").Value = ""
                                        End If
                                        dgv_hdld.Rows(i).Cells("cln_Tungay").Value = IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", Format(CType(db.Rows(i)("TuNgay"), DateTime), "dd-MM-yyyy"), "")
                                        If (db.Rows(i)("DenNgay").ToString().Trim() <> "") Then
                                            dgv_hdld.Rows(i).Cells("cln_Denngay").Value = Format(CType(db.Rows(i)("DenNgay"), DateTime), "dd-MM-yyyy")
                                        Else
                                            dgv_hdld.Rows(i).Cells("cln_Denngay").Value = ""
                                        End If

                                        'Load tên gọi - Hình thức trả lương
                                        If (db.Rows(i)("IdHT_TraLuong").ToString() <> "" And db.Rows(i)("IdHT_TraLuong").ToString() <> "0") Then
                                            strSQL = String.Format("Select id, ten_goi From DanhMuc Where id = {0} and id_goc = 40 and Status = 1", CType(db.Rows(i)("IdHT_TraLuong").ToString(), Integer))
                                            dgv_hdld.Rows(i).Cells("cln_HTTL").Value = _PersonnelFile.GetNameByCode(strSQL)
                                        Else
                                            dgv_hdld.Rows(i).Cells("cln_HTTL").Value = ""
                                        End If
                                        'Load tên gọi - Bậc lương
                                        If (db.Rows(i)("IdBacLuong").ToString() <> "" And db.Rows(i)("IdBacLuong").ToString() <> "0") Then
                                            strSQL = String.Format("Select IdBacLuong, Mota From BacLuong Where IdBacLuong = {0} and Status = 1", CType(db.Rows(i)("IdBacLuong").ToString(), Integer))
                                            dgv_hdld.Rows(i).Cells("cln_Bacluong").Value = _PersonnelFile.GetNameByCode(strSQL)
                                        Else
                                            dgv_hdld.Rows(i).Cells("cln_Bacluong").Value = ""
                                        End If
                                        dgv_hdld.Rows(i).Cells("cln_Heso").Value = IIf(db.Rows(i)("HeSo").ToString() <> "", db.Rows(i)("HeSo").ToString(), "")
                                        dgv_hdld.Rows(i).Cells("cln_Tyle").Value = IIf(db.Rows(i)("TyleHuong").ToString() <> "", db.Rows(i)("TyleHuong").ToString(), "")
                                        dgv_hdld.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                                    Next
                                End If
                            End If
                        End Using

                        If _IdCanBo <> "" Then
                            If checkRight_CreateRecord(_IdCanBo) Then
                                btn_chapnhan.Enabled = True
                                btn_xoadl.Enabled = True
                            Else
                                btn_chapnhan.Enabled = False
                                btn_xoadl.Enabled = False
                            End If
                        End If

                    End If
                End If
                dgv_hdld_CellClick(sender, Nothing)
                tv_main.SelectedNode.Expand()
            End If
            TotalCheckBoxes = dgv_hdld.RowCount
            TotalCheckedCheckBoxes = 0
        Catch ex As Exception
            MessageBox.Show("Hiển thị dữ liệu khi click vào cây dữ liệu: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_hdld_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hdld.CellClick
        Try
            ResetAll_Controls()
            If (dgv_hdld.Rows.Count > 0) Then
                If ((dgv_hdld.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    FillData_SelectRow(dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_hdld_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hdld.KeyUp
        dgv_hdld_CellClick(sender, Nothing)
    End Sub

    Private Sub cb_hd_nghidinh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_hd_nghidinh.SelectedIndexChanged
        cb_hd_bangluong.Items.Clear()
        arr_BangLuong.Clear()
        lbl_bl_code.Text = ""
        If (cb_hd_nghidinh.SelectedIndex > 0 And cb_hd_nghidinh.Items.Count <> 0) Then
            Dim NghiDinhId As Integer = CType(arr_NghiDinhLuong(IIf(cb_hd_nghidinh.SelectedIndex > 0, cb_hd_nghidinh.SelectedIndex, "0")), Integer)
            Dim strSQL As String = ""

            If (NghiDinhId > 0) Then
                strSQL = String.Format("Select IdBangLuong, Mota From BangLuong Where IdND_Luong = {0}", NghiDinhId)
                arr_BangLuong = _Globals.Bind_ComBoBox(cb_hd_bangluong, strSQL, "---Bảng lương---")
            End If

            strSQL = String.Format("Select * From NghiDinhLuong Where IdNDLuong = {0}", NghiDinhId)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        lbl_bl_code.Text = db.Rows(0)("TenND").ToString().Trim()
                    End If
                End If
            End Using
        End If
        cb_hd_bangluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_hd_bangluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_hd_bangluong.SelectedIndexChanged
        cb_hd_ngachluong.Items.Clear()
        arr_NgachLuong.Clear()
        lbl_bl_code.Text = ""
        If (cb_hd_bangluong.SelectedIndex > 0 And cb_hd_bangluong.Items.Count <> 0) Then
            Dim BangLuongId As Integer = CType(arr_BangLuong(IIf(cb_hd_bangluong.SelectedIndex > 0, cb_hd_bangluong.SelectedIndex, "0")), Integer)
            Dim strSQL As String = ""

            If (BangLuongId > 0) Then
                strSQL = String.Format("Select IdNgachLuong, Mota From NgachLuong Where IdBangLuong = {0}", BangLuongId)
                arr_NgachLuong = _Globals.Bind_ComBoBox(cb_hd_ngachluong, strSQL, "---Ngạch lương---")
            End If

            strSQL = String.Format("Select * From BangLuong Where IdBangLuong = {0}", BangLuongId)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        lbl_bl_code.Text = db.Rows(0)("BangLuong").ToString().Trim()
                    End If
                End If
            End Using
        End If
        cb_hd_ngachluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_hd_ngachluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_hd_ngachluong.SelectedIndexChanged
        cb_hd_bacluong.Items.Clear()
        arr_BacLuong.Clear()
        If (cb_hd_ngachluong.SelectedIndex > 0 And cb_hd_ngachluong.Items.Count <> 0) Then
            Dim NgachLuongId As Integer = CType(arr_NgachLuong(IIf(cb_hd_ngachluong.SelectedIndex > 0, cb_hd_ngachluong.SelectedIndex, "0")), Integer)
            If (NgachLuongId > 0) Then
                Dim strSQL As String = String.Format("Select IdBacLuong, BacLuong From BacLuong Where IdNgachLuong = {0}", NgachLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            cb_hd_bacluong.Items.Add("---Bậc lương---")
                            arr_BacLuong.Add(0)
                            For i As Integer = 0 To db.Rows.Count - 1
                                cb_hd_bacluong.Items.Add(db.Rows(i)("BacLuong").ToString())
                                arr_BacLuong.Add(IIf(db.Rows(i)("IdBacLuong").ToString() <> "", db.Rows(i)("IdBacLuong").ToString(), ""))
                            Next
                            If (cb_hd_bacluong.Items.Count <> 0) Then
                                cb_hd_bacluong.SelectedIndex = 0
                            End If
                        End If
                    End If
                End Using
                'Load thông tin tên bảng lương ra theo dõi
                strSQL = String.Format("Select * From NgachLuong Where IdNgachLuong = {0}", NgachLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            lbl_ngl_code.Text = db.Rows(0)("NgachLuong").ToString().Trim()
                        End If
                    End If
                End Using
            End If
        End If
        cb_hd_bacluong_SelectedIndexChanged(sender, Nothing)
    End Sub

    Private Sub cb_hd_bacluong_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_hd_bacluong.SelectedIndexChanged
        edt_hd_heso.Text = ""
        If (cb_hd_bacluong.SelectedIndex > 0 And cb_hd_bacluong.Items.Count <> 0) Then
            Dim BacLuongId As Integer = CType(arr_BacLuong(IIf(cb_hd_bacluong.SelectedIndex > 0, cb_hd_bacluong.SelectedIndex, "0")), Integer)
            If (BacLuongId > 0) Then
                Dim strSQL As String = String.Format("Select * From BacLuong Where IdBacLuong = {0}", BacLuongId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            edt_hd_heso.Text = IIf(db.Rows(0)("Heso").ToString().Trim() <> "", CType(db.Rows(0)("Heso"), Double).ToString("N2"), "0")
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    'Kiểm tra nếu là hợp đồng không xác định thời hạn thì --> Không cho người dùng nhập Đến Ngày
    Private Sub cb_hd_loaihd_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_hd_loaihd.SelectedIndexChanged
        _CodeLoaiHD = ""
        dtpk_hd_denngay.Enabled = True
        dtpk_hd_denngay.Text = DateTime.Now.ToShortDateString()
        If (cb_hd_loaihd.SelectedIndex > 0 And cb_hd_loaihd.Items.Count <> 0) Then
            Dim _IdLoaiHD As Integer = CType(arr_LoaiHd(IIf(cb_hd_loaihd.SelectedIndex > 0, cb_hd_loaihd.SelectedIndex, "0")), Integer)
            Dim strSQL As String = String.Format("Select id, ma_so From DanhMuc Where id = {0} And id_goc = 30 and Status = 1", _IdLoaiHD)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        _CodeLoaiHD = db.Rows(0)("ma_so").ToString().Trim()
                        If (_CodeLoaiHD = "3006") Then
                            labSoHD.Text = "Số QĐ"
                            labLoaiHD.Text = "Quyết định"
                            cb_hd_httraluong.Enabled = False
                            cb_hd_nghidinh.Enabled = False
                            cb_hd_bangluong.Enabled = False
                            cb_hd_ngachluong.Enabled = False
                            cb_hd_bacluong.Enabled = False
                            edt_hd_heso.Enabled = False
                            edt_hd_tyle.Enabled = False
                            txtNoiLamViec.Enabled = False
                        Else
                            labSoHD.Text = "Số hợp đồng"
                            labLoaiHD.Text = "Loại hợp đồng"
                            txtNoiLamViec.Enabled = True
                            cb_hd_httraluong.Enabled = True
                            cb_hd_nghidinh.Enabled = True
                            cb_hd_bangluong.Enabled = True
                            cb_hd_ngachluong.Enabled = True
                            cb_hd_bacluong.Enabled = True
                            edt_hd_heso.Enabled = True
                            edt_hd_tyle.Enabled = True
                            If (_CodeLoaiHD = "3001") Then
                                dtpk_hd_denngay.Checked = False
                                dtpk_hd_denngay.Enabled = False
                            End If
                        End If
                    End If
                End If
            End Using
        End If
    End Sub

    Private Sub lkl_xemct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkl_xemct.LinkClicked
        If (_IdCanBo <> "") Then
            Dim obj_detail As frmHS_ChiTiet = New frmHS_ChiTiet()
            'Lấy danh sách mảng các id hiện có trên lưới dl
            Dim arrRows As ArrayList = New ArrayList()
            arrRows.Add(_IdCanBo)
            obj_detail.RecordCurrent = 0
            obj_detail.Records = 1
            obj_detail.IdCanBo = _IdCanBo
            obj_detail._IdDonviHT = _IdDonviHT
            obj_detail.arr_RecordId = arrRows
            obj_detail.ShowDialog()
        Else
            Return
        End If
    End Sub

    Private Sub btn_huybo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_huybo.Click
        Dim _currRow As String = _RowId
        ResetAll_Controls()
        If (dgv_hdld.Rows.Count <= 0) Then Return
        If (_currRow = "") Then _currRow = dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString()
        dgv_hdld.CurrentRow.Selected = False
        dgv_hdld.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hdld, "cln_Code")).Selected = True
        If (dgv_hdld.Rows.Count > 0) Then
            If ((dgv_hdld.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                FillData_SelectRow(_currRow)
            End If
        End If
    End Sub

    Private Sub btn_addnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.Click
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ cần thêm mới hợp đồng lao động!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        If (Globals.Roles.IndexOf(";136;") < 0) Then
            MessageBox.Show("Bạn không có quyền thực hiện thêm mới Hợp đồng lao động!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
            dgv_hdld_CellClick(sender, Nothing)
            btn_addnew.Visible = False
            Return
        End If

        'If (IsAdd()) Then
        ResetAll_Controls()
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_hdld)
        If (dgv_hdld.Rows.Count <= 0) Then
            'Kiểm tra xem có phải là cán bộ chuyển từ Tập sự sang không
            FillData_Contract()
        End If
        'Thực hiện load các thông tin mặc định có sẵn
        edt_hd_nguoiky.Text = "" '_PersonnelFile.GetVarNam("GIAMDOC")
        cb_hd_chucvu.SelectedIndex = CType(arr_Chucvu_Hd.IndexOf(_PersonnelFile.GetDefault_ChucVu(_IdCanBo).ToString()), Integer)
        cb_hd_httraluong.SelectedIndex = IIf(cb_hd_httraluong.FindString("Tháng") > 0, cb_hd_httraluong.FindString("Tháng"), 0)
        cb_hd_nghidinh.SelectedIndex = IIf(cb_hd_nghidinh.Items.Count <> 0, 1, 0)
        Dim strSQL As String = ""
        'Kiểm tra xem đã có hợp đồng Ngắn hạn nào trước đó của cán bộ này chưa. Nếu có thì thực hiện load các thông tin có sẵn
        strSQL = "SELECT * FROM HS_HDLD WHERE IdCanBo = '" + _IdCanBo + "' And Not (DenNgay Is Null Or DenNgay = '' Or DenNgay = '1900/01/01') Order By DenNgay Desc"
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    dtpk_hd_ngayhl.Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime).AddDays(1)
                End If
            End If
        End Using
        ActiveControl = edt_hd_sohd
        'End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu Có sẵn hợp đồng lao động bên Hồ sơ Cán bộ Tập sự
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FillData_Contract()
        Dim strSQL As String = String.Format("Select * From HSCB_TS Where idNew = '{0}' And TrangThai = 1", _IdCanBo)
        Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (_db Is Nothing) Then
                If (_db.Rows.Count > 0) Then
                    'Lấy được id Cán bộ Tập sự
                    If (_db.Rows(0)("Id").ToString() <> "") Then
                        Using db As DataTable = _PersonnelFile.GetHSCB_TS_HDLD_GetSearch(_db.Rows(0)("Id").ToString(), "", 0, "")
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    'Thực hiện fill dữ liệu liên quan đến Hợp đồng lao động
                                    edt_hd_sohd.Text = db.Rows(0)("SoQD").ToString().Trim()
                                    edt_hd_nguoiky.Text = db.Rows(0)("NguoiQD").ToString().Trim()
                                    cb_hd_chucvu.SelectedIndex = CType(arr_Chucvu_Hd.IndexOf(db.Rows(0)("IdCV_Nguoi_QD").ToString()), Integer)
                                    cb_hd_httraluong.SelectedIndex = CType(arr_HTTL.IndexOf(db.Rows(0)("IdHT_TraLuong").ToString()), Integer)
                                    If db.Rows(0)("NgayQD").ToString().Trim() <> "" Then
                                        dtpk_hd_ngayky.Value = CType(db.Rows(0)("NgayQD").ToString(), DateTime)
                                        dtpk_hd_ngayhl.Value = CType(db.Rows(0)("NgayQD").ToString(), DateTime)
                                    End If
                                    If db.Rows(0)("Tungay").ToString().Trim() <> "" Then
                                        dtpk_hd_tungay.Value = CType(db.Rows(0)("Tungay").ToString(), DateTime)
                                    End If
                                    If db.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                        If (CType(db.Rows(0)("DenNgay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                            dtpk_hd_denngay.Checked = False
                                        Else
                                            dtpk_hd_denngay.Checked = True
                                            dtpk_hd_denngay.Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime)
                                        End If
                                    Else
                                        dtpk_hd_denngay.Checked = True
                                    End If

                                    Dim _NghiDinhLuongId As Integer = 0
                                    Dim _BangLuongId As Integer = 0
                                    Dim _NgachLuongId As Integer = 0
                                    Dim _BacLuongId As Integer = 0
                                    _BacLuongId = IIf(db.Rows(0)("IdBacLuong").ToString() <> "0", CType(db.Rows(0)("IdBacLuong").ToString(), Integer), 0) ' Id bậc lương
                                    If (_BacLuongId > 0) Then
                                        _NgachLuongId = _Labour.GetIdNgachLuong(_BacLuongId)
                                        If (_NgachLuongId > 0) Then
                                            _BangLuongId = _Labour.GetIdBangLuong(_NgachLuongId)
                                        End If
                                        If (_BangLuongId > 0) Then
                                            _NghiDinhLuongId = _Labour.GetIdNghiDinhLuong(_BangLuongId)
                                        End If
                                        cb_hd_nghidinh.SelectedIndex = CType(arr_NghiDinhLuong.IndexOf(_NghiDinhLuongId.ToString()), Integer)
                                        cb_hd_nghidinh_SelectedIndexChanged(Nothing, Nothing)
                                        cb_hd_bangluong.SelectedIndex = CType(arr_BangLuong.IndexOf(_BangLuongId.ToString()), Integer)
                                        cb_hd_bangluong_SelectedIndexChanged(Nothing, Nothing)
                                        cb_hd_ngachluong.SelectedIndex = CType(arr_NgachLuong.IndexOf(_NgachLuongId.ToString()), Integer)
                                        cb_hd_ngachluong_SelectedIndexChanged(Nothing, Nothing)
                                        cb_hd_bacluong.SelectedIndex = CType(arr_BacLuong.IndexOf(db.Rows(0)("IdBacLuong").ToString()), Integer)
                                        cb_hd_bacluong_SelectedIndexChanged(Nothing, Nothing)
                                        edt_hd_tyle.Text = IIf(db.Rows(0)("TyleHuong").ToString().Trim() <> "", CType(db.Rows(0)("TyleHuong"), Double).ToString("N2"), "0")
                                    End If
                                    MessageBox.Show("Thông tin hợp đồng lao động của cán bộ '" + lbl_hoten.Text.Trim() + "' được lấy từ hồ sơ hợp đồng lao động của cán bộ tập sự mà người sử dụng đã khai báo trước đó!" + vbCrLf + "Lưu ý: Các thông tin hợp đồng lao động được hiển thị kế thừa giúp người sử dụng không phải nhập liệu lại, người sử dụng có thể nhập lại sau đó cập nhật.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                                    Return
                                End If
                            End If
                        End Using
                    End If
                End If
            End If
        End Using
    End Sub

    Private Sub btn_xoadl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_xoadl.Click
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ cần xoá hồ sơ Hợp đồng lao động!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        If (Globals.Roles.IndexOf(";138;") < 0) Then
            MessageBox.Show("Bạn không có quyền thực xoá Hợp đồng lao động!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
            btn_xoadl.Visible = False
            dgv_hdld_CellClick(sender, Nothing)
            Return
        End If
        If (dgv_hdld.Rows.Count <= 0) Then Return
        Try
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _count As Int16 = 0
            If (dgv_hdld.Rows.Count > 0) Then
                If ((dgv_hdld.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    For i As Integer = 0 To dgv_hdld.Rows.Count - 1
                        If (dgv_hdld.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                            If (CType(dgv_hdld.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                _count = _count + 1
                                arr_Del.Add(dgv_hdld.Rows(i).Cells("cln_Code").Value.ToString())
                            End If
                        End If
                    Next
                End If
            End If
            If (_count > 0) Then
                Dim _mess As String = IIf(_count = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hợp đồng lao động đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    For i As Int16 = 0 To arr_Del.Count - 1
                        _Labour.Delete_LaborContract(arr_Del(i).ToString())
                    Next
                    'Thực hiện Load thông tin sau khi xoá song --> Thành công
                    dgv_hdld.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_hdld)
                    dgv_hdld_CellClick(sender, Nothing)
                    MessageBox.Show("Bạn đã xoá thành công thông tin hợp đồng lao động của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Else
                    arr_Del.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_hdld)
                End If
            Else
                MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Cập nhật xoá bỏ hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_chapnhan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_chapnhan.Click
        If (_IdCanBo = "") Then Return
        Try
            If (_RowId <> "") Then
                If (Globals.Roles.IndexOf(";137;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi hợp đồng lao động của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                    Dim _currRow As String = _RowId
                    dgv_hdld.CurrentRow.Selected = False
                    dgv_hdld.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hdld, "cln_Code")).Selected = True
                    'Gọi lại sự kiện cell click của lưới dữ liệu
                    ResetAll_Controls()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_hdld)
                    If (dgv_hdld.Rows.Count > 0) Then
                        If ((dgv_hdld.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            FillData_SelectRow(_currRow)
                        End If
                    End If
                    Return
                End If
            End If
            If (IsValid()) Then
                Dim _currRow As String = ""

                Dim obj_labour As clsHS_Hdld.HS_HDLD = New clsHS_Hdld.HS_HDLD()
                obj_labour.IdCanBo = _IdCanBo
                obj_labour.SoHD = Globals.Find_Replace(edt_hd_sohd.Text.ToString().Trim())
                obj_labour.IdLoaiHD = CType(IIf(arr_LoaiHd.Count > 0, arr_LoaiHd(cb_hd_loaihd.SelectedIndex), "0"), Integer)
                obj_labour.Ngay_HL = dtpk_hd_ngayhl.Value
                obj_labour.NgayKy_HD = dtpk_hd_ngayky.Value
                obj_labour.NguoKy_QD = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hd_nguoiky.Text.ToString().Trim()))
                obj_labour.IdCV_Nguoiky_QD = CType(IIf(arr_Chucvu_Hd.Count > 0, arr_Chucvu_Hd(cb_hd_chucvu.SelectedIndex), "0"), Integer)

                obj_labour.TuNgay = dtpk_hd_tungay.Value
                obj_labour.DenNgay = IIf(dtpk_hd_denngay.Checked, dtpk_hd_denngay.Value, DateTime.Parse("01/01/1900"))
                obj_labour.IdHT_TraLuong = CType(IIf(arr_HTTL.Count > 0, arr_HTTL(cb_hd_httraluong.SelectedIndex), "0"), Integer)

                If (cb_hd_bacluong.Items.Count <> 0) Then
                    obj_labour.IdBacLuong = CType(IIf(arr_BacLuong.Count > 0, arr_BacLuong(cb_hd_bacluong.SelectedIndex), "0"), Integer)
                    obj_labour.HeSo = IIf(edt_hd_heso.Text.ToString() <> "", CType(edt_hd_heso.Text.ToString().Trim(), Double), 0)
                Else
                    obj_labour.IdBacLuong = 0
                    obj_labour.HeSo = 0
                End If
                obj_labour.TyleHuong = IIf(edt_hd_tyle.Text.ToString() <> "", CType(edt_hd_tyle.Text.ToString().Trim(), Double), 0)
                obj_labour.DVKyHDLD = standardizeString(txtDVKyHD.Text)
                obj_labour.NoiLamViec = standardizeString(txtNoiLamViec.Text)
                obj_labour.GhiChu = Globals.Find_Replace(edt_hd_ghichu.Text.ToString().Trim())
                Dim strName As String = IIf(lbl_gioitinh.Text.Trim() = "Nam", "ông: ", "bà: ") + lbl_hoten.Text.Trim()

                If (_RowId = "") Then
                    _currRow = _Labour.Insert_LaborContract(obj_labour)
                Else
                    obj_labour.IdCB_HDLD = _RowId
                    _Labour.Update_LaborContract(obj_labour)
                    _currRow = _RowId
                End If
                'Thực hiện fill lại dữ liệu - Hợp đồng lao động
                'ResetAll_Controls()
                dgv_hdld.Rows.Clear()
                tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                'Select item của cây dữ liệu
                Dim node As TreeNode = Nothing
                If (strNode <> "") Then
                    node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                    If Not (node Is Nothing) Then
                        tv_main.SelectedNode = node
                    End If
                End If

                dgv_hdld.CurrentRow.Selected = False
                dgv_hdld.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hdld, "cln_Code")).Selected = True

                'Gọi lại sự kiện cell click của lưới dữ liệu
                ResetAll_Controls()
                ckb_ChoiceAll.Checked = False
                HeaderCheckBoxClick(ckb_ChoiceAll, dgv_hdld)
                If (dgv_hdld.Rows.Count > 0) Then
                    If ((dgv_hdld.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        FillData_SelectRow(_currRow)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Cập nhật hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub
#End Region

#Region "---> Events: Các sự kiện ngoại lệ người dùng <---"
    Private Sub frmHS_HDLD_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub edt_hd_sohd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_sohd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hd_sohd.Text.Trim() <> "") Then
                edt_hd_sohd.Text = edt_hd_sohd.Text.ToUpper()
                cb_hd_loaihd.Focus()
            Else
                edt_hd_sohd.Focus()
            End If
        End If
    End Sub

    Private Sub edt_hd_sohd_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hd_sohd.Leave
        If (edt_hd_sohd.Text.Trim() <> "") Then
            edt_hd_sohd.Text = edt_hd_sohd.Text.ToUpper()
            cb_hd_loaihd.Focus()
        Else
            edt_hd_sohd.Focus()
        End If
    End Sub

    Private Sub cb_hd_loaihd_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_loaihd.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hd_ngayhl.Focus()
        End If
    End Sub

    Private Sub dtpk_hd_ngayhl_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hd_ngayhl.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hd_ngayky.Focus()
        End If
    End Sub

    Private Sub dtpk_hd_ngayhl_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_hd_ngayhl.ValueChanged
        If (_RowId = "") Then   'Trường hợp thêm mới --> Ngày bắt đầu = Ngày hiệu lực của Hợp đồng
            dtpk_hd_tungay.Value = dtpk_hd_ngayhl.Value
            If (_CodeLoaiHD <> "3001") Then     'Hợp đồng có thời hạn
                dtpk_hd_denngay.Value = dtpk_hd_ngayhl.Value.AddYears(1)
            End If
            'Ngày ký mặc định trước ngày hiệu lực 10 ngày
            dtpk_hd_ngayky.Value = dtpk_hd_ngayhl.Value.AddDays(-10)
        End If
    End Sub

    Private Sub dtpk_hd_ngayky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hd_ngayky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hd_nguoiky.Focus()
        End If
    End Sub

    Private Sub edt_hd_nguoiky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_nguoiky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_hd_tungay.Focus()
        End If
    End Sub

    Private Sub edt_hd_nguoiky_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_nguoiky.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_hd_ngayky.Focus()
        End If
    End Sub

    Private Sub dtpk_hd_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hd_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hd_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_hd_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hd_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hd_chucvu.Focus()
        End If
    End Sub

    Private Sub cb_hd_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hd_httraluong.Focus()
        End If
    End Sub

    Private Sub cb_hd_httraluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_httraluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hd_nghidinh.Focus()
        End If
    End Sub

    Private Sub cb_hd_nghidinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_nghidinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hd_bangluong.Focus()
        End If
    End Sub

    Private Sub cb_hd_bangluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_bangluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hd_ngachluong.Focus()
        End If
    End Sub

    Private Sub cb_hd_ngachluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_ngachluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hd_bacluong.Focus()
        End If
    End Sub

    Private Sub cb_hd_bacluong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hd_bacluong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hd_heso.Focus()
        End If
    End Sub

    Private Sub edt_hd_heso_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_heso.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_hd_tyle.Focus()
        End If
    End Sub

    Private Sub edt_hd_heso_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_heso.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_hd_bacluong.Focus()
        End If
    End Sub

    Private Sub edt_hd_heso_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_hd_heso.KeyPress
        Dim decimalString As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
        Dim decimalChar As Char = Convert.ToChar(decimalString)
        If (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
        ElseIf (e.KeyChar = decimalString And edt_hd_heso.Text.IndexOf(decimalString) = -1) Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub edt_hd_tyle_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_tyle.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hd_tyle.Text.Trim() = "") Then
                edt_hd_tyle.Text = "0"
            End If
            edt_hd_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_hd_tyle_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hd_tyle.Leave
        If (edt_hd_tyle.Text.Trim() = "") Then
            edt_hd_tyle.Text = "0"
        End If
    End Sub

    Private Sub edt_hd_tyle_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_hd_tyle.KeyPress
        Dim decimalString As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
        Dim decimalChar As Char = Convert.ToChar(decimalString)
        If (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
        ElseIf (e.KeyChar = decimalString And edt_hd_tyle.Text.IndexOf(decimalString) = -1) Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub edt_hd_tyle_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_tyle.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hd_heso.Focus()
        End If
    End Sub

    Private Sub edt_hd_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_chapnhan.Focus()
        End If
    End Sub

    Private Sub edt_hd_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hd_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hd_tyle.Focus()
        End If
    End Sub
#End Region

#Region "---> Events: Các Hàm và sự kiện liên quan đến Check all các items trên lưới dữ liệu <---"
    Private Sub ckb_ChoiceAll_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Space) Then
            HeaderCheckBoxClick(CType(sender, CheckBox), dgv_hdld)
        End If
    End Sub

    Private Sub ckb_ChoiceAll_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        HeaderCheckBoxClick(CType(sender, CheckBox), dgv_hdld)
    End Sub

    Private Sub dgv_hdld_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        'Sự kiện này thực hiện khi click các checked items trên lưới dữ liệu thì sẽ Checked hoặc UnChecked Checkall
        If CType(sender, DataGridView).Columns(e.ColumnIndex).Name = "cln_Choice" And e.RowIndex >= 0 Then
            If Not IsHeaderCheckBoxClicked Then
                Dim _vCellCheck As DataGridViewCheckBoxCell
                _vCellCheck = CType(dgv_hdld("cln_Choice", e.RowIndex), DataGridViewCheckBoxCell)
                RowCheckBoxClick(_vCellCheck)
            End If
        End If
    End Sub

    Private Sub dgv_hdld_CurrentCellDirtyStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (dgv_hdld.Rows.Count > 0) Then
            If ((dgv_hdld.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hdld.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Dim vCellCheck As DataGridViewCheckBoxCell
                'Checking whether the Datagridview Checkbox column is the first column
                If dgv_hdld.CurrentCellAddress.X = 0 Then
                    vCellCheck = dgv_hdld.CurrentRow.Cells("cln_Choice")
                    If (dgv_hdld.IsCurrentCellDirty) Then 'Checking for dirty cell
                        dgv_hdld.CommitEdit(DataGridViewDataErrorContexts.Commit) 'If it is dirty, making them to commit
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub dgv_hdld_CellPainting(ByVal sender As System.Object, ByVal e As DataGridViewCellPaintingEventArgs)
        'Căn chỉnh cho Checkbox nằm vào giữa tiêu đề của cột Chọn xoá
        If (e.RowIndex = -1 And e.ColumnIndex = 0) Then
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
        dgv_hdld.Controls.Add(ckb_ChoiceAll)
    End Sub

    ''' <summary>
    '''  Hàm thực hiện căn chỉnh cho CheckBox nằm giữa cột tiêu đề của lưới dữ liệu
    ''' </summary>
    ''' <param name="_ColumnIndex"></param>
    ''' <param name="_RowIndex"></param>
    ''' <remarks></remarks>
    Private Sub ResetHeaderCheckBoxLocation(ByVal _ColumnIndex As Integer, ByVal _RowIndex As Integer)
        'Get the column header cell bounds
        Dim oRectangle As Rectangle = dgv_hdld.GetCellDisplayRectangle(_ColumnIndex, _RowIndex, True)
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

End Class