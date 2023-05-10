Public Class DMLKhacForm

#Region "---> Defined parametter and properties <---"
    Dim _DanhMuc As clsHT_DanhMuc = New clsHT_DanhMuc
    Dim _Globals As Globals = New Globals
    Private _Labour As clsHS_Hdld = New clsHS_Hdld()
    Private _SqlHelper As DBAccess

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        _SqlHelper = New DBAccess
    End Sub

    Private _RecordId As String
    Public Property RecordId() As String
        Get
            Return _RecordId
        End Get
        Set(ByVal value As String)
            _RecordId = value
        End Set
    End Property

    ''' <summary>
    ''' Chỉ số xác định danh mục lương được gọi cập nhật. Với quy định:
    '''                    _State = 4: Nghị định lương
    '''                    _State = 5: Bảng lương
    '''                    _State = 6: Ngạch lương
    '''                    _State = 7: Bậc lương
    '''                    _State = 8: Lương cơ bản
    '''                    _State = 9: Mức phụ cấp
    ''' </summary>
    ''' <remarks></remarks>
    Public _Status As Byte

    Private arrRoot As ArrayList = New ArrayList()
    Private arrChild As ArrayList = New ArrayList()
    Private strSQL As String = ""

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
#End Region

#Region "---> Functions main: Các hàm chính <---"

    ''' <summary>
    ''' Hàm thiết lập các controls cho phù hợp với từng danh mục lương
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Create_Controls()
        edt_luongcb.Visible = False
        dtpk_ngayhuong.Visible = False
        dtpk_ngayky.Visible = False
        pnl_control_7.Visible = False
        pnl_control_8.Visible = False
        pnl_control_9.Visible = False
        lbl_div_7.Visible = False
        lbl_div_8.Visible = False
        lbl_div_9.Visible = False
        lbl_time.Visible = False
        Select Case _Status
            Case 4              'Nghị định lương
                lbl_title.Text = " CẬP NHẬT DANH MỤC - NGHỊ ĐỊNH LƯƠNG"
                cb_main_1.Visible = False
                cb_main_2.Visible = False
                edt_dl_1.Dock = DockStyle.Fill
                pnl_control_5.Height = 65
                edt_dl_5.Multiline = True

            Case 5              'Bảng lương
                lbl_title.Text = " CẬP NHẬT DANH MỤC - BẢNG LƯƠNG"
                edt_dl_1.Visible = False
                cb_main_2.Visible = False
                lbl_title_1.Text = "Nghị định lương "
                lbl_title_2.Text = "Bảng lương "
                lbl_title_3.Text = "Mô tả chi tiết "

                pnl_control_3.Height = 81
                edt_dl_3.Multiline = True
                pnl_control_4.Visible = False
                lbl_div_4.Visible = False
                pnl_control_5.Visible = False
                lbl_div_5.Visible = False
                Me.Height = 261

                'Fill data Nghị định lương vào ComboBox Nghị định lương
                arrRoot.Clear()
                cb_main_1.Items.Clear()
                strSQL = String.Format("Select IdNDLuong, TenND From NghiDinhLuong Order by TenND asc")
                arrRoot = _Globals.Bind_ComBoBox(cb_main_1, strSQL, "---Nghị định lương---")

            Case 6              'Ngạch lương
                lbl_time.Visible = True
                lbl_time.Text = " tháng"
                lbl_title.Text = " CẬP NHẬT DANH MỤC - NGẠCH LƯƠNG"
                edt_dl_1.Visible = False
                cb_main_2.Visible = False
                lbl_title_1.Text = "Bảng lương "
                lbl_title_2.Text = "Ngạch lương "
                lbl_title_3.Text = "Phân loại "
                lbl_title_4.Text = "Mô tả "
                lbl_title_5.Text = "Thời gian nâng "
                pnl_control_4.Height = 65
                edt_dl_4.Multiline = True
                pnl_control_5.Height = 22

                lbl_title_1.Width = 99
                lbl_title_2.Width = 99
                lbl_title_3.Width = 99
                lbl_title_4.Width = 99
                lbl_title_5.Width = 99
                lbl_title_6.Width = 99

                'Fill data Bảng lương vào ComboBox Bảng lương
                arrRoot.Clear()
                cb_main_1.Items.Clear()
                strSQL = String.Format("Select IdBangLuong, MoTa From BangLuong")
                arrRoot = _Globals.Bind_ComBoBox(cb_main_1, strSQL, "---Bảng lương---")

            Case 7              'Bậc lương
                lbl_title.Text = " CẬP NHẬT DANH MỤC - BẬC LƯƠNG"
                edt_dl_1.Visible = False
                edt_dl_2.Visible = False
                cb_main_2.Dock = DockStyle.Fill
                pnl_control_5.Height = 65
                edt_dl_5.Multiline = True

                lbl_title_1.Text = "Bảng lương "
                lbl_title_2.Text = "Ngạch lương "
                lbl_title_3.Text = "Bậc lương "
                lbl_title_4.Text = "Hệ số "
                lbl_title_5.Text = "Mô tả "

                'Fill data Bảng lương vào ComboBox Bảng lương
                arrRoot.Clear()
                cb_main_1.Items.Clear()
                strSQL = String.Format("Select IdBangLuong, MoTa From BangLuong")
                arrRoot = _Globals.Bind_ComBoBox(cb_main_1, strSQL, "---Bảng lương---")
                cb_main_1_SelectedIndexChanged(Nothing, Nothing)

            Case 8              'Lương cơ bản
                Me.Height = 357
                edt_dl_1.Visible = False
                cb_main_1.Visible = False
                cb_main_2.Visible = False
                edt_dl_3.Visible = False
                edt_dl_5.Visible = False
                edt_luongcb.Visible = True
                dtpk_ngayhuong.Visible = True
                dtpk_ngayhuong.Width = 191
                dtpk_ngayky.Visible = True
                dtpk_ngayky.Width = 191
                pnl_control_7.Visible = True
                pnl_control_8.Visible = True
                pnl_control_9.Visible = True
                lbl_div_7.Visible = True
                lbl_div_8.Visible = True
                lbl_div_9.Visible = True
                pnl_control_7.Height = 22
                pnl_control_8.Height = 22
                pnl_control_9.Height = 22

                edt_luongcb.Dock = DockStyle.Fill
                lbl_title_1.Text = "Lương tối thiểu "
                lbl_title_2.Text = "Hệ số ngành "
                lbl_title_3.Text = "Ngày áp dụng "
                lbl_title_4.Text = "Số quyết định "
                lbl_title_5.Text = "Ngày ký "
                pnl_control_5.Height = 22

                lbl_title_7.Text = "Người ký "
                lbl_title_8.Text = "Chức vụ "
                lbl_title_9.Text = "Ghi chú "
                pnl_control_9.Height = 65
                edt_dl_9.Multiline = True
                'Fill data vào combobox - Chức vụ người quyết định
                arrRoot.Clear()
                cb_main_3.Items.Clear()
                strSQL = String.Format("Select id, ten_goi From DanhMuc Where id_goc = 14 and ma_so IN ('1401','1402','1403','1404','1410','1411')")
                arrRoot = _Globals.Bind_ComBoBox(cb_main_3, strSQL, "---Chức vụ người ký---")

            Case 9              'Mức phụ cấp
                lbl_title.Text = " CẬP NHẬT DANH MỤC - MỨC PHỤ CẤP"
                edt_dl_1.Visible = False
                cb_main_2.Visible = False
                lbl_title_1.Text = "Loại phụ cấp "
                lbl_title_2.Text = "Mức phụ cấp "
                pnl_control_3.Visible = False
                pnl_control_4.Visible = False
                pnl_control_5.Visible = False
                lbl_div_4.Visible = False
                lbl_div_5.Visible = False
                lbl_div_6.Visible = False

                lbl_div_1.Height = 21
                Me.Height = 201

                'Fill data Nghị định lương vào ComboBox Nghị định lương
                arrRoot.Clear()
                cb_main_1.Items.Clear()
                arrRoot = _Globals.Bind_ComBoBox(cb_main_1, clsHT_DanhMuc.Sql_MucPhCap, "---Loại phụ cấp---")

            Case 10              'Tham số hệ thống Lương
                lbl_title.Text = " CẬP NHẬT THAM SỐ LƯƠNG HỆ THỐNG"
                If (_RecordId > 0) Then     ' Trường hợp sửa đổi
                    cb_main_1.Visible = False
                    edt_dl_1.Dock = DockStyle.Fill
                Else                        ' Trường hợp thêm mới dữ liệu
                    edt_dl_1.Visible = False
                    cb_main_1.Visible = True
                    'Fill data vào combobox - luôn
                    cb_main_1.Items.Clear()
                    cb_main_1.Items.Add("--- Tên gọi tham số ---")
                    strSQL = String.Format("Select Distinct NameVar from SalaryVar")
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    cb_main_1.Items.Add(db.Rows(i)("NameVar").ToString())
                                Next
                            End If
                        End If
                    End Using
                End If

                cb_main_2.Visible = False

                pnl_control_4.Visible = False
                edt_dl_5.Visible = False
                lbl_div_5.Visible = False
                lbl_time.Visible = False
                dtpk_ngayky.Visible = True

                lbl_title_1.Text = "Tên gọi tham số "
                lbl_title_2.Text = "Giá trị tham số "
                lbl_title_3.Text = "Tỷ lệ tương ứng "
                lbl_title_5.Text = "Ngày áp dụng "
                lbl_div_7.Visible = True
                pnl_control_7.Visible = True
                lbl_title_7.Text = "Mô tả chi tiết "

                pnl_control_7.Height = 65
                edt_dl_7.Multiline = True
                dtpk_ngayky.Width = 121
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện reset controls về trạng thái ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetControls()
        ckb_lock.Checked = False
        ckb_lientuc.Checked = False
        Select Case _Status
            Case 4              'Nghị định lương
                edt_dl_1.Text = ""
                edt_dl_2.Text = ""
                edt_dl_3.Text = ""
                edt_dl_4.Text = ""
                edt_dl_5.Text = ""

                ActiveControl = edt_dl_1

            Case 5              'Bảng lương
                cb_main_1.SelectedIndex = 0
                edt_dl_2.Text = ""
                edt_dl_3.Text = ""
                ActiveControl = cb_main_1

            Case 6              'Ngạch lương
                cb_main_1.SelectedIndex = 0
                edt_dl_2.Text = ""
                edt_dl_3.Text = ""
                edt_dl_4.Text = ""
                edt_dl_5.Text = ""
                ActiveControl = cb_main_1

            Case 7              'Bậc lương
                cb_main_1.SelectedIndex = 0
                cb_main_1_SelectedIndexChanged(Nothing, Nothing)
                edt_dl_3.Text = ""
                edt_dl_4.Text = ""
                edt_dl_5.Text = ""
                ActiveControl = cb_main_1

            Case 8              'Lương cơ bản
                edt_luongcb.Text = "0"
                edt_dl_2.Text = ""
                dtpk_ngayhuong.Text = DateTime.Now.ToShortDateString()
                edt_dl_4.Text = ""
                dtpk_ngayky.Text = DateTime.Now.ToShortDateString()
                edt_dl_7.Text = ""
                edt_dl_9.Text = ""
                ActiveControl = edt_luongcb

            Case 9              'Mức phụ cấp
                cb_main_1.SelectedIndex = 0
                edt_dl_2.Text = ""
                ActiveControl = cb_main_1

            Case 10              'Tham số Lương hệ thống
                ActiveControl = edt_dl_1
                If (_RecordId > 0) Then
                    edt_dl_1.Text = ""
                Else
                    cb_main_1.SelectedIndex = 0
                End If
                edt_dl_2.Text = ""
                edt_dl_3.Text = ""
                dtpk_ngayky.Text = DateTime.Now.ToShortDateString()
                edt_dl_7.Text = ""
                ckb_lock.Checked = False
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra điều kiện hợp lệ nhập liệu
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Select Case _Status
            Case 4          'Nghị định lương
                If (edt_dl_1.Text.Trim() = "") Then
                    MessageBox.Show("Mã nghị định lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_1
                    Return False
                End If
                If (edt_dl_2.Text.Trim() = "") Then
                    MessageBox.Show("Số nghị định lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_2
                    Return False
                End If
                If (edt_dl_3.Text.Trim() = "") Then
                    MessageBox.Show("Tên nghị định lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_3
                    Return False
                End If
                If (edt_dl_4.Text.Trim() = "") Then
                    MessageBox.Show("Phân loại nghị định lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_4
                    Return False
                End If
                If (edt_dl_1.Text.Trim() <> "") Then
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from NghiDinhLuong Where MaND = N'{0}' and IdNDLuong <> {1}", Globals.Find_Replace(edt_dl_1.Text.Trim().ToString()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from NghiDinhLuong Where MaND = N'{0}'", Globals.Find_Replace(edt_dl_1.Text.Trim().ToString()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Mã nghị định lương '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_1.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_1
                                Return False
                            End If
                        End If
                    End Using
                End If
                If (edt_dl_2.Text.Trim() <> "") Then
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from NghiDinhLuong Where SoND = N'{0}' and IdNDLuong <> {1}", Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from NghiDinhLuong Where SoND = N'{0}'", Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Số nghị định lương '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_2.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_2
                                Return False
                            End If
                        End If
                    End Using
                End If
                If (edt_dl_3.Text.Trim() <> "") Then
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from NghiDinhLuong Where TenND = N'{0}' and IdNDLuong <> {1}", Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from NghiDinhLuong Where TenND = N'{0}'", Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Tên nghị định lương '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_3.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_3
                                Return False
                            End If
                        End If
                    End Using
                End If

            Case 5          'Bảng lương
                If (cb_main_1.SelectedIndex <= 0 And cb_main_1.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn nghị định lương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_main_1
                    Return False
                End If
                If (edt_dl_2.Text.Trim() = "") Then
                    MessageBox.Show("Bảng lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_2
                    Return False
                End If
                If (edt_dl_3.Text.Trim() = "") Then
                    MessageBox.Show("Mô tả chi tiết không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_3
                    Return False
                End If

                'Trong cùng một nghị định lương thì bảng lương không trùng
                If (edt_dl_2.Text.Trim() <> "") Then
                    Dim _IdNdLuong As Integer = 0
                    _IdNdLuong = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Int32)
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from BangLuong Where BangLuong = N'{0}' and IdBangLuong <> {1} and IdND_Luong = {2}", Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()), _RecordId, _IdNdLuong)
                    Else
                        strSQL = String.Format("Select * from BangLuong Where BangLuong = N'{0}' and IdND_Luong = {1}", Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()), _IdNdLuong)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Bảng lương '{0}' thuộc nghị định lương '{1}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_2.Text.Trim().ToString(), cb_main_1.SelectedItem.ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_2
                                Return False
                            End If
                        End If
                    End Using
                End If
                If (edt_dl_3.Text.Trim() <> "") Then
                    Dim _IdNdLuong As Integer = 0
                    _IdNdLuong = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Int32)
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from BangLuong Where MoTa = N'{0}' and IdBangLuong <> {1} and IdND_Luong = {2}", Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()), _RecordId, _IdNdLuong)
                    Else
                        strSQL = String.Format("Select * from BangLuong Where MoTa = N'{0}' and IdND_Luong = {1}", Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()), _IdNdLuong)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Mô tả chi tiết '{0}' của bảng lương thuộc nghị định lương '{1}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_3.Text.Trim().ToString(), cb_main_1.SelectedItem.ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_3
                                Return False
                            End If
                        End If
                    End Using
                End If

            Case 6          'Ngạch lương
                If (cb_main_1.SelectedIndex <= 0 And cb_main_1.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn bảng lương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_main_1
                    Return False
                End If
                If (edt_dl_2.Text.Trim() = "") Then
                    MessageBox.Show("Ngạch lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_2
                    Return False
                End If
                If (edt_dl_3.Text.Trim() = "") Then
                    MessageBox.Show("Phân loại ngạch lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_3
                    Return False
                End If
                If (edt_dl_4.Text.Trim() = "") Then
                    MessageBox.Show("Mô tả chi tiết ngạch lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_4
                    Return False
                End If
                'Bắt điều kiện trùng dữ liệu Ngạch lương trong cùng một bảng lương
                If (edt_dl_2.Text.Trim() <> "") Then
                    Dim _IdBangLuong As Integer = 0
                    _IdBangLuong = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Int32)
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from NgachLuong Where NgachLuong = N'{0}' and IdNgachLuong <> {1} and IdBangLuong = {2}", Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()), _RecordId, _IdBangLuong)
                    Else
                        strSQL = String.Format("Select * from NgachLuong Where NgachLuong = N'{0}' and IdBangLuong = {1}", Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()), _IdBangLuong)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Ngạch lương '{0}' thuộc bảng lương '{1}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_2.Text.Trim().ToString(), cb_main_1.SelectedItem.ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_2
                                Return False
                            End If
                        End If
                    End Using
                End If
                If (edt_dl_4.Text.Trim() <> "") Then
                    Dim _IdBangLuong As Integer = 0
                    _IdBangLuong = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Int32)
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from NgachLuong Where Mota = N'{0}' and IdNgachLuong <> {1} and IdBangLuong = {2}", Globals.Find_Replace(edt_dl_4.Text.Trim().ToString()), _RecordId, _IdBangLuong)
                    Else
                        strSQL = String.Format("Select * from NgachLuong Where Mota = N'{0}' and IdBangLuong = {1}", Globals.Find_Replace(edt_dl_4.Text.Trim().ToString()), _IdBangLuong)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Mô tả chi tiết ngạch lương '{0}' thuộc bảng lương '{1}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_4.Text.Trim().ToString(), cb_main_1.SelectedItem.ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_4
                                Return False
                            End If
                        End If
                    End Using
                End If

            Case 7          'Bậc lương
                If (cb_main_1.SelectedIndex <= 0 And cb_main_1.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn bảng lương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_main_1
                    Return False
                End If
                If (cb_main_2.SelectedIndex <= 0 And cb_main_2.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn ngạch lương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_main_2
                    Return False
                End If

                If (edt_dl_3.Text.Trim() = "") Then
                    MessageBox.Show("Bậc lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_3
                    Return False
                End If
                If (edt_dl_4.Text.Trim() = "") Then
                    MessageBox.Show("Hệ số lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_4
                    Return False
                End If
                If (edt_dl_5.Text.Trim() = "") Then
                    MessageBox.Show("Mô tả chi tiết bậc lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_5
                    Return False
                End If

                If (edt_dl_3.Text.Trim() <> "") Then
                    Dim _IdNgachLuong As Integer = 0
                    _IdNgachLuong = CType(IIf(arrChild.Count > 0, arrChild(cb_main_2.SelectedIndex), "0"), Int32)
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from BacLuong Where BacLuong = N'{0}' and IdBacLuong <> {1} and IdNgachLuong = {2}", Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()), _RecordId, _IdNgachLuong)
                    Else
                        strSQL = String.Format("Select * from BacLuong Where BacLuong = N'{0}' and IdNgachLuong = {1}", Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()), _IdNgachLuong)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Bậc lương '{0}' thuộc ngạch lương '{1}' đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_3.Text.Trim().ToString(), cb_main_2.SelectedItem.ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_dl_3
                                Return False
                            End If
                        End If
                    End Using
                End If

            Case 8          'Lương cơ bản
                If (edt_luongcb.Text.Trim() = "" Or CType(MoneyValue(edt_luongcb.Text.Trim()), Double) <= 0) Then
                    MessageBox.Show("Lương tối thiểu không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_luongcb
                    Return False
                End If
                If (edt_dl_2.Text.Trim() = "") Then
                    MessageBox.Show("Hệ số ngành không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_2
                    Return False
                End If
                If (edt_dl_2.Text.Trim() <> "") Then
                    If (CType(edt_dl_2.Text.Trim(), Double) <= 0) Then
                        MessageBox.Show("Hệ số ngành không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_dl_2
                        Return False
                    End If
                End If
                If (edt_dl_7.Text.Trim() = "") Then
                    MessageBox.Show("Người ký quyết định không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_7
                    Return False
                End If
                If (cb_main_3.SelectedIndex <= 0 And cb_main_3.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn chức vụ người ký!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_main_3
                    Return False
                End If

            Case 9          'Mức phụ cấp
                If (cb_main_1.SelectedIndex <= 0 And cb_main_1.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn loại phụ cấp!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_main_1
                    Return False
                End If
                If (edt_dl_2.Text.Trim() = "") Then
                    MessageBox.Show("Mức phụ cấp không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_2
                    Return False
                End If
                If (edt_dl_2.Text.Trim() <> "") Then
                    If (CType(edt_dl_2.Text.Trim(), Double) <= 0) Then
                        MessageBox.Show("Mức phụ cấp không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_dl_2
                        Return False
                    End If
                End If

            Case 10          'Tham số Lương hệ thống
                If (_RecordId > 0) Then
                    If (edt_dl_1.Text.Trim() = "") Then
                        MessageBox.Show("Tên gọi tham số lương hệ thống không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_dl_1
                        Return False
                    End If
                Else
                    If (cb_main_1.SelectedIndex <= 0) Then
                        MessageBox.Show("Tên gọi tham số lương hệ thống không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_main_1
                        Return False
                    End If
                End If
                
                'Bắt điều kiện trùng tên tham số Hệ thống
                If (_RecordId > 0) Then
                    strSQL = String.Format("Select * from SalaryVar Where NameVar = N'{0}' and IdSalaryVar <> {1} and DateApply = '{2}'", Globals.Find_Replace(edt_dl_1.Text.Trim().ToString()), _RecordId, Format(dtpk_ngayky.Value, "MM/dd/yyyy"))
                Else
                    strSQL = String.Format("Select * from SalaryVar Where NameVar = N'{0}' and DateApply = '{1}'", Globals.Find_Replace(edt_dl_1.Text.Trim().ToString()), Format(dtpk_ngayky.Value, "MM/dd/yyyy"))
                End If
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            MessageBox.Show(String.Format("Tên gọi '{0}' của tham số lương hệ thống đã tồn tại. Vui lòng kiểm tra lại!", edt_dl_1.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                            ActiveControl = edt_dl_1
                            Return False
                        End If
                    End If
                End Using

                If (edt_dl_2.Text.Trim() = "") Then
                    MessageBox.Show("Giá trị tham số không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_2
                    Return False
                End If
                If (edt_dl_3.Text.Trim() = "") Then
                    MessageBox.Show("Tỷ lệ quy định tương ứng với giá trị tham số không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_3
                    Return False
                End If
                'Bắt điều kiện tỷ lệ phải tương ứng với giá trị của tham số
                If (Globals.Splip_Strings(edt_dl_2.Text.Trim().ToString()).Count <> Globals.Splip_Strings(edt_dl_3.Text.Trim().ToString()).Count) Then
                    MessageBox.Show("Tỷ lệ và giá trị tham số phải tương ứng với nhau!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_3
                    Return False
                End If
                If (edt_dl_7.Text.Trim() = "") Then
                    MessageBox.Show("Mô tả chi tiết tham số không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_dl_7
                    Return False
                End If

        End Select
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu khi sửa đổi danh mục lương
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Fill_Data()
        Dim dr As DataRow
        dr = _DanhMuc.GetDataForId(_RecordId, _Status)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                ckb_lock.Checked = IIf(dr("Status").ToString().Trim() = "1", False, True)
                Select Case _Status
                    Case 4              'Nghị định lương
                        edt_dl_1.Text = dr("MaND").ToString()
                        edt_dl_2.Text = dr("SoND").ToString()
                        edt_dl_3.Text = dr("TenND").ToString()
                        edt_dl_4.Text = dr("LoaiND").ToString()
                        edt_dl_5.Text = dr("GhiChu").ToString()
                    Case 5              'Bảng lương
                        cb_main_1.SelectedIndex = CType(arrRoot.IndexOf(dr("IdND_Luong").ToString()), Integer)
                        edt_dl_2.Text = dr("BangLuong").ToString()
                        edt_dl_3.Text = dr("MoTa").ToString()
                        edt_dl_2.ReadOnly = True
                    Case 6              'Ngạch lương
                        cb_main_1.SelectedIndex = CType(arrRoot.IndexOf(dr("IdBangLuong").ToString()), Integer)
                        edt_dl_2.Text = dr("NgachLuong").ToString()
                        edt_dl_3.Text = dr("Loai").ToString()
                        edt_dl_4.Text = dr("Mota").ToString()
                        edt_dl_5.Text = dr("TimeNangNgach").ToString()
                        edt_dl_2.ReadOnly = True
                    Case 7              'Bậc lương
                        Dim _IdNgachLuong As Integer = 0
                        Dim _BangLuongId As Integer = 0
                        _IdNgachLuong = CType(dr("IdNgachLuong").ToString(), Integer)
                        If (_IdNgachLuong > 0) Then
                            _BangLuongId = _Labour.GetIdBangLuong(_IdNgachLuong)
                        End If
                        cb_main_1.SelectedIndex = CType(arrRoot.IndexOf(_BangLuongId.ToString()), Int32)
                        cb_main_1_SelectedIndexChanged(Nothing, Nothing)
                        cb_main_2.SelectedIndex = CType(arrChild.IndexOf(_IdNgachLuong.ToString()), Int32)

                        edt_dl_3.Text = dr("BacLuong").ToString()
                        edt_dl_4.Text = dr("HeSo").ToString()
                        edt_dl_5.Text = dr("Mota").ToString()
                        edt_dl_3.ReadOnly = True
                    Case 8              'Lương cơ bản
                        edt_luongcb.Text = dr("LuongCoBan").ToString().Trim()
                        edt_dl_2.Text = dr("HeSoNganh").ToString().Trim()
                        If dr("NgayHuong").ToString().Trim() <> "" Then
                            dtpk_ngayhuong.Value = CType(dr("NgayHuong").ToString(), DateTime)
                        End If
                        edt_dl_4.Text = dr("SoQD").ToString().Trim()
                        If dr("NgayQD").ToString().Trim() <> "" Then
                            dtpk_ngayky.Value = CType(dr("NgayQD").ToString(), DateTime)
                        End If
                        edt_dl_7.Text = dr("NguoiQD").ToString().Trim()
                        cb_main_3.SelectedIndex = CType(arrRoot.IndexOf(dr("IdCV_Nguoi_QD").ToString()), Integer)
                        edt_dl_9.Text = dr("Ghichu").ToString().Trim()

                    Case 9              'Mức phụ cấp
                        cb_main_1.SelectedIndex = CType(arrRoot.IndexOf(dr("IdLoai_PhC").ToString()), Integer)
                        edt_dl_2.Text = dr("Muc_PhC").ToString().Trim()

                    Case 10              'Tham số lương hệ thống
                        edt_dl_1.ReadOnly = True
                        edt_dl_1.Text = dr("NameVar").ToString().Trim()
                        edt_dl_2.Text = dr("Value").ToString().Trim().Replace("#", "; ")
                        edt_dl_3.Text = dr("Rate").ToString().Trim().Replace("#", "; ")
                        dtpk_ngayky.Value = CType(dr("DateApply").ToString(), DateTime)
                        edt_dl_7.Text = dr("Descript").ToString().Trim().Replace("#", "; ")
                        ckb_lock.Checked = IIf(dr("Status").ToString().Trim() = "1", False, True)

                End Select
            End If
        End If
    End Sub

#End Region

#Region "---> Events main: Các sự kiện chính <---"

    Private Sub DMLKhacForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Create_Controls()
        ResetControls()
        If (_RecordId > 0) Then
            ckb_lientuc.Visible = False
            Fill_Data()
        End If
    End Sub

    Private Sub cb_main_1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_main_1.SelectedIndexChanged
        If (_Status = 7) Then
            cb_main_2.Items.Clear()
            arrChild.Clear()
            If (cb_main_1.SelectedIndex > 0 And cb_main_1.Items.Count <> 0) Then
                Dim BangLuongId As Int32 = CType(arrRoot(IIf(cb_main_1.SelectedIndex > 0, cb_main_1.SelectedIndex, "0")), Int32)
                If (BangLuongId > 0) Then
                    Dim strSQL As String = String.Format("Select IdNgachLuong, Mota From NgachLuong Where IdBangLuong = {0}", BangLuongId)
                    arrChild = _Globals.Bind_ComBoBox(cb_main_2, strSQL, "---Ngạch lương---")
                    strSQL = String.Format("Select * From BangLuong Where IdBangLuong = {0}", BangLuongId)
                End If
            End If
            If (cb_main_2.Items.Count > 0) Then
                cb_main_2.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btn_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_reset.Click
        ResetControls()
        If (_RecordId > 0) Then
            Fill_Data()
        End If
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Close()
    End Sub

    Private Sub DMLKhacForm_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If (IsValid()) Then
            Select Case _Status
                Case 4              'Nghị định lương
                    Dim obj_ngdluong As clsHT_DanhMuc.NghiDinhLuong = New clsHT_DanhMuc.NghiDinhLuong()
                    obj_ngdluong.MaND = Globals.Find_Replace(edt_dl_1.Text.Trim().ToString())
                    obj_ngdluong.SoND = Globals.Find_Replace(edt_dl_2.Text.Trim().ToString())
                    obj_ngdluong.TenND = Globals.Find_Replace(edt_dl_3.Text.Trim().ToString())
                    obj_ngdluong.LoaiND = Globals.Find_Replace(edt_dl_4.Text.Trim().ToString()).ToUpper()
                    obj_ngdluong.GhiChu = Globals.Find_Replace(edt_dl_5.Text.Trim().ToString())
                    obj_ngdluong.Status = IIf(ckb_lock.Checked = True, 0, 1)
                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        obj_ngdluong.IdNDLuong = _RecordId
                        _DanhMuc.Save_NghiDinhLuong(obj_ngdluong)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _DanhMuc.Save_NghiDinhLuong(obj_ngdluong)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        edt_dl_1.Text = ""
                        edt_dl_2.Text = ""
                        edt_dl_3.Text = ""
                        edt_dl_4.Text = ""
                        edt_dl_5.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = edt_dl_1
                        End If
                    End If

                Case 5              'Bảng lương
                    Dim obj_bangluong As clsHT_DanhMuc.BangLuong = New clsHT_DanhMuc.BangLuong()
                    obj_bangluong.IdND_Luong = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Integer)
                    obj_bangluong.BangLuong = Globals.Find_Replace(edt_dl_2.Text.Trim().ToString())
                    obj_bangluong.MoTa = Globals.Find_Replace(edt_dl_3.Text.Trim().ToString())
                    obj_bangluong.Status = IIf(ckb_lock.Checked = True, 0, 1)
                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        obj_bangluong.IdBangLuong = _RecordId
                        _DanhMuc.Save_BangLuong(obj_bangluong)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _DanhMuc.Save_BangLuong(obj_bangluong)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        cb_main_1.SelectedIndex = 0
                        edt_dl_2.Text = ""
                        edt_dl_3.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = cb_main_1
                        End If
                    End If

                Case 6              'Ngạch lương
                    Dim obj_ngachluong As clsHT_DanhMuc.NgachLuong = New clsHT_DanhMuc.NgachLuong()
                    obj_ngachluong.IdBangLuong = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Integer)
                    obj_ngachluong.NgachLuong = Globals.Find_Replace(edt_dl_2.Text.Trim().ToString())
                    obj_ngachluong.Loai = Globals.Find_Replace(edt_dl_3.Text.Trim().ToString())
                    obj_ngachluong.Mota = Globals.Find_Replace(edt_dl_4.Text.Trim().ToString())
                    obj_ngachluong.TimeNangNgach = Globals.Find_Replace(edt_dl_5.Text.Trim().ToString())
                    obj_ngachluong.Status = IIf(ckb_lock.Checked = True, 0, 1)

                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        obj_ngachluong.IdNgachLuong = _RecordId
                        _DanhMuc.Save_NgachLuong(obj_ngachluong)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _DanhMuc.Save_NgachLuong(obj_ngachluong)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        cb_main_1.SelectedIndex = 0
                        edt_dl_2.Text = ""
                        edt_dl_3.Text = ""
                        edt_dl_4.Text = ""
                        edt_dl_5.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = cb_main_1
                        End If
                    End If

                Case 7              'Bậc lương
                    Dim obj_bacluong As clsHT_DanhMuc.BacLuong = New clsHT_DanhMuc.BacLuong()
                    obj_bacluong.IdNgachLuong = CType(IIf(arrChild.Count > 0, arrChild(cb_main_2.SelectedIndex), "0"), Integer)
                    obj_bacluong.BacLuong = IIf(edt_dl_3.Text.ToString() <> "", CType(edt_dl_3.Text.ToString().Trim(), Integer), 0)
                    obj_bacluong.Heso = IIf(edt_dl_4.Text.ToString() <> "", CType(edt_dl_4.Text.ToString().Trim(), Double), 0)
                    obj_bacluong.Mota = Globals.Find_Replace(edt_dl_5.Text.Trim().ToString())
                    obj_bacluong.Status = IIf(ckb_lock.Checked = True, 0, 1)
                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        obj_bacluong.IdBacLuong = _RecordId
                        _DanhMuc.Save_BacLuong(obj_bacluong)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _DanhMuc.Save_BacLuong(obj_bacluong)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        cb_main_1.SelectedIndex = 0
                        cb_main_1_SelectedIndexChanged(Nothing, Nothing)
                        edt_dl_3.Text = ""
                        edt_dl_4.Text = ""
                        edt_dl_5.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = cb_main_1
                        End If
                    End If

                Case 8              'Lương cơ bản
                    Dim obj_luongcb As clsHT_DanhMuc.TienLuong = New clsHT_DanhMuc.TienLuong()
                    obj_luongcb.LuongCoBan = IIf(edt_luongcb.Text.Trim() <> "", CType(MoneyValue(edt_luongcb.Text.Trim()), Double), 0)
                    obj_luongcb.HeSoNganh = IIf(edt_dl_2.Text.ToString() <> "", CType(edt_dl_2.Text.ToString().Trim(), Double), 0)
                    obj_luongcb.NgayHuong = dtpk_ngayhuong.Value
                    obj_luongcb.SoQD = Globals.Find_Replace(edt_dl_4.Text.Trim().ToString())
                    obj_luongcb.NgayQD = dtpk_ngayky.Value
                    obj_luongcb.NguoiQD = Globals.Find_Replace(edt_dl_7.Text.Trim().ToString())
                    obj_luongcb.IdCV_Nguoi_QD = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_3.SelectedIndex), "0"), Integer)
                    obj_luongcb.Ghichu = Globals.Find_Replace(edt_dl_9.Text.Trim().ToString())
                    obj_luongcb.Status = IIf(ckb_lock.Checked = True, 0, 1)

                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        obj_luongcb.IdTienLuong = _RecordId
                        _DanhMuc.Save_TienLuong(obj_luongcb)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _DanhMuc.Save_TienLuong(obj_luongcb)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        edt_luongcb.Text = "0"
                        edt_dl_2.Text = ""
                        dtpk_ngayhuong.Text = DateTime.Now.ToShortDateString()
                        edt_dl_4.Text = ""
                        dtpk_ngayky.Text = DateTime.Now.ToShortDateString()
                        edt_dl_7.Text = ""
                        edt_dl_9.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = edt_luongcb
                        End If
                    End If

                Case 9              'Mức phụ cấp
                    Dim obj_mucphucap As clsHT_DanhMuc.MucPhuCap = New clsHT_DanhMuc.MucPhuCap()
                    obj_mucphucap.IdLoai_PhC = CType(IIf(arrRoot.Count > 0, arrRoot(cb_main_1.SelectedIndex), "0"), Integer)
                    obj_mucphucap.Muc_PhC = IIf(edt_dl_2.Text.ToString() <> "", CType(edt_dl_2.Text.ToString().Trim(), Double), 0)
                    obj_mucphucap.Status = IIf(ckb_lock.Checked = True, 0, 1)


                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        obj_mucphucap.IdMuc_PhC = _RecordId
                        _DanhMuc.Save_MucPhuCap(obj_mucphucap)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _DanhMuc.Save_MucPhuCap(obj_mucphucap)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        cb_main_1.SelectedIndex = 0
                        edt_dl_2.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = cb_main_1
                        End If
                    End If

                Case 10              'Tham số Lương hệ thống

                    Dim _Value As String = Globals.Find_Replace(edt_dl_2.Text.Trim().ToString()).Replace(";", "#").Replace(" ", "")
                    While _Value.EndsWith("#") 'Bỏ dấu chấm phẩy ở cuối chuỗi đi
                        _Value = _Value.Substring(0, _Value.Length - 1)
                    End While

                    Dim _Rate As String = Globals.Find_Replace(edt_dl_3.Text.Trim().ToString()).Replace(";", "#").Replace(" ", "")
                    While _Rate.EndsWith("#") 'Bỏ dấu chấm phẩy ở cuối chuỗi đi
                        _Rate = _Rate.Substring(0, _Rate.Length - 1)
                    End While

                    Dim _Descript As String = Globals.Find_Replace(edt_dl_7.Text.Trim())
                    Dim _Status As Byte = IIf(ckb_lock.Checked = True, 0, 1)
                    Dim _Name As String = ""
                    If (_RecordId > 0) Then         'Trường hợp sửa đổi dữ liệu
                        _Name = Globals.Find_Replace(edt_dl_1.Text.Trim().ToString()).ToUpper()
                        _DanhMuc.Save_SalaryVar(_RecordId, _Name, _Value, _Rate, dtpk_ngayky.Value, _Descript, _Status)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    Else                            'Trường hợp thêm mới dữ liệu
                        _Name = cb_main_1.SelectedItem.ToString().Trim()
                        _DanhMuc.Save_SalaryVar(0, _Name, _Value, _Rate, dtpk_ngayky.Value, _Descript, _Status)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        edt_dl_1.Text = ""
                        edt_dl_2.Text = ""
                        edt_dl_3.Text = ""
                        dtpk_ngayky.Text = DateTime.Now.ToShortDateString()
                        edt_dl_7.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = edt_dl_1
                        End If
                    End If
            End Select
        End If
    End Sub

#End Region

#Region "---> Events main: Các sự kiện ngoại lệ <---"

    Private Sub edt_dl_1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_1.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_dl_2.Focus()
        End If
    End Sub

    Private Sub cb_main_1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_main_1.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (_Status = 7) Then
                cb_main_2.Focus()
            Else
                edt_dl_2.Focus()
            End If
        End If
    End Sub

    Private Sub edt_luongcb_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_luongcb.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_luongcb.Text.Trim() = "") Then
                edt_luongcb.Text = "0"
            End If
            edt_dl_2.Focus()
        End If
    End Sub

    Private Sub edt_luongcb_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_luongcb.Leave
        If (edt_luongcb.Text.Trim() = "") Then
            edt_luongcb.Text = "0"
        End If
    End Sub

    Private Sub edt_luongcb_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_luongcb.TextChanged
        Try
            edt_luongcb = formatMoneyinTextbox(edt_luongcb)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub edt_dl_2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_2.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (_Status = 9) Then
                ckb_lock.Focus()
            ElseIf (_Status = 8) Then
                dtpk_ngayhuong.Focus()
            Else
                edt_dl_3.Focus()
            End If
        End If
    End Sub

    Private Sub edt_dl_2_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_2.KeyUp
        If (e.KeyCode = Keys.Up) Then
            If (_Status = 5 Or _Status = 6 Or _Status = 9) Then
                cb_main_1.Focus()
            ElseIf (_Status = 8) Then
                edt_luongcb.Focus()
            Else
                edt_dl_1.Focus()
            End If
        End If
    End Sub

    Private Sub cb_main_2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_main_2.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_dl_3.Focus()
        End If
    End Sub

    Private Sub dtpk_ngayhuong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngayhuong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_dl_4
        End If
    End Sub

    Private Sub dtpk_ngayky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngayky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_dl_7
        End If
    End Sub

    Private Sub cb_main_3_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_main_3.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_dl_9
        End If
    End Sub

    Private Sub edt_dl_3_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_3.KeyDown
        If (_Status <> 5) Then
            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
                If (_Status = 10) Then
                    dtpk_ngayky.Focus()
                Else
                    edt_dl_4.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub edt_dl_3_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_3.KeyUp
        If (_Status <> 5) Then
            If (e.KeyCode = Keys.Up) Then
                If (_Status = 7) Then
                    ActiveControl = cb_main_2
                Else
                    ActiveControl = edt_dl_2
                End If
            End If
        End If
    End Sub

    Private Sub edt_dl_4_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_4.KeyDown
        If (_Status <> 6) Then
            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
                If (_Status = 8) Then
                    ActiveControl = dtpk_ngayky
                Else
                    ActiveControl = edt_dl_5
                End If

            End If
        End If
    End Sub

    Private Sub edt_dl_4_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_4.KeyUp
        If (_Status <> 6) Then
            If (e.KeyCode = Keys.Up) Then
                If (_Status = 8) Then
                    ActiveControl = dtpk_ngayhuong
                Else
                    ActiveControl = edt_dl_3
                End If
            End If
        End If
    End Sub

    Private Sub edt_dl_5_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_5.KeyDown
        If (_Status <> 4 And _Status <> 7) Then
            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
                ActiveControl = ckb_lock
            End If
        End If
    End Sub

    Private Sub edt_dl_5_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_5.KeyUp
        If (_Status <> 4 And _Status <> 7) Then
            If (e.KeyCode = Keys.Up) Then
                ActiveControl = edt_dl_4
            End If
        End If
    End Sub

    Private Sub edt_dl_7_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_7.KeyDown
        If (_Status = 10) Then
            If (e.KeyCode = Keys.Tab) Then
                ActiveControl = ckb_lock
            End If
        Else
            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
                ActiveControl = cb_main_3
            End If
        End If
    End Sub

    Private Sub edt_dl_7_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dl_7.KeyUp
        If (_Status <> 10) Then
            If (e.KeyCode = Keys.Up) Then
                ActiveControl = dtpk_ngayky
            End If
        End If
    End Sub

    Private Sub ckb_lock_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ckb_lock.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            btn_save.Focus()
        End If
    End Sub

    'Các sự kiện ngoại lệ về nhập dl số
    Private Sub edt_dl_4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dl_4.KeyPress
        If (_Status = 7) Then
            Dim decimalString As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
            Dim decimalChar As Char = Convert.ToChar(decimalString)
            If (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
            ElseIf (e.KeyChar = decimalString And edt_dl_4.Text.IndexOf(decimalString) = -1) Then
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub edt_dl_2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dl_2.KeyPress
        If (_Status = 8 Or _Status = 9) Then
            Dim decimalString As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
            Dim decimalChar As Char = Convert.ToChar(decimalString)
            If (Char.IsDigit(e.KeyChar) Or Char.IsControl(e.KeyChar)) Then
            ElseIf (e.KeyChar = decimalString And edt_dl_2.Text.IndexOf(decimalString) = -1) Then
            Else
                e.Handled = True
            End If
        ElseIf (_Status = 10) Then
            Dim KeyAscii As Integer
            KeyAscii = Asc(e.KeyChar)

            If (KeyAscii > 59 And KeyAscii <> 95) Or KeyAscii = 58 Or (KeyAscii <= 47 And KeyAscii <> 46 And KeyAscii <> 32 And KeyAscii <> 8 And KeyAscii <> 13) Then
                KeyAscii = 0
            End If

            If KeyAscii = 0 Then
                e.Handled = True
            Else
                e.Handled = False
            End If
        End If
    End Sub

    Private Sub edt_dl_5_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dl_5.KeyPress
        If (_Status = 6) Then
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
        End If
    End Sub

    Private Sub edt_dl_3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_dl_3.KeyPress
        If (_Status = 7) Then
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
        ElseIf (_Status = 10) Then
            Dim KeyAscii As Integer
            KeyAscii = Asc(e.KeyChar)

            If (KeyAscii > 59 And KeyAscii <> 95) Or KeyAscii = 58 Or (KeyAscii <= 47 And KeyAscii <> 46 And KeyAscii <> 32 And KeyAscii <> 8 And KeyAscii <> 13) Then
                KeyAscii = 0
            End If

            If KeyAscii = 0 Then
                e.Handled = True
            Else
                e.Handled = False
            End If
        End If
    End Sub
#End Region

End Class