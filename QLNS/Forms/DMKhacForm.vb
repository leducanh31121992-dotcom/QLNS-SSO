Public Class DMKhacForm

#Region "---> Defined parametter and properties <---"
    Dim _DanhMuc As clsHT_DanhMuc = New clsHT_DanhMuc

    'Chỉ số xác định giao diện hiển thị. Quy định 1-Hệ thống DM. 2-Chi nhánh. 3-Địa danh
    Public _FlagState As Byte

    'Chỉ số xác định khi _FlagState = 1 --> Quy định: 0-Danh mục chung  1-Phòng ban  2-Quốc gia
    Public _FlagChild As Byte
    'Trạng thái xác định: False - Thêm mới. True - Sửa đổi
    Public ValStatus As Boolean

    Private _RootId As String
    Public Property RootId() As String
        Get
            Return _RootId
        End Get
        Set(ByVal value As String)
            _RootId = value
        End Set
    End Property

    Private _RecordId As String
    Public Property RecordId() As String
        Get
            Return _RecordId
        End Get
        Set(ByVal value As String)
            _RecordId = value
        End Set
    End Property

    Private _SqlHelper As DBAccess

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        _SqlHelper = New DBAccess
    End Sub

    'Mảng chứa chỉ số xác định các items trực thuộc
    Private arrTT_Code() As String = {"1", "2", "3", "4", "5", "7", "9"}

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

#Region "---> Functions: Các hàm dùng chung <---"
    Private Sub ResetControls()
        edt_code.Text = ""
        edt_name.Text = ""
        ckb_lock.Checked = False
        ckb_lientuc.Checked = False
        pnl_tructhuoc.Visible = False
        lbl_div_1.Visible = False   'Luôn luôn ẩn
        lbl_div_0.Visible = False
        If (_FlagState = 2) Then    'Trường hợp là chi nhánh
            lbl_div_0.Visible = True
            edt_alias.Text = ""
            edt_address.Text = ""
            edt_tel.Text = ""
            edt_email.Text = ""
            edt_website.Text = ""
            edt_unitnumber.Text = "0"
            Me.Height = 317
            Me.Width = 515
            pnl_branch.Visible = True
        Else                        'Trường hợp còn lại
            pnl_branch.Visible = False
            Me.Height = 191
            Me.Width = 441
            If (_RootId = 100) Then 'Trường hợp cập nhật Phòng ban - Có trực thuộc
                Me.Height = 304
                Me.Width = 485
                pnl_tructhuoc.Visible = True
                lbl_div_0.Visible = True
                pnl_tructhuoc.Height = 123
            Else
                lbl_div.Height = 10
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra điều kiện cập nhật
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        If (edt_code.Text.Trim() = "") Then
            MessageBox.Show("Mã hiệu danh mục không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_code
            Return False
        End If
        If (edt_name.Text.Trim() = "") Then
            MessageBox.Show("Tên gọi danh mục không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_name
            Return False
        End If
        If (_FlagState = 2) Then
            If (edt_alias.Text.Trim() = "") Then
                MessageBox.Show("Tên viết tắt chi nhánh không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = edt_alias
                Return False
            End If
            If (edt_email.Text.Trim() <> "") Then
                If (Not Globals.IsEmail(edt_email.Text.Trim().ToString())) Then
                    MessageBox.Show("Địa chỉ e-mail của cán bộ không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_email
                    Return False
                End If
            End If
        End If

        'Kiểm tra dữ liệu trùng mã hiệu và tên gọi
        Dim strSQL As String = ""
        Select Case _FlagState
            Case 1  'HỆ THỐNG DANH MỤC CHUNG
                If (_FlagChild = 0) Then
                    If (edt_code.Text.Trim() <> "") Then
                        If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi
                            strSQL = String.Format("Select * From DanhMuc Where ma_so = '{0}' and id <> {1} and id_goc = {2}", Globals.Find_Replace(edt_code.Text.Trim()), _RecordId, _RootId)
                        Else
                            strSQL = String.Format("Select * From DanhMuc Where ma_so = '{0}' and id_goc = {1}", Globals.Find_Replace(edt_code.Text.Trim()), _RootId)
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show(String.Format("Mã hiệu '{0}' của danh mục đã tồn tại. Vui lòng kiểm tra lại!", edt_code.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_code
                                    Return False
                                End If
                            End If
                        End Using
                    End If

                    If (edt_name.Text.Trim() <> "") Then
                        If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi
                            strSQL = String.Format("Select * from DanhMuc Where ten_goi = N'{0}' and id <> {1} and id_goc = {2}", Globals.Find_Replace(edt_name.Text.Trim()), _RecordId, _RootId)
                        Else
                            strSQL = String.Format("Select * from DanhMuc Where ten_goi = N'{0}' and id_goc = {1}", Globals.Find_Replace(edt_name.Text.Trim()), _RootId)
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show(String.Format("Tên gọi '{0}' của danh mục đã tồn tại. Vui lòng kiểm tra lại!", edt_name.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_name
                                    Return False
                                End If
                            End If
                        End Using
                    End If
                ElseIf (_FlagChild = 1) Then    'Phòng ban
                    If (edt_code.Text.Trim() <> "") Then
                        If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi
                            strSQL = String.Format("Select * from PhongBan Where ma_so = '{0}' and id <> {1}", Globals.Find_Replace(edt_code.Text.Trim()), _RecordId)
                        Else
                            strSQL = String.Format("Select * from PhongBan Where ma_so = '{0}'", Globals.Find_Replace(edt_code.Text.Trim()))
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show(String.Format("Mã hiệu Phòng ban '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_code.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_code
                                    Return False
                                End If
                            End If
                        End Using
                    End If

                    If (edt_name.Text.Trim() <> "") Then
                        If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi
                            strSQL = String.Format("Select * from PhongBan Where ten_phong = N'{0}' and id <> {1}", Globals.Find_Replace(edt_name.Text.Trim()), _RecordId)
                        Else
                            strSQL = String.Format("Select * from PhongBan Where ten_phong = N'{0}'", Globals.Find_Replace(edt_name.Text.Trim()))
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show(String.Format("Tên gọi Phòng ban '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_name.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_name
                                    Return False
                                End If
                            End If
                        End Using
                    End If
                Else                            'Quốc gia
                    If (edt_code.Text.Trim() <> "") Then
                        If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi
                            strSQL = String.Format("Select * from QuocGia Where ma_so = '{0}' and id <> {1}", Globals.Find_Replace(edt_code.Text.Trim()), _RecordId)
                        Else
                            strSQL = String.Format("Select * from QuocGia Where ma_so = '{0}'", Globals.Find_Replace(edt_code.Text.Trim()))
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show(String.Format("Mã hiệu Quốc gia '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_code.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_code
                                    Return False
                                End If
                            End If
                        End Using
                    End If

                    If (edt_name.Text.Trim() <> "") Then
                        If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi
                            strSQL = String.Format("Select * from QuocGia Where ten_goi = N'{0}' and id <> {1}", Globals.Find_Replace(edt_name.Text.Trim()), _RecordId)
                        Else
                            strSQL = String.Format("Select * from QuocGia Where ten_goi = N'{0}'", Globals.Find_Replace(edt_name.Text.Trim()))
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    MessageBox.Show(String.Format("Tên gọi Quốc gia '{0}' đã tồn tại. Vui lòng kiểm tra lại!", edt_name.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_name
                                    Return False
                                End If
                            End If
                        End Using
                    End If
                End If
            Case 2  'CHI NHÁNH
                'Giới hạn không cho tên viết tắt vượt quá 4 ký tự
                If (edt_alias.Text.Trim().Length > 4) Then
                    MessageBox.Show("Tên viết tắt chi nhánh không được vượt quá 4 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_alias
                    Return False
                End If
                'Bắt trùng: Tên gọi - Mã hiệu - Tên viết tắt - Địa chỉ e-mail (Nêu có địa chỉ e-mail)
                'Bắt trường hợp trung mã --> Không cho phép trùng mã trong tất cả các chi nhánh
                If (edt_code.Text.Trim() <> "") Then
                    If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi địa danh
                        strSQL = String.Format("Select * from ChiNhanh Where ma_so = '{0}' and id <> {1}", Globals.Find_Replace(edt_code.Text.Trim()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from ChiNhanh Where ma_so = '{0}'", Globals.Find_Replace(edt_code.Text.Trim()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Mã hiệu '{0}' của chi nhánh đã tồn tại. Vui lòng kiểm tra lại!", edt_code.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_code
                                Return False
                            End If
                        End If
                    End Using
                End If

                'Bắt trùng tên Chi nhánh
                If (edt_name.Text.Trim() <> "") Then
                    'Trường hợp: _RootId = 1 hoặc 0 đặc biệt: Không cho phép trùng với bất kỳ tên của Chi nhánh nào
                    If (ValStatus = True) Then
                        strSQL = String.Format("Select * from ChiNhanh Where ten_goi = N'{0}' and id <> {1}", Globals.Find_Replace(edt_name.Text.Trim()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from ChiNhanh Where ten_goi = N'{0}'", Globals.Find_Replace(edt_name.Text.Trim()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Tên gọi '{0}' của chi nhánh đã tồn tại. Vui lòng kiểm tra lại!", edt_name.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_name
                                Return False
                            End If
                        End If
                    End Using
                End If
                'Bắt trùng tên viết tắt
                If (edt_alias.Text.Trim() <> "") Then
                    If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi địa danh
                        strSQL = String.Format("Select * from ChiNhanh Where ten_vt = N'{0}' and id <> {1}", Globals.Find_Replace(edt_alias.Text.Trim()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from ChiNhanh Where ten_vt = N'{0}'", Globals.Find_Replace(edt_alias.Text.Trim()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Tên viết tắt '{0}' của chi nhánh đã tồn tại. Vui lòng kiểm tra lại!", edt_alias.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_alias
                                Return False
                            End If
                        End If
                    End Using
                End If
                'Bắt trùng tên địa chỉ e-mail nếu đã nhập
                If (edt_email.Text.Trim() <> "") Then
                    If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi địa danh
                        strSQL = String.Format("Select * from ChiNhanh Where email = N'{0}' and id <> {1}", Globals.Find_Replace(edt_email.Text.Trim()), _RecordId)
                    Else
                        strSQL = String.Format("Select * from ChiNhanh Where email = N'{0}'", Globals.Find_Replace(edt_email.Text.Trim()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Địa chỉ e-mail '{0}' của chi nhánh đã tồn tại. Vui lòng kiểm tra lại!", edt_email.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_email
                                Return False
                            End If
                        End If
                    End Using
                End If
            Case 3  'ĐỊA DANH
                If (edt_code.Text.Trim() <> "") Then
                    If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi địa danh
                        strSQL = String.Format("Select * From DiaDanh Where ma_so = '{0}' and id <> {1}", Globals.Find_Replace(edt_code.Text.Trim()), _RecordId)
                    Else
                        strSQL = String.Format("Select * From DiaDanh Where ma_so = '{0}'", Globals.Find_Replace(edt_code.Text.Trim()))
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Mã hiệu '{0}' của địa danh đã tồn tại. Vui lòng kiểm tra lại!", edt_code.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_code
                                Return False
                            End If
                        End If
                    End Using
                End If
                If (edt_name.Text.Trim() <> "") Then
                    If (ValStatus = True) Then   'Bắt trùng trong trường hợp sửa đổi địa danh
                        strSQL = String.Format("Select * From DiaDanh Where ten_goi = N'{0}' And id <> {1} and id_goc = {2}", Globals.Find_Replace(edt_name.Text.Trim()), _RecordId, _RootId)
                    Else
                        strSQL = String.Format("Select * From DiaDanh Where ten_goi = N'{0}' And id_goc = {1}", Globals.Find_Replace(edt_name.Text.Trim()), _RootId)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Tên gọi '{0}' của địa danh đã tồn tại. Vui lòng kiểm tra lại!", edt_name.Text.Trim().ToString()), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_name
                                Return False
                            End If
                        End If
                    End Using
                End If
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Hàm lấy chuỗi id trực thuộc khi cập nhật Phòng ban
    ''' </summary>
    ''' <returns>Chuỗi id trực thuộc</returns>
    ''' <remarks></remarks>
    Private Function GetTTId() As String
        Dim result As String = ""
        If (clb_tructhuoc.Items.Count > 0) Then
            For Each indexChecked As Int32 In clb_tructhuoc.CheckedIndices
                result = result + arrTT_Code(indexChecked).ToString() + ","
            Next
            If (result <> "") Then
                While result.EndsWith(",") 'Bỏ dấu chấm phẩy ở cuối chuỗi đi
                    result = result.Substring(0, result.Length - 1)
                End While
            End If
        End If
        Return result
    End Function

    ''' <summary>
    ''' Hàm thực thiện Checked các items tương ứng theo chuỗi id
    ''' </summary>
    ''' <param name="strId">Chuỗi id</param>
    ''' <remarks></remarks>
    Private Sub SetItemsTT(ByVal strId As String)
        If (strId = "") Then Return

        Dim arrCode As ArrayList = New ArrayList
        For j As Integer = 0 To arrTT_Code.Length - 1
            arrCode.Add(arrTT_Code(j))
        Next

        Dim strTemp As String = strId + ","
        Dim i As Int32 = 0
        While (strTemp.Trim() <> "")
            i = strTemp.IndexOf(",")
            If (i > 0) Then
                Dim k As Int32 = CType(arrCode.IndexOf(strTemp.Substring(0, i)), Int32)
                clb_tructhuoc.SetItemChecked(k, True)
                strTemp = strTemp.Substring(i + 1)
            Else
                strTemp = ""
            End If
        End While
    End Sub

    ''' <summary>
    ''' Hàm thực hiện load dữ liệu ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Fill_Data()
        If (ValStatus = False) Then     'Trường hợp thêm mới
            edt_code.Text = _DanhMuc.GetCode(_RootId, _FlagState, _FlagChild, _Node)
            ckb_lientuc.Visible = True
            'If edt_code.Text.Trim().Substring(Len(edt_code.Text.Trim()) - 2) = "00" Then
            '    edt_unitnumber.ReadOnly = True
            'End If
        Else                            'Trường hợp sửa đổi
            edt_unitnumber.ReadOnly = False
            edt_code.ReadOnly = True
            ckb_lientuc.Visible = False
            'Thực hiện load dữ liệu ra sửa đổi
            Select Case _FlagState
                Case 1      'Hệ thống danh mục
                    If (_RootId = 100) Then         'Phòng ban
                        Dim dr As DataRow
                        dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 1)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                edt_code.Text = dr("ma_so").ToString().Trim()
                                edt_name.Text = dr("ten_goi").ToString().Trim()
                                SetItemsTT(dr("truc_thuoc").ToString().Trim())
                                ckb_lock.Checked = IIf(dr("Status").ToString().Trim() = "1", False, True)
                            End If
                        End If
                    ElseIf (_RootId = 101) Then     'Quốc gia
                        Dim dr As DataRow
                        dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 2)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                edt_code.Text = dr("ma_so").ToString().Trim()
                                edt_name.Text = dr("ten_goi").ToString().Trim()
                                ckb_lock.Checked = IIf(dr("Status").ToString().Trim() = "1", False, True)
                            End If
                        End If
                    Else                            'Danh mục chung
                        Dim dr As DataRow
                        dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 0)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                edt_code.Text = dr("ma_so").ToString().Trim()
                                edt_name.Text = dr("ten_goi").ToString().Trim()
                                ckb_lock.Checked = IIf(dr("Status").ToString().Trim() = "1", False, True)
                            End If
                        End If
                    End If
                Case 2      'Chi nhánh
                    Dim strSQL As String = String.Format("Select * From ChiNhanh Where id = {0}", _RecordId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                Dim sCode As String = db.Rows(0)("ma_so").ToString().Trim()
                                edt_code.Text = db.Rows(0)("ma_so").ToString().Trim()
                                edt_name.Text = db.Rows(0)("ten_goi").ToString().Trim()
                                edt_alias.Text = db.Rows(0)("ten_vt").ToString().Trim()
                                edt_address.Text = db.Rows(0)("dia_chi").ToString().Trim()
                                edt_tel.Text = db.Rows(0)("dien_thoai").ToString().Trim()
                                edt_email.Text = db.Rows(0)("email").ToString().Trim()
                                edt_website.Text = db.Rows(0)("website").ToString().Trim()
                                lbl_soxaphuong.Text = "Số xã phường  "
                                lbl_soxaphuong.Width = 107
                                If sCode.Substring(0, 2) <> "00" And sCode.Substring(0, 2) <> "10" Then
                                    'Xét trường hợp là Chi nhánh tỉnh (mã có 2 ký tự cuối là 00)
                                    'If sCode.Substring(Len(sCode) - 2) = "00" Then
                                    '    lbl_soxaphuong.Text = "Số lượng xã (phường) trong tỉnh  "
                                    '    lbl_soxaphuong.Width = 221
                                    '    edt_unitnumber.Text = _DanhMuc.GetVillagers(_RecordId)
                                    '    edt_unitnumber.ReadOnly = True
                                    'Else    'Xét trường hợp là Phòng giao dịch

                                    'End If
                                    edt_unitnumber.Text = IIf(db.Rows(0)("solg_xaphuong").ToString().Trim() <> "", db.Rows(0)("solg_xaphuong").ToString().Trim(), "0")
                                Else
                                    edt_unitnumber.Text = 0
                                End If
                                ckb_lock.Checked = IIf(db.Rows(0)("Status").ToString().Trim() = "1", False, True)
                            End If
                        End If
                    End Using
                Case 3      'Địa danh
                    If (_RootId = 0) Then   'Sửa đổi dl - Địa danh Tỉnh - TP
                        Dim strSQL As String = String.Format("Select * from DiaDanh Where id = {0}", _RecordId)
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    edt_code.Text = db.Rows(0)("ma_so").ToString().Trim()
                                    edt_name.Text = db.Rows(0)("ten_goi").ToString().Trim()
                                    ckb_lock.Checked = IIf(db.Rows(0)("Status").ToString().Trim() = "1", False, True)
                                End If
                            End If
                        End Using
                    Else
                        Dim dr As DataRow
                        dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 4)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                edt_code.Text = dr("ma_so").ToString().Trim()
                                edt_name.Text = dr("ten_goi").ToString().Trim()
                                ckb_lock.Checked = IIf(dr("Status").ToString().Trim() = "1", False, True)
                            End If
                        End If
                    End If
            End Select
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện dùng chung <---"
    Private Sub DMKhacForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ResetControls()
        Fill_Data()
        ActiveControl = edt_code
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If (IsValid()) Then
            Select Case _FlagState
                Case 1  'HỆ THỐNG DANH MỤC CHUNG
                    If (_RootId = 100) Then         'Xét trường hợp cập nhật thêm mới Phòng ban
                        Dim obj_debt As clsHT_DanhMuc.PhongBan = New clsHT_DanhMuc.PhongBan()
                        obj_debt.Code = Globals.Find_Replace(edt_code.Text.Trim().ToString())
                        obj_debt.Name = Globals.Find_Replace(edt_name.Text.Trim().ToString())
                        obj_debt.TrucThuoc = GetTTId()
                        obj_debt.Status = IIf(ckb_lock.Checked = True, 0, 1)
                        If (ValStatus = False) Then     'Trường hợp cập nhật Thêm mới
                            _DanhMuc.Save_PhongBan(obj_debt)
                            If (ckb_lientuc.Checked = False) Then
                                If Not (Progress_Changed Is Nothing) Then
                                    Progress_Changed()
                                End If
                                Hide()
                            End If
                            'Load lại trạng thái cho Controls
                            edt_code.Text = ""
                            edt_code.Text = _DanhMuc.GetCode(_RootId, _FlagState, _FlagChild, _Node)
                            edt_name.Text = ""
                            ckb_lock.Checked = False
                            If (clb_tructhuoc.Items.Count > 0) Then
                                For i As Int32 = 0 To clb_tructhuoc.Items.Count - 1
                                    clb_tructhuoc.SetItemCheckState(i, CheckState.Unchecked)
                                Next
                            End If
                        Else                           'Trường hợp cập nhật Sửa đổi
                            obj_debt.Id = _RecordId
                            _DanhMuc.Save_PhongBan(obj_debt)
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                    ElseIf (_RootId = 101) Then     'Xét trường hợp cập nhật thêm mới Quốc gia
                        Dim obj_quocgia As clsHT_DanhMuc.QuocGia = New clsHT_DanhMuc.QuocGia()
                        obj_quocgia.Code = Globals.Find_Replace(edt_code.Text.Trim().ToString())
                        obj_quocgia.Name = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_name.Text.Trim().ToString()))
                        obj_quocgia.Status = IIf(ckb_lock.Checked = True, 0, 1)
                        If (ValStatus = False) Then     'Trường hợp cập nhật Thêm mới
                            _DanhMuc.Save_QuocGia(obj_quocgia)
                            If (ckb_lientuc.Checked = False) Then
                                If Not (Progress_Changed Is Nothing) Then
                                    Progress_Changed()
                                End If
                                Hide()
                            End If
                            'Load lại trạng thái cho Controls
                            edt_code.Text = ""
                            edt_code.Text = _DanhMuc.GetCode(0, _FlagState, _FlagChild, _Node)
                            edt_name.Text = ""
                            ckb_lock.Checked = False
                        Else                           'Trường hợp cập nhật Sửa đổi
                            obj_quocgia.Id = _RecordId
                            _DanhMuc.Save_QuocGia(obj_quocgia)
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                    Else                            'Xét trường hợp cập nhật thêm mới Danh mục chung
                        Dim obj_list As clsHT_DanhMuc.DanhMuc = New clsHT_DanhMuc.DanhMuc()
                        obj_list.Code = Globals.Find_Replace(edt_code.Text.Trim().ToString())
                        obj_list.Name = Globals.Find_Replace(edt_name.Text.Trim().ToString())
                        obj_list.Status = IIf(ckb_lock.Checked = True, 0, 1)
                        obj_list.RootId = _RootId
                        If (ValStatus = False) Then     'Trường hợp cập nhật Thêm mới
                            _DanhMuc.Save_List(obj_list)
                            If (ckb_lientuc.Checked = False) Then
                                If Not (Progress_Changed Is Nothing) Then
                                    Progress_Changed()
                                End If
                                Close()
                            End If
                            'Load lại trạng thái cho Controls
                            edt_code.Text = ""
                            edt_code.Text = _DanhMuc.GetCode(_RootId, _FlagState, _FlagChild, _Node)
                            edt_name.Text = ""
                            ckb_lock.Checked = False
                            If (ckb_lientuc.Checked = True) Then
                                ActiveControl = edt_code
                            End If
                        Else                           'Trường hợp cập nhật Sửa đổi
                            obj_list.Id = _RecordId
                            _DanhMuc.Save_List(obj_list)
                            If (Progress_Changed IsNot Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                    End If

                Case 2  'DANH MỤC - CHI NHÁNH
                    Dim obj_branch As clsHT_DanhMuc.ChiNhanh = New clsHT_DanhMuc.ChiNhanh()
                    obj_branch.Code = Globals.Find_Replace(edt_code.Text.Trim().ToString())
                    obj_branch.Name = Globals.Find_Replace(edt_name.Text.Trim().ToString())
                    obj_branch.AliasBranch = Globals.Find_Replace(edt_alias.Text.Trim().ToString())
                    obj_branch.Address = Globals.Find_Replace(edt_address.Text.Trim().ToString())
                    obj_branch.Tel = Globals.Find_Replace(edt_tel.Text.Trim().ToString())
                    obj_branch.Email = Globals.Find_Replace(edt_email.Text.Trim().ToString())
                    obj_branch.Web = Globals.Find_Replace(edt_website.Text.Trim().ToString())
                    obj_branch.Status = IIf(ckb_lock.Checked = True, 0, 1)
                    If obj_branch.Code <> "" And obj_branch.Code.Substring(Len(obj_branch.Code) - 2) <> "00" Then
                        obj_branch.Villagers = IIf(edt_unitnumber.Text.Trim() <> "", CType(edt_unitnumber.Text.Trim(), Integer), 0)
                    Else        'Nếu là tỉnh thì không cập nhật số lượng xã phường
                        obj_branch.Villagers = 0
                    End If
                    obj_branch.RootId = _RootId
                    If (ValStatus = False) Then     'Trường hợp cập nhật Thêm mới
                        _DanhMuc.Save_Branch(obj_branch)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        edt_code.Text = ""
                        edt_code.Text = _DanhMuc.GetCode(_RootId, _FlagState, _FlagChild, _Node)
                        edt_name.Text = ""
                        edt_alias.Text = ""
                        edt_address.Text = ""
                        edt_tel.Text = ""
                        edt_email.Text = ""
                        edt_website.Text = ""
                        edt_unitnumber.Text = "0"
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = edt_code
                        End If
                    Else                           'Trường hợp cập nhật Sửa đổi
                        obj_branch.Id = _RecordId
                        _DanhMuc.Save_Branch(obj_branch)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    End If
                Case 3  'DANH MỤC - ĐỊA DANH
                    Dim obj_place As clsHT_DanhMuc.DiaDanh = New clsHT_DanhMuc.DiaDanh()
                    obj_place.Code = Globals.Find_Replace(edt_code.Text.Trim().ToString())
                    obj_place.Name = Globals.Find_Replace(edt_name.Text.Trim().ToString())
                    obj_place.Status = IIf(ckb_lock.Checked = True, 0, 1)
                    obj_place.RootId = _RootId
                    'CẬP NHẬT ĐỊA DANH
                    If (ValStatus = False) Then     'Trường hợp cập nhật Thêm mới
                        _DanhMuc.Save_DiaDanh(obj_place)
                        If (ckb_lientuc.Checked = False) Then
                            If Not (Progress_Changed Is Nothing) Then
                                Progress_Changed()
                            End If
                            Close()
                        End If
                        'Load lại trạng thái cho Controls
                        edt_code.Text = ""
                        edt_code.Text = _DanhMuc.GetCode(_RootId, _FlagState, _FlagChild, _Node)
                        edt_name.Text = ""
                        ckb_lock.Checked = False
                        If (ckb_lientuc.Checked = True) Then
                            ActiveControl = edt_code
                        End If
                    Else                           'Trường hợp cập nhật Sửa đổi
                        obj_place.Id = _RecordId
                        _DanhMuc.Save_DiaDanh(obj_place)
                        If Not (Progress_Changed Is Nothing) Then
                            Progress_Changed()
                        End If
                        Close()
                    End If
            End Select
        End If
    End Sub

    Private Sub btn_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_reset.Click
        Fill_Data()
    End Sub

    Private Sub DMKhacForm_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Close()
    End Sub
#End Region

#Region "---> Events: Các sự kiện Ngoại lệ người dùng <---"
    Private Sub edt_code_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_code.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_code.Text.Trim() <> "") Then
                edt_name.Focus()
            Else
                edt_code.Focus()
            End If
        End If
    End Sub

    Private Sub edt_name_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_name.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_name.Text.Trim() <> "") Then
                If (_FlagState = 3) Then
                    edt_alias.Focus()
                Else
                    If (_FlagChild = 1) Then
                        clb_tructhuoc.Focus()
                    Else
                        btn_save.Focus()
                    End If
                End If
            Else
                edt_name.Focus()
            End If
        End If
    End Sub

    Private Sub edt_name_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_name.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_code
        End If
    End Sub

    Private Sub clb_tructhuoc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles clb_tructhuoc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            btn_save.Focus()
        End If
    End Sub

    Private Sub edt_alias_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_alias.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_address
        End If
    End Sub

    Private Sub edt_alias_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_alias.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_name
        End If
    End Sub

    Private Sub edt_address_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_address.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_tel
        End If
    End Sub

    Private Sub edt_address_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_address.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_alias
        End If
    End Sub

    Private Sub edt_tel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_tel.KeyPress
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

    Private Sub edt_tel_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tel.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_email
        End If
    End Sub

    Private Sub edt_tel_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tel.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_address
        End If
    End Sub

    Private Sub edt_email_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_email.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_website
        End If
    End Sub

    Private Sub edt_email_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_email.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_tel
        End If
    End Sub

    Private Sub edt_website_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_website.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_unitnumber
        End If
    End Sub

    Private Sub edt_website_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_website.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_email
        End If
    End Sub

    Private Sub edt_unitnumber_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_unitnumber.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If edt_unitnumber.Text.Trim() = "" Then
                edt_unitnumber.Text = "0"
            End If
            ActiveControl = btn_save
        End If
    End Sub

    Private Sub edt_unitnumber_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_unitnumber.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_website
        End If
    End Sub

    Private Sub edt_unitnumber_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_unitnumber.KeyPress
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

    Private Sub edt_unitnumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_unitnumber.Leave
        If edt_unitnumber.Text.Trim() = "" Then
            edt_unitnumber.Text = "0"
        End If
    End Sub
#End Region

End Class