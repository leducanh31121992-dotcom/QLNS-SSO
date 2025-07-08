Public Class frmHS_CanBoTS

#Region "--->Khai báo thuộc tính và Khởi tạo đối tượng<---"
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _Globals As Globals = New Globals()
    Private _nodeCurrent As String = ""
    Private obj_capnhat As frmCN_CanBoTS
    Private obj_hdld_ts As frmHS_HdldTS
    Private obj_viewdetail As frmTS_ChiTiet
    Private ChiNhanhId As Integer = 0
    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

    Private _Records As Int16    'Tổng số bản ghi tìm thấy
    'Lưu danh sách các cán bộ cần chuyển từ lao động dài hạn sang lao động ngắn hạn
    Private arrHuman As ArrayList = New ArrayList()
    Private _IsUpdate As Byte = 0        'Biến lưu trạng thái khi người dùng click thêm mới hay sửa đổi
    'Khai báo các thông tin add CheckBox vào cột tiêu đề chọn cả các items Check trên lưới
    Private ckb_ChoiceAll As CheckBox = Nothing   'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0          'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0   'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False  'Cờ báo việc Checkall
#End Region
    
#Region "--->Functions: Các hàm dùng chính<---"
    Private Sub ResetAll_Controls()
        dgv_main.Rows.Clear()
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
        tv_main.CollapseAll()
        ChiNhanhId = 0
    End Sub

    ''' <summary>
    ''' Hàm thực hiện load lại các thông tin sau khi một sự kiện hoàn thành
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad_Infors()
        Try
            'Load lại khi click vào Hợp đồng Lao động
            tv_main.Nodes.Clear()
            _HS_CanBo.Fill_Tree_HSCBTS(tv_main)

            If _IsUpdate = 3 Then
                ckb_ChoiceAll.Checked = False
                HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                ckb_chon_dscanbo.Checked = False
                ckb_chon_dscanbo_CheckedChanged(Nothing, Nothing)
                If (cb_dscanbo.Items.Count <> 0) Then cb_dscanbo.SelectedIndex = 0
                obj_hdld_ts.Dispose()
                Return
            End If
            'Load lại với trường hợp thêm mới hoặc sửa đổi hay sau khi click thêm, sửa mà lại không cập nhật (Tức click close form cập nhật dl)
            Dim _RecordId As String = obj_capnhat.IdCanBo.Trim().ToString()
            If (_IsUpdate = 0 Or _RecordId = "") Then
                ckb_ChoiceAll.Checked = False
                HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                ckb_chon_dscanbo.Checked = False
                ckb_chon_dscanbo_CheckedChanged(Nothing, Nothing)
                If (cb_dscanbo.Items.Count <> 0) Then cb_dscanbo.SelectedIndex = 0
                obj_capnhat.Dispose()
                Return
            End If
            'Thực hiện tìm đến Node vừa cập nhật
            ResetAll_Controls()
            tv_main.Refresh()
            Dim _isNode As TreeNode = Nothing
            Dim _Name As String = ""
            Dim _ChiNhanhId As Integer = 0
            Dim dr As DataRow
            dr = _HS_CanBo.GetTrainee(0, _RecordId, "", "", CType("1", Byte), CType("0", Byte))
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    _Name = dr("HoTen").ToString().Trim()
                    _ChiNhanhId = dr("IdChiNhanh").ToString()
                    If (dr("IdPhongBan").ToString().Substring(0, 3) = "DV_") Then
                        _ChiNhanhId = dr("IdPhongBan").ToString().Replace("DV_", "")
                    End If
                End If
            End If
            Dim _NodeFind As String = ""
            _NodeFind = "00CN" & _ChiNhanhId
            _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
            If Not (_isNode Is Nothing) Then
                'Select Node cha (Tương ứng node chi nhánh cấp 1 hoặc tương đương)
                tv_main.SelectedNode = _isNode
                'Bắt đầu tìm Node cán bộ thuộc chi nhánh vừa Select
                Dim _isNodeParent As TreeNode = Nothing
                _isNodeParent = _isNode
                _NodeFind = "01CN" + _RecordId.ToString().Trim()
                _isNode = Nothing
                _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_isNode Is Nothing) Then
                    Dim _index As Integer = _isNode.Index
                    _isNodeParent.Nodes.Remove(_isNode)
                    _isNode = _isNodeParent.Nodes.Insert(_index, _RecordId, _Name)
                    _isNode.Tag = "01CN" + _RecordId
                    tv_main.SelectedNode = _isNode
                Else
                    If _IsUpdate = 1 Then
                        _isNode = _isNodeParent.Nodes.Add(_RecordId, _Name)
                        _isNode.Tag = "01CN" + _RecordId
                        tv_main.SelectedNode = _isNode
                    End If
                End If
                obj_capnhat.Dispose()
                _IsUpdate = 0
            End If
        Catch ex As Exception
            MessageBox.Show("Thiết lập lại trạng thái sau khi cập nhật: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện check quyền của thành viền được phép thao tác với các chức năng nhất định
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        pnl_chuyen.Enabled = False
        Dim _roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";183;184;185;186;", _roles)) Then
            btn_hdld_ts.Enabled = False
        End If
        If (_roles.IndexOf(";179;") < 0) Then
            btn_view_detail.Enabled = False
        End If
        If (_roles.IndexOf(";180;") < 0) Then
            btn_add.Enabled = False
        Else
            pnl_chuyen.Enabled = True
        End If
        If (_roles.IndexOf(";181;") < 0) Then
            btn_edit.Enabled = False
        End If
        If (_roles.IndexOf(";182;") < 0) Then
            btn_delete.Enabled = False
            dgv_main.Columns("cln_Choice").Visible = False
            ckb_ChoiceAll.Visible = False
        End If
    End Sub
#End Region
    
#Region "--->Events: Các sự kiện chính<---"
    Private Sub frmHS_CanBoTS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _HS_CanBo.Create_Frame(dgv_main, 1)
        AddHeaderCheckBox()
        'Gọi một số sự kiện liên quan đến Check all items trên lưới dữ liệu
        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged
        AddHandler dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged
        _HS_CanBo.Fill_Tree_HSCBTS(tv_main)
        ResetAll_Controls()
        Check_Permits()
        ckb_chon_dscanbo.Checked = False
        ckb_chon_dscanbo_CheckedChanged(Nothing, Nothing)
        _IsUpdate = 0
        'pnl_chuyen.Visible = False
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        Dim strSQLTS As String = ""
        ckb_ChoiceAll.Checked = False
        dgv_main.Rows.Clear()
        _Records = 0
        Dim _GridRows As Integer = 0
        _nodeCurrent = ""
        ckb_chon_dscanbo.Checked = False
        ckb_chon_dscanbo_CheckedChanged(sender, Nothing)
        If (tv_main.Nodes.Count > 0) Then
            'Thực hiện load dữ liệu các node con
            Dim IdBranch As Integer = 0
            If tv_main.SelectedNode.GetNodeCount(True) = 0 Then
                If Not (tv_main.SelectedNode.IsExpanded) Then
                    If (tv_main.SelectedNode.Tag.ToString().Trim().Substring(0, 4) = "00CN") Then
                        IdBranch = CType(tv_main.SelectedNode.Tag.ToString().Trim().Substring(4), Integer)
                        _HS_CanBo.Fill_Node(tv_main.SelectedNode, IdBranch)
                    End If
                End If
            End If
            _GridRows = 0
            _nodeCurrent = tv_main.SelectedNode.Tag.ToString().Trim()
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00") Then
                ChiNhanhId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                Using db As DataTable = _HS_CanBo.GetAll_Trainee_New(ChiNhanhId, "", "", "", CType("1", Byte), CType("0", Byte))
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_main.Rows.Add()
                                Try
                                    dgv_main.Rows(i).Cells("cln_Id").Value = db.Rows(i)("Id").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_MaCB").Value = db.Rows(i)("MaCB").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_HoTen").Value = db.Rows(i)("HoTen").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_GioiTinh").Value = db.Rows(i)("GioiTinh_HT").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_NgaySinh").Value = db.Rows(i)("NgaySinh_HT").ToString().Trim()
                                    'Lấy thông tin Phòng ban
                                    dgv_main.Rows(i).Cells("cln_PhongBan").Value = db.Rows(i)("DonVi_HT").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_SoCMT").Value = db.Rows(i)("CMT_So").ToString()
                                    dgv_main.Rows(i).Cells("cln_NgayCap").Value = db.Rows(i)("CMT_NgayCap_HT").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_NoiCap").Value = db.Rows(i)("CMT_NoiCap").ToString().Trim()

                                    dgv_main.Rows(i).Cells("cln_DiaChi").Value = db.Rows(i)("ThT_DiaChi_ALL").ToString().Trim()
                                    'Điện thoại của cán bộ. Đầu tiên fill - di động nếu không có di động -> fill điện thoại cơ quan của cán bộ
                                    If db.Rows(i)("DienThoai_DD").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_DD").ToString().Trim()
                                    Else
                                        If db.Rows(i)("DienThoai_CQ").ToString().Trim() <> "" Then
                                            dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_CQ").ToString().Trim()
                                        Else : dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_NR").ToString().Trim()
                                        End If
                                    End If
                                    dgv_main.Rows(i).Cells("cln_Email").Value = db.Rows(i)("Email").ToString().Trim()

                                    'Lấy thông tin Hợp đồng lao động sớm nhất của cán bộ tập sự
                                    dgv_main.Rows(i).Cells("cln_LoaiHinh").Value = clsHS_CanBo.GetLoaiHinh(CType(db.Rows(i)("Loai").ToString().Trim(), Byte))
                                    dgv_main.Rows(i).Cells("cln_HT_TraLuong").Value = db.Rows(i)("HT_TraLuong_HT").ToString().Trim()
                                    If db.Rows(i)("Tungay").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_NgayBD").Value = CType(db.Rows(i)("Tungay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                    Else
                                        dgv_main.Rows(i).Cells("cln_NgayBD").Value = ""
                                    End If
                                    If db.Rows(i)("DenNgay").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_NgayKT").Value = CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                    Else
                                        dgv_main.Rows(i).Cells("cln_NgayKT").Value = ""
                                    End If
                                    dgv_main.Rows(i).Cells("cln_GioLV").Value = db.Rows(i)("TuGio").ToString().Trim() + " - " + db.Rows(i)("DenGio").ToString().Trim()
                                    If (db.Rows(i)("IdBacLuong").ToString().Trim() <> "0") Then
                                        dgv_main.Rows(i).Cells("cln_HT_HuongLuong").Value = "Thang - bảng lương"
                                        dgv_main.Rows(i).Cells("cln_TienLuong").Value = ""
                                        dgv_main.Rows(i).Cells("cln_BacLuong").Value = db.Rows(i)("BacLuong_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_Heso").Value = IIf(db.Rows(i)("Heso").ToString().Trim() <> "", CType(db.Rows(i)("Heso"), Double).ToString("N2"), "0")
                                        dgv_main.Rows(i).Cells("cln_TyLe").Value = db.Rows(i)("TyleHuong").ToString().Trim()
                                    Else
                                        dgv_main.Rows(i).Cells("cln_HT_HuongLuong").Value = "Lương trọn gói"
                                        dgv_main.Rows(i).Cells("cln_BacLuong").Value = ""
                                        dgv_main.Rows(i).Cells("cln_Heso").Value = ""
                                        dgv_main.Rows(i).Cells("cln_TyLe").Value = ""
                                        dgv_main.Rows(i).Cells("cln_TienLuong").Value = IIf(db.Rows(i)("Tien_Luong").ToString().Trim() <> "", Double.Parse(db.Rows(i)("Tien_Luong").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0")
                                    End If
                                    dgv_main.Rows(i).Cells("cln_Cv_DamNhan").Value = db.Rows(i)("CongViec").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_HieuLuc").Value = db.Rows(i)("CanBo_HieuLuc").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_GhiChu_HieuLuc").Value = db.Rows(i)("GhiChu_HieuLuc").ToString().Trim()
                                Catch ex As Exception
                                    MessageBox.Show("Lỗi: " & ex.Message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                End Try
                            Next
                            _Records = db.Rows.Count
                            _GridRows = db.Rows.Count
                        End If
                    End If
                End Using
            End If

            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "01") Then
                _nodeCurrent = tv_main.SelectedNode.Parent.Tag.ToString().Trim()
                If _nodeCurrent = "00CN1" Then
                    ChiNhanhId = 1
                Else
                    If Not (tv_main.SelectedNode.Parent.Parent Is Nothing) Then
                        If Not (tv_main.SelectedNode.Parent.Parent.Tag Is Nothing) Then
                            ChiNhanhId = CType(tv_main.SelectedNode.Parent.Parent.Tag.ToString().Substring(4), Integer)
                        Else
                            ChiNhanhId = CType(tv_main.SelectedNode.Parent.Tag.ToString().Substring(4), Integer)
                        End If
                    Else
                        ChiNhanhId = CType(tv_main.SelectedNode.Parent.Tag.ToString().Substring(4), Integer)
                    End If
                End If

                Dim Code As String = tv_main.SelectedNode.Tag.ToString().Substring(4)
                Using db As DataTable = _HS_CanBo.GetAll_Trainee_New(ChiNhanhId, Code, "", "", CType("1", Byte), CType("0", Byte))
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            _GridRows = db.Rows.Count
                            dgv_main.Rows.Add()
                            dgv_main.Rows(0).Cells("cln_Id").Value = db.Rows(0)("Id").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                            dgv_main.Rows(0).Cells("cln_MaCB").Value = db.Rows(0)("MaCB").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_HoTen").Value = db.Rows(0)("HoTen").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_GioiTinh").Value = db.Rows(0)("GioiTinh_HT").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_NgaySinh").Value = db.Rows(0)("NgaySinh_HT").ToString().Trim()
                            'Lấy thông tin Phòng ban
                            dgv_main.Rows(0).Cells("cln_PhongBan").Value = db.Rows(0)("DonVi_HT").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_SoCMT").Value = db.Rows(0)("CMT_So").ToString()
                            dgv_main.Rows(0).Cells("cln_NgayCap").Value = db.Rows(0)("CMT_NgayCap_HT").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_NoiCap").Value = db.Rows(0)("CMT_NoiCap").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_DiaChi").Value = db.Rows(0)("ThT_DiaChi_ALL").ToString().Trim()
                            'Điện thoại của cán bộ. Đầu tiên fill - di động nếu không có di động -> fill điện thoại cơ quan của cán bộ
                            If db.Rows(0)("DienThoai_DD").ToString().Trim() <> "" Then
                                dgv_main.Rows(0).Cells("cln_DienThoai").Value = db.Rows(0)("DienThoai_DD").ToString().Trim()
                            Else
                                If db.Rows(0)("DienThoai_CQ").ToString().Trim() <> "" Then
                                    dgv_main.Rows(0).Cells("cln_DienThoai").Value = db.Rows(0)("DienThoai_CQ").ToString().Trim()
                                Else : dgv_main.Rows(0).Cells("cln_DienThoai").Value = db.Rows(0)("DienThoai_NR").ToString().Trim()
                                End If
                            End If
                            dgv_main.Rows(0).Cells("cln_Email").Value = db.Rows(0)("Email").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_LoaiHinh").Value = clsHS_CanBo.GetLoaiHinh(CType(db.Rows(0)("Loai").ToString().Trim(), Byte))
                            dgv_main.Rows(0).Cells("cln_HT_TraLuong").Value = db.Rows(0)("HT_TraLuong_HT").ToString().Trim()
                            If db.Rows(0)("Tungay").ToString().Trim() <> "" Then
                                dgv_main.Rows(0).Cells("cln_NgayBD").Value = CType(db.Rows(0)("Tungay").ToString(), DateTime).ToString("dd-MM-yyyy")
                            Else
                                dgv_main.Rows(0).Cells("cln_NgayBD").Value = ""
                            End If
                            If db.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                dgv_main.Rows(0).Cells("cln_NgayKT").Value = CType(db.Rows(0)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                            Else
                                dgv_main.Rows(0).Cells("cln_NgayKT").Value = ""
                            End If
                            dgv_main.Rows(0).Cells("cln_GioLV").Value = db.Rows(0)("TuGio").ToString().Trim() + " - " + db.Rows(0)("DenGio").ToString().Trim()
                            If (db.Rows(0)("IdBacLuong").ToString().Trim() <> "0") Then
                                dgv_main.Rows(0).Cells("cln_HT_HuongLuong").Value = "Thang - bảng lương"
                                dgv_main.Rows(0).Cells("cln_TienLuong").Value = ""

                                dgv_main.Rows(0).Cells("cln_BacLuong").Value = db.Rows(0)("BacLuong_HT").ToString().Trim()
                                dgv_main.Rows(0).Cells("cln_Heso").Value = IIf(db.Rows(0)("Heso").ToString().Trim() <> "", CType(db.Rows(0)("Heso"), Double).ToString("N2"), "0")
                                dgv_main.Rows(0).Cells("cln_TyLe").Value = db.Rows(0)("TyleHuong").ToString().Trim()
                            Else
                                dgv_main.Rows(0).Cells("cln_HT_HuongLuong").Value = "Lương trọn gói"
                                dgv_main.Rows(0).Cells("cln_BacLuong").Value = ""
                                dgv_main.Rows(0).Cells("cln_Heso").Value = ""
                                dgv_main.Rows(0).Cells("cln_TyLe").Value = ""
                                dgv_main.Rows(0).Cells("cln_TienLuong").Value = IIf(db.Rows(0)("Tien_Luong").ToString().Trim() <> "", Double.Parse(db.Rows(0)("Tien_Luong").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0")
                            End If
                            dgv_main.Rows(0).Cells("cln_Cv_DamNhan").Value = db.Rows(0)("CongViec").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_HieuLuc").Value = db.Rows(0)("CanBo_HieuLuc").ToString().Trim()
                            dgv_main.Rows(0).Cells("cln_GhiChu_HieuLuc").Value = db.Rows(0)("GhiChu_HieuLuc").ToString().Trim()
                        End If
                    End If
                End Using

            End If

            'Bôi đậm mầu các dòng bản ghi cán bộ Ngắn hạn hết hiệu lực HĐLD
            If _GridRows > 0 Then
                For i As Integer = 0 To dgv_main.Rows.Count - 1
                    dgv_main.Rows(i).DefaultCellStyle.ForeColor = Color.Navy
                    Dim sHieuLuc As String = dgv_main.Item("cln_HieuLuc", i).Value.ToString().Trim()
                    If sHieuLuc = "0" Then
                        'dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(97, 135, 214)
                        dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.Orange
                        dgv_main.Rows(i).DefaultCellStyle.ForeColor = System.Drawing.Color.Red
                    End If
                Next
            End If

            tv_main.SelectedNode.Expand()
            dgv_main_CellClick(sender, Nothing)
        End If
        'Thực hiện fill danh sách cán bộ theo đơn vị vừa select
        cb_dscanbo.Items.Clear()
        arrHuman.Clear()
        If (_nodeCurrent <> "") Then
            Dim strSQL As String = String.Format("Select IdCanBo, HoTen From HS_CanBo Where (IdNew = '{0}' Or IdNew Is NULL) And IdDonVi = {1} Order by HoTen", "", CType(_nodeCurrent.Substring(4).ToString().Trim(), Integer))
            arrHuman = _Globals.Bind_ComBoBox(cb_dscanbo, strSQL, "---Danh sách cán bộ---")
            cb_dscanbo.SelectedIndex = 0
        End If
        TotalCheckBoxes = dgv_main.RowCount
        TotalCheckedCheckBoxes = 0
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        obj_capnhat = New frmCN_CanBoTS()
        'Kiểm tra xem có phải chuyển từ LĐ dài hạn sang LĐ ngắn hạn không
        If (cb_dscanbo.Items.Count <> 0 And ckb_chon_dscanbo.Checked = True) Then
            If (cb_dscanbo.SelectedIndex > 0) Then
                obj_capnhat.IdCB = IIf(arrHuman.Count > 0, arrHuman(cb_dscanbo.SelectedIndex), "")
            Else
                cb_dscanbo.Focus()
                Return
            End If
        End If
        obj_capnhat._FlagEvent = 1
        obj_capnhat.IdCanBo = ""
        obj_capnhat.Node = _nodeCurrent
        _IsUpdate = 1
        obj_capnhat.Progress_Changed = New frmCN_CanBoTS.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
        obj_capnhat.ShowDialog()
    End Sub

    Private Sub btn_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit.Click
        If (dgv_main.Rows.Count > 0) Then
            obj_capnhat = New frmCN_CanBoTS
            obj_capnhat.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
            obj_capnhat._FlagEvent = 2
            _IsUpdate = 2
            obj_capnhat.Node = _nodeCurrent
            obj_capnhat.Progress_Changed = New frmCN_CanBoTS.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
            obj_capnhat.ShowDialog()
        Else
            MessageBox.Show("Bạn chưa chọn dữ liệu cán bộ tập sự cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btn_view_detail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_view_detail.Click
        Try
            If (dgv_main.Rows.Count > 0) Then
                'Lấy danh sách mảng các id hiện có trên lưới dl
                Dim arrRows As ArrayList = New ArrayList()
                For i As Integer = 0 To dgv_main.Rows.Count - 1
                    arrRows.Add(dgv_main.Rows(i).Cells("cln_Id").Value.ToString())
                Next
                obj_viewdetail = New frmTS_ChiTiet
                obj_viewdetail.RecordCurrent = CType(dgv_main.CurrentRow.Index, Integer) 'Bản ghi hiện tại đang được Click hay select
                If (_Records = 0) Then
                    obj_viewdetail.Records = 1
                Else
                    obj_viewdetail.Records = _Records 'Tổng số bản ghi đang hiển thị ra trên lưới dữ liệu
                End If

                obj_viewdetail.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
                obj_viewdetail.arr_RecordId = arrRows
                _IsUpdate = 0
                obj_viewdetail.ShowDialog()
            Else
                MessageBox.Show("Bạn chưa chọn dữ liệu cán bộ tập sự cần xem!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xem chi tiết hồ sơ cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If (dgv_main.Rows.Count <= 0) Then Return
        Try
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _count As Int16 = 0
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_MaCB").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_MaCB").Value.ToString() <> "")) Then
                    For i As Integer = 0 To dgv_main.Rows.Count - 1
                        If (dgv_main.Rows(i).Cells("cln_MaCB").Value IsNot Nothing) Then
                            If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                _count += 1
                                arr_Del.Add(dgv_main.Rows(i).Cells("cln_Id").Value.ToString())
                            End If
                        End If
                    Next
                End If
            End If
            If (_count > 0) Then
                Dim _mess As String = IIf(_count = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ cán bộ tập sự đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    Dim arrNode As ArrayList = New ArrayList()

                    For i As Int16 = 0 To arr_Del.Count - 1
                        Dim _IdTSDel As String = arr_Del(i).ToString().Trim()
                        arrNode.Add(_IdTSDel)

                        'Thực hiện xoá dữ liệu ảnh Hồ sơ cán bộ tập sự
                        Dim _pathDel As String = ""
                        Dim dr As DataRow
                        dr = _HS_CanBo.GetTrainee(0, _IdTSDel, "", "", CType("1", Byte), CType("1", Byte))
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                _pathDel = dr("AnhThe").ToString().Trim()
                            End If
                        End If

                        'Thực hiện xoá dữ liệu Hố sơ cán bộ tập sự
                        _HS_CanBo.Delete_HSCB_TS(_IdTSDel)
                        'Kiểm tra xem đây có phải là cán bộ trước chuyển từ Cán bộ chính sang lao động ngắn hạn không ?
                        'Nếu phải thực hiện Update lại IdNew trong bảng HS_CanBo thành NULL
                        If (_IdTSDel <> "") Then
                            'Thực hiện cập nhật idnew (id của cán bộ tập sự) vào hồ sơ cán bộ
                            _HS_CanBo.UpdateIdNew_HS_CanBo(_IdTSDel, "", 2)
                        End If

                        'Xoá ảnh lưu trong thư mục đính kèm
                        _pathDel = String.Format("{0}{1}", Application.StartupPath, _pathDel)
                        Globals.FileDao.DeleteFile(_pathDel)

                        'Kiểm tra xem trong thư mục đính kèm này còn file nào không--> Nếu không con file ảnh nào thì xoá luôn thư mục
                        Dim n As Integer = _pathDel.LastIndexOf("\") + 1
                        Dim strDir As String = _pathDel.Substring(0, n - 1)
                        Dim countFiles As Int16 = Globals.FileDao.GetCountFiles(strDir)
                        If (countFiles = 0) Then
                            Globals.FileDao.DeleteFolder(strDir)
                        End If
                    Next

                    'Thực hiện select lại treeview dữ liệu Cơ cấu tổ chức
                    Dim _isNodeParent As TreeNode = Nothing
                    '--> Duyết mảng id cán bộ tập sự vừa xoá song.
                    If arrNode.Count > 0 Then
                        For i As Integer = 0 To arrNode.Count - 1
                            Dim _TagNode As String = "01CN" + arrNode(i).ToString()
                            Dim _IsNode As TreeNode = Globals.TreeViewFindNode(tv_main.Nodes, _TagNode)
                            If Not (_IsNode Is Nothing) Then
                                _isNodeParent = _IsNode.Parent
                                _isNodeParent.Nodes.Remove(_IsNode)
                            End If
                        Next
                    End If

                    ResetAll_Controls()
                    ckb_chon_dscanbo.Checked = False
                    ckb_chon_dscanbo_CheckedChanged(sender, Nothing)
                    If Not (_isNodeParent Is Nothing) Then
                        tv_main.SelectedNode = _isNodeParent
                    End If
                    MessageBox.Show("Bạn đã xoá thành công thông tin hồ sơ cán bộ tập sự!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                Else
                    arr_Del.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            Else
                MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_hdld_ts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_hdld_ts.Click
        If (dgv_main.Rows.Count > 0) Then
            obj_hdld_ts = New frmHS_HdldTS
            obj_hdld_ts.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
            obj_hdld_ts.Node = _nodeCurrent
            _IsUpdate = 3
            obj_hdld_ts.Progress_Changed = New frmHS_HdldTS.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
            obj_hdld_ts.ShowDialog()
        Else
            MessageBox.Show("Bạn chưa chọn dữ liệu cán bộ tập sự!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub ckb_chon_dscanbo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_chon_dscanbo.CheckedChanged
        If (ckb_chon_dscanbo.Checked = True) Then
            cb_dscanbo.Enabled = True
            lbl_div_cmd0.Visible = False
            btn_edit.Visible = False
            lbl_div_cmd1.Visible = False
            btn_delete.Visible = False
            lbl_div_cmd2.Visible = False
            btn_view_detail.Visible = False
            lbl_div_cmd3.Visible = False
            btn_hdld_ts.Visible = False
            btn_add.Text = "&Chuyển"
            btn_add.Width = 71
        Else
            cb_dscanbo.Enabled = False
            lbl_div_cmd0.Visible = True
            btn_edit.Visible = True
            lbl_div_cmd1.Visible = True
            btn_delete.Visible = True
            lbl_div_cmd2.Visible = True
            btn_view_detail.Visible = True
            lbl_div_cmd3.Visible = True
            btn_hdld_ts.Visible = True
            btn_add.Text = "&Thêm"
            btn_add.Width = 61
        End If
    End Sub

    Private Sub dgv_main_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellClick
        pnl_chuyen.Enabled = False
        If (dgv_main.Rows.Count > 0) Then
            pnl_chuyen.Enabled = True
            btn_edit.Enabled = True
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                Dim sIdNew As String = ""
                sIdNew = SoftSqlHelper.GetString(String.Format("Select Top 1 IsNull(IdNew,'') From HSCB_TS Where Id='{0}'", dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()), "")
                If (sIdNew <> "") Then
                    pnl_chuyen.Enabled = False
                    btn_edit.Enabled = False
                End If
            End If
        End If
    End Sub

    Private Sub dgv_main_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Close()
    End Sub

    Private Sub frmHS_CanBoTS_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
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

End Class
