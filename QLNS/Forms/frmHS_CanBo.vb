Public Class frmHS_CanBo

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _ListDocument As clsHT_DanhMuc = New clsHT_DanhMuc()
    Private _Globals As Globals = New Globals()
    Private arrTapsu As ArrayList = New ArrayList
    Private strNode As String = ""
    Private obj_capnhat As frmCN_HSCanBo
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _Records As Int16    'Tổng số bản ghi tìm thấy
    'Biến lưu lại đơn vị Hiện tại của cán bộ khi Select tới TreeView
    Private _IdDonviHT As Integer = IdDONVI
    Private _BranchTag As String = ""
    Private _IsUpdate As Byte = 0       'Biến lưu trạng thái người dùng click thêm mới hay sửa đổi dữ liệu (1: Thêm mới - 2: Sửa đổi)
    Private _IdNew As String = ""       'Lưu lại giá trị IdNew  của cán bộ nếu có

    'Khai báo các thông tin add CheckBox vào cột tiêu đề chọn cả các items Check trên lưới
    Private ckb_ChoiceAll As CheckBox = Nothing   'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0          'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0   'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False  'Cờ báo việc Checkall

#End Region

#Region "---> Functions: Các hàm chính dùng chung <---"
    ''' <summary>
    ''' Hàm thực hiện reset lại trạng thái ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Load_Infors()
        rb_hsnhansu.Checked = True
        dgv_main.Rows.Clear()
        ckb_chon_dltapsu.Checked = False
        ckb_chon_dltapsu_CheckedChanged(Nothing, Nothing)
        'tv_main.CollapseAll()
        _IdNew = ""
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
    End Sub

    ''' <summary>
    ''' Hàm thực hiện load lại các thông tin sau khi một sự kiện hoàn thành
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad_Infors()
        Try
            Dim vIdHuman As String = obj_capnhat.IdCanBo.ToString().Trim()
            'Trường hợp không liên quan đến cập nhật hoặc thêm mới mà không thêm lại thực hiện click close form cập nhật dữ liệu
            If (_IsUpdate = 0 Or vIdHuman = "") Then
                rb_hsnhansu.Checked = True
                ckb_ChoiceAll.Checked = False
                HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                ckb_chon_dltapsu.Checked = False
                ckb_chon_dltapsu_CheckedChanged(Nothing, Nothing)
                If (cb_cbtapsu.Items.Count <> 0) Then cb_cbtapsu.SelectedIndex = 0
                obj_capnhat.Dispose()
                Return
            End If
            Load_Infors()
            tv_main.Refresh()
            'Thực hiện lấy thông tin IdCanBo, IdDonVi, IdPhongBan
            Dim _isNode As TreeNode = Nothing
            Dim _DonviId As Integer = 0         'Id đơn vị của cán bộ
            Dim _DonviNewId As Integer = 0         'Id đơn vị của cán bộ
            Dim _CodeBranch As String = ""      'Mã hiệu của chi nhánh cap tinh hoac tuong duong
            Dim _PhongBanId As Integer = 0      'Id phòng ban của cán bộ
            Dim _Name As String = ""
            Dim dr As DataRow
            dr = _HS_CanBo.GetHuman_ForCode(vIdHuman)
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    _Name = dr("HoTen").ToString().Trim()
                    _DonviId = IIf(dr("IdDonVi").ToString().Trim() <> "", CType(dr("IdDonVi").ToString().Trim(), Integer), 0)
                    If (_DonviId > 0) Then
                        'Lấy mã hiệu đơn vị của cán bộ
                        _CodeBranch = _ListDocument.GetCodeForId(_DonviId)
                        Using db As DataTable = _HS_CanBo.GetAll_Decision(vIdHuman, _DonviId, 1)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    _DonviNewId = IIf(db.Rows(0)("IdDonVi_Moi").ToString().Trim() <> "", CType(db.Rows(0)("IdDonVi_Moi").ToString().Trim(), Integer), 0)
                                    _PhongBanId = IIf(db.Rows(0)("IdPhong_Moi").ToString().Trim() <> "", CType(db.Rows(0)("IdPhong_Moi").ToString().Trim(), Integer), 0)
                                End If
                            End If
                        End Using
                    End If
                End If
            End If
            'Bắt đầu xử lý đến node trên cây dữ liệu (Tìm kiếm và select node hoặc add node nếu trường hợp thêm mới dữ liệu)
            '   Nếu là Cài đặt tại Hội sở chính (Nhìn thấy được toàn bộ trong hệ thống)
            Dim _NodeFind As String = ""
            If (DONVI = gMaDonViTW) Then
                _NodeFind = "ROOT_1_" & gMaDonViTW
                'Thực hiện tìm Cấp 0
                _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_isNode Is Nothing) Then
                    tv_main.SelectedNode = _isNode
                    _NodeFind = "DV_" + _DonviId.ToString() + "_" + _CodeBranch
                End If
            Else
                _NodeFind = "ROOT_" + _DonviId.ToString() + "_" + _CodeBranch
            End If

            'Tìm cấp 1 (Gốc root: Ví dụ "ROOT_42_60900" (Nếu chỉ cài đặt tại chi nhánh) Hoặc "DV_42_60900" (Nếu cài đặt tại hội sở))
            If (_NodeFind <> "") Then
                _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_isNode Is Nothing) Then
                    tv_main.SelectedNode = _isNode
                End If
            End If
            'Tìm cấp 2: Kiểm tra xem có cấp đến PGD không hay chỉ Phòng ban của Chi nhánh hoặc HSC
            Dim _isNodeParent As TreeNode = Nothing     'Lấy node cha của cán bộ để nếu thêm mới thực hiện add cán bộ đó vào node này
            If (_DonviId <> _DonviNewId) Then       'Có cấp tới PGD ("DV_459_60906")
                _NodeFind = "DV_" + _DonviNewId.ToString() + "_" + _ListDocument.GetCodeForId(_DonviNewId).ToString().Trim()
                If (_NodeFind <> "") Then
                    _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                    If Not (_isNode Is Nothing) Then
                        tv_main.SelectedNode = _isNode
                        ' Tìm cấp 3: Tìm đến phòng ban của PGD ("PB_459_21")
                        _NodeFind = "PB_" + _DonviNewId.ToString() + "_" + _PhongBanId.ToString()
                        If (_NodeFind <> "") Then
                            _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                            If Not (_isNode Is Nothing) Then
                                tv_main.SelectedNode = _isNode
                            End If
                        End If
                    End If
                End If
            Else        'Tìm tới cấp Phòng ban của chi nhánh Tỉnh hoặc tương đương ("PB_42_17")
                _NodeFind = "PB_" + _DonviId.ToString() + "_" + _PhongBanId.ToString().Trim()
                If (_NodeFind <> "") Then
                    _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                    If Not (_isNode Is Nothing) Then
                        tv_main.SelectedNode = _isNode
                    End If
                End If
            End If

            _isNodeParent = _isNode
            'Bắt đầu tìm đến cán bộ vừa cập nhật
            _NodeFind = "CB_" + vIdHuman
            _isNode = Nothing
            _isNode = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
            If Not (_isNode Is Nothing) Then
                Dim _index As Integer = _isNode.Index
                _isNodeParent.Nodes.Remove(_isNode)
                _isNode = _isNodeParent.Nodes.Insert(_index, vIdHuman, _Name)
                _isNode.Tag = "CB_" + vIdHuman
                tv_main.SelectedNode = _isNode
            Else
                If _IsUpdate = 1 Then
                    _isNode = _isNodeParent.Nodes.Add(vIdHuman, _Name)
                    _isNode.Tag = "CB_" + vIdHuman
                    tv_main.SelectedNode = _isNode
                End If
            End If
            obj_capnhat.Dispose()
            _IsUpdate = 0
        Catch ex As Exception
            'MessageBox.Show("Thiết lập lại thông tin sau khi cập nhật hồ sơ cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện phân quyền thành viên thao tác với module
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        pnl_hs_tapsu.Enabled = False
        Dim _roles As String = Globals.Roles
        'Nếu không có quyền nào liên quan đến thông tin cán bộ
        If Not (Globals.IsIntersect(";71;72;73;74;", _roles)) Then
            rb_hskhac.Checked = True
            rb_hskhac_CheckedChanged(Nothing, Nothing)
            pnl_luachon.Visible = False
            btn_delete.Enabled = False
            dgv_main.Columns("cln_Choice").Visible = False
            ckb_ChoiceAll.Visible = False
            btn_view_detail.Enabled = False
        End If
        'Nếu không có quyền nào liên quan đến hồ sơ khác
        If Not (Globals.IsIntersect(";75;76;77;78;79;80;81;82;83;84;85;86;87;88;89;90;91;92;93;94;95;96;97;98;", _roles)) Then
            rb_hsnhansu.Checked = True
            rb_hsnhansu_CheckedChanged(Nothing, Nothing)
            pnl_luachon.Visible = False
        End If
        'Xem - Thêm - Sửa - Xoá hồ sơ cán bộ
        If (_roles.IndexOf(";71;") < 0) Then
            btn_view_detail.Enabled = False
        End If
        If (_roles.IndexOf(";72;") < 0) Then
            btn_add.Enabled = False
        Else
            pnl_hs_tapsu.Enabled = True
        End If
        If (_roles.IndexOf(";73;") < 0) Then
            btn_edit.Enabled = False
        End If
        If (_roles.IndexOf(";74;") < 0) Then
            btn_delete.Enabled = False
            dgv_main.Columns("cln_Choice").Visible = False
            ckb_ChoiceAll.Visible = False
        End If

        'Thực hiện Check quyền thao tác với các menu

        '1. Với quyền thao tác Công tác Đoàn thể và quá trình sinh hoạt
        If Not (Globals.IsIntersect(";163;164;165;166;220;221;222;223;225;226;227;228;", _roles)) Then
            mnu_ctdang_dt.Enabled = False
        End If
        '2. Với quyền thao tác Hợp đồng lao động của Cán bộ
        If Not (Globals.IsIntersect(";135;136;137;138;", _roles)) Then
            mnu_hdld.Enabled = False
        End If
        '3. Với quyền thao tác Hồ sơ Chính sách Lao động của Cán bộ
        If Not (Globals.IsIntersect(";119;120;121;122;123;124;125;126;127;128;129;130;131;132;133;134;", _roles)) Then
            mnu_cslaodong.Enabled = False
        Else
            If Not (Globals.IsIntersect(";119;120;121;122;", _roles)) Then
                mnu_csld_bhyt.Enabled = False
            End If
            If Not (Globals.IsIntersect(";123;124;125;126;", _roles)) Then
                mnu_csld_khamsk.Enabled = False
            End If
            If Not (Globals.IsIntersect(";127;128;129;130;", _roles)) Then
                mnu_csld_qlnc.Enabled = False
            End If
            If Not (Globals.IsIntersect(";131;132;133;134;", _roles)) Then
                mnu_csld_hstainan.Enabled = False
            End If
        End If

        '4. Với quyền thao tác Hồ sơ Công tác của cán bộ
        If Not (Globals.IsIntersect(";75;76;77;78;", _roles)) Then
            mnu_hscongtac.Enabled = False
        End If
        '5. Với quyền thao tác Hồ sơ Xuất ngoại của cán bộ
        If Not (Globals.IsIntersect(";79;80;81;82;", _roles)) Then
            mnu_hsxuatngoai.Enabled = False
        End If
        '6. Với quyền thao tác Hồ sơ Hộ chiếu của cán bộ
        If Not (Globals.IsIntersect(";83;84;85;86;", _roles)) Then
            mnu_hshochieu.Enabled = False
        End If
        '7. Với quyền thao tác Hồ sơ Lực lượng vũ trang của cán bộ
        If Not (Globals.IsIntersect(";87;88;89;90;", _roles)) Then
            mnu_hsllvt.Enabled = False
        End If
        '8. Với quyền thao tác Hồ sơ cũ của cán bộ
        If Not (Globals.IsIntersect(";91;92;93;94;", _roles)) Then
            mnu_hscu.Enabled = False
        End If
        '9. Với quyền thao tác Hồ sơ Học hàm - Học vị của cán bộ
        If Not (Globals.IsIntersect(";95;96;97;98;", _roles)) Then
            mnu_hhhv.Enabled = False
        End If
        '10. Với quyền thao tác Hồ sơ Gia đình cán bộ
        If Not (Globals.IsIntersect(";167;168;169;170;171;172;173;174;175;176;177;178;", _roles)) Then
            mnu_hsgdcanbo.Enabled = False
        End If
        '10. Với quyền thao tác Xuất dữ liệu lý lịch cán bộ
        If Not (Globals.IsIntersect(";229;", _roles)) Then
            mnu_exportcv.Enabled = False
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHS_CanBo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _HS_CanBo.Create_Frame(dgv_main, 0)
        AddHeaderCheckBox()
        'Gọi một số sự kiện liên quan đến Check all items trên lưới dữ liệu
        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged
        AddHandler dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged

        _HS_CanBo.Fill_Tree(tv_main)
        'If tv_main.Nodes.Count > 0 Then tv_main.Nodes(0).ExpandAll()

        arrTapsu.Clear()
        Load_Infors()
        _BranchTag = ""
        Check_Permits()
        _IsUpdate = 0
        'If DONVI = gMaDonViTW Then
        '    btn_add.Enabled = False
        'End If
        'Chỉ user cấp Admin mới được thiết lập quyền sử dụng
        btn_QuyenSD.Enabled = (Globals.Group.ToUpper = "OPERATOR" Or Globals.Group.ToUpper = "U01")
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        obj_capnhat = New frmCN_HSCanBo()
        obj_capnhat.FlagHS = IIf(rb_hsnhansu.Checked = True, False, True)
        obj_capnhat.IdCanBo = ""
        ' Nếu thêm mới các hồ sơ liên quan nhân sự - Thì kiểm tra xem đã chọn một nhân sự nào đó trên lưới chưa
        If (rb_hskhac.Checked = True) Then
            If (dgv_main.Rows.Count > 0) Then
                obj_capnhat.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
            Else
                MessageBox.Show("Hãy chọn cán bộ trên lưới mà bạn muốn cập nhật hồ sơ khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                dgv_main.Focus()
                Return
            End If
        End If
        If (cb_cbtapsu.Items.Count <> 0 And ckb_chon_dltapsu.Checked = True) Then
            If (cb_cbtapsu.SelectedIndex > 0) Then
                obj_capnhat.IdCB_TS = IIf(arrTapsu.Count > 0, arrTapsu(cb_cbtapsu.SelectedIndex), "")
            Else
                MessageBox.Show("Hãy chọn cán bộ mà bạn muốn chuyển!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                cb_cbtapsu.Focus()
                Return
            End If
        End If
        obj_capnhat.FlagEvent = 1
        obj_capnhat.Node = strNode
        If (_IdDonviHT > 0) Then obj_capnhat._IdDonviHT = _IdDonviHT
        _IsUpdate = 1
        obj_capnhat.Progress_Changed = New frmCN_HSCanBo.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
        obj_capnhat.ShowDialog()
    End Sub

    Private Sub btn_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit.Click
        If (dgv_main.Rows.Count > 0) Then
            obj_capnhat = New frmCN_HSCanBo
            obj_capnhat.FlagHS = IIf(rb_hsnhansu.Checked = True, False, True)
            obj_capnhat.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
            obj_capnhat.FlagEvent = 2
            If rb_hsnhansu.Checked = True Then _IsUpdate = 2
            obj_capnhat.Node = strNode
            obj_capnhat._IdDonviHT = _IdDonviHT
            obj_capnhat.Progress_Changed = New frmCN_HSCanBo.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
            obj_capnhat.ShowDialog()
        Else
            MessageBox.Show("Bạn chưa chọn dữ liệu nhân sự cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If (dgv_main.Rows.Count <= 0) Then Return
        Try
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _count As Int32 = 0
            Dim _FlagUse As Boolean = False 'False: Thực hiện xoá được. True: Không xoá được
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_MaCB").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_MaCB").Value.ToString() <> "")) Then
                    arr_Del.Add(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString())
                    _count += 1
                    For i As Int32 = 0 To dgv_main.Rows.Count - 1
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
                If (MessageBox.Show("             Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ cán bộ đã chọn không?" & vbCrLf & "(Chú ý: Khi thực hiện xoá hồ sơ cán bộ thì các hồ sơ khác liên quan của cán bộ cũng sẽ được xoá)", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    'Thực hiện xoá hồ sơ cán bộ và các hồ sơ liên quan của cán bộ
                    Dim arrNode As ArrayList = New ArrayList()
                    Dim strSQL As String = ""
                    For i As Int32 = 0 To arr_Del.Count - 1
                        arrNode.Add(arr_Del(i).ToString())
                        ' Lấy id đơn vị từ id cán bộ
                        Dim strPath As String = ""
                        Dim dr As DataRow
                        dr = _HS_CanBo.GetHuman_ForCode(arr_Del(i).ToString())
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                strPath = dr("AnhThe").ToString()
                            End If
                        End If
                        _HS_CanBo.Delete_Human(arr_Del(i).ToString())
                        'Thực hiện update lại bảng HSCB_TS nếu cán bộ này được chuyển từ tập sự sang
                        _HS_CanBo.Update_HSCB_TS_WhenMove("", arr_Del(i).ToString(), 0, 1)
                        'Thực hiện xoá dữ liệu ảnh hồ sơ cán bộ
                        strPath = String.Format("{0}{1}", Application.StartupPath, strPath)
                        Globals.FileDao.DeleteFile(strPath)

                        'Kiểm tra xem trong thư mục đính kèm này còn file nào không--> Nếu không con file ảnh nào thì xoá luôn thư mục
                        Dim n As Int32 = strPath.LastIndexOf("\") + 1
                        Dim strDir As String = strPath.Substring(0, n - 1)
                        Dim countFiles As Int16 = Globals.FileDao.GetCountFiles(strDir)
                        If (countFiles = 0) Then
                            Globals.FileDao.DeleteFolder(strDir)
                        End If
                    Next
                    'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                    MessageBox.Show("Bạn đã xoá thành công hồ sơ của cán bộ và các hồ sơ liên quan đến cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                    rb_hsnhansu.Checked = True
                    dgv_main.Rows.Clear()
                    'Thực hiên Remove các node vừa xoá bỏ
                    If arrNode.Count > 0 Then
                        For i As Int32 = 0 To arrNode.Count - 1
                            Dim nodeHuman As TreeNode = Globals.TreeViewFindNode(tv_main.Nodes, "CB_" + arrNode(i).ToString())
                            If Not (nodeHuman Is Nothing) Then
                                Dim parentNode As TreeNode = nodeHuman.Parent
                                parentNode.Nodes.Remove(nodeHuman)
                            End If
                        Next
                    End If
                    tv_main.CollapseAll()
                    'Select item của cây dữ liệu
                    Dim _node As TreeNode = Nothing
                    If (strNode <> "") Then
                        _node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (_node Is Nothing) Then
                            tv_main.SelectedNode = _node
                        End If
                    End If
                Else ' Không xoá nữa -> Bắt khi người dùng chọn No
                    arr_Del.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            Else
                MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ cán: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_view_detail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_view_detail.Click
        Try
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    If (Globals.Roles.IndexOf(";71;") < 0) Then
                        Return
                    End If
                    Dim obj_detail As frmHS_ChiTiet = New frmHS_ChiTiet()
                    'Lấy danh sách mảng các id hiện có trên lưới dl
                    Dim arrRows As ArrayList = New ArrayList()
                    For i As Integer = 0 To dgv_main.Rows.Count - 1
                        arrRows.Add(dgv_main.Rows(i).Cells("cln_Id").Value.ToString())
                    Next
                    obj_detail.RecordCurrent = CType(dgv_main.CurrentRow.Index, Int32) 'Bản ghi hiện tại đang được Click hay select
                    If (_Records = 0) Then
                        obj_detail.Records = 1
                    Else
                        obj_detail.Records = _Records 'Tổng số bản ghi đang hiển thị ra trên lưới dữ liệu
                    End If
                    obj_detail.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
                    obj_detail._IdDonviHT = _IdDonviHT
                    obj_detail.arr_RecordId = arrRows
                    obj_detail.ShowDialog()
                End If
            Else
                MessageBox.Show("Bạn chưa chọn dữ liệu cán bộ cần xem!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xem chi tiết hồ sơ cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        'My_MessageBox(tv_main.SelectedNode.Tag & ".")

        _Records = 0
        ckb_ChoiceAll.Checked = False
        dgv_main.Rows.Clear()
        strNode = ""
        _BranchTag = ""
        ckb_chon_dltapsu.Checked = False
        ckb_chon_dltapsu_CheckedChanged(sender, Nothing)
        If (tv_main.Nodes.Count > 0) Then
            Dim arrElement() As String
            If tv_main.SelectedNode.GetNodeCount(True) = 0 Then
                If Not (tv_main.SelectedNode.IsExpanded) Then
                    arrElement = tv_main.SelectedNode.Tag.ToString().Trim().Split("_")
                    If (arrElement.Length > 0) Then
                        If (arrElement(0) = "DV") Then
                            _HS_CanBo.Fill_Node(tv_main.SelectedNode, arrElement(1), arrElement(2))
                        ElseIf arrElement(0) = "PB" Then
                            _HS_CanBo.Fill_NodeCanbo(tv_main.SelectedNode, arrElement(1), arrElement(2), True)
                        End If
                    End If
                End If
            End If
            strNode = tv_main.SelectedNode.Tag.ToString()
            _BranchTag = strNode
            'Thực hiện fill dữ liệu danh sách cán bộ ra lưới dữ liệu
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "PB") Then
                _BranchTag = tv_main.SelectedNode.Parent.Tag.ToString()
                dgv_main.Rows.Clear()
                'Lấy Id - Phòng ban từ Tag của treeview
                Dim deptId As Int32 = CType(tv_main.SelectedNode.Tag.ToString().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1), Int32)
                'Lấy chỉ số xác định chi nhánh - đơn vị
                Dim parent_node As String = tv_main.SelectedNode.Parent.Tag.ToString()
                If (parent_node.IndexOf("_") > 0) Then
                    Dim arrId As String() = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
                    If (arrId.Length > 0) Then
                        Dim unitId As Int32 = CType(arrId(1).ToString(), Int32)
                        Dim db As DataTable = New DataTable
                        'Dim sSQL As String = ""
                        'sSQL = getSQL_DSCB(unitId, deptId)
                        Dim k As Integer = 0
                        Dim dbconn As DBAccess = New DBAccess
                        Dim MaSo_DVCurr As String = ""
                        Dim IdGoc_DVCurr As Integer = 0
                        IdGoc_DVCurr = dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & unitId)
                        If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = unitId
                        MaSo_DVCurr = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)
                        'k = sSQL.IndexOf("F")
                        'sSQL = "SELECT * " & sSQL.Substring(k - 1)
                        'db = _HS_CanBo.GetAllHuman(deptId, unitId)

                        db = _HS_CanBo.GetHS_CanBo_Search("", "", "", "", unitId, "", "", deptId, "", "", 0, "", "", "", "", "", "", "", "", "", "", CType("0", Byte), CType("0", Byte), "", "", CType("0", Integer), "", "", "", "", CType("0", Byte), CType("1", Byte))
                        'db = dbconn.SelectDBRows(sSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                Dim _rows As Int32 = db.Rows.Count - 1
                                For i As Integer = 0 To _rows
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = My_CStr(db.Rows(i)("IdCanBo").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = db.Rows(i)("STT").ToString() 'CType(i + 1, String)
                                    dgv_main.Rows(i).Cells("cln_MaCB").Value = My_CStr(db.Rows(i)("MaCB").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_HoTen").Value = My_CStr(db.Rows(i)("HoTen").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Username").Value = My_CStr(db.Rows(i)("Login_UserName"), "")
                                    dgv_main.Rows(i).Cells("cln_Quyen").Value = My_CStr(db.Rows(i)("Id_Nhom"), "")

                                    If MaSo_DVCurr = gMaDonViTW Then
                                        dgv_main.Rows(i).Cells("cln_TWQuanLy").Value = "X"
                                    Else
                                        dgv_main.Rows(i).Cells("cln_TWQuanLy").Value = IIf(My_CBool(db.Rows(i)("CapQuanLy"), False) = True, "X", "")
                                    End If
                                    dgv_main.Rows(i).Cells("cln_GioiTinh").Value = db.Rows(i)("GioiTinh_HT").ToString()
                                    dgv_main.Rows(i).Cells("cln_NgaySinh").Value = db.Rows(i)("NgaySinh_HT").ToString()
                                    'List tên đơn vị từ Id đơn vị ra lưới dữ liệu
                                    dgv_main.Rows(i).Cells("cln_DonVi").Value = db.Rows(i)("GhiChu_BS").ToString() '_HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh where id = {0}", CType(db.Rows(i)("IdDonVi").ToString(), Int32)))
                                    If My_CInt(MaSo_DVCurr.Substring(0, 4), -1) <= 1 Then
                                        If IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi")) Then
                                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
                                        Else
                                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
                                        End If
                                    Else
                                        If (IdGoc_DVCurr <> dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id_goc=" & IdGoc_DVCurr & " and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) And IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi"))) Then
                                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
                                        Else
                                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
                                        End If
                                    End If
                                    
                                    dgv_main.Rows(i).Cells("cln_SoCMT").Value = db.Rows(i)("CMT_So").ToString()
                                    dgv_main.Rows(i).Cells("cln_DiaChi").Value = db.Rows(i)("ThT_DiaChi_HT").ToString()
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
                                    If db.Rows(i)("Ngay_BienChe").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_BienChe").Value = CType(db.Rows(i)("Ngay_BienChe").ToString(), DateTime).ToString("dd-MM-yyyy")
                                    Else
                                        dgv_main.Rows(i).Cells("cln_BienChe").Value = ""
                                    End If
                                    dgv_main.Rows(i).Cells("cln_HonNhan").Value = db.Rows(i)("HonNhan_HT").ToString()
                                Next
                                _Records = db.Rows.Count
                            End If
                        End If
                        If (db IsNot Nothing) Then db.Clear()
                    End If
                End If
            End If

            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "CB") Then
                _BranchTag = tv_main.SelectedNode.Parent.Parent.Tag.ToString()
                Dim _IdHuman As String = tv_main.SelectedNode.Tag.ToString().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1)
                If (_IdHuman <> "") Then
                    dgv_main.Rows.Clear()
                    'Fill bản ghi thông tin nhân sự theo Mã cán bộ truyền vào
                    Dim db As DataTable = New DataTable
                    Dim strParent As String() = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
                    Dim sSQL As String = ""
                    sSQL = getSQL_DSCB(strParent(1), strParent(2))
                    Dim k As Integer = 0
                    Dim dbconn As DBAccess = New DBAccess
                    Dim MaSo_DVCurr As String = ""
                    Dim IdGoc_DVCurr As Integer = 0
                    IdGoc_DVCurr = dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & strParent(1))
                    If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = strParent(1)
                    MaSo_DVCurr = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)
                    k = sSQL.IndexOf("F")
                    sSQL = "SELECT * " & sSQL.Substring(k - 1)
                    k = sSQL.ToUpper.IndexOf("ORDER BY", 1)
                    sSQL = sSQL.Substring(0, k) & " And t2.IdCanBo='" & _IdHuman & "' "

                    If checkRight_CreateRecord(_IdHuman) Then
                        btn_edit.Enabled = True
                        btn_delete.Enabled = True
                    Else
                        btn_edit.Enabled = False
                        btn_delete.Enabled = False
                    End If

                    'Dim dr As DataRow
                    'dr = _HS_CanBo.GetHuman_ForCode(_IdHuman)
                    db = _HS_CanBo.GetHS_CanBo_Search("", _IdHuman, "", "", CType(strParent(1), Integer), "", "", CType(strParent(2), Integer), "", "", 0, "", "", "", "", "", "", "", "", "", "", CType("0", Byte), CType("0", Byte), "", "", CType("0", Integer), "", "", "", "", CType("0", Byte), CType("1", Byte))
                    'db = dbconn.SelectDBRows(sSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                dgv_main.Rows.Add()
                                dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdCanBo").ToString() <> "", db.Rows(i)("IdCanBo").ToString(), "")
                                dgv_main.Rows(i).Cells("cln_STT").Value = db.Rows(i)("STT").ToString() 'CType(i + 1, String)
                                dgv_main.Rows(i).Cells("cln_MaCB").Value = IIf(db.Rows(i)("MaCB").ToString() <> "", db.Rows(i)("MaCB").ToString(), "")
                                dgv_main.Rows(i).Cells("cln_HoTen").Value = IIf(db.Rows(i)("HoTen").ToString() <> "", db.Rows(i)("HoTen").ToString(), "")

                                dgv_main.Rows(i).Cells("cln_Username").Value = My_CStr(db.Rows(i)("Login_UserName"), "")
                                dgv_main.Rows(i).Cells("cln_Quyen").Value = My_CStr(db.Rows(i)("Id_Nhom"), "")

                                If MaSo_DVCurr = gMaDonViTW Then
                                    dgv_main.Rows(i).Cells("cln_TWQuanLy").Value = "X"
                                Else
                                    dgv_main.Rows(i).Cells("cln_TWQuanLy").Value = IIf(My_CBool(db.Rows(i)("CapQuanLy"), False) = True, "X", "")
                                End If

                                dgv_main.Rows(i).Cells("cln_GioiTinh").Value = db.Rows(i)("GioiTinh_HT").ToString()
                                dgv_main.Rows(i).Cells("cln_NgaySinh").Value = db.Rows(i)("NgaySinh_HT").ToString()
                                'List tên đơn vị từ Id đơn vị ra lưới dữ liệu
                                dgv_main.Rows(i).Cells("cln_DonVi").Value = db.Rows(i)("GhiChu_BS").ToString() '_HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh where id = {0}", CType(db.Rows(i)("IdDonVi").ToString(), Int32)))
                                If My_CInt(MaSo_DVCurr.Substring(0, 4), -1) <= 1 Then
                                    If IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi")) Then
                                        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
                                    Else
                                        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
                                    End If
                                Else
                                    If (IdGoc_DVCurr <> dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id_goc=" & IdGoc_DVCurr & " and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) And IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi"))) Then
                                        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
                                    Else
                                        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
                                    End If
                                End If
                                dgv_main.Rows(i).Cells("cln_SoCMT").Value = db.Rows(i)("CMT_So").ToString()
                                dgv_main.Rows(i).Cells("cln_DiaChi").Value = db.Rows(i)("ThT_DiaChi_HT").ToString() '_HS_CanBo.GetAddress(CType(IIf(db.Rows(i)("IdThT_Tinh").ToString() <> "", db.Rows(i)("IdThT_Tinh").ToString(), "0"), Int32), CType(IIf(db.Rows(i)("IdThT_Huyen").ToString() <> "", db.Rows(i)("IdThT_Huyen").ToString(), "0"), Int32), db.Rows(i)("ThT_Diachi").ToString().Trim())

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
                                If db.Rows(i)("Ngay_BienChe").ToString().Trim() <> "" Then

                                    dgv_main.Rows(i).Cells("cln_BienChe").Value = CType(db.Rows(i)("Ngay_BienChe").ToString(), DateTime).ToString("dd-MM-yyyy")
                                Else
                                    dgv_main.Rows(i).Cells("cln_BienChe").Value = ""
                                End If
                                dgv_main.Rows(i).Cells("cln_HonNhan").Value = db.Rows(i)("HonNhan_HT").ToString()
                            Next
                            _Records = db.Rows.Count
                        End If
                    End If
                    If (db IsNot Nothing) Then db.Clear()
                End If
            End If
            tv_main.SelectedNode.Expand()
        End If
        'Thực hiện fill danh sách cán bộ tập sự để chuyển thành lao động dài hạn
        '--> Lấy id đơn vị vừa chọn

        Dim strSQL As String = ""
        Dim BranchId As String = ""
        If (_BranchTag <> "") Then
            Dim arrElement() As String = _BranchTag.Split("_")
            If (arrElement.Length > 0) Then
                If _HS_CanBo.GetTrucThuoc(arrElement(2), CType(arrElement(1).ToString().Trim(), Integer)) = 4 Then 'Globals.GetAppSetting("HUYEN").Trim() Then
                    'Thực hiện lấy id chi nhánh tỉnh từ id pgd
                    BranchId = _ListDocument.GetRootId(CType(arrElement(1).ToString().Trim(), Integer))
                Else
                    BranchId = arrElement(1).ToString().Trim()
                End If
            End If
        End If
        '--> Bắt đầu fill dữ liệu
        cb_cbtapsu.Items.Clear()
        arrTapsu.Clear()
        If (BranchId <> "") Then
            strSQL = String.Format("Select Id, HoTen from HSCB_TS Where TrangThai = 0 and IdChiNhanh = {0} Order by HoTen", CType(BranchId, Integer))
            arrTapsu = _Globals.Bind_ComBoBox(cb_cbtapsu, strSQL, "---Danh sách cán bộ tập sự---")
            If (cb_cbtapsu.Items.Count <> 0) Then
                cb_cbtapsu.SelectedIndex = 0
            End If
            'Lấy Chỉ số xác định đơn vị Hiện tại của cán bộ (Xét với trường hợp cài đặt đơn vị '000100')
            If (DONVI = gMaDonViTW) Then
                _IdDonviHT = CType(BranchId, Integer)
            End If
        End If
        TotalCheckBoxes = dgv_main.RowCount
        TotalCheckedCheckBoxes = 0
        If dgv_main.Rows.Count > 0 Then dgv_main_CellClick(sender, Nothing)

    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Close()
    End Sub

    Private Sub frmHS_CanBo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub ckb_chon_dltapsu_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_chon_dltapsu.CheckedChanged
        If (ckb_chon_dltapsu.Checked = True) Then
            cb_cbtapsu.Enabled = True
            lbl_div_2.Visible = False
            btn_view_detail.Visible = False
            lbl_div_1.Visible = False
            btn_delete.Visible = False
            lbl_div_0.Visible = False
            btn_edit.Visible = False
            btn_add.Text = "&Chuyển"
        Else
            cb_cbtapsu.Enabled = False
            btn_view_detail.Visible = True
            btn_edit.Visible = True
            btn_delete.Visible = True
            lbl_div_0.Visible = True
            lbl_div_1.Visible = True
            lbl_div_2.Visible = True
            btn_add.Text = "&Thêm"
        End If
    End Sub

    Private Sub rb_hsnhansu_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_hsnhansu.CheckedChanged
        If (rb_hsnhansu.Checked = True) Then
            pnl_hs_tapsu.Visible = True
            btn_add.Visible = True
            lbl_div_main.Visible = True
            btn_edit.Width = 61
            btn_edit.Text = "&Sửa"
            If (_IdNew <> "") Then btn_edit.Enabled = False
        End If
    End Sub

    Private Sub rb_hskhac_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_hskhac.CheckedChanged
        If (rb_hskhac.Checked = True) Then
            btn_add.Visible = False
            lbl_div_main.Visible = False
            If (_IdNew <> "") Then
                btn_edit.Enabled = True
            End If
            btn_edit.Width = 105
            btn_edit.Text = "&Vào hồ sơ khác"
            pnl_hs_tapsu.Visible = False
        End If
    End Sub

    Private Sub dgv_main_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellClick
        _IdNew = ""
        pnl_hs_tapsu.Enabled = True
        If (dgv_main.Rows.Count > 0) Then
            pnl_hs_tapsu.Enabled = True
            btn_add.Enabled = True
            btn_edit.Enabled = True
            btn_delete.Enabled = True
            If DONVI = gMaDonViTW Or Globals.Group.ToUpper = "OPERATOR" Or Globals.Group.ToUpper = "U01" Then btn_CapQL.Enabled = True 'User TW hoặc user cấp Admin Chi nhánh đc chuyển Cấp quản lý

            'Nếu người đang sử dụng là thuộc hội sở chính thì kiểm tra xem hồ sơ đang chọn có thuộc TW quản lý ko
            If (IdDONVI = 1 AndAlso dgv_main.CurrentRow.Cells("cln_TWQuanLy").Value = "") Or (IdDONVI <> 1 AndAlso dgv_main.CurrentRow.Cells("cln_TWQuanLy").Value = "X") Then
                btn_add.Enabled = False
                btn_edit.Enabled = False
                btn_delete.Enabled = False
                btn_CapQL.Enabled = False
            End If
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                Dim dr As DataRow
                dr = _HS_CanBo.GetHuman_ForCode(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString())
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        'Nếu là cán bộ chính thức đã chuyển sang lao động ngắn hạn
                        If (dr("IdNew").ToString().Trim() <> "") Then
                            _IdNew = dr("IdNew").ToString().Trim()
                            pnl_hs_tapsu.Enabled = False
                            btn_add.Enabled = False
                            If (rb_hskhac.Checked = False) Then
                                btn_edit.Enabled = False
                                btn_CapQL.Enabled = False
                            End If
                        End If
                    End If
                End If
            End If

            '-------------
            'If DONVI = "000100" Then btn_delete.Enabled = True
        End If
    End Sub

    Private Sub dgv_main_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub
#End Region

#Region "---> Events: Hiển thị menu gọi các hồ sơ khác của cán bộ <---"
    Private Sub dgv_main_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgv_main.MouseUp
        If (e.Button = Windows.Forms.MouseButtons.Right) Then
            Dim _htinfor As DataGridView.HitTestInfo = dgv_main.HitTest(e.X, e.Y)
            If (_htinfor.RowIndex >= 0) Then
                If (_htinfor.Type = DataGridViewHitTestType.Cell And _htinfor.RowIndex = _htinfor.RowIndex) Then
                    ctmn_main.Show(Me.dgv_main, New Point(e.X, e.Y))
                End If
                dgv_main.Rows(_htinfor.RowIndex).Selected = True
                dgv_main.CurrentCell = dgv_main(_htinfor.ColumnIndex, _htinfor.RowIndex)
            End If
        End If
    End Sub

    Private Sub mnu_QDNhansu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnu_QDNhansu.Click
        Dim frm As New frmQuyetDinhNS
        Dim arr() As String = _BranchTag.Split("_")
        frm.IDCB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        frm.IDPB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Phong").Value.ToString()
        frm.IDDV_HS_Canbo = arr(1)
        frm.ShowDialog()
    End Sub

    Private Sub mnu_Dtao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnu_Dtao.Click
        Dim frm As New frmDaoTao
        Dim arr() As String = _BranchTag.Split("_")
        frm.IDCB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        frm.IDPB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Phong").Value.ToString()
        frm.IDDV_HS_Canbo = arr(1)
        frm.ShowDialog()
    End Sub

    Private Sub mnu_DanhGia_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnu_DanhGia.Click
        Dim frm As New frmDanhGiaCanBo
        Dim arr() As String = _BranchTag.Split("_")
        frm.IDCB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        frm.IDPB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Phong").Value.ToString()
        frm.IDDV_HS_Canbo = arr(1)
        frm.ShowDialog()
    End Sub

    Private Sub mnu_KTKL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnu_KTKL.Click
        Dim frm As New frmKhenThuongKiLuat
        Dim arr() As String = _BranchTag.Split("_")
        frm.IDCB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        frm.IDPB_HS_Canbo = dgv_main.CurrentRow.Cells("cln_Phong").Value.ToString()
        frm.IDDV_HS_Canbo = arr(1)
        frm.ShowDialog()
    End Sub

    Private Sub mnu_ctdang_dt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_ctdang_dt.Click
        Dim obj_hsdoanthe As New frmHS_DoanThe
        obj_hsdoanthe.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hsdoanthe.TagNode = _BranchTag
        obj_hsdoanthe.FlagShow = True
        obj_hsdoanthe.ShowDialog()
    End Sub

    Private Sub mnu_hdld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hdld.Click
        Dim obj_hdld As New frmHS_HDLD
        obj_hdld.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hdld.TagNode = _BranchTag
        obj_hdld.FlagShow = True
        obj_hdld.ShowDialog()
    End Sub

    Private Sub mnu_csld_bhyt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_csld_bhyt.Click
        Dim obj_bhyt As New frmHS_CsLaoDong
        obj_bhyt.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_bhyt.FlagShow = True
        obj_bhyt.TagNode = _BranchTag
        obj_bhyt.TabVal = 0
        obj_bhyt.ShowDialog()
    End Sub

    Private Sub mnu_csld_qlnc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_csld_qlnc.Click
        Dim obj_qlnc As New frmHS_CsLaoDong
        obj_qlnc.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_qlnc.FlagShow = True
        obj_qlnc.TagNode = _BranchTag
        obj_qlnc.TabVal = 1
        obj_qlnc.ShowDialog()
    End Sub

    Private Sub mnu_csld_hstainan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_csld_hstainan.Click
        Dim obj_hstainan As New frmHS_CsLaoDong
        obj_hstainan.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hstainan.FlagShow = True
        obj_hstainan.TagNode = _BranchTag
        obj_hstainan.TabVal = 2
        obj_hstainan.ShowDialog()
    End Sub

    Private Sub mnu_csld_khamsk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_csld_khamsk.Click
        Dim obj_khamsk As New frmHS_CsLaoDong
        obj_khamsk.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_khamsk.FlagShow = True
        obj_khamsk.TagNode = _BranchTag
        obj_khamsk.TabVal = 3
        obj_khamsk.ShowDialog()
    End Sub

    Private Sub mnu_hscongtac_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hscongtac.Click
        Dim obj_hscongtac As New frmCN_HSCanBo
        obj_hscongtac.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hscongtac.FlagHS = True
        obj_hscongtac.TabVal = 5
        obj_hscongtac.ShowDialog()
    End Sub

    Private Sub mnu_hsxuatngoai_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hsxuatngoai.Click
        Dim obj_hsxuatngoai As New frmCN_HSCanBo
        obj_hsxuatngoai.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hsxuatngoai.FlagHS = True
        obj_hsxuatngoai.TabVal = 1
        obj_hsxuatngoai.ShowDialog()
    End Sub

    Private Sub mnu_hshochieu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hshochieu.Click
        Dim obj_hshochieu As New frmCN_HSCanBo
        obj_hshochieu.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hshochieu.FlagHS = True
        obj_hshochieu.TabVal = 2
        obj_hshochieu.ShowDialog()
    End Sub

    Private Sub mnu_hsllvt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hsllvt.Click
        Dim obj_hsllvt As New frmCN_HSCanBo
        obj_hsllvt.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_hsllvt.FlagHS = True
        obj_hsllvt.TabVal = 3
        obj_hsllvt.ShowDialog()
    End Sub

    Private Sub mnu_hscu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hscu.Click
        Dim obj_olddocument As New frmCN_HSCanBo
        obj_olddocument.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_olddocument.FlagHS = True
        obj_olddocument.TabVal = 6
        obj_olddocument.ShowDialog()
    End Sub

    Private Sub mnu_hhhv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hhhv.Click
        Dim obj_education As New frmCN_HSCanBo
        obj_education.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_education.FlagHS = True
        obj_education.TabVal = 4
        obj_education.ShowDialog()
    End Sub

    'Menu gọi hồ sơ Gia đình của cán bộ bạn click chuột phải
    Private Sub mnu_hsgdcanbo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_hsgdcanbo.Click
        Dim obj_gdcb As New frmHS_GDCB
        obj_gdcb.HumanId = dgv_main.CurrentRow.Cells("cln_Id").Value.ToString()
        obj_gdcb.TagNode = _BranchTag
        obj_gdcb.FlagShow = True
        obj_gdcb.ShowDialog()
    End Sub

#End Region

#Region "---> Events: Hiển thị menu Xuất CV của cán bộ ra Word <---"
    Private Sub mnu_exportcv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_exportcv.Click
        If (dgv_main.Rows.Count > 0) Then
            Me.Cursor = Cursors.WaitCursor
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                _HS_CanBo.ExportWord_CurriculumVitae(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString())
            End If
            Me.Cursor = Cursors.Default
        End If
    End Sub

    '''' <summary>
    '''' Hàm thực hiện xuất dữ liệu Hồ sơ cán bộ thành mẫu 02C (Lý lịch cá nhân) ra file Word
    '''' </summary>
    '''' <param name="_RowId">Chuỗi chỉ số xác định bản ghi của cán bộ</param>
    '''' <remarks></remarks>
    'Private Sub ExportWord_CurriculumVitae(ByVal _RowId As String)
    '    Dim _sDiaban As String = _HS_CanBo.GetVarNam("DIABAN").ToString().Trim()
    '    Dim _sDonviTT As String = ""        'Biến lưu chuỗi Đơn vị Trực thuộc
    '    Dim _sDonviCS As String = ""        'Biến lưu chuỗi Đơn vị Cơ sở
    '    Dim _IdDvParent As Integer = 0          'Biến lưu Chỉ số đơn vị cấp nhỏ nhất của Cán bộ
    '    Dim _IdDvChild As Integer = 0           'Biến lưu Chỉ số đơn vị cấp nhỏ nhất của Cán bộ
    '    Dim _IdPb As Integer = 0                'Biến lưu chỉ số Phòng ban của cán bộ
    '    Dim _IdCv As Integer = 0                'Biến lưu chỉ số xác định chức vụ Cán bộ
    '    Dim _IdChmon As Integer = 0             'Biến lưu chỉ số xác Chuyên môn mới nhất của cán bộ (Công việc chính đang làm)
    '    Dim _sVal As String = ""
    '    Dim strSQL As String = ""

    '    Dim dr As DataRow
    '    dr = _HS_CanBo.GetHuman_ForCode(_RowId)
    '    If Not (dr Is Nothing) Then
    '        If (dr.Table.Rows.Count > 0) Then
    '            'Lấy các thông tin chính để xuất dữ liệu phần đầu
    '            _IdDvParent = dr("IdDonVi").ToString().Trim()
    '            'Lấy dữ liệu Đơn vị cấp dưới của cán bộ đang công tác, Phòng ban, chức vụ
    '            Dim HS_QDNhansu As QDNhanSu = New QDNhanSu
    '            HS_QDNhansu = HS_QDNhansu.getFinalRecord(_RowId)
    '            If Not (HS_QDNhansu Is Nothing) Then
    '                If (HS_QDNhansu.IdQDNhanSu <> "") Then
    '                    _IdCv = IIf(HS_QDNhansu.IdChucvu_Moi <> 0, HS_QDNhansu.IdChucvu_Moi, 0)
    '                    _IdPb = IIf(HS_QDNhansu.IdPhong_Moi <> 0, HS_QDNhansu.IdPhong_Moi, 0)
    '                    _IdChmon = IIf(HS_QDNhansu.IdChuyenMon_Moi <> 0, HS_QDNhansu.IdChuyenMon_Moi, 0)
    '                    _IdDvChild = IIf(HS_QDNhansu.IdDonvi_Moi <> 0, HS_QDNhansu.IdDonvi_Moi, 0)
    '                End If
    '            End If
    '            strSQL = String.Format("Select id, ten_goi from ChiNhanh Where Status = 1 and id = {0} And id_goc IN (0,1)", _IdDvParent)
    '            _sDonviTT = "Đơn vị trực thuộc: " + Replace_Branch(_HS_CanBo.GetNameByCode(strSQL).ToString())
    '            If (_IdDvParent <> _IdDvChild) Then 'Phòng giao dịch
    '                strSQL = String.Format("Select id, ten_goi from ChiNhanh Where Status = 1 and id = {0} And id_goc > 1", _IdDvChild)
    '                _sDonviCS = "Đơn vị cơ sở: " + Replace_Branch(_HS_CanBo.GetNameByCode(strSQL).ToString())
    '            ElseIf (_IdDvParent = _IdDvChild) Then 'Phòng ban của tỉnh hoặc Hội sở
    '                strSQL = String.Format("Select id, ten_phong From PhongBan Where Status = 1 And id = {0}", _IdPb)
    '                _sDonviCS = "Đơn vị cơ sở: " + Replace_Branch(_HS_CanBo.GetNameByCode(strSQL).ToString())
    '            End If

    '            'Bắt đầu khai báo và xuất dữ liệu ra Word
    '            Dim objWordApp As New Word.Application      'Tạo Đối tượng Word Application
    '            Dim objDocument As New Word.Document        'Tạo đối tượng Word Document
    '            Dim objTable As Word.Table                  'Tạo đối tượng Word Table
    '            'Start Word and open the document template.
    '            objWordApp.Visible = True                          'And show word screen (False: Hide msword cho đến khi xuất hết thông tin)
    '            objWordApp.Activate()
    '            objDocument = objWordApp.Documents.Add                  'Add một Document vào trong Application (Create a new word document)
    '            'Biến lưu vị trí select hiện hành
    '            Dim objselection As Word.Selection
    '            'Gán vị trí hiện hành trong Document vào biến selection
    '            objselection = objDocument.Application.Selection()

    '            'Định dạng Paragraph
    '            objselection.Font.Color = Word.WdColor.wdColorAutomatic
    '            objselection.Font.Size = 13
    '            objselection.Font.Name = "Times New Roman"
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

    '            'Tạo một bảng gồm 3 dòng và 3 cột
    '            'objTable = objselection.Tables.Add(objDocument.Bookmarks.Item("\endofdoc").Range, 3, 3)
    '            objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 3, 3)
    '            objTable.Columns(1).Width = 180.0F
    '            objTable.Columns(2).Width = 162.0F
    '            objTable.Columns(3).Width = 100.0F
    '            objTable.Cell(1, 2).Merge(objTable.Cell(3, 2))
    '            objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '            objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '            objTable.LeftPadding = -10
    '            objTable.RightPadding = -10
    '            objTable.Cell(1, 2).Range.InsertAfter("LÝ LỊCH CÁ NHÂN")
    '            objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '            objTable.Cell(1, 2).Range.Font.Name = "Times New Roman"
    '            objTable.Cell(1, 2).Range.Font.Size = 20
    '            objTable.Cell(1, 2).Range.Font.Bold = 1

    '            objTable.Cell(1, 1).Range.InsertAfter("Tỉnh (TP): " + _sDiaban)
    '            objTable.Cell(2, 1).Range.InsertAfter(_sDonviTT)
    '            objTable.Cell(3, 1).Range.InsertAfter(_sDonviCS)
    '            objTable.Cell(1, 3).Range.InsertAfter("Mẫu 01/LLCB")
    '            objTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '            objTable.Cell(3, 3).Range.InsertAfter("...................................")
    '            objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.InsertParagraph()

    '            'Xóa định dạng Paragraph trước
    '            objselection.ClearFormatting()
    '            'Định dạng lại Paragraph
    '            objselection.Font.Color = Word.WdColor.wdColorAutomatic
    '            objselection.Font.Size = 12
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '            'Add Text vào Paragraph
    '            objselection.TypeText("Số hiệu cán bộ, công chức")
    '            objselection.Paragraphs.SpaceAfter = 14

    '            'BẮT ĐẦU XUẤT DỮ LIỆU CHÍNH CỦA HỒ SƠ CÁN BỘ
    '            'Tạo một Paragraph mới
    '            objselection.TypeParagraph()
    '            objselection.Font.Size = 12
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            objselection.Paragraphs.SpaceAfter = 4
    '            objselection.TypeText("1. Họ và tên ")
    '            objselection.Font.Size = 10
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("(Viết chữ in hoa): ")
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            _sVal = dr("HoTen").ToString().Trim().ToUpper() & vbTab & vbTab & vbTab & vbTab & "2. Nam, Nữ: " & IIf(dr("GioiTinh").ToString() = False, " Nam", " Nữ")
    '            objselection.TypeText(_sVal + vbCrLf)
    '            objselection.TypeText("3. Các tên gọi khác: " + dr("TenThuongGoi").ToString().Trim() + vbCrLf)

    '            '4.5.6. Write 4: Chức vụ (Đảng, đoàn thể, chính quyền, kể cả chức vụ kiêm nhiệm)
    '            '--->Lấy chức vụ đảng trong thời gian hiện tại (Tức chức vụ đó thì ngày hiện tại phải thuộc khoảng từ ngày 1 đến ngày 2)
    '            Dim arrPartys() As String       'Lấy mảng các thông tin về đảng của cán bộ
    '            arrPartys = GetPartys(_RowId)
    '            'Lấy Chức vụ đảng của cán bộ (Chức vụ chính và Chức vụ Kiêm)
    '            If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
    '                _sVal = arrPartys(0) + "; " + arrPartys(4)
    '            Else
    '                _sVal = IIf(arrPartys(0) <> "", arrPartys(0), IIf(arrPartys(4) <> "", arrPartys(4), ""))
    '            End If
    '            Dim _sDoanvien As String = GetValue(_RowId, 1) 'Lấy chuỗi thông tin đoàn
    '            If (_sDoanvien <> "" And _sDoanvien.IndexOf(";") >= 0) Then
    '                If (_sVal <> "") Then
    '                    _sVal += "; " + _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
    '                Else
    '                    _sVal = _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
    '                End If
    '            End If
    '            If (_sVal <> "") Then
    '                _sVal += "; " + _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
    '            Else
    '                _sVal = _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
    '            End If
    '            _sVal = _sVal.Replace(";;", ";").Replace("; ;", ";").Trim()
    '            While _sVal.EndsWith("-") 'Bỏ dấu gạch ngang (-) ở cuối chuỗi nếu có
    '                _sVal = _sVal.Substring(0, _sVal.Length - 1).Trim()
    '            End While
    '            If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
    '                _sVal = _sVal.Substring(1).Trim()
    '            End If
    '            While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
    '                _sVal = _sVal.Substring(0, _sVal.Length - 1)
    '            End While
    '            If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
    '                _sVal += vbTab & vbTab & vbTab
    '            Else
    '                _sVal += vbTab
    '            End If
    '            objselection.TypeText("4. Chức vụ ")
    '            'Xóa định dạng Paragraph cũ
    '            objselection.Font.Size = 10
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("(Đảng, đoàn thể, chính quyền, kể cả chức vụ kiêm nghiệm): ")
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            objselection.TypeText(_sVal)

    '            'Lấy phụ cấp chức vụ và Phụ cấp trách nhiệm
    '            _sVal = GetValue(_RowId, 2)
    '            _sVal = "5. Phụ cấp chức vụ: " + _sVal.Substring(0, _sVal.IndexOf(";")) + vbTab + vbTab + "6. Phụ cấp trách nhiệm: " + _sVal.Substring(_sVal.IndexOf(";") + 1) + vbCrLf
    '            objselection.TypeText(_sVal)
    '            '7. Write 7: Cấp uỷ hiện tai, cấp uỷ kiêm
    '            _sVal = "7. Cấp uỷ hiện tại: " + arrPartys(1) & vbTab & vbTab & vbTab + ", Cấp uỷ kiêm: " + arrPartys(5) + vbCrLf
    '            objselection.TypeText(_sVal)

    '            '8. Write 8: Ngày tháng năm sinh, nơi sinh của cán bộ
    '            _sVal = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
    '            objselection.TypeText("8. Ngày sinh " & _sVal.Substring(0, 2) & " tháng " & _sVal.Substring(3, 2) & " năm " & _sVal.Substring(6) & vbTab & vbTab & "9. Nơi sinh: ")
    '            _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdNS_Tinh").ToString() <> "", dr("IdNS_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Huyen").ToString() <> "", dr("IdNS_Huyen").ToString(), "0"), Int32), dr("NS_DChi").ToString().Trim())
    '            objselection.TypeText(_sVal & vbCrLf)
    '            _sVal = IIf(dr("IdQuocTich").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from QuocGia Where id = {0} and Status = 1", CType(dr("IdQuocTich").ToString(), Int32))), "")
    '            objselection.TypeText("10. Quốc tịch: " & _sVal & vbTab & vbTab)

    '            _sVal = IIf(dr("IdDanToc").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 21 and Status = 1", CType(dr("IdDanToc").ToString(), Int32))), "")
    '            objselection.TypeText("11. Dân tộc: " & _sVal & vbTab & vbTab)
    '            _sVal = IIf(dr("IdTonGiao").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 24 and Status = 1", CType(dr("IdTonGiao").ToString(), Int32))), "")
    '            objselection.TypeText("12. Tôn giáo: " & _sVal & vbTab & vbTab & vbCrLf)
    '            _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdNQ_Tinh").ToString() <> "", dr("IdNQ_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNQ_Huyen").ToString() <> "", dr("IdNQ_Huyen").ToString(), "0"), Int32), dr("NQ_DChi").ToString().Trim())
    '            objselection.TypeText("13. Quê quán: " & _sVal & vbCrLf)
    '            _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdThT_Tinh").ToString() <> "", dr("IdThT_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdThT_Huyen").ToString() <> "", dr("IdThT_Huyen").ToString(), "0"), Int32), dr("ThT_Diachi").ToString().Trim())
    '            objselection.TypeText("14. Nơi đăng ký Hộ khẩu thường trú: " & _sVal & vbCrLf)

    '            _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdTTr_Tinh").ToString() <> "", dr("IdTTr_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdTTr_Huyen").ToString() <> "", dr("IdTTr_Huyen").ToString(), "0"), Int32), dr("TTr_Diachi").ToString().Trim())
    '            objselection.TypeText("15. Nơi ở hiện nay: " & _sVal & vbCrLf)
    '            _sVal = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
    '            objselection.TypeText("16. Giấy CMND số: " & dr("CMT_So").ToString().Trim() & vbTab & "Ngày cấp: " & _sVal & " " & vbTab & " Nơi cấp: " & dr("CMT_NoiCap").ToString().Trim() & vbCrLf)
    '            objselection.TypeText("17. Điện thoại: ")
    '            'Xóa định dạng Paragraph cũ
    '            objselection.Font.Italic = 1
    '            objselection.TypeText(" Di động: " & dr("DienThoai_DD").ToString().Trim() & vbTab & " Cơ quan: " & dr("DienThoai_CQ").ToString().Trim() & vbTab & " Nhà riêng: " & dr("DienThoai_NR").ToString().Trim())
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            objselection.Paragraphs.SpaceAfter = 4
    '            objselection.TypeText(vbCrLf)
    '            objselection.TypeText("18. Địa chỉ e-mail: " & dr("Emai").ToString().Trim() & vbCrLf)
    '            '19. Lấy dữ liệu tình trạng Sức khoẻ của cán bộ trong hồ sơ Sức khoẻ của Cán bộ
    '            Dim arrTemp() As String = GetValue(_RowId, 3).Split(";")
    '            objselection.TypeText("19. Tình trạng sức khoẻ: " & arrTemp(2).Trim() & "," & vbTab & " Cao: " & arrTemp(0).Trim() & "," & vbTab & " Cân Nặng: " & arrTemp(1).Trim() & " (kg)," & vbTab & " Nhóm máu: " & dr("NhomMau").ToString().Trim() & vbCrLf)
    '            '20. Trình độ học vấn: Giáo dục phổ thông (Lớp mấy):
    '            _sVal = IIf(dr("IdTrinhDoVH").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 6 and Status = 1", CType(dr("IdTrinhDoVH").ToString(), Int32))), "")
    '            objselection.TypeText("20. Trình độ học vấn: Giáo dục phổ thông (Lớp mấy): " & _sVal & vbCrLf)
    '            'Thực hiện lấy thông tin Học vị - Chuyên ngành đào tạo - Trường đào tạo - Năm tốt nghiệp
    '            objselection.TypeText("21. Trình độ chuyên môn")
    '            objselection.Font.Size = 10
    '            objselection.TypeText(" (Trung cấp, Cử nhân, Kỹ sư) ")
    '            objselection.Font.Size = 12
    '            _sVal = GetValue(_RowId, 4)  '+ ", " + vbTab + GetValue(_RowId, 5)
    '            objselection.Font.Italic = 1
    '            objselection.TypeText(_sVal & vbCrLf)
    '            objselection.Font.Italic = 0
    '            'Lấy thông tin Học hàm
    '            '_sVal = GetValue(_RowId, 6) & vbCrLf
    '            'objselection.TypeText("22. Học hàm cao nhất: " + _sVal)

    '            objselection.TypeText("22. Học hàm cao nhất: ")
    '            objselection.Font.Italic = 1
    '            _sVal = GetValue(_RowId, 11)
    '            objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
    '            objselection.Font.Italic = 0


    '            'Lấy thông tin Lý luận chính trị - Ngoại ngữ - Tin học
    '            _sVal = IIf(dr("IdTrinhDoCT").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 36 and Status = 1", CType(dr("IdTrinhDoCT").ToString(), Int32))), "")
    '            objselection.TypeText("23. Lý luận chính trị: " + _sVal & vbTab & vbTab + "24. Ngoại ngữ: " + GetValue(_RowId, 7) & vbTab & vbTab + "25. Tin học: " + GetValue(_RowId, 8) + vbCrLf)
    '            objselection.Font.Size = 10
    '            objselection.TypeText("(Cao cấp, Trung cấp, Sơ cấp)" + vbTab + vbTab + "(T.sĩ, ĐH, C.chỉ Anh, Nga, Pháp A/B/C,...)" + vbTab + "   (T.sĩ, ĐH, C.chỉ Anh, Nga, Pháp A/B/C,...)" + vbCrLf)
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            _sVal = IIf(_IdChmon > 0, _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} and id_goc = 12 and Status = 1", _IdChmon)), "")
    '            objselection.TypeText("26. Công việc chính đang làm: " + _sVal + vbCrLf)
    '            strSQL = "SELECT (Select Mota from NgachLuong Where IdNgachLuong = b.IdNgachLuong And Status = 1) As NgachLuong,"
    '            strSQL += "a.IdLuongCB,b.BacLuong,a.HeSoLuong,Convert(Varchar, Ngay_Huong, 103) as NgayApDung  From HS_LuongCB a, BacLuong b"
    '            strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdBacLuong = b.IdBacLuong And b.Status = 1) Order by a.Ngay_Huong Desc"
    '            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If Not (db Is Nothing) Then
    '                    If db.Rows.Count > 0 Then
    '                        objselection.TypeText("27. Ngạch công chức, viên chức: " + db.Rows(0)("NgachLuong").ToString().Trim() + vbTab + vbTab)
    '                        objselection.TypeText("Bậc lương: " + db.Rows(0)("BacLuong").ToString().Trim() + vbTab + vbTab)
    '                        objselection.TypeText("Hệ số: " + CType(db.Rows(0)("HeSoLuong"), Double).ToString("N2") + vbTab)
    '                        If (db.Rows(0)("NgayApDung").ToString().Trim() <> "") Then
    '                            Dim arrDate() As String = db.Rows(0)("NgayApDung").ToString().Trim().Split("/")
    '                            objselection.TypeText("từ " + arrDate(0) + " tháng " + arrDate(1) + " năm " + arrDate(2) + vbCrLf)
    '                        Else
    '                            objselection.TypeText("từ    tháng    năm    " + vbCrLf)
    '                        End If
    '                    Else
    '                        objselection.TypeText("27. Ngạch công chức, viên chức:              Bậc lương:      Hệ số:      từ    tháng    năm" + vbCrLf)
    '                    End If
    '                Else
    '                    objselection.TypeText("27. Ngạch công chức, viên chức:              Bậc lương:      Hệ số:      từ    tháng    năm" + vbCrLf)
    '                End If
    '            End Using
    '            '28. Thành phần gia đình xuất thân (công nhân, nông dân, cán bộ, công chức, trí thức, quân nhân, dân nghèo thành thị, tiểu thương, tiểu chư, tư sản,...):
    '            objselection.TypeText("28. Thành phần gia đình xuất thân ")
    '            objselection.Font.Size = 10
    '            objselection.TypeText("(công nhân, nông dân, cán bộ, công chức, trí thức, quân nhân, dân nghèo thành thị, tiểu thương, tiểu chư, tư sản,...)")
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            _sVal = IIf(dr("IdThanhPhanGD").ToString() <> "", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 13 and Status = 1", CType(dr("IdThanhPhanGD").ToString(), Integer))), "")
    '            objselection.TypeText(": " + _sVal & vbCrLf)

    '            '29. Ưu tiên gia đình
    '            _sVal = IIf(dr("IdUT_GDinh").ToString() <> "", _HS_CanBo.GetNameByCode(String.Format("Select id, ma_so from DanhMuc Where id = {0} And id_goc = 19 and Status = 1", CType(dr("IdUT_GDinh").ToString(), Integer))), "")
    '            objselection.TypeText("29. Ưu tiên gia đình: " & vbTab & vbTab)
    '            If (String.Compare(_sVal, "1901", True) = 0) Then
    '                objselection.InsertSymbol(254, "Wingdings")
    '            Else
    '                objselection.InsertSymbol(112, "Wingdings")
    '            End If
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("- Gia đình cách mạng" & vbTab & vbTab)
    '            objselection.Font.Italic = 0

    '            If (String.Compare(_sVal, "1902", True) = 0) Then
    '                objselection.InsertSymbol(254, "Wingdings")
    '            Else
    '                objselection.InsertSymbol(112, "Wingdings")
    '            End If
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("- Gia đình liệt sỹ" & vbTab & vbTab)
    '            objselection.Font.Italic = 0
    '            If (String.Compare(_sVal, "1903", True) = 0) Then
    '                objselection.InsertSymbol(254, "Wingdings")
    '            Else
    '                objselection.InsertSymbol(112, "Wingdings")
    '            End If
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("- Không" & vbCrLf)
    '            objselection.Font.Italic = 0

    '            '30. Ưu tiên Bản thân cán bộ
    '            objselection.TypeText("30. Ưu tiên bản thân: " & vbTab & vbTab)
    '            If (dr("IdUT_BThan").ToString().Trim() <> "" And dr("IdUT_BThan").ToString().Trim() <> "0") Then
    '                _sVal = dr("IdUT_BThan").ToString().Trim()
    '                If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
    '                    _sVal = _sVal.Substring(1).Trim()
    '                End If
    '                While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
    '                    _sVal = _sVal.Substring(0, _sVal.Length - 1)
    '                End While
    '                _sVal = _sVal.Trim()
    '                Dim arrUTBT() As String = _sVal.Split(";")
    '                Dim lengths As Integer = arrUTBT.Length - 1
    '                For i As Integer = 0 To lengths
    '                    _sVal = IIf(arrUTBT(i).ToString() <> "", _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DanhMuc Where id = {0} And id_goc = 18 and Status = 1", CType(arrUTBT(i).ToString(), Integer))), "")
    '                    If (i > 1) Then
    '                        objselection.TypeText(vbCrLf + vbTab + vbTab + vbTab + vbTab)
    '                    End If
    '                    objselection.InsertSymbol(254, "Wingdings")
    '                    objselection.Font.Italic = 1
    '                    objselection.TypeText("- " + _sVal + vbTab + vbTab)
    '                    objselection.Font.Italic = 0
    '                Next
    '            Else
    '                objselection.Font.Italic = 0
    '                objselection.InsertSymbol(112, "Wingdings")
    '                objselection.Font.Italic = 1
    '                objselection.TypeText("- Thương binh hạng ...." + vbTab + vbTab)

    '                objselection.Font.Italic = 0
    '                objselection.InsertSymbol(112, "Wingdings")
    '                objselection.Font.Italic = 1
    '                objselection.TypeText("- Bệnh binh binh hạng ...." + vbCrLf)
    '                objselection.TypeText(vbTab + vbTab + vbTab + vbTab)
    '                objselection.Font.Italic = 0
    '                objselection.InsertSymbol(112, "Wingdings")
    '                objselection.Font.Italic = 1
    '                objselection.TypeText("- Anh hùng lao động" + vbTab + vbTab)
    '                objselection.Font.Italic = 0
    '                objselection.InsertSymbol(112, "Wingdings")
    '                objselection.Font.Italic = 1
    '                objselection.TypeText("- Anh hùng lực lượng vũ trang" + vbCrLf)
    '                objselection.TypeText(vbTab + vbTab + vbTab + vbTab)
    '                objselection.Font.Italic = 0
    '                objselection.InsertSymbol(112, "Wingdings")
    '                objselection.Font.Italic = 1
    '                objselection.TypeText("- Con liệt sỹ" + vbTab + vbTab + vbTab + vbTab)
    '                objselection.Font.Italic = 0
    '                objselection.InsertSymbol(112, "Wingdings")
    '                objselection.Font.Italic = 1
    '                objselection.TypeText("- Không")
    '            End If
    '            objselection.Font.Italic = 0
    '            objselection.TypeText(vbCrLf)
    '            '31. Nghề nghiệp bản thân trước khi được tuyển dụng (ghi nghề được đào tạo hoặc công nhân (thợ gì), làm ruộng, buôn bán, học sinh,...)
    '            objselection.TypeText("31. Nghề nghiệp bản thân trước khi được tuyển dụng ")
    '            objselection.Font.Size = 10
    '            objselection.TypeText("(ghi nghề được đào tạo hoặc công nhân (thợ gì), làm ruộng, buôn bán, học sinh,...)")
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            _sVal = GetValue(_RowId, 9)
    '            'objselection.TypeText("32. Ngày được tuyển dụng: " + vbTab + vbTab + " Vào cơ quan nào, ở đâu: " + vbCrLf)

    '            objselection.TypeText("32. Ngày được tuyển dụng: ")
    '            objselection.Font.Italic = 1
    '            If dr("Ngay_ThamNien").ToString().Trim <> "" Then
    '                _sVal = IIf(CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy"), vbTab + vbTab)
    '            Else
    '                _sVal = vbTab + vbTab
    '            End If
    '            objselection.TypeText(_sVal + vbTab)
    '            objselection.Font.Italic = 0

    '            objselection.TypeText("Vào cơ quan nào, ở đâu: ")
    '            _sVal = GetValue(_RowId, 9)
    '            objselection.Font.Italic = 1
    '            objselection.TypeText(IIf(_sVal <> "", _sVal, "................") & vbCrLf)
    '            objselection.Font.Italic = 0

    '            objselection.TypeText(": " + _sVal + vbCrLf)
    '            objselection.TypeText("33. Ngày vào ngành Ngân hàng: " + _sVal + vbCrLf)
    '            If dr("Ngay_VBSP").ToString() <> "" Then _sVal = IIf(CType(dr("Ngay_VBSP").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("Ngay_VBSP").ToString(), DateTime).ToString("dd/MM/yyyy"), "")
    '            objselection.TypeText("34. Ngày vào NHCSXH: " + IIf(_sVal <> "", _sVal + vbTab + vbTab, vbTab + vbTab + vbTab + vbTab) + ", Tuyển dụng hoặc tiếp nhận: ")
    '            strSQL = "SELECT b.ten_goi,a.* FROM QDNhanSu a, DanhMuc b WHERE a.IdCanBo = '" + _RowId + "' And a.IsKiemNhiem = 0 And"
    '            strSQL += String.Format(" (a.IdDonvi_moi IN (Select id From ChiNhanh Where id = {0} or id_goc = {0}))", _IdDvParent)
    '            strSQL += " And (a.IdLoaiQD = b.id And b.id_goc = 15 And b.Status = 1 And b.ma_so IN('1508','1509')) ORDER BY NgayKy_QD ASC"
    '            Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If Not (_db Is Nothing) Then
    '                    If (_db.Rows.Count > 0) Then
    '                        _sVal = _db.Rows(0)("ten_goi").ToString().Trim()
    '                    End If
    '                End If
    '            End Using
    '            objselection.TypeText(_sVal + vbCrLf)
    '            objselection.TypeText("35. Sở trường công tác: " + dr("SoTruong_CT").ToString().Trim() + "," + vbTab + " Công việc đã làm lâu nhất: " + dr("CV_Lau").ToString().Trim() + vbCrLf)
    '            _sVal = ""
    '            If dr("CM_Ngay").ToString() <> "" Then _sVal = IIf(CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd/MM/yyyy"), "")

    '            objselection.TypeText("36. Ngày tham gia cách mạng: " + IIf(_sVal <> "", _sVal, vbTab + vbTab) + " Trong tổ chức nào: " + dr("CM_ToChuc").ToString().Trim() + vbCrLf)
    '            objselection.TypeText("37. Ngày vào Đảng Cộng sản Việt Nam: " + IIf(arrPartys(2) <> "", arrPartys(2) + ", ", vbTab + vbTab))
    '            objselection.TypeText(" Ngày chính thức: " + arrPartys(3) + vbCrLf)
    '            '38. Ngày tham gia các tổ chức chính trị, xã hội (Ngày vào Đoàn TNCSHCM, Công đoàn, Hội,...)
    '            objselection.TypeText("38. Ngày tham gia các tổ chức chính trị, xã hội ")
    '            objselection.Font.Size = 10
    '            objselection.TypeText("(Ngày vào Đoàn TNCSHCM, Công đoàn, Hội,...)")
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            _sVal = ""
    '            If (_sDoanvien <> "") Then
    '                If (_sDoanvien.IndexOf(";") >= 0) Then
    '                    _sVal = _sDoanvien.Substring(0, _sDoanvien.IndexOf(";"))
    '                End If
    '            End If
    '            objselection.TypeText(": " & _sVal & vbCrLf)
    '            _sVal = ""
    '            '39. Ngày nhập ngũ, Ngày xuất ngũ, Quân hàm, chức vụ cao nhất (Năm)
    '            strSQL = " SELECT a.IdHSLLVT,a.TuNgay,a.DenNgay,a.IdQuanHam, b.ten_goi As QuanHam,a.ChucVu,a.DonVi FROM HS_LLVT a, DanhMuc b "
    '            strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdQuanHam = b.id and b.Status = 1 and b.id_goc = 2) Order by a.TuNgay ASC"
    '            Dim ngay_nhap As String = ""
    '            Dim ngay_xuat As String = ""
    '            Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If Not (_db Is Nothing) Then
    '                    If (_db.Rows.Count > 0) Then
    '                        Dim _IdMin As Integer = CType(_db.Rows(0)("IdQuanHam").ToString().Trim(), Integer)
    '                        Dim _ChuvuQN As String = _db.Rows(0)("ChucVu").ToString().Trim()
    '                        ngay_nhap = CType(_db.Rows(0)("TuNgay").ToString(), DateTime).ToString("dd/MM/yyyy")
    '                        For i As Integer = 0 To _db.Rows.Count - 1
    '                            If (CType(_db.Rows(i)("IdQuanHam").ToString().Trim(), Integer) < _IdMin) Then
    '                                _IdMin = CType(_db.Rows(i)("IdQuanHam").ToString().Trim(), Integer)
    '                                _ChuvuQN = _db.Rows(i)("ChucVu").ToString().Trim()
    '                            End If

    '                            If (i = _db.Rows.Count - 1) Then
    '                                ngay_xuat = CType(_db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy")
    '                            End If
    '                        Next
    '                        _sVal = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DanhMuc Where id = {0} And id_goc = 2 and Status = 1", _IdMin)) + ", " + _ChuvuQN
    '                    End If
    '                End If
    '            End Using
    '            objselection.TypeText("39. Ngày nhập ngũ: " + IIf(ngay_nhap <> "", ngay_nhap, vbTab + vbTab) + ", Ngày xuất ngũ: " + IIf(ngay_xuat <> "", ngay_xuat, vbTab + vbTab) + ", Quân hàm, chức vụ cao nhất (năm): " + _sVal + vbCrLf)

    '            '40. Xuất dữ liệu: Quá trình đào tạo, Bồi dưỡng về chuyên môn, nghiệp vụ, lý luận chính trị, ngoại ngữ, tin học
    '            objselection.Font.Bold = 1
    '            objselection.TypeText("40. Quá trình đào tạo, Bồi dưỡng về chuyên môn, nghiệp vụ, lý luận chính trị, ngoại ngữ, tin học" + vbCrLf)
    '            objselection.Font.Bold = 0
    '            objselection.Font.Size = 11
    '            Dim _rowsTable As Integer = 2
    '            strSQL = "SELECT CoSo_DT,(Select ten_goi From DanhMuc Where IdChuyenNganhDT = Id and id_goc = 11 And Status =1) As ChuyenNganh,"
    '            strSQL += "NganhHoc,(Convert(Varchar, TuNgay, 103) + ' - ' + Convert(Varchar, DenNgay, 103)) As ThoiGian,"
    '            strSQL += "(Select ten_goi From DanhMuc Where id = IdHinhThucDT And id_goc = 8 And Status = 1) As HinhThuc,"
    '            strSQL += "(CASE VBCC WHEN 1 THEN 'VB'ELSE 'CC' END) As VBCC,"
    '            strSQL += "(Select ten_goi From DanhMuc Where id = IdTrinhDo And id_goc = 38 And Status = 1) As TrinhDo"
    '            strSQL += " FROM HS_DTVBCC WHERE IdCanBo = '" + _RowId + "' And HoanThanh = 1 Order By DenNgay Asc"
    '            Using db_vb As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_vb Is Nothing Or db_vb.Rows.Count = 0) Then
    '                    _rowsTable = 2
    '                Else
    '                    _rowsTable = db_vb.Rows.Count + 1
    '                End If
    '                'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 5 cột
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Tên trường")
    '                objTable.Cell(1, 2).Range.InsertAfter("Chuyên ngành - Khoa")
    '                objTable.Cell(1, 3).Range.InsertAfter("Thời gian học")
    '                objTable.Cell(1, 4).Range.InsertAfter("Hình thức học")
    '                objTable.Cell(1, 5).Range.InsertAfter("Văn bằng, chứng chỉ, trình độ gì")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                If Not (db_vb Is Nothing) Then
    '                    If (db_vb.Rows.Count > 0) Then
    '                        For i As Integer = 0 To db_vb.Rows.Count - 1
    '                            objTable.Cell(i + 2, 1).Range.InsertAfter(db_vb.Rows(i)("CoSo_DT").ToString().Trim())
    '                            objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 2).Range.InsertAfter(db_vb.Rows(i)("ChuyenNganh").ToString().Trim() + " - " + db_vb.Rows(i)("NganhHoc").ToString().Trim())
    '                            objTable.Cell(i + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 3).Range.InsertAfter(db_vb.Rows(i)("ThoiGian").ToString().Trim())
    '                            objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 2, 4).Range.InsertAfter(db_vb.Rows(i)("HinhThuc").ToString().Trim())
    '                            objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 2, 5).Range.InsertAfter(db_vb.Rows(i)("VBCC").ToString().Trim() + "-" + db_vb.Rows(i)("TrinhDo").ToString().Trim())
    '                            objTable.Cell(i + 2, 5).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                        Next
    '                    End If
    '                End If
    '                objTable.Columns(1).PreferredWidth = 23
    '                objTable.Columns(2).PreferredWidth = 27
    '                objTable.Columns(3).PreferredWidth = 27
    '                objTable.Columns(4).PreferredWidth = 13
    '                objTable.Columns(5).PreferredWidth = 17
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Font.Size = 10
    '            objselection.Range.Paragraphs.SpaceAfter = 4
    '            _sVal = "Ghi chú: Hình thức học: Chính quy, tại chức, chuyên tu, bồi dưỡng,.../ Văn bằng: Tiến sỹ, Phó TS, Thạc sỹ, Cử nhân, kỹ sư,..." + vbCrLf
    '            objselection.TypeText(_sVal)

    '            objselection.Font.Size = 12
    '            '41. Tóm tắt quá trình công tác:
    '            objselection.Font.Bold = 1
    '            objselection.TypeText("41. Tóm tắt quá trình công tác:" + vbCrLf)
    '            objselection.Font.Bold = 0
    '            objselection.Font.Size = 11
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            Using db_ct As DataTable = listQDNhansu(_RowId, 2)
    '                If (db_ct Is Nothing Or db_ct.Rows.Count = 0) Then
    '                    _rowsTable = 2
    '                Else
    '                    _rowsTable = db_ct.Rows.Count + 1
    '                End If
    '                'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 2 cột
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 2)
    '                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.PreferredWidth = 100
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Từ ngày tháng năm -" + vbCrLf + " Đến ngày tháng năm")
    '                objTable.Cell(1, 2).Range.InsertAfter("Chức danh, chức vụ, đơn vị công tác (Đảng, Chính quyền, Đoàn thể)")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                If Not (db_ct Is Nothing) Then
    '                    If (db_ct.Rows.Count > 0) Then
    '                        Dim row As Integer = 0
    '                        For i As Integer = 0 To db_ct.Rows.Count - 1
    '                            Dim ThoiGian As String = ""
    '                            If CInt(db_ct.Rows(i)("IsQD_NHCS")) = 0 Then
    '                                If i = 0 Then
    '                                    row = i
    '                                    ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
    '                                    Dim j As Integer = i + 1
    '                                    While j < db_ct.Rows.Count
    '                                        Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
    '                                            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516", "1532", "1534", "1535"
    '                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
    '                                                Exit While
    '                                            Case Else
    '                                                j = j + 1
    '                                        End Select
    '                                    End While
    '                                    objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                                    objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
    '                                    objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
    '                                    objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                                Else
    '                                    If Not (db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
    '                                            Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
    '                                            Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
    '                                        ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
    '                                        If Not (db_ct.Rows(i)("DenNgay") Is DBNull.Value) Then
    '                                            If Not (CType(db_ct.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
    '                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(db_ct.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
    '                                                row = row + 1
    '                                                objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                                                objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
    '                                                objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                                objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
    '                                                objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                                            Else

    '                                                Dim j As Integer = i + 1
    '                                                While j < db_ct.Rows.Count
    '                                                    Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
    '                                                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516", "1532", "1534", "1535"
    '                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
    '                                                            Exit While
    '                                                        Case Else
    '                                                            j = j + 1
    '                                                    End Select
    '                                                End While
    '                                                row = row + 1
    '                                                objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                                                objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
    '                                                objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                                objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
    '                                                objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F

    '                                            End If
    '                                        Else

    '                                            Dim j As Integer = i + 1
    '                                            While j < db_ct.Rows.Count
    '                                                Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
    '                                                    Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516", "1532", "1534", "1535"
    '                                                        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
    '                                                        Exit While
    '                                                    Case Else
    '                                                        j = j + 1
    '                                                End Select
    '                                            End While
    '                                            row = row + 1
    '                                            objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                                            objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
    '                                            objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                            objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
    '                                            objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                                        End If
    '                                    End If
    '                                End If
    '                            Else
    '                                'ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
    '                                'Dim j As Integer = i + 1
    '                                'While j < db_ct.Rows.Count
    '                                '    If (CInt(db_ct.Rows(i)("IDDonVi_Moi")) = CInt(db_ct.Rows(j)("IDDonVi_Moi")) And CInt(db_ct.Rows(i)("IdPhong_Moi")) = CInt(db_ct.Rows(j)("IdPhong_Moi")) And CInt(db_ct.Rows(i)("IDChucVu_Moi")) = CInt(db_ct.Rows(j)("IDChucVu_Moi"))) Then
    '                                '        j = j + 1
    '                                '    Else
    '                                '        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
    '                                '        i = j - 1
    '                                '        Exit While
    '                                '    End If
    '                                'End While
    '                                'If i <> 0 Then row = row + 1
    '                                'objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                                'objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
    '                                'objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                'objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
    '                                'objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                                'If j = db_ct.Rows.Count Then Exit For
    '                                ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
    '                                Dim j As Integer = i + 1
    '                                While j < db_ct.Rows.Count
    '                                    Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
    '                                        Case "1534", "1535"
    '                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
    '                                            Exit While
    '                                        Case Else
    '                                            If (CInt(db_ct.Rows(i)("IDDonVi_Moi")) = CInt(db_ct.Rows(j)("IDDonVi_Moi")) And CInt(db_ct.Rows(i)("IdPhong_Moi")) = CInt(db_ct.Rows(j)("IdPhong_Moi")) And CInt(db_ct.Rows(i)("IDChucVu_Moi")) = CInt(db_ct.Rows(j)("IDChucVu_Moi"))) Then
    '                                                j = j + 1
    '                                            Else
    '                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
    '                                                i = j - 1
    '                                                Exit While
    '                                            End If
    '                                    End Select
    '                                End While
    '                                If i <> 0 Then row = row + 1
    '                                objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                                objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
    '                                objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1534" Then
    '                                    objTable.Cell(row + 2, 2).Range.InsertAfter(db_ct.Rows(i)("LoaiQD").ToString())
    '                                Else
    '                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
    '                                End If
    '                                objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                                If j = db_ct.Rows.Count Then Exit For
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                objTable.Columns(1).PreferredWidth = 24
    '                objTable.Columns(2).PreferredWidth = 76
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Font.Name = "Times New Roman"
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            objselection.Paragraphs.SpaceBefore = 0
    '            objselection.Paragraphs.SpaceAfter = 4
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            objselection.TypeText("42. Đặc điểm, lý lịch bản thân: " + vbCrLf)
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 1
    '            If (dr("DacDiem_BT").ToString().Trim() <> "") Then
    '                objselection.TypeText(vbTab + "- " + dr("DacDiem_BT").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
    '            Else
    '                _sVal = "      a - Khai rõ: bị bắt, bị tù (từ ngày tháng năm nào đến ngày tháng năm nào, ở đâu), đã khai báo cho ai, những vấn đề gì:" + vbCrLf
    '                _sVal += "     .........................................................................................................................................................................." + vbCrLf
    '                _sVal += "     .........................................................................................................................................................................." + vbCrLf
    '                objselection.TypeText(_sVal)
    '            End If
    '            If (dr("GhiChu").ToString().Trim() <> "") Then
    '                objselection.TypeText(vbTab + "- " + dr("GhiChu").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
    '            Else
    '                _sVal = "      b - Bản thân có làm việc trong chế độ cũ (Cơ quan, đơn vị nào, địa điểm, chức danh, chức vụ, thời gian làm việc):" + vbCrLf
    '                _sVal += "     ..........................................................................................................................................................................." + vbCrLf
    '                _sVal += "     ..........................................................................................................................................................................." + vbCrLf
    '                objselection.TypeText(_sVal)
    '            End If
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            objselection.TypeText("43. Quan hệ với nước ngoài: " + vbCrLf)
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 1
    '            If (dr("QuanHe_Nguoi_NN").ToString().Trim() <> "") Then
    '                objselection.TypeText(vbTab + "- " + dr("QuanHe_Nguoi_NN").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
    '            Else
    '                _sVal = "      Tham gia hoặc có quan hệ với các tổ chức chính trị, kinh tế, xã hội nào ở nước ngoài (làm gì, tổ chức nào, đặt trụ sở ở đâu?): .................................................................................................................................." + vbCrLf
    '                _sVal += "     ............................................................................................................................................................................" + vbCrLf
    '                _sVal += "     ............................................................................................................................................................................" + vbCrLf
    '                _sVal += "     ............................................................................................................................................................................" + vbCrLf
    '                _sVal += "     Có thân nhân (Bố mẹ, vợ, chồng, con, anh chị em ruột) ở nước ngoài (làm gì, địa chỉ?): ....................................................." + vbCrLf
    '                _sVal += "     ............................................................................................................................................................................" + vbCrLf
    '                _sVal += "     ............................................................................................................................................................................" + vbCrLf
    '                _sVal += "     ............................................................................................................................................................................" + vbCrLf
    '                objselection.TypeText(_sVal)
    '            End If
    '            '44. Thông tin tài khoản ngân hàng
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            objselection.TypeText("44. Thông tin tài khoản ngân hàng: " + vbTab)
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("- Mã khách hàng: " + dr("NH_MaKH").ToString().Trim() + vbTab + " - Số tài khoản: " + dr("NH_SoTK").ToString().Trim() + vbCrLf)
    '            objselection.Font.Size = 10
    '            objselection.TypeText(vbTab + "(Tài khoản trả lương)" + vbTab)
    '            objselection.Font.Size = 12
    '            objselection.TypeText(vbTab + "- Tên ngân hàng mở tài khoản: " + dr("NH_TenNH").ToString().Trim() + vbCrLf)
    '            objselection.Font.Italic = 0
    '            objselection.TypeText("45. Mã số thế cá nhân: " + dr("MaSoThue").ToString().Trim() + vbCrLf)
    '            objselection.TypeText("46. Bảo hiểm xã hội: " + vbTab)
    '            objselection.Font.Italic = 1

    '            objselection.TypeText("- Số sổ: " + dr("BHXH_SoSo").ToString().Trim() + vbTab)
    '            _sVal = ""
    '            If dr("BHXH_NgaySo").ToString() <> "" Then _sVal = IIf(CType(dr("BHXH_NgaySo").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("BHXH_NgaySo").ToString(), DateTime).ToString("dd/MM/yyyy"), "")
    '            objselection.TypeText("- Ngày cấp sổ: " + _sVal + vbTab)
    '            _sVal = ""
    '            If dr("BHXH_NgayBatDau").ToString() <> "" Then _sVal = IIf(CType(dr("BHXH_NgayBatDau").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("BHXH_NgayBatDau").ToString(), DateTime).ToString("dd/MM/yyyy"), "")
    '            objselection.TypeText("- Ngày bắt đầu tham gia BHXH: " + _sVal + vbCrLf)
    '            objselection.TypeText(vbTab + vbTab + vbTab + "- Nơi cấp sổ: " + dr("BHXH_NoiLam").ToString().Trim() + vbCrLf)
    '            objselection.Font.Italic = 0
    '            '47. Tham gia Đoàn công tác nước ngoài (do NHCSXH cử)
    '            objselection.TypeText("47. Tham gia Đoàn công tác nước ngoài (do NHCSXH cử)" + vbCrLf)
    '            strSQL = "SELECT a.SoQD,Convert(Varchar,a.NgayKy_QD,103) As NgayKy_QD,b.ten_goi as NuocDen,a.MucDich,Convert(Varchar,a.TuNgay,103) As TG_TuNgay,"
    '            strSQL += "Convert(Varchar,a.DenNgay,103) As TG_DenNgay,a.NguoiKy_QD FROM HS_XuatNgoai a, QuocGia b "
    '            strSQL += " WHERE a.IdCanBo = '" + _RowId + "' And (a.IdNuocDen = b.id and b.Status = 1) Order by a.TuNgay Asc"
    '            objselection.Font.Bold = 0
    '            objselection.Font.Size = 11
    '            Using db_xn As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_xn Is Nothing Or db_xn.Rows.Count = 0) Then
    '                    _rowsTable = 3
    '                Else
    '                    _rowsTable = db_xn.Rows.Count + 2
    '                End If
    '                'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 4 cột
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 7)
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Columns(1).PreferredWidth = 7
    '                objTable.Columns(2).PreferredWidth = 7
    '                objTable.Columns(3).PreferredWidth = 14
    '                objTable.Columns(4).PreferredWidth = 32
    '                objTable.Columns(5).PreferredWidth = 7
    '                objTable.Columns(6).PreferredWidth = 7
    '                objTable.Columns(7).PreferredWidth = 12
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Quyết định số")
    '                objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
    '                objTable.Cell(1, 3).Range.InsertAfter("Nước đến")
    '                objTable.Cell(1, 4).Range.InsertAfter("Mục đích")
    '                objTable.Cell(1, 5).Range.InsertAfter("Thời gian công tác")
    '                objTable.Cell(1, 7).Range.InsertAfter("Người ký quyết định")
    '                objTable.Cell(2, 5).Range.InsertAfter("Từ ngày, tháng")
    '                objTable.Cell(2, 6).Range.InsertAfter("Đến ngày, tháng")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Rows(2).Range.Font.Bold = 1
    '                If Not (db_xn Is Nothing) Then
    '                    If (db_xn.Rows.Count > 0) Then
    '                        For i As Integer = 0 To db_xn.Rows.Count - 1
    '                            objTable.Cell(i + 3, 1).Range.InsertAfter(db_xn.Rows(i)("SoQD").ToString().Trim())
    '                            objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 3, 2).Range.InsertAfter(db_xn.Rows(i)("NgayKy_QD").ToString().Trim())
    '                            objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 3).Range.InsertAfter(db_xn.Rows(i)("NuocDen").ToString().Trim())
    '                            objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 4).Range.InsertAfter(db_xn.Rows(i)("MucDich").ToString().Trim())
    '                            objTable.Cell(i + 3, 4).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 3, 5).Range.InsertAfter(db_xn.Rows(i)("TG_TuNgay").ToString().Trim())
    '                            objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 6).Range.InsertAfter(db_xn.Rows(i)("TG_DenNgay").ToString().Trim())
    '                            objTable.Cell(i + 3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 7).Range.InsertAfter(db_xn.Rows(i)("NguoiKy_QD").ToString().Trim())
    '                            objTable.Cell(i + 3, 7).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                        Next
    '                    End If
    '                End If
    '                objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
    '                objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
    '                objTable.Cell(1, 3).Merge(objTable.Cell(2, 3))
    '                objTable.Cell(1, 4).Merge(objTable.Cell(2, 4))
    '                objTable.Cell(1, 7).Merge(objTable.Cell(2, 7))
    '                objTable.Cell(1, 5).Merge(objTable.Cell(1, 6))
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            '48. Hồ sơ Hộ chiếu của cán bộ
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Color = Word.WdColor.wdColorAutomatic
    '            objselection.Font.Name = "Times New Roman"
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 0
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            objselection.TypeText("48. Hộ chiếu" + vbCrLf)
    '            strSQL = "SELECT So_HoChieu, (Case Loai_HC When 1 Then N'Phổ thông' Else ((Case Loai_HC When 2 Then N'Công vụ' Else N'Ngoại giao' End)) End) As LoaiHC,Convert(Varchar,NgayCap,103) As NgCapHC, NoiCap, Convert(Varchar,NgayHH,103) As NgHetHanHC,"
    '            strSQL += "(Case TinhTrang When 1 Then N'Còn hiệu lực' Else ((Case TinhTrang When 2 Then N'Hết hạn' Else N'Mất' End)) End) As TinhTrang, GhiChu FROM HS_HoChieu Where IdCanBo = '" + _RowId + "' Order by NgayCap Asc"
    '            objselection.Font.Bold = 0
    '            objselection.Font.Size = 11
    '            Using db_hc As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_hc Is Nothing Or db_hc.Rows.Count = 0) Then
    '                    _rowsTable = 2
    '                Else
    '                    _rowsTable = db_hc.Rows.Count + 1
    '                End If
    '                'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 7 cột
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 7)
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '                objTable.Columns(1).PreferredWidth = 10
    '                objTable.Columns(2).PreferredWidth = 13
    '                objTable.Columns(3).PreferredWidth = 12
    '                objTable.Columns(4).PreferredWidth = 26
    '                objTable.Columns(5).PreferredWidth = 12
    '                objTable.Columns(6).PreferredWidth = 13
    '                objTable.Columns(7).PreferredWidth = 20
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Số hộ chiếu")
    '                objTable.Cell(1, 2).Range.InsertAfter("Loại hộ chiếu")
    '                objTable.Cell(1, 3).Range.InsertAfter("Ngày cấp")
    '                objTable.Cell(1, 4).Range.InsertAfter("Nơi cấp")
    '                objTable.Cell(1, 5).Range.InsertAfter("Ngày hết hạn")
    '                objTable.Cell(1, 6).Range.InsertAfter("Tình trạng")
    '                objTable.Cell(1, 7).Range.InsertAfter("Ghi chú thêm")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                If Not (db_hc Is Nothing) Then
    '                    If (db_hc.Rows.Count > 0) Then
    '                        For i As Integer = 0 To db_hc.Rows.Count - 1
    '                            objTable.Cell(i + 2, 1).Range.InsertAfter(db_hc.Rows(i)("So_HoChieu").ToString().Trim())
    '                            objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 2).Range.InsertAfter(db_hc.Rows(i)("LoaiHC").ToString().Trim())
    '                            objTable.Cell(i + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 3).Range.InsertAfter(db_hc.Rows(i)("NgCapHC").ToString().Trim())
    '                            objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 2, 4).Range.InsertAfter(db_hc.Rows(i)("NoiCap").ToString().Trim())
    '                            objTable.Cell(i + 2, 4).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 5).Range.InsertAfter(db_hc.Rows(i)("NgHetHanHC").ToString().Trim())
    '                            objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 2, 6).Range.InsertAfter(db_hc.Rows(i)("TinhTrang").ToString().Trim())
    '                            objTable.Cell(i + 2, 6).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 7).Range.InsertAfter(db_hc.Rows(i)("GhiChu").ToString().Trim())
    '                            objTable.Cell(i + 2, 7).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                        Next
    '                    End If
    '                End If
    '            End Using

    '            '49. Quan hệ gia đình
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Color = Word.WdColor.wdColorAutomatic
    '            objselection.Font.Name = "Times New Roman"
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 1
    '            objselection.Font.Bold = 1
    '            objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            objselection.TypeText("49. Quan hệ gia đình:" + vbCrLf)
    '            objselection.Font.Bold = 0
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.TypeText("a. Bố, Mẹ, Vợ (Chồng), các con, anh chị em ruột" + vbCrLf)
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 0
    '            'Tạo một bảng có _rowsTable dòng và 4 cột
    '            'Tạo một bảng có _rowsTable dòng và 4 cột
    '            strSQL = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2303','2304','2314','2313','2312','2317','2305','2320','2315','2316'))"
    '            Using db_rows As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_rows Is Nothing Or CType(db_rows.Rows(0)(0), Byte) = 0) Then
    '                    _rowsTable = 6
    '                Else
    '                    Dim db As DBAccess = New DBAccess
    '                    Dim ssql As String = ""
    '                    'Dem So con
    '                    ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2312','2317'))"
    '                    If db.getNumber(ssql) = 0 Then
    '                        _rowsTable = 5
    '                    Else
    '                        _rowsTable = 4 + db.getNumber(ssql)
    '                    End If
    '                    'Dem So anh chi em
    '                    ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316'))"
    '                    _rowsTable = _rowsTable + db.getNumber(ssql)
    '                    '_rowsTable = CType(db_rows.Rows(0)(0), Byte) + 1
    '                    If _rowsTable <= 6 Then _rowsTable = 6
    '                End If
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
    '                objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
    '                objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
    '                objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
    '                objTable.Cell(1, 4).Range.Font.Bold = 0
    '                objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(2, 1).Range.InsertAfter("Bố")
    '                objTable.Cell(2, 1).Range.Font.Bold = 1
    '                objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
    '                objTable.Cell(3, 1).Range.Font.Bold = 1
    '                objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                If (dr("GioiTinh").ToString() = False) Then
    '                    objTable.Cell(4, 1).Range.InsertAfter("Vợ" + vbCrLf)
    '                Else
    '                    objTable.Cell(4, 1).Range.InsertAfter("Chồng" + vbCrLf)
    '                End If
    '                objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(4, 1).Range.Font.Bold = 1
    '                'If _rowsTable = 6 And CType(db_rows.Rows(0)(0), Byte) = 0 Then
    '                '    objTable.Cell(5, 1).Range.InsertAfter("Các" + vbCrLf + "con")
    '                '    objTable.Cell(5, 1).Range.Font.Bold = 1
    '                '    objTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                '    objTable.Cell(6, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
    '                '    objTable.Cell(6, 1).Range.Font.Bold = 1
    '                '    objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                'End If
    '                For i As Byte = 2 To _rowsTable
    '                    objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
    '                Next
    '                If Not (db_rows Is Nothing) Then
    '                    If (db_rows.Rows.Count > 0) Then
    '                        'Lấy thông tin Bố đẻ
    '                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2303')"
    '                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                            If Not (db_gd Is Nothing) Then
    '                                If (db_gd.Rows.Count > 0) Then
    '                                    objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                    _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                                    objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                                    objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
    '                                    objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
    '                                End If
    '                            End If
    '                        End Using

    '                        'Lấy thông tin Mẹ đẻ
    '                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2304')"
    '                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                            If Not (db_gd Is Nothing) Then
    '                                If (db_gd.Rows.Count > 0) Then
    '                                    objTable.Cell(3, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(3, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                    _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                                    objTable.Cell(3, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                                    objTable.Cell(3, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
    '                                    objTable.Cell(3, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
    '                                End If
    '                            End If
    '                        End Using

    '                        'Lấy thông tin Vợ hoặc chồng
    '                        Dim rowCVC As Integer = 4  'Danh dau dong Chong/Vo/Con
    '                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2314','2313'))"
    '                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                            If Not (db_gd Is Nothing) Then
    '                                If (db_gd.Rows.Count > 0) Then
    '                                    objTable.Cell(4, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(4, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                    _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                                    objTable.Cell(4, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                                    objTable.Cell(4, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
    '                                    objTable.Cell(4, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
    '                                Else
    '                                    objTable.Cell(4, 2).Range.InsertAfter(vbCrLf)
    '                                    objTable.Cell(4, 3).Range.InsertAfter(vbCrLf)
    '                                    objTable.Cell(4, 4).Range.InsertAfter(vbCrLf)
    '                                End If
    '                            End If
    '                        End Using

    '                        'Lấy thông tin Các con đẻ
    '                        Dim _RowCon As Byte = 0
    '                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2312','2317'))"
    '                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                            If db_gd.Rows.Count = 0 Then
    '                                objTable.Cell(5, 1).Range.InsertAfter("Các" + vbCrLf + "con")
    '                                objTable.Cell(5, 1).Range.Font.Bold = 1
    '                                objTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            End If
    '                            If Not (db_gd Is Nothing) Then
    '                                If (db_gd.Rows.Count > 0) Then
    '                                    _RowCon = db_gd.Rows.Count
    '                                    For i As Integer = 0 To db_gd.Rows.Count - 1
    '                                        If (i = 0) Then
    '                                            objTable.Cell(i + 5, 1).Range.InsertAfter("Các" + vbCrLf + "con")
    '                                        End If
    '                                        objTable.Cell(i + 5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                        objTable.Cell(i + 5, 1).Range.Font.Bold = 1
    '                                        objTable.Cell(i + 5, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                                        objTable.Cell(i + 5, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                                        objTable.Cell(i + 5, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                        _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                                        objTable.Cell(i + 5, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                                        objTable.Cell(i + 5, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
    '                                        objTable.Cell(i + 5, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
    '                                        If (i > 0) Then objTable.Rows(i + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
    '                                    Next
    '                                End If
    '                            End If
    '                            '                                If _RowCon = db_gd.Rows.Count Then objTable.Rows(_RowCon - 1 + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                        End Using

    '                        'Lấy thông tin Anh chị em ruột thịt
    '                        Dim _rowAnhChi As Integer = 0
    '                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316'))"
    '                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                            If db_gd.Rows.Count = 0 Then
    '                                If _RowCon = 0 Then
    '                                    objTable.Cell(6, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
    '                                    objTable.Cell(6, 1).Range.Font.Bold = 1
    '                                    objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                Else
    '                                    objTable.Cell(_RowCon + 5, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
    '                                    objTable.Cell(_RowCon + 5, 1).Range.Font.Bold = 1
    '                                    objTable.Cell(_RowCon + 5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                End If
    '                            End If

    '                            If Not (db_gd Is Nothing) Then
    '                                If (db_gd.Rows.Count > 0) Then
    '                                    _rowAnhChi = db_gd.Rows.Count
    '                                    If _RowCon = 0 Then _RowCon = 1
    '                                    For i As Integer = 0 To db_gd.Rows.Count - 1
    '                                        If (i = 0) Then
    '                                            objTable.Cell(i + _RowCon + 5, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
    '                                            objTable.Rows(i + _RowCon + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                        End If
    '                                        objTable.Cell(i + _RowCon + 5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                        objTable.Cell(i + _RowCon + 5, 1).Range.Font.Bold = 1
    '                                        objTable.Cell(i + _RowCon + 5, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                                        objTable.Cell(i + _RowCon + 5, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                                        objTable.Cell(i + _RowCon + 5, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                        _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                                        objTable.Cell(i + _RowCon + 5, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                                        objTable.Cell(i + _RowCon + 5, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
    '                                        objTable.Cell(i + _RowCon + 5, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
    '                                        'If (i > 0) Then objTable.Rows(i + _RowCon + 3).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
    '                                        If (i > 0) Then objTable.Rows(i + _RowCon + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
    '                                    Next
    '                                    objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                    objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                    objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                    objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                End If
    '                            End If
    '                        End Using
    '                        'Merge ô cuối trước
    '                        If (_rowAnhChi > 1) Then
    '                            'objTable.Cell(_RowCon + 4, 1).Merge(objTable.Cell(_rowsTable, 1))
    '                            objTable.Cell(_RowCon + 5, 1).Merge(objTable.Cell(_rowsTable, 1))
    '                        End If
    '                        'If _RowCon > 1 Then objTable.Cell(4, 1).Merge(objTable.Cell(3 + _RowCon, 1))
    '                        If _RowCon > 1 Then objTable.Cell(5, 1).Merge(objTable.Cell(5 + _RowCon - 1, 1))

    '                    End If
    '                End If
    '                'Thực hiện Merge Cell Anh chị em ruột
    '                objTable.Columns(1).PreferredWidth = 9
    '                objTable.Columns(2).PreferredWidth = 18
    '                objTable.Columns(3).PreferredWidth = 9
    '                objTable.Columns(4).PreferredWidth = 59
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            'Mục 49 b (Bố mẹ, anh chị em ruột bên vợ hoặc chồng)
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("b. Bố, mẹ, anh chị em ruột bên vợ hoặc chồng: " + vbCrLf)
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 0
    '            strSQL = "SELECT count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2326','2327','2328','2329','2330','2331'))"
    '            Using db_gdvochong As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_gdvochong Is Nothing Or CType(db_gdvochong.Rows(0)(0), Byte) = 0) Then
    '                    _rowsTable = 4
    '                Else
    '                    _rowsTable = CType(db_gdvochong.Rows(0)(0), Byte) + 3
    '                    If _rowsTable <= 4 Then _rowsTable = 4
    '                End If
    '                'Tạo một bảng có 4 dòng và 4 cột
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
    '                objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
    '                objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
    '                objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
    '                objTable.Cell(1, 4).Range.Font.Bold = 0
    '                objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(2, 1).Range.InsertAfter("Bố")
    '                objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
    '                objTable.Cell(4, 1).Range.InsertAfter("Anh" + vbCrLf + "chị em" + vbCrLf + "ruột")
    '                objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Cell(2, 1).Range.Font.Bold = 1
    '                objTable.Cell(3, 1).Range.Font.Bold = 1
    '                objTable.Cell(4, 1).Range.Font.Bold = 1
    '                For i As Integer = 2 To _rowsTable
    '                    objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
    '                Next
    '                'Lấy thông tin Bố vợ (Bố chồng)
    '                strSQL = "SELECT Top 1 * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2322','2324'))"
    '                Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                    If Not (db_gd Is Nothing) Then
    '                        If (db_gd.Rows.Count > 0) Then
    '                            objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                            objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                            objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                            objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                            objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
    '                            objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
    '                        End If
    '                    End If
    '                End Using
    '                'Lấy thông tin Mẹ vợ (Mẹ chồng)
    '                strSQL = "SELECT Top 1 * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2323','2325'))"
    '                Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                    If Not (db_gd Is Nothing) Then
    '                        If (db_gd.Rows.Count > 0) Then
    '                            objTable.Cell(3, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                            objTable.Cell(3, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                            objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                            objTable.Cell(3, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                            objTable.Cell(3, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
    '                            objTable.Cell(3, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
    '                        End If
    '                    End If
    '                End Using
    '                If (CType(db_gdvochong.Rows(0)(0), Byte) > 0) Then
    '                    'Lấy thông tin Anh chị em ruột bên vợ (hoặc bên chồng)
    '                    strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2326','2327','2328','2329','2330','2331'))"
    '                    Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                        If Not (db_gd Is Nothing) Then
    '                            If (db_gd.Rows.Count > 0) Then
    '                                For i As Integer = 0 To db_gd.Rows.Count - 1
    '                                    objTable.Cell(i + 4, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(i + 4, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
    '                                    objTable.Cell(i + 4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                                    _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
    '                                    objTable.Cell(i + 4, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
    '                                    objTable.Cell(i + 4, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
    '                                    objTable.Cell(i + 4, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
    '                                    If (i > 0) Then
    '                                        objTable.Rows(i + 4).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
    '                                    End If
    '                                Next
    '                                'objTable.Cell(_rowsTable, 1).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                                'Thực hiện Merge Cell Anh chị em ruột
    '                                If (_rowsTable > 4) Then objTable.Cell(4, 1).Merge(objTable.Cell(_rowsTable, 1))
    '                            End If
    '                        End If
    '                    End Using
    '                End If
    '                objTable.Columns(1).PreferredWidth = 9
    '                objTable.Columns(2).PreferredWidth = 18
    '                objTable.Columns(3).PreferredWidth = 9
    '                objTable.Columns(4).PreferredWidth = 59
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            '49c. Hồ sơ Giảm trừ gia cảnh của cán bộ
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 1
    '            objselection.TypeText("c. Hồ sơ giảm trừ gia cảnh" + vbCrLf)
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 0
    '            strSQL = "SELECT b.ten_goi as QuanHe,a.HoTenNguoiPT,Convert(Varchar,a.TuNgay,103) as Tg_TuNgay,Convert(Varchar,a.DenNgay,103) as Tg_DenNgay,a.GhiChu FROM HS_GTGC a, DanhMuc b Where (a.IdQuanHe = b.id And b.id_goc = 23) And a.IdCanBo = '" + _RowId + "' Order by TuNgay Asc"
    '            Using db_gtgc As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_gtgc Is Nothing Or db_gtgc.Rows.Count = 0) Then
    '                    _rowsTable = 3
    '                Else
    '                    _rowsTable = db_gtgc.Rows.Count + 2
    '                End If
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
    '                objTable.Cell(1, 2).Range.InsertAfter("Họ và tên người phụ thuộc")
    '                objTable.Cell(1, 3).Range.InsertAfter("Thời gian kê khai")
    '                objTable.Cell(2, 3).Range.InsertAfter("Từ ngày")
    '                objTable.Cell(2, 4).Range.InsertAfter("Đến ngày")
    '                objTable.Cell(1, 5).Range.InsertAfter("Ghi chú")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Rows(2).Range.Font.Bold = 1
    '                If Not (db_gtgc Is Nothing) Then
    '                    If (db_gtgc.Rows.Count > 0) Then
    '                        For i As Integer = 0 To db_gtgc.Rows.Count - 1
    '                            objTable.Cell(i + 3, 1).Range.InsertAfter(db_gtgc.Rows(i)("QuanHe").ToString().Trim())
    '                            objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 3, 2).Range.InsertAfter(db_gtgc.Rows(i)("HoTenNguoiPT").ToString().Trim())
    '                            objTable.Cell(i + 3, 2).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 3, 3).Range.InsertAfter(db_gtgc.Rows(i)("Tg_TuNgay").ToString().Trim())
    '                            objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 4).Range.InsertAfter(db_gtgc.Rows(i)("Tg_DenNgay").ToString().Trim())
    '                            objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 5).Range.InsertAfter(db_gtgc.Rows(i)("GhiChu").ToString().Trim())
    '                            objTable.Cell(i + 3, 5).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                        Next
    '                    End If
    '                End If
    '                objTable.Columns(1).PreferredWidth = 10
    '                objTable.Columns(2).PreferredWidth = 40
    '                objTable.Columns(3).PreferredWidth = 10
    '                objTable.Columns(4).PreferredWidth = 10
    '                objTable.Columns(5).PreferredWidth = 28
    '                objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
    '                objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
    '                objTable.Cell(1, 5).Merge(objTable.Cell(2, 5))
    '                objTable.Cell(1, 3).Merge(objTable.Cell(1, 4))
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            ''50. Hồ sơ bổ nhiệm (Nếu có)
    '            'objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            'objselection.Paragraphs.SpaceBefore = 10
    '            'objselection.Font.Size = 12
    '            'objselection.Font.Italic = 1
    '            'objselection.Font.Bold = 1
    '            'objselection.TypeText("50. Hồ sơ bổ nhiệm (nếu có)" + vbCrLf)
    '            'objselection.Paragraphs.SpaceBefore = 3
    '            'objselection.Paragraphs.SpaceAfter = 3
    '            'objselection.Font.Size = 11
    '            'objselection.Font.Italic = 0
    '            'objselection.Font.Bold = 0
    '            ''Tạo một bảng gồm 5 cột và _rowsTable dòng
    '            'strSQL = "SELECT So_QD,Convert(Varchar,NgayKy_QD,103) As NgayKy,"
    '            'strSQL += "(Select ten_goi From DanhMuc Where id = IdChucVu_Moi and id_goc = 14) As ChucDanh,"
    '            'strSQL += "Convert(Varchar,NgayHL,103) As NgayHL,Convert(Varchar,NgayBoNhiem_TT,103) As NgayBoNhiem_TT,"
    '            'strSQL += "NguoiKy_QD FROM QDNhanSu WHERE IdLoaiQD IN (Select [id] From DanhMuc Where ma_so IN ('1502','1503','1504')) "
    '            'strSQL += " And IsKiemNhiem = 0 And IdCanBo = 'CNTT00000000016' Order By NgayKy_QD Asc"
    '            'Using db_hsbn As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '            '    If (db_hsbn Is Nothing Or db_hsbn.Rows.Count = 0) Then
    '            '        _rowsTable = 3
    '            '    Else
    '            '        _rowsTable = db_hsbn.Rows.Count + 2
    '            '    End If
    '            '    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 6)
    '            '    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '            '    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '            '    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '            '    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '            '    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '            '    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '            '    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '            '    objTable.Range.Paragraphs.SpaceBefore = 2
    '            '    objTable.Range.Paragraphs.SpaceAfter = 2
    '            '    'Add tiêu đề của các cột trong Bảng
    '            '    objTable.Cell(1, 1).Range.InsertAfter("Số QĐ")
    '            '    objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
    '            '    objTable.Cell(1, 3).Range.InsertAfter("Chức danh")
    '            '    objTable.Cell(1, 4).Range.InsertAfter("Hiệu lực thi hành")
    '            '    objTable.Cell(2, 4).Range.InsertAfter("Từ ngày")
    '            '    objTable.Cell(2, 5).Range.InsertAfter("Đến ngày")
    '            '    objTable.Cell(1, 6).Range.InsertAfter("Người ký QĐ")
    '            '    objTable.Rows(1).Range.Font.Bold = 1
    '            '    objTable.Rows(2).Range.Font.Bold = 1
    '            '    If Not (db_hsbn Is Nothing) Then
    '            '        If (db_hsbn.Rows.Count > 0) Then
    '            '            For i As Integer = 0 To db_hsbn.Rows.Count - 1
    '            '                objTable.Cell(i + 3, 1).Range.InsertAfter(db_hsbn.Rows(i)("So_QD").ToString().Trim())
    '            '                objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
    '            '                objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            '                objTable.Cell(i + 3, 2).Range.InsertAfter(db_hsbn.Rows(i)("NgayKy").ToString().Trim())
    '            '                objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '            '                objTable.Cell(i + 3, 3).Range.InsertAfter(db_hsbn.Rows(i)("ChucDanh").ToString().Trim())
    '            '                objTable.Cell(i + 3, 3).Range.Paragraphs.LeftIndent = 3.0F
    '            '                objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            '                objTable.Cell(i + 3, 4).Range.InsertAfter(db_hsbn.Rows(i)("NgayHL").ToString().Trim())
    '            '                objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '            '                objTable.Cell(i + 3, 5).Range.InsertAfter(db_hsbn.Rows(i)("NgayBoNhiem_TT").ToString().Trim())
    '            '                objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '            '                objTable.Cell(i + 3, 6).Range.InsertAfter(db_hsbn.Rows(i)("NguoiKy_QD").ToString().Trim())
    '            '                objTable.Cell(i + 3, 6).Range.Paragraphs.LeftIndent = 3.0F
    '            '                objTable.Cell(i + 3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '            '            Next
    '            '        End If
    '            '    End If
    '            '    objTable.Columns(1).PreferredWidth = 10
    '            '    objTable.Columns(2).PreferredWidth = 10
    '            '    objTable.Columns(3).PreferredWidth = 30
    '            '    objTable.Columns(4).PreferredWidth = 10
    '            '    objTable.Columns(5).PreferredWidth = 10
    '            '    objTable.Columns(6).PreferredWidth = 28
    '            '    objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
    '            '    objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
    '            '    objTable.Cell(1, 3).Merge(objTable.Cell(2, 3))
    '            '    objTable.Cell(1, 6).Merge(objTable.Cell(2, 6))
    '            '    objTable.Cell(1, 4).Merge(objTable.Cell(1, 5))
    '            '    objTable.LeftPadding = -3
    '            '    objTable.RightPadding = -4
    '            'End Using

    '            '51. Khen thưởng
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 1
    '            objselection.Font.Bold = 1
    '            objselection.TypeText("50. Khen thưởng" + vbCrLf)
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 0
    '            objselection.Font.Bold = 0
    '            strSQL = "SELECT SoQD, Convert(Varchar,NgayQD,103) as NgayQD,(Select ten_goi From DanhMuc Where id = IdCapKT And id_goc = 4) As ThamQuyenQD,"
    '            strSQL += "(Select DanhHieu_HinhThuc From ThiDuaKhenThuong Where idTDKT = IdDanhHieuHinhThuc) As DanhHieu,'' As HinhThuc"
    '            strSQL += " FROM HS_khenthuong t1, HS_KhenThuong_CT t2 where t1.IdKhenThuong=t2.IdKhenthuong AND t2.IdCN_TT = '" + _RowId + "'"
    '            strSQL += " ORDER BY NgayQD ASC"
    '            Using db_khenthuong As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_khenthuong Is Nothing Or db_khenthuong.Rows.Count = 0) Then
    '                    _rowsTable = 2
    '                Else
    '                    _rowsTable = db_khenthuong.Rows.Count + 1
    '                End If
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Số QĐ")
    '                objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
    '                objTable.Cell(1, 3).Range.InsertAfter("Thẩm quyền QĐ")
    '                objTable.Cell(1, 4).Range.InsertAfter("Danh hiệu" + vbCrLf + "(từ chiến sỹ thi đua cơ sở trở lên)")
    '                objTable.Cell(1, 5).Range.InsertAfter("Hình thức" + vbCrLf + "(từ Bằng khen)")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                If Not (db_khenthuong Is Nothing) Then
    '                    If (db_khenthuong.Rows.Count > 0) Then
    '                        For i As Integer = 0 To db_khenthuong.Rows.Count - 1
    '                            objTable.Cell(i + 2, 1).Range.InsertAfter(db_khenthuong.Rows(i)("SoQD").ToString().Trim())
    '                            objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 2).Range.InsertAfter(db_khenthuong.Rows(i)("NgayQD").ToString().Trim())
    '                            objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 2, 3).Range.InsertAfter(db_khenthuong.Rows(i)("ThamQuyenQD").ToString().Trim())
    '                            objTable.Cell(i + 2, 3).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 4).Range.InsertAfter(db_khenthuong.Rows(i)("DanhHieu").ToString().Trim())
    '                            objTable.Cell(i + 2, 4).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 2, 5).Range.InsertAfter(db_khenthuong.Rows(i)("HinhThuc").ToString().Trim())
    '                            objTable.Cell(i + 2, 5).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                        Next
    '                    End If
    '                End If
    '                objTable.Columns(1).PreferredWidth = 10
    '                objTable.Columns(2).PreferredWidth = 10
    '                objTable.Columns(3).PreferredWidth = 20
    '                objTable.Columns(4).PreferredWidth = 38
    '                objTable.Columns(5).PreferredWidth = 20
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            '52. Kỷ luật
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Size = 12
    '            objselection.Font.Italic = 1
    '            objselection.Font.Bold = 1
    '            objselection.TypeText("51. Kỷ luật" + vbCrLf)
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.Font.Size = 11
    '            objselection.Font.Italic = 0
    '            objselection.Font.Bold = 0
    '            strSQL = "SELECT SoQD,Convert(Varchar,NgayQD,103) As NgayKyQD,(Select ten_goi From DanhMuc Where id = IdHinhThucKL And id_goc = 5) As HinhThucKL,"
    '            strSQL += "Convert(Varchar,TuNgay,103) As Tg_TuNgay,Convert(Varchar,DenNgay,103) As Tg_DenNgay, NguoiQD"
    '            strSQL += " FROM HS_KiLuat Where IdCanBo = '" + _RowId + "' Order By TuNgay"
    '            Using db_kiluat As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_kiluat Is Nothing Or db_kiluat.Rows.Count = 0) Then
    '                    _rowsTable = 3
    '                Else
    '                    _rowsTable = db_kiluat.Rows.Count + 2
    '                End If
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 6)
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Số QĐ")
    '                objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
    '                objTable.Cell(1, 3).Range.InsertAfter("Hình thức kỷ luật")
    '                objTable.Cell(1, 4).Range.InsertAfter("Hiệu lực thi hành")
    '                objTable.Cell(2, 4).Range.InsertAfter("Từ ngày")
    '                objTable.Cell(2, 5).Range.InsertAfter("Đến ngày")
    '                objTable.Cell(1, 6).Range.InsertAfter("Người ký QĐ")
    '                objTable.Rows(1).Range.Font.Bold = 1
    '                objTable.Rows(2).Range.Font.Bold = 1
    '                If Not (db_kiluat Is Nothing) Then
    '                    If (db_kiluat.Rows.Count > 0) Then
    '                        For i As Integer = 0 To db_kiluat.Rows.Count - 1
    '                            objTable.Cell(i + 3, 1).Range.InsertAfter(db_kiluat.Rows(i)("SoQD").ToString().Trim())
    '                            objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 3, 2).Range.InsertAfter(db_kiluat.Rows(i)("NgayKyQD").ToString().Trim())
    '                            objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 3).Range.InsertAfter(db_kiluat.Rows(i)("HinhThucKL").ToString().Trim())
    '                            objTable.Cell(i + 3, 3).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                            objTable.Cell(i + 3, 4).Range.InsertAfter(db_kiluat.Rows(i)("Tg_TuNgay").ToString().Trim())
    '                            objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 5).Range.InsertAfter(db_kiluat.Rows(i)("Tg_DenNgay").ToString().Trim())
    '                            objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                            objTable.Cell(i + 3, 6).Range.InsertAfter(db_kiluat.Rows(i)("NguoiQD").ToString().Trim())
    '                            objTable.Cell(i + 3, 6).Range.Paragraphs.LeftIndent = 3.0F
    '                            objTable.Cell(i + 3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                        Next
    '                    End If
    '                End If
    '                objTable.Columns(1).PreferredWidth = 10
    '                objTable.Columns(2).PreferredWidth = 10
    '                objTable.Columns(3).PreferredWidth = 30
    '                objTable.Columns(4).PreferredWidth = 10
    '                objTable.Columns(5).PreferredWidth = 10
    '                objTable.Columns(6).PreferredWidth = 28
    '                objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
    '                objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
    '                objTable.Cell(1, 3).Merge(objTable.Cell(2, 3))
    '                objTable.Cell(1, 6).Merge(objTable.Cell(2, 6))
    '                objTable.Cell(1, 4).Merge(objTable.Cell(1, 5))
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            '53. Kỷ luật
    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Size = 12
    '            objselection.TypeText("52. Hoàn cảnh kinh tế gia đình" + vbCrLf)
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.TypeText("- Quá trình lương của bản thân:" + vbCrLf)
    '            strSQL = "SELECT (Select BacLuong From BacLuong Where IdBacLuong = HS_LuongCB.IdBacLuong) As BacLuong,"
    '            strSQL += "(Select NgachLuong + ' ('+Mota+')' From NgachLuong Where IdNgachLuong in (Select IdNgachLuong From BacLuong Where IdBacLuong = HS_LuongCB.IdBacLuong)) As NgachLuong,"
    '            strSQL += "CONVERT(VarChar,HeSoLuong) As HeSo,Convert(Varchar,Ngay_Huong,103) As NgayHuong,SoQD,Convert(Varchar,NgayQD,103) As NgayQD,NguoiQD"
    '            strSQL += " FROM HS_LuongCB WHERE IdCanBo = '" + _RowId + "' Order By HeSoLuong Asc"
    '            Using db_qtluong As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If (db_qtluong Is Nothing Or db_qtluong.Rows.Count = 0) Then
    '                    _rowsTable = 10
    '                Else
    '                    _rowsTable = 6
    '                End If
    '                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 4, _rowsTable)
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '                objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
    '                objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
    '                objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
    '                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '                objTable.Range.Paragraphs.SpaceBefore = 2
    '                objTable.Range.Paragraphs.SpaceAfter = 2
    '                For i As Integer = 1 To 4
    '                    objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
    '                Next
    '                objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                'Add tiêu đề của các cột trong Bảng
    '                objTable.Cell(1, 1).Range.InsertAfter("Tháng/Năm")
    '                objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
    '                objTable.Cell(2, 1).Range.InsertAfter("Ngạch")
    '                objTable.Cell(3, 1).Range.InsertAfter("Bậc lương")
    '                objTable.Cell(4, 1).Range.InsertAfter("Hệ số lương")
    '                If Not (db_qtluong Is Nothing) Then
    '                    If (db_qtluong.Rows.Count > 0) Then
    '                        Dim _columns As Byte = 6  'Dùng biến này làm biến đếm ngược
    '                        For i As Integer = db_qtluong.Rows.Count - 1 To 0 Step -1
    '                            objTable.Cell(1, _columns).Range.InsertAfter(db_qtluong.Rows(i)("NgayHuong").ToString().Trim().Substring(3))
    '                            objTable.Cell(2, _columns).Range.InsertAfter(db_qtluong.Rows(i)("NgachLuong").ToString().Trim())
    '                            objTable.Cell(3, _columns).Range.InsertAfter(db_qtluong.Rows(i)("BacLuong").ToString().Trim())
    '                            objTable.Cell(4, _columns).Range.InsertAfter(db_qtluong.Rows(i)("HeSo").ToString().Trim())
    '                            _columns -= 1
    '                            If _columns = 1 Then
    '                                Exit For
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                objTable.LeftPadding = -3
    '                objTable.RightPadding = -4
    '            End Using

    '            objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
    '            objselection.Paragraphs.SpaceBefore = 10
    '            objselection.Font.Size = 12
    '            objselection.TypeText("- Nguồn thu nhập chính của gia đình (hàng năm):" + vbCrLf)
    '            objselection.Paragraphs.SpaceBefore = 3
    '            objselection.Paragraphs.SpaceAfter = 3
    '            objselection.TypeText(vbTab + vbTab + "+ Lương: ..............................................................................................................." + vbCrLf)
    '            objselection.TypeText(vbTab + vbTab + "+ Các nguồn khác: ........................................................................................................." + vbCrLf)
    '            objselection.TypeText("- Nhà ở:" + vbTab + "+ Được cấp, được thuê, loại nhà: ....................," + " tổng diện tích sử dụng: .............." + "m2" + vbCrLf)
    '            objselection.TypeText(vbTab + vbTab + "+ Nhà tự mua, tự xây, loại nhà: .....................," + " tổng diện tích sử dụng: ................" + "m2" + vbCrLf)
    '            objselection.TypeText("- Đất ở:" + vbTab + "+ Đất được cấp: ................m2," + vbTab + vbTab + vbTab + "+ Đất được mua: ....................." + "m2" + vbCrLf)
    '            objselection.TypeText("- Đất sản xuất, kinh doanh (tổng diện tích được cấp, tự mua, tự khai phá,...): ......................................" + vbCrLf)
    '            objselection.TypeText("........................................................................................................................................................" + vbCrLf + vbCrLf + vbCrLf + vbCrLf)

    '            objselection.Font.Size = 13
    '            objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 5, 2)
    '            objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '            objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '            objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
    '            objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
    '            objTable.PreferredWidth = 100
    '            objTable.Range.Paragraphs.SpaceBefore = 0
    '            objTable.Range.Paragraphs.SpaceAfter = 0

    '            'objTable.Range.Font.Bold = 0
    '            objTable.Cell(1, 1).Range.InsertAfter("Người khai")
    '            objTable.Cell(1, 1).Range.Font.Bold = 1
    '            objTable.Cell(2, 1).Range.InsertAfter("Tôi xin cam đoan những")
    '            objTable.Cell(3, 1).Range.InsertAfter("lời khai trên đây là đúng sự thật")
    '            objTable.Cell(4, 1).Range.InsertAfter("(Ký tên)")
    '            objTable.Cell(5, 1).Range.InsertAfter(vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + dr("HoTen").ToString().Trim().ToUpper())
    '            objTable.Cell(5, 1).Range.Font.Bold = 1
    '            _sVal = _sDiaban + ", Ngày " + DateTime.Now.ToString("dd") + " tháng " + DateTime.Now.ToString("MM") + " năm " + DateTime.Now.ToString("yyyy")
    '            objTable.Cell(2, 2).Range.InsertAfter(_sVal)
    '            objTable.Cell(2, 2).Range.Font.Italic = 1
    '            objTable.Cell(3, 2).Range.InsertAfter("Xác nhận của cơ quan quản lý")
    '            objTable.Cell(3, 2).Range.Font.Bold = 1
    '            objTable.Cell(5, 2).Range.InsertAfter(vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf)
    '            objTable.Cell(1, 1).Merge(objTable.Cell(4, 1))
    '            objTable.Cell(1, 2).Merge(objTable.Cell(4, 2))
    '            objTable.LeftPadding = -3
    '            objTable.RightPadding = -4


    '            'Định dạng thêm về trang word khi xuất ra
    '            objWordApp.ActiveDocument.PageSetup.HeaderDistance = 3
    '            objWordApp.ActiveDocument.PageSetup.FooterDistance = 3
    '            objWordApp.ActiveDocument.PageSetup.LeftMargin = 25         'Quy ước:   72 Points = 1 Inch.
    '            objWordApp.ActiveDocument.PageSetup.RightMargin = 20
    '            objWordApp.ActiveDocument.PageSetup.TopMargin = 25
    '            objWordApp.ActiveDocument.PageSetup.BottomMargin = 25
    '            objWordApp.ActiveDocument.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4
    '            'objWord.ActiveDocument.PageSetup.HeaderDistance = 100
    '            objWordApp.ActiveDocument.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait   'Định dạng Khổ dấy dọc
    '            objWordApp.ActiveDocument.ShowGrammaticalErrors = False     'Không cho hiển thị các đường viền check chính tả mầu xanh
    '            objWordApp.ActiveDocument.ShowSpellingErrors = False        'Không cho hiển thị các đường viền check chính tả mầu đỏ
    '            'Kiểm tra chưa có thư mục lưu lại file CV đã xuất thì thực hiện tạo Thư mục này
    '            'If Not Globals.FileDao.IsDirectory(Application.StartupPath & "\_LLCN\") Then
    '            '    System.IO.Directory.CreateDirectory(Application.StartupPath & "\_LLCN\")
    '            'End If
    '            ''7. Xuất nội dung ra file word (kết thúc quá trình lưu tập tin và mở msword lên)
    '            'objDocument.SaveAs(Application.StartupPath & "\_LLCN\" + getFullName(_RowId.ToString().Trim(), False) + ".doc")

    '            If Not Globals.FileDao.IsDirectory(gBackUpFileDir & "\LLCN\") Then
    '                System.IO.Directory.CreateDirectory(gBackUpFileDir & "\LLCN\")
    '            End If
    '            '7. Xuất nội dung ra file word (kết thúc quá trình lưu tập tin và mở msword lên)
    '            objDocument.SaveAs(gBackUpFileDir & "\LLCN\" + getFullName(_RowId.ToString().Trim(), False) + ".doc")
    '            objWordApp = Nothing
    '            Clipboard.Clear()
    '        End If
    '    End If
    'End Sub

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

    Private Sub bntImportExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImportExcel.Click
        If DONVI = gMaDonViTW Then
            Dim frmImportCBInf As QLNS.frmImportExcel_CanBoInf = New frmImportExcel_CanBoInf
            frmImportCBInf.ShowDialog()
        Else
            Dim frmImportCB As QLNS.frmImportExcelCB = New frmImportExcelCB
            frmImportCB.ShowDialog()
        End If
    End Sub

    Private Sub bntConvertMaCB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntConvertMaCB.Click
        Dim frmConvertMaCB As QLNS.frmChuyenDoiMaCB = New frmChuyenDoiMaCB
        frmConvertMaCB.ShowDialog()
    End Sub

    Private Sub bntImportHSCB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImportHSCB.Click
        Dim frmImportHSCB As QLNS.frmImportHSCB = New frmImportHSCB
        frmImportHSCB.ShowDialog()
    End Sub

    Private Sub mnu_export_2C_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_export_2C.Click
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                Me.Cursor = Cursors.WaitCursor
                _HS_CanBo.ExportWord_2C(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString())
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Sub chkHienThi_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkHienThi.CheckedChanged
        dgv_main.Columns("cln_Username").Visible = chkHienThi.Checked
        dgv_main.Columns("cln_TWQuanLy").Visible = chkHienThi.Checked
        dgv_main.Columns("cln_Quyen").Visible = chkHienThi.Checked
    End Sub

    Private Sub btnQuyenSD_Click(sender As System.Object, e As System.EventArgs) Handles btn_QuyenSD.Click
        If (dgv_main.Rows.Count > 0) Then
            'Lấy ID Đơn vị
            Dim IdGoc_DVCurr As Integer = 0, MaSo_DVCurr As String = ""
            IdGoc_DVCurr = _SqlHelper.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & _BranchTag.Split("_")(1))
            If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = _BranchTag.Split("_")(1)
            MaSo_DVCurr = _SqlHelper.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)

            Dim frm As New frmHT_QuyenSuDung
            frm.HoTen = dgv_main.CurrentRow.Cells("cln_HoTen").Value
            frm.Username = dgv_main.CurrentRow.Cells("cln_Username").Value
            frm.IdCanBo = dgv_main.CurrentRow.Cells("cln_Id").Value
            If dgv_main.CurrentRow.Cells("cln_Username").Value.ToString.Trim = "" Then frm.MaPOS = MaSo_DVCurr 'Trường hợp thêm mới user thì lấy POS hiện tại

            frm.ShowDialog()
            tv_main_AfterSelect(sender, Nothing)
        Else
            MessageBox.Show("Bạn chưa chọn dữ liệu nhân sự cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btnCapQL_Click(sender As System.Object, e As System.EventArgs) Handles btn_CapQL.Click
        If (dgv_main.Rows.Count > 0) Then
            Dim IdGoc_DVCurr As Integer = 0, MaSo_DVCurr As Integer = 0
            IdGoc_DVCurr = _SqlHelper.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & _BranchTag.Split("_")(1))
            If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = _BranchTag.Split("_")(1)
            MaSo_DVCurr = _SqlHelper.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)

            'Nếu là hồ sơ của Hội sở chính thì luôn thuộc TW quản lý
            If MaSo_DVCurr = gMaDonViTW Then
                My_MessageBox("Cấp quản lý của cán bộ Hội sở chính luôn là Trung ương.")
            Else
                Select Case dgv_main.CurrentRow.Cells("cln_TWQuanLy").Value.ToString
                    Case ""
                        'Nếu cấp quản lý đang là đơn vị
                        If My_MessageBox("Chắc chắn chuyển hồ sơ sang cấp Trung ương quản lý?") = vbYes Then
                            _SqlHelper.executeSQL("UPDATE HS_CanBo SET CapQuanLy = 1 WHERE IdCanBo = '" & dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() & "'")
                            dgv_main.CurrentRow.Cells("cln_TWQuanLy").Value = "X"
                        End If
                    Case "X"
                        If My_MessageBox("Chắc chắn chuyển hồ sơ sang cấp Đơn vị quản lý?") = vbYes Then
                            _SqlHelper.executeSQL("UPDATE HS_CanBo SET CapQuanLy = 0 WHERE idCanBo = '" & dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() & "'")
                            dgv_main.CurrentRow.Cells("cln_TWQuanLy").Value = ""
                            'dgv_main.CurrentRow.Cells("cln_TWQuanLy").Value = _SqlHelper.getNumber("SELECT CapQUanLy FROM HS_CanBo WHERE IdCanBo = '" & dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() & "'")
                        End If
                End Select
            End If
        Else
            MessageBox.Show("Bạn chưa chọn dữ liệu nhân sự cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub
End Class



