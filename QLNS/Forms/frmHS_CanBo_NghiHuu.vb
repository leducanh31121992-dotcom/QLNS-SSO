Public Class frmHS_CanBo_NghiHuu

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
        dgv_main.Rows.Clear()
        tv_main.CollapseAll()
        _IdNew = ""
    End Sub

    ''' <summary>
    ''' Hàm thực hiện load lại các thông tin sau khi một sự kiện hoàn thành
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad_Infors()
        Try
            Dim vIdHuman As String = obj_capnhat.IdCanBo.ToString().Trim()
            'Trường hợp không liên quan đến cập nhật hoặc thêm mới mà không thêm lại thực hiện click close form cập nhật dữ liệu
            'If (_IsUpdate = 0 Or vIdHuman = "") Then
            '    ckb_ChoiceAll.Checked = False
            '    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
            '    obj_capnhat.Dispose()
            '    Return
            'End If
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
            MessageBox.Show("Thiết lập lại thông tin sau khi cập nhật hồ sơ cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện phân quyền thành viên thao tác với module
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        'Nếu không có quyền nào liên quan đến thông tin cán bộ
        If Not (Globals.IsIntersect(";71;72;73;74;", _roles)) Then
            dgv_main.Columns("cln_Choice").Visible = False
            'ckb_ChoiceAll.Visible = False
            btn_view_detail.Enabled = False
        End If
        'Xem - Thêm - Sửa - Xoá hồ sơ cán bộ
        If (_roles.IndexOf(";71;") < 0) Then
            btn_view_detail.Enabled = False
        End If

        If (_roles.IndexOf(";74;") < 0) Then
            dgv_main.Columns("cln_Choice").Visible = False
            'ckb_ChoiceAll.Visible = False
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHS_CanBo_NghiHuu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _HS_CanBo.Create_Frame(dgv_main, 18)
        'AddHeaderCheckBox()
        ''Gọi một số sự kiện liên quan đến Check all items trên lưới dữ liệu
        'AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        'AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        'AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged
        'AddHandler dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        'AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged

        _HS_CanBo.Fill_Tree_NghiHuu_ChuyenCT(tv_main)
        arrTapsu.Clear()
        Load_Infors()
        _BranchTag = ""
        Check_Permits()
        _IsUpdate = 0
        'If DONVI = gMaDonViTW Then
        '    btn_add.Enabled = False
        'End If
    End Sub

    Private Sub btn_view_detail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_view_detail.Click
        Try
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    If (Globals.Roles.IndexOf(";71;") < 0) Then
                        Return
                    End If
                    Dim obj_detail As frmHS_NghiHuu_ChuyenCT_ChiTiet = New frmHS_NghiHuu_ChuyenCT_ChiTiet()
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

    Private Sub loadGrid(ByVal vIdDonvi As Integer)
        Dim db As DataTable = New DataTable
        Dim sSQL As String = ""
        Dim k As Integer = 0
        sSQL = getSQL_DSCB(vIdDonvi, 0, False, False, False, False, False, False, False, True)
        Dim dbconn As DBAccess = New DBAccess
        Dim MaSo_DVCurr As String = ""
        Dim IdGoc_DVCurr As Integer = 0
        IdGoc_DVCurr = dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & vIdDonvi)
        If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = vIdDonvi
        MaSo_DVCurr = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)
        k = sSQL.IndexOf("F")       'Vị trí bắt đầu chữ From
        sSQL = "SELECT * " & sSQL.Substring(k - 1)
        db = dbconn.SelectDBRows(sSQL)
        If Not (db Is Nothing) Then
            If (db.Rows.Count > 0) Then
                Dim _rows As Int32 = db.Rows.Count - 1
                For i As Integer = 0 To _rows
                    dgv_main.Rows.Add()
                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdCanBo").ToString() <> "", db.Rows(i)("IdCanBo").ToString(), "")
                    dgv_main.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                    dgv_main.Rows(i).Cells("cln_MaCB").Value = IIf(db.Rows(i)("MaCB").ToString() <> "", db.Rows(i)("MaCB").ToString(), "")
                    dgv_main.Rows(i).Cells("cln_HoTen").Value = IIf(db.Rows(i)("HoTen").ToString() <> "", db.Rows(i)("HoTen").ToString(), "")
                    dgv_main.Rows(i).Cells("cln_GioiTinh").Value = IIf(db.Rows(i)("GioiTinh").ToString() = False, "Nam", "Nữ")
                    If db.Rows(i)("NgaySinh").ToString().Trim() <> "" Then
                        dgv_main.Rows(i).Cells("cln_NgaySinh").Value = CType(db.Rows(i)("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy")
                    Else
                        dgv_main.Rows(i).Cells("cln_NgaySinh").Value = ""
                    End If
                    If vIdDonvi = 1 Then
                        If vIdDonvi = CInt(db.Rows(i)("IdDonVi_Moi")) Then
                            dgv_main.Rows(i).Cells("cln_GhiChu").Value = dbconn.getString("SELECT (SELECT ten_goi FROM DanhMuc WHERE [id]=IdLoaiQD) FROM HS_CBThoiviec WHERE IdCanbo='" & db.Rows(i)("IdCanBo").ToString() & "' And IsQD_NHCS = 1")
                        Else
                            dgv_main.Rows(i).Cells("cln_GhiChu").Value = "Chuyển công tác"
                        End If
                    Else
                        dgv_main.Rows(i).Cells("cln_GhiChu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                        'If dbconn.getNumber("SELECT count(*) FROM ChiNhanh WHERE (Id=" & vIdDonvi & " or id_goc=" & vIdDonvi & " ) and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) > 0 Then
                        '    dgv_main.Rows(i).Cells("cln_GhiChu").Value = dbconn.getString("SELECT (SELECT X.Ten_goi From DanhMuc X Where X.Id=IdLoaiQd) From HS_CBThoiviec WHERE IdCanBo='" & db.Rows(i)("IdCanBo").ToString() & "' And IsQD_NHCS = 1 ")
                        'Else
                        '    dgv_main.Rows(i).Cells("cln_GhiChu").Value = "Chuyển công tác"
                        'End If
                    End If

                    dgv_main.Rows(i).Cells("cln_DonVi").Value = dbconn.getString("SELECT ten_goi from ChiNhanh WHERE [id]=" & CInt(db.Rows(i)("IdDonVi_Moi")))
                    dgv_main.Rows(i).Cells("cln_Phong").Value = dbconn.getString("SELECT ten_phong FROM PhongBan WHERE [id]=" & CInt(db.Rows(i)("IdPhong_Moi")))
                    'dgv_main.Rows(i).Cells("cln_DonVi").Value = _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh where id = {0}", CType(db.Rows(i)("IdDonVi").ToString(), Int32)))
                    'If My_CInt(MaSo_DV.Substring(0, 4), -1) <= 1 Then
                    '    If IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi")) Then
                    '        dgv_main.Rows(i).Cells("cln_Phong").Value = dbconn.getString("SELECT ten_phong FROM PhongBan WHERE [id]=" & CInt(db.Rows(i)("IdPhong_Cu")))
                    '    Else
                    '        dgv_main.Rows(i).Cells("cln_Phong").Value = dbconn.getString("SELECT ten_phong FROM PhongBan WHERE [id]=" & CInt(db.Rows(i)("IdPhong_Moi")))
                    '    End If
                    'Else
                    '    If (IdGoc_DVCurr <> dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id_goc=" & IdGoc_DVCurr & " and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) And IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi"))) Then
                    '        dgv_main.Rows(i).Cells("cln_Phong").Value = dbconn.getString("SELECT ten_phong FROM PhongBan WHERE [id]=" & CInt(db.Rows(i)("IdPhong_Cu")))
                    '    Else
                    '        dgv_main.Rows(i).Cells("cln_Phong").Value = dbconn.getString("SELECT ten_phong FROM PhongBan WHERE [id]=" & CInt(db.Rows(i)("IdPhong_Moi")))
                    '    End If
                    'End If

                Next
                _Records = db.Rows.Count
            End If
        End If
        If (db IsNot Nothing) Then db.Clear()
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        _Records = 0
        'ckb_ChoiceAll.Checked = False
        dgv_main.Rows.Clear()
        strNode = ""
        _BranchTag = ""
        If (tv_main.Nodes.Count > 0) Then
            Dim arrElement() As String
            If tv_main.SelectedNode.GetNodeCount(True) = 0 Then
                If Not (tv_main.SelectedNode.IsExpanded) Then
                    arrElement = tv_main.SelectedNode.Tag.ToString().Trim().Split("_")
                    If (arrElement.Length > 0) Then
                        If (arrElement(0) = "DV") Then
                            _HS_CanBo.Fill_Node_NghiHuu_ChuyenCT(tv_main.SelectedNode, arrElement(1), arrElement(2))
                            '_HS_CanBo.Fill_Node(tv_main.SelectedNode, arrElement(1), arrElement(2))
                            'ElseIf arrElement(0) = "PB" Then
                            '    _HS_CanBo.Fill_NodeCanbo(tv_main.SelectedNode, arrElement(1), arrElement(2), True, True)
                        End If
                    End If
                End If
            End If

            'Thực hiện fill dữ liệu danh sách cán bộ ra lưới dữ liệu
            Dim arrId As String()
            arrId = tv_main.SelectedNode.Tag.ToString().Split("_")
            dgv_main.Rows.Clear()
            loadGrid(arrId(1))


            'If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "PB") Then
            '    _BranchTag = tv_main.SelectedNode.Parent.Tag.ToString()
            '    dgv_main.Rows.Clear()
            '    'Lấy Id - Phòng ban từ Tag của treeview
            '    Dim deptId As Int32 = CType(tv_main.SelectedNode.Tag.ToString().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1), Int32)
            '    'Lấy chỉ số xác định chi nhánh - đơn vị
            '    Dim parent_node As String = tv_main.SelectedNode.Parent.Tag.ToString()
            '    If (parent_node.IndexOf("_") > 0) Then
            '        Dim arrId As String() = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
            '        If (arrId.Length > 0) Then
            '            Dim unitId As Int32 = CType(arrId(1).ToString(), Int32)
            '            Dim db As DataTable = New DataTable
            '            Dim sSQL As String = ""
            '            sSQL = getSQL_DSCB(unitId, deptId, False, False, False, False, False, False, False, True)
            '            Dim k As Integer = 0
            '            Dim dbconn As DBAccess = New DBAccess
            '            Dim MaSo_DVCurr As String = ""
            '            Dim IdGoc_DVCurr As Integer = 0
            '            IdGoc_DVCurr = dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & unitId)
            '            If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = unitId
            '            MaSo_DVCurr = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)
            '            k = sSQL.IndexOf("F")
            '            sSQL = "SELECT * " & sSQL.Substring(k - 1)
            '            'db = _HS_CanBo.GetAllHuman(deptId, unitId)
            '            db = dbconn.SelectDBRows(sSQL)
            '            If Not (db Is Nothing) Then
            '                If (db.Rows.Count > 0) Then
            '                    Dim _rows As Int32 = db.Rows.Count - 1
            '                    For i As Integer = 0 To _rows
            '                        dgv_main.Rows.Add()
            '                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdCanBo").ToString() <> "", db.Rows(i)("IdCanBo").ToString(), "")
            '                        dgv_main.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
            '                        dgv_main.Rows(i).Cells("cln_MaCB").Value = IIf(db.Rows(i)("MaCB").ToString() <> "", db.Rows(i)("MaCB").ToString(), "")
            '                        dgv_main.Rows(i).Cells("cln_HoTen").Value = IIf(db.Rows(i)("HoTen").ToString() <> "", db.Rows(i)("HoTen").ToString(), "")
            '                        dgv_main.Rows(i).Cells("cln_GioiTinh").Value = IIf(db.Rows(i)("GioiTinh").ToString() = False, "Nam", "Nữ")
            '                        If db.Rows(i)("NgaySinh").ToString().Trim() <> "" Then
            '                            dgv_main.Rows(i).Cells("cln_NgaySinh").Value = CType(db.Rows(i)("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy")
            '                        Else
            '                            dgv_main.Rows(i).Cells("cln_NgaySinh").Value = ""
            '                        End If
            '                        'List tên đơn vị từ Id đơn vị ra lưới dữ liệu
            '                        dgv_main.Rows(i).Cells("cln_DonVi").Value = _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh where id = {0}", CType(db.Rows(i)("IdDonVi").ToString(), Int32)))
            '                        If My_CInt(MaSo_DV.Substring(0, 4), -1) <= 1 Then
            '                            If IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi")) Then
            '                                dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
            '                            Else
            '                                dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
            '                            End If
            '                        Else
            '                            If (IdGoc_DVCurr <> dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id_goc=" & IdGoc_DVCurr & " and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) And IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi"))) Then
            '                                dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
            '                            Else
            '                                dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
            '                            End If
            '                        End If
            '                        'If My_CInt(DONVI.Substring(0, 4), -1) <= 1 Then
            '                        '    If IdDONVI <> CInt(db.Rows(i)("IdDonVi_Moi")) Then
            '                        '        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
            '                        '    Else
            '                        '        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
            '                        '    End If
            '                        'Else
            '                        '    If (IdDONVI <> dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id_goc=" & IdDONVI & " and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) And IdDONVI <> CInt(db.Rows(i)("IdDonVi_Moi"))) Then
            '                        '        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
            '                        '    Else
            '                        '        dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
            '                        '    End If
            '                        'End If
            '                        dgv_main.Rows(i).Cells("cln_SoCMT").Value = db.Rows(i)("CMT_So").ToString()
            '                        dgv_main.Rows(i).Cells("cln_DiaChi").Value = _HS_CanBo.GetAddress(CType(IIf(db.Rows(i)("IdThT_Tinh").ToString() <> "", db.Rows(i)("IdThT_Tinh").ToString(), "0"), Int32), CType(IIf(db.Rows(i)("IdThT_Huyen").ToString() <> "", db.Rows(i)("IdThT_Huyen").ToString(), "0"), Int32), db.Rows(i)("ThT_Diachi").ToString().Trim())
            '                        'Điện thoại của cán bộ. Đầu tiên fill - di động nếu không có di động -> fill điện thoại cơ quan của cán bộ
            '                        If db.Rows(i)("DienThoai_DD").ToString().Trim() <> "" Then
            '                            dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_DD").ToString().Trim()
            '                        Else
            '                            If db.Rows(i)("DienThoai_CQ").ToString().Trim() <> "" Then
            '                                dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_CQ").ToString().Trim()
            '                            Else : dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_NR").ToString().Trim()
            '                            End If
            '                        End If
            '                        dgv_main.Rows(i).Cells("cln_Email").Value = db.Rows(i)("Emai").ToString().Trim()
            '                        If db.Rows(i)("Ngay_BienChe").ToString().Trim() <> "" Then
            '                            dgv_main.Rows(i).Cells("cln_BienChe").Value = CType(db.Rows(i)("Ngay_BienChe").ToString(), DateTime).ToString("dd-MM-yyyy")
            '                        Else
            '                            dgv_main.Rows(i).Cells("cln_BienChe").Value = ""
            '                        End If
            '                    Next
            '                    _Records = db.Rows.Count
            '                End If
            '            End If
            '            If (db IsNot Nothing) Then db.Clear()
            '        End If
            '    End If
            'End If

            'If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "CB") Then
            '    _BranchTag = tv_main.SelectedNode.Parent.Parent.Tag.ToString()
            '    Dim _IdHuman As String = tv_main.SelectedNode.Tag.ToString().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1)
            '    If (_IdHuman <> "") Then
            '        dgv_main.Rows.Clear()
            '        'Fill bản ghi thông tin nhân sự theo Mã cán bộ truyền vào
            '        Dim db As DataTable = New DataTable
            '        Dim strParent As String() = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
            '        Dim sSQL As String = ""
            '        sSQL = getSQL_DSCB(strParent(1), strParent(2))
            '        Dim k As Integer = 0
            '        Dim dbconn As DBAccess = New DBAccess
            '        Dim MaSo_DVCurr As String = ""
            '        Dim IdGoc_DVCurr As Integer = 0
            '        IdGoc_DVCurr = dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & strParent(1))
            '        If IdGoc_DVCurr = 0 Or IdGoc_DVCurr = 1 Then IdGoc_DVCurr = strParent(1)
            '        MaSo_DVCurr = dbconn.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & IdGoc_DVCurr)
            '        k = sSQL.IndexOf("F")
            '        sSQL = "SELECT * " & sSQL.Substring(k - 1)
            '        k = sSQL.ToUpper.IndexOf("ORDER BY", 1)
            '        sSQL = sSQL.Substring(0, k) & " And t2.IdCanBo='" & _IdHuman & "' "

            '        'Dim dr As DataRow
            '        'dr = _HS_CanBo.GetHuman_ForCode(_IdHuman)
            '        db = dbconn.SelectDBRows(sSQL)
            '        If Not (db Is Nothing) Then
            '            If (db.Rows.Count > 0) Then
            '                Dim _rows As Int32 = db.Rows.Count - 1
            '                For i As Integer = 0 To _rows
            '                    dgv_main.Rows.Add()
            '                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdCanBo").ToString() <> "", db.Rows(i)("IdCanBo").ToString(), "")
            '                    dgv_main.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
            '                    dgv_main.Rows(i).Cells("cln_MaCB").Value = IIf(db.Rows(i)("MaCB").ToString() <> "", db.Rows(i)("MaCB").ToString(), "")
            '                    dgv_main.Rows(i).Cells("cln_HoTen").Value = IIf(db.Rows(i)("HoTen").ToString() <> "", db.Rows(i)("HoTen").ToString(), "")
            '                    dgv_main.Rows(i).Cells("cln_GioiTinh").Value = IIf(db.Rows(i)("GioiTinh").ToString() = False, "Nam", "Nữ")
            '                    If db.Rows(i)("NgaySinh").ToString().Trim() <> "" Then
            '                        dgv_main.Rows(i).Cells("cln_NgaySinh").Value = CType(db.Rows(i)("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy")
            '                    Else
            '                        dgv_main.Rows(i).Cells("cln_NgaySinh").Value = ""
            '                    End If
            '                    'List tên đơn vị từ Id đơn vị ra lưới dữ liệu
            '                    dgv_main.Rows(i).Cells("cln_DonVi").Value = _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh where id = {0}", CType(db.Rows(i)("IdDonVi").ToString(), Int32)))
            '                    If My_CInt(MaSo_DV.Substring(0, 4), -1) <= 1 Then
            '                        If IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi")) Then
            '                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
            '                        Else
            '                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
            '                        End If
            '                    Else
            '                        If (IdGoc_DVCurr <> dbconn.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id_goc=" & IdGoc_DVCurr & " and id=" & CInt(db.Rows(i)("IdDonVi_Moi"))) And IdGoc_DVCurr <> CInt(db.Rows(i)("IdDonVi_Moi"))) Then
            '                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Cu"))
            '                        Else
            '                            dgv_main.Rows(i).Cells("cln_Phong").Value = CInt(db.Rows(i)("IdPhong_Moi"))
            '                        End If
            '                    End If
            '                    dgv_main.Rows(i).Cells("cln_SoCMT").Value = db.Rows(i)("CMT_So").ToString()
            '                    dgv_main.Rows(i).Cells("cln_DiaChi").Value = _HS_CanBo.GetAddress(CType(IIf(db.Rows(i)("IdThT_Tinh").ToString() <> "", db.Rows(i)("IdThT_Tinh").ToString(), "0"), Int32), CType(IIf(db.Rows(i)("IdThT_Huyen").ToString() <> "", db.Rows(i)("IdThT_Huyen").ToString(), "0"), Int32), db.Rows(i)("ThT_Diachi").ToString().Trim())

            '                    'Điện thoại của cán bộ. Đầu tiên fill - di động nếu không có di động -> fill điện thoại cơ quan của cán bộ
            '                    If db.Rows(i)("DienThoai_DD").ToString().Trim() <> "" Then
            '                        dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_DD").ToString().Trim()
            '                    Else
            '                        If db.Rows(i)("DienThoai_CQ").ToString().Trim() <> "" Then
            '                            dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_CQ").ToString().Trim()
            '                        Else : dgv_main.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai_NR").ToString().Trim()
            '                        End If
            '                    End If
            '                    dgv_main.Rows(i).Cells("cln_Email").Value = db.Rows(i)("Emai").ToString().Trim()
            '                    If db.Rows(i)("Ngay_BienChe").ToString().Trim() <> "" Then

            '                        dgv_main.Rows(i).Cells("cln_BienChe").Value = CType(db.Rows(i)("Ngay_BienChe").ToString(), DateTime).ToString("dd-MM-yyyy")
            '                    Else
            '                        dgv_main.Rows(i).Cells("cln_BienChe").Value = ""
            '                    End If
            '                Next
            '                _Records = db.Rows.Count
            '            End If
            '        End If
            '        If (db IsNot Nothing) Then db.Clear()
            '    End If
            'End If
            tv_main.SelectedNode.Expand()
        End If
        'Thực hiện fill danh sách cán bộ tập sự để chuyển thành lao động dài hạn
        '--> Lấy id đơn vị vừa chọn

        Dim strSQL As String = ""
        Dim BranchId As String = ""
        If (_BranchTag <> "") Then
            Dim arrElement() As String = _BranchTag.Split("_")
            If (arrElement.Length > 0) Then
                If (arrElement(2).Substring(3).Trim() <> "00") Then
                    'Thực hiện lấy id chi nhánh tỉnh từ id pgd
                    BranchId = _ListDocument.GetRootId(CType(arrElement(1).ToString().Trim(), Integer))
                Else
                    BranchId = arrElement(1).ToString().Trim()
                End If
            End If
        End If
        TotalCheckBoxes = dgv_main.RowCount
        TotalCheckedCheckBoxes = 0
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Close()
    End Sub

    Private Sub frmHS_CanBo_NghiHuu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub dgv_main_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellClick
        _IdNew = ""
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                Dim dr As DataRow
                dr = _HS_CanBo.GetHuman_ForCode(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString())
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        'Nếu là cán bộ chính thức đã chuyển sang lao động ngắn hạn
                        If (dr("IdNew").ToString().Trim() <> "") Then
                            _IdNew = dr("IdNew").ToString().Trim()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub dgv_main_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub
#End Region

End Class



