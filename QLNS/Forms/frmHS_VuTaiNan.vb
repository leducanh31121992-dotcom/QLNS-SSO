Public Class frmHS_VuTaiNan

#Region "---> The profile of the accident <---"
    Private _HS_CsLaodong As clsHS_CsLaodong = New clsHS_CsLaodong()
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess
    Private arr_LoaiTn As ArrayList = New ArrayList
    Private arr_NguyenNhan As ArrayList = New ArrayList
    'Biến lưu Chỉ số xác định sự kiện người dùng. Với 1-Thêm mới  2-Sửa đổi

    Private _CodeId As String = ""
    Private _ParentNode As String = ""
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

    'Khai báo thuộc tính lấy và trả giá trị Index đang Select bên gọi đến
    Private _Index As Int32
    Public Property Index() As Int32
        Get
            Return _Index
        End Get
        Set(ByVal value As Int32)
            _Index = value
        End Set
    End Property
#End Region

#Region "---> Functions: Các hàm chính <---"
    ''' <summary>
    ''' Hàm thực hiện Reset lại các controls về trạng thái ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls()
        _CodeId = ""
        If _ParentNode <> "" And arr_LoaiTn.Count <> 0 Then
            Dim _index As Integer = CType(arr_LoaiTn.IndexOf(_ParentNode.Substring(5).ToString()), Int32)
            If (_index >= 0 And cb_loai_tn.Items.Count >= _index) Then
                cb_loai_tn.SelectedIndex = _index
            End If
        Else
            cb_loai_tn.SelectedIndex = 0
        End If
        edt_tenvu_tn.Text = ""
        dtpk_ngaytn.Text = DateTime.Now.ToShortDateString()
        cb_vitri.SelectedIndex = 0
        edt_noi_tn.Text = ""
        cb_nguyennhan.SelectedIndex = 0
        edt_sotien_thiethai.Text = "0"
        dtpk_ngay_bc.Text = DateTime.Now.ToShortDateString()
        dtpk_ngay_dtra.Text = DateTime.Now.ToShortDateString()
        edt_nguoi_dtra.Text = ""
        edt_mota_vutn.Text = ""
        edt_denghi.Text = ""
        edt_ghichu.Text = ""
        ckb_chonca.Checked = False
        dgv_main.Rows.Clear()
    End Sub

    Private Function Valid() As Boolean
        If (cb_loai_tn.SelectedIndex <= 0 And cb_loai_tn.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn loại tai nạn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_loai_tn
            Return False
        End If
        If (edt_tenvu_tn.Text.Trim() = "") Then
            MessageBox.Show("Tên vụ tai nạn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_tenvu_tn
            Return False
        End If
        If (cb_vitri.SelectedIndex <= 0 And cb_vitri.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn vị trí xẩy ra tai nạn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_vitri
            Return False
        End If
        If (edt_noi_tn.Text.Trim() = "") Then
            MessageBox.Show("Nơi tai nạn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_noi_tn
            Return False
        End If
        If (cb_nguyennhan.SelectedIndex <= 0 And cb_nguyennhan.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn nguyên nhân xẩy ra tai nạn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_nguyennhan
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện select row trên lưới dữ liệu --> Fill dữ liệu vào các controls
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi (Id record) đang select</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String)
        Dim dr As DataRow
        _CodeId = _Idnew
        dr = _HS_CsLaodong.GetRecord_VuTaiNan(_CodeId)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                cb_loai_tn.SelectedIndex = CType(arr_LoaiTn.IndexOf(dr("IdLoaiTN").ToString()), Int32)
                edt_tenvu_tn.Text = dr("TenVu_TN").ToString().Trim()
                If dr("Ngay_TN").ToString().Trim() <> "" Then
                    dtpk_ngaytn.Value = CType(dr("Ngay_TN").ToString(), DateTime)
                End If
                cb_vitri.SelectedIndex = 0
                If (dr("NgoaiDonVi").ToString().Trim() = "False") Then
                    cb_vitri.SelectedIndex = 1
                Else
                    cb_vitri.SelectedIndex = 2
                End If
                edt_noi_tn.Text = dr("Noi_TN").ToString().Trim()
                cb_nguyennhan.SelectedIndex = CType(arr_NguyenNhan.IndexOf(dr("IdNguyenNhanTN").ToString()), Int32)
                edt_sotien_thiethai.Text = dr("SoTien_ThietHai").ToString().Trim()

                If dr("Ngay_BaoCao").ToString().Trim() <> "" Then
                    dtpk_ngay_bc.Value = CType(dr("Ngay_BaoCao").ToString(), DateTime)
                End If

                If dr("Ngay_DieuTra").ToString().Trim() <> "" Then
                    dtpk_ngay_bc.Value = CType(dr("Ngay_DieuTra").ToString(), DateTime)
                End If
                edt_nguoi_dtra.Text = dr("Nguoi_DieuTra").ToString().Trim()
                edt_mota_vutn.Text = dr("MoTa_TN").ToString()
                edt_denghi.Text = dr("DeNghi").ToString()
                edt_ghichu.Text = dr("GhiChu").ToString()
            End If
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính<---"
    Private Sub frmHS_HSVuTaiNan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _ParentNode = ""
        _HS_CsLaodong.BindData_TreeView(tv_main)
        arr_LoaiTn.Clear()
        cb_loai_tn.Items.Clear()
        arr_LoaiTn = _Globals.Bind_ComBoBox(cb_loai_tn, clsHT_DanhMuc.Sql_LoaiTn, "---Chọn loại tai nạn---")

        arr_NguyenNhan.Clear()
        cb_nguyennhan.Items.Clear()
        arr_NguyenNhan = _Globals.Bind_ComBoBox(cb_nguyennhan, clsHT_DanhMuc.Sql_NgNhanTn, "---Chọn nguyên nhân tai nạn---")

        _HS_CsLaodong.Create_Frame(dgv_main, 0)
        ResetAll_Controls()
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        ResetAll_Controls()
        _ParentNode = ""
        If (tv_main.Nodes.Count > 0) Then
            'Select parent node in treeview
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00") Then
                _ParentNode = tv_main.SelectedNode.Tag.ToString().Trim()
                'Lấy Id loại tai nạn-để fill data hồ sơ vụ tai nạn theo id này
                Using db As DataTable = _HS_CsLaodong.GetAll_HS_VuTaiNan(CType(_ParentNode.Substring(5), Int32))
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_main.Rows.Add()
                                dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                dgv_main.Rows(i).Cells("cln_Code").Value = db.Rows(i)("IdVuTaiNan").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_Tengoi").Value = db.Rows(i)("TenVu_TN").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_LoaiTN").Value = db.Rows(i)("LoaiTN").ToString().Trim()

                                If db.Rows(i)("Ngay_TN").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_NgayTN").Value = CType(db.Rows(i)("Ngay_TN").ToString(), DateTime).ToString("dd-MM-yyyy")
                                Else
                                    dgv_main.Rows(i).Cells("cln_NgayTN").Value = ""
                                End If

                                If (db.Rows(i)("NgoaiDonVi").ToString().Trim() = "False") Then
                                    dgv_main.Rows(i).Cells("cln_ViTri").Value = "Trong đơn vị"
                                ElseIf (db.Rows(i)("NgoaiDonVi").ToString().Trim() = "True") Then
                                    dgv_main.Rows(i).Cells("cln_ViTri").Value = "Ngoài đơn vị"
                                Else
                                    dgv_main.Rows(i).Cells("cln_ViTri").Value = ""
                                End If

                                dgv_main.Rows(i).Cells("cln_NoiTN").Value = db.Rows(i)("Noi_TN").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_NgNhan").Value = db.Rows(i)("NguyenNhan").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_SoTien").Value = IIf(db.Rows(i)("SoTien_ThietHai").ToString().Trim() <> "", Double.Parse(db.Rows(i)("SoTien_ThietHai").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0")

                                If db.Rows(i)("Ngay_BaoCao").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_Ngay_BC").Value = CType(db.Rows(i)("Ngay_BaoCao").ToString(), DateTime).ToString("dd-MM-yyyy")
                                Else
                                    dgv_main.Rows(i).Cells("cln_Ngay_BC").Value = ""
                                End If

                                If db.Rows(i)("Ngay_DieuTra").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_Ngay_DT").Value = CType(db.Rows(i)("Ngay_DieuTra").ToString(), DateTime).ToString("dd-MM-yyyy")
                                Else
                                    dgv_main.Rows(i).Cells("cln_Ngay_DT").Value = ""
                                End If

                                dgv_main.Rows(i).Cells("cln_NguoiDT").Value = db.Rows(i)("Nguoi_DieuTra").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_MoTa").Value = db.Rows(i)("MoTa_TN").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_DeNghi").Value = db.Rows(i)("DeNghi").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_GhiChu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                            Next
                        End If
                    End If
                End Using
            End If

            'Select child node in treeview
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "01") Then
                _ParentNode = tv_main.SelectedNode.Parent.Tag.ToString()
                Dim Code As String = tv_main.SelectedNode.Tag.ToString().Substring(5)
                Dim dr As DataRow
                dr = _HS_CsLaodong.GetRecord_VuTaiNan(Code)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        dgv_main.Rows.Add()
                        dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                        dgv_main.Rows(0).Cells("cln_Code").Value = dr("IdVuTaiNan").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_Tengoi").Value = dr("TenVu_TN").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_LoaiTN").Value = dr("LoaiTN").ToString().Trim()

                        If dr("Ngay_TN").ToString().Trim() <> "" Then
                            dgv_main.Rows(0).Cells("cln_NgayTN").Value = CType(dr("Ngay_TN").ToString(), DateTime).ToString("dd-MM-yyyy")
                        Else
                            dgv_main.Rows(0).Cells("cln_NgayTN").Value = ""
                        End If

                        If (dr("NgoaiDonVi").ToString().Trim() = "1") Then
                            dgv_main.Rows(0).Cells("cln_ViTri").Value = "Trong đơn vị"
                        ElseIf (dr("NgoaiDonVi").ToString().Trim() = "2") Then
                            dgv_main.Rows(0).Cells("cln_ViTri").Value = "Ngoài đơn vị"
                        Else
                            dgv_main.Rows(0).Cells("cln_ViTri").Value = ""
                        End If

                        dgv_main.Rows(0).Cells("cln_NoiTN").Value = dr("Noi_TN").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_NgNhan").Value = dr("NguyenNhan").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_SoTien").Value = IIf(dr("SoTien_ThietHai").ToString().Trim() <> "", Double.Parse(dr("SoTien_ThietHai").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0")

                        If dr("Ngay_BaoCao").ToString().Trim() <> "" Then
                            dgv_main.Rows(0).Cells("cln_Ngay_BC").Value = CType(dr("Ngay_BaoCao").ToString(), DateTime).ToString("dd-MM-yyyy")
                        Else
                            dgv_main.Rows(0).Cells("cln_Ngay_BC").Value = ""
                        End If

                        If dr("Ngay_DieuTra").ToString().Trim() <> "" Then
                            dgv_main.Rows(0).Cells("cln_Ngay_DT").Value = CType(dr("Ngay_DieuTra").ToString(), DateTime).ToString("dd-MM-yyyy")
                        Else
                            dgv_main.Rows(0).Cells("cln_Ngay_DT").Value = ""
                        End If

                        dgv_main.Rows(0).Cells("cln_NguoiDT").Value = dr("Nguoi_DieuTra").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_MoTa").Value = dr("MoTa_TN").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_DeNghi").Value = dr("DeNghi").ToString().Trim()
                        dgv_main.Rows(0).Cells("cln_GhiChu").Value = dr("GhiChu").ToString().Trim()
                    End If
                End If
            End If
            dgv_main_CellClick(sender, Nothing)
        End If
    End Sub

    Private Sub btn_addnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_addnew.Click
        _CodeId = ""
        If _ParentNode <> "" And arr_LoaiTn.Count <> 0 Then
            Dim _index As Integer = CType(arr_LoaiTn.IndexOf(_ParentNode.Substring(5).ToString()), Int32)
            If (_index >= 0 And cb_loai_tn.Items.Count >= _index) Then
                cb_loai_tn.SelectedIndex = _index
            End If
        Else
            cb_loai_tn.SelectedIndex = 0
        End If
        edt_tenvu_tn.Text = ""
        dtpk_ngaytn.Text = DateTime.Now.ToShortDateString()
        cb_vitri.SelectedIndex = 0
        edt_noi_tn.Text = ""
        cb_nguyennhan.SelectedIndex = 0
        edt_sotien_thiethai.Text = ""
        dtpk_ngay_bc.Text = DateTime.Now.ToShortDateString()
        dtpk_ngay_dtra.Text = DateTime.Now.ToShortDateString()
        edt_nguoi_dtra.Text = ""
        edt_mota_vutn.Text = ""
        edt_denghi.Text = ""
        edt_ghichu.Text = ""
        ckb_chonca.Checked = False
        ActiveControl = cb_loai_tn
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        If (Valid()) Then
            Dim obj_hs_vutainan As clsHS_CsLaodong.HS_VuTaiNan = New clsHS_CsLaodong.HS_VuTaiNan()
            obj_hs_vutainan.IdLoaiTN = CType(IIf(arr_LoaiTn.Count > 0, arr_LoaiTn(cb_loai_tn.SelectedIndex), "0"), Int32)
            obj_hs_vutainan.TenVu_TN = Globals.Find_Replace(edt_tenvu_tn.Text.ToString().Trim())
            obj_hs_vutainan.Ngay_TN = dtpk_ngaytn.Value

            If (cb_vitri.SelectedIndex = 1) Then
                obj_hs_vutainan.NgoaiDonVi = 0
            ElseIf (cb_vitri.SelectedIndex = 2) Then
                obj_hs_vutainan.NgoaiDonVi = 1
            End If
            obj_hs_vutainan.Noi_TN = Globals.Find_Replace(edt_noi_tn.Text.ToString().Trim())
            obj_hs_vutainan.IdNguyenNhanTN = CType(IIf(arr_NguyenNhan.Count > 0, arr_NguyenNhan(cb_nguyennhan.SelectedIndex), "0"), Int32)
            obj_hs_vutainan.SoTien_ThietHai = IIf(edt_sotien_thiethai.Text.Trim() <> "", CType(MoneyValue(edt_sotien_thiethai.Text.Trim()), Double), 0)
            obj_hs_vutainan.Ngay_BaoCao = dtpk_ngay_bc.Value
            obj_hs_vutainan.Ngay_DieuTra = dtpk_ngay_dtra.Value
            obj_hs_vutainan.Nguoi_DieuTra = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_nguoi_dtra.Text.ToString().Trim()))
            obj_hs_vutainan.MoTa_TN = Globals.Find_Replace(edt_mota_vutn.Text.ToString().Trim())
            obj_hs_vutainan.DeNghi = Globals.Find_Replace(edt_denghi.Text.ToString().Trim())
            obj_hs_vutainan.GhiChu = Globals.Find_Replace(edt_ghichu.Text.ToString().Trim())
            Dim _currRow As String = ""
            If (_CodeId = "") Then
                _currRow = _HS_CsLaodong.Insert_HS_VuTaiNan(obj_hs_vutainan)
            Else
                obj_hs_vutainan.IdVuTaiNan = _CodeId
                _HS_CsLaodong.Update_HS_VuTaiNan(obj_hs_vutainan)
                _currRow = _CodeId
            End If

            'Load lại các thông tin sau khi cập nhật song dữ liệu
            Dim strNode As String = "01DM_" & _currRow
            _HS_CsLaodong.BindData_TreeView(tv_main)
            ResetAll_Controls()

            'Select item của cây dữ liệu
            Dim node As TreeNode = Nothing
            If (strNode <> "") Then
                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                If Not (node Is Nothing) Then
                    tv_main.SelectedNode = node
                End If
            End If

            'Tìm lại dòng đang select trước đó
            dgv_main.CurrentRow.Selected = False
            dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
            'Select row trên lưới dữ liệu
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(_currRow)
                End If
            End If
        End If
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        Try
            If (dgv_main.Rows.Count <= 0) Then
                Return
            End If
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _count As Int16 = 0
            Dim FlagOK As Boolean = False
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    For i As Int32 = 0 To dgv_main.Rows.Count - 1
                        If (dgv_main.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                            If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                _count += 1
                                Dim strSQL As String = ""
                                strSQL = String.Format("Select * From CB_TaiNan Where IdVuTN = '{0}'", dgv_main.Rows(i).Cells("cln_Code").Value.ToString())
                                _SqlHelper = New DBAccess()
                                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                    If (db Is Nothing Or db.Rows.Count = 0) Then
                                        arr_Del.Add(dgv_main.Rows(i).Cells("cln_Code").Value.ToString())
                                    Else
                                        FlagOK = True
                                    End If
                                End Using
                            End If
                        End If
                    Next
                End If
            End If
            If (_count > 0) Then
                Dim _mess As String = IIf(_count = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ vụ tai nạn đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    If (FlagOK = True) Then
                        MessageBox.Show("Hồ sơ vụ tai nạn bạn chọn hiện đang được dùng, không thể xoá được!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        arr_Del.Clear()
                        Globals.Check_All_Items(dgv_main, False)
                        Dim strNode As String = _ParentNode
                        ResetAll_Controls()
                        tv_main.CollapseAll()
                        'Select item của cây dữ liệu
                        Dim _node As TreeNode = Nothing
                        If (strNode <> "") Then
                            _node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                            If Not (_node Is Nothing) Then
                                tv_main.SelectedNode = _node
                            End If
                        End If
                        dgv_main_CellClick(sender, Nothing)
                    Else
                        'Thực hiện xoá dữ liệu khi đã chấp thuận
                        For i As Int16 = 0 To arr_Del.Count - 1
                            _HS_CsLaodong.Delete_HS_VuTaiNan(arr_Del(i).ToString())
                        Next
                        'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                        MessageBox.Show("Bạn đã xoá thành công thông tin hồ sơ vụ tai nạn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Dim strNode As String = _ParentNode
                        _HS_CsLaodong.BindData_TreeView(tv_main)
                        ResetAll_Controls()

                        'Select item của cây dữ liệu
                        Dim _node As TreeNode = Nothing
                        If (strNode <> "") Then
                            _node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                            If Not (_node Is Nothing) Then
                                tv_main.SelectedNode = _node
                            End If
                        End If
                        dgv_main_CellClick(sender, Nothing)
                    End If
                Else
                    arr_Del.Clear()
                    Globals.Check_All_Items(dgv_main, False)
                    ckb_chonca.Checked = False
                End If
            Else
                MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                ckb_chonca.Checked = False
            End If
        Catch ex As Exception
            MessageBox.Show("Cập nhật xoá bỏ hồ sơ vụ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub ckb_chonca_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_chonca.CheckedChanged
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Globals.Check_All_Items(dgv_main, ckb_chonca.Checked)
            Else
                ckb_chonca.Checked = False
                Return
            End If
        Else
            ckb_chonca.Checked = False
            Return
        End If
    End Sub

    Private Sub dgv_main_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellClick
        Try
            _CodeId = ""
            If _ParentNode <> "" And arr_LoaiTn.Count <> 0 Then
                Dim _index As Integer = CType(arr_LoaiTn.IndexOf(_ParentNode.Substring(5).ToString()), Int32)
                If (_index >= 0 And cb_loai_tn.Items.Count >= _index) Then
                    cb_loai_tn.SelectedIndex = _index
                End If
            Else
                cb_loai_tn.SelectedIndex = 0
            End If
            edt_tenvu_tn.Text = ""
            dtpk_ngaytn.Text = DateTime.Now.ToShortDateString()
            cb_vitri.SelectedIndex = 0
            edt_noi_tn.Text = ""
            cb_nguyennhan.SelectedIndex = 0
            edt_sotien_thiethai.Text = "0"
            dtpk_ngay_bc.Text = DateTime.Now.ToShortDateString()
            dtpk_ngay_dtra.Text = DateTime.Now.ToShortDateString()
            edt_nguoi_dtra.Text = ""
            edt_mota_vutn.Text = ""
            edt_denghi.Text = ""
            edt_ghichu.Text = ""
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(dgv_main.CurrentRow.Cells("cln_Code").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ vụ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_main_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub

    Private Sub btn_huybo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_huybo.Click
        Dim _currRow As String = _CodeId
        _CodeId = ""
        If _ParentNode <> "" And arr_LoaiTn.Count <> 0 Then
            Dim _index As Integer = CType(arr_LoaiTn.IndexOf(_ParentNode.Substring(5).ToString()), Int32)
            If (_index >= 0 And cb_loai_tn.Items.Count >= _index) Then
                cb_loai_tn.SelectedIndex = _index
            End If
        Else
            cb_loai_tn.SelectedIndex = 0
        End If
        edt_tenvu_tn.Text = ""
        dtpk_ngaytn.Text = DateTime.Now.ToShortDateString()
        cb_vitri.SelectedIndex = 0
        edt_noi_tn.Text = ""
        cb_nguyennhan.SelectedIndex = 0
        edt_sotien_thiethai.Text = ""
        dtpk_ngay_bc.Text = DateTime.Now.ToShortDateString()
        dtpk_ngay_dtra.Text = DateTime.Now.ToShortDateString()
        edt_nguoi_dtra.Text = ""
        edt_mota_vutn.Text = ""
        edt_denghi.Text = ""
        edt_ghichu.Text = ""
        ckb_chonca.Checked = False

        If (dgv_main.Rows.Count <= 0) Then Return
        If (_currRow = "") Then _currRow = dgv_main.CurrentRow.Cells("cln_Code").Value.ToString()
        dgv_main.CurrentRow.Selected = False
        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                SelectRow(_currRow)
            End If
        End If
    End Sub

    Private Sub frmHS_HSVuTaiNan_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub frmHS_VuTaiNan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

#Region "---> Events: Các sự kiện ngoại lệ - Di chuyển controls <---"
    Private Sub cb_loai_tn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_loai_tn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_tenvu_tn.Focus()
        End If
    End Sub

    Private Sub edt_tenvu_tn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tenvu_tn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_ngaytn.Focus()
        End If
    End Sub

    Private Sub edt_tenvu_tn_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tenvu_tn.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_loai_tn.Focus()
        End If
    End Sub

    Private Sub dtpk_ngaytn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngaytn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_vitri.Focus()
        End If
    End Sub

    Private Sub cb_vitri_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_vitri.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_noi_tn.Focus()
        End If
    End Sub

    Private Sub edt_noi_tn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_noi_tn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            cb_nguyennhan.Focus()
        End If
    End Sub

    Private Sub edt_noi_tn_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_noi_tn.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_vitri.Focus()
        End If
    End Sub

    Private Sub cb_nguyennhan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_nguyennhan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_sotien_thiethai.Focus()
        End If
    End Sub

    Private Sub edt_sotien_thiethai_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sotien_thiethai.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_sotien_thiethai.Text = "") Then
                edt_sotien_thiethai.Text = "0"
            End If
            dtpk_ngay_bc.Focus()
        End If
    End Sub

    Private Sub edt_sotien_thiethai_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_sotien_thiethai.Leave
        If (edt_sotien_thiethai.Text = "") Then
            edt_sotien_thiethai.Text = "0"
        End If
    End Sub

    Private Sub edt_sotien_thiethai_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_sotien_thiethai.TextChanged
        Try
            edt_sotien_thiethai = formatMoneyinTextbox(edt_sotien_thiethai)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub edt_sotien_thiethai_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_sotien_thiethai.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_nguyennhan.Focus()
        End If
    End Sub

    Private Sub dtpk_ngay_bc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_bc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_ngay_dtra.Focus()
        End If
    End Sub

    Private Sub dtpk_ngay_dtra_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_ngay_dtra.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_nguoi_dtra.Focus()
        End If
    End Sub

    Private Sub edt_nguoi_dtra_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nguoi_dtra.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_mota_vutn.Focus()
        End If
    End Sub

    Private Sub edt_nguoi_dtra_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nguoi_dtra.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_ngay_dtra.Focus()
        End If
    End Sub

    Private Sub edt_mota_vutn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_mota_vutn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_denghi.Focus()
        End If
    End Sub

    Private Sub edt_mota_vutn_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_mota_vutn.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_nguoi_dtra.Focus()
        End If
    End Sub

    Private Sub edt_denghi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_denghi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_denghi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_denghi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_mota_vutn.Focus()
        End If
    End Sub

    Private Sub edt_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_denghi.Focus()
        End If
    End Sub
#End Region

End Class