Public Class frmQuyetDinhNS
    Inherits System.Windows.Forms.Form
    Public IDCB_HS_Canbo As String = ""   ' Id Cán bộ truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDDV_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDPB_HS_Canbo As String = 0    ' Id Phong/ban truyền từ menu ngữ cảnh trong HS_Canbo
    Private dbconn As DBAccess
    Private init As Boolean = False
    Private init_L As Boolean = False
    Private init_P As Boolean = False
    Private initTV As Boolean = False
    Private idCanBo As String = ""
    Private idQDNhansu As String = ""
    Private idLuong As String = ""
    Private idPhucap As String = ""
    Private idQDThoiViec As String = ""
    Private initQDL As Boolean = False
    Private initQDP As Boolean = False
    Private idQDLuong As String = ""
    Private idQDPhucap As String = ""
    Private idQDKhac As String = ""
    Private IdBangLuong As Integer = 0

    Private IdDonViCu As Integer = 0
    Private IdPhongCu As Integer = 0
    Private IdChucVuCu As Integer = 0
    Private IdChuyenMonCu As Integer = 0
    Private SoQDCU As String = ""
    Private NgayHLCU_Max As Date
    Private NgayHLCU As Date
    Private NgayQDCU As Date

    Private Sub frmQuyetDinhNS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lkl_xemct.Enabled = False
        'Check quyền thành viên
        Dim roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";99;100;101;102;", roles)) Then
            If tabControlDieuChuyen.TabPages.Contains(tabQDNhanSu) Then tabControlDieuChuyen.TabPages.Remove(tabQDNhanSu)
        End If

        If Not (Globals.IsIntersect(";139;140;141;142;", roles)) Then
            If tabControlDieuChuyen.TabPages.Contains(tabThoiViec) Then tabControlDieuChuyen.TabPages.Remove(tabThoiViec)
        End If

        If Not (Globals.IsIntersect(";187;188;189;190;", roles)) Then
            If tabControlDieuChuyen.TabPages.Contains(tabQDLuong) Then tabControlDieuChuyen.TabPages.Remove(tabQDLuong)
        End If
        If Not (Globals.IsIntersect(";191;192;193;194;", roles)) Then
            If tabControlDieuChuyen.TabPages.Contains(tabQDPhuCap) Then tabControlDieuChuyen.TabPages.Remove(tabQDPhuCap)
        End If
        'QDNhansu
        If (Globals.Roles.IndexOf(";99;") < 0) Then   ' xem
            gridQDNhanSu.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";100;") < 0) Then  ' Them
            bntNew.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";102;") < 0) Then  ' Xoá
            bntDelete.Enabled = False
        End If
        'QD Thoi Viec
        If (Globals.Roles.IndexOf(";139;") < 0) Then
            gridThoiViec.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";140;") < 0) Then
            bntNew_TV.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";142;") < 0) Then
            bntDelete_TV.Enabled = False
        End If
        ' QD Luong
        If (Globals.Roles.IndexOf(";187;") < 0) Then   ' xem
            gridQDLuong.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";188;") < 0) Then  ' Them
            bntNew.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";190;") < 0) Then  ' Xoá
            bntDelete.Enabled = False
        End If
        'QD Phu cap
        If (Globals.Roles.IndexOf(";191;") < 0) Then
            gridQDPhuCap.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";192;") < 0) Then
            bntNew_PC.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";194;") < 0) Then
            bntDelete_PC.Enabled = False
        End If
        
        If IDCB_HS_Canbo = "" Then
            bindTreeview(treeCocau)
            treeCocau.ExpandAll()
        Else
            'Gọi từ Hồ sơ cán bộ qua menu ngữ cảnh
            Dim maCB As String = getCanBo_Ma(IDCB_HS_Canbo, False)
            Dim _NodeFind As TreeNode = New TreeNode
            Dim _node As TreeNode = New TreeNode
            Dim m_QDNS As QDNhanSu = New QDNhanSu

            'Dim m_QDNS As QDNhanSu = New QDNhanSu
            'm_QDNS = m_QDNS.getFinalRecord(IDCB_HS_Canbo, Nothing, True, True)
            '_NodeFind.Name = IDCB_HS_Canbo
            '_NodeFind.Tag = "CB_" & IDCB_HS_Canbo & "_" & maCB & "_" & m_QDNS.IdDonvi_Moi & "_" & m_QDNS.IdPhong_Moi
            'bindTreeview(treeCocau, False, m_QDNS.IdPhong_Moi)

            bindTreeview(treeCocau, False, IDPB_HS_Canbo, IDDV_HS_Canbo)
            treeCocau.ExpandAll()

            m_QDNS = m_QDNS.getFinalRecord(IDCB_HS_Canbo, Nothing, True, True)
            _NodeFind.Name = IDCB_HS_Canbo
            _NodeFind.Tag = "CB_" & IDCB_HS_Canbo & "_" & maCB & "_" & m_QDNS.IdDonvi_Moi & "_" & m_QDNS.IdPhong_Moi

            _node = findNode(treeCocau.Nodes, _NodeFind.Tag)
            If Not _node Is Nothing Then
                treeCocau.SelectedNode = _node
                idCanBo = IDCB_HS_Canbo
                OverInfCB(idCanBo, maCB)
                initTabQDNhansu()
            End If

            If checkRight_CreateRecord(IDCB_HS_Canbo) OrElse (IdDONVI = 1 AndAlso chkTWQuanLy.Checked = True) Then
                bntAdd.Enabled = True
                bntUpdate_TV.Enabled = True
                bntUpdate_L.Enabled = True
                bntUpdate_PC.Enabled = True
                bntUpdate_K.Enabled = True
                bntDelete.Enabled = True
                bntDelete_TV.Enabled = True
                bntDelete_L.Enabled = True
                bntDelete_PC.Enabled = True
                bntDelete_K.Enabled = True
            Else
                bntAdd.Enabled = False
                bntUpdate_TV.Enabled = False
                bntUpdate_L.Enabled = False
                bntUpdate_PC.Enabled = False
                bntUpdate_K.Enabled = False
                bntDelete.Enabled = False
                bntDelete_TV.Enabled = False
                bntDelete_L.Enabled = False
                bntDelete_PC.Enabled = False
                bntDelete_K.Enabled = False
            End If
        End If
    End Sub

    Private Sub frmQuyetDinhNS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub treeCocau_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeCocau.AfterSelect
        'My_MessageBox(treeCocau.SelectedNode.Tag & ".")
        If e.Action <> TreeViewAction.Unknown Then
            Dim arr() As String
            arr = treeCocau.SelectedNode.Tag.ToString.Split("_")
            labAlert_QD.Text = ""
            labAlert_TV.Text = ""
            Select Case arr(0)
                Case "DV"
                    blankOverInfCB()
                    If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                        If Not treeCocau.SelectedNode.IsExpanded Then
                            getNode(treeCocau.SelectedNode, arr(1), arr(2))
                        End If
                    End If
                    treeCocau.SelectedNode.Expand()
                Case "PB"
                    blankOverInfCB()
                    If tabControlDieuChuyen.SelectedTab.Name = "tabQDNhanSu" Then
                        While treeCocau.SelectedNode.Nodes.Count > 0
                            treeCocau.SelectedNode.Nodes.Remove(treeCocau.SelectedNode.FirstNode)
                        End While
                        If treeCocau.SelectedNode.IsSelected Then
                            getNodeCanbo(treeCocau.SelectedNode, arr(1), arr(2))
                        End If
                        gridQDNhanSu.DataSource = Nothing
                        blankQDNhansu()
                    Else
                        If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                            If Not treeCocau.SelectedNode.IsExpanded Then
                                'getNodeCanbo(treeCocau.SelectedNode, arr(1), arr(2), False, True)
                                getNodeCanbo(treeCocau.SelectedNode, arr(1), arr(2))
                            End If
                        End If
                        gridThoiViec.DataSource = Nothing
                        blankQDThoiViec()
                    End If
                    treeCocau.SelectedNode.Expand()
                Case "CB"
                    'Bắt quyền xem chi tiết hồ sơ cán bộ
                    If (Globals.Roles.IndexOf(";71;") < 0) Then
                        lkl_xemct.Enabled = False
                    Else
                        lkl_xemct.Enabled = True
                    End If
                    idCanBo = arr(1)
                    OverInfCB(idCanBo, arr(2))
                    idQDNhansu = ""
                    idLuong = ""
                    idPhucap = ""
                    idQDThoiViec = ""
                    idQDLuong = ""
                    idQDPhucap = ""
                    idQDKhac = ""
                    Select Case tabControlDieuChuyen.SelectedTab.Name
                        Case "tabQDNhanSu"
                            initTabQDNhansu()
                        Case "tabThoiViec"
                            initTabQDThoiViec()
                        Case "tabQDLuong"
                            initTabQDLuong()
                        Case "tabQDPhuCap"
                            initTabQDPhuCap()
                        Case Else
                            initTabQDKhac()
                    End Select

                    If checkRight_CreateRecord(idCanBo) OrElse (IdDONVI = 1 AndAlso chkTWQuanLy.Checked = True) Then
                        bntAdd.Enabled = True
                        bntUpdate_TV.Enabled = True
                        bntUpdate_L.Enabled = True
                        bntUpdate_PC.Enabled = True
                        bntUpdate_K.Enabled = True
                        bntDelete.Enabled = True
                        bntDelete_TV.Enabled = True
                        bntDelete_L.Enabled = True
                        bntDelete_PC.Enabled = True
                        bntDelete_K.Enabled = True
                    Else
                        bntAdd.Enabled = False
                        bntUpdate_TV.Enabled = False
                        bntUpdate_L.Enabled = False
                        bntUpdate_PC.Enabled = False
                        bntUpdate_K.Enabled = False
                        bntDelete.Enabled = False
                        bntDelete_TV.Enabled = False
                        bntDelete_L.Enabled = False
                        bntDelete_PC.Enabled = False
                        bntDelete_K.Enabled = False
                    End If
            End Select
        End If
    End Sub

    Private Sub tabControlDieuChuyen_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabControlDieuChuyen.SelectedIndexChanged
        callTab()
    End Sub

    Private Sub lkl_xemct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkl_xemct.LinkClicked
        If (idCanBo <> "") Then
            Dim obj_detail As frmHS_ChiTiet = New frmHS_ChiTiet()
            'Lấy danh sách mảng các id hiện có trên lưới dl
            Dim arrRows As ArrayList = New ArrayList()
            arrRows.Add(idCanBo)
            obj_detail.RecordCurrent = 0
            obj_detail.Records = 1
            obj_detail.IdCanBo = idCanBo
            obj_detail.arr_RecordId = arrRows
            obj_detail.ShowDialog()
        Else
            Return
        End If
    End Sub

#Region "private function"

    Private Sub OverInfCB(ByVal vIdCanbo As String, ByVal vMaCb As String)
        Dim dt As DataTable
        dbconn = New DBAccess
        dt = dbconn.SelectDBRows("SELECT Hoten, Gioitinh, Ngaysinh, CMT_So, CapQuanLy FROM HS_Canbo WHERE idCanbo='" & vIdCanbo & "' and MaCB='" & vMaCb & "'")
        If dt.Rows.Count > 0 Then
            txtCanbo.Text = dt.Rows(0).Item("Hoten")
            txtMaCB.Text = vMaCb
            txtNgaysinhCB.Text = DateTimeUtil.getShortDate(dt.Rows(0).Item("Ngaysinh"))
            txtCMT.Text = dt.Rows(0).Item("CMT_So")
            If dt.Rows(0).Item("Gioitinh") Then
                txtGioiTinh.Text = "Nữ"
            Else
                txtGioiTinh.Text = "Nam"
            End If
            chkTWQuanLy.Checked = My_CBool(dt.Rows(0).Item("CapQuanLy"), False)
        End If
        lkl_xemct.Enabled = True
    End Sub

    Private Sub callTab()
        Try
            Dim arr() As String
            Dim _node As TreeNode
            _node = treeCocau.SelectedNode.Parent
            arr = _node.Tag.ToString.Split("_")
            If arr(0) <> "PB" Then
                _node = treeCocau.SelectedNode
                arr = _node.Tag.ToString.Split("_")
            End If
            Select Case tabControlDieuChuyen.SelectedTab.Name
                Case "tabQDNhanSu"
                    If arr(2) <> 0 Then refreshNodeNOTCanbo(treeCocau, _node, arr(1), arr(2), False, False, idCanBo)
                    initTabQDNhansu()
                Case "tabThoiViec"
                    'If arr(2) <> 0 Then refreshNodeNOTCanbo(treeCocau, _node, arr(1), arr(2), False, True, idCanBo)
                    initTabQDThoiViec()
                Case "tabQDLuong"
                    initTabQDLuong()
                Case "tabQDPhuCap"
                    initTabQDPhuCap()
                Case Else
                    initTabQDKhac()
            End Select
        Catch ex As Exception
        End Try
    End Sub

    Private Sub blankOverInfCB()
        idCanBo = ""
        idQDNhansu = ""
        idLuong = ""
        idPhucap = ""
        idQDThoiViec = ""
        idQDLuong = ""
        idQDPhucap = ""
        idQDKhac = ""
        txtCanbo.Text = ""
        txtMaCB.Text = ""
        txtNgaysinhCB.Text = ""
        txtGioiTinh.Text = ""
        txtCMT.Text = ""
        lkl_xemct.Enabled = False
        chkTWQuanLy.Checked = False
    End Sub

#End Region

#Region "tab Quyết định nhân sự"

    Private Sub initTabQDNhansu()
        cbIsQD_NHCS.Checked = True
        If Not init Then
            init = True
            bindCboLoaiQD()
            cboCVNguoiQD.DataSource = listChucVuQuyenRaQD()
            bindCboDonVi()
            bindCboChuyenMon()
        End If
        getIdBangNgachBac(idCanBo, IdBangLuong, 0, 0, 0)
        bindGridQDNhansu(idCanBo)
        fillQDNhansu(idQDNhansu, idCanBo)
        gridQDNhanSu.Focus()
    End Sub

    Private Sub bindCboLoaiQD()
        cboLoaiQD.DataSource = listDanhmucQDNhanSu()
    End Sub

    Private Sub bindCboDonVi()
        cboDonviMoi.DataSource = listDonvi_New(False, False, False, False, True)
    End Sub

    Private Sub bindCboPhongMoi(ByVal vIdDonviMoi As Integer)
        Try
            If vIdDonviMoi > 0 Then
                cboPhongMoi.DataSource = listPhong(vIdDonviMoi, False, False, False)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bindCboChucVu(ByVal vIdDonviMoi As Integer)
        cboChucvuMoi.DataSource = listChucVu(vIdDonviMoi)
    End Sub

    Private Sub bindCboChuyenMon()
        cboChuyenmonMoi.DataSource = listDanhmuc(12, False, True)
    End Sub

    Private Sub bindGridQDNhansu(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridQDNhanSu.AutoGenerateColumns = False
                gridQDNhanSu.DataSource = listQDNhansu(vIdCanbo, 1)
                Dim i As Integer = 0
                While i <= gridQDNhanSu.Rows.Count - 1
                    If Trim(gridQDNhanSu.Rows(i).Cells(0).Value) = idQDNhansu Then
                        gridQDNhanSu.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            Else
                gridQDNhanSu.DataSource = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách quyết định nhân sự !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabQDNhanSu() As String
        Dim strReturn As String = ""
        Try
            If cbIsQD_NHCS.Checked = True Then
                If txtSoQD.Text = "" Then
                    txtSoQD.Focus()
                    strReturn = "Chưa nhập Số quyết định !"
                    Exit Try
                Else
                    txtSoQD.Text = txtSoQD.Text.Replace("_", "-")
                End If
            End If
            If dpkNgayHL.Value.Date > Now.Date Then
                strReturn = "Ngày hiệu lực phải nhỏ hơn hoặc bằng ngày hiện tại. Hãy nhập lại!"
                dpkNgayHL.Focus()
                Exit Try
            End If

            If dpkNgayHL.Value.Date > NgayHLCU_Max Then
                If cbIsQD_NHCS.Checked = False Then
                    strReturn = "Ngày hiệu lực của quyết định ngoài NHCS phải nhỏ hơn ngày quyết định hiện tại của NHCS." & vbCr & " Hãy nhập lại!"
                    dpkNgayHL.Focus()
                    Exit Try
                Else
                    If DONVI = gMaDonViTW Then
                        'If dbconn.getNumber("SELECT count(*) FROM chinhanh where (id_goc=0 or id_goc=1) and id=" & CInt(cboDonviMoi.SelectedValue)) <= 0 Then
                        If dbconn.getNumber("SELECT count(Id) FROM ChiNhanh where status=1 and [id]=" & CInt(cboDonviMoi.SelectedValue)) <= 0 Then
                            strReturn = "Với giá trị Ngày hiệu lực lớn nhất, giá trị Đơn vị mới không hợp lệ. Hãy nhập lại!"
                            dpkNgayHL.Focus()
                            Exit Try
                        Else
                            'kiem tra id phong cuaQd mới nhất có status
                            If dbconn.getNumber("SELECT count(*) FROM PhongBan WHERE id=" & CInt(cboPhongMoi.SelectedValue) & " and status=0") > 0 Then
                                strReturn = "Tại thời điểm quyết định có hiệu lực, Phòng/Ban trên không tồn tại. Hãy nhập lại!"
                                dpkNgayHL.Focus()
                                Exit Try
                            End If
                        End If
                    Else
                        If dbconn.getNumber("SELECT count(*) FROM chinhanh WHERE id=" & CInt(cboDonviMoi.SelectedValue) & " or (id_goc in (select id_goc from chinhanh where id=" & CInt(cboDonviMoi.SelectedValue) & ") and id>69)  or (id in (select id_goc from chinhanh where id=" & CInt(cboDonviMoi.SelectedValue) & ") and id_goc<>0)") <= 0 Then
                            strReturn = "Tại thời điểm quyết định có hiệu lực, Đơn vị trên không hợp lệ. Hãy nhập lại!"
                            dpkNgayHL.Focus()
                            Exit Try
                            'kiem tra id phong cuaQd mới nhất có status
                            If dbconn.getNumber("SELECT count(*) FROM PhongBan WHERE id=" & CInt(cboPhongMoi.SelectedValue) & " and status=0") > 0 Then
                                strReturn = "Tại thời điểm quyết định có hiệu lực, Phòng/Ban trên không tồn tại. Hãy nhập lại!"
                                dpkNgayHL.Focus()
                                Exit Try
                            End If
                        End If
                    End If
                End If
            End If

            If DONVI = gMaDonViTW Then
                If dpkNgayHL.Value.Date > NgayHLCU_Max Then
                    'If CInt(cboDonviMoi.SelectedValue) <> 1 Then
                    '    If Not ((dbconn.getString("SELECT ma_so FROM PhongBan WHERE id=" & CInt(cboPhongMoi.SelectedValue)) = "14") Or (dbconn.getString("SELECT ma_so FROM PhongBan WHERE id=" & CInt(cboPhongMoi.SelectedValue)) = "15" And dbconn.getString("SELECT ma_so FROM DanhMuc WHERE id=" & CInt(cboChucvuMoi.SelectedValue)) = "1412") Or (dbconn.getString("SELECT ma_so FROM PhongBan WHERE id=" & CInt(cboPhongMoi.SelectedValue)) = "17" And dbconn.getString("SELECT ma_so FROM DanhMuc WHERE id=" & CInt(cboChucvuMoi.SelectedValue)) = "1412")) Then
                    '        strReturn = "Không nhập được quyết định cho cán bộ không thuộc Hội sở chính. Hãy nhập lại!"
                    '        dpkNgayHL_L.Focus()
                    '        Exit Try
                    '    End If
                    'End If
                    If IdDonViCu <> 1 Then
                        If Not ((dbconn.getString("SELECT ma_so FROM PhongBan WHERE id=" & IdPhongCu) = "14") Or (dbconn.getString("SELECT ma_so FROM PhongBan WHERE id=" & IdPhongCu) = "15" And dbconn.getString("SELECT ma_so FROM DanhMuc WHERE id=" & IdChucVuCu) = "1412") Or (dbconn.getString("SELECT ma_so FROM PhongBan WHERE id=" & IdPhongCu) = "17" And dbconn.getString("SELECT ma_so FROM DanhMuc WHERE id=" & IdChucVuCu) = "1412")) Then
                            strReturn = "Không nhập được quyết định cho cán bộ không thuộc Hội sở chính. Hãy nhập lại!"
                            dpkNgayHL_L.Focus()
                            Exit Try
                        End If
                    End If
                End If
            End If

            'If idQDNhansu = "" Then
            '    If checkQuyetDinh("QDNHANSU", idCanBo, txtSoQD.Text) Then
            '        txtSoQD.Text = ""
            '        txtSoQD.Focus()
            '        strReturn = "Số quyết định nhân sự của cán bộ đã tồn tại. Hãy nhập lại!"
            '        Exit Try
            '    End If
            'End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankQDNhansu()
        Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
        m_QDNhanSu = m_QDNhanSu.getFinalRecord(idCanBo)
        cbIsQD_NHCS.Checked = True
        If m_QDNhanSu.IdQDNhanSu <> "" Then
            IdDonViCu = m_QDNhanSu.IdDonvi_Moi
            IdPhongCu = m_QDNhanSu.IdPhong_Moi
            IdChucVuCu = m_QDNhanSu.IdChucvu_Moi
            IdChuyenMonCu = m_QDNhanSu.IdChuyenMon_Moi
            txtDonViCu.Text = getDonvi(IdDonViCu)
            txtPhongCu.Text = getPhong(IdPhongCu)
            txtChucVuCu.Text = getDanhmuc_Name(14, IdChucVuCu)
            txtChuyenMonCu.Text = getDanhmuc_Name(12, IdChuyenMonCu)
        Else
            txtDonViCu.Text = ""
            txtPhongCu.Text = ""
            txtChucVuCu.Text = ""
            txtChuyenMonCu.Text = ""
        End If
        grpTTCu.Enabled = False
        idQDNhansu = ""
        idLuong = ""
        idPhucap = ""
        bntDelete.Enabled = False
        txtSoQD.Text = ""
        dpkNgayQD.Value = Date.Now
        dpkNgayHL.Value = Date.Now
        txtDVraQD.Text = getDonvi(IdDONVI)
        txtNguoiQD.Text = GIAMDOC
        cboDonviMoi.SelectedValue = IdDONVI
        If CAP = 1 Then
            cboCVNguoiQD.SelectedValue = getDanhmuc_ID("1402")
        Else
            cboCVNguoiQD.SelectedValue = getDanhmuc_ID("1410")
        End If
        cbLuong.Checked = False
        cbPhucap.Checked = False
        dpkNgayTL.Checked = False
        dpkNgayBN.Checked = False
        cbActive.Checked = True
        cboNghiDinh.SelectedValue = 0
        cboBang.SelectedValue = 0
        cboNgach.SelectedValue = 0
        cboBac.SelectedValue = 0
        txtGhichu.Text = ""
        labAlert_QD.Text = ""
        NgayHLCU = Nothing
        SoQDCU = ""
        NgayQDCU = Nothing
        cbIsQD_NHCS.Focus()
    End Sub

    Private Sub InfLuong(ByVal active As Boolean)
        cboNghiDinh.Enabled = active
        cboBang.Enabled = active
        cboNgach.Enabled = active
        cboBac.Enabled = active
        txtHeso.Enabled = active
        txtNoiDung_L.Enabled = active
        If active And (Not init_L) Then
            init_L = True
            cboNghiDinh.DataSource = listNghiDinhLuong(True)
        End If
    End Sub

    Private Sub InfPhucap(ByVal active As Boolean)
        cboLoaiPC.Enabled = active
        cboMucPC.Enabled = active
        txtNoiDung_PC.Enabled = active
        If active And (Not init_P) Then
            init_P = True
            cboLoaiPC.DataSource = listDanhmuc(31, True)
        End If
    End Sub

    Private Sub fillQDNhansu(ByRef vIdQDNhanSu As String, Optional ByVal vIdCanBo As String = "")
        Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
        Dim m_Luong As LuongCanBo = New LuongCanBo
        Dim m_Phucap As Phucap = New Phucap
        Dim vIdNghiDinhLuong, vIdBangluong, vIdNgachluong, vIdLoaiPC As Integer

        labAlert_QD.Text = ""
        If (Globals.Roles.IndexOf(";102;") < 0) Then  ' Xoá
            bntDelete.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDelete.Enabled = True
                Else
                    bntDelete.Enabled = False
                End If
            End If
        End If
        If vIdCanBo <> "" Then
            m_QDNhanSu = m_QDNhanSu.getFinalRecord(vIdCanBo, Nothing, True, True)
            NgayHLCU_Max = m_QDNhanSu.NgayHL
        End If

        If vIdQDNhanSu <> "" Then
            m_QDNhanSu = m_QDNhanSu.getRecord(vIdQDNhanSu)
        Else
            If vIdCanBo <> "" Then m_QDNhanSu = m_QDNhanSu.getFinalRecord(vIdCanBo, Nothing, True, True)
        End If
        vIdQDNhanSu = m_QDNhanSu.IdQDNhanSu
        txtSoQD.Text = m_QDNhanSu.So_QD
        dpkNgayQD.Value = m_QDNhanSu.NgayKy_QD
        dpkNgayHL.Value = m_QDNhanSu.NgayHL
        If m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue Then
            dpkNgayBN.Checked = False
        Else
            dpkNgayBN.Checked = True
            dpkNgayBN.Value = m_QDNhanSu.NgayBoNhiem_TT
        End If
        If m_QDNhanSu.NgayThoiLuong = DateTime.MinValue Then
            dpkNgayTL.Checked = False
        Else
            dpkNgayTL.Checked = True
            dpkNgayTL.Value = m_QDNhanSu.NgayThoiLuong
        End If
        txtNguoiQD.Text = m_QDNhanSu.NguoiKy_QD
        txtDVraQD.Text = m_QDNhanSu.DVraQD
        cboLoaiQD.SelectedValue = m_QDNhanSu.IdLoaiQD
        If dbconn.getNumber("SELECT id FROM DanhMuc WHERE id=" & m_QDNhanSu.IdLoaiQD & " And ma_so in ('1518','1519')") > 0 Then
            grpTTCu.Text = "Thông tin hiện tại"
        Else
            grpTTCu.Text = "Thông tin cũ"
        End If
        If m_QDNhanSu.IsQD_NHCS Then
            cbIsQD_NHCS.Checked = True
            txtCVNguoiQD.Visible = False
            cboCVNguoiQD.Visible = True
            labQuyetDinh.Visible = False
            txtLoaiQD.Visible = False
            labDenNgay.Visible = False
            dpkDenNgay.Visible = False
            pnlQD_IsNotNHCS.Visible = False
            pnlL_IsNotNHCS.Visible = False
            pnlPC_IsNotNHCS.Visible = False

            pnlQD_IsNHCS.Visible = True
            pnlL_IsNHCS.Visible = True
            pnlPC_IsNHCS.Visible = True
            grpTTCu.Visible = True
            grpTTCu.Enabled = False

            cboCVNguoiQD.SelectedValue = m_QDNhanSu.idCV_Nguoiky_QD
            IdDonViCu = m_QDNhanSu.IdDonvi_Cu
            IdPhongCu = m_QDNhanSu.IdPhong_Cu
            IdChucVuCu = m_QDNhanSu.IdChucvu_Cu
            IdChuyenMonCu = m_QDNhanSu.IdChuyenMon_Cu
            txtDonViCu.Text = getDonvi(IdDonViCu)
            txtPhongCu.Text = getPhong(IdPhongCu)
            txtChucVuCu.Text = getDanhmuc_Name(14, IdChucVuCu)
            txtChuyenMonCu.Text = getDanhmuc_Name(12, IdChuyenMonCu)
            cboDonviMoi.SelectedValue = m_QDNhanSu.IdDonvi_Moi
            cboPhongMoi.SelectedValue = m_QDNhanSu.IdPhong_Moi
            cboChucvuMoi.SelectedValue = m_QDNhanSu.IdChucvu_Moi
            cboChuyenmonMoi.SelectedValue = m_QDNhanSu.IdChuyenMon_Moi
        Else
            cbIsQD_NHCS.Checked = False
            cboCVNguoiQD.Visible = False
            txtCVNguoiQD.Visible = True
            labQuyetDinh.Visible = True
            txtLoaiQD.Visible = True
            grpTTCu.Visible = False
            pnlQD_IsNHCS.Visible = False
            pnlQD_IsNotNHCS.Visible = True
            txtNoiDung_QD.Enabled = True
            labDenNgay.Visible = True
            dpkDenNgay.Visible = True
            txtCVNguoiQD.Text = m_QDNhanSu.CV_NguoiKy_QD
            txtLoaiQD.Text = m_QDNhanSu.LoaiQD
            If m_QDNhanSu.DenNgay = DateTime.MinValue Then
                dpkDenNgay.Checked = False
            Else
                dpkDenNgay.Checked = True
                dpkDenNgay.Value = m_QDNhanSu.DenNgay
            End If
            txtNoiDung_QD.Text = m_QDNhanSu.NoiDung
        End If

        If m_QDNhanSu.Active Then
            cbActive.Checked = True
        Else : cbActive.Checked = False
        End If
        txtGhichu.Text = m_QDNhanSu.GhiChu

        NgayHLCU = m_QDNhanSu.NgayHL
        SoQDCU = m_QDNhanSu.So_QD
        NgayQDCU = m_QDNhanSu.NgayKy_QD

        ' fill thông tin lương trong quyết định
        idLuong = m_Luong.getID(idCanBo, NgayHLCU, SoQDCU, NgayQDCU)
        If idLuong <> "" Then
            cbLuong.Checked = True
            m_Luong = m_Luong.getRecord(idLuong)
            If m_QDNhanSu.IsQD_NHCS Then
                txtHeso.Text = m_Luong.HeSoLuong
                getIdNDBangNgach(m_Luong.IdBacLuong, vIdNghiDinhLuong, vIdBangluong, vIdNgachluong)
                cboNghiDinh.SelectedValue = vIdNghiDinhLuong
                cboBang.SelectedValue = vIdBangluong
                cboNgach.SelectedValue = vIdNgachluong
                cboBac.SelectedValue = m_Luong.IdBacLuong
            Else
                txtNoiDung_L.Enabled = True
                txtNoiDung_L.Text = m_Luong.NoiDung
            End If
        Else
            cbLuong.Checked = False
            cboNghiDinh.SelectedValue = 0
            cboBang.SelectedValue = 0
            cboNgach.SelectedValue = 0
            cboBac.SelectedValue = 0
            txtNoiDung_L.Text = ""
        End If

        ' fill thông tin phụ cấp liên quan trong quyết định
        idPhucap = m_Phucap.getID(idCanBo, NgayHLCU, SoQDCU, NgayQDCU)
        If idPhucap <> "" Then
            cbPhucap.Checked = True
            m_Phucap = m_Phucap.getRecord(idPhucap)
            If m_QDNhanSu.IsQD_NHCS Then
                vIdLoaiPC = getIdLoaiPC(m_Phucap.IdMucPC)
                cboLoaiPC.SelectedValue = vIdLoaiPC
                cboMucPC.SelectedValue = m_Phucap.IdMucPC
            Else
                txtNoiDung_PC.Text = m_Phucap.NoiDung
            End If
        Else
            cbPhucap.Checked = False
            cboMucPC.SelectedValue = 0
            cboLoaiPC.SelectedValue = 0
            txtNoiDung_PC.Text = ""
        End If

        'Nếu QD nhân sự cuối cùng của cán bộ thuộc chi nhánh khác trong NHCS thì hiển thị nút xuất hồ sơ thông tin cán bộ
        Dim QDNhanSufinalNHCS As QDNhanSu = New QDNhanSu
        QDNhanSufinalNHCS = QDNhanSufinalNHCS.getFinalRecord(vIdCanBo, Nothing, True)
        If QDNhanSufinalNHCS.IdQDNhanSu <> "" And QDNhanSufinalNHCS.IdDonvi_Moi <> IdDONVI Then
            bntExportHSCB.Enabled = True
        Else
            bntExportHSCB.Enabled = False
        End If
    End Sub

    Public Sub getOldInfQDNhanSu(ByVal vIdQDNhanSu As String, ByVal vIdCanBo As String, ByRef vIdDonViCu As Integer, ByRef vDonViCu As String, ByRef vIdPhongCu As Integer, ByRef vPhongCu As String, ByRef vIdChucVuCu As Integer, ByRef vChucVuCu As String, ByRef vIdChuyenMonCu As Integer, ByRef vChuyenMonCu As String)
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 IdDonVi_Moi as IdDonVi_Cu, (SELECT ten_goi FROM chinhanh WHERE Id= IdDonVi_Moi) as DonViCu," & _
                   " IdPhong_Moi as IdPhong_Cu, (SELECT ten_phong FROM PHONGBAN WHERE Id=IdPhong_Moi) as PhongCu," & _
                   " IdChucVu_Moi as IdChucVu_Cu, (SELECT ten_goi FROM Danhmuc WHERE id= IdChucVu_Moi) as ChucVuCu," & _
                   " IdChuyenMon_Moi as IdChuyenMon_Cu, (SELECT ten_goi FROM Danhmuc WHERE id= IdChuyenMon_Moi) as ChuyenMonCu" & _
                   " FROM QDNhanSu WHERE ngayHL<(select b.ngayHL FROM QDnhansu b where b.IDQDNhanSu='" & vIdQDNhanSu & "' AND b.IsQD_NHCS=1) AND idcanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 order by NgayHL desc"
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            dt = db.SelectDBRows(strSql)
            If dt.Rows.Count > 0 Then
                vIdDonViCu = dt.Rows(0).Item("IdDonVi_Cu")
                vIdPhongCu = dt.Rows(0).Item("IdPhong_Cu")
                vIdChucVuCu = dt.Rows(0).Item("IdChucVu_Cu")
                vIdChuyenMonCu = dt.Rows(0).Item("IdChuyenMon_Cu")
                vDonViCu = dt.Rows(0).Item("DonViCu")
                vPhongCu = dt.Rows(0).Item("PhongCu")
                vChucVuCu = dt.Rows(0).Item("ChucVuCu")
                vChuyenMonCu = dt.Rows(0).Item("ChuyenMonCu")
            Else
                vIdDonViCu = 0
                vIdPhongCu = 0
                vIdChucVuCu = 0
                vIdChuyenMonCu = 0
                vDonViCu = ""
                vPhongCu = ""
                vChucVuCu = ""
                vChuyenMonCu = ""
            End If
        Catch ex As Exception
            vDonViCu = ""
            vPhongCu = ""
            vChucVuCu = ""
            vChuyenMonCu = ""
        End Try
    End Sub

    Private Function updateQDNhanSu(ByRef vIdQDNhanSu As String) As Boolean
        Try
            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
            Dim dateHL As Date
            Dim dateQD As Date

            dateHL = DateTimeUtil.getDateCurrTime(dpkNgayHL.Text)
            dateQD = DateTimeUtil.getDate(dpkNgayQD.Text)
            m_QDNhanSu.IdQDNhanSu = vIdQDNhanSu
            m_QDNhanSu.IdCanBo = idCanBo
            m_QDNhanSu.So_QD = txtSoQD.Text.Trim
            m_QDNhanSu.NgayKy_QD = dateQD
            m_QDNhanSu.NguoiKy_QD = standardizeName(txtNguoiQD.Text)
            m_QDNhanSu.NgayHL = dateHL
            If dpkNgayBN.Checked Then
                m_QDNhanSu.NgayBoNhiem_TT = DateTimeUtil.getDate(dpkNgayBN.Text)
            Else
                m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
            End If
            If dpkNgayTL.Checked Then
                m_QDNhanSu.NgayThoiLuong = DateTimeUtil.getDate(dpkNgayTL.Text)
            Else
                m_QDNhanSu.NgayThoiLuong = DateTime.MinValue
            End If
            m_QDNhanSu.DVraQD = standardizeString(txtDVraQD.Text.Trim)
            m_QDNhanSu.IdLoaiQD = CInt(cboLoaiQD.SelectedValue)
            If cbIsQD_NHCS.Checked Then
                Dim m_QDNSfinal As QDNhanSu = New QDNhanSu
                m_QDNSfinal = m_QDNSfinal.getFinalRecord(idCanBo, dpkNgayHL.Value)
                m_QDNhanSu.IsQD_NHCS = 1
                m_QDNhanSu.IdDonvi_Cu = m_QDNSfinal.IdDonvi_Moi
                m_QDNhanSu.IdPhong_Cu = m_QDNSfinal.IdPhong_Moi
                m_QDNhanSu.IdChucvu_Cu = m_QDNSfinal.IdChucvu_Moi
                m_QDNhanSu.IdChuyenMon_Cu = m_QDNSfinal.IdChuyenMon_Moi
                m_QDNhanSu.IdDonvi_Moi = CInt(cboDonviMoi.SelectedValue)
                m_QDNhanSu.IdPhong_Moi = CInt(cboPhongMoi.SelectedValue)
                m_QDNhanSu.IdChucvu_Moi = CInt(cboChucvuMoi.SelectedValue)
                m_QDNhanSu.IdChuyenMon_Moi = CInt(cboChuyenmonMoi.SelectedValue)
                m_QDNhanSu.idCV_Nguoiky_QD = CInt(cboCVNguoiQD.SelectedValue)
                m_QDNhanSu.CV_NguoiKy_QD = ""
                m_QDNhanSu.LoaiQD = ""
                m_QDNhanSu.DenNgay = DateTime.MinValue
                m_QDNhanSu.NoiDung = ""
            Else
                m_QDNhanSu.IsQD_NHCS = 0
                m_QDNhanSu.IdDonvi_Cu = 0
                m_QDNhanSu.IdPhong_Cu = 0
                m_QDNhanSu.IdChucvu_Cu = 0
                m_QDNhanSu.IdChuyenMon_Cu = 0
                m_QDNhanSu.IdDonvi_Moi = 0
                m_QDNhanSu.IdPhong_Moi = 0
                m_QDNhanSu.IdChucvu_Moi = 0
                If dpkDenNgay.Checked Then
                    m_QDNhanSu.DenNgay = DateTimeUtil.getDate(dpkDenNgay.Text)
                Else
                    m_QDNhanSu.DenNgay = DateTime.MinValue
                End If
                m_QDNhanSu.CV_NguoiKy_QD = standardizeString(txtCVNguoiQD.Text.Trim)
                m_QDNhanSu.idCV_Nguoiky_QD = 0
                m_QDNhanSu.LoaiQD = standardizeString(txtLoaiQD.Text.Trim)
                m_QDNhanSu.NoiDung = standardizeString(txtNoiDung_QD.Text)
            End If

            Select Case getDanhmuc_MaSo(m_QDNhanSu.IdLoaiQD)
                Case "1518"
                    ' Kiem nhiem
                    m_QDNhanSu.IsKiemNhiem = 1
                Case "1519"
                    ' Thoi kiem nhiem
                    m_QDNhanSu.IsKiemNhiem = 2
                Case Else
                    m_QDNhanSu.IsKiemNhiem = 0
            End Select

            If cbActive.Checked Then
                m_QDNhanSu.Active = 1
            Else
                m_QDNhanSu.Active = 0
            End If
            m_QDNhanSu.GhiChu = standardizeString(txtGhichu.Text)

            If vIdQDNhanSu <> "" Then
                m_QDNhanSu.Update()
            Else
                vIdQDNhanSu = m_QDNhanSu.Add()
            End If
            If m_QDNhanSu.IdQDNhanSu <> "" Then
                If m_QDNhanSu.IsKiemNhiem = 2 Then
                    'cap nhat lai QD kiem nhiem truoc do
                    dbconn.executeSQL("UPDATE QDNhanSu SET IsKiemNhiem=2 WHERE IdCanBo='" & idCanBo & "' and IsKiemNhiem=1 and IdDonvi_Moi=" & m_QDNhanSu.IdDonvi_Moi & " and IdPhong_Moi=" & m_QDNhanSu.IdPhong_Moi & " and IdChucVu_Moi=" & m_QDNhanSu.IdChucvu_Moi)
                End If
                Dim m_Luong As LuongCanBo = New LuongCanBo
                Dim m_PhuCap As Phucap = New Phucap
                Dim maQDNhanSu As String = ""
                Dim IDLoaiQDLuong As String = ""
                Dim db As DBAccess = New DBAccess
                m_Luong.IdLuongCB = idLuong
                m_Luong.IdCanBo = idCanBo
                If cbLuong.Checked Then
                    m_Luong.Ngay_Huong = dateHL
                    m_Luong.NgayLen_DK = dateHL.AddMonths(getTimeNangBac(cboNgach.SelectedValue))
                    m_Luong.DVraQD = m_QDNhanSu.DVraQD
                    If cbIsQD_NHCS.Checked Then
                        m_Luong.IsQD_NHCS = 1
                        m_Luong.NoiDung = ""
                        m_Luong.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                        m_Luong.CV_NguoiKy_QD = ""
                        m_Luong.IdBacLuong = CInt(cboBac.SelectedValue)
                        m_Luong.HeSoLuong = CDbl(txtHeso.Text)
                        m_Luong.LoaiQD = ""
                    Else
                        m_Luong.IdCV_Nguoi_QD = 0
                        m_Luong.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                        m_Luong.NoiDung = txtNoiDung_L.Text.Trim
                        m_Luong.IsQD_NHCS = 0
                        m_Luong.LoaiQD = m_QDNhanSu.LoaiQD
                    End If
                    maQDNhanSu = getDanhmuc_MaSo(m_QDNhanSu.IdLoaiQD)
                    IDLoaiQDLuong = db.getNumber("SELECT ID FROM DanhMuc WHERE ma_so ='43" & (CInt(maQDNhanSu.Substring(2, 1)) + 1) & maQDNhanSu.Substring(3, 1) & "' and ma_so in ('4311','4312','4313','4314','4315','4317','4318','4319','4327','4328','4329','4330','4331','4332','4342','4343','4323','4324','4325')")
                    If IDLoaiQDLuong = 0 Then
                        IDLoaiQDLuong = getDanhmuc_ID("4323")
                    End If
                    m_Luong.IdLoaiQD = IDLoaiQDLuong
                    m_Luong.SoQD = m_QDNhanSu.So_QD
                    m_Luong.NgayQD = dateQD
                    m_Luong.NguoiQD = m_QDNhanSu.NguoiKy_QD
                    If m_QDNhanSu.GhiChu.Length > 100 Then
                        m_Luong.GhiChu = m_QDNhanSu.GhiChu.Substring(0, 99)
                    Else
                        m_Luong.GhiChu = m_QDNhanSu.GhiChu
                    End If
                    If idLuong <> "" Then
                        m_Luong.Update()
                    Else
                        ' Kiem tra thông tin lương trong QDNS neu trung với thông tin lương hiện đang hưởng của CB thì ko cho nhập sang HS Luong
                        Dim m_LuongFinal As LuongCanBo = New LuongCanBo
                        m_LuongFinal = m_LuongFinal.getFinalRecord(idCanBo)
                        If Not (m_LuongFinal.IdBacLuong = m_Luong.IdBacLuong And m_LuongFinal.IdLoaiQD = m_Luong.IdLoaiQD And m_Luong.IsQD_NHCS) Then
                            m_Luong.Add()
                        End If
                    End If
                Else
                    If idLuong <> "" Then
                        m_Luong.Delete()
                    End If
                End If
                m_PhuCap.IdCB_PhuCap = idPhucap
                m_PhuCap.IdCanBo = idCanBo
                If cbPhucap.Checked Then
                    If cbIsQD_NHCS.Checked Then
                        m_PhuCap.IsQD_NHCS = 1
                        m_PhuCap.NoiDung = ""
                        m_PhuCap.IdMucPC = CInt(cboMucPC.SelectedValue)
                        m_PhuCap.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                        m_PhuCap.CV_NguoiKy_QD = ""
                    Else
                        m_PhuCap.IdCV_Nguoi_QD = 0
                        m_PhuCap.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                        m_PhuCap.IsQD_NHCS = 0
                        m_PhuCap.NoiDung = txtNoiDung_PC.Text.Trim
                    End If
                    m_PhuCap.TuNgay = DateTimeUtil.getDate(dpkNgayHL.Text)  'dateHL
                    m_PhuCap.DenNgay = DateTime.MinValue
                    m_PhuCap.SoQD = m_QDNhanSu.So_QD
                    m_PhuCap.NgayQD = dateQD
                    m_PhuCap.NguoiQD = m_QDNhanSu.NguoiKy_QD
                    m_PhuCap.DVraQD = m_QDNhanSu.DVraQD
                    If m_QDNhanSu.GhiChu.Length > 100 Then
                        m_PhuCap.GhiChu = m_QDNhanSu.GhiChu.Substring(0, 99)
                    Else
                        m_PhuCap.GhiChu = m_QDNhanSu.GhiChu
                    End If
                    If idPhucap <> "" Then
                        Dim dtPC As DataTable
                        Dim iPC As Integer = 0
                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap FROM HS_PhucapCB WHERE IdCanbo='" & idCanBo & "' and Datediff(day,TuNgay,'" & NgayHLCU & "')=0 and Upper(SoQD)=Upper(N'" & SoQDCU & "') and NgayQD='" & NgayQDCU & "' and IdMucPC<>" & m_PhuCap.IdMucPC)
                        m_PhuCap.Update()
                        ' Update tiếp với các Phụ cấp khác có cùng số QĐ và ngày
                        If dtPC.Rows.Count > 0 Then
                            For iPC = 0 To dtPC.Rows.Count - 1
                                m_PhuCap.IdCB_PhuCap = dtPC.Rows(iPC).Item("IdCB_PhuCap")
                                m_PhuCap.UpdateNotFull()
                            Next
                        End If
                    Else
                        Dim dtPC As DataTable
                        m_PhuCap.Add()
                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap, DenNgay FROM HS_PhucapCB t1, MucPhuCap t2 WHERE t1.IdMucPC=t2.IdMuc_PhC and idcanbo='" & idCanBo & "' and IsQD_NHCS=1 and IdLoai_PhC=" & CInt(cboLoaiPC.SelectedValue) & " order by TuNgay desc")
                        If dtPC.Rows.Count >= 2 Then
                            If dtPC.Rows(1).Item("DenNgay") Is DBNull.Value Then
                                db.executeSQL("UPDATE HS_PhuCapCB Set DenNgay='" & m_PhuCap.TuNgay.AddDays(-1) & "' WHERE IdCB_PhuCap= '" & dtPC.Rows(1).Item("IdCB_PhuCap") & "'")
                            End If
                        End If
                    End If
                Else
                    If idPhucap <> "" Then
                        m_PhuCap.Delete()
                    End If
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'Public Function findStringInValueMember(ByVal cbo As ComboBox, ByVal criteria As String)
    '    Dim item As Object
    '    Dim s As System.Type
    '    ' Case insensitive
    '    Dim lowerCriteria As String = criteria.ToLower
    '    ' Keep track of index for return value
    '    Dim index As Integer = 0

    '    ' Get the system type of the datasource
    '    s = cbo.DataSource.GetType

    '    ' Use reflection to get the property info for the combobox value member
    '    Dim valueMemberProp As System.Reflection.PropertyInfo = s.GetProperty(cbo.ValueMember)

    '    ' Loop through each item in the data source
    '    For Each item In cbo.DataSource
    '        Dim strValueMember As String
    '        ' Get the value of the value member as a string
    '        strValueMember = valueMemberProp.GetValue(item, Nothing)

    '        ' Do lowercase comparision (case insensitve again)
    '        If strValueMember.ToLower.StartsWith(lowerCriteria) Then
    '            ' Return index if found
    '            Return index
    '        End If
    '        index += 1
    '    Next
    '    ' Not found, return -1
    '    Return -1
    'End Function

    Private Sub cbIsQD_NHCS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbIsQD_NHCS.CheckedChanged
        If cbIsQD_NHCS.Checked Then
            txtCVNguoiQD.Visible = False
            cboCVNguoiQD.Visible = True
            labQuyetDinh.Visible = False
            txtLoaiQD.Visible = False
            labNgayHL.Text = "Ngày hiệu lực"
            labDenNgay.Visible = False
            dpkDenNgay.Visible = False
            pnlQD_IsNotNHCS.Visible = False
            pnlQD_IsNHCS.Visible = True

            pnlL_IsNotNHCS.Visible = False
            pnlPC_IsNotNHCS.Visible = False
            pnlL_IsNHCS.Visible = True
            pnlPC_IsNHCS.Visible = True
            grpTTCu.Visible = True
            grpTTCu.Enabled = False
            If idQDNhansu = "" Then txtDVraQD.Text = getDonvi(IdDONVI)
        Else
            cboCVNguoiQD.Visible = False
            txtCVNguoiQD.Visible = True
            labNgayHL.Text = "Từ ngày"
            labQuyetDinh.Visible = True
            txtLoaiQD.Visible = True
            grpTTCu.Visible = False
            pnlQD_IsNHCS.Visible = False
            pnlQD_IsNotNHCS.Visible = True
            labDenNgay.Visible = True
            dpkDenNgay.Visible = True
            pnlQD_IsNotNHCS.Enabled = True
            labDenNgay.Enabled = True
            dpkDenNgay.Enabled = True

            pnlL_IsNHCS.Visible = False
            pnlPC_IsNHCS.Visible = False
            pnlL_IsNotNHCS.Visible = True
            pnlPC_IsNotNHCS.Visible = True
            pnlL_IsNotNHCS.Enabled = True
            txtNoiDung_L.Enabled = False
            pnlPC_IsNotNHCS.Enabled = True
        End If
    End Sub

    Private Sub cboDonviMoi_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDonviMoi.SelectedValueChanged
        If Not (cboDonviMoi.SelectedValue Is Nothing) Then
            bindCboPhongMoi(CInt(cboDonviMoi.SelectedValue))
            bindCboChucVu(CInt(cboDonviMoi.SelectedValue))
        End If
    End Sub

    Private Sub cboNgach_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboNgach.SelectedValueChanged
        If Not (cboNgach.SelectedValue Is Nothing) Then
            If CInt(cboNgach.SelectedValue) > 0 Then cboBac.DataSource = listBacLuong(CInt(cboNgach.SelectedValue), 0, True)
        End If
    End Sub

    Private Sub cboNghiDinh_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboNghiDinh.SelectedValueChanged
        If Not (cboNghiDinh.SelectedValue Is Nothing) Then
            If CInt(cboNghiDinh.SelectedValue) > 0 Then cboBang.DataSource = listBangLuong(CInt(cboNghiDinh.SelectedValue), True)
        End If
    End Sub

    Private Sub cboBang_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBang.SelectedValueChanged
        If Not (cboBang.SelectedValue Is Nothing) Then
            If CInt(cboBang.SelectedValue) > 0 Then cboNgach.DataSource = listNgachLuong(CInt(cboBang.SelectedValue), True)
        End If
    End Sub

    Private Sub cboBac_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBac.SelectedValueChanged
        If Not (cboBac.SelectedValue Is Nothing) Then
            If CInt(cboBac.SelectedValue) > 0 Then
                Try
                    txtHeso.Text = getHeso(CInt(cboBac.SelectedValue))
                Catch ex As Exception
                    txtHeso.Text = ""
                End Try
            End If
        End If
    End Sub

    Private Sub cboLoaiPC_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLoaiPC.SelectedValueChanged
        If Not (cboLoaiPC.SelectedValue Is Nothing) Then
            If CInt(cboLoaiPC.SelectedValue) > 0 Then cboMucPC.DataSource = listMucPC(CInt(cboLoaiPC.SelectedValue))
        End If
    End Sub

    Private Sub cbLuong_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbLuong.CheckedChanged
        InfLuong(cbLuong.Checked)
    End Sub

    Private Sub cbPhucap_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPhucap.CheckedChanged
        InfPhucap(cbPhucap.Checked)
    End Sub

    'Private Sub cboDonviMoi_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboDonviMoi.KeyPress
    '    'Dim i As Integer
    '    'i = findStringInValueMember(cboDonviMoi, cboDonviMoi.Text)
    '    'Dim asd As Integer = 0
    '    'Dim db As DBAccess
    '    'Dim strsql As String
    '    'strsql = " SELECT id as Value, ten_goi as Display FROM ten_goi like %'" & cboDonviMoi.Text & "'% Order by Ma_so"
    '    'cboDonviMoi.Update = db.SelectDBRows(strsql)
    '    cboDonviMoi.FindStringExact(cboDonviMoi.Text)
    'End Sub

    Private Sub gridQDNhanSu_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridQDNhanSu.CellClick
        Try
            idQDNhansu = gridQDNhanSu.CurrentRow.Cells("IdQDNS").Value.ToString
            fillQDNhansu(idQDNhansu)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDNhanSu_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridQDNhanSu.CellFormatting
        Try
            If gridQDNhanSu.Columns(e.ColumnIndex).Name = "NgayHL" Then
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew.Click
        blankQDNhansu()
    End Sub

    Private Sub bntAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Try
            Dim lab_ErrQDNS As String = ""
            If (idQDNhansu <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";101;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu quyết định nhân sự!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridQDNhanSu_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Else
                lab_ErrQDNS = checkTabQDNhanSu()
                If lab_ErrQDNS <> "" Then
                    MessageBox.Show(lab_ErrQDNS, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    Exit Sub
                End If
                Dim IDDV_Moi, IDP_Moi, IDDV_Cu, IDP_Cu As Integer
                IDDV_Moi = CInt(cboDonviMoi.SelectedValue)
                IDP_Moi = CInt(cboPhongMoi.SelectedValue)
                IDDV_Cu = IdDonViCu
                IDP_Cu = IdPhongCu
                If updateQDNhanSu(idQDNhansu) Then
                    Dim lastQD As QDNhanSu = New QDNhanSu
                    lastQD = lastQD.getFinalRecord(idCanBo)
                    If lastQD.IdDonvi_Moi = IDDV_Moi And lastQD.IdPhong_Moi = IDP_Moi Then
                        Dim childNode As TreeNode
                        Dim _node As TreeNode
                        Dim isExits As Boolean = False

                        _node = treeCocau.SelectedNode.Parent
                        If Not (IDDV_Cu = 0 And IDP_Cu = 0) Then
                            refreshNodeNOTCanbo(treeCocau, _node, IDDV_Cu, IDP_Cu, False, False, idCanBo)
                        End If
                        For Each childNode In treeCocau.Nodes(0).Nodes
                            Dim arr1() As String
                            Dim arr0() As String
                            arr1 = childNode.Tag.ToString.Split("_")
                            arr0 = childNode.Parent.Tag.ToString.Split("_")
                            If arr1(0) = "DV" Then
                                Dim IdGoc As Integer = 0
                                Dim db As DBAccess = New DBAccess
                                Dim DP_Node As TreeNode
                                IdGoc = db.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id in (SELECT Id_goc FROM ChiNhanh WHERE Id=" & arr1(1) & ")")
                                If IdGoc = 1 Then
                                    For Each DP_Node In childNode.Nodes
                                        Dim arr1_DP() As String
                                        Dim arr0_DP() As String
                                        arr1_DP = DP_Node.Tag.ToString.Split("_")
                                        arr0_DP = DP_Node.Parent.Tag.ToString.Split("_")
                                        If (arr0_DP(1) = IDDV_Moi And arr1_DP(2) = IDP_Moi) Then
                                            refreshNodeNOTCanbo(treeCocau, DP_Node, arr0_DP(1), arr1_DP(2), False, False, idCanBo)
                                            isExits = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else
                                If (arr0(1) = IDDV_Moi And arr1(2) = IDP_Moi) Then
                                    refreshNodeNOTCanbo(treeCocau, childNode, arr0(1), arr1(2), False, False, idCanBo)
                                    isExits = True
                                    Exit For
                                End If
                            End If
                            If isExits Then Exit For
                        Next
                    End If
                    If cbActive.Checked Then
                        bindGridQDNhansu(idCanBo)
                        fillQDNhansu(idQDNhansu, idCanBo)
                    End If
                    labAlert_QD.Text = "Ghi dữ liệu thành công!"
                Else
                    MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel.Click
        If idQDNhansu = "" Then
            blankQDNhansu()
        Else
            fillQDNhansu(idQDNhansu)
        End If
    End Sub

    Private Sub bntDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete.Click
        Try
            labAlert_QD.Text = ""
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If gridQDNhanSu.Rows.Count = 1 Then
                MessageBox.Show("Quyết định này không được xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridQDNhanSu.Rows.Count > 0) Then
                For i As Int32 = 0 To gridQDNhanSu.Rows.Count - 1
                    If (CType(gridQDNhanSu.Rows(i).Cells("cln_cbQDNS").Value, Boolean) = True) Then
                        arrDel.Add(gridQDNhanSu.Rows(i).Cells("IdQDNS").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idQDNhansu <> "") Then arrDel.Add(idQDNhansu.ToString)

            If (arrDel.Count > 0) Then
                Dim IDDV_Moi, IDP_Moi As Integer
                Dim IDQD_Final As String
                Dim m_QDNhanSuFinal As QDNhanSu = New QDNhanSu
                m_QDNhanSuFinal = m_QDNhanSuFinal.getFinalRecord(idCanBo, Nothing, True)
                IDQD_Final = m_QDNhanSuFinal.IdQDNhanSu
                IDDV_Moi = m_QDNhanSuFinal.IdDonvi_Moi
                IDP_Moi = m_QDNhanSuFinal.IdPhong_Moi
                If arrDel.Count = 1 And arrDel(0).ToString = IDQD_Final Then
                    MessageBox.Show("Quyết định này không được xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Exit Sub
                End If
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        'Dim IDDV_Moi, IDP_Moi, IDDV_Cu, IDP_Cu As Integer
                        'Dim m_QDNhanSuFinal As QDNhanSu = New QDNhanSu
                        'm_QDNhanSuFinal = m_QDNhanSuFinal.getFinalRecord(idCanBo)
                        'IDDV_Cu = m_QDNhanSuFinal.IdDonvi_Moi
                        'IDP_Cu = m_QDNhanSuFinal.IdPhong_Moi
                        'Dim i As Integer
                        'For i = 0 To arrDel.Count - 1
                        '    Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                        '    m_QDNhanSu.IdQDNhanSu = arrDel(i).ToString()
                        '    m_QDNhanSu.Delete()
                        'Next
                        'm_QDNhanSuFinal = New QDNhanSu
                        'm_QDNhanSuFinal = m_QDNhanSuFinal.getFinalRecord(idCanBo)
                        'IDDV_Moi = m_QDNhanSuFinal.IdDonvi_Moi
                        'IDP_Moi = m_QDNhanSuFinal.IdPhong_Moi

                        'Dim childNode As TreeNode
                        'Dim _node As TreeNode
                        'Dim isExits As Boolean = False

                        '_node = treeCocau.SelectedNode.Parent
                        'refreshNodeNOTCanbo(treeCocau, _node, IDDV_Cu, IDP_Cu, False, False, idCanBo)
                        'For Each childNode In treeCocau.Nodes(0).Nodes
                        '    Dim arr1() As String
                        '    Dim arr0() As String
                        '    arr1 = childNode.Tag.ToString.Split("_")
                        '    arr0 = childNode.Parent.Tag.ToString.Split("_")
                        '    If (arr0(1) = IDDV_Moi And arr1(2) = IDP_Moi) Then
                        '        refreshNodeNOTCanbo(treeCocau, childNode, arr0(1), arr1(2), False, False, idCanBo)
                        '        isExits = True
                        '        Exit For
                        '    End If
                        'Next

                        ' xu ly ko cho xoa QD Nhan su hien tai cua can bo

                        '''''''''''''''''''''''''''''''''''''''''''''''''

                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            If IDQD_Final <> arrDel(i).ToString() Then
                                Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                                m_QDNhanSu.IdQDNhanSu = arrDel(i).ToString()
                                m_QDNhanSu.Delete()
                            End If
                        Next
                        Dim lastQD As QDNhanSu = New QDNhanSu
                        lastQD = lastQD.getFinalRecord(idCanBo)
                        If Not (lastQD.IdDonvi_Moi = IDDV_Moi And lastQD.IdPhong_Moi = IDP_Moi) Then
                            Dim childNode As TreeNode
                            'Dim _node As TreeNode
                            Dim isExits As Boolean = False

                            '_node = treeCocau.SelectedNode.Parent
                            'refreshNodeNOTCanbo(treeCocau, _node, IDDV_Cu, IDP_Cu, False, False, idCanBo)
                            For Each childNode In treeCocau.Nodes(0).Nodes
                                Dim arr1() As String
                                Dim arr0() As String
                                arr1 = childNode.Tag.ToString.Split("_")
                                arr0 = childNode.Parent.Tag.ToString.Split("_")
                                If (arr0(1) = IDDV_Moi And arr1(2) = IDP_Moi) Then
                                    refreshNodeNOTCanbo(treeCocau, childNode, arr0(1), arr1(2), False, False, idCanBo)
                                    isExits = True
                                    Exit For
                                End If
                            Next
                        End If
                        If cbActive.Checked Then
                            bindGridQDNhansu(idCanBo)
                            fillQDNhansu("", idCanBo)
                        End If
                        labAlert_QD.Text = "Xoá dữ liệu thành công!"
                    Catch ex As Exception
                        MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Else
                MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub gridQDNhanSu_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridQDNhanSu.KeyUp
        gridQDNhanSu_CellClick(sender, Nothing)
    End Sub

    'Private Sub dpkNgayHL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpkNgayHL.LostFocus
    '    Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
    '    m_QDNhanSu = m_QDNhanSu.getFinalRecord(idCanBo, dpkNgayHL.Value)
    '    If m_QDNhanSu.IdQDNhanSu <> "" Then
    '        If idQDNhansu <> "" Then
    '            getOldInfQDNhanSu(m_QDNhanSu.IdQDNhanSu, idCanBo, IdDonViCu, txtDonViCu.Text, IdPhongCu, txtPhongCu.Text, IdChucVuCu, txtChucVuCu.Text, IdChuyenMonCu, txtChuyenMonCu.Text)
    '        Else
    '            IdDonViCu = m_QDNhanSu.IdDonvi_Moi
    '            txtDonViCu.Text = getDonvi(IdDonViCu)
    '            IdPhongCu = m_QDNhanSu.IdPhong_Moi
    '            txtPhongCu.Text = getPhong(IdPhongCu)
    '            IdChucVuCu = m_QDNhanSu.IdChucvu_Moi
    '            txtChucVuCu.Text = getDanhmuc_Name(14, IdChucVuCu)
    '            IdChuyenMonCu = m_QDNhanSu.IdChuyenMon_Moi
    '            txtChuyenMonCu.Text = getDanhmuc_Name(12, IdChuyenMonCu)
    '        End If
    '    Else
    '        txtDonViCu.Text = ""
    '        txtPhongCu.Text = ""
    '        txtChucVuCu.Text = ""
    '        txtChuyenMonCu.Text = ""
    '    End If
    'End Sub

    Private Sub txtNguoiQD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNguoiQD.LostFocus
        txtNguoiQD.Text = standardizeName(txtNguoiQD.Text)
    End Sub

    Private Sub txtGhichu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGhichu.LostFocus
        txtGhichu.Text = standardizeString(txtGhichu.Text)
    End Sub

#End Region

#Region "tab Quyết định thôi viêc của cán bộ"
    Private Sub initTabQDThoiViec()
        cbIsQD_NHCS_TV.Checked = True
        If Not initTV Then
            initTV = True
            bindCboLoaiQD_TV()
            bindCboLyDo_TV()
            bindCboChucvuNguoiKyQD_TV()
        End If
        bindGridTV(idCanBo)
        fillQDThoiViec(idQDThoiViec, idCanBo)
        gridThoiViec.Focus()
    End Sub

    Private Sub bindCboLoaiQD_TV()
        cboQDThoiViec.DataSource = listDanhmuc(37)
    End Sub

    Private Sub bindCboLyDo_TV()
        cboLyDo_TV.DataSource = listDanhmuc(22)
    End Sub

    Private Sub bindCboChucvuNguoiKyQD_TV()
        cboCVNguoiKyQD_TV.DataSource = listChucVuQuyenRaQD()
    End Sub

    Private Sub bindGridTV(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridThoiViec.AutoGenerateColumns = False
                Dim m_ThoiViec As CBThoiViec = New CBThoiViec
                gridThoiViec.DataSource = m_ThoiViec.getAllByCanbo(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridThoiViec.Rows.Count - 1
                    If Trim(gridThoiViec.Rows(i).Cells(0).Value) = idQDThoiViec Then
                        gridThoiViec.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Quyết định thôi việc cán bộ !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabQDThoiviec() As String
        Dim strReturn As String = ""
        Try
            If (idCanBo = "") Then
                strReturn = "Bạn chưa chọn cán bộ cần cập nhật cán bộ thôi việc"
                Exit Try
            End If
            If (cboQDThoiViec.Items.Count = 0) Then
                cboQDThoiViec.Focus()
                strReturn = "Bạn chưa chọn loại quyết định thôi việc!"
                Exit Try
            End If
            If (txtNguoiKyQD_TV.Text.Trim() = "") Then
                txtNguoiKyQD_TV.Focus()
                strReturn = "Người ký quyết định không được để trống!"
                Exit Try
            End If
            If (cboCVNguoiKyQD_TV.Items.Count = 0) Then
                cboCVNguoiKyQD_TV.Focus()
                strReturn = "Bạn chưa chọn chức vụ người ký quyết định thôi việc!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankQDThoiViec()
        idQDThoiViec = ""
        txtNguoiKyQD_TV.Text = ""
        dpkNgayKy_TV.Value = Date.Now
        dpkNgayHL_TV.Value = Date.Now
        bntDelete_TV.Enabled = False
        txtTroCap_TV.Text = ""
        txtTroCapKhac.Text = ""
        txtTienBoiThuong.Text = ""
        txtTienThuHoi.Text = ""
        txtGhiChu_TV.Text = ""
        txtSoQD_TV.Text = ""
        txtDVraQD_TV.Text = ""
        labAlert_TV.Text = ""
        cbIsQD_NHCS_TV.Focus()
    End Sub

    Private Sub fillQDThoiViec(ByRef vIdQDThoiViec As String, Optional ByVal vIdCanBo As String = "")
        Dim m_QDThoiViec As CBThoiViec = New CBThoiViec
        labAlert_TV.Text = ""
        If vIdCanBo = "" Then
            m_QDThoiViec = m_QDThoiViec.getRecord(vIdQDThoiViec)
        Else
            m_QDThoiViec = m_QDThoiViec.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";142;") < 0) Then  ' Xoá
            bntDelete_TV.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDelete_TV.Enabled = True
                Else
                    bntDelete_TV.Enabled = False
                End If
            End If
        End If
        If m_QDThoiViec.IdCBThoiViec <> "" Then
            vIdQDThoiViec = m_QDThoiViec.IdCBThoiViec
            cboQDThoiViec.SelectedValue = m_QDThoiViec.IdLoaiQD
            txtNguoiKyQD_TV.Text = m_QDThoiViec.NguoiKy_QD
            dpkNgayKy_TV.Value = m_QDThoiViec.NgayKy_QD
            dpkNgayHL_TV.Value = m_QDThoiViec.Ngay_HL
            cboLyDo_TV.SelectedValue = m_QDThoiViec.IdLyDo
            txtTroCap_TV.Text = m_QDThoiViec.TroCap_ThoiViec.ToString
            txtTroCapKhac.Text = m_QDThoiViec.TroCap_Khac.ToString
            txtTienBoiThuong.Text = m_QDThoiViec.SoTien_BoiThuong
            txtTienThuHoi.Text = m_QDThoiViec.SoTien_ThuHoi
            txtSoQD_TV.Text = m_QDThoiViec.So_QD
            txtDVraQD.Text = m_QDThoiViec.DVraQD
            If m_QDThoiViec.IsQD_NHCS Then
                cbIsQD_NHCS_TV.Checked = True
                txtCVNguoiKyQD_TV.Visible = False
                cboCVNguoiKyQD_TV.Visible = True
                cboCVNguoiKyQD_TV.SelectedValue = m_QDThoiViec.IdCV_NguoKy_QD
            Else
                cbIsQD_NHCS_TV.Checked = False
                txtCVNguoiKyQD_TV.Visible = True
                cboCVNguoiKyQD_TV.Visible = False
                txtCVNguoiKyQD_TV.Text = m_QDThoiViec.CV_NguoiKy_QD
            End If
            txtGhiChu_TV.Text = m_QDThoiViec.GhiChu
        Else
            blankQDThoiViec()
        End If
    End Sub

    Private Function updateQDThoiViec(ByVal vIdQDThoiViec As String) As Boolean
        Try
            Dim m_QDThoiViec As CBThoiViec = New CBThoiViec

            m_QDThoiViec.IdCBThoiViec = vIdQDThoiViec
            m_QDThoiViec.IdCanBo = idCanBo
            m_QDThoiViec.IdLoaiQD = CInt(cboQDThoiViec.SelectedValue)
            m_QDThoiViec.IdLyDo = CInt(cboLyDo_TV.SelectedValue)
            m_QDThoiViec.NgayKy_QD = DateTimeUtil.getDate(dpkNgayKy_TV.Text)
            m_QDThoiViec.Ngay_HL = DateTimeUtil.getDate(dpkNgayHL_TV.Text)
            m_QDThoiViec.NguoiKy_QD = standardizeName(txtNguoiKyQD_TV.Text)
            m_QDThoiViec.TroCap_ThoiViec = N2Number(MoneyValue(txtTroCap_TV.Text.ToString))
            m_QDThoiViec.TroCap_Khac = N2Number(MoneyValue(txtTroCapKhac.Text.ToString))
            m_QDThoiViec.SoTien_BoiThuong = N2Number(MoneyValue(txtTienBoiThuong.Text.ToString))
            m_QDThoiViec.SoTien_ThuHoi = N2Number(MoneyValue(txtTienThuHoi.Text.ToString))
            m_QDThoiViec.So_QD = txtSoQD_TV.Text.Trim
            m_QDThoiViec.DVraQD = standardizeString(txtDVraQD.Text.Trim)
            If cbIsQD_NHCS_TV.Checked Then
                m_QDThoiViec.IsQD_NHCS = 1
                m_QDThoiViec.IdCV_NguoKy_QD = CInt(cboCVNguoiKyQD_TV.SelectedValue)
                m_QDThoiViec.CV_NguoiKy_QD = ""
            Else
                m_QDThoiViec.IsQD_NHCS = 0
                m_QDThoiViec.IdCV_NguoKy_QD = 0
                m_QDThoiViec.CV_NguoiKy_QD = standardizeString(txtCVNguoiKyQD_TV.Text.Trim)
            End If
            m_QDThoiViec.GhiChu = standardizeString(txtGhiChu_TV.Text)
            If vIdQDThoiViec <> "" Then
                m_QDThoiViec.Update()
            Else
                idQDThoiViec = m_QDThoiViec.Add()
                If m_QDThoiViec.IdCBThoiViec = "" Then
                    Return False
                    Exit Function
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Sub cbIsQD_NHCS_TV_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbIsQD_NHCS_TV.CheckedChanged
        If cbIsQD_NHCS_TV.Checked Then
            txtCVNguoiKyQD_TV.Visible = False
            cboCVNguoiKyQD_TV.Visible = True
        Else
            cboCVNguoiKyQD_TV.Visible = False
            txtCVNguoiKyQD_TV.Visible = True
        End If
    End Sub

    Private Sub txtTroCap_TV_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTroCap_TV.TextChanged
        Try
            txtTroCap_TV = formatMoneyinTextbox(txtTroCap_TV)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTroCapKhac_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTroCapKhac.TextChanged
        Try
            txtTroCapKhac = formatMoneyinTextbox(txtTroCapKhac)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridThoiViec_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridThoiViec.CellClick
        Try
            idQDThoiViec = gridThoiViec.CurrentRow.Cells("IdCBThoiViec").Value.ToString
            fillQDThoiViec(idQDThoiViec)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridThoiViec_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridThoiViec.CellFormatting
        Try
            If gridThoiViec.Columns(e.ColumnIndex).Name = "Ngay_HL" Then
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
            If gridThoiViec.Columns(e.ColumnIndex).Name = "IdLoaiQD_TV" Then
                e.Value = getDanhmuc_Name(37, CInt(e.Value))
            End If
            If gridThoiViec.Columns(e.ColumnIndex).Name = "IdLyDo" Then
                e.Value = getDanhmuc_Name(22, CInt(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntNew_TV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew_TV.Click
        blankQDThoiViec()
    End Sub

    Private Sub bntUpdate_TV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate_TV.Click
        Try
            Dim lab_ErrTV As String = ""
            If (idQDThoiViec <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";141;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu quyết định thôi việc!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridThoiViec_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_ErrTV = checkTabQDThoiviec()
            If lab_ErrTV <> "" Then
                MessageBox.Show(lab_ErrTV, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateQDThoiViec(idQDThoiViec) Then
                bindGridTV(idCanBo)
                labAlert_TV.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_TV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel_TV.Click
        If idQDThoiViec = "" Then
            blankQDThoiViec()
        Else
            fillQDThoiViec(idQDThoiViec)
        End If
    End Sub

    Private Sub bntDelete_TV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete_TV.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridThoiViec.Rows.Count > 0) Then
                For i As Int32 = 0 To gridThoiViec.Rows.Count - 1
                    If (CType(gridThoiViec.Rows(i).Cells("cln_cbThoiViec").Value, Boolean) = True) Then
                        arrDel.Add(gridThoiViec.Rows(i).Cells("IdCBThoiViec").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idQDThoiViec <> "") Then arrDel.Add(idQDThoiViec.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        labAlert_TV.Text = ""
                        For i = 0 To arrDel.Count - 1
                            Dim m_QDThoiViec As CBThoiViec = New CBThoiViec
                            m_QDThoiViec.IdCBThoiViec = arrDel(i).ToString()
                            m_QDThoiViec.Delete()
                        Next
                        blankQDThoiViec()
                        bindGridTV(idCanBo)
                        labAlert_TV.Text = "Xoá dữ liệu thành công!"
                    Catch ex As Exception
                        MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Else
                MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntClose_TV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose_TV.Click
        Close()
    End Sub

    Private Sub gridThoiViec_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridThoiViec.KeyUp
        gridThoiViec_CellClick(sender, Nothing)
    End Sub

    Private Sub txtNguoiKyQD_TV_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNguoiKyQD_TV.LostFocus
        txtNguoiKyQD_TV.Text = standardizeName(txtNguoiKyQD_TV.Text)
    End Sub

    Private Sub txtGhiChu_TV_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGhiChu_TV.LostFocus
        txtGhiChu_TV.Text = standardizeString(txtGhiChu_TV.Text)
    End Sub

    Private Sub txtTienBoiThuong_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTienBoiThuong.TextChanged
        Try
            txtTienBoiThuong = formatMoneyinTextbox(txtTienBoiThuong)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTienThuHoi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTienThuHoi.TextChanged
        Try
            txtTienThuHoi = formatMoneyinTextbox(txtTienThuHoi)
        Catch ex As Exception
        End Try
    End Sub

#End Region

#Region "tab Quyết định Lương"

    Private Sub initTabQDLuong()
        cbIsQD_NHCS_L.Checked = True
        If Not initQDL Then
            cboLoaiQDLuong.DataSource = listLoaiQDLuong()
            cboCVNguoiKyQD_L.DataSource = listChucVuQuyenRaQD()
            cboQDNghiDinh.DataSource = listNghiDinhLuong()
            If cboQDNghiDinh.SelectedValue = 0 Then cboQDNghiDinh.SelectedValue = 1
            initQDL = True
        End If
        bindGridQDLuong(idCanBo)
        fillQDLuong(idQDLuong, idCanBo)
        cbIsQD_NHCS_L.Focus()
    End Sub

    Private Sub bindGridQDLuong(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridQDLuong.AutoGenerateColumns = False
                gridQDLuong.DataSource = listQDLuong(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridQDLuong.Rows.Count - 1
                    If Trim(gridQDLuong.Rows(i).Cells(0).Value) = idQDLuong Then
                        gridQDLuong.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách các quyết định Lương của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabQDLuong() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_L.Text = "" Then
                txtSoQD_L.Focus()
                strReturn = "Chưa nhập Số quyết định lương!"
                Exit Try
            End If
            If dpkNgayHL_L.Value.Date > Now.Date Then
                strReturn = "Ngày hiệu lực phải nhỏ hơn hoặc bằng ngày hiện tại. Hãy nhập lại!"
                dpkNgayHL_L.Focus()
                Exit Try
            End If
            'If idQDLuong = "" Then
            '    If checkQuyetDinh("LUONG", idCanBo, txtSoQD_L.Text) Then
            '        txtSoQD_L.Text = ""
            '        txtSoQD_L.Focus()
            '        strReturn = "Số quyết định lương của cán bộ đã tồn tại. Hãy nhập lại!"
            '        Exit Try
            '    End If
            'End If
            If txtNguoiKyQD_L.Text = "" Then
                txtNguoiKyQD_L.Focus()
                strReturn = "Chưa nhập Người quyết định !"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankQDLuong()
        idQDLuong = ""
        bntDelete_L.Enabled = False
        cbIsQD_NHCS_L.Checked = True
        txtSoQD_L.ReadOnly = False
        cboLoaiQDLuong.DataSource = listLoaiQDLuong()
        txtSoQD_L.Text = ""
        txtDVraQD_L.Text = getDonvi(IdDONVI)
        dpkNgayKy_L.Value = Date.Now
        txtNguoiKyQD_L.Text = GIAMDOC
        dpkNgayHL_L.Value = Date.Now
        dpkNgayLenLTT.Value = Date.Now
        txtGhiChu_L.Text = ""
        labAlert_L.Text = ""
        cbIsQD_NHCS_L.Focus()
    End Sub

    Private Sub fillQDLuong(ByRef vIdLuongCB As String, Optional ByVal vIdCanBo As String = "")
        Dim m_LuongCb As LuongCanBo = New LuongCanBo
        labAlert_L.Text = ""
        Dim vIdNghiDinhluong, vIdBangluong, vIdNgachluong As Integer
        If vIdLuongCB <> "" Then
            m_LuongCb = m_LuongCb.getRecord(vIdLuongCB)
        Else
            If vIdCanBo <> "" Then m_LuongCb = m_LuongCb.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";190;") < 0) Then  ' Xoá
            bntDelete_L.Enabled = False
        Else
            If m_LuongCb.IdCanBo <> "" Then
                If checkRight_CreateRecord(m_LuongCb.IdCanBo) Then
                    bntDelete_L.Enabled = True
                Else
                    bntDelete_L.Enabled = False
                End If
            End If
        End If
        If m_LuongCb.IdLuongCB <> "" Then
            vIdLuongCB = m_LuongCb.IdLuongCB
            txtSoQD_L.Text = m_LuongCb.SoQD
            txtDVraQD_L.Text = m_LuongCb.DVraQD
            dpkNgayKy_L.Value = m_LuongCb.NgayQD
            txtNguoiKyQD_L.Text = m_LuongCb.NguoiQD
            dpkNgayHL_L.Value = m_LuongCb.Ngay_Huong
            dpkNgayLenLTT.Value = m_LuongCb.NgayLen_DK
            txtGhiChu_L.Text = m_LuongCb.GhiChu
            If m_LuongCb.IsQD_NHCS Then
                cbIsQD_NHCS_L.Checked = True
                txtCVNguoiKyQD_L.Visible = False
                cboCVNguoiKyQD_L.Visible = True
                txtLoaiQDLuong.Visible = False
                cboLoaiQDLuong.Visible = True
                pnlQDL_IsNotNHCS.Visible = False
                pnlQDL_IsNHCS.Visible = True
                cboCVNguoiKyQD_L.SelectedValue = m_LuongCb.IdCV_Nguoi_QD
                If checkQDLuong_ReadOnly(m_LuongCb.IdLoaiQD) Then
                    cboLoaiQDLuong.DataSource = listDanhmuc(43, False, False, m_LuongCb.IdLoaiQD)
                Else
                    cboLoaiQDLuong.DataSource = listLoaiQDLuong()
                End If
                cboLoaiQDLuong.SelectedValue = m_LuongCb.IdLoaiQD
                getIdNDBangNgach(m_LuongCb.IdBacLuong, vIdNghiDinhluong, vIdBangluong, vIdNgachluong)
                cboQDNghiDinh.SelectedValue = vIdNghiDinhluong
                cboQDBang.SelectedValue = vIdBangluong
                cboQDNgach.SelectedValue = vIdNgachluong
                cboQDBac.SelectedValue = m_LuongCb.IdBacLuong
                txtQDHeso.Text = m_LuongCb.HeSoLuong
            Else
                cbIsQD_NHCS_L.Checked = False
                cboCVNguoiKyQD_L.Visible = False
                txtCVNguoiKyQD_L.Visible = True
                cboLoaiQDLuong.Visible = False
                txtLoaiQDLuong.Visible = True
                pnlQDL_IsNHCS.Visible = False
                pnlQDL_IsNotNHCS.Visible = True
                pnlQDL_IsNotNHCS.Enabled = True
                If m_LuongCb.IdLoaiQD > 0 Then
                    txtLoaiQDLuong.Text = getDanhmuc_Name(43, m_LuongCb.IdLoaiQD)
                    txtLoaiQDLuong.ReadOnly = True
                Else
                    txtLoaiQDLuong.ReadOnly = False
                    txtLoaiQDLuong.Text = m_LuongCb.LoaiQD
                End If
                txtCVNguoiKyQD_L.Text = m_LuongCb.CV_NguoiKy_QD
                txtNoiDungQD_L.Text = m_LuongCb.NoiDung
            End If
            If checkQDLuong_ReadOnly(m_LuongCb.IdLoaiQD) Then
                txtSoQD_L.ReadOnly = True
            Else
                txtSoQD_L.ReadOnly = False
            End If
        Else
            blankQDLuong()
        End If
    End Sub

    Private Function updateQDLuong(ByVal vIdLuongCB As String) As Boolean
        Try
            Dim m_LuongCB As LuongCanBo = New LuongCanBo
            m_LuongCB.IdLuongCB = vIdLuongCB
            m_LuongCB.IdCanBo = idCanBo
            m_LuongCB.Ngay_Huong = DateTimeUtil.getDate(dpkNgayHL_L.Text)
            m_LuongCB.NgayLen_DK = DateTimeUtil.getDate(dpkNgayLenLTT.Text)
            m_LuongCB.SoQD = txtSoQD_L.Text
            m_LuongCB.DVraQD = standardizeString(txtDVraQD_L.Text.Trim)
            m_LuongCB.NgayQD = DateTimeUtil.getDate(dpkNgayKy_L.Text)
            m_LuongCB.NguoiQD = standardizeName(txtNguoiKyQD_L.Text)
            m_LuongCB.GhiChu = standardizeString(txtGhiChu_L.Text)
            If cbIsQD_NHCS_L.Checked Then
                m_LuongCB.IsQD_NHCS = 1
                m_LuongCB.IdCV_Nguoi_QD = CInt(cboCVNguoiKyQD_L.SelectedValue)
                m_LuongCB.CV_NguoiKy_QD = ""
                m_LuongCB.IdLoaiQD = CInt(cboLoaiQDLuong.SelectedValue)
                m_LuongCB.LoaiQD = ""
                m_LuongCB.IdBacLuong = CInt(cboQDBac.SelectedValue)
                m_LuongCB.HeSoLuong = CDbl(txtQDHeso.Text)
                m_LuongCB.NoiDung = ""
            Else
                m_LuongCB.IsQD_NHCS = 0
                m_LuongCB.IdBacLuong = 0
                m_LuongCB.HeSoLuong = 0
                m_LuongCB.IdCV_Nguoi_QD = 0
                m_LuongCB.CV_NguoiKy_QD = standardizeString(txtCVNguoiKyQD_L.Text.Trim)
                m_LuongCB.IdLoaiQD = 0
                m_LuongCB.LoaiQD = standardizeString(txtLoaiQDLuong.Text.Trim)
                m_LuongCB.NoiDung = standardizeString(txtNoiDungQD_L.Text.Trim)
            End If
            If idQDLuong <> "" Then
                m_LuongCB.Update()
            Else
                m_LuongCB.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub cbIsQD_NHCS_L_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbIsQD_NHCS_L.CheckedChanged
        If cbIsQD_NHCS_L.Checked Then
            txtCVNguoiKyQD_L.Visible = False
            cboCVNguoiKyQD_L.Visible = True
            txtLoaiQDLuong.Visible = False
            cboLoaiQDLuong.Visible = True
            pnlQDL_IsNotNHCS.Visible = False
            pnlQDL_IsNHCS.Visible = True
            If idQDLuong = "" Then txtDVraQD_L.Text = getDonvi(IdDONVI)
        Else
            txtCVNguoiKyQD_L.Visible = True
            cboCVNguoiKyQD_L.Visible = False
            cboLoaiQDLuong.Visible = False
            txtLoaiQDLuong.Visible = True
            txtLoaiQDLuong.ReadOnly = False
            txtLoaiQDLuong.Enabled = True
            txtLoaiQDLuong.Text = ""
            pnlQDL_IsNHCS.Visible = False
            pnlQDL_IsNotNHCS.Visible = True
            pnlQDL_IsNotNHCS.Enabled = True
        End If
    End Sub

    Private Sub txtNguoiKyQD_L_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNguoiKyQD_L.LostFocus
        txtNguoiKyQD_L.Text = standardizeName(txtNguoiKyQD_L.Text)
    End Sub

    Private Sub cboQDNghiDinh_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQDNghiDinh.SelectedValueChanged
        If Not (cboQDNghiDinh.SelectedValue Is Nothing) Then
            cboQDBang.DataSource = listBangLuong(CInt(cboQDNghiDinh.SelectedValue))
        End If
    End Sub

    Private Sub cboQDBang_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQDBang.SelectedValueChanged
        If Not (cboQDBang.SelectedValue Is Nothing) Then
            cboQDNgach.DataSource = listNgachLuong(CInt(cboQDBang.SelectedValue))
        End If
    End Sub

    Private Sub cboQDNgach_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQDNgach.SelectedValueChanged
        If Not (cboQDNgach.SelectedValue Is Nothing) Then
            cboQDBac.DataSource = listBacLuong(CInt(cboQDNgach.SelectedValue))
        End If
    End Sub

    Private Sub cboQDBac_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQDBac.SelectedValueChanged
        If Not (cboQDBac.SelectedValue Is Nothing) Then
            Try
                txtQDHeso.Text = getHeso(CInt(cboQDBac.SelectedValue))
            Catch ex As Exception
                txtQDHeso.Text = ""
            End Try
        End If

    End Sub

    Private Sub dpkNgayHL_L_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dpkNgayHL_L.LostFocus
        dpkNgayLenLTT.Value = dpkNgayHL_L.Value.AddMonths(getTimeNangBac(CInt(cboQDNgach.SelectedValue)))
    End Sub

    Private Sub gridQDLuong_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridQDLuong.CellClick
        Try
            idQDLuong = gridQDLuong.CurrentRow.Cells("IdLCB").Value.ToString
            fillQDLuong(idQDLuong)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDLuong_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridQDLuong.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridQDLuong.Columns("NgayHL_L").Index Then
                e.FormattingApplied = True
                row = gridQDLuong.Rows(e.RowIndex)
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDLuong_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridQDLuong.KeyUp
        gridQDLuong_CellClick(sender, Nothing)
    End Sub

    Private Sub bntNew_L_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntNew_L.Click
        blankQDLuong()
    End Sub

    Private Sub bntUpdate_L_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUpdate_L.Click
        Try
            Dim lab_Err1 As String = ""
            If (idQDLuong <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";189;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu lương chuyên môn nghiệp vụ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridQdLuong_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err1 = checkTabQDLuong()
            If lab_Err1 <> "" Then
                MessageBox.Show(lab_Err1, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If

            If updateQDLuong(idQDLuong) Then
                bindGridQDLuong(idCanBo)
                labAlert_L.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_L_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntCancel_L.Click
        If idQDLuong = "" Then
            blankQDLuong()
        Else
            fillQDLuong(idQDLuong)
        End If
    End Sub

    Private Sub bntDelete_L_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntDelete_L.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridQDLuong.Rows.Count > 0) Then
                For i As Int32 = 0 To gridQDLuong.Rows.Count - 1
                    If (CType(gridQDLuong.Rows(i).Cells("cln_cbL").Value, Boolean) = True) Then
                        arrDel.Add(gridQDLuong.Rows(i).Cells("IdLCB").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idQDLuong <> "") Then arrDel.Add(idQDLuong.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        labAlert_L.Text = ""
                        For i = 0 To arrDel.Count - 1
                            Dim m_LuongCanBo As LuongCanBo = New LuongCanBo
                            m_LuongCanBo.IdLuongCB = arrDel(i).ToString()
                            m_LuongCanBo.Delete()
                        Next
                        blankQDLuong()
                        bindGridQDLuong(idCanBo)
                        labAlert_L.Text = "Xoá dữ liệu thành công!"
                    Catch ex As Exception
                        MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Else
                MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntClose_L_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClose_L.Click
        Close()
    End Sub

#End Region

#Region "tab Quyết định Phu cap"
    Private Sub initTabQDPhuCap()
        cbIsQD_NHCS_PC.Checked = True
        If Not initQDP Then
            cboCVNguoiKyQD_PC.DataSource = listChucVuQuyenRaQD()
            cboLoaiQD_PC.DataSource = listDanhmuc(31)
            initQDP = True
        End If
        bindGridQDPhucap(idCanBo)
        fillQDPhucap(idQDPhucap, idCanBo)
        cbIsQD_NHCS_PC.Focus()
    End Sub

    Private Sub bindGridQDPhucap(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridQDPhuCap.AutoGenerateColumns = False
                gridQDPhuCap.DataSource = listQDPhucap(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridQDPhuCap.Rows.Count - 1
                    If Trim(gridQDPhuCap.Rows(i).Cells(0).Value) = idQDPhucap Then
                        gridQDPhuCap.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Phụ cấp của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabQDPhucap() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_PC.Text = "" Then
                txtSoQD_PC.Focus()
                strReturn = "Chưa nhập Số quyết định !"
                Exit Try
            End If
            If dpkNgayHL_PC.Value.Date > Now.Date Then
                strReturn = "Ngày hiệu lực phải nhỏ hơn hoặc bằng ngày hiện tại. Hãy nhập lại!"
                dpkNgayHL_PC.Focus()
                Exit Try
            End If
            'If idQDPhucap = "" Then
            '    If checkQuyetDinh("PHUCAP", idCanBo, txtSoQD_PC.Text) Then
            '        txtSoQD_PC.Text = ""
            '        txtSoQD_PC.Focus()
            '        strReturn = "Số quyết định phụ cấp của cán bộ đã tồn tại. Hãy nhập lại!"
            '        Exit Try
            '    End If
            'End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankQDPhucap()
        idQDPhucap = ""
        bntDelete_PC.Enabled = False
        cbIsQD_NHCS_PC.Checked = True
        txtSoQD_PC.Text = ""
        txtDVraQD_PC.Text = getDonvi(IdDONVI)
        dpkNgayKy_PC.Value = Date.Now
        txtNguoiKyQD_PC.Text = GIAMDOC
        dpkNgayHL_PC.Value = Date.Now
        dpkDenNgay_PC.Checked = False
        txtGhichu_PC.Text = ""
        labAlert_P.Text = ""
        NgayHLCU = Nothing
        SoQDCU = ""
        NgayQDCU = Nothing
        cbIsQD_NHCS_PC.Focus()
    End Sub

    Private Sub fillQDPhucap(ByRef vIdPhucapCB As String, Optional ByVal vIdCanBo As String = "")
        Dim m_PhucapCB As Phucap = New Phucap
        labAlert_P.Text = ""
        Dim vIdLoaiPC As Integer
        If vIdPhucapCB <> "" Then
            m_PhucapCB = m_PhucapCB.getRecord(vIdPhucapCB)
        Else
            If vIdCanBo <> "" Then m_PhucapCB = m_PhucapCB.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";194;") < 0) Then  ' Xoá
            bntDelete_PC.Enabled = False
        Else
            If m_PhucapCB.IdCanBo <> "" Then
                If checkRight_CreateRecord(m_PhucapCB.IdCanBo) Then
                    bntDelete_PC.Enabled = True
                Else
                    bntDelete_PC.Enabled = False
                End If
            End If
        End If
        If m_PhucapCB.IdCB_PhuCap <> "" Then
            vIdPhucapCB = m_PhucapCB.IdCB_PhuCap
            txtSoQD_PC.Text = m_PhucapCB.SoQD
            dpkNgayKy_PC.Value = m_PhucapCB.NgayQD
            txtNguoiKyQD_PC.Text = m_PhucapCB.NguoiQD
            dpkNgayHL_PC.Value = m_PhucapCB.TuNgay
            If m_PhucapCB.DenNgay = DateTime.MinValue Then
                dpkDenNgay_PC.Checked = False
            Else
                dpkDenNgay_PC.Checked = True
                dpkDenNgay_PC.Value = m_PhucapCB.DenNgay
            End If
            vIdLoaiPC = getIdLoaiPC(m_PhucapCB.IdMucPC)
            txtDVraQD_PC.Text = m_PhucapCB.DVraQD
            If m_PhucapCB.IsQD_NHCS Then
                cbIsQD_NHCS_PC.Checked = True
                txtCVNguoiKyQD_PC.Visible = False
                cboCVNguoiKyQD_PC.Visible = True
                pnlQDPC_IsNotNHCS.Visible = False
                pnlQDPC_IsNHCS.Visible = True
                cboCVNguoiKyQD_PC.SelectedValue = m_PhucapCB.IdCV_Nguoi_QD
                cboLoaiQD_PC.SelectedValue = vIdLoaiPC
                cboMucQD_PC.SelectedValue = m_PhucapCB.IdMucPC
            Else
                cbIsQD_NHCS_PC.Checked = False
                cboCVNguoiKyQD_PC.Visible = False
                txtCVNguoiKyQD_PC.Visible = True
                pnlQDPC_IsNHCS.Visible = False
                pnlQDPC_IsNotNHCS.Visible = True
                pnlQDPC_IsNotNHCS.Enabled = True
                txtNoiDungQD_PC.Enabled = True
                txtCVNguoiKyQD_PC.Text = m_PhucapCB.CV_NguoiKy_QD
                txtNoiDungQD_PC.Text = m_PhucapCB.NoiDung
            End If
            txtGhichu_PC.Text = m_PhucapCB.GhiChu

            NgayHLCU = m_PhucapCB.TuNgay
            SoQDCU = m_PhucapCB.SoQD
            NgayQDCU = m_PhucapCB.NgayQD
        Else
            blankQDPhucap()
        End If
    End Sub

    Private Function updateQDPhucap(ByVal vIdPhucap As String) As Boolean
        Try
            Dim m_PhucapCB As Phucap = New Phucap
            Dim dtPC As DataTable
            Dim iPC As Integer = 0
            Dim db As DBAccess = New DBAccess

            m_PhucapCB.IdCB_PhuCap = vIdPhucap
            m_PhucapCB.IdCanBo = idCanBo

            m_PhucapCB.TuNgay = DateTimeUtil.getDate(dpkNgayHL_PC.Text)
            If dpkDenNgay_PC.Checked Then
                m_PhucapCB.DenNgay = DateTimeUtil.getDate(dpkDenNgay_PC.Text)
            Else
                m_PhucapCB.DenNgay = DateTime.MinValue
            End If
            m_PhucapCB.SoQD = txtSoQD_PC.Text
            m_PhucapCB.NgayQD = DateTimeUtil.getDate(dpkNgayKy_PC.Text)
            m_PhucapCB.NguoiQD = standardizeName(txtNguoiKyQD_PC.Text)
            m_PhucapCB.DVraQD = standardizeString(txtDVraQD_PC.Text.Trim)
            If cbIsQD_NHCS_PC.Checked Then
                m_PhucapCB.IsQD_NHCS = 1
                m_PhucapCB.IdCV_Nguoi_QD = CInt(cboCVNguoiKyQD_PC.SelectedValue)
                m_PhucapCB.CV_NguoiKy_QD = ""
                m_PhucapCB.IdMucPC = CInt(cboMucQD_PC.SelectedValue)
                m_PhucapCB.NoiDung = ""
            Else
                cbIsQD_NHCS_PC.Checked = False
                m_PhucapCB.IdCV_Nguoi_QD = 0
                m_PhucapCB.CV_NguoiKy_QD = standardizeString(txtCVNguoiKyQD_PC.Text.Trim)
                m_PhucapCB.IdMucPC = 0
                m_PhucapCB.NoiDung = standardizeString(txtNoiDungQD_PC.Text)
            End If
            m_PhucapCB.GhiChu = standardizeString(txtGhichu_PC.Text)
            If idQDPhucap <> "" Then
                dtPC = db.SelectDBRows("SELECT IdCB_PhuCap FROM HS_PhucapCB WHERE IdCanbo='" & idCanBo & "' and Datediff(day,TuNgay,'" & NgayHLCU & "')=0 and Upper(SoQD)=Upper(N'" & SoQDCU & "') and NgayQD='" & NgayQDCU & "' and IdMucPC<>" & m_PhucapCB.IdMucPC)
                m_PhucapCB.Update()
                ' Update tiếp với các Phụ cấp khác có cùng số QĐ và ngày
                If dtPC.Rows.Count > 0 Then
                    For iPC = 0 To dtPC.Rows.Count - 1
                        m_PhucapCB.IdCB_PhuCap = dtPC.Rows(iPC).Item("IdCB_PhuCap")
                        m_PhucapCB.UpdateNotFull()
                    Next
                End If
            Else
                m_PhucapCB.Add()
                dtPC = db.SelectDBRows("SELECT IdCB_PhuCap, DenNgay FROM HS_PhucapCB t1, MucPhuCap t2 WHERE t1.IdMucPC=t2.IdMuc_PhC and idcanbo='" & idCanBo & "' and IsQD_NHCS=1 and IdLoai_PhC=" & CInt(cboLoaiQD_PC.SelectedValue) & " order by TuNgay desc")
                If dtPC.Rows.Count >= 2 Then
                    If dtPC.Rows(1).Item("DenNgay") Is DBNull.Value Then
                        db.executeSQL("UPDATE HS_PhuCapCB Set DenNgay='" & m_PhucapCB.TuNgay.AddDays(-1) & "' WHERE IdCB_PhuCap= '" & dtPC.Rows(1).Item("IdCB_PhuCap") & "'")
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub txtNguoiKyQD_PC_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNguoiKyQD_PC.LostFocus
        txtNguoiKyQD_PC.Text = standardizeName(txtNguoiKyQD_PC.Text)
    End Sub

    Private Sub cboLoaiQD_PC_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLoaiQD_PC.SelectedValueChanged
        If Not (cboLoaiQD_PC.SelectedValue Is Nothing) Then
            cboMucQD_PC.DataSource = listMucPC(CInt(cboLoaiQD_PC.SelectedValue))
        End If
    End Sub

    Private Sub gridQDPhuCap_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridQDPhuCap.CellClick
        Try
            idQDPhucap = gridQDPhuCap.CurrentRow.Cells("IdPC").Value.ToString
            fillQDPhucap(idQDPhucap)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDPhuCap_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridQDPhuCap.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridQDPhuCap.Columns("NgayHL_PC").Index Then
                e.FormattingApplied = True
                row = gridQDPhuCap.Rows(e.RowIndex)
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDPhuCap_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridQDPhuCap.KeyUp
        gridQDPhuCap_CellClick(sender, Nothing)
    End Sub

    Private Sub cbIsQD_NHCS_PC_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbIsQD_NHCS_PC.CheckedChanged
        If cbIsQD_NHCS_PC.Checked Then
            txtCVNguoiKyQD_PC.Visible = False
            cboCVNguoiKyQD_PC.Visible = True
            pnlQDPC_IsNotNHCS.Visible = False
            pnlQDPC_IsNHCS.Visible = True
            If idQDPhucap = "" Then txtDVraQD_PC.Text = getDonvi(IdDONVI)
        Else
            txtCVNguoiKyQD_PC.Visible = True
            cboCVNguoiKyQD_PC.Visible = False
            pnlQDPC_IsNHCS.Visible = False
            pnlQDPC_IsNotNHCS.Visible = True
        End If
    End Sub

    Private Sub bntNew_PC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntNew_PC.Click
        blankQDPhucap()
    End Sub

    Private Sub bntUpdate_PC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUpdate_PC.Click
        Try
            Dim lab_Err2 As String = ""
            If (idQDPhucap <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";193;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu về phụ cấp lương!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridQDPhuCap_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err2 = checkTabQDPhucap()
            If lab_Err2 <> "" Then
                MessageBox.Show(lab_Err2, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            'If checkQuyetDinh("PHUCAP", idCanBo, txtSoQD_PC.Text) Then
            '    If MessageBox.Show("Số quyết định phụ cấp của cán bộ đã tồn tại." & vbCr & "Bạn muốn nhập tiếp thông tin cho số quyết định này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
            '        If updateQDPhucap(idQDPhucap) Then
            '            bindGridQDPhucap(idCanBo)
            '            labAlert_P.Text = "Ghi dữ liệu thành công!"
            '        Else
            '            MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            '        End If
            '    End If
            'Else
            If updateQDPhucap(idQDPhucap) Then
                bindGridQDPhucap(idCanBo)
                labAlert_P.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_PC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntCancel_PC.Click
        If idQDPhucap = "" Then
            blankQDPhucap()
        Else
            fillQDPhucap(idQDPhucap)
        End If
    End Sub

    Private Sub bntDelete_PC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntDelete_PC.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridQDPhuCap.Rows.Count > 0) Then
                For i As Int32 = 0 To gridQDPhuCap.Rows.Count - 1
                    If (CType(gridQDPhuCap.Rows(i).Cells("cln_cbPC").Value, Boolean) = True) Then
                        arrDel.Add(gridQDPhuCap.Rows(i).Cells("IdPC").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idQDPhucap <> "") Then arrDel.Add(idQDPhucap.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        labAlert_P.Text = ""
                        For i = 0 To arrDel.Count - 1
                            Dim m_Phucap As Phucap = New Phucap
                            m_Phucap.IdCB_PhuCap = arrDel(i).ToString()
                            m_Phucap.Delete()
                        Next
                        blankQDPhucap()
                        bindGridQDPhucap(idCanBo)
                        labAlert_P.Text = "Xoá dữ liệu thành công!"
                    Catch ex As Exception
                        MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Else
                MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntClose_PC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClose_PC.Click
        Close()
    End Sub

#End Region

#Region "tab Quyết định Khac"

    Private Sub initTabQDKhac()
        bindGridQDKhac(idCanBo)
        fillQDKhac(idQDKhac, idCanBo)
        txtSoQD_K.Focus()
    End Sub

    Private Sub bindGridQDKhac(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridQDKhac.AutoGenerateColumns = False
                gridQDKhac.DataSource = listQDKhac(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridQDKhac.Rows.Count - 1
                    If Trim(gridQDKhac.Rows(i).Cells(0).Value) = idQDKhac Then
                        gridQDKhac.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Phụ cấp của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabQDKhac() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_K.Text = "" Then
                txtSoQD_K.Focus()
                strReturn = "Chưa nhập Số quyết định !"
                Exit Try
            End If
            'If idQDKhac = "" Then
            '    If checkQuyetDinh("QDKHAC", idCanBo, txtSoQD_K.Text) Then
            '        txtSoQD_K.Text = ""
            '        txtSoQD_K.Focus()
            '        strReturn = "Số quyết định khác của cán bộ đã tồn tại. Hãy nhập lại!"
            '        Exit Try
            '    End If
            'End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankQDKhac()
        idQDKhac = ""
        bntDelete_K.Enabled = False
        txtSoQD_K.Text = ""
        dpkNgayKy_K.Value = Date.Now
        txtNguoiKyQD_K.Text = ""
        txtNoiDungQD_K.Text = ""
        txtGhiChu_K.Text = ""
        labAlert_K.Text = ""
        txtSoQD_K.Focus()
    End Sub

    Private Sub fillQDKhac(ByRef vIdQDKhacCB As String, Optional ByVal vIdCanBo As String = "")
        Dim m_QDKhacCB As QDKhac = New QDKhac
        labAlert_K.Text = ""
        If vIdQDKhacCB <> "" Then
            m_QDKhacCB = m_QDKhacCB.getRecord(vIdQDKhacCB)
        Else
            If vIdCanBo <> "" Then m_QDKhacCB = m_QDKhacCB.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";102;") < 0) Then  ' Xoá
            bntDelete_K.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDelete_K.Enabled = True
                Else
                    bntDelete_K.Enabled = False
                End If
            End If
        End If
        If m_QDKhacCB.IdQDKhac <> "" Then
            vIdQDKhacCB = m_QDKhacCB.IdQDKhac
            txtSoQD_K.Text = m_QDKhacCB.So_QD
            dpkNgayKy_K.Value = m_QDKhacCB.NgayKy_QD
            txtCVNguoiKyQD_K.Text = m_QDKhacCB.CV_Nguoiky_QD
            txtNguoiKyQD_K.Text = m_QDKhacCB.NguoiKy_QD
            txtDVraQD_K.Text = m_QDKhacCB.DVraQD
            txtNoiDungQD_K.Text = m_QDKhacCB.NoiDung
            txtGhiChu_K.Text = m_QDKhacCB.GhiChu
        Else
            blankQDPhucap()
        End If
    End Sub

    Private Function updateQDKhac(ByVal vIdQDKhacCB As String) As Boolean
        Try
            Dim m_QDKhacCB As QDKhac = New QDKhac
            m_QDKhacCB.IdQDKhac = vIdQDKhacCB
            m_QDKhacCB.IdCanBo = idCanBo
            m_QDKhacCB.So_QD = txtSoQD_K.Text
            m_QDKhacCB.NgayKy_QD = DateTimeUtil.getDate(dpkNgayKy_K.Text)
            m_QDKhacCB.NguoiKy_QD = standardizeName(txtNguoiKyQD_K.Text)
            m_QDKhacCB.CV_Nguoiky_QD = standardizeName(txtCVNguoiKyQD_K.Text)
            m_QDKhacCB.DVraQD = txtDVraQD_K.Text.Trim
            m_QDKhacCB.NoiDung = standardizeString(txtNoiDungQD_K.Text)
            m_QDKhacCB.GhiChu = standardizeString(txtGhiChu_K.Text)
            If vIdQDKhacCB <> "" Then
                m_QDKhacCB.Update()
            Else
                m_QDKhacCB.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

    Private Sub txtNguoiKyQD_K_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNguoiKyQD_K.LostFocus
        txtNguoiKyQD_K.Text = standardizeName(txtNguoiKyQD_K.Text)
    End Sub

    Private Sub gridQDKhac_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridQDKhac.CellClick
        Try
            idQDKhac = gridQDKhac.CurrentRow.Cells("IdQDK").Value.ToString
            fillQDKhac(idQDKhac)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDKhac_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridQDKhac.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridQDKhac.Columns("NgayHL_K").Index Then
                e.FormattingApplied = True
                row = gridQDKhac.Rows(e.RowIndex)
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridQDKhac_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridQDKhac.KeyUp
        gridQDKhac_CellClick(sender, Nothing)
    End Sub

    Private Sub bntNew_K_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntNew_K.Click
        blankQDKhac()
    End Sub

    Private Sub bntUpdate_K_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUpdate_K.Click
        Try
            Dim lab_Err1 As String = ""
            If (idQDKhac <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";101;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu quyết định nhân sự khác của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridQDKhac_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err1 = checkTabQDKhac()
            If lab_Err1 <> "" Then
                MessageBox.Show(lab_Err1, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If

            If updateQDKhac(idQDKhac) Then
                bindGridQDKhac(idCanBo)
                labAlert_K.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_K_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntCancel_K.Click
        If idQDKhac = "" Then
            blankQDKhac()
        Else
            fillQDKhac(idQDKhac)
        End If
    End Sub

    Private Sub bntDelete_K_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntDelete_K.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridQDKhac.Rows.Count > 0) Then
                For i As Int32 = 0 To gridQDKhac.Rows.Count - 1
                    If (CType(gridQDKhac.Rows(i).Cells("cln_cbK").Value, Boolean) = True) Then
                        arrDel.Add(gridQDKhac.Rows(i).Cells("IdQDK").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idQDLuong <> "") Then arrDel.Add(idQDLuong.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        labAlert_K.Text = ""
                        For i = 0 To arrDel.Count - 1
                            Dim m_QDKhacCB As QDKhac = New QDKhac
                            m_QDKhacCB.IdQDKhac = arrDel(i).ToString()
                            m_QDKhacCB.Delete()
                        Next
                        blankQDKhac()
                        bindGridQDKhac(idCanBo)
                        labAlert_K.Text = "Xoá dữ liệu thành công!"
                    Catch ex As Exception
                        MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Else
                MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntClose_K_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClose_K.Click
        Close()
    End Sub

    Private Sub bntExportHSCB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExportHSCB.Click
        Dim frmEx As frmExportHSCB = New frmExportHSCB
        Dim QDNhanSufinalNHCS As QDNhanSu = New QDNhanSu
        Dim IDCN As Integer
        QDNhanSufinalNHCS = QDNhanSufinalNHCS.getFinalRecord(idCanBo, Nothing, True)
        frmEx.IDCB_Moved = QDNhanSufinalNHCS.IdCanBo
        IDCN = dbconn.getString("SELECT id FROM ChiNhanh WHERE (id=" & QDNhanSufinalNHCS.IdDonvi_Moi & " and id_goc=1) or (id=" & QDNhanSufinalNHCS.IdDonvi_Moi & " and id_goc=0) or (id in (SELECT id_goc FROM Chinhanh WHERE id=" & QDNhanSufinalNHCS.IdDonvi_Moi & ") and id_goc=1)")
        frmEx.IDCN_Moved = IDCN
        frmEx.TG_Chuyen = QDNhanSufinalNHCS.NgayHL.Date
        frmEx.ShowDialog()
    End Sub

    Private Sub chkTWQuanLy_CheckedChanged(sender As Object, e As EventArgs) Handles chkTWQuanLy.CheckedChanged

    End Sub

End Class