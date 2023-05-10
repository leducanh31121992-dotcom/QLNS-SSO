Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmLuong
    Inherits System.Windows.Forms.Form
    Public IDCB_HS_Canbo As String = ""   ' Id Cán bộ truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDDV_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDPB_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Private dbconn As DBAccess
    Private idCanBo As String = ""
    Private initQDL As Boolean = False
    Private initQDP As Boolean = False
    Private initTG As Boolean = False
    Private idQDLuong As String = ""
    Private idQDPhucap As String = ""
    Private idThemgio As String = ""
    Private idDVTT As Integer = 0
    Private idPTT As Integer = 0
    Private TS As Boolean = False
    Private SoQDCU As String = ""
    Private NgayHLCU As Date
    Private NgayQDCU As Date
    Private tpl As New List(Of TabPage)

    Private Sub frmLuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmLuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        lkl_xemct.Enabled = False
        'Check quyền thành viên
        Dim roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";187;188;189;190;", roles)) Then
            If TabControlLuong.TabPages.Contains(tabLuong) Then TabControlLuong.TabPages.Remove(tabLuong)
        End If
        If Not (Globals.IsIntersect(";191;192;193;194;", roles)) Then
            If TabControlLuong.TabPages.Contains(tabPhuCap) Then TabControlLuong.TabPages.Remove(tabPhuCap)
        End If
        If Not (Globals.IsIntersect(";195;196;197;198;", roles)) Then
            If TabControlLuong.TabPages.Contains(tabThemGio) Then TabControlLuong.TabPages.Remove(tabThemGio)
        End If
        If Not (Globals.IsIntersect(";236;237;", roles)) Then
            If TabControlLuong.TabPages.Contains(tabTK) Then TabControlLuong.TabPages.Remove(tabTK)
        End If

        If (Globals.Roles.IndexOf(";187;") < 0) Then   ' xem
            gridQDLuong.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";188;") < 0) Then  ' Them
            bntNew_L.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";190;") < 0) Then  ' Xoá
            bntDelete_L.Enabled = False
        End If

        If (Globals.Roles.IndexOf(";191;") < 0) Then
            gridQDPhuCap.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";192;") < 0) Then
            bntNew_PC.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";194;") < 0) Then
            bntDelete_PC.Enabled = False
        End If

        If (Globals.Roles.IndexOf(";195;") < 0) Then
            gridThemgio.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";196;") < 0) Then
            bntNew_TG.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";198;") < 0) Then
            bntDelete_TG.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";237;") < 0) Then
            bntUpdate_TaiKhoan.Enabled = False
        End If
        If IDCB_HS_Canbo = "" Then
            bindTreeview(treeCocau)
            treeCocau.ExpandAll()
            callTab()
        Else
            'Gọi từ Hồ sơ cán bộ qua menu ngữ cảnh
            Dim maCB As String = getCanBo_Ma(IDCB_HS_Canbo, False)
            Dim m_QDNS As QDNhanSu = New QDNhanSu
            Dim _NodeFind As TreeNode = New TreeNode
            Dim _node As TreeNode = New TreeNode

            m_QDNS = m_QDNS.getFinalRecord(IDCB_HS_Canbo, Nothing, True, True)
            _NodeFind.Name = IDCB_HS_Canbo
            _NodeFind.Tag = "CB_" & IDCB_HS_Canbo & "_" & maCB & "_" & m_QDNS.IdDonvi_Moi & "_" & m_QDNS.IdPhong_Moi

            bindTreeview(treeCocau, False, IDPB_HS_Canbo, IDDV_HS_Canbo)
            treeCocau.ExpandAll()

            _node = findNode(treeCocau.Nodes, _NodeFind.Tag)
            If Not _node Is Nothing Then
                treeCocau.SelectedNode = _node
                idCanBo = IDCB_HS_Canbo
                OverInfCB(idCanBo, maCB)
                initTabQDLuong()
            End If

            If checkRight_CreateRecord(IDCB_HS_Canbo) Then
                bntUpdate_L.Enabled = True
                bntUpdate_PC.Enabled = True
                bntUpdate_TG.Enabled = True
                bntUpdate_TaiKhoan.Enabled = True
                bntDelete_L.Enabled = True
                bntDelete_PC.Enabled = True
                bntDelete_TG.Enabled = True
            Else
                bntUpdate_L.Enabled = False
                bntUpdate_PC.Enabled = False
                bntUpdate_TG.Enabled = False
                bntUpdate_TaiKhoan.Enabled = False
                bntDelete_L.Enabled = False
                bntDelete_PC.Enabled = False
                bntDelete_TG.Enabled = False
            End If
        End If
    End Sub

    Private Sub treeCocau_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeCocau.AfterSelect
        Dim arr() As String
        arr = treeCocau.SelectedNode.Tag.ToString.Split("_")
        labAlert_L.Text = ""
        labAlert_P.Text = ""
        labAlert_T.Text = ""
        blankQDLuong()
        Select Case arr(0)
            Case "DV"
                blankOverviewInfCB()
                blankQDLuong()
                blankQDPhucap()
                blankLuongThemgio()
                If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not treeCocau.SelectedNode.IsExpanded Then
                        getNode(treeCocau.SelectedNode, arr(1), arr(2), True)
                    End If
                End If
                treeCocau.SelectedNode.Expand()
            Case "PB"
                blankOverviewInfCB()
                blankQDLuong()
                blankQDPhucap()
                blankLuongThemgio()
                If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not treeCocau.SelectedNode.IsExpanded Then
                        getNodeCanbo(treeCocau.SelectedNode, arr(1), arr(2), True)
                    End If
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
                idDVTT = arr(3)
                idPTT = arr(4)
                OverInfCB(idCanBo, arr(2))
                idQDLuong = ""
                idQDPhucap = ""
                idThemgio = ""
                If arr(2) = "TS" Then
                    If Not TS Then
                        TS = True
                        tpl.Add(TabControlLuong.TabPages(0))
                        TabControlLuong.TabPages.Remove(TabControlLuong.TabPages(0))
                        tpl.Add(TabControlLuong.TabPages(0))
                        TabControlLuong.TabPages.Remove(TabControlLuong.TabPages(0))
                        tpl.Add(TabControlLuong.TabPages(0))
                        TabControlLuong.TabPages.Remove(TabControlLuong.TabPages(0))
                        initTabTKThuHuong()
                    End If
                Else
                    If TS Then
                        TS = False
                        TabControlLuong.TabPages.Insert(0, tpl(0))
                        tpl.RemoveAt(0)
                        TabControlLuong.TabPages.Insert(1, tpl(0))
                        tpl.RemoveAt(0)
                        TabControlLuong.TabPages.Insert(2, tpl(0))
                        tpl.RemoveAt(0)
                    End If
                    Select Case TabControlLuong.SelectedTab.Name
                        Case "tabLuong"
                            initTabQDLuong()
                            'bindGridQDLuong(idCanBo)
                            'fillQDLuong(idQDLuong, idCanBo)
                            'gridQDLuong.Focus()
                        Case "tabPhuCap"
                            initTabQDPhuCap()
                            'bindGridQDPhucap(idCanBo)
                            'fillQDPhucap(idQDPhucap, idCanBo)
                            'gridQDPhuCap.Focus()
                        Case "tabThemGio"
                            initTabThemgio()
                            'bindGridThemgio(idCanBo)
                            'fillThemgio(idThemgio, idCanBo)
                            'gridThemgio.Focus()
                        Case Else
                            initTabTKThuHuong()
                    End Select
                End If

                If idCanBo <> "" Then
                    If checkRight_CreateRecord(idCanBo) Then
                        bntUpdate_L.Enabled = True
                        bntUpdate_PC.Enabled = True
                        bntUpdate_TG.Enabled = True
                        bntUpdate_TaiKhoan.Enabled = True
                        bntDelete_L.Enabled = True
                        bntDelete_PC.Enabled = True
                        bntDelete_TG.Enabled = True
                    Else
                        bntUpdate_L.Enabled = False
                        bntUpdate_PC.Enabled = False
                        bntUpdate_TG.Enabled = False
                        bntUpdate_TaiKhoan.Enabled = False
                        bntDelete_L.Enabled = False
                        bntDelete_PC.Enabled = False
                        bntDelete_TG.Enabled = False
                    End If
                End If
                
        End Select
    End Sub

    Private Sub TabControlLuong_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlLuong.SelectedIndexChanged
        callTab()
    End Sub

#Region "private function"

    Private Sub blankOverviewInfCB()
        idCanBo = ""
        idQDLuong = ""
        idQDPhucap = ""
        idThemgio = ""
        idDVTT = 0
        idPTT = 0
        txtCanbo.Text = ""
        txtMaCB.Text = ""
        txtNgaysinhCB.Text = ""
        txtGioiTinh.Text = ""
        txtCMT.Text = ""
        gridQDLuong.DataSource = Nothing
        gridQDPhuCap.DataSource = Nothing
        gridThemgio.DataSource = Nothing
        lkl_xemct.Enabled = False
    End Sub

    Private Sub callTab()
        Try
            Select Case TabControlLuong.SelectedTab.Name
                Case "tabLuong"
                    initTabQDLuong()
                Case "tabPhuCap"
                    initTabQDPhuCap()
                Case "tabThemGio"
                    initTabThemgio()
                Case Else
                    initTabTKThuHuong()
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub OverInfCB(ByVal vIdCanbo As String, ByVal vMaCb As String)
        Dim strSql As String = ""
        Dim dt As DataTable
        dbconn = New DBAccess
        If vMaCb = "TS" Then
            strSql = "SELECT Hoten, Gioitinh, Ngaysinh, CMT_So, MaCB FROM HSCB_TS WHERE id='" & vIdCanbo & "'"
        Else
            strSql = "SELECT Hoten, Gioitinh, Ngaysinh, CMT_So, MaCB FROM HS_Canbo WHERE idCanbo='" & vIdCanbo & "' and MaCB='" & vMaCb & "'"
        End If
        dt = dbconn.SelectDBRows(strSql)
        If dt.Rows.Count > 0 Then
            txtCanbo.Text = dt.Rows(0).Item("Hoten")
            txtMaCB.Text = dt.Rows(0).Item("MaCB")
            txtNgaysinhCB.Text = DateTimeUtil.getShortDate(dt.Rows(0).Item("Ngaysinh"))
            txtCMT.Text = dt.Rows(0).Item("CMT_So")
            If dt.Rows(0).Item("Gioitinh") Then
                txtGioiTinh.Text = "Nữ"
            Else
                txtGioiTinh.Text = "Nam"
            End If
        End If
        lkl_xemct.Enabled = True
    End Sub

    Private Sub lkl_xemct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
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

#End Region

#Region "tab QD Luong"
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
            If idQDLuong = "" Then
                If checkQuyetDinh("LUONG", idCanBo, txtSoQD_L.Text) Then
                    txtSoQD_L.Text = ""
                    txtSoQD_L.Focus()
                    strReturn = "Số quyết định lương của cán bộ đã tồn tại. Hãy nhập lại!"
                    Exit Try
                End If
            End If
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
        cboLoaiQDLuong.DataSource = listLoaiQDLuong()
        bntDelete_L.Enabled = False
        cbIsQD_NHCS_L.Checked = True
        txtSoQD_L.ReadOnly = False
        txtSoQD_L.Text = ""
        txtDVraQD_L.Text = getDonvi(IdDONVI)
        dpkNgayKy_L.Value = Date.Now
        txtNguoiKyQD_L.Text = ""
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
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
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
                idQDLuong = m_LuongCB.Add()
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

    Private Sub cboCVNguoiKyQD_L_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Private Sub dpkNgayHL_L_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
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
                    gridQDLuong_CellClick(sender, Nothing)
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
        txtNguoiKyQD_PC.Text = ""
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
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
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
            Dim db As DBAccess = New DBAccess
            Dim dtPC As DataTable
            Dim iPC As Integer = 0

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
                idQDPhucap = m_PhucapCB.Add()
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

    Private Sub txtNguoiKyQD_PC_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
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
            If checkQuyetDinh("PHUCAP", idCanBo, txtSoQD_PC.Text) And idQDPhucap = "" Then
                If MessageBox.Show("Số quyết định phụ cấp của cán bộ đã tồn tại." & vbCr & "Bạn muốn nhập tiếp thông tin cho số quyết định này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    If updateQDPhucap(idQDPhucap) Then
                        bindGridQDPhucap(idCanBo)
                        labAlert_P.Text = "Ghi dữ liệu thành công!"
                    Else
                        MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                End If
            Else
                If updateQDPhucap(idQDPhucap) Then
                    bindGridQDPhucap(idCanBo)
                    labAlert_P.Text = "Ghi dữ liệu thành công!"
                Else
                    MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End If
            End If
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

#Region "Tab Tài khoản thụ hưong"

    Private Sub initTabTKThuHuong()
        Try
            If idCanBo <> "" Then
                Dim strSql As String
                If TS Then
                    strSql = "SELECT NH_SHTK as NH_SoTK, NH_Ten_NH as NH_TenNH, MaSoThue, BHXH_SoSo, BHXH_NgayLam as BHXH_NgaySo, BHXH_NgayDong as BHXH_NgayBatDau, BHXH_NoiLam FROM HSCB_TS WHERE  Id ='" & idCanBo & "'"
                Else
                    strSql = "SELECT NH_SoTK, NH_TenNH, MaSoThue, BHXH_SoSo, BHXH_NgaySo, BHXH_NgayBatDau, BHXH_NoiLam FROM HS_CanBo WHERE  IdCanbo ='" & idCanBo & "'"
                End If

                Dim dt As DataTable
                Dim db As DBAccess = New DBAccess
                dt = db.SelectDBRows(strSql)
                txtSoTaiKhoan.Text = dt.Rows(0).Item("NH_SoTK")
                txtNganHang.Text = dt.Rows(0).Item("NH_TenNH")
                txtMaSoThue.Text = dt.Rows(0).Item("MaSoThue")

                txtSoSo.Text = dt.Rows(0).Item("BHXH_SoSo")
                If (dt.Rows(0).Item("BHXH_NgaySo") Is DBNull.Value) Then
                    dpkNgayLamSo.Checked = False
                Else
                    If dt.Rows(0).Item("BHXH_NgaySo") = DateTime.MinValue Then
                        dpkNgayLamSo.Checked = False
                    Else
                        dpkNgayLamSo.Checked = True
                        dpkNgayLamSo.Value = dt.Rows(0).Item("BHXH_NgaySo")
                    End If
                End If
                If (dt.Rows(0).Item("BHXH_NgayBatDau") Is DBNull.Value) Then
                    dpkNgayDong.Checked = False
                Else
                    If dt.Rows(0).Item("BHXH_NgayBatDau") = DateTime.MinValue Then
                        dpkNgayDong.Checked = False
                    Else
                        dpkNgayDong.Checked = True
                        dpkNgayDong.Value = dt.Rows(0).Item("BHXH_NgayBatDau")
                    End If
                End If
                txtNoiLamSo.Text = dt.Rows(0).Item("BHXH_NoiLam")
                txtSoTaiKhoan.Focus()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function updateTaiKhoan(ByVal vIdCanBo As String, ByVal vSoTaiKhoan As String, ByVal vNganHang As String, ByVal vMaSoThue As String, ByVal vBHXH_SoSo As String, ByVal vBHXH_NgaySo As Date, ByVal vBHXH_NgayBatDau As Date, ByVal vBHXH_NoiLam As String) As Boolean
        Try
            Dim strSql As String
            Dim db As DBAccess = New DBAccess
            If TS Then
                strSql = " UPDATE HSCB_TS  " & _
                         " SET NH_SHTK=N'" & vSoTaiKhoan & "', NH_Ten_NH=N'" & vNganHang & "', MaSoThue='" & vMaSoThue & "', " & _
                         "     BHXH_SoSo='" & vBHXH_SoSo & "', BHXH_NgayLam='" & vBHXH_NgaySo & "', BHXH_NgayDong='" & vBHXH_NgayBatDau & "', BHXH_NoiLam=N'" & vBHXH_NoiLam & "' " & _
                         " WHERE Id='" & vIdCanBo & "'"
            Else
                strSql = " UPDATE HS_CanBo  " & _
                         " SET NH_SoTK=N'" & vSoTaiKhoan & "', NH_TenNH=N'" & vNganHang & "', MaSoThue='" & vMaSoThue & "', " & _
                         "     BHXH_SoSo='" & vBHXH_SoSo & "', BHXH_NgaySo='" & vBHXH_NgaySo & "', BHXH_NgayBatDau='" & vBHXH_NgayBatDau & "', BHXH_NoiLam=N'" & vBHXH_NoiLam & "' " & _
                         " WHERE IdCanBo='" & vIdCanBo & "'"
            End If
            db.executeSQL(strSql)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

    Private Sub bntUpdate_TaiKhoan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUpdate_TaiKhoan.Click
        Try
            Dim lab_Err2 As String = ""
            If (Globals.Roles.IndexOf(";189;") < 0) Then
                MessageBox.Show("Bạn không có quyền sửa đổi thông tin tài khoản lương của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim NgayLamSo As Date
            Dim NgayDongBHXH As Date
            If dpkNgayLamSo.Checked Then
                NgayLamSo = DateTimeUtil.getDate(dpkNgayLamSo.Text)
            Else
                NgayLamSo = DateTime.MinValue
            End If
            If dpkNgayDong.Checked Then
                NgayDongBHXH = DateTimeUtil.getDate(dpkNgayDong.Text)
            Else
                NgayDongBHXH = DateTime.MinValue
            End If
            If updateTaiKhoan(idCanBo, txtSoTaiKhoan.Text.Trim, txtNganHang.Text.Trim, txtMaSoThue.Text.Trim, txtSoSo.Text.Trim, NgayLamSo, NgayDongBHXH, txtNoiLamSo.Text.Trim) Then
                MessageBox.Show("Ghi dữ liệu thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    Private Sub bntClose_TaiKhoan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClose_TaiKhoan.Click
        Close()
    End Sub

#Region "Tab Lương làm thêm giờ"

    Private Sub initTabThemgio()
        If Not initTG Then
            cboHinhThucTToan.DataSource = listDanhmuc(44)
            initTG = True
        End If
        bindGridThemgio(idCanBo)
        fillThemgio(idThemgio, idCanBo)
        gridThemgio.Focus()
    End Sub

    Private Sub bindGridThemgio(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridThemgio.AutoGenerateColumns = False
                Dim m_Themgio As Themgio = New Themgio
                gridThemgio.DataSource = m_Themgio.getAllByCanbo(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridThemgio.Rows.Count - 1
                    If Trim(gridThemgio.Rows(i).Cells(0).Value) = idThemgio Then
                        gridThemgio.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Lương làm thêm giờ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabThemgio() As String
        Dim strReturn As String = ""
        Try
            If Replace(txtG150.Text.Trim, "0", "") = "" And Replace(txtG200.Text.Trim, "0", "") = "" And Replace(txtG300.Text.Trim, "0", "") = "" And Replace(txtG_LamDemNgT.Text.Trim, "0", "") = "" And Replace(txtG_LamDemNgT.Text.Trim, "0", "") = "" And Replace(txtG_LamDemTBCN.Text.Trim, "0", "") = "" Then
                strReturn = "Chưa nhập số giờ làm thêm!"
                txtG150.Focus()
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankLuongThemgio()
        idThemgio = ""
        bntDelete_TG.Enabled = False
        txtG150.Text = ""
        txtG200.Text = ""
        txtG300.Text = ""
        txtG_LamDemNgT.Text = ""
        txtG_LamDemTBCN.Text = ""
        txtG_TToan_50.Text = ""
        txtG_TToan_100.Text = ""
        txtG_TToan_150.Text = ""
        txtG_TToan_200.Text = ""
        txtG_TToan_300.Text = ""
        txtG_NghiBu.Text = ""
        txtTrathem.Text = ""
        dtpkNam.Value = Now
        numThang.Value = Now.Month
        txtGhichu_TG.Text = ""
        labAlert_T.Text = ""
        dtpkNam.Focus()
    End Sub

    Private Sub fillThemgio(ByRef vIdThemgio As String, Optional ByVal vIdCanBo As String = "")
        Dim m_Themgio As Themgio = New Themgio
        labAlert_T.Text = ""
        bntDelete_TG.Enabled = True
        If vIdThemgio <> "" Then
            m_Themgio = m_Themgio.getRecord(vIdThemgio)
        Else
            m_Themgio = m_Themgio.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";198;") < 0) Then  ' Xoá
            bntDelete_TG.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDelete_TG.Enabled = True
                Else
                    bntDelete_TG.Enabled = False
                End If
            End If
        End If
        If m_Themgio.IdThemGio <> "" Then
            vIdThemgio = m_Themgio.IdThemGio
            Dim vDate As DateTime = New DateTime(CInt(m_Themgio.Nam.ToString()), 1, 1)
            dtpkNam.Value = vDate
            numThang.Value = m_Themgio.Thang
            txtG150.Text = m_Themgio.SoGio150.ToString
            txtG200.Text = m_Themgio.SoGio200.ToString
            txtG300.Text = m_Themgio.SoGio300.ToString
            txtG_LamDemNgT.Text = m_Themgio.SoGioLamDemNgayThuong.ToString
            txtG_LamDemTBCN.Text = m_Themgio.SoGioLamDemTBayCNhat.ToString
            cboHinhThucTToan.SelectedValue = m_Themgio.IDHT_ThanhToan
            txtG_TToan_50.Text = m_Themgio.TT_50.ToString
            txtG_TToan_100.Text = m_Themgio.TT_100.ToString
            txtG_TToan_150.Text = m_Themgio.TT_150.ToString
            txtG_TToan_200.Text = m_Themgio.TT_200.ToString
            txtG_TToan_300.Text = m_Themgio.TT_300.ToString
            txtG_NghiBu.Text = m_Themgio.TT_GioNghibu.ToString
            txtTrathem.Text = m_Themgio.TraThem.ToString
            txtGhichu_TG.Text = m_Themgio.GhiChu
        Else
            blankLuongThemgio()
        End If
    End Sub

    Private Function updateLuongThemgio(ByVal vIdThemgio As String) As Boolean
        Try
            Dim m_Themgio As Themgio = New Themgio
            Dim smartReader As SmartDataReader = New SmartDataReader
            If txtG150.Text.Trim = "" Then txtG150.Text = "0"
            If txtG200.Text.Trim = "" Then txtG200.Text = "0"
            If txtG300.Text.Trim = "" Then txtG300.Text = "0"
            If txtG_LamDemNgT.Text.Trim = "" Then txtG_LamDemNgT.Text = "0"
            If txtG_LamDemTBCN.Text.Trim = "" Then txtG_LamDemTBCN.Text = "0"
            If txtG_TToan_50.Text.Trim = "" Then txtG_TToan_50.Text = "0"
            If txtG_TToan_100.Text.Trim = "" Then txtG_TToan_100.Text = "0"
            If txtG_TToan_150.Text.Trim = "" Then txtG_TToan_150.Text = "0"
            If txtG_TToan_200.Text.Trim = "" Then txtG_TToan_200.Text = "0"
            If txtG_TToan_300.Text.Trim = "" Then txtG_TToan_300.Text = "0"
            If txtG_NghiBu.Text.Trim = "" Then txtG_NghiBu.Text = "0"
            If txtTrathem.Text.Trim = "" Then txtTrathem.Text = "0"
            m_Themgio.IdThemGio = vIdThemgio
            m_Themgio.IdDonVi_TG = idDVTT
            m_Themgio.IdPhong = idPTT
            m_Themgio.IdCanBo = idCanBo
            m_Themgio.Nam = dtpkNam.Value.Year
            m_Themgio.Thang = numThang.Value
            m_Themgio.SoGio150 = CDec(txtG150.Text.ToString)
            m_Themgio.SoGio200 = CDec(txtG200.Text.ToString)
            m_Themgio.SoGio300 = CDec(txtG300.Text.ToString)
            m_Themgio.SoGioLamDemNgayThuong = CDec(txtG_LamDemNgT.Text.ToString)
            m_Themgio.SoGioLamDemTBayCNhat = CDec(txtG_LamDemTBCN.Text.ToString)
            m_Themgio.IDHT_ThanhToan = CInt(cboHinhThucTToan.SelectedValue)
            m_Themgio.TT_50 = CDec(txtG_TToan_50.Text.ToString)
            m_Themgio.TT_100 = CDec(txtG_TToan_100.Text.ToString)
            m_Themgio.TT_150 = CDec(txtG_TToan_150.Text.ToString)
            m_Themgio.TT_200 = CDec(txtG_TToan_200.Text.ToString)
            m_Themgio.TT_300 = CDec(txtG_TToan_300.Text.ToString)
            m_Themgio.TT_GioNghibu = CDec(txtG_NghiBu.Text.ToString)
            m_Themgio.TraThem = N2Number(MoneyValue(txtTrathem.Text.ToString))
            m_Themgio.ThucLinh = 0
            m_Themgio.GhiChu = standardizeString(txtGhichu_TG.Text)
            If vIdThemgio <> "" Then
                m_Themgio.Update()
            Else
                idThemgio = m_Themgio.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub txtG150_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG150.TextChanged
        txtG150.Text = formatDouble(txtG150.Text.Trim)
    End Sub

    Private Sub txtG200_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG200.TextChanged
        txtG200.Text = formatDouble(txtG200.Text.Trim)
    End Sub

    Private Sub txtG300_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG300.TextChanged
        txtG300.Text = formatDouble(txtG300.Text.Trim)
    End Sub

    Private Sub txtG_LamDemNgT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_LamDemNgT.TextChanged
        txtG_LamDemNgT.Text = formatDouble(txtG_LamDemNgT.Text.Trim)
    End Sub

    Private Sub txtG_LamDemTBCN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_LamDemTBCN.TextChanged
        txtG_LamDemTBCN.Text = formatDouble(txtG_LamDemTBCN.Text.Trim)
    End Sub

    Private Sub txtG_TToan_50_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_TToan_50.TextChanged
        txtG_TToan_50.Text = formatDouble(txtG_TToan_50.Text.Trim)
    End Sub

    Private Sub txtG_TToan_100_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_TToan_100.TextChanged
        txtG_TToan_100.Text = formatDouble(txtG_TToan_100.Text.Trim)
    End Sub

    Private Sub txtG_TToan_150_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_TToan_150.TextChanged
        txtG_TToan_150.Text = formatDouble(txtG_TToan_150.Text.Trim)
    End Sub

    Private Sub txtG_TToan_200_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_TToan_200.TextChanged
        txtG_TToan_200.Text = formatDouble(txtG_TToan_200.Text.Trim)
    End Sub

    Private Sub txtG_TToan_300_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_TToan_300.TextChanged
        txtG_TToan_300.Text = formatDouble(txtG_TToan_300.Text.Trim)
    End Sub

    Private Sub txtG_NghiBu_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtG_NghiBu.TextChanged
        txtG_NghiBu.Text = formatDouble(txtG_NghiBu.Text.Trim)
    End Sub

    Private Sub cboHinhThucTToan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboHinhThucTToan.LostFocus
        If Not cboHinhThucTToan.SelectedValue Is Nothing Then
            Try
                If txtG150.Text.Trim = "" Then txtG150.Text = "0"
                If txtG200.Text.Trim = "" Then txtG200.Text = "0"
                If txtG300.Text.Trim = "" Then txtG300.Text = "0"
                If txtG_LamDemNgT.Text.Trim = "" Then txtG_LamDemNgT.Text = "0"
                If txtG_LamDemTBCN.Text.Trim = "" Then txtG_LamDemTBCN.Text = "0"
                Select Case getDanhmuc_MaSo(cboHinhThucTToan.SelectedValue)
                    Case "4401"  ' Thanh toan 100%
                        txtG_TToan_50.Text = "0"
                        txtG_TToan_100.Text = "0"
                        txtG_NghiBu.Text = "0"
                        txtG_TToan_150.Text = CDec(txtG150.Text.ToString) + CDec(txtG_LamDemNgT.Text.ToString)
                        txtG_TToan_200.Text = CDec(txtG200.Text.ToString) + CDec(txtG_LamDemTBCN.Text.ToString)
                        txtG_TToan_300.Text = CDec(txtG300.Text.ToString)
                        If CDec(txtG_LamDemNgT.Text.ToString) <= 0 And CDec(txtG_LamDemTBCN.Text.ToString) <= 0 Then
                            txtTrathem.Text = "0"
                        Else
                            Dim LuongGioCanBo As Long
                            LuongGioCanBo = getLuongGioCanBo(idCanBo, numThang.Value, dtpkNam.Value.Year)
                            txtTrathem.Text = LuongGioCanBo * 1.5 * CDec(txtG_LamDemNgT.Text.ToString) * 0.3 + LuongGioCanBo * 2 * CDec(txtG_LamDemTBCN.Text.ToString) * 0.3
                        End If
                    Case "4402"  ' nghi bu, thanh toan chenh lech
                        txtG_TToan_50.Text = CDec(txtG150.Text.ToString) + CDec(txtG_LamDemNgT.Text.ToString)
                        txtG_TToan_100.Text = CDec(txtG200.Text.ToString) + CDec(txtG_LamDemTBCN.Text.ToString)
                        txtG_NghiBu.Text = CDec(txtG150.Text.ToString) + CDec(txtG200.Text.ToString) + CDec(txtG_LamDemNgT.Text.ToString) + CDec(txtG_LamDemTBCN.Text.ToString)
                        txtG_TToan_150.Text = "0"
                        txtG_TToan_200.Text = "0"
                        txtG_TToan_300.Text = "0"
                        If CDec(txtG_LamDemNgT.Text.ToString) <= 0 And CDec(txtG_LamDemTBCN.Text.ToString) <= 0 Then
                            txtTrathem.Text = "0"
                        Else
                            Dim LuongGioCanBo As Long
                            LuongGioCanBo = getLuongGioCanBo(idCanBo, numThang.Value, dtpkNam.Value.Year)
                            txtTrathem.Text = LuongGioCanBo * 1.5 * CDec(txtG_LamDemNgT.Text.ToString) * 0.3 + LuongGioCanBo * 2 * CDec(txtG_LamDemTBCN.Text.ToString) * 0.3
                        End If
                    Case Else  ' Do nguoi dung tu nhap gia tri
                        txtG_TToan_50.Text = "0"
                        txtG_TToan_100.Text = "0"
                        txtG_TToan_150.Text = "0"
                        txtG_TToan_200.Text = "0"
                        txtG_TToan_300.Text = "0"
                        txtG_NghiBu.Text = "0"
                End Select
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub cboHinhThucTToan_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboHinhThucTToan.SelectedValueChanged
        If Not cboHinhThucTToan.SelectedValue Is Nothing Then
            Try
                If txtG150.Text.Trim = "" Then txtG150.Text = "0"
                If txtG200.Text.Trim = "" Then txtG200.Text = "0"
                If txtG300.Text.Trim = "" Then txtG300.Text = "0"
                If txtG_LamDemNgT.Text.Trim = "" Then txtG_LamDemNgT.Text = "0"
                If txtG_LamDemTBCN.Text.Trim = "" Then txtG_LamDemTBCN.Text = "0"
                Select Case getDanhmuc_MaSo(cboHinhThucTToan.SelectedValue)
                    Case "4401"  ' Thanh toan 100%
                        txtG_TToan_50.Text = "0"
                        txtG_TToan_100.Text = "0"
                        txtG_NghiBu.Text = "0"
                        txtG_TToan_150.Text = CDec(txtG150.Text.ToString) + CDec(txtG_LamDemNgT.Text.ToString)
                        txtG_TToan_200.Text = CDec(txtG200.Text.ToString) + CDec(txtG_LamDemTBCN.Text.ToString)
                        txtG_TToan_300.Text = CDec(txtG300.Text.ToString)
                        If CDec(txtG_LamDemNgT.Text.ToString) <= 0 And CDec(txtG_LamDemTBCN.Text.ToString) <= 0 Then
                            txtTrathem.Text = "0"
                        Else
                            Dim LuongGioCanBo As Long
                            LuongGioCanBo = getLuongGioCanBo(idCanBo, numThang.Value, dtpkNam.Value.Year)
                            txtTrathem.Text = LuongGioCanBo * 1.5 * CDec(txtG_LamDemNgT.Text.ToString) * 0.3 + LuongGioCanBo * 2 * CDec(txtG_LamDemTBCN.Text.ToString) * 0.3
                        End If
                    Case "4402"  ' nghi bu, thanh toan chenh lech
                        txtG_TToan_50.Text = CDec(txtG150.Text.ToString) + CDec(txtG_LamDemNgT.Text.ToString)
                        txtG_TToan_100.Text = CDec(txtG200.Text.ToString) + CDec(txtG_LamDemTBCN.Text.ToString)
                        txtG_NghiBu.Text = CDec(txtG150.Text.ToString) + CDec(txtG200.Text.ToString) + CDec(txtG_LamDemNgT.Text.ToString) + CDec(txtG_LamDemTBCN.Text.ToString)
                        txtG_TToan_150.Text = "0"
                        txtG_TToan_200.Text = "0"
                        txtG_TToan_300.Text = "0"
                        If CDec(txtG_LamDemNgT.Text.ToString) <= 0 And CDec(txtG_LamDemTBCN.Text.ToString) <= 0 Then
                            txtTrathem.Text = "0"
                        Else
                            Dim LuongGioCanBo As Long
                            LuongGioCanBo = getLuongGioCanBo(idCanBo, numThang.Value, dtpkNam.Value.Year)
                            txtTrathem.Text = LuongGioCanBo * 1.5 * CDec(txtG_LamDemNgT.Text.ToString) * 0.3 + LuongGioCanBo * 2 * CDec(txtG_LamDemTBCN.Text.ToString) * 0.3
                        End If
                    Case Else  ' Do nguoi dung tu nhap gia tri
                        txtG_TToan_50.Text = "0"
                        txtG_TToan_100.Text = "0"
                        txtG_TToan_150.Text = "0"
                        txtG_TToan_200.Text = "0"
                        txtG_TToan_300.Text = "0"
                        txtG_NghiBu.Text = "0"
                End Select
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub gridThemgio_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridThemgio.CellClick
        Try
            gridThemgio.Rows(e.RowIndex).Selected = True
            idThemgio = gridThemgio.CurrentRow.Cells("IdLuongThemGio").Value.ToString
            fillThemgio(idThemgio)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridThemgio_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridThemgio.KeyUp
        gridThemgio_CellClick(sender, Nothing)
    End Sub

    Private Sub bntNew_TG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntNew_TG.Click
        blankLuongThemgio()
    End Sub

    Private Sub bntUpdate_TG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUpdate_TG.Click
        Try
            Dim lab_Err3 As String = ""
            If (idThemgio <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";197;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu về thêm giờ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridThemgio_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err3 = checkTabThemgio()
            If lab_Err3 <> "" Then
                MessageBox.Show(lab_Err3, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateLuongThemgio(idThemgio) Then
                bindGridThemgio(idCanBo)
                labAlert_T.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_TG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntCancel_TG.Click
        If idThemgio = "" Then
            blankLuongThemgio()
        Else
            fillThemgio(idThemgio)
        End If
    End Sub

    Private Sub bntDelete_TG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntDelete_TG.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridThemgio.Rows.Count > 0) Then
                For i As Int32 = 0 To gridThemgio.Rows.Count - 1
                    If (CType(gridThemgio.Rows(i).Cells("cln_cbTG").Value, Boolean) = True) Then
                        arrDel.Add(gridThemgio.Rows(i).Cells("IdLuongThemGio").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idThemgio <> "") Then arrDel.Add(idThemgio.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_Themgio As Themgio = New Themgio
                            m_Themgio.IdThemGio = arrDel(i).ToString()
                            m_Themgio.Delete()
                        Next
                        blankLuongThemgio()
                        bindGridThemgio(idCanBo)
                        labAlert_T.Text = "Xoá dữ liệu thành công!"
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

    Private Sub bntClose_TG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClose_TG.Click
        Close()
    End Sub

    Private Sub txtTrathem_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTrathem.TextChanged
        Try
            txtTrathem = formatMoneyinTextbox(txtTrathem)
        Catch ex As Exception
        End Try
    End Sub

#End Region
  
   
End Class