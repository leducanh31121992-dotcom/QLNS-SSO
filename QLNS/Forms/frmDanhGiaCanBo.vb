Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmDanhGiaCanBo
    Inherits System.Windows.Forms.Form
    Public IDCB_HS_Canbo As String = ""   ' Id Cán bộ truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDDV_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDPB_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Private dbconn As DBAccess
    Private idCanBo As String = ""
    Private idDGCB As String = ""
    Private idTNLD As String = ""
    Private idNXLD As String = ""
    Private init_D As Boolean
    Private init_T As Boolean
    Private init_N As Boolean
    Private initQH As Boolean = False
    Private idQHCB As String = ""

    Private Sub frmDanhGiaCanBo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        lkl_xemct.Enabled = False
        'Check quyền thành viên
        Dim roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";151;152;153;154;", roles)) Then
            If tabControlDanhGia.TabPages.Contains(tabDGCB) Then tabControlDanhGia.TabPages.Remove(tabDGCB)
        End If
        If Not (Globals.IsIntersect(";155;156;157;158;", roles)) Then
            If tabControlDanhGia.TabPages.Contains(tabTNLD) Then tabControlDanhGia.TabPages.Remove(tabTNLD)
        End If
        If Not (Globals.IsIntersect(";103;104;105;106;", roles)) Then
            If tabControlDanhGia.TabPages.Contains(tabQHCB) Then tabControlDanhGia.TabPages.Remove(tabQHCB)
        End If
        If (Globals.Roles.IndexOf(";151;") < 0) Then
            gridDGCB.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";152;") < 0) Then
            bntNewDG.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";154;") < 0) Then
            bntDeleteDG.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";155;") < 0) Then
            gridTNLD.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";156;") < 0) Then
            bntNewTN.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";158;") < 0) Then
            bntDeleteTN.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";103;") < 0) Then
            gridQHCB.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";104;") < 0) Then
            bntNewQH.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";106;") < 0) Then
            bntDeleteQH.Enabled = False
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
                initTabDGCB()
            End If

            If checkRight_CreateRecord(IDCB_HS_Canbo) Then
                bntUpdateDG.Enabled = True
                bntUpdateTN.Enabled = True
                bntAddQH.Enabled = True
                bntDeleteDG.Enabled = True
                bntDeleteTN.Enabled = True
                bntDeleteQH.Enabled = True
            Else
                bntUpdateDG.Enabled = False
                bntUpdateTN.Enabled = False
                bntAddQH.Enabled = False
                bntDeleteDG.Enabled = False
                bntDeleteTN.Enabled = False
                bntDeleteQH.Enabled = False
            End If
        End If

    End Sub

    Private Sub frmDanhGiaCanBo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
        Dim arr() As String
        arr = treeCocau.SelectedNode.Tag.ToString.Split("_")
        labAlert_DG.Text = ""
        labAlert_TN.Text = ""
        Select Case arr(0)
            Case "DV"
                blankOverInfCB()
                blankDGCB()
                blankTNLD()
                blankQHCB()
                If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not treeCocau.SelectedNode.IsExpanded Then
                        getNode(treeCocau.SelectedNode, arr(1), arr(2))
                    End If
                End If
                treeCocau.SelectedNode.Expand()
            Case "PB"
                blankOverInfCB()
                blankDGCB()
                blankTNLD()
                blankQHCB()
                If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not treeCocau.SelectedNode.IsExpanded Then
                        getNodeCanbo(treeCocau.SelectedNode, arr(1), arr(2))
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
                OverInfCB(idCanBo, arr(2))
                If tabControlDanhGia.SelectedTab.Name = "tabDGCB" Then
                    initTabDGCB()
                ElseIf tabControlDanhGia.SelectedTab.Name = "tabTNLD" Then
                    initTabTNLD()
                Else
                    initTabQHCB()
                End If

                If checkRight_CreateRecord(idCanBo) Then
                    bntUpdateDG.Enabled = True
                    bntUpdateTN.Enabled = True
                    bntAddQH.Enabled = True
                    bntDeleteDG.Enabled = True
                    bntDeleteTN.Enabled = True
                    bntDeleteQH.Enabled = True
                Else
                    bntUpdateDG.Enabled = False
                    bntUpdateTN.Enabled = False
                    bntAddQH.Enabled = False
                    bntDeleteDG.Enabled = False
                    bntDeleteTN.Enabled = False
                    bntDeleteQH.Enabled = False
                End If
        End Select
    End Sub

    Private Sub tabControlDanhGia_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabControlDanhGia.SelectedIndexChanged
        callTab()
    End Sub

    Private Sub lkl_xemct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkl_xemct.LinkClicked
        If (idCanBo <> "") Then
            Dim obj_detail As frmHS_ChiTiet = New frmHS_ChiTiet()
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
    Private Sub callTab()
        Try
            If tabControlDanhGia.SelectedTab.Name = "tabDGCB" Then
                initTabDGCB()
            ElseIf tabControlDanhGia.SelectedTab.Name = "tabTNLD" Then
                initTabTNLD()
            Else
                initTabQHCB()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub OverInfCB(ByVal vIdCanbo As String, ByVal vMaCb As String)
        Dim dt As DataTable
        dbconn = New DBAccess
        dt = dbconn.SelectDBRows("SELECT Hoten, Gioitinh, Ngaysinh, CMT_So FROM HS_Canbo WHERE idCanbo='" & vIdCanbo & "' and MaCB='" & vMaCb & "'")
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
        End If
        lkl_xemct.Enabled = True
    End Sub

    Private Sub blankOverInfCB()
        idCanBo = ""
        idDGCB = ""
        idTNLD = ""
        idNXLD = ""
        txtCanbo.Text = ""
        txtMaCB.Text = ""
        txtNgaysinhCB.Text = ""
        txtGioiTinh.Text = ""
        txtCMT.Text = ""
        gridDGCB.DataSource = Nothing
        gridTNLD.DataSource = Nothing
        lkl_xemct.Enabled = False
    End Sub
#End Region

#Region "tab Danh gia can bo"

    Private Sub initTabDGCB()
        If Not init_D Then
            bindCboDanhGiaCV()
            init_D = True
        End If
        bindGridDGCB(idCanBo)
        txtNamDG.Text = (Now.Year - 1).ToString
        fillDGCB(idDGCB, idCanBo)
        gridDGCB.Focus()
    End Sub

    Private Sub bindCboDanhGiaCV()
        cboDanhgiaCV.DataSource = listDanhmuc(1)
    End Sub

    Private Sub bindGridDGCB(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridDGCB.AutoGenerateColumns = False
                Dim m_DanhGiaCB As DanhGiaCB = New DanhGiaCB
                gridDGCB.DataSource = m_DanhGiaCB.getAllByCanbo(vIdCanbo)
                'fillDGCB(0, vIdCanbo)
                Dim i As Integer = 0
                While i <= gridDGCB.Rows.Count - 1
                    If Trim(gridDGCB.Rows(i).Cells(0).Value) = idDGCB Then
                        gridDGCB.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Đánh giá cán bộ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabDGCB() As String
        Dim strReturn As String = ""
        Try
            If txtNamDG.Text = "" Or txtNamDG.Text = "0" Then
                txtNamDG.Focus()
                strReturn = "Chưa nhập năm đánh giá!"
                Exit Try
            End If
            If txtDotDG.Text = "" Or txtDotDG.Text = "0" Then
                txtDotDG.Focus()
                strReturn = "Chưa nhập đợt đánh giá!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankDGCB()
        idDGCB = ""
        txtNamDG.Text = ""
        txtDotDG.Text = ""
        txtCongviecDN.Text = ""
        txtKetquaTH.Text = ""
        txtPhamchatCT.Text = ""
        txtDaoducLS.Text = ""
        txtNhanxetTT.Text = ""
        txtNhanxetBT.Text = ""
        txtGhichuDG.Text = ""
        txtNamDG.Text = Now.Year.ToString
        bntDeleteDG.Enabled = False
        labAlert_DG.Text = ""
        txtNamDG.Focus()
    End Sub

    Private Sub fillDGCB(ByRef vIdDGCB As String, Optional ByVal vIdCanBo As String = "")
        Dim m_DanhGiaCB As DanhGiaCB = New DanhGiaCB
        labAlert_DG.Text = ""
        If vIdCanBo = "" Then
            m_DanhGiaCB = m_DanhGiaCB.getRecord(vIdDGCB)
        Else
            m_DanhGiaCB = m_DanhGiaCB.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";154;") < 0) Then  ' Xoá
            bntDeleteDG.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteDG.Enabled = True
                Else
                    bntDeleteDG.Enabled = False
                End If
            End If
        End If
        If m_DanhGiaCB.IdDGCB <> "" Then
            vIdDGCB = m_DanhGiaCB.IdDGCB
            txtNamDG.Text = m_DanhGiaCB.Nam_DG
            txtDotDG.Text = m_DanhGiaCB.Dot_DG
            txtCongviecDN.Text = m_DanhGiaCB.CongViec_DN
            txtKetquaTH.Text = m_DanhGiaCB.KetQua_TH
            txtPhamchatCT.Text = m_DanhGiaCB.PhamChat_CT
            txtDaoducLS.Text = m_DanhGiaCB.DaoDuc_LS
            txtNhanxetTT.Text = m_DanhGiaCB.NhanXet_TT
            txtNhanxetBT.Text = m_DanhGiaCB.NhanXet_BT
            cboDanhgiaCV.SelectedValue = m_DanhGiaCB.IdDanhGiaCV
            txtGhichuDG.Text = m_DanhGiaCB.GhiChu
        Else
            blankDGCB()
        End If
    End Sub

    Private Function updateDGCB(ByVal vIdDGCB As String) As Boolean
        Try
            Dim m_DanhGiaCB As DanhGiaCB = New DanhGiaCB
            m_DanhGiaCB.IdDGCB = vIdDGCB
            m_DanhGiaCB.Nam_DG = txtNamDG.Text
            m_DanhGiaCB.Dot_DG = txtDotDG.Text
            m_DanhGiaCB.IdCanBo = idCanBo
            m_DanhGiaCB.CongViec_DN = txtCongviecDN.Text
            m_DanhGiaCB.KetQua_TH = txtKetquaTH.Text
            m_DanhGiaCB.PhamChat_CT = txtPhamchatCT.Text
            m_DanhGiaCB.DaoDuc_LS = txtDaoducLS.Text
            m_DanhGiaCB.NhanXet_TT = txtNhanxetTT.Text
            m_DanhGiaCB.NhanXet_BT = txtNhanxetBT.Text
            m_DanhGiaCB.IdDanhGiaCV = CInt(cboDanhgiaCV.SelectedValue)
            m_DanhGiaCB.GhiChu = txtGhichuDG.Text
            If idDGCB <> "" Then
                m_DanhGiaCB.Update()
            Else
                idDGCB = m_DanhGiaCB.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub gridDGCB_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridDGCB.CellClick
        Try
            'gridDGCB.Rows(e.RowIndex).Selected = True
            idDGCB = gridDGCB.CurrentRow.Cells("Id_DGCB").Value.ToString
            fillDGCB(idDGCB)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridDGCB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridDGCB.KeyUp
        gridDGCB_CellClick(sender, Nothing)
    End Sub

    Private Sub gridDGCB_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridDGCB.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridDGCB.Columns("IdDanhGiaCV").Index Then
                e.FormattingApplied = True
                row = gridDGCB.Rows(e.RowIndex)
                e.Value = getDanhmuc_Name(1, CInt(row.Cells("IdDanhGiaCV").Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtNamDG_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNamDG.TextChanged
        txtNamDG.Text = Val(txtNamDG.Text.Trim)
    End Sub

    Private Sub txtDotDG_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDotDG.TextChanged
        txtDotDG.Text = Val(txtDotDG.Text.Trim)
    End Sub

    Private Sub bntUpdateDG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateDG.Click
        Dim lab_ErrDGCB As String = ""
        If (idDGCB <> "") Then   ' Sửa
            If (Globals.Roles.IndexOf(";153;") < 0) Then
                MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu đánh giá cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                gridDGCB_CellClick(sender, Nothing)
                Return
            End If
        End If
        If idCanBo = "" Then
            MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        lab_ErrDGCB = checkTabDGCB()
        If lab_ErrDGCB <> "" Then
            MessageBox.Show(lab_ErrDGCB, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If updateDGCB(idDGCB) Then
            bindGridDGCB(idCanBo)
            labAlert_DG.Text = "Ghi dữ liệu thành công!"
        Else
            MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub bntNewDG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewDG.Click
        blankDGCB()
    End Sub

    Private Sub bntDeleteDG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteDG.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridDGCB.Rows.Count > 0) Then
                For i As Int32 = 0 To gridDGCB.Rows.Count - 1
                    If (CType(gridDGCB.Rows(i).Cells("cln_cbDGCB").Value, Boolean) = True) Then
                        arrDel.Add(gridDGCB.Rows(i).Cells("Id_DGCB").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idDGCB <> "") Then arrDel.Add(idDGCB.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_DanhGiaCB As DanhGiaCB = New DanhGiaCB
                            m_DanhGiaCB.IdDGCB = arrDel(i).ToString()
                            m_DanhGiaCB.Delete()
                        Next
                        blankDGCB()
                        bindGridDGCB(idCanBo)
                        labAlert_DG.Text = "Xoá dữ liệu thành công!"
                    Catch ex As Exception
                        MessageBox.Show("Xoá dữ liệu không thành công:" & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Else
                MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công:" & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancelDG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelDG.Click
        If idDGCB = "" Then
            blankDGCB()
        Else
            fillDGCB(idDGCB)
        End If
    End Sub

    Private Sub bntCloseDG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseDG.Click
        Close()
    End Sub

#End Region

#Region "tab Tin nhiem lanh dao"

    Private Sub initTabTNLD()
        If Not init_T Then
            bindCboChucDanhTN()
            init_T = True
        End If
        bindGridTNLD(idCanBo)
        dpkNgayTN.Value = Date.Now
        fillTNLD(idTNLD, idCanBo)
        gridTNLD.Focus()
    End Sub

    Private Sub bindCboChucDanhTN()
        cboChucdanhTN.DataSource = listDanhmuc(14)
    End Sub

    Private Sub bindGridTNLD(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridTNLD.AutoGenerateColumns = False
                Dim m_TinNhiemLD As TinNhiemLD = New TinNhiemLD
                gridTNLD.DataSource = m_TinNhiemLD.getAllByCanbo(vIdCanbo)
                'fillTNLD(0, vIdCanbo)
                Dim i As Integer = 0
                While i <= gridTNLD.Rows.Count - 1
                    If Trim(gridTNLD.Rows(i).Cells(0).Value) = idTNLD Then
                        gridTNLD.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách bỏ phiếu tín nhiệm lãnh đạo !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabTNLD() As String
        Dim strReturn As String = ""
        Try
            If txtTongPhieuTN.Text = "" Or txtTongPhieuTN.Text = "0" Then
                txtTongPhieuTN.Focus()
                strReturn = "Chưa nhập Tổng số phiếu!"
                Exit Try
            End If
            If (txtPhieuTN.Text.Trim <> "" Or txtPhieuTN.Text.Trim <> "0") And CInt(txtPhieuTN.Text.ToString) > CInt(txtTongPhieuTN.Text.ToString) Then
                txtPhieuTN.Focus()
                strReturn = "Số phiếu tín nhiệm phải nhỏ hơn tổng số phiếu!"
                Exit Try
            End If
            If (txtPhieuKTN_NL.Text.Trim <> "" Or txtPhieuKTN_NL.Text.Trim <> "0") And CInt(txtPhieuKTN_NL.Text.ToString) > CInt(txtTongPhieuTN.Text.ToString) Then
                txtPhieuKTN_NL.Focus()
                strReturn = "Số phiếu không tín nhiệm năng lực phải nhỏ hơn tổng số phiếu!"
                Exit Try
            End If
            If (txtPhieuKTN_DD.Text.Trim <> "" Or txtPhieuKTN_DD.Text.Trim <> "0") And CInt(txtPhieuKTN_DD.Text.ToString) > CInt(txtTongPhieuTN.Text) Then
                txtPhieuKTN_DD.Focus()
                strReturn = "Số phiếu không tín nhiệm đạo đức phải nhỏ hơn tổng số phiếu!"
                Exit Try
            End If
            If (txtPhieuKTN_K.Text.Trim <> "" Or txtPhieuKTN_K.Text.Trim <> "0") And CInt(txtPhieuKTN_K.Text) > CInt(txtTongPhieuTN.Text) Then
                txtPhieuKTN_K.Focus()
                strReturn = "Số phiếu không tín nhiệm với các lý do khác phải nhỏ hơn tổng số phiếu!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankTNLD()
        idTNLD = ""
        dpkNgayTN.Value = Date.Now
        txtTongPhieuTN.Text = ""
        txtPhieuTN.Text = "0"
        txtTyleTN.Text = "0"
        txtPhieuKTN_NL.Text = "0"
        txtTyleKTN_NL.Text = "0"
        txtPhieuKTN_DD.Text = "0"
        txtTyleKTN_DD.Text = "0"
        txtPhieuKTN_K.Text = "0"
        txtTyleKTN_K.Text = "0"
        txtGhichuTN.Text = ""
        bntDeleteTN.Enabled = False
        labAlert_TN.Text = ""
        dpkNgayTN.Focus()
    End Sub

    Private Sub fillTNLD(ByRef vIdTNLD As String, Optional ByVal vIdCanBo As String = "")
        Dim m_TinNhiemLD As TinNhiemLD = New TinNhiemLD
        labAlert_TN.Text = ""
        If vIdCanBo = "" Then
            m_TinNhiemLD = m_TinNhiemLD.getRecord(vIdTNLD)
        Else
            m_TinNhiemLD = m_TinNhiemLD.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";158;") < 0) Then  ' Xoá
            bntDeleteTN.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteTN.Enabled = True
                Else
                    bntDeleteTN.Enabled = True
                End If
            End If
        End If
        If m_TinNhiemLD.IdTNhLD <> "" Then
            vIdTNLD = m_TinNhiemLD.IdTNhLD
            dpkNgayTN.Value = m_TinNhiemLD.Ngay
            txtTongPhieuTN.Text = m_TinNhiemLD.TongPhieu
            txtPhieuTN.Text = m_TinNhiemLD.SoPhieu_TN
            txtTyleTN.Text = m_TinNhiemLD.TyLe_TN
            txtPhieuKTN_NL.Text = m_TinNhiemLD.Phieu_KTN_NL
            txtTyleKTN_NL.Text = m_TinNhiemLD.TyLe_KTN_NL
            txtPhieuKTN_DD.Text = m_TinNhiemLD.TyLe_KTN_DD
            txtTyleKTN_DD.Text = m_TinNhiemLD.TyLe_KTN_DD
            txtPhieuKTN_K.Text = m_TinNhiemLD.Phieu_KTN_khac
            txtTyleKTN_K.Text = m_TinNhiemLD.TyLe_KTN_khac
            cboChucdanhTN.SelectedValue = m_TinNhiemLD.IdChucDanh
            txtGhichuTN.Text = m_TinNhiemLD.GhiChu
        Else
            blankTNLD()
        End If
    End Sub

    Private Function updateTNLD(ByVal vIdTNLD As String) As Boolean
        Try
            Dim m_TinNhiemLD As TinNhiemLD = New TinNhiemLD
            m_TinNhiemLD.IdTNhLD = vIdTNLD
            m_TinNhiemLD.Ngay = DateTimeUtil.getDate(dpkNgayTN.Text)
            m_TinNhiemLD.TongPhieu = txtTongPhieuTN.Text
            m_TinNhiemLD.IdCanBo = idCanBo
            m_TinNhiemLD.SoPhieu_TN = txtPhieuTN.Text
            m_TinNhiemLD.TyLe_TN = txtTyleTN.Text
            m_TinNhiemLD.Phieu_KTN_NL = txtPhieuKTN_NL.Text
            m_TinNhiemLD.TyLe_KTN_NL = txtTyleKTN_NL.Text
            m_TinNhiemLD.TyLe_KTN_DD = txtPhieuKTN_DD.Text
            m_TinNhiemLD.TyLe_KTN_DD = txtTyleKTN_DD.Text
            m_TinNhiemLD.Phieu_KTN_khac = txtPhieuKTN_K.Text
            m_TinNhiemLD.TyLe_KTN_khac = txtTyleKTN_K.Text
            m_TinNhiemLD.IdChucDanh = cboChucdanhTN.SelectedValue
            m_TinNhiemLD.GhiChu = txtGhichuTN.Text
            If idTNLD <> "" Then
                m_TinNhiemLD.Update()
            Else
                idTNLD = m_TinNhiemLD.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub txtTongPhieuTN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTongPhieuTN.TextChanged
        txtTongPhieuTN.Text = Val(txtTongPhieuTN.Text.Trim)
    End Sub

    Private Sub txtPhieuTN_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuTN.LostFocus
        Try
            If (txtPhieuTN.Text.Trim <> "" Or txtPhieuTN.Text.Trim <> "0") And (txtTongPhieuTN.Text.Trim <> "" Or txtTongPhieuTN.Text.Trim <> "0") Then
                txtTyleTN.Text = Format(CInt(txtPhieuTN.Text.ToString) * 100 / CInt(txtTongPhieuTN.Text.ToString), "#.#0; (#.#0)")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtPhieuTN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuTN.TextChanged
        txtPhieuTN.Text = Val(txtPhieuTN.Text.Trim)
    End Sub

    Private Sub txtPhieuKTN_NL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuKTN_NL.LostFocus
        Try
            If (txtPhieuKTN_NL.Text.Trim <> "" Or txtPhieuKTN_NL.Text.Trim <> "0") And (txtTongPhieuTN.Text.Trim <> "" Or txtTongPhieuTN.Text.Trim <> "0") Then
                txtTyleKTN_NL.Text = Format(CInt(txtPhieuKTN_NL.Text.ToString) * 100 / CInt(txtTongPhieuTN.Text.ToString), "#.#0; (#.#0)")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtPhieuKTN_NL_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuKTN_NL.TextChanged
        txtPhieuKTN_NL.Text = Val(txtPhieuKTN_NL.Text.Trim)
    End Sub

    Private Sub txtPhieuKTN_DD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuKTN_DD.LostFocus
        Try
            If (txtPhieuKTN_DD.Text.Trim <> "" Or txtPhieuKTN_DD.Text.Trim <> "0") And (txtTongPhieuTN.Text.Trim <> "" Or txtTongPhieuTN.Text.Trim <> "0") Then
                txtTyleKTN_DD.Text = Format(CInt(txtPhieuKTN_DD.Text.ToString) * 100 / CInt(txtTongPhieuTN.Text.ToString), "#.#0; (#.#0)")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtPhieuKTN_DD_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuKTN_DD.TextChanged
        txtPhieuKTN_DD.Text = Val(txtPhieuKTN_DD.Text.Trim)
    End Sub

    Private Sub txtPhieuKTN_K_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuKTN_K.LostFocus
        Try
            If (txtPhieuKTN_K.Text.Trim <> "" Or txtPhieuKTN_K.Text.Trim <> "0") And (txtTongPhieuTN.Text.Trim <> "" Or txtTongPhieuTN.Text.Trim <> "0") Then
                txtTyleKTN_K.Text = Format(CInt(txtPhieuKTN_K.Text.ToString) * 100 / CInt(txtTongPhieuTN.Text.ToString), "#.#0; (#.#0)")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtPhieuKTN_K_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhieuKTN_K.TextChanged
        txtPhieuKTN_K.Text = Val(txtPhieuKTN_K.Text.Trim)
    End Sub

    Private Sub gridTNLD_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridTNLD.CellClick
        Try
            'gridTNLD.Rows(e.RowIndex).Selected = True
            idTNLD = gridTNLD.CurrentRow.Cells("IdTNhLD").Value.ToString
            fillTNLD(idTNLD)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridTNLD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridTNLD.KeyUp
        gridTNLD_CellClick(sender, Nothing)
    End Sub

    Private Sub gridTNLD_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridTNLD.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridTNLD.Columns("IdChucDanh").Index Then
                e.FormattingApplied = True
                row = gridTNLD.Rows(e.RowIndex)
                e.Value = getDanhmuc_Name(14, CInt(row.Cells("IdChucDanh").Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdateTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateTN.Click
        Try
            Dim lab_ErrPTN As String
            If (idTNLD <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";157;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu tín nhiệm lãnh đạo!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridTNLD_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_ErrPTN = checkTabTNLD()
            If lab_ErrPTN <> "" Then
                MessageBox.Show(lab_ErrPTN, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateTNLD(idTNLD) Then
                bindGridTNLD(idCanBo)
                labAlert_TN.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Ghi dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntNewTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewTN.Click
        blankTNLD()
    End Sub

    Private Sub bntDeleteTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteTN.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridTNLD.Rows.Count > 0) Then
                For i As Int32 = 0 To gridTNLD.Rows.Count - 1
                    If (CType(gridTNLD.Rows(i).Cells("cln_cbTNLD").Value, Boolean) = True) Then
                        arrDel.Add(gridTNLD.Rows(i).Cells("IdTNhLD").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idTNLD <> "") Then arrDel.Add(idTNLD.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        For i As Int32 = 0 To arrDel.Count - 1
                            Dim m_TinNhiemLD As TinNhiemLD = New TinNhiemLD
                            m_TinNhiemLD.IdTNhLD = arrDel(i).ToString()
                            m_TinNhiemLD.Delete()
                        Next
                        blankTNLD()
                        bindGridTNLD(idCanBo)
                        labAlert_TN.Text = "Xoá dữ liệu thành công!"
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

    Private Sub bntCancelTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelTN.Click
        If idTNLD = "" Then
            blankTNLD()
        Else
            fillTNLD(idTNLD)
        End If
    End Sub

    Private Sub bntCloseTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseTN.Click
        Close()
    End Sub

#End Region

#Region "tab Quy hoạch cán bộ"

    Private Sub initTabQHCB()
        Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
        m_QDNhanSu = m_QDNhanSu.getFinalRecord(idCanBo)
        If Not initQH Then
            initQH = True
            bindCboChucVuQH(m_QDNhanSu.IdChucvu_Moi)
            cboChuyenmonHT.DataSource = listDanhmuc(12, False, True)
            cboChuyenmonHT.SelectedValue = m_QDNhanSu.IdChuyenMon_Moi
            bindCboChuyenMonQH()
            bindCboChinhTriQH()
            bindCboNN_TH()
        End If
        bindGridQHCB(idCanBo)
        fillQHCB(idQHCB, idCanBo)
        gridQHCB.Focus()
    End Sub

    Private Sub bindCboChucVuQH(ByVal IdChucVuHT As Integer)
        'Dim dt As DataTable = New DataTable
        'Dim dt1 As DataTable = New DataTable("tb1")
        'Dim dt2 As DataTable = New DataTable("tb2")
        'Dim dt3 As DataTable = New DataTable("tb3")
        'Dim dt4 As DataTable = New DataTable("tb4")
        'Dim dt1, dt2, dt3, dt4 As DataTable
        'dt = listDanhmuc(14)
        'dt1 = dt
        'dt2 = dt
        'dt3 = dt
        'dt4 = dt

        'cboChucvuQH.DataBindings.Add()
        cboChucvuHT.DataSource = listDanhmuc(14)
        cboChucvuHT.SelectedValue = IdChucVuHT
        cboChucvuQH.DataSource = listDanhmuc(45, True)
        cboChucvuLK.DataSource = listDanhmuc(45, True)
        cboChucvuTT.DataSource = listDanhmuc(45, True)
    End Sub

    Private Sub bindCboChuyenMonQH()
        cboChuyenmonKH.DataSource = listDanhmuc(12, True, True)
    End Sub

    Private Sub bindCboChinhTriQH()
        cboChinhtriHT.DataSource = listDanhmuc(36, True)
        cboChinhtriKH.DataSource = listDanhmuc(36, True)
    End Sub

    Private Sub bindCboNN_TH()
        cboNgoainguHT.DataSource = listDanhmuc(38, True)
        cboNgoainguKH.DataSource = listDanhmuc(38, True)
        cboTinhocHT.DataSource = listDanhmuc(38, True)
        cboTinhocKH.DataSource = listDanhmuc(38, True)
    End Sub

    Private Sub bindGridQHCB(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridQHCB.AutoGenerateColumns = False
                'Dim m_QHCB As QuyHoachCB = New QuyHoachCB
                'gridQHCB.DataSource = m_QHCB.getAllByCanbo(vIdCanbo)
                gridQHCB.DataSource = listQHCB(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridQHCB.Rows.Count - 1
                    If Trim(gridQHCB.Rows(i).Cells(0).Value) = idQHCB Then
                        gridQHCB.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Quy hoạch cán bộ !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function listQHCB(ByVal vIdCanBo As String) As DataTable
        Try
            Dim strSql As String
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            strSql = "SELECT IdQHCB, Nam_QH, Dot_QH, ((SELECT Ten_goi FROM DanhMuc Where id=IdChucDanh_QH_HT)+' '+DonviQH_HT) as QH_HT, ((SELECT Ten_goi FROM DanhMuc Where id=IdChucDanh_QH_LK)+' '+DonviQH_LK) as QH_LK  FROM HS_QHCB WHERE idCanbo='" & vIdCanBo & "' Order by Nam_QH desc, Dot_QH desc"
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function checkTabQHCB() As String
        Dim strReturn As String = ""
        Try
            If (Val(txtNam.Text) = 0 Or Val(txtNam.Text) < 2003) Then
                txtNam.Focus()
                strReturn = "Năm quy hoạch không hợp lệ. Hãy nhập lại!"
                Exit Try
            End If
            If txtDot.Text = "0" Then
                txtDot.Focus()
                strReturn = "Chưa nhập Đợt!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankQHCB()
        Dim m_QDnhansu As QDNhanSu = New QDNhanSu
        m_QDnhansu = m_QDnhansu.getFinalRecord(idCanBo)
        cboChucvuHT.SelectedValue = m_QDnhansu.IdChucvu_Moi
        cboChuyenmonHT.SelectedValue = m_QDnhansu.IdChuyenMon_Moi
        cboChinhtriHT.SelectedValue = GetValueHT(idCanBo, 1)
        cboNgoainguHT.SelectedValue = GetValueHT(idCanBo, 2)
        cboTinhocHT.SelectedValue = GetValueHT(idCanBo, 3)
        idQHCB = ""
        bntDeleteQH.Enabled = False
        txtNam.Text = Now.Year.ToString
        txtDot.Text = "0"
        txtKetquaHT.Text = ""
        txtKetquaLK.Text = ""
        txtKetquaTT.Text = ""
        txtKetquaLTHT.Text = ""
        txtKetquaLTLK.Text = ""
        txtKetquaLTTT.Text = ""
        txtGhichuQH.Text = ""
        labAlert_QH.Text = ""
        txtNam.Focus()
    End Sub

    Private Sub fillQHCB(ByRef vIdQHCB As String, Optional ByVal vIdCanBo As String = "")
        Dim m_QHCB As QuyHoachCB = New QuyHoachCB
        labAlert_QH.Text = ""
        If vIdCanBo = "" Then
            m_QHCB = m_QHCB.getRecord(vIdQHCB)
        Else
            m_QHCB = m_QHCB.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";106;") < 0) Then
            bntDeleteQH.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteQH.Enabled = True
                Else
                    bntDeleteQH.Enabled = False
                End If
            End If
        End If
        If m_QHCB.IdQHCB <> "" Then
            vIdQHCB = m_QHCB.IdQHCB
            txtNam.Text = m_QHCB.Nam_QH
            txtDot.Text = m_QHCB.Dot_QH
            cboChucvuHT.SelectedValue = m_QHCB.IdChucDanh_HT
            cboChuyenmonHT.SelectedValue = m_QHCB.IdTrinhDo_CM_HT
            cboChinhtriHT.SelectedValue = m_QHCB.IdTrinhDo_CT_HT
            cboNgoainguHT.SelectedValue = m_QHCB.TrinhDo_NN_HT
            cboTinhocHT.SelectedValue = m_QHCB.TrinhDo_TH_HT
            cboChucvuQH.SelectedValue = m_QHCB.IdChucDanh_QH_HT
            cboChucvuLK.SelectedValue = m_QHCB.IdChucDanh_QH_LK
            cboChucvuTT.SelectedValue = m_QHCB.IdChucDanh_QH_5Nam
            txtKetquaHT.Text = m_QHCB.KQ_KiemPhieu_HT
            txtKetquaLK.Text = m_QHCB.KQ_KiemPhieu_LK
            txtKetquaTT.Text = m_QHCB.KQ_KiemPhieu_5N
            txtKetquaLTHT.Text = m_QHCB.KQ_KiemPhieu_LTHT
            txtKetquaLTLK.Text = m_QHCB.KQ_KiemPhieu_LTLK
            txtKetquaLTTT.Text = m_QHCB.KQ_KiemPhieu_LT5N
            cboChuyenmonKH.SelectedValue = m_QHCB.IdKH_DTBD_CM
            cboChinhtriKH.SelectedValue = m_QHCB.IdKH_DTBD_CT
            cboNgoainguKH.SelectedValue = m_QHCB.KH_DTBD_NN
            cboTinhocKH.SelectedValue = m_QHCB.KH_DTBD_TH
            txtGhichuQH.Text = m_QHCB.GhiChu
            txtQH_HT.Text = m_QHCB.DonviQH_HT
            txtQH_LK.Text = m_QHCB.DonviQH_LK
            txtQH_TT.Text = m_QHCB.DonviQH_5Nam
        Else
            blankQHCB()
        End If
    End Sub

    Private Function updateQHCB(ByVal vIdQHCB As String) As Boolean
        Try
            Dim m_QHCB As QuyHoachCB = New QuyHoachCB
            m_QHCB.IdQHCB = vIdQHCB
            m_QHCB.Nam_QH = CInt(txtNam.Text)
            m_QHCB.Dot_QH = CInt(txtDot.Text)
            m_QHCB.IdCanBo = idCanBo
            m_QHCB.IdChucDanh_HT = CInt(cboChucvuHT.SelectedValue)
            m_QHCB.IdTrinhDo_CM_HT = CInt(cboChuyenmonHT.SelectedValue)
            m_QHCB.IdTrinhDo_CT_HT = CInt(cboChinhtriHT.SelectedValue)
            m_QHCB.TrinhDo_NN_HT = CInt(cboNgoainguHT.SelectedValue)
            m_QHCB.TrinhDo_TH_HT = CInt(cboTinhocHT.SelectedValue)
            m_QHCB.IdChucDanh_QH_HT = CInt(cboChucvuQH.SelectedValue)
            m_QHCB.IdChucDanh_QH_LK = CInt(cboChucvuLK.SelectedValue)
            m_QHCB.IdChucDanh_QH_5Nam = CInt(cboChucvuTT.SelectedValue)
            m_QHCB.KQ_KiemPhieu_HT = txtKetquaHT.Text
            m_QHCB.KQ_KiemPhieu_LK = txtKetquaLK.Text
            m_QHCB.KQ_KiemPhieu_5N = txtKetquaTT.Text
            m_QHCB.KQ_KiemPhieu_LTHT = txtKetquaLTHT.Text
            m_QHCB.KQ_KiemPhieu_LTLK = txtKetquaLTLK.Text
            m_QHCB.KQ_KiemPhieu_LT5N = txtKetquaLTTT.Text
            m_QHCB.IdKH_DTBD_CM = CInt(cboChuyenmonKH.SelectedValue)
            m_QHCB.IdKH_DTBD_CT = CInt(cboChinhtriKH.SelectedValue)
            m_QHCB.KH_DTBD_NN = CInt(cboNgoainguKH.SelectedValue)
            m_QHCB.KH_DTBD_TH = CInt(cboTinhocKH.SelectedValue)
            m_QHCB.GhiChu = txtGhichuQH.Text
            m_QHCB.DonviQH_HT = standardizeString(txtQH_HT.Text)
            m_QHCB.DonviQH_LK = standardizeString(txtQH_LK.Text)
            m_QHCB.DonviQH_5Nam = standardizeString(txtQH_TT.Text)
            If vIdQHCB <> "" Then
                m_QHCB.Update()
            Else
                m_QHCB.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function GetValueHT(ByVal vIdCanBo As String, ByVal vIsFlag As Byte) As Integer
        Dim _result As Integer = 0
        Dim strSQL As String = ""
        Dim db As DBAccess = New DBAccess
        Select Case vIsFlag
            Case 1
                _result = db.getNumber("SELECT IdTrinhDoCT FROM HS_Canbo WHERE IdCanbo='" & vIdCanBo & "'")
            Case 2      'Lấy thông tin Trình độ cao nhất về Ngoại ngữ
                strSQL = "SELECT a.IdTrinhDo"
                strSQL += " FROM HS_DTVBCC a, DanhMuc b Where a.IdCanBo = '" + vIdCanBo + "' And "
                strSQL += " (a.IdLoaiVBCC = b.id And b.id_goc = 33 and b.Status = 1 And b.ma_so = '3302')"
                strSQL += " And a.HoanThanh = 1 And (a.NgayHH >= GetDate() Or a.NgayHH Is Null)  Order by Ma_so, a.VBCC, a.NamTN Desc"
                Using dt As DataTable = db.SelectDBRows(strSQL)
                    If Not (dt Is Nothing) Then
                        If dt.Rows.Count > 0 Then
                            _result = dt.Rows(0)("IdTrinhDo").ToString().Trim()
                        End If
                    End If
                End Using
                If _result.ToString.Trim = "" Or _result.ToString.Trim = "A" Or _result.ToString.Trim = "B" Then
                    If db.getString("select IdDTVBCC from HS_DTVBCC WHERE IDcanbo='" + vIdCanBo + "' and IdtrinhDo in (SELECT id FROM danhmuc Where id_goc=38 and ma_so in ('3801','3802','3803'))") <> "" Then
                        _result = db.getNumber("SELECT id FROM Danhmuc WHERE id_goc=38 and ma_so='3812'")
                    End If
                End If
            Case 3      'Lấy thông tin Trình độ cao nhất về tin học
                _result = getChinhDo_Max(vIdCanBo, "1113")
        End Select
        Return _result
    End Function

    Private Sub bntNewQH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewQH.Click
        blankQHCB()
    End Sub

    Private Sub bntAddQH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntAddQH.Click
        Try
            Dim lab_ErrQHCB As String = ""
            If (idQHCB <> "") Then
                If (Globals.Roles.IndexOf(";105;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu quy hoạch cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridQHCB_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Else
                lab_ErrQHCB = checkTabQHCB()
                If lab_ErrQHCB <> "" Then
                    MessageBox.Show(lab_ErrQHCB, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    Exit Sub
                End If
                If updateQHCB(idQHCB) Then
                    initTabQHCB()
                    labAlert_QH.Text = "Ghi dữ liệu thành công!"
                Else
                    MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_QH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel_QH.Click
        If idQHCB = "" Then
            blankQHCB()
        Else
            fillQHCB(idQHCB)
        End If
    End Sub

    Private Sub bntDeleteQH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteQH.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridQHCB.Rows.Count > 0) Then
                For i As Int32 = 0 To gridQHCB.Rows.Count - 1
                    If (CType(gridQHCB.Rows(i).Cells("cln_cbQHCB").Value, Boolean) = True) Then
                        arrDel.Add(gridQHCB.Rows(i).Cells("IdQH").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idQHCB <> "") Then arrDel.Add(idQHCB.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_QuyHoachCB As QuyHoachCB = New QuyHoachCB
                            m_QuyHoachCB.IdQHCB = arrDel(i).ToString()
                            m_QuyHoachCB.Delete()
                        Next
                        blankQHCB()
                        bindGridQHCB(idCanBo)
                        labAlert_QH.Text = "Xoá dữ liệu thành công!"
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

    Private Sub bntCloseQH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseQH.Click
        Close()
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNam.TextChanged
        txtNam.Text = Val(txtNam.Text.Trim)
    End Sub

    Private Sub txtDot_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDot.TextChanged
        txtDot.Text = Val(txtDot.Text.Trim)
    End Sub

    Private Sub gridQHCB_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridQHCB.CellClick
        Try
            idQHCB = gridQHCB.CurrentRow.Cells("IdQH").Value.ToString
            fillQHCB(idQHCB)
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub gridQHCB_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridQHCB.CellFormatting
    '    Try
    '        If gridQHCB.Columns(e.ColumnIndex).Name = "IdChucDanh_QH_HT" Then
    '            e.Value = getDanhmuc_Name(14, CInt(e.Value))
    '        End If
    '        If gridQHCB.Columns(e.ColumnIndex).Name = "IdChucDanh_QH_LK" Then
    '            e.Value = getDanhmuc_Name(14, CInt(e.Value))
    '        End If
    '    Catch ex As Exception
    '    End Try
    'End Sub

    Private Sub gridQHCB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridQHCB.KeyUp
        gridQHCB_CellClick(sender, Nothing)
    End Sub

    Private Sub txtGhichuQH_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGhichuQH.LostFocus
        txtGhichuQH.Text = standardizeString(txtGhichuQH.Text)
    End Sub

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

   
End Class