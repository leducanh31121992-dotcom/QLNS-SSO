Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmDaoTao
    Inherits System.Windows.Forms.Form
    Public IDCB_HS_Canbo As String = ""   ' Id Cán bộ truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDDV_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDPB_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Private dbconn As DBAccess
    Private idCanBo As String = ""
    Private idNCKH As String = ""
    Private idDTVBCC As String = ""
    Private idNghiPhep As String = ""
    Private init_N As Boolean
    Private init_D As Boolean
    Dim frmDeTai As frmHS_DeTaiNCKH

    Private Sub frmDaoTao_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
        lkl_xemct.Enabled = False
        'Check quyền thành viên
        Dim roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";143;144;145;146;", roles)) Then
            If tabControlDaotao.TabPages.Contains(tabDTVBCC) Then tabControlDaotao.TabPages.Remove(tabDTVBCC)
        End If
        If Not (Globals.IsIntersect(";147;148;149;150;", roles)) Then
            If tabControlDaotao.TabPages.Contains(tabNCKH) Then tabControlDaotao.TabPages.Remove(tabNCKH)
        End If
        If (Globals.Roles.IndexOf(";143;") < 0) Then   ' xem
            gridDTVBCC.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";144;") < 0) Then  ' Them
            bntNew_DT.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";146;") < 0) Then  ' Xoá
            bntDeleteDT.Enabled = False
        End If

        If (Globals.Roles.IndexOf(";147;") < 0) Then
            gridNCKH.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";148;") < 0) Then
            bntNewNC.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";150;") < 0) Then
            bntDeleteNC.Enabled = False
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
                initTabDTVBCC()
            End If

            If checkRight_CreateRecord(IDCB_HS_Canbo) Then
                bntUpdate_DT.Enabled = True
                bntUpdateNC.Enabled = True
                bntDeleteDT.Enabled = True
                bntDeleteNC.Enabled = True
            Else
                bntUpdate_DT.Enabled = False
                bntUpdateNC.Enabled = False
                bntDeleteDT.Enabled = False
                bntDeleteNC.Enabled = False
            End If
        End If
    End Sub

    Private Sub frmDaoTao_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
        labAlert_DT.Text = ""
        labAlert_NC.Text = ""
        Select Case arr(0)
            Case "DV"
                blankOverInfCB()
                blankDTVBCC()
                blankNCKH()
                If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not treeCocau.SelectedNode.IsExpanded Then
                        getNode(treeCocau.SelectedNode, arr(1), arr(2))
                    End If
                End If
                treeCocau.SelectedNode.Expand()
            Case "PB"
                blankOverInfCB()
                blankDTVBCC()
                blankNCKH()
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
                If tabControlDaotao.SelectedTab.Name = "tabNCKH" Then
                    initTabNCKH()
                Else
                    initTabDTVBCC()
                End If

                If checkRight_CreateRecord(idCanBo) Then
                    bntUpdate_DT.Enabled = True
                    bntUpdateNC.Enabled = True
                    bntDeleteDT.Enabled = True
                    bntDeleteNC.Enabled = True
                Else
                    bntUpdate_DT.Enabled = False
                    bntUpdateNC.Enabled = False
                    bntDeleteDT.Enabled = False
                    bntDeleteNC.Enabled = False
                End If
        End Select
    End Sub

    Private Sub tabControlDaotao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabControlDaotao.SelectedIndexChanged
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
            If tabControlDaotao.SelectedTab.Name = "tabNCKH" Then
                initTabNCKH()
            Else
                initTabDTVBCC()
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
        idNCKH = ""
        idDTVBCC = ""
        txtCanbo.Text = ""
        txtMaCB.Text = ""
        txtNgaysinhCB.Text = ""
        txtGioiTinh.Text = ""
        txtCMT.Text = ""
        gridNCKH.DataSource = Nothing
        gridDTVBCC.DataSource = Nothing
        lkl_xemct.Enabled = False
    End Sub
#End Region

#Region "Tab Nghiên cứu khoa học"
    Private Sub initTabNCKH()
        If Not init_N Then
            bindCboCapDeTai()
            bindCboDeTai(CInt(cboCapDeTai.SelectedValue))
            init_N = True
        End If
        bindGridNCKH(idCanBo)
        fillNCKH(idNCKH, idCanBo)
        gridNCKH.Focus()
    End Sub

    Private Sub bindCboCapDeTai()
        cboCapDeTai.DataSource = listDanhmuc(7, True)
    End Sub

    Private Sub bindCboDeTai(ByVal vIdCapDeTai As Integer)
        cboDeTai.DataSource = listDeTai(vIdCapDeTai)
    End Sub

    Private Sub bindGridNCKH(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridNCKH.AutoGenerateColumns = False
                Dim m_NCKH As NghienCuuKhoaHoc = New NghienCuuKhoaHoc
                gridNCKH.DataSource = m_NCKH.getAllByCanbo(vIdCanbo)
                'fillNCKH(0, vIdCanbo)
                Dim i As Integer = 0
                While i <= gridNCKH.Rows.Count - 1
                    If Trim(gridNCKH.Rows(i).Cells(0).Value) = idNCKH Then
                        gridNCKH.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Nghiên cứu khoa học của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabNCKH() As String
        Dim strReturn As String = ""
        Try
            If cboDeTai.SelectedValue = "0" Then
                cboDeTai.Focus()
                strReturn = "Chưa chọn tên đề tài!"
                Exit Try
            End If
            If idNCKH = "" Then
                If checkDeTaiCB_NCKH(cboDeTai.SelectedValue, idCanBo) Then
                    cboDeTai.Focus()
                    strReturn = "Cán bộ đã tham gia nghiên cứu đề tài khoa học này. Hãy chọn lại!"
                    Exit Try
                End If
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Public Function checkDeTaiCB_NCKH(ByVal vIdDeTai As String, ByVal vIdCanbo As String) As Boolean
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT dbo.existsDeTaiCB_NCKH('" & vIdDeTai & "','" & vIdCanbo & "')")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return CBool(ReturnValue)
    End Function

    Private Sub blankNCKH()
        idNCKH = ""
        bntDeleteNC.Enabled = False
        txtGhichu_NC.Text = ""
        cboDeTai.SelectedValue = 0
        labAlert_NC.Text = ""
        cboDeTai.Focus()
    End Sub

    Private Sub fillNCKH(ByRef vIdNCKH As String, Optional ByVal vIdCanBo As String = "")
        Dim m_NCKH As NghienCuuKhoaHoc = New NghienCuuKhoaHoc
        labAlert_NC.Text = ""
        If vIdCanBo = "" Then
            m_NCKH = m_NCKH.getRecord(vIdNCKH)
        Else
            m_NCKH = m_NCKH.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";150;") < 0) Then  ' Xoá
            bntDeleteNC.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteNC.Enabled = True
                Else
                    bntDeleteNC.Enabled = False
                End If
            End If
        End If
        If m_NCKH.IdCBNCKH <> "" Then
            vIdNCKH = m_NCKH.IdCBNCKH
            cboCapDeTai.SelectedValue = getIdCapDeTai(m_NCKH.IdDeTai)
            cboDeTai.SelectedValue = m_NCKH.IdDeTai
            If m_NCKH.ChuNhiem Then
                rdChuNhiem.Checked = True
                rdThanhVien.Checked = False
            Else
                rdChuNhiem.Checked = False
                rdThanhVien.Checked = True
            End If
            txtGhichu_NC.Text = m_NCKH.GhiChu
        Else
            blankNCKH()
        End If
    End Sub

    Private Function updateNCKH(ByVal vIdNCKH As String) As Boolean
        Try
            Dim m_NCKH As NghienCuuKhoaHoc = New NghienCuuKhoaHoc
            m_NCKH.IdCBNCKH = vIdNCKH
            m_NCKH.IdCanBo = idCanBo
            m_NCKH.IdDeTai = cboDeTai.SelectedValue
            If rdChuNhiem.Checked Then
                m_NCKH.ChuNhiem = 1
            Else
                m_NCKH.ChuNhiem = 0
            End If
            m_NCKH.GhiChu = txtGhichu_NC.Text
            If vIdNCKH <> "" Then
                m_NCKH.Update()
            Else
                idNCKH = m_NCKH.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function listDeTai(ByVal vIdCapDeTai As Integer) As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        If vIdCapDeTai = 0 Then
            strSql = "SELECT '0' as Value, N'0-----Tên đề tài-----0' as Display  UNION SELECT idDeTai as Value, TenDeTai as Display  FROM DeTaiNCKH ORDER BY display"
        Else
            strSql = "SELECT '0' as Value, N'0-----Tên đề tài-----0' as Display  UNION SELECT idDeTai as Value, TenDeTai as Display  FROM DeTaiNCKH  WHERE IdCapDeTai=" & vIdCapDeTai & " ORDER BY display"
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    Private Function getTenDeTai(ByVal vIdDeTai As String) As String
        Dim strValue As String = ""
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        strSql = "SELECT TenDeTai FROM DeTaiNCKH WHERE idDeTai='" & vIdDeTai & "'"
        strValue = dbconn.getString(strSql).Trim
        Return strValue
    End Function

    Private Function getIdCapDeTai(ByVal vIdDeTai As String) As String
        Dim strValue As String = ""
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        strSql = "SELECT IdCapDeTai FROM DeTaiNCKH WHERE idDeTai='" & vIdDeTai & "'"
        strValue = dbconn.getString(strSql).Trim
        Return strValue
    End Function

    Private Sub ReLoad_TabNCKH()
        frmDeTai.Dispose()
        bindCboCapDeTai()
        cboCapDeTai.SelectedValue = frmDeTai.CapDeTai
        bindCboDeTai(frmDeTai.CapDeTai)
        cboDeTai.SelectedValue = frmDeTai.Value
        bindGridNCKH(idCanBo)
    End Sub

    Private Sub cboCapDeTai_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        bindCboDeTai(CInt(cboCapDeTai.SelectedValue))
    End Sub

    Private Sub gridNCKH_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            'gridNCKH.Rows(e.RowIndex).Selected = True
            idNCKH = gridNCKH.CurrentRow.Cells("IdCBNCKH").Value.ToString
            fillNCKH(idNCKH)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridNCKH_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridNCKH.Columns("IdDeTai").Index Then
                e.FormattingApplied = True
                row = gridNCKH.Rows(e.RowIndex)
                e.Value = getTenDeTai(row.Cells("IdDeTai").Value)
            End If
            If e.ColumnIndex = gridNCKH.Columns("ChuNhiem").Index Then
                row = gridNCKH.Rows(e.RowIndex)
                If row.Cells("ChuNhiem").Value Then
                    e.Value = "x"
                Else
                    e.Value = ""
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdateNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateNC.Click
        Try
            Dim lab_ErrNC As String = ""
            If (idNCKH <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";101;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu nghiên cứu khoa học của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridNCKH_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_ErrNC = checkTabNCKH()
            If lab_ErrNC <> "" Then
                MessageBox.Show(lab_ErrNC, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateNCKH(idNCKH) Then
                bindGridNCKH(idCanBo)
                'blankNCKH()
                labAlert_NC.Text = "Ghi dữ liệu thành công!"
                'MessageBox.Show("Ghi dữ liệu thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntNewNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewNC.Click
        blankNCKH()
    End Sub

    Private Sub bntDeleteNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteNC.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridNCKH.Rows.Count > 0) Then
                For i As Int32 = 0 To gridNCKH.Rows.Count - 1
                    If (CType(gridNCKH.Rows(i).Cells("cln_cbNCKH").Value, Boolean) = True) Then
                        arrDel.Add(gridNCKH.Rows(i).Cells("IdCBNCKH").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idNCKH <> "") Then arrDel.Add(idNCKH.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_NCKH As NghienCuuKhoaHoc = New NghienCuuKhoaHoc
                            m_NCKH.IdCBNCKH = arrDel(i).ToString()
                            m_NCKH.Delete()
                        Next
                        blankNCKH()
                        bindGridNCKH(idCanBo)
                        labAlert_NC.Text = "Xoá dữ liệu thành công!"
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

    Private Sub bntCancelNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelNC.Click
        If idNCKH = "" Then
            blankNCKH()
        Else
            fillNCKH(idNCKH)
        End If
    End Sub

    Private Sub bntCloseNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseNC.Click
        Close()
    End Sub

    Private Sub bntOpenfrm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntOpenfrm.Click
        frmDeTai = New frmHS_DeTaiNCKH
        If Not (cboDeTai.SelectedValue Is Nothing) Then
            frmDeTai.Value = cboDeTai.SelectedValue.ToString
            frmDeTai.CapDeTai = CInt(cboCapDeTai.SelectedValue)
        Else
            frmDeTai.Value = ""
            frmDeTai.CapDeTai = 0
        End If
        frmDeTai.Progress_Changed = New frmHS_DeTaiNCKH.ProgressChangedEventHandler(AddressOf ReLoad_TabNCKH)
        frmDeTai.ShowDialog()
    End Sub

    Private Sub gridNCKH_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        gridNCKH_CellClick(sender, Nothing)
    End Sub

#End Region

#Region "tab Dao Tao Van Bang Chung Chi"

    Private Sub initTabDTVBCC()
        If Not init_D Then
            pnlTotNghiep.Enabled = False
            cboLoaiVBCC.DataSource = listDanhmuc(33, False, True)
            cboChuyenNganh.DataSource = listDanhmuc(11, False, True)
            cboHe.DataSource = listDanhmuc(8)
            cboTrinhDo.DataSource = listDanhmuc(38)
            cboNuocDT.DataSource = listQuocGia()
            cboNuocDT.SelectedValue = 1
            bindCboXepLoai()
            init_D = True
        End If
        bindGridDTVBCC(idCanBo)
        fillDTVBCC(idDTVBCC, idCanBo)
        gridDTVBCC.Focus()
    End Sub

    Private Sub bindCboXepLoai()
        cboXepLoai.Items.Add("")
        cboXepLoai.Items.Add("Xuất sắc")
        cboXepLoai.Items.Add("Giỏi")
        cboXepLoai.Items.Add("Khá")
        cboXepLoai.Items.Add("Trung bình khá")
        cboXepLoai.Items.Add("Trung binh")
        cboXepLoai.Items.Add("Không xếp loại")
        If cboXepLoai.SelectedValue = 0 Then
            cboXepLoai.SelectedIndex = 1
        End If
    End Sub

    Private Sub bindGridDTVBCC(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridDTVBCC.AutoGenerateColumns = False
                Dim m_DTVBCC As DaoTaoVanBangChungChi = New DaoTaoVanBangChungChi
                gridDTVBCC.DataSource = m_DTVBCC.getAllByCanbo(vIdCanbo)
                'fillDTVBCC(0, vIdCanbo)
                Dim i As Integer = 0
                While i <= gridDTVBCC.Rows.Count - 1
                    If Trim(gridDTVBCC.Rows(i).Cells(0).Value) = idDTVBCC Then
                        gridDTVBCC.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Quá trình đào tạo văn bằng chứng chỉ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabDTVBCC() As String
        Dim strReturn As String = ""
        Try
            If txtCoSo.Text = "" Then
                txtCoSo.Focus()
                strReturn = "Chưa nhập cơ sở đào tạo!"
                Exit Try
            End If
            If txtNganhHoc.Text = "" Then
                txtNganhHoc.Focus()
                strReturn = "Chưa nhập ngành học!"
                Exit Try
            End If
            If dpkTuNgay_DT.Value > dpkDenNgay_DT.Value Then
                dpkTuNgay_DT.Focus()
                strReturn = "Ngày bắt đầu phải nhỏ hơn hoặc cùng ngày với ngày kết thúc đào tạo. Hãy nhập lại!"
                Exit Try
            End If
            If cbCuDiHoc.Checked Then
                If Not dpkTuNgay_DT.Checked Then
                    dpkTuNgay_DT.Focus()
                    strReturn = "Trường hợp đơn vị cử đi học phải nhập thời gian bắt đầu!"
                    Exit Try
                End If
                If Not dpkDenNgay_DT.Checked Then
                    dpkDenNgay_DT.Focus()
                    strReturn = "Trường hợp đơn vị cử đi học phải nhập thời gian kết thúc!"
                    Exit Try
                End If
            End If
            If cbTotNghiep.Checked Then
                'If cboXepLoai.SelectedIndex = 0 Then
                '    cboXepLoai.Focus()
                '    strReturn = "Chưa chọn xếp loại!"
                '    Exit Try
                'End If
                If txtTenVBCC.Text.Trim = "" Then
                    txtTenVBCC.Focus()
                    strReturn = "Chưa nhập tên VBCC!"
                    Exit Try
                End If
                'If txtSoVBCC.Text.Trim = "" Then
                '    txtSoVBCC.Focus()
                '    strReturn = "Chưa nhập số VBCC!"
                '    Exit Try
                'End If
                If txtNamTN.Text.Trim = "" Then
                    txtNamTN.Focus()
                    strReturn = "Chưa nhập năm tốt nghiệp!"
                    Exit Try
                End If
                If txtNguoiKy.Text.Trim = "" Then
                    txtNguoiKy.Focus()
                    strReturn = "Chưa nhập người ký!"
                    Exit Try
                End If
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankDTVBCC()
        idDTVBCC = ""
        txtCoSo.Text = ""
        txtNganhHoc.Text = ""
        txtLop.Text = ""
        txtTyleHuong.Text = "100"
        cbTotNghiep.Checked = False
        txtTenVBCC.Text = ""
        txtNamTN.Text = ""
        txtSoVBCC.Text = ""
        txtNguoiKy.Text = ""
        txtGhiChu_DT.Text = ""
        bntDeleteDT.Enabled = False
        labAlert_DT.Text = ""
        cbCuDiHoc.Focus()
    End Sub

    Private Sub fillDTVBCC(ByRef vIdDTVBCC As String, Optional ByVal vIdCanBo As String = "")
        Dim m_DTVBCC As DaoTaoVanBangChungChi = New DaoTaoVanBangChungChi

        labAlert_DT.Text = ""
        If vIdCanBo = "" Then
            m_DTVBCC = m_DTVBCC.getRecord(vIdDTVBCC)
        Else
            m_DTVBCC = m_DTVBCC.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";146;") < 0) Then  ' Xoá
            bntDeleteDT.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteDT.Enabled = True
                Else
                    bntDeleteDT.Enabled = False
                End If
            End If
        End If
        If m_DTVBCC.IdDTVBCC <> "" Then
            vIdDTVBCC = m_DTVBCC.IdDTVBCC
            If m_DTVBCC.CuDiHoc Then
                cbCuDiHoc.Checked = True
            Else : cbCuDiHoc.Checked = False
            End If
            If m_DTVBCC.VBCC Then
                rdVanBang.Checked = True
                rdChungChi.Checked = False
            Else
                rdVanBang.Checked = False
                rdChungChi.Checked = True
            End If
            cboLoaiVBCC.SelectedValue = m_DTVBCC.IdLoaiVBCC
            cboHe.SelectedValue = m_DTVBCC.IdHinhThucDT
            txtCoSo.Text = m_DTVBCC.CoSo_DT
            cboNuocDT.SelectedValue = m_DTVBCC.IdNuocDT
            cboChuyenNganh.SelectedValue = m_DTVBCC.IdChuyenNganhDT
            cboTrinhDo.SelectedValue = m_DTVBCC.IdTrinhDo
            txtNganhHoc.Text = m_DTVBCC.NganhHoc
            txtLop.Text = m_DTVBCC.Lop
            txtTyleHuong.Text = m_DTVBCC.TyleHuong.ToString
            If m_DTVBCC.TuNgay = DateTime.MinValue Then
                dpkTuNgay_DT.Checked = False
            Else
                dpkTuNgay_DT.Checked = True
                dpkTuNgay_DT.Value = m_DTVBCC.TuNgay
            End If
            If m_DTVBCC.DenNgay = DateTime.MinValue Then
                dpkDenNgay_DT.Checked = False
            Else
                dpkDenNgay_DT.Checked = True
                dpkDenNgay_DT.Value = m_DTVBCC.DenNgay
            End If
            If m_DTVBCC.HoanThanh Then
                cbTotNghiep.Checked = True
                pnlTotNghiep.Enabled = True
            Else
                cbTotNghiep.Checked = False
                dpkNgayCap.Checked = False
                dpkNgayHL.Checked = False
                dpkNgayHH.Checked = False
                pnlTotNghiep.Enabled = False
            End If
            txtTenVBCC.Text = m_DTVBCC.Ten_VBCC
            txtNamTN.Text = m_DTVBCC.NamTN
            cboXepLoai.SelectedIndex = m_DTVBCC.XepLoai
            txtSoVBCC.Text = m_DTVBCC.So_VBCC
            If m_DTVBCC.NgayCap = DateTime.MinValue Then
                dpkNgayCap.Checked = False
            Else
                dpkNgayCap.Checked = True
                dpkNgayCap.Value = m_DTVBCC.NgayCap
            End If
            txtNguoiKy.Text = m_DTVBCC.NguoiKy
            If m_DTVBCC.NgayHL = DateTime.MinValue Then
                dpkNgayHL.Checked = False
            Else
                dpkNgayHL.Checked = True
                dpkNgayHL.Value = m_DTVBCC.NgayHL
            End If
            If m_DTVBCC.NgayHH = DateTime.MinValue Or m_DTVBCC.NgayHH = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy") Then
                dpkNgayHH.Checked = False
                'If m_DTVBCC.NgayHH = DateTime.MinValue Then
                '    dpkNgayHH.Value = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy")
                'Else
                '    dpkNgayHH.Value = m_DTVBCC.NgayHH
                'End If
            Else
                dpkNgayHH.Checked = True
                dpkNgayHH.Value = m_DTVBCC.NgayHH
            End If
            txtGhiChu_DT.Text = m_DTVBCC.GhiChu
            idNghiPhep = getID_HS_NghiPhep(idCanBo, m_DTVBCC.TuNgay, m_DTVBCC.TyleHuong, m_DTVBCC.CoSo_DT, m_DTVBCC.NganhHoc)
        Else
            blankDTVBCC()
        End If
    End Sub

    Private Function updateDTVBCC(ByVal vIdDTVBCC As String) As Boolean
        Try
            Dim m_DTVBCC As DaoTaoVanBangChungChi = New DaoTaoVanBangChungChi

            m_DTVBCC.IdDTVBCC = vIdDTVBCC
            m_DTVBCC.IdCanBo = idCanBo
            If cbCuDiHoc.Checked Then
                m_DTVBCC.CuDiHoc = 1
                If txtTyleHuong.Text = "" Then txtTyleHuong.Text = "0"
                m_DTVBCC.TyleHuong = CDbl(txtTyleHuong.Text.ToString)
            Else
                m_DTVBCC.CuDiHoc = 0
                m_DTVBCC.TyleHuong = 0
            End If
            If rdVanBang.Checked Then
                m_DTVBCC.VBCC = 1
            Else
                m_DTVBCC.VBCC = 0
            End If
            m_DTVBCC.IdLoaiVBCC = CInt(cboLoaiVBCC.SelectedValue)
            m_DTVBCC.IdHinhThucDT = CInt(cboHe.SelectedValue)
            m_DTVBCC.CoSo_DT = txtCoSo.Text
            m_DTVBCC.IdNuocDT = CInt(cboNuocDT.SelectedValue)
            m_DTVBCC.IdChuyenNganhDT = CInt(cboChuyenNganh.SelectedValue)
            m_DTVBCC.IdTrinhDo = CInt(cboTrinhDo.SelectedValue)
            m_DTVBCC.NganhHoc = txtNganhHoc.Text
            m_DTVBCC.Lop = txtLop.Text
            Globals.Logger.Error("Cập nhật đào tạo/VBCC frmDaoTao\updateDTVBCC => Vào đến 1: ")
            'If dpkTuNgay_DT.Checked Then
            '    m_DTVBCC.TuNgay = DateTimeUtil.getDate(dpkTuNgay_DT.Text)
            'Else
            '    m_DTVBCC.TuNgay = DateTime.MinValue
            'End If
            m_DTVBCC.TuNgay = IIf(dpkTuNgay_DT.Checked, dpkTuNgay_DT.Value, DateTime.Parse("01/01/1900"))

            'If dpkDenNgay_DT.Checked Then
            '    m_DTVBCC.DenNgay = DateTimeUtil.getDate(dpkDenNgay_DT.Text)
            'Else
            '    m_DTVBCC.DenNgay = DateTime.MinValue
            'End If
            m_DTVBCC.DenNgay = IIf(dpkDenNgay_DT.Checked, dpkDenNgay_DT.Value, DateTime.Parse("01/01/1900"))

            Globals.Logger.Error("Cập nhật đào tạo/VBCC frmDaoTao\updateDTVBCC => Vào đến 2: ")
            If cbTotNghiep.Checked Then
                m_DTVBCC.HoanThanh = 1
            Else
                m_DTVBCC.HoanThanh = 0
            End If
            Globals.Logger.Error("Cập nhật đào tạo/VBCC frmDaoTao\updateDTVBCC => Vào đến 3: ")
            m_DTVBCC.Ten_VBCC = txtTenVBCC.Text
            If txtNamTN.Text.Trim = "" Then txtNamTN.Text = "0"
            m_DTVBCC.NamTN = txtNamTN.Text
            m_DTVBCC.XepLoai = CByte(cboXepLoai.SelectedIndex)
            m_DTVBCC.So_VBCC = txtSoVBCC.Text
            m_DTVBCC.NgayCap = IIf(dpkNgayCap.Checked, dpkNgayCap.Value, DateTime.Parse("01/01/1900"))
            m_DTVBCC.NguoiKy = standardizeName(txtNguoiKy.Text)
            m_DTVBCC.NgayHL = IIf(dpkNgayHL.Checked, dpkNgayHL.Value, DateTime.Parse("01/01/1900"))
            If dpkNgayHH.Checked = True Then
                m_DTVBCC.NgayHH = dpkNgayHH.Value
            Else
                m_DTVBCC.NgayHH = DateTimeUtil.StringToDateTime("01/01/1900", "dd/MM/yyyy")
            End If

            m_DTVBCC.GhiChu = txtGhiChu_DT.Text
            If vIdDTVBCC <> "" Then
                m_DTVBCC.Update()
                'Dim m_HSCsLaoDong As clsHS_CsLaodong = New clsHS_CsLaodong
                'Dim m_NghiPhep As clsHS_CsLaodong.HS_NghiPhep = New clsHS_CsLaodong.HS_NghiPhep
                'If m_DTVBCC.CuDiHoc Then
                '    Dim vLydoNghi As String = ""
                '    vLydoNghi = "Cử đi học ngành " & m_DTVBCC.NganhHoc
                '    m_NghiPhep.IdNghiPhep = idNghiPhep
                '    m_NghiPhep.IdCanBo = idCanBo
                '    m_NghiPhep.Nam = dpkTuNgay_DT.Value.Year
                '    m_NghiPhep.TuNgay = m_DTVBCC.TuNgay
                '    m_NghiPhep.DenNgay = m_DTVBCC.DenNgay
                '    m_NghiPhep.SoNgay = countworkdays(m_DTVBCC.TuNgay, m_DTVBCC.DenNgay)
                '    m_NghiPhep.IdLoaiNghi = getDanhmuc_ID("3404")
                '    m_NghiPhep.NoiNghi = m_DTVBCC.CoSo_DT
                '    m_NghiPhep.LyDo = vLydoNghi
                '    m_NghiPhep.TyleHuong = m_DTVBCC.TyleHuong
                '    m_NghiPhep.TienTroCap = 0
                '    m_NghiPhep.NguoiKy_Nghi = ""
                '    m_NghiPhep.GhiChu = m_DTVBCC.GhiChu
                '    If m_NghiPhep.IdNghiPhep <> "" Then
                '        m_HSCsLaoDong.Update_HS_NghiPhep(m_NghiPhep)
                '    Else
                '        m_HSCsLaoDong.Insert_HS_NghiPhep(m_NghiPhep)
                '    End If
                'Else
                '    If idNghiPhep <> "" Then
                '        m_HSCsLaoDong.Delete_HS_NghiPhep(idNghiPhep)
                '        Exit Try
                '    End If
                'End If
            Else
                idDTVBCC = m_DTVBCC.Add()
                If m_DTVBCC.IdDTVBCC = "" Then
                    Return False
                    Exit Function
                End If
            End If
            'If vIdDTVBCC <> "" Then
            '    Dim MaChuyenMon As String = getDanhmuc_MaSo(m_DTVBCC.IdChuyenNganhDT)
            '    If MaChuyenMon = "1122" Then
            '        Dim IdTrinhDoCT_MAX As Integer = 0
            '        Dim MaTrinhDo_MAX As String = ""
            '        Dim IdTrinhDoCT_Moi As Integer = 0
            '        IdTrinhDoCT_MAX = getChinhDo_Max(m_DTVBCC.IdCanBo, "1122")
            '        MaTrinhDo_MAX = getDanhmuc_MaSo(IdTrinhDoCT_MAX)
            '        If MaTrinhDo_MAX = "3807" Then
            '            IdTrinhDoCT_Moi = getDanhmuc_ID("3601")
            '        ElseIf MaTrinhDo_MAX = "3808" Then
            '            IdTrinhDoCT_Moi = getDanhmuc_ID("3602")
            '        ElseIf MaTrinhDo_MAX = "3809" Then
            '            IdTrinhDoCT_Moi = getDanhmuc_ID("3603")
            '        End If
            '        If IdTrinhDoCT_Moi <> 0 Then dbconn.executeSQL("UPDATE HS_Canbo SET IdTrinhDoCT=" & IdTrinhDoCT_Moi & " WHERE IdCanBo='" & m_DTVBCC.IdCanBo & "'")
            '    End If

            'End If

            Return True
        Catch ex As Exception
            Return False
            Globals.Logger.Error("Cập nhật đào tạo/VBCC frmDaoTao\updateDTVBCC: " + ex.Message)
        End Try
        Return True
    End Function

    Private Sub txtTyleHuong_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtTyleHuong.Text.Trim = "" Then
            txtTyleHuong.Text = "0"
        End If
    End Sub

    Private Sub txtTyleHuong_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTyleHuong.TextChanged
        If txtTyleHuong.Text <> "" Then
            txtTyleHuong.Text = formatDouble(txtTyleHuong.Text.Trim)
            txtTyleHuong.SelectionStart = txtTyleHuong.Text.Length
        End If
    End Sub

    Private Sub cbTotNghiep_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTotNghiep.CheckedChanged
        If cbTotNghiep.Checked Then
            pnlTotNghiep.Enabled = True
            If dpkDenNgay_DT.Checked Then txtNamTN.Text = (dpkDenNgay_DT.Value).Year
        Else
            pnlTotNghiep.Enabled = False
        End If
    End Sub

    Private Sub bntUpdate_DT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate_DT.Click
        Try
            Dim lab_ErrDT As String = ""
            If (idDTVBCC <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";145;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu đào tạo văn bằng chứng chỉ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridDTVBCC_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_ErrDT = checkTabDTVBCC()
            If lab_ErrDT <> "" Then
                MessageBox.Show(lab_ErrDT, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateDTVBCC(idDTVBCC) Then
                bindGridDTVBCC(idCanBo)
                'blankDTVBCC()
                labAlert_DT.Text = "Ghi dữ liệu thành công!"
                'MessageBox.Show("Ghi dữ liệu thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Cập nhật thông tin đào tạo frmDaoTao\bntUpdate_DT_Click: " + ex.Message)
        End Try
    End Sub

    Private Sub bntNew_DT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew_DT.Click
        blankDTVBCC()
    End Sub

    Private Sub gridDTVBCC_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridDTVBCC.CellClick
        Try
            'gridDTVBCC.Rows(e.RowIndex).Selected = True
            idDTVBCC = gridDTVBCC.CurrentRow.Cells("Id_DTVBCC").Value.ToString
            fillDTVBCC(idDTVBCC)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridDTVBCC_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridDTVBCC.CellFormatting
        Try
            If gridDTVBCC.Columns(e.ColumnIndex).Name = "TuNgay" Then
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
            If gridDTVBCC.Columns(e.ColumnIndex).Name = "DenNgay" Then
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
            If gridDTVBCC.Columns(e.ColumnIndex).Name = "IdHinhThucDT" Then
                e.Value = getDanhmuc_Name(8, CInt(e.Value))
            End If
            If gridDTVBCC.Columns(e.ColumnIndex).Name = "IdTrinhDo" Then
                e.Value = getDanhmuc_Name(38, CInt(e.Value))
            End If
            If gridDTVBCC.Columns(e.ColumnIndex).Name = "IdChuyenNganhDT" Then
                e.Value = getDanhmuc_Name(11, CInt(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntDeleteDT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteDT.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridDTVBCC.Rows.Count > 0) Then
                For i As Int32 = 0 To gridDTVBCC.Rows.Count - 1
                    If (CType(gridDTVBCC.Rows(i).Cells("cln_cbDTVBCC").Value, Boolean) = True) Then
                        arrDel.Add(gridDTVBCC.Rows(i).Cells("Id_DTVBCC").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idDTVBCC <> "") Then arrDel.Add(idDTVBCC.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_DTVBCC As DaoTaoVanBangChungChi = New DaoTaoVanBangChungChi
                            m_DTVBCC.IdDTVBCC = arrDel(i).ToString()
                            m_DTVBCC.Delete()
                        Next
                        blankDTVBCC()
                        bindGridDTVBCC(idCanBo)
                        labAlert_DT.Text = "Xoá dữ liệu thành công!"
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

    Private Sub bntCancel_DT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel_DT.Click
        If idDTVBCC = "" Then
            blankDTVBCC()
        Else
            fillDTVBCC(idDTVBCC)
        End If
    End Sub

    Private Sub bntClose_DT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose_DT.Click
        Close()
    End Sub

    Private Sub cbCuDiHoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCuDiHoc.CheckedChanged
        If cbCuDiHoc.Checked Then
            txtTyleHuong.Enabled = True
            txtTyleHuong.Text = 100
        Else
            txtTyleHuong.Enabled = False
            txtTyleHuong.Text = 0
        End If
    End Sub

    Private Sub gridDTVBCC_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridDTVBCC.KeyUp
        gridDTVBCC_CellClick(sender, Nothing)
    End Sub

    Private Sub dpkNgayHL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        dpkNgayCap.Value = dpkNgayHL.Value
    End Sub

    Private Sub dpkDenNgay_DT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        If dpkDenNgay_DT.Checked Then txtNamTN.Text = (dpkDenNgay_DT.Value).Year
    End Sub

    Private Sub txtNguoiKy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        txtNguoiKy.Text = standardizeName(txtNguoiKy.Text)
    End Sub

    Private Sub txtGhiChu_DT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        txtGhiChu_DT.Text = standardizeString(txtGhiChu_DT.Text)
    End Sub

    Private Sub txtGhichu_NC_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        txtGhichu_NC.Text = standardizeString(txtGhichu_NC.Text)
    End Sub
#End Region

End Class