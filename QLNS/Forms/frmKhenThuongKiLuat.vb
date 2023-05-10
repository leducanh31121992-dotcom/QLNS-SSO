Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core

Public Class frmKhenThuongKiLuat
    Inherits System.Windows.Forms.Form
    Public IDCB_HS_Canbo As String = ""   ' Id Cán bộ truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDDV_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Public IDPB_HS_Canbo As String = 0    ' Id Đơn vị truyền từ menu ngữ cảnh trong HS_Canbo
    Private dbconn As DBAccess
    Private init_KT As Boolean = False
    Private init_KL As Boolean = False
    Private init_L As Boolean = False
    Private init_P As Boolean = False
    Private init_CV As Boolean = False
    Private idCanBo As String = ""
    Private idKT As String = ""
    Private idKT_CT As String = ""
    Private idKL As String = ""
    Private IdBangLuong As Integer = 0
    Private IdNgachLuong As Integer = 0
    Private IdBacLuong As Integer = 0
    Private IdLoaiPC As Integer = 0
    Private IdMucPC As Integer = 0
    Private IdDV As Integer = 0
    Private IdPhong As Integer = 0
    Private IdChucVu As Integer = 0
    Private IdChuyenMon As Integer = 0
    Private Heso As String = ""
    Private idLuong As String = ""
    Private idPhuCap As String = ""
    Private idQDNhanSu As String = ""
    Private ComDset As New DataSet

    Private Sub frmKhenThuongKiLuat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      
        lkl_xemct.Enabled = False
        'Check quyền thành viên
        Dim roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";107;108;109;110;", roles)) Then
            If tabControlKTKL.TabPages.Contains(tabKT) Then tabControlKTKL.TabPages.Remove(tabKT)
        End If
        If Not (Globals.IsIntersect(";111;112;113;114;", roles)) Then
            If tabControlKTKL.TabPages.Contains(tabKL) Then tabControlKTKL.TabPages.Remove(tabKL)
        End If

        If (Globals.Roles.IndexOf(";107;") < 0) Then   ' xem
            gridKT.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";108;") < 0) Then  ' Them
            bntNewKT.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";110;") < 0) Then  ' Xoá
            bntDeleteKT.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";111;") < 0) Then
            gridKL.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";112;") < 0) Then
            bntNewKL.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";114;") < 0) Then
            bntDeleteKL.Enabled = False
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
                initTabKT()
            End If

            If checkRight_CreateRecord(IDCB_HS_Canbo) Then
                bntUpdateKL.Enabled = True
                bntUpdateKT.Enabled = True
                bntDeleteKL.Enabled = True
                bntDeleteKT.Enabled = True
            Else
                bntUpdateKL.Enabled = False
                bntUpdateKT.Enabled = False
                bntDeleteKL.Enabled = False
                bntDeleteKT.Enabled = False
            End If
        End If
    End Sub

    Private Sub frmKhenThuongKiLuat_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
        labAlert_KT.Text = ""
        labAlert_KL.Text = ""
        Select Case arr(0)
            Case "DV"
                blankOverInfCB()
                blankKL()
                blankKT()
                If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not treeCocau.SelectedNode.IsExpanded Then
                        getNode(treeCocau.SelectedNode, arr(1), arr(2))
                    End If
                End If
                treeCocau.SelectedNode.Expand()
            Case "PB"
                blankOverInfCB()
                blankKL()
                blankKT()
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
                If tabControlKTKL.SelectedTab.Name = "tabKT" Then
                    initTabKT()
                Else
                    initTabKL()
                End If

                If idCanBo <> "" Then
                    If checkRight_CreateRecord(idCanBo) Then
                        bntUpdateKL.Enabled = True
                        bntUpdateKT.Enabled = True
                        bntDeleteKL.Enabled = True
                        bntDeleteKT.Enabled = True
                    Else
                        bntUpdateKL.Enabled = False
                        bntUpdateKT.Enabled = False
                        bntDeleteKL.Enabled = False
                        bntDeleteKT.Enabled = False
                    End If
                End If
                
        End Select
    End Sub

    Private Sub tabControlKTKL_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabControlKTKL.SelectedIndexChanged
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

    Private Sub callTab()
        Try
            If tabControlKTKL.SelectedTab.Name = "tabKT" Then
                initTabKT()
            Else
                initTabKL()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub OverInfCB(ByVal vIdCanbo As String, ByVal vMaCb As String)
        Dim dt As System.Data.DataTable
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
        getIdBangNgachBac(idCanBo, IdBangLuong, IdNgachLuong, IdBacLuong, Heso)
        lkl_xemct.Enabled = True
    End Sub

    Private Sub blankOverInfCB()
        idCanBo = ""
        idKT = ""
        idKL = ""
        txtCanbo.Text = ""
        txtMaCB.Text = ""
        txtNgaysinhCB.Text = ""
        txtGioiTinh.Text = ""
        txtCMT.Text = ""
        gridKT.DataSource = Nothing
        gridKL.DataSource = Nothing
        lkl_xemct.Enabled = False
    End Sub
#End Region

#Region "tab Khen Thuong"

    Private Sub initTabKT()
        If Not init_KT Then
            bindCboKT()
            init_KT = True
        End If
        bindGridKT(idCanBo, 0)
        fillKT(idKT, idKT_CT, idCanBo)
        gridKT.Focus()
    End Sub

    Private Sub bindCboKT()
        cboCap_KT.DataSource = listDanhmuc(4, False, True)
        cboChucVuKyQD.DataSource = listChucVuQDKhenThuong(4)
        cboDanhhieuTD.DataSource = listKhenThuong(True, False, KT_ChuyenMon, True)
    End Sub

    Private Sub bindGridKT(ByVal vIdCanbo As String, ByVal vNamKT As Integer)
        Try
            Dim dt As System.Data.DataTable = New System.Data.DataTable
            dt = getKhenThuongDetail_CaNhan(vIdCanbo, KT_ChuyenMon, vNamKT, 1, 1).Tables("tbKTCN")
            gridKT.DataSource = Nothing
            gridKT.Columns.Clear()
            gridKT.Rows.Clear()
            gridKT.Columns.Add("IDKT", "IDKT")
            gridKT.Columns.Add("IDKT_CT", "IDKT_CT")
            Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            ckb_choice.Name = "cln_cbKT"
            ckb_choice.HeaderText = "Chọn"
            ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.Width = 43
            gridKT.Columns.Add(ckb_choice)
            gridKT.Columns.Add("NamKT", "Năm")
            gridKT.Columns.Add("QD", "Quyết định")
            gridKT.Columns.Add("KT", "KT/Đề nghị")
            gridKT.Columns.Add("CapKT", "Cấp khen thưởng")
            gridKT.Columns.Add("DHHT", "Danh hiệu/hình thức")
            gridKT.Columns.Add("NoiDung", "Nội dung khen thưởng")
            gridKT.Columns.Add("NguoiKyQD", "Người kí quyết định/đề nghị")
            gridKT.Columns(0).Visible = False
            gridKT.Columns(1).Visible = False
            gridKT.Columns(3).Width = 50
            gridKT.Columns(4).Width = 200
            gridKT.Columns(5).Width = 100
            gridKT.Columns(6).Width = 150
            gridKT.Columns(7).Width = 300
            gridKT.Columns(8).Width = 300
            gridKT.Columns(9).Width = 250
            If dt.Rows.Count > 0 Then
                Dim tableStyle As DataGridTableStyle = New DataGridTableStyle()
                idKT = dt.Rows(0)("IdKhenThuong").ToString().Trim()
                idKT_CT = dt.Rows(0)("IdKhenThuong_CT").ToString().Trim()
                For i As Integer = 0 To dt.Rows.Count - 1
                    gridKT.Rows.Add()
                    gridKT.Rows(i).Cells("IDKT").Value = dt.Rows(i)("IdKhenThuong").ToString().Trim()
                    gridKT.Rows(i).Cells("IDKT_CT").Value = dt.Rows(i)("IdKhenThuong_CT").ToString().Trim()
                    gridKT.Rows(i).Cells("NamKT").Value = dt.Rows(i)("NamKT").ToString().Trim()
                    gridKT.Rows(i).Cells("QD").Value = dt.Rows(i)("SoQD").ToString().Trim() & " ngày " & DateTimeUtil.getShortDate(CDate(dt.Rows(i)("NgayQD").ToString().Trim()))
                    gridKT.Rows(i).Cells("KT").Value = IIf(dt.Rows(i)("KhenThuong"), "Khen thưởng", "Đề nghị KT")
                    gridKT.Rows(i).Cells("CapKT").Value = dt.Rows(i)("CapKT").ToString().Trim()
                    gridKT.Rows(i).Cells("DHHT").Value = dt.Rows(i)("DanhHieuHinhThuc").ToString().Trim()
                    gridKT.Rows(i).Cells("NoiDung").Value = dt.Rows(i)("NoiDung").ToString().Trim()
                    gridKT.Rows(i).Cells("NguoiKyQD").Value = dt.Rows(i)("ChucVuKyQD").ToString().Trim() & " " & dt.Rows(i)("NguoiKyQD").ToString().Trim()
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Khen thưởng của cán bộ !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabKT() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_KT.Text = "" Then
                txtSoQD_KT.Focus()
                strReturn = "Chưa nhập Số quyết định!"
                Exit Try
            Else
                txtSoQD_KT.Text = txtSoQD_KT.Text.Replace("_", "-")
            End If
            If txtNguoiQD_KT.Text = "" Then
                txtNguoiQD_KT.Focus()
                strReturn = "Chưa nhập Người ký quyết đinh khen thưởng!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankKT()
        idKT = ""
        idKT_CT = ""
        bntDeleteKT.Enabled = False
        txtSoQD_KT.Text = ""
        txtNguoiQD_KT.Text = ""
        txtNoiDung_KT.Text = ""
        labAlert_KT.Text = ""
        txtSoQD_KT.Focus()
    End Sub

    Private Sub fillKT(ByRef vIdKT As String, ByVal vIdKT_CT As String, Optional ByVal vIdCanBo As String = "")
        Dim m_KhenThuong As KhenThuong = New KhenThuong
        Dim m_KhenThuong_CT As KhenThuong_CT = New KhenThuong_CT
        labAlert_KT.Text = ""
        m_KhenThuong = m_KhenThuong.getRecord(vIdKT)
        m_KhenThuong_CT = m_KhenThuong_CT.getRecord(vIdKT_CT)
        If (Globals.Roles.IndexOf(";110;") < 0) Then  ' Xoá
            bntDeleteKT.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteKT.Enabled = True
                Else
                    bntDeleteKT.Enabled = False
                End If
            End If
        End If
        If m_KhenThuong.IdKhenThuong <> "" And m_KhenThuong_CT.IdKhenThuong_CT <> "" Then
            vIdKT = m_KhenThuong.IdKhenThuong
            vIdKT_CT = m_KhenThuong_CT.IdKhenThuong_CT
            If m_KhenThuong.KhenThuong Then
                rdKhenThuong.Checked = True
                rdDeNghi.Checked = False
            Else
                rdKhenThuong.Checked = False
                rdDeNghi.Checked = True
            End If
            Dim vDate As DateTime = New DateTime(CInt(m_KhenThuong.NamKT.ToString()), 1, 1)
            dtpkNam.Value = vDate
            If m_KhenThuong.DinhKy Then
                rdDK.Checked = True
                rdDX.Checked = False
            Else
                rdDK.Checked = False
                rdDX.Checked = True
            End If
            cboCap_KT.SelectedValue = m_KhenThuong.IdCapKT
            txtSoQD_KT.Text = m_KhenThuong.SoQD
            dpkNgayQD_KT.Value = m_KhenThuong.NgayQD
            cboChucVuKyQD.SelectedValue = m_KhenThuong.IdChucVuKyQD
            txtNguoiQD_KT.Text = m_KhenThuong.NguoiKyQD
            cboDanhhieuTD.SelectedValue = m_KhenThuong_CT.IdDanhHieuHinhThuc
            txtNoiDung_KT.Text = m_KhenThuong_CT.GhiChu
        Else
            blankKT()
        End If
    End Sub

    Private Function updateKT(ByVal vIdKT As String, ByVal vIdKT_CT As String) As Boolean
        Try
            Dim DaDuyet As Boolean = True
            Dim m_KhenThuong As KhenThuong = New KhenThuong
            Dim m_KhenThuong_CT As KhenThuong_CT = New KhenThuong_CT

            If vIdKT = "" Then vIdKT = checkQuyetDinhKhenThuong(CInt(dtpkNam.Text), "N'" & standardizeString(txtSoQD_KT.Text) & "'", CInt(cboCap_KT.SelectedValue))
            'Update HS_KhenThuong
            m_KhenThuong.IdKhenThuong = vIdKT
            If rdKhenThuong.Checked Then
                m_KhenThuong.KhenThuong = 1
                DaDuyet = True
            Else
                m_KhenThuong.KhenThuong = 0
                DaDuyet = False
            End If
            If rdDK.Checked Then
                m_KhenThuong.DinhKy = 1
            Else
                m_KhenThuong.DinhKy = 0
            End If
            m_KhenThuong.CN_TT = 0
            m_KhenThuong.NamKT = CInt(dtpkNam.Text)
            m_KhenThuong.SoQD = standardizeString(txtSoQD_KT.Text.Trim)
            m_KhenThuong.NgayQD = DateTimeUtil.getDate(dpkNgayQD_KT.Text)
            m_KhenThuong.NguoiKyQD = standardizeName(txtNguoiQD_KT.Text)
            m_KhenThuong.IdCapKT = CInt(cboCap_KT.SelectedValue)
            m_KhenThuong.IdChucVuKyQD = CInt(cboChucVuKyQD.SelectedValue)
            If vIdKT <> "" Then
                m_KhenThuong.UpdatenotFull()
            Else
                m_KhenThuong.NoiDungKT = txtNoiDung_KT.Text
                m_KhenThuong.KT_ChuyenMon = KT_ChuyenMon
                m_KhenThuong.GhiChu = ""
                vIdKT = m_KhenThuong.Add()
            End If

            'Update vao HS_KhenThuong_CT
            m_KhenThuong_CT.IdKhenThuong_CT = vIdKT_CT
            m_KhenThuong_CT.IdKhenThuong = vIdKT
            m_KhenThuong_CT.IdDanhHieuHinhThuc = CInt(cboDanhhieuTD.SelectedValue)
            m_KhenThuong_CT.GhiChu = standardizeString(txtNoiDung_KT.Text)
            If vIdKT_CT <> "" Then
                m_KhenThuong_CT.UpdatenotFull()
            Else
                Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                m_QDNhanSu = m_QDNhanSu.getFinalRecord(idCanBo)
                m_KhenThuong_CT.IdDonVi_KTCT = m_QDNhanSu.IdDonvi_Moi
                m_KhenThuong_CT.IdPhong = m_QDNhanSu.IdPhong_Moi
                m_KhenThuong_CT.CN_TT = 1
                m_KhenThuong_CT.DaDuyet = DaDuyet
                m_KhenThuong_CT.IdKhenThuong_CT_parent = ""
                m_KhenThuong_CT.IdCN_TT = idCanBo
                m_KhenThuong_CT.TenCN_TT = ""
                vIdKT_CT = m_KhenThuong_CT.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function countDoiThuongKT(ByVal vIdKhenThuong As String) As Integer
        Try
            Return dbconn.getNumber("SELECT count(IdKhenThuong_CT) FROM HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong WHERE b.IdKhenThuong='" & vIdKhenThuong & "'")
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Sub gridKT_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridKT.CellClick
        Try
            idKT = gridKT.CurrentRow.Cells("IDKT").Value.ToString
            idKT_CT = gridKT.CurrentRow.Cells("IDKT_CT").Value.ToString
            fillKT(idKT, idKT_CT, idCanBo)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdateKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateKT.Click
        Try
            Dim lab_ErrKT As String = ""
            If (idKT <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";109;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu khen thưởng cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridKT_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_ErrKT = checkTabKT()
            If lab_ErrKT <> "" Then
                MessageBox.Show(lab_ErrKT, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateKT(idKT, idKT_CT) Then
                bindGridKT(idCanBo, 0)
                labAlert_KT.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntNewKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewKT.Click
        blankKT()
    End Sub

    Private Sub bntDeleteKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteKT.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridKT.Rows.Count > 0) Then
                For i As Int32 = 0 To gridKT.Rows.Count - 1
                    If (CType(gridKT.Rows(i).Cells("cln_cbKT").Value, Boolean) = True) Then
                        arrDel.Add(gridKT.Rows(i).Cells("IDKT_CT").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idKT <> "") Then arrDel.Add(idKT_CT.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = System.Windows.Forms.DialogResult.OK Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_KhenThuong As KhenThuong = New KhenThuong
                            Dim m_KhenThuong_CT As KhenThuong_CT = New KhenThuong_CT
                            m_KhenThuong_CT.IdKhenThuong_CT = arrDel(i).ToString()
                            m_KhenThuong_CT.Delete()
                            ' Xoa QD khen thuong neu so QD khen thuong_CT tuong ung=0
                            If countDoiThuongKT(idKT) = 0 Then
                                m_KhenThuong.IdKhenThuong = idKT
                                m_KhenThuong.Delete()
                            End If
                        Next
                        blankKT()
                        bindGridKT(idCanBo, 0)
                        labAlert_KT.Text = "Xoá dữ liệu thành công!"
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

    Private Sub bntCancelKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelKT.Click
        If idKT = "" Then
            blankKT()
        Else
            fillKT(idKT, idKT_CT, idCanBo)
        End If
    End Sub

    Private Sub bntCloseKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseKT.Click
        Close()
    End Sub

    Private Sub gridKT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridKT.KeyUp
        gridKT_CellClick(sender, Nothing)
    End Sub

#End Region

#Region "tab Ki luat"

    Private Sub initTabKL()
        init_L = False
        init_P = False
        init_CV = False
        If Not init_KL Then
            Dim m_QD As QDNhanSu = New QDNhanSu
            m_QD = m_QD.getFinalRecord(idCanBo, Nothing, True)
            IdDV = m_QD.IdDonvi_Moi
            IdPhong = m_QD.IdPhong_Moi
            IdChucVu = m_QD.IdChucvu_Moi
            IdChuyenMon = m_QD.IdChuyenMon_Moi
            getIdBangNgachBac(idCanBo, IdBangLuong, IdNgachLuong, IdBacLuong, Heso)
            getIdLoaiMucPC(idCanBo, IdLoaiPC, IdMucPC)
            cboCap_KL.DataSource = listDanhmuc(4, False, True)
            cboHinhthuc_KL.DataSource = listDanhmuc(5)
            init_KL = True
        Else
            cboHinhthuc_KL.SelectedIndex = 1
        End If
        bindGridKL(idCanBo)
        fillKL(idKL, idCanBo)
        gridKL.Focus()
    End Sub

    Private Sub bindGridKL(ByVal vIdCanbo As String)
        Try
            If vIdCanbo <> "" Then
                gridKL.AutoGenerateColumns = False
                Dim m_KiLuatCB As KiLuatCB = New KiLuatCB
                gridKL.DataSource = m_KiLuatCB.getAllByCanbo(vIdCanbo)
                Dim i As Integer = 0
                While i <= gridKL.Rows.Count - 1
                    If Trim(gridKL.Rows(i).Cells(0).Value) = idKL Then
                        gridKL.Rows(i).Selected = True
                        Exit While
                    End If
                    i = i + 1
                End While
            End If
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Kỉ luật của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabKL() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_KL.Text = "" Then
                txtSoQD_KL.Focus()
                strReturn = "Chưa nhập số quyết định!"
                Exit Try
            Else
                txtSoQD_KL.Text = txtSoQD_KL.Text.Replace("_", "-")
            End If
            If idKL = "" Then
                If checkQuyetDinh("KILUAT", idCanBo, txtSoQD_KL.Text) Then
                    txtSoQD_KL.Text = ""
                    txtSoQD_KL.Focus()
                    strReturn = "Số quyết định kỉ luật cán bộ đã tồn tại. Hãy nhập lại!"
                    Exit Try
                End If
            End If
            If txtLydo_KL.Text = "" Then
                txtLydo_KL.Focus()
                strReturn = "Chưa nhập lý do kỉ luật!"
                Exit Try
            End If
            If txtNguoiQD_KL.Text = "" Then
                txtNguoiQD_KL.Focus()
                strReturn = "Chưa nhập người quyết định!"
                Exit Try
            End If
            If cbLuong.Checked Then
                If cboNgach.SelectedValue = 0 Then
                    cboNgach.Focus()
                    strReturn = "Chưa nhập chọn Ngạch lương!"
                    Exit Try
                End If
                If cboBac.SelectedValue = 0 Then
                    cboBac.Focus()
                    strReturn = "Chưa nhập chọn Bậc lương!"
                    Exit Try
                End If
            End If
            If cbPhucap.Checked Then
                If cboLoaiPC.SelectedValue = 0 Then
                    cboLoaiPC.Focus()
                    strReturn = "Chưa nhập chọn Loại phụ cấp!"
                    Exit Try
                End If
                If cboMucPC.SelectedValue = 0 Then
                    cboMucPC.Focus()
                    strReturn = "Chưa nhập chọn Mức phụ cấp!"
                    Exit Try
                End If
            End If
            If cbChucvu.Checked Then
                If cboDonviMoi.SelectedValue = 0 Then
                    cboDonviMoi.Focus()
                    strReturn = "Chưa nhập chọn Đơn vị!"
                    Exit Try
                End If
                If cboPhongMoi.SelectedValue = 0 Then
                    cboPhongMoi.Focus()
                    strReturn = "Chưa nhập chọn Phòng ban!"
                    Exit Try
                End If
                If cboChucvuMoi.SelectedValue = 0 Then
                    cboChucvuMoi.Focus()
                    strReturn = "Chưa nhập chọn Chưc vụ!"
                    Exit Try
                End If
                If cboChuyenmonMoi.SelectedValue = 0 Then
                    cboChuyenmonMoi.Focus()
                    strReturn = "Chưa nhập chọn Chuyên môn!"
                    Exit Try
                End If
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankKL()
        idKL = ""
        idLuong = ""
        idPhuCap = ""
        idQDNhanSu = ""
        pnlLPC.Enabled = False
        bntDeleteKL.Enabled = False
        txtSoQD_KL.Text = ""
        dpkNgayQD_KL.Value = Now
        dpkTuNgay.Value = Now
        dpkDenNgay.Checked = False
        txtNguoiQD_KL.Text = ""
        txtLydo_KL.Text = ""
        txtTrachnhiemVC.Text = ""
        txtSoThang.Text = "0"
        txtGhichu_KL.Text = ""
        pnlLPC.Enabled = False
        cbLuong.Checked = False
        cbPhucap.Checked = False
        cbChucvu.Checked = False
        cboNgach.SelectedValue = 0
        cboBac.SelectedValue = 0
        txtHeso.Text = ""
        cboLoaiPC.SelectedValue = 0
        cboMucPC.SelectedValue = 0
        cboDonviMoi.SelectedValue = 0
        cboPhongMoi.SelectedValue = 0
        cboChuyenmonMoi.SelectedValue = 0
        cboHinhthuc_KL.SelectedValue = 135
        labAlert_KL.Text = ""
        txtSoQD_KL.Focus()
    End Sub

    Private Sub bindCboNgachluong(ByVal vIdBangluong As Integer)
        If vIdBangluong > 0 Then
            cboNgach.DataSource = listNgachLuong(vIdBangluong, True)
        End If
    End Sub

    Private Sub bindCboBacluong(ByVal vIdNgachluong As Integer, Optional ByVal MaxVar As Integer = 0)
        Try
            If vIdNgachluong > 0 Then
                cboBac.DataSource = listBacLuong(vIdNgachluong, MaxVar, True)
                getHeso(CInt(cboBac.SelectedValue.ToString))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bindCboLoaiPC()
        cboLoaiPC.DataSource = listDanhmuc(31, True)
    End Sub

    Private Sub bindCboMucPC(ByVal vIdLoaiPC As Integer)
        If vIdLoaiPC > 0 Then
            cboMucPC.DataSource = listMucPC(vIdLoaiPC, True)
        End If
    End Sub

    Private Sub bindCboPhongMoi(ByVal vIdDonviMoi As Integer)
        Try
            If vIdDonviMoi > 0 Then
                cboPhongMoi.DataSource = listPhong(vIdDonviMoi, True)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bindCboChucVu(ByVal vIdDonviMoi As Integer)
        cboChucvuMoi.DataSource = listChucVu(vIdDonviMoi, True)
    End Sub

    Private Sub InfLuong(ByVal active As Boolean)
        cboNgach.Enabled = active
        cboBac.Enabled = active
        txtHeso.Enabled = active
        If active And Not init_L Then
            init_L = True
            bindCboNgachluong(IdBangLuong)
        End If
        If active Then
            cboNgach.SelectedValue = IdNgachLuong
            cboBac.SelectedValue = IdBacLuong
            txtHeso.Text = Heso
        End If
    End Sub

    Private Sub InfPhucap(ByVal active As Boolean)
        cboLoaiPC.Enabled = active
        cboMucPC.Enabled = active
        If active And Not init_P Then
            init_P = True
            bindCboLoaiPC()
        End If
        If active Then
            cboLoaiPC.SelectedValue = IdLoaiPC
            cboMucPC.SelectedValue = IdMucPC
        End If
    End Sub

    Private Sub InfChucvu(ByVal active As Boolean)
        cboDonviMoi.Enabled = active
        cboPhongMoi.Enabled = active
        cboChucvuMoi.Enabled = active
        cboChuyenmonMoi.Enabled = active
        If active And (Not init_CV) Then
            init_CV = True
            cboDonviMoi.DataSource = listDonvi(False, True, False, False, False, "       ")
            cboChuyenmonMoi.DataSource = listDanhmuc(12, True)
        End If
        If active Then
            cboDonviMoi.SelectedValue = IdDV
            cboPhongMoi.SelectedValue = IdPhong
            cboChucvuMoi.SelectedValue = IdChucVu
            cboChuyenmonMoi.SelectedValue = IdChuyenMon
        End If
    End Sub

    Private Sub fillKL(ByRef vIdKL As String, Optional ByVal vIdCanBo As String = "")
        Dim m_KiLuatCB As KiLuatCB = New KiLuatCB
        Dim m_Luong As LuongCanBo = New LuongCanBo
        Dim m_Phucap As Phucap = New Phucap
        Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
        Dim vIdNghiDinhluong, vIdBangluong, vIdNgachluong, vIdLoaiPC As Integer

        labAlert_KL.Text = ""
        If vIdCanBo = "" Then
            m_KiLuatCB = m_KiLuatCB.getRecord(vIdKL)
        Else
            m_KiLuatCB = m_KiLuatCB.getFinalRecord(vIdCanBo)
        End If
        If (Globals.Roles.IndexOf(";114;") < 0) Then  ' Xoá
            bntDeleteKL.Enabled = False
        Else
            If vIdCanBo <> "" Then
                If checkRight_CreateRecord(vIdCanBo) Then
                    bntDeleteKL.Enabled = True
                Else
                    bntDeleteKL.Enabled = False
                End If
            End If
        End If
        If m_KiLuatCB.IdKiLuat_CB <> "" Then
            vIdKL = m_KiLuatCB.IdKiLuat_CB
            cboHinhthuc_KL.SelectedValue = m_KiLuatCB.IdHinhThucKL
            txtLydo_KL.Text = m_KiLuatCB.LyDo
            dpkTuNgay.Value = m_KiLuatCB.TuNgay
            If m_KiLuatCB.DenNgay = DateTime.MinValue Then
                dpkDenNgay.Checked = False
            Else
                dpkDenNgay.Checked = True
                dpkDenNgay.Value = m_KiLuatCB.DenNgay
            End If
            txtSoQD_KL.Text = m_KiLuatCB.SoQD
            dpkNgayQD_KL.Value = m_KiLuatCB.NgayQD
            txtNguoiQD_KL.Text = m_KiLuatCB.NguoiQD
            cboCap_KL.SelectedValue = m_KiLuatCB.IdCapKiLuat
            txtTrachnhiemVC.Text = m_KiLuatCB.TrachNhiem_VC
            idLuong = m_KiLuatCB.IdLuong
            txtGhichu_KL.Text = m_KiLuatCB.GhiChu

            ' fill thông tin lương trong quyết định
            If idLuong <> "" Then
                If m_KiLuatCB.KeoDaiNangLuong = 0 Then
                    pnlTNangLuong.Enabled = False
                    txtSoThang.Text = "0"
                    pnlLPC.Enabled = True
                    cbLuong.Checked = True
                    m_Luong = m_Luong.getRecord(idLuong)
                    getIdNDBangNgach(m_Luong.IdBacLuong, vIdNghiDinhluong, vIdBangluong, vIdNgachluong)
                    cboNgach.SelectedValue = vIdNgachluong
                    cboBac.SelectedValue = m_Luong.IdBacLuong
                    txtHeso.Text = m_Luong.HeSoLuong
                Else
                    pnlTNangLuong.Enabled = True
                    cbLuong.Checked = False
                    txtSoThang.Text = m_KiLuatCB.KeoDaiNangLuong.ToString
                End If
            Else
                cbLuong.Checked = False
                cboNgach.SelectedValue = 0
                cboBac.SelectedValue = 0
                txtHeso.Text = ""
            End If

            ' fill thông tin phụ cấp liên quan trong quyết định
            idPhuCap = m_Phucap.getID(idCanBo, m_KiLuatCB.TuNgay, m_KiLuatCB.SoQD, m_KiLuatCB.NgayQD)
            If idPhuCap <> "" Then
                If Not pnlLPC.Enabled Then pnlLPC.Enabled = True
                cbPhucap.Checked = True
                m_Phucap = m_Phucap.getRecord(idPhuCap)
                cboMucPC.SelectedValue = m_Phucap.IdMucPC
                vIdLoaiPC = getIdLoaiPC(m_Phucap.IdMucPC)
                cboLoaiPC.SelectedValue = vIdLoaiPC
            Else
                cbPhucap.Checked = False
                cboLoaiPC.SelectedValue = 0
                cboMucPC.SelectedValue = 0
            End If

            ' fill thông tin chức vụ liên quan trong quyết định
            idQDNhanSu = m_QDNhanSu.getID(idCanBo, m_KiLuatCB.TuNgay, m_KiLuatCB.SoQD, m_KiLuatCB.NgayQD)
            If idQDNhanSu <> "" Then
                If Not pnlLPC.Enabled Then pnlLPC.Enabled = True
                cbChucvu.Checked = True
                m_QDNhanSu = m_QDNhanSu.getRecord(idQDNhanSu)
                cboDonviMoi.SelectedValue = m_QDNhanSu.IdDonvi_Moi
                cboPhongMoi.SelectedValue = m_QDNhanSu.IdPhong_Moi
                cboChucvuMoi.SelectedValue = m_QDNhanSu.IdChucvu_Moi
                cboChuyenmonMoi.SelectedValue = m_QDNhanSu.IdChuyenMon_Moi
            Else
                cbChucvu.Checked = False
                cboDonviMoi.SelectedValue = 0
                cboPhongMoi.SelectedValue = 0
                cboChucvuMoi.SelectedValue = 0
                cboChuyenmonMoi.SelectedValue = 0
            End If
        Else
            blankKL()
        End If
    End Sub

    Private Function updateKL(ByVal vIdKL As String, Optional ByVal isSaThai As Boolean = False, Optional ByVal isKeoDaiTGNangLuong As Boolean = False) As Boolean
        Try
            Dim m_KiLuatCB As KiLuatCB = New KiLuatCB
            Dim m_luong As LuongCanBo = New LuongCanBo
            Dim m_PhuCap As Phucap = New Phucap
            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
            Dim m_clsHDLD As clsHS_Hdld = New clsHS_Hdld()
            Dim m_CBThoiViec As clsHS_Hdld.HS_CBThoiViec = New clsHS_Hdld.HS_CBThoiViec()
            Dim dateHL As Date
            Dim dateQD As Date

            dateHL = DateTimeUtil.getDateCurrTime(dpkTuNgay.Text)
            dateQD = DateTimeUtil.getDate(dpkNgayQD_KL.Text)
            m_KiLuatCB.IdKiLuat_CB = vIdKL
            m_KiLuatCB.IdCanBo = idCanBo
            m_KiLuatCB.IdHinhThucKL = CInt(cboHinhthuc_KL.SelectedValue)
            m_KiLuatCB.LyDo = txtLydo_KL.Text
            m_KiLuatCB.TuNgay = DateTimeUtil.getDate(dpkTuNgay.Text)
            If dpkDenNgay.Checked Then
                m_KiLuatCB.DenNgay = DateTimeUtil.getDate(dpkDenNgay.Text)
            Else
                m_KiLuatCB.DenNgay = DateTime.MinValue
            End If
            m_KiLuatCB.SoQD = txtSoQD_KL.Text
            m_KiLuatCB.NgayQD = DateTimeUtil.getDate(dpkNgayQD_KL.Text)
            m_KiLuatCB.NguoiQD = standardizeString(txtNguoiQD_KL.Text)
            m_KiLuatCB.IdCapKiLuat = CInt(cboCap_KL.SelectedValue)
            m_KiLuatCB.TrachNhiem_VC = txtTrachnhiemVC.Text
            m_KiLuatCB.GhiChu = txtGhichu_KL.Text

            If isKeoDaiTGNangLuong Then
                m_luong = m_luong.getFinalRecord(idCanBo, dateHL)
                m_luong.NgayLen_DK = m_luong.NgayLen_DK.AddMonths(CInt(txtSoThang.Text))
                m_luong.SoQD = txtSoQD_KL.Text
                m_luong.NgayQD = dateQD
                m_luong.NguoiQD = standardizeString(txtNguoiQD_KL.Text)
                m_luong.IdCV_Nguoi_QD = IIf(DONVI = "VBSP", 296, 304) ' Tong giam doc or Giam doc
                m_luong.IdLoaiQD = getDanhmuc_ID("4338")
                m_luong.Ngay_Huong = m_KiLuatCB.TuNgay
                m_luong.GhiChu = "Theo QĐ kỉ luật kéo dài thời gian nâng bậc lương. " & txtGhichu_KL.Text
                If idLuong <> "" Then
                    m_luong.IdLuongCB = ""
                    m_luong.Update()
                Else
                    m_luong.Add()
                End If
            Else
                If idLuong <> "" Then
                    m_luong.Delete()
                End If
            End If

            If Not (txtSoThang.Text <> "" And txtSoThang.Text <> "0") Then
                m_KiLuatCB.IdLuong = ""
                m_KiLuatCB.KeoDaiNangLuong = 0
            Else
                If m_luong.IdLuongCB Is Nothing Then m_luong.IdLuongCB = ""
                m_KiLuatCB.IdLuong = m_luong.IdLuongCB
                m_KiLuatCB.KeoDaiNangLuong = CInt(txtSoThang.Text)
            End If

            If idKL <> "" Then
                m_KiLuatCB.Update()
            Else
                idKL = m_KiLuatCB.Add()
            End If
            If m_KiLuatCB.IdKiLuat_CB <> "" Then
                If cbLuong.Checked Then
                    Dim db As DBAccess = New DBAccess
                    Dim maLoaiQDKiLuat As String = ""
                    Dim IdLoaiQDLuong As Integer
                    m_luong.IdLuongCB = idLuong
                    m_luong.IdCanBo = idCanBo
                    m_luong.IdBacLuong = CInt(cboBac.SelectedValue)
                    m_luong.HeSoLuong = CDbl(txtHeso.Text)
                    m_luong.Ngay_Huong = dateHL
                    m_luong.NgayLen_DK = dateHL.AddMonths(getTimeNangBac(CInt(cboNgach.SelectedValue)))
                    m_luong.SoQD = m_KiLuatCB.SoQD
                    m_luong.NgayQD = dateQD
                    m_luong.NguoiQD = standardizeString(txtNguoiQD_KL.Text)
                    m_luong.IdCV_Nguoi_QD = IIf(DONVI = "VBSP", 296, 304) ' Tong giam doc or Giam doc
                    m_luong.IsQD_NHCS = 1
                    m_luong.DVraQD = "Ngân hàng CSXH"
                    m_luong.NoiDung = ""
                    m_luong.CV_NguoiKy_QD = ""
                    maLoaiQDKiLuat = getDanhmuc_MaSo(m_KiLuatCB.IdHinhThucKL)
                    IdLoaiQDLuong = db.getNumber("SELECT ID FROM DanhMuc WHERE ma_so ='43" & (CInt(maLoaiQDKiLuat.Substring(2, 1)) + 3) & maLoaiQDKiLuat.Substring(3, 1) & "' and ma_so in ('4333','4334','4335','4338','4339','4340')")
                    If IdLoaiQDLuong = 0 Then
                        IdLoaiQDLuong = getDanhmuc_ID("4341")
                    End If
                    m_luong.IdLoaiQD = IdLoaiQDLuong
                    m_luong.LoaiQD = ""
                    m_luong.GhiChu = "Kèm theo QĐ kỉ luật. " & txtGhichu_KL.Text
                    If idLuong <> "" Then
                        m_luong.Update()
                    Else
                        m_luong.Add()
                    End If
                    m_KiLuatCB.IdLuong = m_luong.IdLuongCB
                    m_KiLuatCB.Update()
                Else
                    If idLuong <> "" Then
                        m_luong.Delete()
                        m_KiLuatCB.IdLuong = ""
                        m_KiLuatCB.Update()
                    End If
                End If
                If cbPhucap.Checked Then
                    m_PhuCap.IdCB_PhuCap = idPhuCap
                    m_PhuCap.IdCanBo = idCanBo
                    m_PhuCap.IdMucPC = CInt(cboMucPC.SelectedValue)
                    m_PhuCap.TuNgay = dateHL
                    m_PhuCap.DenNgay = DateTime.MinValue
                    m_PhuCap.SoQD = txtSoQD_KL.Text
                    m_PhuCap.NgayQD = dateQD
                    m_PhuCap.NguoiQD = standardizeString(txtNguoiQD_KL.Text)
                    m_PhuCap.IdCV_Nguoi_QD = IIf(DONVI = "VBSP", 296, 304) ' Tong giam doc or Giam doc
                    m_PhuCap.GhiChu = "Kèm theo QĐ kỉ luật. " & txtGhichu_KL.Text
                    m_PhuCap.IsQD_NHCS = 1
                    m_PhuCap.NoiDung = ""
                    m_PhuCap.DVraQD = ""
                    m_PhuCap.CV_NguoiKy_QD = ""
                    If idPhuCap <> "" Then
                        m_PhuCap.Update()
                    Else
                        m_PhuCap.Add()
                    End If
                Else
                    If idPhuCap <> "" Then
                        m_PhuCap.Delete()
                    End If
                End If
                If cbChucvu.Checked Then  ' m_QDNhanSu
                    Dim m_QDNhanSu_Cu As QDNhanSu = New QDNhanSu
                    m_QDNhanSu_Cu = m_QDNhanSu_Cu.getFinalRecord(idCanBo, dateHL)
                    m_QDNhanSu.IdQDNhanSu = idQDNhanSu
                    m_QDNhanSu.IdCanBo = idCanBo
                    m_QDNhanSu.So_QD = txtSoQD_KL.Text
                    m_QDNhanSu.NgayKy_QD = dateQD
                    m_QDNhanSu.NguoiKy_QD = standardizeString(txtNguoiQD_KL.Text)
                    m_QDNhanSu.IdLoaiQD = getDanhmuc_ID("1510")
                    m_QDNhanSu.NgayHL = dateHL
                    m_QDNhanSu.idCV_Nguoiky_QD = IIf(DONVI = "VBSP", 296, 304) ' Tong giam doc or Giam doc
                    m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
                    m_QDNhanSu.NgayThoiLuong = DateTime.MinValue
                    m_QDNhanSu.IdDonvi_Cu = m_QDNhanSu_Cu.IdDonvi_Moi
                    m_QDNhanSu.IdPhong_Cu = m_QDNhanSu_Cu.IdPhong_Moi
                    m_QDNhanSu.IdChucvu_Cu = m_QDNhanSu_Cu.IdChucvu_Moi
                    m_QDNhanSu.IdChuyenMon_Cu = m_QDNhanSu_Cu.IdChuyenMon_Moi
                    m_QDNhanSu.IdDonvi_Moi = CInt(cboDonviMoi.SelectedValue)
                    m_QDNhanSu.IdPhong_Moi = CInt(cboPhongMoi.SelectedValue)
                    m_QDNhanSu.IdChucvu_Moi = CInt(cboChucvuMoi.SelectedValue)
                    m_QDNhanSu.IdChuyenMon_Moi = CInt(cboChuyenmonMoi.SelectedValue)
                    m_QDNhanSu.Active = True
                    m_QDNhanSu.IsKiemNhiem = False
                    m_QDNhanSu.IsQD_NHCS = 1
                    m_QDNhanSu.DenNgay = DateTime.MinValue
                    m_QDNhanSu.NoiDung = ""
                    m_QDNhanSu.DVraQD = ""
                    m_QDNhanSu.CV_NguoiKy_QD = ""
                    m_QDNhanSu.LoaiQD = ""
                    m_QDNhanSu.GhiChu = "Kèm theo QĐ kỉ luật. " & txtGhichu_KL.Text
                    If idQDNhanSu <> "" Then
                        m_QDNhanSu.Update()
                    Else
                        m_QDNhanSu.Add()
                    End If
                Else
                    If idQDNhanSu <> "" Then
                        m_QDNhanSu.Delete()
                    End If
                End If
                If isSaThai Then   'Thoi Viec
                    m_CBThoiViec.IdCanBo = idCanBo
                    m_CBThoiViec.IdLoaiQD = getDanhmuc_ID("3702")
                    m_CBThoiViec.LyDo = txtLydo_KL.Text.Trim
                    m_CBThoiViec.NgayKy_QD = DateTimeUtil.getDate(dpkNgayQD_KL.Text)
                    m_CBThoiViec.Ngay_HL = DateTimeUtil.getDate(dpkTuNgay.Text)
                    m_CBThoiViec.NguoiKy_QD = standardizeString(txtNguoiQD_KL.Text)
                    m_CBThoiViec.IdCV_NguoKy_QD = 304 ' Giam doc don vi
                    m_CBThoiViec.TroCap_ThoiViec = 0
                    m_CBThoiViec.TroCap_Khac = 0
                    m_CBThoiViec.GhiChu = "Kèm theo QĐ kỉ luật: Sa thải. " & txtTrachnhiemVC.Text.Trim & " " & txtGhichu_KL.Text
                    m_clsHDLD.Insert_StopWork(m_CBThoiViec)
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub cbLuong_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbLuong.CheckedChanged
        InfLuong(cbLuong.Checked)
    End Sub

    Private Sub cbPhucap_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPhucap.CheckedChanged
        InfPhucap(cbPhucap.Checked)
    End Sub

    Private Sub cbChucvu_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbChucvu.CheckedChanged
        InfChucvu(cbChucvu.Checked)
    End Sub

    Private Sub cboNgach_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboNgach.SelectedIndexChanged
        If Not (cboNgach.SelectedValue Is Nothing) Then
            bindCboBacluong(CInt(cboNgach.SelectedValue), IdBacLuong)
        End If
    End Sub

    Private Sub cboBac_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBac.SelectedIndexChanged
        Try
            txtHeso.Text = getHeso(CInt(cboBac.SelectedValue.ToString))
        Catch ex As Exception
            txtHeso.Text = ""
        End Try
    End Sub

    Private Sub cboDonviMoi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDonviMoi.SelectedIndexChanged
        If Not (cboDonviMoi.SelectedValue Is Nothing) Then
            bindCboPhongMoi(CInt(cboDonviMoi.SelectedValue))
            bindCboChucVu(CInt(cboDonviMoi.SelectedValue))
        End If
    End Sub

    Private Sub cboHinhthuc_KL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHinhthuc_KL.SelectedIndexChanged
        If idCanBo <> "" Then
            Dim MaSo_HTKL As String = getDanhmuc_MaSo(cboHinhthuc_KL.SelectedValue)
            If (MaSo_HTKL = "0508") Then
                pnlTNangLuong.Enabled = True
                pnlLPC.Enabled = False
                InfLuong(False)
                InfPhucap(False)
                InfChucvu(False)
            Else
                If (MaSo_HTKL = "0503" Or MaSo_HTKL = "0504" Or MaSo_HTKL = "0505" Or MaSo_HTKL = "0509" Or MaSo_HTKL = "0510") Then
                    pnlLPC.Enabled = True
                    txtSoThang.Text = "0"
                    pnlTNangLuong.Enabled = False
                    InfLuong(cbLuong.Checked)
                    InfPhucap(cbPhucap.Checked)
                    InfChucvu(cbChucvu.Checked)
                Else
                    txtSoThang.Text = "0"
                    pnlTNangLuong.Enabled = False
                    pnlLPC.Enabled = False
                    InfLuong(False)
                    InfPhucap(False)
                    InfChucvu(False)
                    cboNgach.SelectedValue = 0
                    cboLoaiPC.SelectedValue = 0
                    cboDonviMoi.SelectedValue = 0
                End If
            End If
        Else
            pnlLPC.Enabled = False
            pnlTNangLuong.Enabled = False
        End If
    End Sub

    Private Sub cboLoaiPC_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLoaiPC.SelectedValueChanged
        If Not (cboLoaiPC.SelectedValue Is Nothing) Then
            If CInt(cboLoaiPC.SelectedValue) > 0 Then bindCboMucPC(CInt(cboLoaiPC.SelectedValue))
        End If
    End Sub

    Private Sub gridKL_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridKL.CellClick
        Try
            'gridKL.Rows(e.RowIndex).Selected = True
            idKL = gridKL.CurrentRow.Cells("IdKiLuat_CB").Value.ToString
            fillKL(idKL)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridKL_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridKL.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridKL.Columns("TuNgay").Index Then
                row = gridKL.Rows(e.RowIndex)
                e.Value = DateTimeUtil.getShortDate(CDate(row.Cells("TuNgay").Value))
            End If
            If e.ColumnIndex = gridKL.Columns("IdHinhThucKL").Index Then
                row = gridKL.Rows(e.RowIndex)
                e.Value = getDanhmuc_Name(5, CInt(row.Cells("IdHinhThucKL").Value))
            End If
            'If e.ColumnIndex = gridKL.Columns("NgayQD").Index Then
            '    row = gridKL.Rows(e.RowIndex)
            '    e.Value = DateTimeUtil.getShortDate(CDate(row.Cells("NgayQD").Value))
            'End If
            If e.ColumnIndex = gridKL.Columns("IdCapKiLuat").Index Then
                row = gridKL.Rows(e.RowIndex)
                e.Value = getDanhmuc_Name(4, CInt(row.Cells("IdCapKiLuat").Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdateKL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateKL.Click
        Try
            Dim lab_ErrKL As String = ""
            Dim isSaThai As Boolean = False
            Dim isKeoDaiTGNangLuong As Boolean = False
            If (idKL <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";113;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu kỉ luật cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridKL_CellClick(sender, Nothing)
                    Return
                End If
            End If
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_ErrKL = checkTabKL()
            If lab_ErrKL <> "" Then
                MessageBox.Show(lab_ErrKL, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim maso_HTKL As String = getDanhmuc_MaSo(CInt(cboHinhthuc_KL.SelectedValue))
            If maso_HTKL = "0506" Or maso_HTKL = "0507" Then
                isSaThai = True
            End If
            If maso_HTKL = "0508" And txtSoThang.Text.ToString.Trim <> "0" Then
                isKeoDaiTGNangLuong = True
            End If
            If updateKL(idKL, isSaThai, isKeoDaiTGNangLuong) Then
                If isSaThai Then
                    blankOverInfCB()
                    blankKL()
                    treeCocau.Nodes.Clear()
                    bindTreeview(treeCocau)
                Else
                    bindGridKL(idCanBo)
                End If
                labAlert_KL.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntNewKL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewKL.Click
        blankKL()
    End Sub

    Private Sub bntDeleteKL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteKL.Click
        Try
            If idCanBo = "" Then
                MessageBox.Show("Chưa chọn Cán bộ !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Dim arrDel As ArrayList = New ArrayList()
            If (gridKL.Rows.Count > 0) Then
                For i As Int32 = 0 To gridKL.Rows.Count - 1
                    If (CType(gridKL.Rows(i).Cells("cln_cbKL").Value, Boolean) = True) Then
                        arrDel.Add(gridKL.Rows(i).Cells("IdKiLuat_CB").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idKL <> "") Then arrDel.Add(idKL.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = System.Windows.Forms.DialogResult.OK Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_KiLuatCB As KiLuatCB = New KiLuatCB
                            m_KiLuatCB.IdKiLuat_CB = arrDel(i).ToString()
                            m_KiLuatCB.Delete()
                        Next
                        blankKL()
                        bindGridKL(idCanBo)
                        labAlert_KL.Text = "Xoá dữ liệu thành công!"
                        'MessageBox.Show("Xoá dữ liệu thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
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

    Private Sub bntCancelKL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelKL.Click
        If idKL = "" Then
            blankKL()
        Else
            fillKL(idKL)
        End If
    End Sub

    Private Sub bntCloseKL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseKL.Click
        Close()
    End Sub

    Private Sub txtSoThang_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoThang.TextChanged
        txtSoThang.Text = Val(txtSoThang.Text.Trim)
    End Sub

    Private Sub gridKL_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridKL.KeyUp
        gridKL_CellClick(sender, Nothing)
    End Sub

#End Region

    'Private Sub bntOpenfrmKTs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntOpenfrmKTs.Click
    '    'Dim frm As frmKhenThuongGroup = New frmKhenThuongGroup
    '    Dim frm As frmTDKTGroup = New frmTDKTGroup
    '    frm.ShowDialog()
    'End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Dim frm As frmDS_TDKT = New frmDS_TDKT
    '    frm.ShowDialog()
    'End Sub

    Private Sub bntExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExport.Click
        Try
            Cursor = Cursors.WaitCursor
            Load_Excel_Details()
            Cursor = Cursors.Default
        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub Load_Excel_Details()

        Dim filename As String
        'Dim ds As DataSet
        Dim col, row As Integer

        'Dim conn As SqlClient.SqlConnection = dbconn.getConnection
        'Dim adp As New SqlClient.SqlDataAdapter(strSQL, conn)
        'ComDset.Reset()
        'adp.Fill(ComDset, "TTbl")
        ComDset = getKhenThuongDetail_CaNhan(idCanBo, KT_ChuyenMon, 0, 0, 1)
        If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
            Exit Sub
        End If
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MessageBox.Show("Excel chưa được cài đặt trên máy. Để tiếp tục yêu cầu cài đặt MS Excel", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If

        Try
            'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            With Excel
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()

                Dim i As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    .cells(1, i).value = ComDset.Tables(0).Columns(col).ColumnName
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    i = 2
                    For row = 0 To ComDset.Tables(0).Rows.Count - 1
                        .Cells(i, k).Value = ComDset.Tables(0).Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next

                SaveToExcel(.ActiveCell.Worksheet, "DS_KhenThuong_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls", True)

            End With
            'Trả lại thiết lập cũ
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            MsgBox(ex.Message)
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

End Class