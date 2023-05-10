Public Class frmTDKTGroup
    Private dbconn As DBAccess
    Private idKT As String = ""
    Private list As New List(Of String())
    Private actIdDonVi As Integer = IdDONVI
    Private actNam As Integer
    Dim frmDSKT As frmChonDS_KhenThuong

    Private Sub frmTDKTGroup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmTDKTGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bindDSKhenThuong(treeCocau, KT_ChuyenMon)
        treeCocau.ExpandAll()
        init()
    End Sub

    Private Sub treeCocau_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeCocau.AfterSelect
        If e.Action <> TreeViewAction.Unknown Then
            Dim arr() As String
            arr = treeCocau.SelectedNode.Tag.ToString.Split("_")
            Select Case arr(0)
                Case "DV"
                    actIdDonVi = arr(1)
                    If treeCocau.SelectedNode.GetNodeCount(True) = 0 Then
                        If Not treeCocau.SelectedNode.IsExpanded Then
                            getKhenThuong_Nam(treeCocau.SelectedNode, arr(1), KT_ChuyenMon)
                        End If
                    End If
                    treeCocau.SelectedNode.Expand()
                Case "NAM"
                    getKhenThuongList(treeCocau.SelectedNode, arr(1), arr(2), KT_ChuyenMon)
                    actNam = arr(2)
                    treeCocau.SelectedNode.Expand()
                Case "CT"
                    idKT = arr(2)
                    labTitle.Text = infQuyetDinhKT(CInt(arr(3)), CBool(arr(4)), CBool(arr(5)), arr(6), arr(7), CInt(arr(8)), CInt(arr(9)))
                    bindGridKT(CInt(arr(3)), CBool(arr(4)), CBool(arr(5)), arr(6), arr(7), CInt(arr(8)), CInt(arr(9)))
                    fillKhenThuong(idKT)
            End Select
        End If
    End Sub

    Private Sub init()
        cboCap_KT.DataSource = listDanhmuc(4, False, True)
        cboChucVuKyQD.DataSource = listChucVuQDKhenThuong(4)
    End Sub

    Private Function infQuyetDinhKT(ByVal vNamKT As Int32, ByVal vKhenThuong As Boolean, ByVal vDinhKy As Boolean, ByVal vSoQD As String, ByVal vNgayQD As String, ByVal vIdCapKT As Int32, ByVal vIdChucVuKyQD As Int32) As String
        infQuyetDinhKT = ""
        infQuyetDinhKT = "Quyết định " & vSoQD & " ngày " & DateTimeUtil.getShortDate(CDate(vNgayQD)) & IIf(vKhenThuong, " Khen thưởng", " Đề nghị khen thưởng") & " năm " & vNamKT
    End Function

    Private Sub bindGridKT(ByVal vNamKT As Int32, ByVal vKhenThuong As Boolean, ByVal vDinhKy As Boolean, ByVal vSoQD As String, ByVal vNgayQD As String, ByVal vIdCapKT As Int32, ByVal vIdChucVuKyQD As Int32)
        Dim dt As DataTable = New DataTable
        'Dim dt_tmp As DataTable = New DataTable
        dt = getKhenThuongDetail(vKhenThuong, vDinhKy, vNamKT, vSoQD, vNgayQD, vIdCapKT, vIdChucVuKyQD)
        fillGridKT(dt)
        'gridKT.DataSource = dt
        If (Not list Is Nothing Or list.Count > 0) Then list.Clear()
        loadKhenThuongList(dt, list)
    End Sub

    Private Sub reLoad_frmTDKT_DanhSach()
        frmDSKT.Dispose()
        list = frmDSKT.list
        setKhenThuongList(frmDSKT.list)
    End Sub

    Private Sub fillGridKT(ByVal dt As DataTable)
        'Khoi tao gridDataview
        gridKT.DataSource = Nothing
        gridKT.Columns.Clear()
        gridKT.Rows.Clear()
        gridKT.Columns.Add("STT", "STT")
        gridKT.Columns.Add("CN_TTstr", "CN/TT")
        gridKT.Columns.Add("DoiTuong", "Đối tượng")
        gridKT.Columns.Add("DanhHieu", "Danh hiệu")
        gridKT.Columns.Add("NoiDung", "Nội dung")
        gridKT.Columns(0).Width = 40
        gridKT.Columns(1).Width = 80
        gridKT.Columns(2).Width = 250
        gridKT.Columns(3).Width = 400
        gridKT.Columns(4).Width = 500
        If dt.Rows.Count > 0 Then
            Dim tableStyle As DataGridTableStyle = New DataGridTableStyle()
            For i As Integer = 0 To dt.Rows.Count - 1
                gridKT.Rows.Add()
                gridKT.Rows(i).Cells("STT").Value = dt.Rows(i)("STT").ToString().Trim()
                gridKT.Rows(i).Cells("CN_TTstr").Value = dt.Rows(i)("CN_TTstr").ToString().Trim()
                gridKT.Rows(i).Cells("DoiTuong").Value = dt.Rows(i)("DoiTuong").ToString().Trim()
                gridKT.Rows(i).Cells("DanhHieu").Value = dt.Rows(i)("DanhHieu").ToString().Trim()
                gridKT.Rows(i).Cells("NoiDung").Value = dt.Rows(i)("NoiDung").ToString().Trim()
            Next
        End If
    End Sub

    Private Sub fillKhenThuong(ByVal IdKT As String)
        Dim m_KhenThuong As KhenThuong = New KhenThuong
        m_KhenThuong = m_KhenThuong.getRecord(IdKT)
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
        txtNoiDung_KT.Text = m_KhenThuong.NoiDungKT
        txtGhichu_KT.Text = m_KhenThuong.GhiChu
        setKhenThuongList(list)
    End Sub

    Private Sub loadKhenThuongList(ByVal dt As DataTable, ByRef list As List(Of String()))
        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            Dim arrKT(6) As String
            arrKT(0) = dt.Rows(i).Item(3)
            arrKT(1) = dt.Rows(i).Item(5)
            arrKT(2) = dt.Rows(i).Item(6)
            arrKT(3) = dt.Rows(i).Item(7)
            arrKT(4) = ""
            arrKT(5) = dt.Rows(i).Item(1)
            arrKT(6) = dt.Rows(i).Item(9)
            If arrKT(0) = 2 Then
                arrKT(3) = dt.Rows(i).Item(8).ToString.Substring(0, dt.Rows(i).Item(8).ToString.IndexOf(":"))
                arrKT(4) = dt.Rows(i).Item(7)
            End If
            list.Add(arrKT)
        Next
    End Sub

    Private Sub setKhenThuongList(ByVal list As List(Of String()), Optional ByVal viewDanhHieu As Boolean = False)
        Dim i As Integer
        Dim arrDel As ArrayList = New ArrayList()
        Dim arrTS As ArrayList = New ArrayList()
        If Not (list Is Nothing) Then
            If list.Count > 0 Then
                lstKhenThuong.ClearSelected()
                lstKhenThuong.Items.Clear()
                For i = 0 To list.Count - 1
                    Dim v_arr() As String
                    Dim v_name As String = ""
                    v_arr = list.Item(i)
                    If viewDanhHieu Then v_name = getDanhHieu(v_arr(5)) & ": "
                    Select Case v_arr(0)
                        Case 1
                            v_name += getFullNamePath(v_arr(3), False, v_arr(2), v_arr(1), IdDONVI, " _ ")
                        Case 2
                            Dim v_arr_TV() As String
                            Dim k As Integer
                            v_name += v_arr(3)
                            v_arr_TV = v_arr(4).Split("_")
                            If v_arr_TV.Length > 0 Then
                                v_name += " ("
                                For k = 0 To v_arr_TV.Length - 1
                                    v_name += getFullName(v_arr_TV(k), False) & ", "
                                Next
                                v_name = v_name.Substring(0, v_name.Length - 2) & ")"
                            End If
                        Case 3
                            v_name += getPhongPath(v_arr(3), v_arr(1), IdDONVI, " _ ")
                        Case 4
                            v_name += getDonviPath(v_arr(3), IdDONVI, " _ ")
                    End Select
                    lstKhenThuong.Items.Add(v_name)
                Next
            End If
        End If
    End Sub

    Private Function checkKhenThuong() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_KT.Text = "" Then
                txtSoQD_KT.Focus()
                strReturn = "Chưa nhập Số quyết định!"
                Exit Try
            End If
            'If idKT = "" Then
            '    If checkQuyetDinhKhenThuong(CInt(dtpkNam.Text), standardizeString(txtSoQD_KT.Text), CInt(cboCap_KT.SelectedValue)) <> "" Then
            '        txtSoQD_KT.Text = ""
            '        txtSoQD_KT.Focus()
            '        strReturn = "Số quyết định khen thưởng đã tồn tại. Hãy nhập lại!"
            '        Exit Try
            '    End If
            'End If
            If txtNguoiQD_KT.Text = "" Then
                txtNguoiQD_KT.Focus()
                strReturn = "Chưa nhập Người ký quyết định khen thưởng!"
                Exit Try
            End If
            If list Is Nothing Or list.Count = 0 Then
                strReturn = "Chưa chọn danh sách cá nhân/tập thể được khen thưởng!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Function updateKhenThuong(ByVal vIdKT As String) As Boolean
        Try
            Dim m_KhenThuong As KhenThuong = New KhenThuong
            Dim DaDuyet As Boolean = True
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
            m_KhenThuong.NoiDungKT = standardizeString(txtNoiDung_KT.Text)
            m_KhenThuong.KT_ChuyenMon = KT_ChuyenMon
            m_KhenThuong.GhiChu = txtGhichu_KT.Text

            If idKT <> "" Then
                m_KhenThuong.Update()
            Else
                idKT = m_KhenThuong.Add()
            End If
            'delKhenThuong_CT(idKT)
            addKhenThuong_CT(idKT, list, DaDuyet)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub blankForm()
        idKT = ""
        rdKhenThuong.Checked = True
        rdDeNghi.Checked = False
        dtpkNam.Value = Now.Date
        rdDK.Checked = True
        rdDX.Checked = False
        txtSoQD_KT.Text = ""
        dpkNgayQD_KT.Value = Now
        txtNguoiQD_KT.Text = ""
        txtNoiDung_KT.Text = ""
        txtGhichu_KT.Text = ""
        lstKhenThuong.ClearSelected()
        lstKhenThuong.Items.Clear()
        If (Not list Is Nothing Or list.Count > 0) Then list.Clear()
        gridKT.Columns().Clear()
        gridKT.DataSource = Nothing
        labAlert.Text = ""
        labTitle.Text = ""
    End Sub

    Private Function delKhenThuong_CT(ByVal vIdKT As String) As Boolean
        Try
            Dim m_KhenThuong_CT As KhenThuong_CT = New KhenThuong_CT
            m_KhenThuong_CT.IdKhenThuong = vIdKT
            m_KhenThuong_CT.DeleteForIdKhenThuong()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function addKhenThuong_CT(ByVal vIdKT As String, ByVal vList As List(Of String()), ByVal DaDuyet As Boolean) As Boolean
        Try
            Dim i, j As Integer
            Dim m_KhenThuong_CT As KhenThuong_CT
            Dim v_arr(6) As String
            Dim arrNhomCB() As String

            For i = 0 To vList.Count - 1
                Dim IdKhenThuong_CT As String = ""
                m_KhenThuong_CT = New KhenThuong_CT
                v_arr = vList.Item(i)
                m_KhenThuong_CT.IdKhenThuong = vIdKT
                ' DV nao nhap doi tuong khen thuong cho DV do, voi DV cap tren se nhin thay DS Khen thuong cua cap duoi
                m_KhenThuong_CT.IdDonVi_KTCT = v_arr(1)
                m_KhenThuong_CT.IdPhong = v_arr(2)
                m_KhenThuong_CT.IdDanhHieuHinhThuc = v_arr(5)
                m_KhenThuong_CT.CN_TT = v_arr(0)
                m_KhenThuong_CT.DaDuyet = DaDuyet
                m_KhenThuong_CT.GhiChu = v_arr(6) ' noi dung khen thuong chi tiet cho tung doi tuong
                m_KhenThuong_CT.IdKhenThuong_CT_parent = ""
                If v_arr(0) = 2 Then
                    m_KhenThuong_CT.IdCN_TT = 0
                    m_KhenThuong_CT.TenCN_TT = v_arr(3)
                Else
                    m_KhenThuong_CT.IdCN_TT = v_arr(3)
                    m_KhenThuong_CT.TenCN_TT = ""
                End If
                IdKhenThuong_CT = m_KhenThuong_CT.Add()
                If IdKhenThuong_CT <> "" And v_arr(0) = 2 Then
                    arrNhomCB = v_arr(4).Split("_")
                    For j = 0 To arrNhomCB.Length - 1
                        m_KhenThuong_CT.IdKhenThuong_CT_parent = IdKhenThuong_CT
                        m_KhenThuong_CT.IdCN_TT = arrNhomCB(j)
                        m_KhenThuong_CT.Add()
                    Next
                End If
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub lnkChoiseCB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkChoiseCB.LinkClicked
        frmDSKT = New frmChonDS_KhenThuong
        frmDSKT.list = list
        frmDSKT.Progress_Changed = New frmChonDS_KhenThuong.ProgressChangedEventHandler(AddressOf ReLoad_frmTDKT_DanhSach)
        frmDSKT.ShowDialog()
    End Sub

    Private Sub bntCloseKT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntCloseKT.Click
        Close()
    End Sub

    Private Sub bntUpdateKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateKT.Click
        Try
            Dim lab_ErrKT As String = ""
            If (idKT <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";109;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu khen thưởng cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Return
                End If
            End If
            lab_ErrKT = checkKhenThuong()
            If lab_ErrKT <> "" Then
                MessageBox.Show(lab_ErrKT, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateKhenThuong(idKT) Then
                'reload QD khen thuong
                bindGridKT(dtpkNam.Value.Year, IIf(rdKhenThuong.Checked, True, False), IIf(rdDK.Checked, True, False), txtSoQD_KT.Text, DateTimeUtil.getDate(dpkNgayQD_KT.Text), CInt(cboCap_KT.SelectedValue), CInt(cboChucVuKyQD.SelectedValue))
                'reload danh sach cac QD khen thuong
                Dim childNode As TreeNode
                treeCocau.Nodes.Clear()
                bindDSKhenThuong(treeCocau, KT_ChuyenMon)
                For Each childNode In treeCocau.Nodes(0).Nodes
                    Dim arr() As String
                    arr = childNode.Tag.ToString.Split("_")
                    If arr(2) = dtpkNam.Value.Year Then
                        getKhenThuongList(childNode, arr(1), arr(2), KT_ChuyenMon)
                        actNam = arr(2)
                        'treeCocau.SelectedNode.Expand()
                        Exit For
                    End If
                Next
                treeCocau.ExpandAll()
                labAlert.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntNewKT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntNewKT.Click
        blankForm()
    End Sub

    Private Sub bntDeleteKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteKT.Click
        If idKT = "" Then
            MessageBox.Show("Hãy chọn lại Hồ sơ khen thưởng cần xóa !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If MessageBox.Show("Bạn có chắc chắn xoá quyết định khen thưởng này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.OK Then
            Try
                Dim m_KhenThuong As KhenThuong = New KhenThuong
                m_KhenThuong.IdKhenThuong = idKT
                m_KhenThuong.Delete()
                blankForm()
                'reload danh sach cac QD khen thuong
                Dim childNode As TreeNode
                treeCocau.Nodes.Clear()
                bindDSKhenThuong(treeCocau, KT_ChuyenMon)
                For Each childNode In treeCocau.Nodes(0).Nodes
                    Dim arr() As String
                    arr = childNode.Tag.ToString.Split("_")
                    If arr(2) = dtpkNam.Value.Year Then
                        getKhenThuongList(childNode, arr(1), arr(2), KT_ChuyenMon)
                        actNam = arr(2)
                        'treeCocau.SelectedNode.Expand()
                        Exit For
                    End If
                Next
                treeCocau.ExpandAll()
                labAlert.Text = "Xoá dữ liệu thành công!"
            Catch ex As Exception
                MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub bntCancelKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelKT.Click
        If idKT = "" Then
            blankForm()
        Else
            fillKhenThuong(idKT)
        End If
    End Sub

    Private Sub rdDeNghi_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdDeNghi.CheckedChanged
        labCapKT.Text = "Cấp đề nghị"
    End Sub

    Private Sub rdKhenThuong_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdKhenThuong.CheckedChanged
        labCapKT.Text = "Cấp khen thưởng"
    End Sub
End Class