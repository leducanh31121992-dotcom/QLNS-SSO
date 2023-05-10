Public Class frmKhenThuongGroup

    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private idKT As String = ""
    Private list As New List(Of String)

    Private Sub frmKhenThuongGroup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmKhenThuongGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bindTreeviewFull(treeCocau, True, False, False, False, False)
        treeCocau.ExpandAll()
        init()
    End Sub

    Private Sub treeCocau_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles treeCocau.NodeMouseClick
        If e.Node.Checked Then
            CheckSelected(e.Node, True)
            idKT = ""
        Else
            ParentUnCheckSelected(e.Node)
            CheckSelected(e.Node, False)
        End If
        list = getList(treeCocau.Nodes)
    End Sub

    Private Sub init()
        bindCboKT()
        bindGridKT(CInt(cboDanhhieuTD.SelectedValue))
        fillKT(idKT)
        gridKT.Focus()
    End Sub

    Private Sub bindCboKT()
        cboHinhthuc_KT.DataSource = listDanhmuc(3, False, True)
        cboCap_KT.DataSource = listDanhmuc(4, False, True)
        cboDanhhieuTD.DataSource = listDanhmuc(25, False, True)
    End Sub

    Private Sub bindGridKT(ByVal vIdDanhHieuTD As Integer)
        'Try
        '    If vIdDanhHieuTD > 0 Then
        '        gridKT.AutoGenerateColumns = False
        '        Dim m_KhenThuongCB As KhenThuongCB = New KhenThuongCB
        '        gridKT.DataSource = m_KhenThuongCB.getAllByDanhHieuTD(vIdDanhHieuTD)
        '        Dim i As Integer = 0
        '        While i <= gridKT.Rows.Count - 1
        '            If Trim(gridKT.Rows(i).Cells(0).Value) = idKT Then
        '                gridKT.Rows(i).Selected = True
        '                Exit While
        '            End If
        '            i = i + 1
        '        End While
        '    End If
        'Catch ex As Exception
        '    MessageBox.Show("Không Load được Danh sách Khen thưởng của cán bộ !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        'End Try
    End Sub

    Private Function checkKhenThuong() As String
        Dim strReturn As String = ""
        Dim arr() As String
        Dim canbo As String
        Dim i As Integer
        Try
            If txtSoQD_KT.Text = "" Then
                txtSoQD_KT.Focus()
                strReturn = "Chưa nhập Số quyết định!"
                Exit Try
            End If
            If idKT = "" Then
                For i = 0 To list.Count - 1
                    canbo = list.Item(i)
                    arr = canbo.Split("_")
                    If checkQuyetDinh("KHENTHUONG", arr(1), txtSoQD_KT.Text) Then
                        txtSoQD_KT.Text = ""
                        txtSoQD_KT.Focus()
                        strReturn = "Số quyết định khen thưởng cán bộ " & getFullName(arr(1), False) & " đã tồn tại. Hãy nhập lại!"
                        Exit Try
                    End If
                Next
            End If
            If txtNguoiQD_KT.Text = "" Then
                txtNguoiQD_KT.Focus()
                strReturn = "Chưa nhập Người quyết đinh khen thưởng!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankForm()
        list.Clear()
        treeCocau.TopNode.Checked = False
        CheckSelected(treeCocau.TopNode, False)
        idKT = ""
        bntDeleteKT.Enabled = False
        txtSoQD_KT.Text = ""
        txtNguoiQD_KT.Text = ""
        txtMuc_KT.Text = ""
        txtGhichu_KT.Text = ""
        cbAllDelete.Checked = False
        txtSoQD_KT.Focus()
    End Sub

    Private Sub fillKT(ByRef vIdKT As String)
        'Dim m_KhenThuongCB As KhenThuongCB = New KhenThuongCB
        'm_KhenThuongCB = m_KhenThuongCB.getRecord(vIdKT)
        'If (Globals.Roles.IndexOf(";110;") < 0) Then  ' Xoá
        '    bntDeleteKT.Enabled = False
        'Else
        '    bntDeleteKT.Enabled = True
        'End If
        'If m_KhenThuongCB.IdKhenThuong_CB <> "" Then
        '    vIdKT = m_KhenThuongCB.IdKhenThuong_CB
        '    txtSoQD_KT.Text = m_KhenThuongCB.So_QD
        '    txtNguoiQD_KT.Text = m_KhenThuongCB.Ngay_QD
        '    txtNguoiQD_KT.Text = m_KhenThuongCB.Nguoi_QD
        '    cboCap_KT.SelectedValue = m_KhenThuongCB.IdCapKT
        '    cboHinhthuc_KT.SelectedValue = m_KhenThuongCB.IdHinhThucKT
        '    cboDanhhieuTD.SelectedValue = m_KhenThuongCB.IdDanhHieuTD
        '    dpkNgayQD_KT.Value = m_KhenThuongCB.Ngay_QD
        '    txtMuc_KT.Text = m_KhenThuongCB.MucKT
        '    txtGhichu_KT.Text = m_KhenThuongCB.GhiChu
        'Else
        '    blankForm()
        'End If
    End Sub

    Private Function updateKhenThuong(ByVal vIdKhenThuong As String, ByRef list As List(Of String)) As String
        Dim i As Integer
        Dim arr() As String
        Dim Canbo As String = ""
        Dim strErr As String = ""
        Dim db As DBAccess = New DBAccess
        Try
            For i = 0 To list.Count - 1
                Canbo = list.Item(i)
                arr = Canbo.Split("_")
                If Not updateKT(vIdKhenThuong, arr(1)) Then
                    strErr = IIf(strErr = "", "", ",") & arr(1)
                End If
            Next
            If strErr <> "" Then
                strErr = "Hãy cập nhật lại thông tin khen thưởng cán bộ: " & strErr
                Exit Try
            End If
            list.Clear()
        Catch ex As Exception
            strErr = "Ghi dữ liệu không thành công: " & ex.Message
        End Try
        Return strErr
    End Function

    Private Function updateKT(ByVal vIdKT As String, ByVal vIdCanBo As String) As Boolean
        'Try
        '    Dim m_KhenThuongCB As KhenThuongCB = New KhenThuongCB
        '    m_KhenThuongCB.IdKhenThuong_CB = vIdKT
        '    m_KhenThuongCB.So_QD = txtSoQD_KT.Text
        '    m_KhenThuongCB.Ngay_QD = DateTimeUtil.getDate(dpkNgayQD_KT.Text)
        '    m_KhenThuongCB.Nguoi_QD = standardizeString(txtNguoiQD_KT.Text)
        '    m_KhenThuongCB.IdCanBo = vIdCanBo
        '    m_KhenThuongCB.IdCapKT = CInt(cboCap_KT.SelectedValue)
        '    m_KhenThuongCB.IdHinhThucKT = CInt(cboHinhthuc_KT.SelectedValue)
        '    m_KhenThuongCB.IdDanhHieuTD = CInt(cboDanhhieuTD.SelectedValue)
        '    m_KhenThuongCB.MucKT = txtMuc_KT.Text
        '    m_KhenThuongCB.GhiChu = txtGhichu_KT.Text
        '    If vIdKT <> "" Then
        '        m_KhenThuongCB.Update()
        '    Else
        '        idKT = m_KhenThuongCB.Add()
        '    End If
        '    Return True
        'Catch ex As Exception
        '    Return False
        'End Try
    End Function

    Private Sub cboDanhhieuTD_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDanhhieuTD.SelectedIndexChanged
        bindGridKT(CInt(cboDanhhieuTD.SelectedValue))
    End Sub

    Private Sub bntNewKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewKT.Click
        blankForm()
    End Sub

    Private Sub bntUpdateKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateKT.Click
        Try
            Dim lab_Err As String = ""
            If (list Is Nothing) Then
                MessageBox.Show("Chưa chọn Cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If (list.Count = 0) Then
                MessageBox.Show("Không có Cán bộ nào trong danh sách cập nhật. Hãy chọn cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err = checkKhenThuong()
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err = updateKhenThuong(idKT, list)
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            bindGridKT(CInt(cboDanhhieuTD.SelectedValue))
            CheckSelected(treeCocau.Nodes(0), False)
            MessageBox.Show("Ghi dữ liệu thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        Catch ex As Exception
            MessageBox.Show("Ghi dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancelKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelKT.Click

    End Sub

    Private Sub bntDeleteKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteKT.Click
        'Try
        '    Dim arrDel As ArrayList = New ArrayList()
        '    If (gridKT.Rows.Count > 0) Then
        '        For i As Int32 = 0 To gridKT.Rows.Count - 1
        '            If (CType(gridKT.Rows(i).Cells("cln_cbKT").Value, Boolean) = True) Then
        '                arrDel.Add(gridKT.Rows(i).Cells("IdKhenThuong_CB").Value.ToString())
        '            End If
        '        Next
        '    End If
        '    If (arrDel.Count = 0) And (idKT <> "") Then arrDel.Add(idKT.ToString)
        '    If (arrDel.Count > 0) Then
        '        If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.OK Then
        '            Try
        '                Dim i As Integer
        '                For i = 0 To arrDel.Count - 1
        '                    Dim m_KhenThuongCB As KhenThuongCB = New KhenThuongCB
        '                    m_KhenThuongCB.IdKhenThuong_CB = arrDel(i).ToString()
        '                    m_KhenThuongCB.Delete()
        '                Next
        '                blankForm()
        '                bindGridKT(CInt(cboDanhhieuTD.SelectedValue))
        '                MessageBox.Show("Xoá dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '            Catch ex As Exception
        '                MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        '            End Try
        '        End If
        '    Else
        '        MessageBox.Show("Bạn chưa chọn bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '    End If
        'Catch ex As Exception
        '    MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        'End Try
    End Sub

    Private Sub bntCloseKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseKT.Click
        Close()
    End Sub

    Private Sub cbAllDelete_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAllDelete.CheckedChanged
        If cbAllDelete.Checked Then
            If (gridKT.Rows.Count > 0) Then
                bntDeleteKT.Enabled = True
                For i As Int32 = 0 To gridKT.Rows.Count - 1
                    gridKT.Rows(i).Cells("cln_cbKT").Value = True
                Next
            End If
        Else
            If (gridKT.Rows.Count > 0) Then
                bntDeleteKT.Enabled = False
                For i As Int32 = 0 To gridKT.Rows.Count - 1
                    gridKT.Rows(i).Cells("cln_cbKT").Value = False
                Next
            End If
        End If
    End Sub

    Private Sub gridKT_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridKT.CellClick
        Try
            list.Clear()
            idKT = gridKT.CurrentRow.Cells("IdKhenThuong_CB").Value.ToString
            list.Add(gridKT.CurrentRow.Cells("IdCanBo").Value.ToString)
            fillKT(idKT)
            'If gridLuongBS.Rows(e.RowIndex).Cells("cln_cbBS").Value = True Then
            '    cbAllDelete.Checked = False
            '    cbAllDelete_CheckedChanged(sender, Nothing)
            'End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridKT_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridKT.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridKT.Columns("Ngay_QD").Index Then
                e.FormattingApplied = True
                row = gridKT.Rows(e.RowIndex)
                e.Value = DateTimeUtil.getShortDate(CDate(row.Cells("Ngay_QD").Value))
            End If
            If e.ColumnIndex = gridKT.Columns("IdDanhHieuTD").Index Then
                e.FormattingApplied = True
                row = gridKT.Rows(e.RowIndex)
                e.Value = getDanhmuc_Name(25, CInt(row.Cells("IdDanhHieuTD").Value))
            End If
            If e.ColumnIndex = gridKT.Columns("IdCanBo").Index Then
                row = gridKT.Rows(e.RowIndex)
                e.Value = getFullName(row.Cells("IdCanBo").Value, False)
            End If
            
        Catch ex As Exception

        End Try
    End Sub
End Class