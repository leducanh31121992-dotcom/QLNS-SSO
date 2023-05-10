Public Class frmLuongBosungGroup

    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private idBosung As String = ""
    Private Tapsu As Boolean = False
    Private list As New List(Of String)
    Private init_B As Boolean = False

    Private Sub frmLuongBosungGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bindTreeviewFull(treeCocau, True, False, False, False, False)
        treeCocau.ExpandAll()
        init()
    End Sub

    Private Sub frmLuongBosungGroup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub treeCocau_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles treeCocau.NodeMouseClick
        If e.Node.Checked Then
            CheckSelected(e.Node, True)
        Else
            ParentUnCheckSelected(e.Node)
            CheckSelected(e.Node, False)
        End If
        list = getList(treeCocau.Nodes)
    End Sub

    Private Sub init()
        txtThangBS.Text = Now.Month.ToString
        txtNamBS.Text = Now.Year.ToString
        If Not init_B Then
            bindCboLoaiBS()
            init_B = True
        End If
        bindGridLuongBosung(CInt(cboLoaiLuongBS.SelectedValue), CInt(txtNamBS.Text.ToString))
        'txtNamBS.Focus()
        gridLuongBS.Focus()
    End Sub

    Private Sub bindCboLoaiBS()
        cboLoaiLuongBS.DataSource = listDanhmuc(42)
    End Sub

    Private Sub bindGridLuongBosung(ByVal vIdLoaiBS As Integer, ByVal vNam As Integer)
        'Try
        '    If vNam = 0 Then
        '        gridLuongBS.DataSource = Nothing
        '    Else
        '        gridLuongBS.AutoGenerateColumns = False
        '        Dim m_LuongBosung As LuongBosung = New LuongBosung
        '        gridLuongBS.DataSource = m_LuongBosung.getDup_CB_TS(IdDONVI, vIdLoaiBS, vNam)
        '    End If
        '    Dim i As Integer = 0
        '    While i <= gridLuongBS.Rows.Count - 1
        '        If Trim(gridLuongBS.Rows(i).Cells(0).Value) = idBosung Then
        '            gridLuongBS.Rows(i).Selected = True
        '            Exit While
        '        End If
        '        i = i + 1
        '    End While
        '    labListBC.Text = "Danh sách lương bổ sung: " & getDanhmuc_Name(42, CInt(cboLoaiLuongBS.SelectedValue)) & " năm " & txtNamBS.Text
        'Catch ex As Exception
        '    MessageBox.Show("Không Load được Danh sách Lương bổ sung của cán bộ:" & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        'End Try
    End Sub

    Private Function checkForm() As String
        Dim strReturn As String = ""
        Try
            If Not (CInt(txtThangBS.Text.Trim) >= 1 And CInt(txtThangBS.Text.Trim) <= 12) Then
                strReturn = "Hãy nhập lại Tháng bổ sung lương !"
                txtThangBS.Focus()
                Exit Try
            End If

            If txtNamBS.Text.Trim.Length <> 4 Then
                strReturn = "Hãy nhập Năm bổ sung lương (YYYY) !"
                txtNamBS.Focus()
                Exit Try
            End If

            If txtHesochi.Text.Trim = "" Then txtHesochi.Text = "0"
            If txtHesochi.Text.Trim = "0" Then
                If (txtThuclinhBS.Text.Trim = "" Or txtThuclinhBS.Text.Trim = "0") Then
                    strReturn = "Hãy nhập số tiền thực lĩnh!"
                    txtThuclinhBS.Focus()
                    Exit Try
                ElseIf (txtThuclinhBS.Text.Trim.Length > 11) Then
                    strReturn = "Hãy nhập lại số tiền thực lĩnh!"
                    txtThuclinhBS.Focus()
                    Exit Try
                End If
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
        idBosung = ""
        bntDeleteBS.Enabled = False
        txtHesochi.Text = ""
        txtThuclinhBS.Text = ""
        txtGhichuBS.Text = ""
        cbAllDelete.Checked = False
        txtThangBS.Focus()
    End Sub

    Private Sub fillBosung(ByVal vIdBosung As String, ByVal TS As Boolean)
        'bntDeleteBS.Enabled = True
        'If TS Then
        '    Dim m_LuongBosung As LuongBoSung_CBTS = New LuongBoSung_CBTS
        '    m_LuongBosung = m_LuongBosung.getRecord(vIdBosung)
        '    cboLoaiLuongBS.SelectedValue = CInt(m_LuongBosung.IdLoaiLuongBS)
        '    txtThangBS.Text = m_LuongBosung.Thang.ToString
        '    txtNamBS.Text = m_LuongBosung.Nam.ToString
        '    txtHesochi.Text = m_LuongBosung.Hesochi.ToString
        '    txtThuclinhBS.Text = m_LuongBosung.ThucLinh.ToString
        '    txtGhichuBS.Text = m_LuongBosung.GhiChu
        'Else
        '    Dim m_LuongBosung As LuongBosung = New LuongBosung
        '    m_LuongBosung = m_LuongBosung.getRecord(vIdBosung)
        '    cboLoaiLuongBS.SelectedValue = CInt(m_LuongBosung.IdLoaiLuongBS)
        '    txtThangBS.Text = m_LuongBosung.Thang.ToString
        '    txtNamBS.Text = m_LuongBosung.Nam.ToString
        '    txtHesochi.Text = m_LuongBosung.Hesochi.ToString
        '    txtThuclinhBS.Text = m_LuongBosung.ThucLinh.ToString
        '    txtGhichuBS.Text = m_LuongBosung.GhiChu
        'End If
    End Sub

    Private Function updateLuongBoSung(ByVal vIdBosung As String, ByVal list As List(Of String)) As String
        Dim i As Integer
        Dim arr() As String
        Dim Canbo As String = ""
        Dim strErr As String = ""
        Dim db As DBAccess = New DBAccess
        Try
            For i = 0 To list.Count - 1
                Canbo = list.Item(i)
                arr = Canbo.Split("_")
                If Not updateCB(vIdBosung, arr(1), arr(0)) Then
                    strErr = IIf(strErr = "", "", ",") & arr(1)
                End If
            Next
            If strErr <> "" Then strErr = "Hãy cập nhật lại thông tin lương bổ sung cán bộ: " & strErr
        Catch ex As Exception
            strErr = "Ghi dữ liệu không thành công: " & ex.Message
        End Try
        Return strErr
    End Function

    Private Function updateCB(ByVal vIdBosung As String, ByVal vIdCanbo As String, ByVal kind As String) As Boolean
        'Try
        '    If kind = "TS" Then
        '        Dim m_LuongBosung As LuongBoSung_CBTS = New LuongBoSung_CBTS
        '        m_LuongBosung.Id = vIdBosung
        '        m_LuongBosung.IdCanBo = vIdCanbo
        '        m_LuongBosung.IdLoaiLuongBS = CInt(cboLoaiLuongBS.SelectedValue)
        '        m_LuongBosung.Thang = CInt(txtThangBS.Text.ToString)
        '        m_LuongBosung.Nam = CInt(txtNamBS.Text.ToString)
        '        If txtHesochi.Text.Trim = "" Then txtHesochi.Text = "0"
        '        If txtHesochi.Text <> "0" Then
        '            m_LuongBosung.Hesochi = CDbl(txtHesochi.Text.ToString)
        '            m_LuongBosung.ThucLinh = CDbl(txtHesochi.Text) * getLuongCanBo(vIdCanbo, CInt(txtThangBS.Text.ToString), CInt(txtNamBS.Text.ToString), True)
        '        Else
        '            m_LuongBosung.Hesochi = 0
        '            m_LuongBosung.ThucLinh = N2Number(MoneyValue(txtThuclinhBS.Text.ToString))
        '        End If
        '        m_LuongBosung.GhiChu = txtGhichuBS.Text
        '        If vIdBosung <> "" Then
        '            m_LuongBosung.Update()
        '        Else
        '            m_LuongBosung.Add()
        '        End If
        '    Else
        '        Dim m_LuongBosung As LuongBosung = New LuongBosung
        '        m_LuongBosung.Id = vIdBosung
        '        m_LuongBosung.IdCanBo = vIdCanbo
        '        m_LuongBosung.IdLoaiLuongBS = CInt(cboLoaiLuongBS.SelectedValue)
        '        m_LuongBosung.Thang = CInt(txtThangBS.Text.ToString)
        '        m_LuongBosung.Nam = CInt(txtNamBS.Text.ToString)
        '        If txtHesochi.Text.Trim = "" Then txtHesochi.Text = "0"
        '        If txtHesochi.Text <> "0" Then
        '            m_LuongBosung.Hesochi = CDbl(txtHesochi.Text.ToString)
        '            m_LuongBosung.ThucLinh = CDbl(txtHesochi.Text) * getLuongCanBo(vIdCanbo, CInt(txtThangBS.Text.ToString), CInt(txtNamBS.Text.ToString), False)
        '        Else
        '            m_LuongBosung.Hesochi = 0
        '            m_LuongBosung.ThucLinh = N2Number(MoneyValue(txtThuclinhBS.Text.ToString))
        '        End If
        '        m_LuongBosung.GhiChu = txtGhichuBS.Text
        '        If vIdBosung <> "" Then
        '            m_LuongBosung.Update()
        '        Else
        '            m_LuongBosung.Add()
        '        End If
        '    End If
        '    Return True
        'Catch ex As Exception
        '    Return False
        'End Try
    End Function

    Private Sub txtHesochi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHesochi.TextChanged
        If txtHesochi.Text <> "" Then
            txtHesochi.Text = formatDouble(txtHesochi.Text.Trim)
            txtHesochi.SelectionStart = txtHesochi.Text.Length
        End If
    End Sub

    Private Sub txtNamBS_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNamBS.TextChanged
        txtNamBS.Text = Val(txtNamBS.Text.Trim)
    End Sub

    Private Sub txtThangBS_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThangBS.TextChanged
        txtThangBS.Text = Val(txtThangBS.Text.Trim)
    End Sub

    Private Sub txtThuclinhBS_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThuclinhBS.TextChanged
        Try
            txtThuclinhBS = formatMoneyinTextbox(txtThuclinhBS)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdateBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdateBS.Click
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
            lab_Err = checkForm()
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            lab_Err = updateLuongBoSung(idBosung, list)
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            blankForm()
            bindGridLuongBosung(CInt(cboLoaiLuongBS.SelectedValue), CInt(txtNamBS.Text.ToString))
        Catch ex As Exception
            MessageBox.Show("Ghi dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntNewBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNewBS.Click
        blankForm()
    End Sub

    Private Sub bntDeleteBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDeleteBS.Click
        'Try
        '    Dim arrDel As ArrayList = New ArrayList()
        '    Dim arrTS As ArrayList = New ArrayList()
        '    If (gridLuongBS.Rows.Count > 0) Then
        '        For i As Int32 = 0 To gridLuongBS.Rows.Count - 1
        '            If (CType(gridLuongBS.Rows(i).Cells("cln_cbBS").Value, Boolean) = True) Then
        '                arrDel.Add(gridLuongBS.Rows(i).Cells("id").Value.ToString())
        '                arrTS.Add(gridLuongBS.Rows(i).Cells("IdLoaiLuongBS").Value.ToString())
        '            End If
        '        Next
        '    End If
        '    If (arrDel.Count = 0) And (idBosung <> "") Then
        '        arrDel.Add(idBosung.ToString())
        '        arrTS.Add(CByte(Tapsu).ToString())
        '    End If
        '    If (arrDel.Count > 0) Then
        '        If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.OK Then
        '            Try
        '                Dim i As Integer
        '                For i = 0 To arrDel.Count - 1
        '                    If arrTS(i) = 0 Then
        '                        Dim m_LuongBosung As LuongBosung = New LuongBosung
        '                        m_LuongBosung.Id = arrDel(i).ToString()
        '                        m_LuongBosung.Delete()
        '                    Else
        '                        Dim m_LuongBosung As LuongBoSung_CBTS = New LuongBoSung_CBTS
        '                        m_LuongBosung.Id = arrDel(i).ToString()
        '                        m_LuongBosung.Delete()
        '                    End If
        '                Next
        '                blankForm()
        '                bindGridLuongBosung(CInt(cboLoaiLuongBS.SelectedValue), CInt(txtNamBS.Text.ToString))
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

    Private Sub gridLuongBS_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridLuongBS.CellClick
        Try
            list.Clear()
            'gridLuongBS.Rows(e.RowIndex).Selected = True
            idBosung = gridLuongBS.CurrentRow.Cells("id").Value.ToString
            Tapsu = CByte(gridLuongBS.CurrentRow.Cells("IdLoaiLuongBS").Value.ToString)
            If Tapsu = 0 Then
                list.Add("CB_" & gridLuongBS.CurrentRow.Cells("IdCanBo_LBS").Value.ToString)
            Else
                list.Add("TS_" & gridLuongBS.CurrentRow.Cells("IdCanBo_LBS").Value.ToString)
            End If
            fillBosung(idBosung, Tapsu)
            'If gridLuongBS.Rows(e.RowIndex).Cells("cln_cbBS").Value = True Then
            '    cbAllDelete.Checked = False
            '    cbAllDelete_CheckedChanged(sender, Nothing)
            'End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridLuongBS_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridLuongBS.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridLuongBS.Columns("IdLoaiLuongBS").Index Then
                row = gridLuongBS.Rows(e.RowIndex)
                If row.Cells("IdLoaiLuongBS").Value.ToString = "1" Then
                    e.Value = "x"
                Else
                    e.Value = ""
                End If
            End If
            If e.ColumnIndex = gridLuongBS.Columns("IdCanbo").Index Then
                row = gridLuongBS.Rows(e.RowIndex)
                e.Value = getFullName(row.Cells("IdCanbo").Value, CByte(row.Cells("IdLoaiLuongBS").Value.ToString))
            End If
            If e.ColumnIndex = gridLuongBS.Columns("ThucLinh").Index Then
                row = gridLuongBS.Rows(e.RowIndex)
                e.Value = formatMoney(row.Cells("ThucLinh").Value.ToString)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboLoaiLuongBS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLoaiLuongBS.SelectedIndexChanged
        bindGridLuongBosung(CInt(cboLoaiLuongBS.SelectedValue), CInt(txtNamBS.Text.ToString))
    End Sub

    Private Sub bntCancelBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelBS.Click
        If idBosung = "" Then
            blankForm()
        Else
            fillBosung(idBosung, Tapsu)
        End If
    End Sub

    Private Sub bntCloseBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseBS.Click
        Close()
    End Sub

    Private Sub gridLuongBS_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridLuongBS.KeyUp
        gridLuongBS_CellClick(sender, Nothing)
    End Sub

    Private Sub cbAllDelete_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAllDelete.CheckedChanged

        If cbAllDelete.Checked Then
            If (gridLuongBS.Rows.Count > 0) Then
                bntDeleteBS.Enabled = True
                For i As Int32 = 0 To gridLuongBS.Rows.Count - 1
                    gridLuongBS.Rows(i).Cells("cln_cbBS").Value = True
                Next
            End If
        Else
            If (gridLuongBS.Rows.Count > 0) Then
                bntDeleteBS.Enabled = False
                For i As Int32 = 0 To gridLuongBS.Rows.Count - 1
                    gridLuongBS.Rows(i).Cells("cln_cbBS").Value = False
                Next
            End If
        End If
    End Sub
End Class