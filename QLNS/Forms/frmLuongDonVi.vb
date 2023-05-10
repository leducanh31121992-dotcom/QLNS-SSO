Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmLuongDonVi
    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private idLuongDonVi As String = ""
    Private init_LDV As Boolean = False

    Private Sub frmLuongDonVi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OverInfLuongDonVi()
        'Thực hiện check quyền thành viên
        If (Globals.Roles.IndexOf(";204;") < 0) Then   ' xem
            gridLuongDonVi.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";205;") < 0) Then  ' Them
            bntNew_CN.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";207;") < 0) Then  ' Xoá
            bntDelete_CN.Enabled = False
        End If
        initLuongDonVi()
    End Sub

    Private Sub frmLuongDonVi_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub OverInfLuongDonVi()
        Dim m_LuongDonVi As LuongDonVi = New LuongDonVi
        Dim m_TienLuong As TienLuong = New TienLuong
        m_LuongDonVi = m_LuongDonVi.getAllByDonVi(IdDONVI)
        m_TienLuong = m_TienLuong.getRecord(m_LuongDonVi.idTienLuong)
        grpDonVi.Text = "Thông số lương hiện hành của đon vị: " & getDonvi(m_LuongDonVi.idChiNhanh)
        labTuNgay.Text = DateTimeUtil.getShortDate(m_LuongDonVi.TuNgay)
        labLuongCoBan.Text = formatMoney(m_TienLuong.LuongCoBan.ToString)
        labHeSoNganh.Text = m_TienLuong.HeSoNganh
    End Sub

    Private Sub initLuongDonVi()
        If Not init_LDV Then
            cboCVNguoiQD_CN.DataSource = listChucVuQuyenRaQD()
            If DONVI = gMaDonViTW Then
                cboDonVi.DataSource = listDonvi()
            Else
                cboDonVi.DataSource = listDonvi(True)
            End If
            bindCboLuongCoBan()
            init_LDV = True
        End If
        bindGridLuongDonVi()
        fillLuongDonVi("", CInt(cboDonVi.SelectedValue))
        cboDonVi.Focus()
    End Sub

    Private Sub bindCboLuongCoBan()
        Dim m_TienLuong As TienLuong = New TienLuong
        cboLuongCoBan.DataSource = m_TienLuong.getAllcurr
    End Sub

    Private Sub bindGridLuongDonVi()
        Try
            gridLuongDonVi.AutoGenerateColumns = False
            Dim m_LuongDonVi As LuongDonVi = New LuongDonVi
            gridLuongDonVi.DataSource = m_LuongDonVi.getAllByDonVi
            Dim i As Integer = 0
            While i <= gridLuongDonVi.Rows.Count - 1
                If Trim(gridLuongDonVi.Rows(i).Cells(0).Value) = idLuongDonVi Then
                    gridLuongDonVi.Rows(i).Selected = True
                    Exit While
                End If
                i = i + 1
            End While
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Thông tin lương đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkTabLuongDonVi() As String
        Dim strReturn As String = ""
        Try
            If txtSoQD_CN.Text = "" Then
                txtSoQD_CN.Focus()
                strReturn = "Chưa nhập Số quyết định !"
                Exit Try
            End If
            If txtNguoiQD_CN.Text = "" Then
                txtNguoiQD_CN.Focus()
                strReturn = "Chưa nhập Người quyết định !"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankLuongDonVi()
        idLuongDonVi = ""
        bntDelete_CN.Enabled = False
        labAlert.Text = ""
        txtSoQD_CN.Text = ""
        txtNguoiQD_CN.Text = ""
        txtGhiChu_CN.Text = ""
        cboDonVi.Focus()
    End Sub

    Private Sub fillLuongDonVi(ByVal vIdLuongDonVi As String, Optional ByVal vIdDonVi As Integer = 0)
        Dim m_LuongDonVi As LuongDonVi = New LuongDonVi
        labAlert.Text = ""
        If vIdLuongDonVi <> "" Then
            m_LuongDonVi = m_LuongDonVi.getRecord(vIdLuongDonVi)
        Else
            m_LuongDonVi = m_LuongDonVi.getAllByDonVi(vIdDonVi)
        End If
        If (Globals.Roles.IndexOf(";207;") < 0) Then  ' Xoá
            bntDelete_CN.Enabled = False
        Else
            bntDelete_CN.Enabled = True
        End If
        If m_LuongDonVi.id <> "" Then
            bntDelete_CN.Enabled = True
            cboDonVi.SelectedValue = m_LuongDonVi.idChiNhanh
            cboLuongCoBan.SelectedValue = m_LuongDonVi.idTienLuong
            txtSoQD_CN.Text = m_LuongDonVi.SoQD
            txtNguoiQD_CN.Text = m_LuongDonVi.NguoiQD
            dpkNgayApDung_CN.Value = m_LuongDonVi.TuNgay
            dpkNgayQD_CN.Value = m_LuongDonVi.NgayQD
            txtGhiChu_CN.Text = m_LuongDonVi.GhiChu
            cboCVNguoiQD_CN.SelectedValue = m_LuongDonVi.IdCV_Nguoi_QD
        Else
            blankLuongDonVi()
        End If
    End Sub

    Private Function updateLuongDonVi(ByVal vIdLuongDonVi As String) As Boolean
        Try
            Dim m_LuongDonVi As LuongDonVi = New LuongDonVi
            m_LuongDonVi.id = vIdLuongDonVi
            m_LuongDonVi.IdCV_Nguoi_QD = CInt(cboCVNguoiQD_CN.SelectedValue)
            m_LuongDonVi.idChiNhanh = CInt(cboDonVi.SelectedValue)
            m_LuongDonVi.idTienLuong = CInt(cboLuongCoBan.SelectedValue)
            m_LuongDonVi.SoQD = txtSoQD_CN.Text
            m_LuongDonVi.NguoiQD = standardizeString(txtNguoiQD_CN.Text)
            m_LuongDonVi.TuNgay = DateTimeUtil.getDate(dpkNgayApDung_CN.Text)
            m_LuongDonVi.NgayQD = DateTimeUtil.getDate(dpkNgayQD_CN.Text)
            m_LuongDonVi.GhiChu = txtGhiChu_CN.Text
            If vIdLuongDonVi <> "" Then
                m_LuongDonVi.Update()
            Else
                m_LuongDonVi.Add()
            End If
            If m_LuongDonVi.id = "" Then Return False
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub gridLuongDonVi_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridLuongDonVi.CellClick
        Try
            'gridLuongDonVi.Rows(e.RowIndex).Selected = True
            idLuongDonVi = gridLuongDonVi.CurrentRow.Cells("idLDV").Value.ToString
            fillLuongDonVi(idLuongDonVi)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridLuongDonVi_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridLuongDonVi.CellFormatting
        Try
            If gridLuongDonVi.Columns(e.ColumnIndex).Name = "idChiNhanh" Then
                e.Value = getDonvi(CInt(e.Value))
            End If
            If gridLuongDonVi.Columns(e.ColumnIndex).Name = "idTienLuong_CN" Then
                Dim m_TienLuong As TienLuong = New TienLuong
                m_TienLuong = m_TienLuong.getRecord(CInt(e.Value))
                e.Value = m_TienLuong.ThongTinChung
            End If
            If gridLuongDonVi.Columns(e.ColumnIndex).Name = "TuNgay" Then
                e.Value = DateTimeUtil.getShortDate(CDate(e.Value))
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdate_CN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate_CN.Click
        Dim lab_Err As String = ""
        If (idLuongDonVi <> "") Then   ' Sửa
            If (Globals.Roles.IndexOf(";206;") < 0) Then
                MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu lương đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                gridLuongDonVi_CellClick(sender, Nothing)
                Return
            End If
        End If
        lab_Err = checkTabLuongDonVi()
        If lab_Err <> "" Then
            MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If updateLuongDonVi(idLuongDonVi) Then
            bindGridLuongDonVi()
            labAlert.Text = "Ghi dữ liệu thành công!"
        Else
            MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub bntNew_CN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew_CN.Click
        blankLuongDonVi()
    End Sub

    Private Sub bntDelete_CN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete_CN.Click
        Try
            If idLuongDonVi <> "" Then
                If MessageBox.Show("Thông tin này có ảnh hưởng đến việc tính lương trong đơn vị. Bạn có chắc chắn xoá thông tin này không ?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Dim m_LuongDonVi As LuongDonVi = New LuongDonVi
                    m_LuongDonVi.id = idLuongDonVi
                    If m_LuongDonVi.Delete Then
                        bindGridLuongDonVi()
                        blankLuongDonVi()
                        labAlert.Text = "Xoá dữ liệu thành công!"
                    Else
                        MessageBox.Show("Xoá dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                End If
            Else
                MessageBox.Show("Không có gì để xoá!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công" & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btnCancel_CN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel_CN.Click
        If idLuongDonVi = "" Then
            blankLuongDonVi()
        Else
            fillLuongDonVi(idLuongDonVi)
        End If
    End Sub

    Private Sub bntClose_CN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose_CN.Click
        Close()
    End Sub

    Private Sub gridLuongDonVi_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridLuongDonVi.KeyUp
        gridLuongDonVi_CellClick(sender, Nothing)
    End Sub
End Class