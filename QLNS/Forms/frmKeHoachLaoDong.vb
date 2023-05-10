Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmKeHoachLaoDong
    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private idKeHoachLD As String = ""
    Private isAll As Boolean = False

    Private Sub frmKeHoachLaoDong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmKeHoachLaoDong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initLuongDonVi()
    End Sub

    Private Sub initLuongDonVi()
        If DONVI = gMaDonViTW Then isAll = True
        cboDonVi.DataSource = listDonvi(True, False, False, True)
        cboPhong.DataSource = listPhong(CInt(cboDonVi.SelectedValue), True)
        bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
        cboDonVi.Focus()
    End Sub

    Private Sub bindGridKeHoachLaoDongDV(ByVal vIdDonVi As Integer, ByVal vNam As Integer)
        Try
            gridKHLD.AutoGenerateColumns = False
            Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
            gridKHLD.DataSource = m_KeHoachLD.getAllByDonVi(vIdDonVi, vNam)
            Dim i As Integer = 0
            While i <= gridKHLD.Rows.Count - 1
                If Trim(gridKHLD.Rows(i).Cells(0).Value) = idKeHoachLD Then
                    gridKHLD.Rows(i).Selected = True
                    Exit While
                End If
                i = i + 1
            End While
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Kế hoạch lao động của đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkForm() As String
        Dim strReturn As String = ""
        Try
            If CInt(cboDonVi.SelectedValue) = 0 Then
                cboDonVi.Focus()
                strReturn = "Chưa chọn chi nhánh!"
                Exit Try
            End If
            If idKeHoachLD = "" Then
                If checkSoKeHoachLaoDong(CInt(cboDonVi.SelectedValue), dtpkNam.Value.Year, numThang.Value) Then
                    strReturn = "Số kế hoạch lao động của đơn vị năm " & dtpkNam.Value.Year & " đã tồn tại. Hãy nhập lại!"
                    Exit Try
                End If
            End If
            If txtSoDaiHan.Text = "" Then
                txtSoDaiHan.Focus()
                strReturn = "Chưa nhập số lao động dài hạn!"
                Exit Try
            End If
            If txtSoNganHan.Text = "" Then
                txtSoNganHan.Focus()
                strReturn = "Chưa nhập số lao động ngắn hạn!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankForm()
        idKeHoachLD = ""
        bntDelete.Enabled = False
        labAlert.Text = ""
        dtpkNam.Value = Now
        numThang.Value = Now.Month
        txtSoDaiHan.Text = ""
        txtSoNganHan.Text = ""
        txtGhiChu.Text = ""
        cboDonVi.Focus()
    End Sub

    Private Sub fillKeHoachLD(ByVal vIdKHLD As String)
        Try
            If vIdKHLD = "" Then
                blankForm()
            Else
                Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
                m_KeHoachLD = m_KeHoachLD.getRecord(vIdKHLD)
                'If (Globals.Roles.IndexOf(";207;") < 0) Then  ' Xoá
                '    bntDelete_CN.Enabled = False
                'Else
                '    bntDelete_CN.Enabled = True
                'End If
                bntDelete.Enabled = True
                labAlert.Text = ""
                Dim vDate As DateTime = New DateTime(CInt(m_KeHoachLD.Nam.ToString()), 1, 1)
                dtpkNam.Value = vDate
                labNam.Text = dtpkNam.Value.Year
                If m_KeHoachLD.Thang = 0 Then
                    numThang.Visible = False
                    rdNam.Checked = True
                Else
                    numThang.Visible = True
                    rdThang.Checked = True
                    numThang.Value = m_KeHoachLD.Thang
                End If

                cboDonVi.SelectedValue = m_KeHoachLD.IdDonVi_KH
                cboPhong.SelectedValue = m_KeHoachLD.IdPhong
                txtSoDaiHan.Text = m_KeHoachLD.DaiHan
                txtSoNganHan.Text = m_KeHoachLD.NganHan
                txtGhiChu.Text = m_KeHoachLD.GhiChu
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Function updateKeHoachLD(ByVal vIdKHLD As String) As Boolean
        Try
            Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
            m_KeHoachLD.IdKHLD = vIdKHLD
            m_KeHoachLD.IdDonVi_KH = CInt(cboDonVi.SelectedValue)
            m_KeHoachLD.IdPhong = CInt(cboPhong.SelectedValue)
            m_KeHoachLD.Nam = dtpkNam.Value.Year
            If rdNam.Checked Then
                m_KeHoachLD.Thang = 0
            Else
                m_KeHoachLD.Thang = numThang.Value
            End If
            m_KeHoachLD.DaiHan = CInt(txtSoDaiHan.Text.Trim)
            m_KeHoachLD.NganHan = CInt(txtSoNganHan.Text.Trim)
            m_KeHoachLD.GhiChu = txtGhiChu.Text
            If vIdKHLD <> "" Then
                m_KeHoachLD.Update()
            Else
                m_KeHoachLD.Add()
            End If
            If m_KeHoachLD.IdKHLD = "" Then Return False
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub txtSoDaiHan_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoDaiHan.TextChanged
        txtSoDaiHan.Text = Val(txtSoDaiHan.Text.Trim)
    End Sub

    Private Sub txtSoNganHan_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoNganHan.TextChanged
        txtSoNganHan.Text = Val(txtSoNganHan.Text.Trim)
    End Sub

    Private Sub gridKHLD_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridKHLD.CellClick
        Try
            'gridLuongDonVi.Rows(e.RowIndex).Selected = True
            idKeHoachLD = gridKHLD.CurrentRow.Cells("idKHLD").Value.ToString
            fillKeHoachLD(idKeHoachLD)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridKHLD_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridKHLD.CellFormatting
        If gridKHLD.Columns(e.ColumnIndex).Name = "Thang" Then
            If e.Value = 0 Then
                e.Value = "Cả năm"
            End If
        End If
        If gridKHLD.Columns(e.ColumnIndex).Name = "Id_DonVi" Then
            e.Value = getDonvi(CInt(e.Value), Not isAll)
        End If
        If gridKHLD.Columns(e.ColumnIndex).Name = "IdPhong" Then
            e.Value = getPhong(CInt(e.Value))
        End If
    End Sub

    Private Sub bntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew.Click
        blankForm()
    End Sub

    Private Sub bntUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate.Click
        Dim lab_Err As String = ""
        'If (idKeHoachLD <> "") Then   ' Sửa
        '    If (Globals.Roles.IndexOf(";206;") < 0) Then
        '        MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu lương đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '        gridKHLD_CellClick(sender, Nothing)
        '        Return
        '    End If
        'End If
        lab_Err = checkForm()
        If lab_Err <> "" Then
            MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If updateKeHoachLD(idKeHoachLD) Then
            bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
            labAlert.Text = "Ghi dữ liệu thành công!"
        Else
            MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If idKeHoachLD = "" Then
            blankForm()
        Else
            fillKeHoachLD(idKeHoachLD)
        End If
    End Sub

    Private Sub bntDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete.Click
        Try
            If idKeHoachLD <> "" Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.OK Then
                    Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
                    m_KeHoachLD.IdKHLD = idKeHoachLD
                    If m_KeHoachLD.Delete Then
                        bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
                        blankForm()
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

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub cboNam_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
        labNam.Text = dtpkNam.Value.Year.ToString
    End Sub

    Private Sub cboDonVi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDonVi.SelectedIndexChanged
        bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
        cboPhong.DataSource = listPhong(CInt(cboDonVi.SelectedValue), True)
    End Sub

    Private Sub gridKHLD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridKHLD.KeyUp
        gridKHLD_CellClick(sender, Nothing)
    End Sub

    Private Sub rdNam_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdNam.CheckedChanged
        numThang.Visible = False
    End Sub

    Private Sub rdThang_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdThang.CheckedChanged
        numThang.Visible = True
    End Sub

    Private Sub dtpkNam_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpkNam.ValueChanged
        bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
    End Sub
End Class