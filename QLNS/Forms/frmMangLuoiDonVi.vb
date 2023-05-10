Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmMangLuoiDonVi
    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private idMangLuoiDV As String = ""
    Private IdDV As Integer
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

    Private Sub frmMangLuoiDonVi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initMangLuoiDonVi()
    End Sub

    Private Sub initMangLuoiDonVi()
        If DONVI = gMaDonViTW Then isAll = True
        cboDV.DataSource = listDonvi(True, False, False, True)
        bindGridMangLuoiDV(IdDONVI, dtpkNam.Value.Year)
        cboDV.Focus()
    End Sub

    Private Sub bindGridMangLuoiDV(ByVal vIdDonVi As Integer, ByVal vNam As Integer)
        Try
            gridMLDV.AutoGenerateColumns = False
            Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
            gridMLDV.DataSource = m_MangLuoiDV.getAllByDonVi(vIdDonVi, vNam)
            Dim i As Integer = 0
            While i <= gridMLDV.Rows.Count - 1
                If Trim(gridMLDV.Rows(i).Cells(0).Value) = idMangLuoiDV Then
                    gridMLDV.Rows(i).Selected = True
                    Exit While
                End If
                i = i + 1
            End While
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách mạng lưới kế hoạch hoạt động của đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkForm() As String
        Dim strReturn As String = ""
        Try
            If CInt(cboDV.SelectedValue) = 0 Then
                cboDV.Focus()
                strReturn = "Chưa chọn chi nhánh!"
                Exit Try
            End If
            If idMangLuoiDV = "" Then
                If checkMangLuoiDV(CInt(cboDV.SelectedValue), dtpkNam.Value.Year, numQuy.Value) Then
                    strReturn = "Số liệu " & getDonvi(CInt(cboDV.SelectedValue)) & " năm " & dtpkNam.Value.Year & " quý " & numQuy.Value & " đã tồn tại. Hãy nhập lại!"
                    Exit Try
                End If
            End If
            If txtSoDiemGD.Text = "" Then
                txtSoDiemGD.Focus()
                strReturn = "Chưa nhập số điểm giao dịch!"
                Exit Try
            End If
            If (txtDuNo.Text.Trim = "" Or txtDuNo.Text.Trim = "0") Then
                strReturn = "Hãy nhập dự nợ!"
                txtDuNo.Focus()
                Exit Try
            ElseIf (txtDuNo.Text.Trim.Length > 11) Then
                strReturn = "Kiểm tra lại dư nợ!"
                txtDuNo.Focus()
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankForm()
        idMangLuoiDV = ""
        bntDelete.Enabled = False
        labAlert.Text = ""
        dtpkNam.Value = Now
        'numQuy.Value = Now.Month
        txtXaPhuong.Text = ""
        txtSoDiemGD.Text = ""
        txtDuNo.Text = ""
        txtGhiChu.Text = ""
        cboDV.Focus()
    End Sub

    Private Sub fillMangLuoiDV(ByVal vidMangLuoiDV As String)
        If vidMangLuoiDV = "" Then
            blankForm()
        Else
            Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
            m_MangLuoiDV = m_MangLuoiDV.getRecord(vidMangLuoiDV)
            'If (Globals.Roles.IndexOf(";207;") < 0) Then  ' Xoá
            '    bntDelete_CN.Enabled = False
            'Else
            '    bntDelete_CN.Enabled = True
            'End If
            bntDelete.Enabled = True
            labAlert.Text = ""
            Dim vDate As DateTime = New DateTime(CInt(m_MangLuoiDV.Nam.ToString()), 1, 1)
            dtpkNam.Value = vDate
            numQuy.Value = m_MangLuoiDV.Quy
            IdDV = m_MangLuoiDV.IdDonVi
            cboDV.SelectedValue = m_MangLuoiDV.IdPGD
            txtXaPhuong.Text = m_MangLuoiDV.SoXaPhuong
            txtSoDiemGD.Text = m_MangLuoiDV.SoDiemGD
            txtDuNo.Text = m_MangLuoiDV.DuNo
            txtGhiChu.Text = m_MangLuoiDV.GhiChu
        End If
        
    End Sub

    Private Function updateMangLuoiDV(ByVal vidMangLuoiDV As String) As Boolean
        Try
            Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
            m_MangLuoiDV.IdMangLuoi = vidMangLuoiDV
            m_MangLuoiDV.IdDonVi = IdDONVI
            m_MangLuoiDV.IdPGD = CInt(cboDV.SelectedValue)
            m_MangLuoiDV.Nam = dtpkNam.Value.Year
            m_MangLuoiDV.Quy = numQuy.Value
            m_MangLuoiDV.SoXaPhuong = CInt(txtXaPhuong.Text.Trim)
            m_MangLuoiDV.SoDiemGD = CInt(txtSoDiemGD.Text.Trim)
            m_MangLuoiDV.DuNo = N2Number(MoneyValue(txtDuNo.Text.ToString))
            'm_MangLuoiDV.DuNo = MoneyValue(txtDuNo.Text.ToString)
            m_MangLuoiDV.GhiChu = txtGhiChu.Text
            If vidMangLuoiDV <> "" Then
                m_MangLuoiDV.Update()
            Else
                m_MangLuoiDV.Add()
            End If
            If m_MangLuoiDV.IdMangLuoi = "" Then Return False
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub txtDuNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDuNo.TextChanged
        Try
            txtDuNo = formatMoneyinTextbox(txtDuNo)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridMLDV_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridMLDV.CellClick
        Try
            idMangLuoiDV = gridMLDV.CurrentRow.Cells("IdML").Value.ToString
            fillMangLuoiDV(idMangLuoiDV)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridMLDV_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridMLDV.CellFormatting
        Try
            If gridMLDV.Columns(e.ColumnIndex).Name = "IdPGD" Then
                e.Value = getDonvi(CInt(e.Value), Not isAll)
            End If
            If gridMLDV.Columns(e.ColumnIndex).Name = "DuNo" Then
                e.Value = formatMoney(e.Value.ToString)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridMLDV_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridMLDV.KeyUp
        gridMLDV_CellClick(sender, Nothing)
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
        If updateMangLuoiDV(idMangLuoiDV) Then
            bindGridMangLuoiDV(IdDONVI, dtpkNam.Value.Year)
            labAlert.Text = "Ghi dữ liệu thành công!"
        Else
            MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If idMangLuoiDV = "" Then
            blankForm()
        Else
            fillMangLuoiDV(idMangLuoiDV)
        End If
    End Sub

    Private Sub bntDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete.Click
        Try
            If idMangLuoiDV <> "" Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
                    m_MangLuoiDV.IdMangLuoi = idMangLuoiDV
                    If m_MangLuoiDV.Delete Then
                        bindGridMangLuoiDV(IdDONVI, dtpkNam.Value.Year)
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

    Private Sub txtSoDiemGD_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoDiemGD.TextChanged
        txtSoDiemGD.Text = Val(txtSoDiemGD.Text.Trim)
    End Sub

    Private Sub dtpkNam_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpkNam.ValueChanged
        bindGridMangLuoiDV(IdDONVI, dtpkNam.Value.Year)
    End Sub
End Class