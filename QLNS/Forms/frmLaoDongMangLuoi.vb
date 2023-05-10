Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmLaoDongMangLuoi
    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private init_LD As Boolean = False
    Private init_ML As Boolean = False
    Private idKeHoachLD As String = ""
    Private idMangLuoiDV As String = ""
    Private IdDV As Integer
    Private isAll As Boolean = False

    Private Sub frmLaoDongMangLuoi_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmLaoDongMangLuoi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If DONVI = gMaDonViTW Then isAll = True
        initKeHoachLaoDong()
    End Sub

#Region "--->Functions<---"

    Private Sub initKeHoachLaoDong()
        cboDonVi.DataSource = listDonvi(False, False, True, True, False, "-------Các đơn vị trực thuộc-------")
        If CInt(cboDonVi.SelectedValue) = IdDONVI Then
            cboPhong.Enabled = True
            cboPhong.DataSource = listPhong(CInt(cboDonVi.SelectedValue))
        Else
            cboPhong.Enabled = False
        End If
        bindGridKeHoachLaoDongDV(cboDonVi.SelectedValue, dtpkNam_LD.Value.Year)
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

    Private Sub blankTabKHLD()
        idKeHoachLD = ""
        btnDelete_LD.Enabled = False
        labAlert.Text = ""
        dtpkNam_LD.Value = Now
        numThang.Value = Now.Month
        txtSoDaiHan.Text = ""
        txtSoNganHan.Text = ""
        txtGhiChu.Text = ""
        cboDonVi.Focus()
    End Sub

    Private Sub fillKeHoachLD(ByVal vIdKHLD As String)
        Try
            If vIdKHLD = "" Then
                blankTabKHLD()
            Else
                Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
                m_KeHoachLD = m_KeHoachLD.getRecord(vIdKHLD)
                'If (Globals.Roles.IndexOf(";207;") < 0) Then  ' Xoá
                '    bntDelete_CN.Enabled = False
                'Else
                '    bntDelete_CN.Enabled = True
                'End If
                btnDelete_LD.Enabled = True
                labAlert.Text = ""
                Dim vDate As DateTime = New DateTime(CInt(m_KeHoachLD.Nam.ToString()), 1, 1)
                dtpkNam_LD.Value = vDate
                labNam.Text = dtpkNam_LD.Value.Year
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
            Dim i As Integer
            m_KeHoachLD.IdKHLD = vIdKHLD
            m_KeHoachLD.IdDonVi_KH = CInt(cboDonVi.SelectedValue)
            m_KeHoachLD.IdPhong = CInt(cboPhong.SelectedValue)
            m_KeHoachLD.Nam = dtpkNam_LD.Value.Year
            m_KeHoachLD.DaiHan = CInt(txtSoDaiHan.Text.Trim)
            m_KeHoachLD.NganHan = CInt(txtSoNganHan.Text.Trim)
            m_KeHoachLD.GhiChu = txtGhiChu.Text
            If rdNam.Checked Then
                m_KeHoachLD.Thang = 0
                'For i = 1 To 12
                '    m_KeHoachLD.Thang = i
                '    If vIdKHLD <> "" Then
                '        m_KeHoachLD.Update()
                '    Else
                '        m_KeHoachLD.Add()
                '    End If
                'Next
            Else
                m_KeHoachLD.Thang = numThang.Value
                'For i = numThang.Value To 12
                '    m_KeHoachLD.Thang = i
                '    If vIdKHLD <> "" Then
                '        m_KeHoachLD.Update()
                '    Else
                '        m_KeHoachLD.Add()
                '    End If
                'Next
            End If
            vIdKHLD = getKHLD_ID(m_KeHoachLD.IdDonVi_KH, m_KeHoachLD.IdPhong, m_KeHoachLD.Thang, m_KeHoachLD.Nam)
            m_KeHoachLD.IdKHLD = vIdKHLD
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

    Private Function getKHLD_ID(ByVal vIdDonvi As Integer, ByVal vIdPhong As Integer, ByVal vThang As Integer, ByVal vNam As Integer) As String
        Dim dbconn As DBAccess = New DBAccess
        Try
            Return dbconn.getString("SELECT IDKHLD FROM KEHOACHLD WHERE IdDonvi=" & vIdDonvi & " AND IDPHONG=" & vIdPhong & " AND THANG=" & vThang & " AND NAM=" & vNam)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function checkTabKHLD() As String
        Dim strReturn As String = ""
        Try
            If CInt(cboDonVi.SelectedValue) = 0 Then
                cboDonVi.Focus()
                strReturn = "Chưa chọn chi nhánh!"
                Exit Try
            End If
            'If idKeHoachLD = "" Then
            '    If checkSoKeHoachLaoDong(CInt(cboDonVi.SelectedValue), dtpkNam_LD.Value.Year, numThang.Value) Then
            '        strReturn = "Số kế hoạch lao động của đơn vị năm " & dtpkNam_LD.Value.Year & " đã tồn tại. Hãy nhập lại!"
            '        Exit Try
            '    End If
            'End If
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

    Private Sub initMangLuoiDonVi()
        If DONVI = gMaDonViTW Then isAll = True
        cboDV.DataSource = listDonvi(True, False, False, True)
        bindGridMangLuoiDV(IdDONVI, dtpkNam_LD.Value.Year)
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

    Private Function checkTabMangLuoi() As String
        Dim strReturn As String = ""
        Try
            If CInt(cboDV.SelectedValue) = 0 Then
                cboDV.Focus()
                strReturn = "Chưa chọn chi nhánh!"
                Exit Try
            End If
            If idMangLuoiDV = "" Then
                If checkMangLuoiDV(CInt(cboDV.SelectedValue), dtpkNam_LD.Value.Year, numQuy.Value) Then
                    strReturn = "Số liệu " & getDonvi(CInt(cboDV.SelectedValue)) & " năm " & dtpkNam_LD.Value.Year & " quý " & numQuy.Value & " đã tồn tại. Hãy nhập lại!"
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

    Private Sub blankTabMangLuoi()
        idMangLuoiDV = ""
        btnDelete_LD.Enabled = False
        labAlert.Text = ""
        dtpkNam_LD.Value = Now
        'numQuy.Value = Now.Month
        txtXaPhuong.Text = ""
        txtSoDiemGD.Text = ""
        txtDuNo.Text = ""
        txtGhiChu.Text = ""
        cboDV.Focus()
    End Sub

    Private Sub fillMangLuoiDV(ByVal vidMangLuoiDV As String)
        If vidMangLuoiDV = "" Then
            blankTabMangLuoi()
        Else
            Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
            m_MangLuoiDV = m_MangLuoiDV.getRecord(vidMangLuoiDV)
            'If (Globals.Roles.IndexOf(";207;") < 0) Then  ' Xoá
            '    bntDelete_CN.Enabled = False
            'Else
            '    bntDelete_CN.Enabled = True
            'End If
            btnDelete_LD.Enabled = True
            labAlert.Text = ""
            Dim vDate As DateTime = New DateTime(CInt(m_MangLuoiDV.Nam.ToString()), 1, 1)
            dtpkNam_LD.Value = vDate
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
            m_MangLuoiDV.Nam = dtpkNam_LD.Value.Year
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

    Private Sub callTab()
        Try
            Select Case tabMain.SelectedTab.Name
                Case "tabMangLuoiDV"
                    initMangLuoiDonVi()
                Case Else
                    initKeHoachLaoDong()
            End Select
        Catch ex As Exception

        End Try
    End Sub

#End Region


    Private Sub tabMain_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabMain.SelectedIndexChanged
        callTab()
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
            'If cboDonVi.SelectedValue = 1 Then
            '    e.Value = getPhong(CInt(e.Value))
            'Else
            '    e.Value = getDonvi(CInt(e.Value), Not isAll)
            'End If
            e.Value = getDonvi(CInt(e.Value), Not isAll)
        End If
        If gridKHLD.Columns(e.ColumnIndex).Name = "Id_Phong" Then
            e.Value = getPhong(CInt(e.Value))
        End If
    End Sub

    Private Sub txtSoDaiHan_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoDaiHan.TextChanged
        txtSoDaiHan.Text = Val(txtSoDaiHan.Text.Trim)
    End Sub

    Private Sub txtSoNganHan_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoNganHan.TextChanged
        txtSoNganHan.Text = Val(txtSoNganHan.Text.Trim)
    End Sub

    Private Sub bntNew_LD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew_LD.Click
        blankTabKHLD()
    End Sub

    Private Sub btnUpdate_LD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate_LD.Click
        Dim lab_Err As String = ""
        'If (idKeHoachLD <> "") Then   ' Sửa
        '    If (Globals.Roles.IndexOf(";206;") < 0) Then
        '        MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu lương đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '        gridKHLD_CellClick(sender, Nothing)
        '        Return
        '    End If
        'End If
        lab_Err = checkTabKHLD()
        If lab_Err <> "" Then
            MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If updateKeHoachLD(idKeHoachLD) Then
            bindGridKeHoachLaoDongDV(cboDonVi.SelectedValue, dtpkNam_LD.Value.Year)
            labAlert.Text = "Ghi dữ liệu thành công!"
        Else
            MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btnCancel_LD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel_LD.Click
        If idKeHoachLD = "" Then
            blankTabKHLD()
        Else
            fillKeHoachLD(idKeHoachLD)
        End If
    End Sub

    Private Sub bntDelete_LD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete_LD.Click
        Try
            If idKeHoachLD <> "" Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.OK Then
                    Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
                    m_KeHoachLD.IdKHLD = idKeHoachLD
                    If m_KeHoachLD.Delete Then
                        bindGridKeHoachLaoDongDV(cboDonVi.SelectedValue, dtpkNam_LD.Value.Year)
                        blankTabKHLD()
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

    Private Sub bntClose_LD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose_LD.Click
        Close()
    End Sub

    Private Sub cboDonVi_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDonVi.SelectedIndexChanged
        If CInt(cboDonVi.SelectedValue) = IdDONVI Then
            cboPhong.Enabled = True
        Else
            cboPhong.Enabled = False
        End If
        bindGridKeHoachLaoDongDV(cboDonVi.SelectedValue, dtpkNam_LD.Value.Year)
        cboPhong.DataSource = listPhong(CInt(cboDonVi.SelectedValue), True)
    End Sub

    Private Sub gridKHLD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridKHLD.KeyUp
        gridKHLD_CellClick(sender, Nothing)
    End Sub

    Private Sub rdNam_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdNam.CheckedChanged
        numThang.Visible = False
    End Sub

    Private Sub rdThang_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdThang.CheckedChanged
        numThang.Visible = True
    End Sub

    'Private Sub dtpkNam_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpkNam.ValueChanged
    '    bindGridKeHoachLaoDongDV(IdDONVI, dtpkNam.Value.Year)
    'End Sub

    Private Sub txtDuNo_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDuNo.VisibleChanged
        Try
            txtDuNo = formatMoneyinTextbox(txtDuNo)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridMLDV_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridMLDV.CellClick
        Try
            idMangLuoiDV = gridMLDV.CurrentRow.Cells("IdML").Value.ToString   'IdMangLuoi
            'idMangLuoiDV = gridMLDV.CurrentRow.Cells("IdMangLuoi").Value.ToString
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

    Private Sub bntNew_ML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew_ML.Click
        blankTabMangLuoi()
    End Sub

    Private Sub bntUpdate_ML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate_ML.Click
        Dim lab_Err As String = ""
        'If (idKeHoachLD <> "") Then   ' Sửa
        '    If (Globals.Roles.IndexOf(";206;") < 0) Then
        '        MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu lương đơn vị!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '        gridKHLD_CellClick(sender, Nothing)
        '        Return
        '    End If
        'End If
        lab_Err = checkTabMangLuoi()
        If lab_Err <> "" Then
            MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If updateMangLuoiDV(idMangLuoiDV) Then
            bindGridMangLuoiDV(IdDONVI, dtpkNam_LD.Value.Year)
            labAlert.Text = "Ghi dữ liệu thành công!"
        Else
            MessageBox.Show("Ghi dữ liệu không thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btnCancel_ML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel_ML.Click
        If idMangLuoiDV = "" Then
            blankTabMangLuoi()
        Else
            fillMangLuoiDV(idMangLuoiDV)
        End If
    End Sub

    Private Sub btnDelete_ML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete_ML.Click
        Try
            If idMangLuoiDV <> "" Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
                    m_MangLuoiDV.IdMangLuoi = idMangLuoiDV
                    If m_MangLuoiDV.Delete Then
                        bindGridMangLuoiDV(IdDONVI, dtpkNam_LD.Value.Year)
                        blankTabMangLuoi()
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

    Private Sub btnClose_ML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose_ML.Click
        Close()
    End Sub

    Private Sub txtSoDiemGD_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoDiemGD.TextChanged
        txtSoDiemGD.Text = Val(txtSoDiemGD.Text.Trim)
    End Sub

    Private Sub dtpkNam_ML_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpkNam_ML.ValueChanged
        bindGridMangLuoiDV(IdDONVI, dtpkNam_ML.Value.Year)
    End Sub

    Private Sub cboDV_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDV.SelectedIndexChanged
        bindGridKeHoachLaoDongDV(cboDV.SelectedValue, dtpkNam_ML.Value.Year)
    End Sub
End Class