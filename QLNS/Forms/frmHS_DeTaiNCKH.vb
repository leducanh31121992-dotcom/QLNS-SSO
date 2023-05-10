Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.DataViewManager
Imports System.Globalization

Public Class frmHS_DeTaiNCKH
    Inherits System.Windows.Forms.Form
    Private dbconn As DBAccess
    Private idDeTai As String = ""
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
    Private _Value As String
    Private _CapDeTai As Integer

    Public Property Value() As String
        Get
            Return _Value
        End Get
        Set(ByVal value As String)
            _Value = value
        End Set
    End Property

    Public Property CapDeTai() As Integer
        Get
            Return _CapDeTai
        End Get
        Set(ByVal value As Integer)
            _CapDeTai = value
        End Set
    End Property

    Private Sub frmHS_DeTaiNCKH_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (Globals.Roles.IndexOf(";147;") < 0) Then
            gridDeTai.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";148;") < 0) Then
            bntNew.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";150;") < 0) Then
            bntDelete.Enabled = False
        End If
        idDeTai = Value
        bindCboCapDeTai()
        cboCapDeTai_NC.SelectedValue = CapDeTai
        bindGridDeTai(cboCapDeTai_NC.SelectedValue)
        fillDeTaiNCKH(Value)
        'gridDeTai.SelectedRows(gridDeTai.Rows.Cells("IdDeTaiNCKH").Value.ToString).Selected = True
        gridDeTai.Focus()
    End Sub

    Private Sub frmHS_DeTaiNCKH_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub frmHS_DeTaiNCKH_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub bindCboCapDeTai()
        cboCapDeTai_NC.DataSource = listDanhmuc(7, True)
        'cboCapDeTai_NC.SelectedValue = vIdCapDeTai
    End Sub

    Private Sub bindGridDeTai(ByVal vIdCapDeTai As Integer)
        Try
            gridDeTai.AutoGenerateColumns = False
            Dim m_DeTai As DeTaiNCKH = New DeTaiNCKH
            If vIdCapDeTai = 0 Then
                gridDeTai.DataSource = m_DeTai.getAll
            Else
                gridDeTai.DataSource = m_DeTai.getAllByCapDeTai(vIdCapDeTai)
            End If
            'fillDeTaiNCKH(idDeTai)
            Dim i As Integer = 0
            While i <= gridDeTai.Rows.Count - 1
                If Trim(gridDeTai.Rows(i).Cells(0).Value) = idDeTai Then
                    gridDeTai.Rows(i).Selected = True
                    Exit While
                End If
                i = i + 1
            End While
        Catch ex As Exception
            MessageBox.Show("Không Load được Danh sách Đề tài nghiên cứu khoa học!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function checkDeTaiNCKH() As String
        Dim strReturn As String = ""
        Try
            If txtTenDeTai_NC.Text = "" Then
                txtTenDeTai_NC.Focus()
                strReturn = "Chưa nhập tên đề tài!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub blankDeTai()
        idDeTai = ""
        bntDelete.Enabled = False
        txtTenDeTai_NC.Text = ""
        txtNoiDung_NC.Text = ""
        txtDonViQL_NC.Text = ""
        txtGhichu_NC.Text = ""
        txtTenDeTai_NC.Focus()
    End Sub

    'Private Sub fillDeTaiNCKH(ByVal vIdDeTai As String, Optional ByVal vIdCapDeTai As Integer = 0)
    Private Sub fillDeTaiNCKH(ByVal vIdDeTai As String)
        Dim m_DeTaiNCKH As DeTaiNCKH = New DeTaiNCKH
        m_DeTaiNCKH = m_DeTaiNCKH.getRecord(vIdDeTai)
        idDeTai = vIdDeTai
        If (Globals.Roles.IndexOf(";150;") < 0) Then  ' Xoá
            bntDelete.Enabled = False
        Else
            bntDelete.Enabled = True
        End If
        If m_DeTaiNCKH.IdDeTai <> "" Then
            cboCapDeTai_NC.SelectedValue = m_DeTaiNCKH.IdCapDeTai
            txtTenDeTai_NC.Text = m_DeTaiNCKH.TenDeTai
            txtNoiDung_NC.Text = m_DeTaiNCKH.NoiDung
            dpkTuNgay_NC.Value = m_DeTaiNCKH.TuNgay
            dpkDenNgay_NC.Value = m_DeTaiNCKH.DenNgay
            If m_DeTaiNCKH.NgayNghiemThu = DateTime.MinValue Then
                dpkNgayNT_NC.Checked = False
            Else
                dpkNgayNT_NC.Checked = True
                dpkNgayNT_NC.Value = m_DeTaiNCKH.NgayNghiemThu
            End If
            txtDonViQL_NC.Text = m_DeTaiNCKH.DonVi_QL
            txtGhichu_NC.Text = m_DeTaiNCKH.GhiChu
            Value = vIdDeTai
            CapDeTai = m_DeTaiNCKH.IdCapDeTai
        Else
            blankDeTai()
        End If
    End Sub

    Private Function updateDeTaiNCKH(ByRef vIdDeTai As String) As Boolean
        Try
            Dim m_DeTaiNCKH As DeTaiNCKH = New DeTaiNCKH
            m_DeTaiNCKH.IdDeTai = vIdDeTai
            m_DeTaiNCKH.IdCapDeTai = CInt(cboCapDeTai_NC.SelectedValue)
            m_DeTaiNCKH.TenDeTai = txtTenDeTai_NC.Text
            m_DeTaiNCKH.NoiDung = txtNoiDung_NC.Text
            m_DeTaiNCKH.TuNgay = DateTimeUtil.getDate(dpkTuNgay_NC.Text)
            m_DeTaiNCKH.DenNgay = DateTimeUtil.getDate(dpkDenNgay_NC.Text)
            If dpkNgayNT_NC.Checked Then
                m_DeTaiNCKH.NgayNghiemThu = DateTimeUtil.getDate(dpkNgayNT_NC.Text)
            Else
                m_DeTaiNCKH.NgayNghiemThu = DateTime.MinValue
            End If
            m_DeTaiNCKH.DonVi_QL = txtDonViQL_NC.Text
            m_DeTaiNCKH.GhiChu = txtGhichu_NC.Text
            If vIdDeTai <> "" Then
                m_DeTaiNCKH.Update()
            Else
                vIdDeTai = m_DeTaiNCKH.Add()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub cboCapDeTai_NC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCapDeTai_NC.SelectedIndexChanged
        'fillDeTaiNCKH(0, cboCapDeTai_NC.SelectedValue.ToString)
        bindGridDeTai(CInt(cboCapDeTai_NC.SelectedValue))
        blankDeTai()
    End Sub

    Private Sub gridDeTai_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridDeTai.CellClick
        Try
            'gridDeTai.Rows(e.RowIndex).Selected = True
            idDeTai = gridDeTai.CurrentRow.Cells("IdDeTaiNCKH").Value.ToString
            'Value = idDeTai
            fillDeTaiNCKH(idDeTai)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gridDeTai_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles gridDeTai.CellFormatting
        Try
            Dim row As DataGridViewRow
            If e.ColumnIndex = gridDeTai.Columns("IdCapDeTai").Index Then
                e.FormattingApplied = True
                row = gridDeTai.Rows(e.RowIndex)
                e.Value = getDanhmuc_Name(7, CInt(row.Cells("IdCapDeTai").Value))
            End If
            'If e.ColumnIndex = gridDeTai.Columns("TuNgay").Index Then
            '    row = gridDeTai.Rows(e.RowIndex)
            '    e.Value = DateTimeUtil.getShortDate(CDate(row.Cells("TuNgay").Value))
            'End If
            'If e.ColumnIndex = gridDeTai.Columns("DenNgay").Index Then
            '    row = gridDeTai.Rows(e.RowIndex)
            '    e.Value = DateTimeUtil.getShortDate(CDate(row.Cells("DenNgay").Value))
            'End If
            'If e.ColumnIndex = gridDeTai.Columns("NgayNghiemThu").Index Then
            '    row = gridDeTai.Rows(e.RowIndex)
            '    e.Value = DateTimeUtil.getShortDate(CDate(row.Cells("NgayNghiemThu").Value))
            'End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate.Click
        Try
            Dim lab_ErrNC As String = ""
            If (idDeTai <> "") Then   ' Sửa
                If (Globals.Roles.IndexOf(";101;") < 0) Then
                    MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu đề tài nghiên cứu khoa học!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    gridDeTai_CellClick(sender, Nothing)
                    Return
                End If
            End If
            lab_ErrNC = checkDeTaiNCKH()
            If lab_ErrNC <> "" Then
                MessageBox.Show(lab_ErrNC, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If updateDeTaiNCKH(idDeTai) Then
                bindGridDeTai(CInt(cboCapDeTai_NC.SelectedValue))
                Value = idDeTai
                CapDeTai = CInt(cboCapDeTai_NC.SelectedValue)
                'blankNCKH()
                MessageBox.Show("Ghi dữ liệu thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete.Click
        Try
            Dim arrDel As ArrayList = New ArrayList()
            If (gridDeTai.Rows.Count > 0) Then
                For i As Int32 = 0 To gridDeTai.Rows.Count - 1
                    If (CType(gridDeTai.Rows(i).Cells("cln_cbNCKH").Value, Boolean) = True) Then
                        arrDel.Add(gridDeTai.Rows(i).Cells("IdDeTaiNCKH").Value.ToString())
                    End If
                Next
            End If
            If (arrDel.Count = 0) And (idDeTai <> "") Then arrDel.Add(idDeTai.ToString)
            If (arrDel.Count > 0) Then
                If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    Try
                        Dim i As Integer
                        For i = 0 To arrDel.Count - 1
                            Dim m_detaiNCKH As DeTaiNCKH = New DeTaiNCKH
                            m_detaiNCKH.IdDeTai = arrDel(i).ToString()
                            If Value = m_detaiNCKH.IdDeTai Then Value = "0"
                            m_detaiNCKH.Delete()
                        Next
                        blankDeTai()
                        bindGridDeTai(CInt(cboCapDeTai_NC.SelectedValue))
                        MessageBox.Show("Xoá dữ liệu thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
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

    Private Sub bntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew.Click
        blankDeTai()
    End Sub

    Private Sub bntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel.Click
        If idDeTai = "" Then
            blankDeTai()
        Else
            fillDeTaiNCKH(idDeTai)
        End If
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub gridDeTai_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gridDeTai.KeyUp
        gridDeTai_CellClick(sender, Nothing)
    End Sub

End Class