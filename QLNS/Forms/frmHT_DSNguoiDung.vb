Public Class frmHT_DSNguoiDung
    Private strSQL As String = ""
    Private _Globals As New Globals
    Private _Conn As New DBAccess
    Private arr_Groups As ArrayList = New ArrayList()

#Region "Form events"
    Private Sub frmHT_DSNguoiDung_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _InitControls()
        btnQuyenQuanLyCN.Visible = False 'Chỉ áp dụng cho nhóm Sử dụng ở cấp TW
    End Sub

    Private Sub cboNhomQuyen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNhomQuyen.SelectedIndexChanged
        _LoadUsers(IIf(arr_Groups.Count > 0, arr_Groups(cboNhomQuyen.SelectedIndex), ""))
        btnQuyenQuanLyCN.Visible = (arr_Groups(cboNhomQuyen.SelectedIndex) = "S02")
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        If dgv.SelectedRows.Count = 0 Then
            My_MessageBox("Chưa có dòng nào được chọn để hủy quyền đăng nhập!")
        ElseIf My_MessageBox("Có chắc chắn xóa thông tin đăng nhập này không?") = Windows.Forms.DialogResult.Yes Then
            For Each r As DataGridViewRow In dgv.SelectedRows
                strSQL = "UPDATE HS_CanBo SET Login_Username = '', Login_Password = '', Login_POS = '', ID_Nhom = '' WHERE IdCanBo = '" & r.Cells("IdCanBo").Value & "'"
                _Conn.executeSQL(strSQL)
                dgv.Rows.Remove(r)
            Next
        End If
    End Sub

    Private Sub btn_Edit_Click(sender As Object, e As EventArgs) Handles btn_Edit.Click
        If dgv.SelectedRows.Count <> 1 Then
            My_MessageBox("Cần chọn 1 dòng để sửa!")
        Else
            Dim r As DataGridViewRow = dgv.SelectedRows(0)
            Dim IdCanBo As String = r.Cells("IdCanbo").Value
            Dim frm As New frmHT_QuyenSuDung
            frm.HoTen = r.Cells("HoTen").Value
            frm.Username = r.Cells("Username").Value
            frm.IdCanBo = r.Cells("IdCanbo").Value
            frm.MaPOS = r.Cells("POS").Value
            frm.ShowDialog()

            'Load lại user
            _LoadUsers(IIf(arr_Groups.Count > 0, arr_Groups(cboNhomQuyen.SelectedIndex), ""))
            For Each row As DataGridViewRow In dgv.Rows
                If row.Cells("IdCanBo").Value = IdCanBo Then
                    row.Selected = True
                    dgv.CurrentCell = row.Cells(0)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub btnQuyenQuanLyCN_Click(sender As Object, e As EventArgs) Handles btnQuyenQuanLyCN.Click
        If dgv.SelectedRows.Count <> 1 Then
            My_MessageBox("Cần chọn 1 dòng để thực hiện phân quyền quản lý Chi nhánh!")
        Else
            Dim r As DataGridViewRow = dgv.SelectedRows(0)
            Dim IdCanBo As String = r.Cells("IdCanbo").Value
            Dim frm As New frmHT_QuyenQuanLyCN(r.Cells("IdCanbo").Value, r.Cells("HoTen").Value)
            frm.ShowDialog()
        End If
    End Sub

    Private Sub btn_back_Click(sender As Object, e As EventArgs) Handles btn_back.Click
        Me.Close()
    End Sub

#End Region

#Region "Local functions"
    Sub _InitControls()
        With dgv
            .Columns.Add("STT", "STT")
            .Columns.Add("HoTen", "Họ tên")
            .Columns.Add("IdCanBo", "Mã CB")
            .Columns.Add("Username", "Tên đăng nhập")
            .Columns.Add("POS", "Mã POS")
            .Columns.Add("DonVi", "Đơn vị")
            .ReadOnly = True
            .AutoResizeColumns()
        End With

        arr_Groups.Clear()
        cboNhomQuyen.Items.Clear()
        If Microsoft.VisualBasic.Left(Globals.Group, 1).ToUpper <> "U" Then
            strSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Order by ma_nhom Asc")
        Else
            strSQL = String.Format("Select ma_nhom,ten_nhom From HT_NhomTV Where LEFT(ma_nhom,1) = 'U' Order by ma_nhom Asc")
        End If
        arr_Groups = _Globals.Bind_ComBoBox(cboNhomQuyen, strSQL, "---Nhóm thành viên---")
    End Sub

    Sub _LoadUsers(Id_Nhom As String)
        dgv.Rows.Clear()
        If DONVI = gMaDonViTW Then
            strSQL = "SELECT T1.IdCanBo, T1.HoTen, T3.Ten_Goi, T3.Ma_So, T1.Login_Username, T1.Login_POS " _
            & "FROM HS_CanBo T1 " _
            & "INNER JOIN (SELECT Z1.* FROM QDNhanSu Z1 INNER JOIN (SELECT IdCanBo,MAX(NgayHL) NgayHL FROM QDNhanSu GROUP BY IdCanBo) Z2 ON Z1.IdCanBo = Z2.IdCanBo AND Z1.NgayHL = Z2.NgayHL) T2 ON T1.IdCanBo = T2.IdCanBo " _
            & "INNER JOIN ChiNhanh T3 ON T2.IdDonVi_Moi = T3.Id " _
            & "WHERE T1.ID_Nhom = '" & Id_Nhom & "' " _
            & "ORDER BY T3.ma_so"
        Else
            strSQL = "SELECT T1.IdCanBo, T1.HoTen, T3.Ten_Goi, T3.Ma_So, T1.Login_Username, T1.Login_POS " _
            & "FROM HS_CanBo T1 " _
            & "INNER JOIN (SELECT Z1.* FROM QDNhanSu Z1 INNER JOIN (SELECT IdCanBo,MAX(NgayHL) NgayHL FROM QDNhanSu GROUP BY IdCanBo) Z2 ON Z1.IdCanBo = Z2.IdCanBo AND Z1.NgayHL = Z2.NgayHL) T2 ON T1.IdCanBo = T2.IdCanBo " _
            & "INNER JOIN ChiNhanh T3 ON T2.IdDonVi_Moi = T3.Id " _
            & "WHERE T1.ID_Nhom = '" & Id_Nhom & "' AND T1.Login_POS = '" & DONVI & "' " _
            & "ORDER BY T3.ma_so"
        End If
        Using db As DataTable = _Conn.getDataTable(strSQL)
            If Not (db Is Nothing) AndAlso (db.Rows.Count > 0) Then
                For i As Integer = 0 To db.Rows.Count - 1
                    dgv.Rows.Add()
                    dgv.Rows(i).Cells("STT").Value = i + 1
                    dgv.Rows(i).Cells("HoTen").Value = db.Rows(i)("HoTen").ToString().Trim()
                    dgv.Rows(i).Cells("Username").Value = db.Rows(i)("Login_Username").ToString().Trim()
                    dgv.Rows(i).Cells("IdCanBo").Value = db.Rows(i)("IdCanBo").ToString().Trim()
                    dgv.Rows(i).Cells("POS").Value = db.Rows(i)("Login_POS").ToString().Trim()
                    dgv.Rows(i).Cells("DonVi").Value = db.Rows(i)("Ten_Goi").ToString().Trim()
                Next
            End If
        End Using
        dgv.AutoResizeColumns()
    End Sub
#End Region
End Class