Public Class frmNN_DSCanBo

#Region "Local Function"
    Sub _Load_DonVi()
        Dim conn As New DBAccess
        cbo_Binding(cboDonVi, conn.getDataTable("SELECT Id, ten_goi FROM ChiNhanh WHERE id_goc <= 1 And [Status]=1 Order By id"))
        If cboDonVi.Items.Count > 0 Then cboDonVi.SelectedIndex = 0
    End Sub

    Sub _Load_DSCanBo()
        chkAll_BC.Checked = False
        chkAll_TT.Checked = False
        dgv.Rows.Clear()
        If My_CInt(cboDonVi.SelectedValue, -1) = -1 Then Exit Sub
        Dim StrSQL As String
        If rbtChuyenNganh.Checked Then
            If rbtTheoDonVi.Checked Then
                StrSQL = "SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY T2.Id) AS Row, T1.IdCanBo, T1.HoTen, T1.NgaySinh, T5.ten_phong + ' - ' + T2.ten_goi, T3.[DBtmp], T3.[Rpt_Ngoainganh] " _
                            & "FROM HS_CanBo T1 " _
                            & "INNER JOIN HS_CBThoiViec T3 ON T1.IdCanBo = T3.IdCanBo AND T3.[isQD_NHCS] = 1 " _
                            & "INNER JOIN (SELECT Z1.* FROM QDNhanSu Z1 INNER JOIN (SELECT IdCanBo,MAX(NgayHL) NgayHL FROM QDNhanSu GROUP BY IdCanBo) Z2 ON Z1.IdCanBo = Z2.IdCanBo AND Z1.NgayHL = Z2.NgayHL) T4 ON T1.IdCanBo = T4.IdCanBo " _
                            & "INNER JOIN PhongBan T5 ON T4.IdPhong_Moi = T5.id " _
                            & "INNER JOIN ChiNhanh T2 ON T4.IdDonVi_Moi = T2.id " _
                            & "WHERE ((T1.GioiTinh = 0 And (DateDiff(Year, T1.NgaySinh, GETDATE()) < 60 Or (DateDiff(Year, T1.NgaySinh, GETDATE()) = 60 And Month(T1.NgaySinh) >= Month(GETDATE())))) Or ((T1.GioiTinh = 1 And DateDiff(Year, T1.NgaySinh, GETDATE()) < 55) Or (DateDiff(Year, T1.NgaySinh, GETDATE()) = 55 And Month(T1.NgaySinh) >= Month(GETDATE())))) " _
                            & "AND T4.IdDonVi_Moi = " & cboDonVi.SelectedValue
            Else
                StrSQL = "SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY T2.Id) AS Row, T1.IdCanBo, T1.HoTen, T1.NgaySinh, T5.ten_phong + ' - ' + T2.ten_goi, T3.[DBtmp], T3.[Rpt_Ngoainganh] " _
                            & "FROM HS_CanBo T1 " _
                            & "INNER JOIN HS_CBThoiViec T3 ON T1.IdCanBo = T3.IdCanBo AND T3.[isQD_NHCS] = 1 " _
                            & "INNER JOIN (SELECT Z1.* FROM QDNhanSu Z1 INNER JOIN (SELECT IdCanBo,MAX(NgayHL) NgayHL FROM QDNhanSu GROUP BY IdCanBo) Z2 ON Z1.IdCanBo = Z2.IdCanBo AND Z1.NgayHL = Z2.NgayHL) T4 ON T1.IdCanBo = T4.IdCanBo " _
                            & "INNER JOIN PhongBan T5 ON T4.IdPhong_Moi = T5.id " _
                            & "INNER JOIN ChiNhanh T2 ON T4.IdDonVi_Moi = T2.id " _
                            & "WHERE ((T1.GioiTinh = 0 And (DateDiff(Year, T1.NgaySinh, GETDATE()) < 60 Or (DateDiff(Year, T1.NgaySinh, GETDATE()) = 60 And Month(T1.NgaySinh) >= Month(GETDATE())))) Or ((T1.GioiTinh = 1 And DateDiff(Year, T1.NgaySinh, GETDATE()) < 55) Or (DateDiff(Year, T1.NgaySinh, GETDATE()) = 55 And Month(T1.NgaySinh) >= Month(GETDATE()))))"
            End If
        Else
            If rbtTheoDonVi.Checked Then
                StrSQL = "SELECT ROW_NUMBER() OVER (ORDER BY T2.Id) AS Row, T1.Id, T1.HoTen, T1.NgaySinh, T2.ten_goi, T1.DBtmp, T1.Rpt_Ngoainganh FROM [HSCB_TS] T1 " _
                            & "INNER JOIN ChiNhanh T2 ON T1.IdChiNhanh = T2.id " _
                            & "WHERE T2.Id = " & cboDonVi.SelectedValue
            Else
                StrSQL = "SELECT ROW_NUMBER() OVER (ORDER BY T2.Id) AS Row, T1.Id, T1.HoTen, T1.NgaySinh, T2.ten_goi, T1.DBtmp, T1.Rpt_Ngoainganh FROM [HSCB_TS] T1 " _
                            & "INNER JOIN ChiNhanh T2 ON T1.IdChiNhanh = T2.id"
            End If
        End If

        Me.Cursor = Cursors.WaitCursor

        Dim conn As New DBAccess
        Dim dt As DataTable = conn.getDataTable(StrSQL)
        DGV_LoadData(dgv, dt)
        Me.Cursor = Cursors.Default
    End Sub

#End Region

    Private Sub frmNN_DSCanBo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _Load_DonVi()
    End Sub

    Private Sub DSCanBo(sender As Object, e As EventArgs) Handles cboDonVi.SelectedIndexChanged, rbtChuyenNganh.CheckedChanged
        _Load_DSCanBo()
    End Sub

    Private Sub btn_Luu_Click(sender As Object, e As EventArgs) Handles btn_Luu.Click
        Dim conn As New DBAccess
        Dim StrSQL As String
        If dgv.Rows.Count > 0 Then
            For Each r As DataGridViewRow In dgv.Rows
                If rbtChuyenNganh.Checked Then
                    StrSQL = "UPDATE HS_CBThoiViec SET DBtmp = " & IIf(r.Cells("TamThoi").Value, 1, 0) & " , Rpt_Ngoainganh = " & IIf(r.Cells("BaoCao").Value, 1, 0) & " " _
                                & "WHERE IdCanBo = '" & r.Cells("ID").Value & "'"
                Else
                    StrSQL = "UPDATE HSCB_TS SET DBtmp = " & IIf(r.Cells("TamThoi").Value, 1, 0) & " , Rpt_Ngoainganh = " & IIf(r.Cells("BaoCao").Value, 1, 0) & " " _
                                & "WHERE Id = '" & r.Cells("ID").Value & "'"
                End If
                conn.executeSQL(StrSQL)
            Next
        End If
        My_MessageBox("Lưu dữ liệu thành công.")
    End Sub

    Private Sub btn_back_Click(sender As Object, e As EventArgs) Handles btn_back.Click
        Me.Close()
    End Sub

    Private Sub chkAllTT_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll_TT.CheckedChanged
        With dgv
            If .Rows.Count > 0 Then
                For Each r As DataGridViewRow In dgv.Rows
                    r.Cells("TamThoi").Value = chkAll_TT.Checked
                Next
            End If
        End With
    End Sub

    Private Sub chkAll_BC_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll_BC.CheckedChanged
        With dgv
            If .Rows.Count > 0 Then
                For Each r As DataGridViewRow In dgv.Rows
                    r.Cells("BaoCao").Value = chkAll_BC.Checked
                Next
            End If
        End With
    End Sub

    Private Sub rbtToanBo_CheckedChanged(sender As Object, e As EventArgs) Handles rbtToanBo.CheckedChanged, rbtTheoDonVi.CheckedChanged
        cboDonVi.Enabled = rbtTheoDonVi.Checked
        _Load_DSCanBo()
    End Sub

End Class