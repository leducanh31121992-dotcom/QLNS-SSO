Public Class frmHT_ChotDuLieu
    Dim MyConn As New DBAccess

    Private Sub frmHT_ChotDuLieu_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim _currentdate As DateTime = MyConn.getDateTime("SELECT GETDATE()")
        With cboChonKy
            For i As Integer = 1 To 6
                cboChonKy.Items.Add(String.Format("{1}{0} - tháng {0} năm {1}", My_CStrByLength(Month(_currentdate), 2, "0"c), Year(_currentdate)))
                _currentdate = _currentdate.AddMonths(-1)
            Next
        End With
        _LayKyChotGanNhat()
    End Sub

    Private Sub cboChonKy_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboChonKy.SelectedIndexChanged
        If cboChonKy.SelectedIndex > -1 Then
            btnChotDuLieu.Enabled = Not _ChotDLKyNayChua(Microsoft.VisualBasic.Left(cboChonKy.Text, 6))
            _LoadDGV(Microsoft.VisualBasic.Left(cboChonKy.Text, 6))
        End If
    End Sub

    Private Sub btnChotDuLieu_Click(sender As System.Object, e As System.EventArgs) Handles btnChotDuLieu.Click
        If My_MessageBox("Có chắc chắn chốt dữ liệu kỳ " & Microsoft.VisualBasic.Left(cboChonKy.Text, 6) & "hay không?") = vbYes Then
            Dim StrSQL As String = "INSERT INTO HT_XacNhanDuLieu(IdDonVi, KyXacNhan, ThoiGian) VALUES (" & IdDONVI & ",'" & Microsoft.VisualBasic.Left(cboChonKy.Text, 6) & "', GETDATE())"
            MyConn.executeSQL(StrSQL)
            btnChotDuLieu.Enabled = False
            _LayKyChotGanNhat()
            _LoadDGV(Microsoft.VisualBasic.Left(cboChonKy.Text, 6))
        End If
    End Sub

    Sub _LayKyChotGanNhat()
        Dim StrSQL As String = "SELECT MAX(KyXacNhan) FROM HT_XacNhanDuLieu WHERE IdDonVi = " & IdDONVI
        lblKyGanNhat.Text = MyConn.getString(StrSQL)
        If lblKyGanNhat.Text = "" Then lblKyGanNhat.Text = "Chưa có kỳ chốt nào"
    End Sub

    Function _ChotDLKyNayChua(sKy As String) As Boolean
        Dim StrSQL As String = "SELECT COUNT(*) FROM HT_XacNhanDuLieu WHERE IdDonVi = " & IdDONVI & " AND KyXacNhan = '" & sKy & "'"
        Return (MyConn.getNumber(StrSQL) > 0)
    End Function

    Sub _LoadDGV(sKy As String)
        DGV.Rows.Clear()
        Dim StrSQL As String = "SELECT T1.ma_so, T1.ten_goi, T2.ThoiGian FROM ChiNhanh T1 INNER JOIN HT_XacNhanDuLieu T2 ON T1.Id = T2.IdDonVi " _
                               & "WHERE T2.KyXacNhan = '" & sKy & "' " _
                               & "ORDER BY T2.ThoiGian DESC"
        Dim dt As DataTable = MyConn.getDataTable(StrSQL)

        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                DGV.Rows.Add({dt.Rows(i)(0), dt.Rows(i)(1), Format(dt.Rows(i)(2), "dd/MM/yyyy - h:mm tt")})
            Next
        End If

    End Sub
End Class