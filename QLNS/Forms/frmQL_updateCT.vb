Public Class frmQL_updateCT

    Private Sub frmQL_updateCT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        getlastestUpdate()
        initGrid()
        bindGrid()

    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub getlastestUpdate()
        Try
            Dim db As DBAccess = New DBAccess
            Dim strSQL As String = "SELECT TOP 1 * FROM updateCT WHERE iddonvi=" & IdDONVI & " Order by date_update desc "
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = db.SelectDBRows(strSQL)
            For j = 0 To dt.Rows.Count - 1
                labVer.Text = IIf(dt.Rows(j)("version").ToString() <> "", dt.Rows(j)("version").ToString(), "")
                labDateVer.Text = DateTimeUtil.getShortDate(dt.Rows(0)("date_version").ToString())
                labDateUp.Text = DateTimeUtil.getLongTimeDateVN(dt.Rows(0)("date_update").ToString())
            Next
        Catch ex As Exception
            MessageBox.Show("Lỗi: không load được thông tin lần update cuối cùng:", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub initGrid()
        gridDS.Columns().Clear()
        gridDS.Columns.Add("STT", "STT")
        gridDS.Columns.Add("version", "Version")
        gridDS.Columns.Add("date_version", "Date Version")
        gridDS.Columns.Add("date_update", "Ngày update")

        'Căn chỉnh tiêu đề
        gridDS.Columns("STT").Width = 40
        gridDS.Columns("version").Width = 80
        gridDS.Columns("date_version").Width = 120
        gridDS.Columns("date_update").Width = 160 'System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        'gridDS.Columns("ChucDanh").Width = 150
        'gridDS.Columns("MaCB_Cu").Width = 90
        'gridDS.Columns("MaCB_Moi").Width = 90
        ''gridDS.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        ''gridDS.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        'gridDS.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        'gridDS.Columns(6).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub bindGrid()
        Try
            Dim db As DBAccess = New DBAccess
            Dim strSQL As String = "SELECT id, version, date_version, date_update FROM UpdateCT WHERE iddonvi=" & IdDONVI
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = db.SelectDBRows(strSQL)
            gridDS.DataSource = Nothing
            gridDS.Refresh()
            For j = 0 To dt.Rows.Count - 1
                gridDS.Rows.Add()
                gridDS.Rows(j).Cells("STT").Value = j + 1
                gridDS.Rows(j).Cells("version").Value = IIf(dt.Rows(j)("version").ToString() <> "", dt.Rows(j)("version").ToString(), "")
                gridDS.Rows(j).Cells("date_version").Value = DateTimeUtil.getShortDate(dt.Rows(j)("date_version").ToString())
                gridDS.Rows(j).Cells("date_update").Value = DateTimeUtil.getLongTimeDateVN(dt.Rows(j)("date_update").ToString())
            Next
        Catch ex As Exception
            MessageBox.Show("Lỗi: không load được danh sách các lần update chương trình:", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
End Class