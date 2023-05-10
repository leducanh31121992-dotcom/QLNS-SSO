Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class frmChuyenDoiMaCB
    Private ComDset As New DataSet

    Private Sub frmChuyenDoiMaCB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmChuyenDoiMaCB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim _roles As String = Globals.Roles
        initGrid()
        bindGrid(True)
        'Nếu có quyền với HS cán bộ thì có quyền chuyển đổi mã
        If Not (Globals.IsIntersect(";71;72;73;74;", _roles)) Then
            Panel1.Visible = False
        End If

    End Sub

    Private Sub HideProgressBar()
        ProgressBar1.Visible = False
    End Sub

    Private Sub ShowProgressBar()
        ProgressBar1.Visible = True
    End Sub

    Private Sub SetProgress(ByVal iVal As Integer)
        ProgressBar1.Value = iVal
        Me.Refresh()
    End Sub

    Private Sub initGrid()
        gridDS_CB.Columns().Clear()
        gridDS_CB.Columns.Add("STT", "STT")
        gridDS_CB.Columns.Add("HoTen", "Họ tên")
        gridDS_CB.Columns.Add("DonVi", "Đơn vị")
        gridDS_CB.Columns.Add("Phong", "Phòng/Ban")
        gridDS_CB.Columns.Add("ChucDanh", "Chức vụ")
        gridDS_CB.Columns.Add("MaCB_Cu", "Mã hiện tại")
        gridDS_CB.Columns.Add("MaCB_Moi", "Mã mới")
        
        'Căn chỉnh tiêu đề
        gridDS_CB.Columns("STT").Width = 40
        gridDS_CB.Columns("HoTen").Width = 150
        gridDS_CB.Columns("DonVi").Width = 150 'System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        gridDS_CB.Columns("Phong").Width = 120
        gridDS_CB.Columns("ChucDanh").Width = 150
        gridDS_CB.Columns("MaCB_Cu").Width = 90
        gridDS_CB.Columns("MaCB_Moi").Width = 90
        'gridDS_CB.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        'gridDS_CB.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        gridDS_CB.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        gridDS_CB.Columns(6).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub bindGrid(ByVal vMa_moi As Boolean)
        Try
            Dim db As DBAccess = New DBAccess
            Dim strSQL As String = "SELECT IdCanbo, HoTen, MaCB FROM HS_canbo WHERE IdDonvi=" & IdDONVI
            Dim dt As DataTable
            Dim j As Integer = 0
            Dim MaCB_Begin As Integer = db.getNumber("SELECT MaCB_Begin FROM ChiNhanh WHERE Id=" & IdDONVI)
            Dim MaCB_End As Integer = db.getNumber("SELECT MaCB_End FROM ChiNhanh WHERE Id=" & IdDONVI)
            strSQL = "SELECT t2.IdCanbo, HoTen, MaCB " & _
                     "FROM HS_Canbo t2, " & _
                     "	   (SELECT t1.idCanbo, t1.idDonvi_Moi, t1.idPhong_Moi, t1.IdChucVu_Moi " & _
                     "	    FROM QDNhansu t1 " & _
                     "      WHERE(t1.isKiemNhiem = 0 AND t1.IsQD_NHCS=1) " & _
                     "	      AND t1.idCanbo not in (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t1.idCanbo AND IsQD_NHCS=1) " & _
                     "	      AND t1.idcanbo not in (SELECT idcanbo FROM QDNhansu t4 WHERE t4.idCanbo=t1.idCanbo AND datediff(second,t1.ngayHL,t4.ngayHL)>0)) t   " & _
                     " WHERE t.idCanbo = t2.idCanbo AND IdDonvi=" & IdDONVI & _
                     " Order by idDonvi_Moi, idPhong_Moi, idChucvu_moi, Hoten "
            dt = db.SelectDBRows(strSQL)
            'gridDS_CB.Columns().Clear()
            gridDS_CB.DataSource = Nothing
            gridDS_CB.Refresh()
            For j = 0 To dt.Rows.Count - 1
                gridDS_CB.Rows.Add()
                gridDS_CB.Rows(j).Cells("STT").Value = j + 1
                gridDS_CB.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                Dim dt1 As DataTable
                dt1 = db.SelectDBRows("SELECT TOP 1 (SELECT ten_goi FROM ChiNhanh where Id=IdDonvi_Moi) as Chinhanh, (SELECT ten_phong FROM PhongBan WHERE Id=IdPhong_Moi) as Phong, (SELECT ten_goi FROM DanhMuc WHERE id_goc=14 and id=Idchucvu_moi) as ChucDanh FROM QDNhansu WHERE IdCanBo='" & dt.Rows(j)("IdCanBo").ToString() & "' order by NgayHL desc")
                gridDS_CB.Rows(j).Cells("DonVi").Value = dt1.Rows(0)("Chinhanh").ToString()
                gridDS_CB.Rows(j).Cells("Phong").Value = dt1.Rows(0)("Phong").ToString()
                gridDS_CB.Rows(j).Cells("ChucDanh").Value = dt1.Rows(0)("ChucDanh").ToString()
                gridDS_CB.Rows(j).Cells("MaCB_Cu").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                If vMa_moi Then
                    If MaCB_Begin < MaCB_End Then
                        gridDS_CB.Rows(j).Cells("MaCB_Moi").Value = formatLenString(5, MaCB_Begin.ToString)
                        MaCB_Begin = MaCB_Begin + 1
                    Else
                        gridDS_CB.Rows(j).Cells("MaCB_Moi").Value = ""
                    End If
                Else
                    gridDS_CB.Rows(j).Cells("MaCB_Moi").Value = ""
                End If

            Next
            If vMa_moi Then
                labDS.Text = "Tổng số cán bộ được đổi mã: " & dt.Rows.Count.ToString & " cán bộ."
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi: không load được danh sách cán bộ:", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntConvert.Click
        If MessageBox.Show("Bạn có đồng ý đổi mã cán bộ sang mã mới không?", "Xác nhận", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.OK Then
            ' Uu tien doi ma theo file excel, co dinh dang: cot A: ma cu cua CB, cot B: ma moi cua CB
            If txtPath.Text.Trim = "" Then
                Dim db As DBAccess = New DBAccess
                Dim cmd As SqlCommand = New SqlCommand("convertMaCB")
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDONVI))
                db.executeSQL(cmd)
            Else
                ImportExcel_MaCB(txtPath.Text, gridDS_CB)
            End If
            bindGrid(False)
            
        End If
    End Sub

    Private Sub cmdChonTM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChonTM.Click
        Dim openFile As New OpenFileDialog
        HideProgressBar()
        Try
            openFile.Filter = "Excel Files (*.xls)|*.xls"
            openFile.ShowDialog()
            txtPath.Text = openFile.FileName
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ImportExcel_MaCB(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)
        ' Cau truc file excel: Cot A: ID CB; cot D: maCB moi
        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
        Dim dsCanBo As String = ""
        Dim db As DBAccess = New DBAccess
        Dim strErr As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(1).Select()

                Dim vIdCB As String = ""
                row = 2
                While .cells(row, "A").value.ToString.Trim <> ""
                    If .cells(row, "A").value.ToString.Trim <> "" And .cells(row, "D").value.ToString.Trim <> "" Then
                        Dim IDCanBo As String = ""
                        Dim MaCB_Moi As String = ""
                        If Not (.cells(row, "A").value Is Nothing) Then
                            IDCanBo = .cells(row, "A").value.ToString.Trim
                        End If
                        If Not (.cells(row, "D").value Is Nothing) Then
                            MaCB_Moi = .cells(row, "D").value.ToString.Trim
                        End If
                        If IDCanBo <> "" And MaCB_Moi <> "" Then db.executeSQL("UPDATE HS_Canbo set Macb='" & MaCB_Moi & "' WHERE IDCanbo='" & IDCanBo & "'")
                    End If
nex:
                    row = row + 1
                End While
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub
End Class