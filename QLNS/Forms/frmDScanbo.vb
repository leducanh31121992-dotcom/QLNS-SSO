Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core

Public Class frmDScanbo

    'Khai báo biến cho từng loại tìm kiếm
    '=1: tìm kiếm Danh sách
    '=2: thống kê làm Thêm giờ
    Public ID_Search As Integer = 0

    Private ComDset As New DataSet
    Private arrWHERE As ArrayList = New ArrayList()

    Private Sub frmDScanbo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmDScanbo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True)
        If ID_Search = 1 Then
            Me.Text = "Quản lý nhân sự: DANH SÁCH CÁN BỘ VIÊN CHỨC"
            grpResult.Text = "Danh sách cán bộ viên chức"
            labThoiDiem.Text = "Đến ngày"
            dpkDenNgay.Visible = True
            dtpkNam.Visible = False
        Else
            Me.Text = "Quản lý nhân sự: THỐNG KÊ LÀM THÊM GIỜ"
            grpResult.Text = "Kết quả thống kê theo năm"
            labThoiDiem.Text = "Năm"
            dpkDenNgay.Visible = False
            dtpkNam.Visible = True
        End If
        AddHandler bntSearch.Click, AddressOf btnUpdateClicked
    End Sub

    Private Sub bntExportFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExportFile.Click
        Try
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            Load_Excel_Details(ComDset)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntRefresh.Click
        blankFrm()
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub blankFrm()
        dpkDenNgay.Enabled = False
    End Sub

    Private Sub bindGridResult(ByVal vIdDonVi As Integer, ByVal vDenNgay As Date, ByVal vNam As Integer)
        Try
            Try
                Dim db As DBAccess = New DBAccess
                Dim dbconn As DBAccess = New DBAccess
                Dim cmd As SqlCommand
                Dim conn As SqlClient.SqlConnection = dbconn.getConnection
                Dim adp As SqlClient.SqlDataAdapter

                Try
                    If ID_Search = 1 Then
                        cmd = New SqlCommand("getListCanBo")
                        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
                        cmd.Parameters.Add(New SqlParameter("@ThoiDiem", vDenNgay))
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Connection = conn
                        conn.Open()
                    Else
                        cmd = New SqlCommand("getListThemGio")
                        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
                        cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Connection = conn
                        conn.Open()
                    End If
                    adp = New SqlDataAdapter(cmd)

                    ComDset.Reset()
                    adp.Fill(ComDset, "TTbl")
                    If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
                        'grpResult.Text = "Danh sách cán bộ viên chức: 0"
                        Exit Sub
                    End If
                    gridResult.AutoGenerateColumns = True
                    gridResult.RowHeadersVisible = False
                    gridResult.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                    gridResult.AllowUserToAddRows = False

                    gridResult.Columns().Clear()
                    gridResult.DataSource = ComDset.Tables(0)
                    'Làm đẹp
                    Grid_Process()
                    'grpResult.Text = "Danh sách cán bộ viên chức: " & ComDset.Tables(0).Rows.Count
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                Finally
                    dbconn.closeConnection(conn)
                End Try
            Catch ex As Exception
                grpResult.Text = "Danh sách cán bộ viên chức: 0"
                MessageBox.Show("Kết quả tra cứu không có: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            End Try
        Catch ex As Exception
            MessageBox.Show("Không load được danh sách: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub Load_Excel_Details(ByVal ds As DataSet)

        Dim filename As String
        Dim col, row As Integer
        If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
            Exit Sub
        End If
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MessageBox.Show("Excel chưa được cài đặt trên máy. Để tiếp tục yêu cầu cài đặt MS Excel", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If

        Try
            'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            With Excel
                .Visible = False
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()

                Dim i As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    .cells(1, i).value = ComDset.Tables(0).Columns(col).ColumnName
                    .cells(1, i).Font.Size = 12
                    .cells(1, i).Font.Bold = True
                    .cells(1, i).EntireColumn.AutoFit()
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    i = 2
                    For row = 0 To ComDset.Tables(0).Rows.Count - 1
                        .Cells(i, k).Value = ComDset.Tables(0).Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next

                If ID_Search = 1 Then
                    filename = "DSCanBo_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                Else
                    filename = "TKThemGio_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                End If

                SaveToExcel(.ActiveCell.Worksheet, filename, True)

                'If Not (System.IO.Directory.Exists(gFileExported)) Then
                '    System.IO.Directory.CreateDirectory(gFileExported)
                'End If
                'If ID_Search = 1 Then
                '    filename = gFileExported & "DSCanBo_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                'Else
                '    filename = gFileExported & "TKThemGio_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                'End If
            End With
            'Trả lại thiết lập cũ
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            MsgBox(ex.Message)
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub btnUpdateClicked(ByVal source As Object, ByVal e As EventArgs)
        Try
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            If ID_Search = 1 Then
                bindGridResult(CInt(cboDonVi.SelectedValue), dpkDenNgay.Value.Date, 0)
            Else
                bindGridResult(CInt(cboDonVi.SelectedValue), Nothing, dtpkNam.Value.Year)
            End If

            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Function DSCanBo(ByVal vIdDonVi As Integer, ByVal vDenNgay As Date) As System.Data.DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd = New SqlCommand("getListCanBo")
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@ThoiDiem", vDenNgay))
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbDS")
            Return ds.Tables("tbDS")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Sub Grid_Process()
        Dim checkCol As Integer = -1
        'Format kết quả cho đẹp
        With gridResult
            If .Rows.Count > 1 Then
                Select Case ID_Search
                    Case 1
                        checkCol = 1
                    Case 2
                        checkCol = 0
                End Select
                If checkCol > -1 Then
                    For i As Integer = 0 To .Rows.Count - 1
                        If My_CInt(.Item(checkCol, i).Value, -1) = -1 Then
                            .Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(.Font, FontStyle.Bold)
                            .Rows(i).DefaultCellStyle.BackColor = Color.LightGray
                        End If
                    Next
                End If
            End If
        End With
    End Sub
End Class