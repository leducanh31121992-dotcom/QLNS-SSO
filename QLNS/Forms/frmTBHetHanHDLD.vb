Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core

Public Class frmTBHetHanHDLD

    Private ComDset As New DataSet
    Private arrWHERE As ArrayList = New ArrayList()

    Private Sub frmTBHetHanHDLD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmTBHetHanHDLD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True)
        cboDonVi.Focus()
        AddHandler cmdReport.Click, AddressOf cmdReportClicked
    End Sub

    Private Sub cmdReportClicked(ByVal source As Object, ByVal e As EventArgs)
        Try
            Dim all As Short
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            If cbAll.Checked Then
                all = 1
            Else
                all = 0
            End If
            bindGridResult(CInt(cboDonVi.SelectedValue), all, dpkDateline.Value)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub createReport(ByVal vIdDonVi As Integer, ByVal vAll As Short, ByVal vNgayhetHan As DateTime)
        Dim dt As System.Data.DataTable
        dt = getTBHetHanHDLD(vIdDonVi, vAll, vNgayhetHan)
        gridResult.AutoGenerateColumns = True
        gridResult.RowHeadersVisible = False
        gridResult.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        gridResult.AllowUserToAddRows = False

        gridResult.Columns().Clear()
        gridResult.DataSource = dt
        'grpResult.Text = "Kết quả tra cứu _ Số bản ghi tìm thấy: " & ComDset.Tables(0).Rows.Count
    End Sub

    Private Sub bindGridResult(ByVal vIdDonVi As Integer, ByVal vAll As Short, ByVal vNgayHetHan As DateTime)
        Try
            Try
                Dim db As DBAccess = New DBAccess
                Dim dbconn As DBAccess = New DBAccess
                Dim cmd As SqlCommand
                Dim conn As SqlClient.SqlConnection = dbconn.getConnection
                Dim adp As SqlClient.SqlDataAdapter

                Try
                    cmd = New SqlCommand("alert_HetHanHDLD")
                    cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
                    cmd.Parameters.Add(New SqlParameter("@All", vAll))
                    cmd.Parameters.Add(New SqlParameter("@Dateline", vNgayHetHan))
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Connection = conn
                    conn.Open()
                    adp = New SqlDataAdapter(cmd)
                    ComDset.Reset()
                    adp.Fill(ComDset, "TTbl")
                    If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
                        Exit Sub
                    End If
                    gridResult.AutoGenerateColumns = True
                    gridResult.RowHeadersVisible = False
                    gridResult.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                    gridResult.AllowUserToAddRows = False

                    gridResult.Columns().Clear()
                    gridResult.DataSource = ComDset.Tables(0)
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

                SaveToExcel(.ActiveCell.Worksheet, "TBHetHanHDLD_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls", True)

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

    Private Sub blankFrm()
        dpkDateline.Enabled = False
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

End Class