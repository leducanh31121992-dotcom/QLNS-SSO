Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.IO


Public Class frmExChamCong

    Private Sub frmExChamCong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Excel As Object = CreateObject("Excel.Application")
        Dim ComDsetCB As New DataSet
        Dim ComDsetCC As New DataSet
        Dim db As DBAccess = New DBAccess
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlClient.SqlConnection = dbconn.getConnection
        Dim strSQL As String = "SELECT MaCB, HoTen FROM HS_CanBo"
        Dim adpCB As New SqlClient.SqlDataAdapter(strSQL, conn)
        strSQL = "SELECT Ten_goi FROM DanhMuc Where Id_goc=34"
        Dim adpCC As New SqlClient.SqlDataAdapter(strSQL, conn)

        If Excel Is Nothing Then
            MessageBox.Show("Excel chưa được cài đặt trên máy. Để tiếp tục yêu cầu cài đặt MS Excel", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        ComDsetCB.Reset()
        adpCB.Fill(ComDsetCB, "TTbl")
        adpCC.Fill(ComDsetCC, "TTbl")
        If ComDsetCB.Tables.Count < 0 Or ComDsetCB.Tables(0).Rows.Count <= 0 Then
            MsgBox("Không có dữ liệu để xuất file")
            Exit Sub
        End If
        exSheet(Excel, ComDsetCB, ComDsetCC)

    End Sub

    Private Sub exSheet(ByVal Excel As Object, ByVal dsCB As DataSet, ByVal dsChamCong As DataSet)
        Dim col As Integer = 0
        Dim row As Integer = 0
        Dim vDate As Date = DateTimeUtil.getDate("01/" & Now.Month & Now.Year)
        Dim aDate As Date = vDate.AddMonths(1)
        Dim days As Integer = aDate.AddDays(-1).Day
        Dim i As Integer = 1
        Dim totalColsheet As Integer = 2 + days + dsChamCong.Tables(0).Rows.Count
        'Try

        '    With Excel
        '        'không cho hiện ứng dụng Excel lên để tránh gây đơ máy
        '        .Visible = False
        '        .SheetsInNewWorkbook = 1
        '        .Workbooks.Add()
        '        .Worksheets(1).Select()
        '        '.Worksheets(1).Columns.AutoFit()
        '        '.Worksheets(1).Columns.EntireColumn.AutoFit()
        '        'caption.FormulaR1C1 = tieude;

        '        'căn lề cho tiêu đề
        '        .Worksheets(1).HorizontalAlignment = Excel.Constants.xlCenter

        '        'Noi dung bang bieu bat dau tu dong thu 3
        '        For row = 3 To dsCB.Tables(0).Rows.Count + 4
        '            If row = 0 Then
        '                'Header
        '                For col = 0 To totalColsheet
        '                    .cells(row, col).Font.Size = 12
        '                    .cells(row, col).Font.Bold = True
        '                    .cells(row, col).EntireColumn.AutoFit()
        '                Next
        '            Else

        '            End If

        '        Next
        '        For col = 0 To totalColsheet
        '            .cells(1, i).value = ds.Tables(0).Columns(col).ColumnName
        '            .cells(1, i).Font.Size = 12
        '            .cells(1, i).Font.Bold = True
        '            '.cells(1, i).DegreeAlignment = 90
        '            '.cells(1, i).VerticalAlignment = -4108
        '            '.cells(1, i).HorizontalAlignment = -4108
        '            'wb.Worksheets("Sheet1").Columns("B:B").NumberFormat = "m/d/yyyy;@
        '            .cells(1, i).EntireColumn.AutoFit()


        '            'If i = 1 Then
        '            '    .cells(1, i).ColumnWidth = 15
        '            'Else
        '            '    If i = 2 Then
        '            '        .cells(1, i).ColumnWidth = 25
        '            '    Else
        '            '        .cells(1, i).ColumnWidth = 20
        '            '    End If
        '            'End If
        '            i += 1
        '        Next
        '        i = 2
        '        Dim k As Integer = 1
        '        For col = 0 To ComDset.Tables(0).Columns.Count - 1
        '            i = 2
        '            For row = 0 To ComDset.Tables(0).Rows.Count - 1
        '                .Cells(i, k).Value = ComDset.Tables(0).Rows(row).ItemArray(col)
        '                i += 1
        '            Next
        '            k += 1
        '        Next
        '        If Not (System.IO.Directory.Exists(gFileExported)) Then
        '            System.IO.Directory.CreateDirectory(gFileExported)
        '        End If
        '        filename = gFileExported & "TraCuu_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
        '        .ActiveCell.Worksheet.SaveAs(filename)
        '    End With

        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
        '    Excel = Nothing
        '    MessageBox.Show("Dữ liệu đã được xuất ra Excel thành công tại địa chỉ " & filename, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        '    MsgBox(ex.Message)
        'End Try

        'Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        'For Each i As Process In pro
        '    i.Kill()
        'Next
    End Sub
End Class