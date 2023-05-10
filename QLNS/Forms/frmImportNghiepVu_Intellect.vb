Option Explicit On
'Option Strict On

Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Text
Imports System
Imports System.Reflection

Public Class frmImportNghiepVu_Intellect

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

    '??c n?i dung t? file excel vao datagrid
    Private Sub ImportExcel(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim db As DBAccess = New DBAccess
        Dim strErr As String = ""

        Dim BatBuoc As Integer = 0
        Dim goc As Integer = 0
        Dim manghiepvu As String = ""
        Dim nghiepvu As String = ""
        Dim pre_CO As String = ""
        Dim pre_KHONG As String = ""
        Dim ghichu As String = ""

        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(1).Select()

                Dim vIdCB As String = ""
                row = 2
                While .cells(row, "A").value.ToString.Trim <> ""
                    If .cells(row, "A").value.ToString.Trim <> "" Then
                        Try
                            If .cells(row, "C").value Is Nothing Then
                                BatBuoc = 0
                            Else
                                BatBuoc = 1
                            End If
                            If .cells(row, "D").value Is Nothing Then
                                goc = 0
                            Else
                                goc = 1
                            End If
                            If .cells(row, "E").value Is Nothing Then
                                manghiepvu = ""
                            Else
                                manghiepvu = .cells(row, "E").value.ToString
                            End If
                            If .cells(row, "F").value Is Nothing Then
                                nghiepvu = ""
                            Else
                                nghiepvu = .cells(row, "F").value.ToString
                            End If
                            If .cells(row, "G").value Is Nothing Then
                                pre_CO = ""
                            Else
                                pre_CO = .cells(row, "G").value.ToString
                            End If
                            If .cells(row, "H").value Is Nothing Then
                                pre_KHONG = ""
                            Else
                                pre_KHONG = .cells(row, "H").value.ToString
                            End If
                            If .cells(row, "I").value Is Nothing Then
                                ghichu = ""
                            Else
                                ghichu = .cells(row, "I").value.ToString
                            End If
                            db.executeSQL("INSERT INTO tbl_NghiepVu (BatBuoc,Goc,MaNghiepVu,NghiepVu,pre_DK_CO, pre_DK_KHONG,GhiChu) values (" & BatBuoc & ", " & goc & ", N'" & manghiepvu & "', N'" & nghiepvu & "', N'" & pre_CO & "', N'" & pre_KHONG & "', N'" & ghichu & "')")
                        Catch ex As Exception

                        End Try
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

        SetProgress(8)
        Dim sqlDSCanBo As String = "SELECT * FROM tbl_NghiepVu "
        Dim dt As DataTable
        Dim j As Integer = 0
        dt = db.SelectDBRows(sqlDSCanBo)
        For j = 0 To dt.Rows.Count - 1
            myDataGrid.Rows.Add()
            myDataGrid.Rows(j).Cells("STT").Value = j + 1
            myDataGrid.Rows(j).Cells("BatBuoc").Value = IIf(dt.Rows(j)("BatBuoc").ToString() <> "", dt.Rows(j)("BatBuoc").ToString(), "")
            myDataGrid.Rows(j).Cells("Goc").Value = IIf(dt.Rows(j)("Goc").ToString() <> "", dt.Rows(j)("Goc").ToString(), "")
            myDataGrid.Rows(j).Cells("MaNghieVu").Value = dt.Rows(0)("MaNghieVu").ToString()
            myDataGrid.Rows(j).Cells("NghiepVu").Value = dt.Rows(0)("NghiepVu").ToString()
            myDataGrid.Rows(j).Cells("Pre_DK_CO").Value = IIf(dt.Rows(j)("Pre_DK_CO").ToString() <> "", dt.Rows(j)("Pre_DK_CO").ToString(), "")
            myDataGrid.Rows(j).Cells("Pre_DK_KHONG").Value = IIf(dt.Rows(j)("Pre_DK_KHONG").ToString() <> "", dt.Rows(j)("Pre_DK_KHONG").ToString(), "")
            myDataGrid.Rows(j).Cells("GhiChu").Value = dt.Rows(0)("GhiChu").ToString()
        Next

        SetProgress(10)

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_DaiMa(ByVal PathExcelFile As String)

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
                Dim col_end_CU As Integer = 0
                Dim daiDuPhong As Integer = 0
                Dim ID_CU As Integer = 0
                row = 1
                While .cells(row, "A").value.ToString.Trim <> ""
                    If row = 1 Then
                        ID_CU = .cells(row, "A").value
                        If ID_CU > 0 Then db.executeSQL("UPDATE ChiNhanh set MaDV='" & .cells(row, "C").value.ToString & "', MaCB_Begin= 1, MaCB_End=" & .cells(row, "D").value & " WHERE ID=" & .cells(row, "A").value)
                    Else
                        If .cells(row, "A").value = 99 Then
                            daiDuPhong = .cells(row, "D").value
                        Else
                            daiDuPhong = 0
                        End If
                        col_end_CU = db.getNumber("SELECT MaCB_end FROM Chinhanh WHERE id=" & ID_CU)
                        db.executeSQL("UPDATE ChiNhanh set MaDV='" & .cells(row, "C").value & "', MaCB_Begin= " & daiDuPhong + col_end_CU + 1 & ", MaCB_End=" & col_end_CU + .cells(row, "D").value & " WHERE ID=" & .cells(row, "A").value)
                        ID_CU = .cells(row, "A").value
                    End If
                    row = row + 1
                End While
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        End Try
        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show(strErr, "Thông báo")
        End If

    End Sub

    Private Sub addCol(ByRef vGrid As DataGridView, ByVal vHeaderText As String, ByVal vDataPropertyName As String, ByVal vName As String)
        Dim colTxt As New DataGridViewTextBoxColumn()
        colTxt.DataPropertyName = vDataPropertyName
        colTxt.HeaderText = vHeaderText
        colTxt.Name = vName
        colTxt.ReadOnly = True
        vGrid.Columns.Add(colTxt)
    End Sub

    'Private Sub initGrid()
    '    gridDS_CB.Columns().Clear()
    '    gridDS_CB.Columns.Add("STT", "STT")
    '    gridDS_CB.Columns.Add("MaCB", "Mã Cán b?")
    '    gridDS_CB.Columns.Add("HoTen", "H? tên")
    '    gridDS_CB.Columns.Add("GioiTinh", "Gi?i tính")
    '    gridDS_CB.Columns.Add("NgaySinh", "Ngày sinh")
    '    gridDS_CB.Columns.Add("CMT_So", "S? CMT")
    '    gridDS_CB.Columns.Add("ChiNhanh", "??n v?")
    '    gridDS_CB.Columns.Add("Phong", "Phòng")
    '    'C?n ch?nh tiêu ??
    '    gridDS_CB.Columns("STT").Width = 50
    '    gridDS_CB.Columns("HoTen").Width = 188
    '    gridDS_CB.Columns("ChiNhanh").Width = 188
    '    gridDS_CB.Columns("Phong").Width = 188
    '    gridDS_CB.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    '    gridDS_CB.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    '    gridDS_CB.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    '    gridDS_CB.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

    'End Sub

    Private Sub initGrid(ByVal sLoaiHS As Integer, ByVal myDataGrid As DataGridView)
        myDataGrid.Columns().Clear()
        myDataGrid.Columns.Add("STT", "STT")
        myDataGrid.Columns.Add("BatBuoc", "Bat buoc")
        myDataGrid.Columns.Add("Goc", "Goc")
        myDataGrid.Columns.Add("MaNghiepVu", "MaNghiepVu")
        myDataGrid.Columns.Add("NghiepVu", "NghiepVu")
        myDataGrid.Columns.Add("Pre_DK_CO", "Pre_DK_CO")
        myDataGrid.Columns.Add("Pre_DK_KHONG", "Pre_DK_KHONG")
        myDataGrid.Columns.Add("GhiChu", "GhiChu")
        myDataGrid.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        myDataGrid.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        myDataGrid.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub bntImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImport.Click
        Try
            ShowProgressBar()
            ProgressBar1.Maximum = 10
            SetProgress(1)
            Me.Refresh()

            initGrid(0, gridDS_CB)
            SetProgress(3)
            ImportExcel(txtPath.Text, gridDS_CB)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmImportNghiepVu_Intellect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("S? d?ng phím t?t: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmImportNghiepVu_Intellect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HideProgressBar()
        If DONVI = gMaDonViTW Then
            bntImportMaCB.Visible = True
        Else
            bntImportMaCB.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImportMaCB.Click
        Try
            ShowProgressBar()
            ProgressBar1.Maximum = 10
            SetProgress(1)
            Me.Refresh()

            initGrid(0, gridDS_CB)
            'initGrid()
            SetProgress(3)
            'ImportExcel(txtPath.Text, gridDS_CB)
            'ImportExcel_TCCB(txtPath.Text, gridDS_CB)
            ImportExcel_DaiMa(txtPath.Text)
        Catch ex As Exception

        End Try
    End Sub
End Class