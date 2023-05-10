Imports System.Data
Imports System.Data.SqlClient
Imports System
Imports System.IO
Imports Syncfusion.XlsIO
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Collections
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class frmTBNangBacLuong
    Private _BaoCaoBLL As clsHS_BaoCao = New clsHS_BaoCao()
    Private Sub frmTBNangBacLuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        chk_TrunguongQL.Visible = IIf(TRUCTHUOC = 1, True, False)
        cboDonVi.DataSource = listDonvi(True)
        txtNam.Text = Now.Year
        chk_TrunguongQL_CheckedChanged(sender, Nothing)
        If Not (cboKy Is Nothing) And cboKy.Items.Count <> 0 Then
            cboKy.SelectedIndex = DateTimeUtil.GetQuarter(DateTime.Now) - 1
        End If
        cboDonVi.Focus()
    End Sub

    Private Sub frmTBNangBacLuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
        Try
            Dim bIsBac_NgachLuong As Byte
            Dim ky As Short
            Dim lab_Err As String = ""
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            lab_Err = checkRpt()
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                labStatusProcess.Text = "Error"
                Cursor = Cursors.Default
                Exit Sub
            End If
           
            If rdBac.Checked Then
                bIsBac_NgachLuong = 1
                ky = CShort(cboKy.SelectedItem)
            Else
                bIsBac_NgachLuong = 2
            End If

            Dim bIsFlagAll As Byte = 0
            If chk_TrunguongQL.Checked Then
                bIsFlagAll = 2
            Else
                bIsFlagAll = IIf(cbAll.Checked, 1, 0)
            End If

            createReport(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), bIsFlagAll, bIsBac_NgachLuong, ky)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtNam.Text = Val(txtNam.Text.Trim)
    End Sub

    Private Sub createReport(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short, ByVal vLoai As SByte, ByVal vKy As Short)
        Dim db_source As DataTable = New DataTable()
        lbl_waitting.Text = ""
        If vLoai = 1 Then
            db_source.TableName = "01BC"
            db_source = _BaoCaoBLL.GetDanhSach_NangLuong(1, vIdDonVi, vYear, vKy, vAll, 1)
            'rpt_TBNangBac.SetDataSource(getTBNangBac(vIdDonVi, vYear, vAll, vKy))
            If (db_source Is Nothing) Then
                lbl_waitting.Text = "Không có cán bộ thuộc danh sách nâng bậc/ngạch lương"
                Return
            Else
                If db_source.Rows.Count <= 0 Then
                    lbl_waitting.Text = "Không có cán bộ thuộc danh sách nâng bậc/ngạch lương"
                    Return
                End If
            End If
            rpt_TBNangBac.SetDataSource(db_source)
            
            rpt_TBNangBac.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
            rpt_TBNangBac.SetParameterValue("tinh", DIABAN)
            rpt_TBNangBac.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
            rpt_TBNangBac.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
            rpt_TBNangBac.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
            rpt_TBNangBac.SetParameterValue("Thoidiem", "Kỳ " & vKy.ToString & " Năm " & txtNam.Text)
            rpt_View.ReportSource = rpt_TBNangBac
        Else
            db_source.TableName = "01BC"
            db_source = _BaoCaoBLL.GetDanhSach_NangLuong(1, vIdDonVi, vYear, vKy, vAll, 3)
            If (db_source Is Nothing) Then
                lbl_waitting.Text = "Không có cán bộ thuộc danh sách nâng bậc/ngạch lương"
                Return
            Else
                If db_source.Rows.Count <= 0 Then
                    lbl_waitting.Text = "Không có cán bộ thuộc danh sách nâng bậc/ngạch lương"
                    Return
                End If
            End If

            'rpt_TBNangNgach.SetDataSource(getTBNangNgach(vIdDonVi, vYear, vAll))
            rpt_TBNangNgach.SetDataSource(db_source)
            rpt_TBNangNgach.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
            rpt_TBNangNgach.SetParameterValue("tinh", DIABAN)
            rpt_TBNangNgach.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
            rpt_TBNangNgach.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
            rpt_TBNangNgach.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
            rpt_TBNangNgach.SetParameterValue("Thoidiem", txtNam.Text)
            rpt_View.ReportSource = rpt_TBNangNgach
        End If

    End Sub

    Private Function checkRpt() As String
        Dim strReturn As String = ""
        Try
            If (txtNam.Text = "") Or (txtNam.Text = "0") Then
                txtNam.Focus()
                strReturn = "Chưa nhập Năm lấy báo cáo!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub rdBac_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdBac.CheckedChanged
        If rdBac.Checked Then
            labKy.Visible = True
            cboKy.Visible = True
        Else
            labKy.Visible = False
            cboKy.Visible = False
        End If
    End Sub

    Private Sub chk_TrunguongQL_CheckedChanged(sender As Object, e As EventArgs) Handles chk_TrunguongQL.CheckedChanged
        cbAll.Enabled = True
        If chk_TrunguongQL.Checked Then
            cbAll.Checked = False
            cbAll.Enabled = False
        End If
    End Sub

    Private Sub btn_exportexcel_Click(sender As Object, e As EventArgs) Handles btn_exportexcel.Click
        Dim bIsFlagAll As Byte = 0
        If chk_TrunguongQL.Checked Then
            bIsFlagAll = 2
        Else
            bIsFlagAll = IIf(cbAll.Checked, 1, 0)
        End If

        Dim bQuyBC As Byte = CShort(cboKy.SelectedItem)
        Dim bLoaiDs As Byte = IIf(rdBac.Checked, 1, 2)
        Dim iDonViId As Integer = CInt(cboDonVi.SelectedValue)
        Dim db_danhsach As DataTable = New DataTable()
        db_danhsach = _BaoCaoBLL.GetDanhSach_NangLuong(bLoaiDs, iDonViId, CInt(txtNam.Text), bQuyBC, bIsFlagAll, IIf(bLoaiDs = 1, 2, 4))
        If (db_danhsach Is Nothing Or db_danhsach.Rows.Count <= 0) Then
            MessageBox.Show("Không có danh sách cán bộ nâng bậc/ngạch lương để xuất file excel. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If

        Cursor = Cursors.WaitCursor
        lbl_waitting.Text = "Waitting for data export..."

        'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Dim _IndexBC As Int32 = IIf(rdBac.Checked, 1, 2)
        'Dim _LoaiCB As Byte = IIf(_IdIndex = 1, 2, IIf(chk_TrunguongQL.Checked = True, 1, 0))

        Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
        Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
        Dim sPath_Export As String = "C:\Temp\"
        Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
        sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")

        sFileExcelTemplate = sPathExcelTemplate + IIf(_IndexBC = 1, "SK_DanhSach_NangBacLuong_01.xlsx", "SK_DanhSach_NangNgachLuong_01.xlsx")
        sFileNameEx = "Sao ke danh sach " + IIf(_IndexBC = 1, "Nang bac luong Quy_", "Nang ngach luong Nam_") + IIf(_IndexBC = 1, bQuyBC.ToString(), txtNam.Text) + "_" + sAutoNumber + ".xlsx"
        Dim iRowStart As Integer = 0, iRowA As Integer = 0

        If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
            Using excelEngine As ExcelEngine = New ExcelEngine()
                Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                Dim worksheet As IWorksheet = workbook.Worksheets(0)
                iRowStart = 7
                If _IndexBC = 1 Or _IndexBC = 2 Then
                    If db_danhsach.Rows(0)("ChiNhanh_HT").ToString() = "1" Then
                        worksheet.Range("A2").Text = "BAN TỔ CHỨC CÁN BỘ"
                    Else
                        worksheet.Range("A2").Text = db_danhsach.Rows(0)("ChiNhanh_HT").ToString().Trim().ToUpper()
                    End If
                    worksheet.Range("A3").Text = IIf(_IndexBC = 1, "THÔNG BÁO DANH SÁCH CÁN BỘ ĐẾN THỜI HẠN NÂNG BẬC LƯƠNG", "THÔNG BÁO DANH SÁCH CÁN BỘ ĐẾN THỜI HẠN NÂNG NGẠCH LƯƠNG")
                    worksheet.Range("A4").Text = IIf(_IndexBC = 1, "QUÝ " + bQuyBC.ToString() + " NĂM " + txtNam.Text.Trim(), " NĂM " + txtNam.Text.Trim())
                End If
                sColNameEnd = IIf(_IndexBC = 1, "N", "K")

                'Create Template Marker Processor
                Dim marker As ITemplateMarkersProcessor = workbook.CreateTemplateMarkersProcessor()

                Dim columnNames(db_danhsach.Columns.Count) As String
                Dim i As Integer = 0
                For Each column As DataColumn In db_danhsach.Columns
                    columnNames(i) = column.ColumnName
                    i += 1
                Next

                For Each c As DataColumn In db_danhsach.Columns
                    Dim columnName As String = c.ColumnName
                    Dim columnData As EnumerableRowCollection(Of Object)
                    If Globals.IsNumeric(c) Then
                        columnData = db_danhsach.AsEnumerable().[Select](Function(r) If(r.Field(Of Object)(columnName), 0))
                    ElseIf c.DataType Is GetType(Date) Then
                        columnData = db_danhsach.AsEnumerable.[Select](Function(r) If(r.Field(Of Object)(columnName), ""))
                    Else
                        columnData = db_danhsach.AsEnumerable.[Select](Function(r) If(("'" & r.Field(Of Object)(columnName)), ""))
                    End If

                    Dim paramDataArray As Object() = columnData.ToArray
                    marker.AddVariable(columnName, paramDataArray)
                Next

                marker.ApplyMarkers()
                workbook.Version = ExcelVersion.Excel2013

                'Căn chỉnh bôi đậm, nghiêng các dòng bản ghi và điền dữ liệu vào một số CELL
                If Not (db_danhsach Is Nothing) Then
                    If (db_danhsach.Rows.Count > 0) Then
                        For iTT As Integer = 0 To db_danhsach.Rows.Count - 1
                            iRowA = (iRowStart + 1 + iTT)
                            Dim iKieuIn As Integer = CType(db_danhsach.Rows(iTT)("KieuIn"), Integer)
                            If (iKieuIn = 0 Or iKieuIn = 1) Then
                                worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Bold = True
                                worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Color = ExcelKnownColors.Red
                            ElseIf (iKieuIn = 2) Then
                                worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Italic = True
                            ElseIf (iKieuIn = 4) Then
                                worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Italic = True
                                worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Bold = True
                            ElseIf (iKieuIn = 5) Then
                                worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Underline = ExcelUnderline.Single
                            End If
                        Next
                    End If
                End If

                If (Not System.IO.Directory.Exists(sPath_Export)) Then
                    Directory.CreateDirectory(sPath_Export)
                End If

                Dim file_name As String = sPath_Export + sFileNameEx
                If Globals.FileDao.IsFile(file_name) Then
                    Globals.FileDao.DeleteFile(file_name)
                End If
                Dim fs As FileStream = File.Create(file_name)
                workbook.SaveAs(fs)
                workbook.Close()
                streamRead.Close()
                fs.Close()
                MessageBox.Show("Dữ liệu được xuất ra file excel thành công theo đường dẫn: [" + file_name + "]", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End Using
        Else
            MessageBox.Show("Không lấy được file excel mẫu. Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
        'Trả lại thiết lập cũ
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        lbl_waitting.Text = ""
        Cursor = Cursors.Default
    End Sub

End Class