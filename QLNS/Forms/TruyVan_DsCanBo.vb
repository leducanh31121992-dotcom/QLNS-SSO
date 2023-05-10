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

Public Class TruyVan_DsCanBo
    Private _ChiNhanhBLL As ChiNhanhBLL = New ChiNhanhBLL()
    Private _HsCanBoBLL As HsCanBoBLL = New HsCanBoBLL()
    Private db_dscanbo As System.Data.DataTable = New System.Data.DataTable()

    Private Sub TruyVan_DsCanBo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Bind Danh sách loại cán bộ
        cb_loaicb.Items.Clear()
        cb_loaicb.Items.Add("Chuyên môn nghiệp vụ")
        cb_loaicb.Items.Add("Bảo vệ, tạp vụ")
        If (cb_loaicb.Items.Count <> 0) Then
            cb_loaicb.SelectedIndex = 0
        End If
        'Ẩn hiện tích chọn Cán bộ do TW quản lý khi chạy theo cấp Chi nhánh hoặc Toàn quốc
        chk_TrunguongQL.Visible = IIf(TRUCTHUOC = 1, True, False)
        chk_TrunguongQL.Checked = False
        'Thực hiện Bind dữ liệu Danh sách đơn vị trực thuộc
        chk_TrunguongQL_CheckedChanged(sender, Nothing)

        cb_donvi.DisplayMember = "Display"
        cb_donvi.ValueMember = "Value"
        dtpk_thoidiem.Text = Globals.GetDateTime_ForServerDB()
        lbl_waitting.Text = ""
    End Sub

    Private Sub chk_TrunguongQL_CheckedChanged(sender As Object, e As EventArgs) Handles chk_TrunguongQL.CheckedChanged
        If (chk_TrunguongQL.Checked = True) Then
            cb_donvi.DataSource = _ChiNhanhBLL.GetListComBo_ChiNhanh(2, "")
        Else
            cb_donvi.DataSource = _ChiNhanhBLL.GetListComBo_ChiNhanh(1, IIf(TRUCTHUOC = 0, "", "---Danh sách đơn vị ---"))
        End If
        If (cb_donvi.Items.Count <> 0) Then
            cb_donvi.SelectedIndex = 0
        End If
    End Sub

    Private Sub TruyVan_DsCanBo_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                db_dscanbo.Dispose()
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub btn_huybo_Click(sender As Object, e As EventArgs) Handles btn_huybo.Click
        If Not (cb_donvi Is Nothing) Then
            If cb_donvi.Items.Count > 0 Then
                cb_donvi.SelectedIndex = 0
            End If
        End If
        If Not (cb_loaicb Is Nothing) Then
            If (cb_loaicb.Items.Count <> 0) Then
                cb_loaicb.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        db_dscanbo.Dispose()
        Close()
    End Sub

    Private Sub cb_loaicb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_loaicb.SelectedIndexChanged
        If (Not (cb_loaicb Is Nothing) And (cb_loaicb.SelectedIndex >= 0) And (cb_loaicb.Items.Count <> 0)) Then
            Dim _IdIndex As Int32 = cb_loaicb.SelectedIndex
            If _IdIndex = 0 Then    'Nếu chọn là cán bộ chuyên môn nghiệp vụ
                chk_TrunguongQL.Checked = False
                chk_TrunguongQL.Visible = IIf(TRUCTHUOC = 1, True, False)
            Else                    'Nếu chọn là cán bộ Bảo vệ, tạp vụ
                chk_TrunguongQL.Checked = False
                chk_TrunguongQL.Visible = False
            End If
            chk_TrunguongQL_CheckedChanged(sender, Nothing)
        End If
    End Sub

    Private Sub btn_tracuu_Click(sender As Object, e As EventArgs) Handles btn_tracuu.Click
        Try
            If (Not (cb_loaicb Is Nothing) And (cb_loaicb.SelectedIndex >= 0) And (cb_loaicb.Items.Count <> 0)) Then
                If chk_TrunguongQL.Checked = False And cb_loaicb.SelectedIndex = 0 And TRUCTHUOC = 1 Then
                    If ((cb_donvi Is Nothing) Or (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0)) Then
                        MessageBox.Show("Bạn chưa chọn đơn vị cần truy vấn thông tin!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_donvi
                        Return
                    End If
                End If
                Cursor = Cursors.WaitCursor
                lbl_waitting.Text = "Waitting ..."
                Dim _IdIndex As Int32 = cb_loaicb.SelectedIndex
                Dim _LoaiCB As Byte = IIf(_IdIndex = 1, 2, IIf(chk_TrunguongQL.Checked = True, 1, 0))
                Dim _IdDonVi As Integer = CInt(cb_donvi.SelectedValue)
                Dim _ThoiDiem As DateTime = DateTimeUtil.getDatetime(dtpk_thoidiem.Value.ToString("dd/MM/yyyy 23:59:59"))
                lbl_tonghop.Text = ""
                dgv_main.AutoGenerateColumns = True
                dgv_main.Columns.Clear()
                dgv_main.ColumnHeadersHeight = 35
                dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                dgv_main.Columns().Clear()

                If Not (db_dscanbo Is Nothing) Then
                    db_dscanbo.Reset()
                End If

                db_dscanbo = _HsCanBoBLL.GetDsCanBos(_LoaiCB, _IdDonVi, _ThoiDiem)

                dgv_main.DataSource = db_dscanbo
                FixGrid_Columns()
                dgv_main.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(204, 209, 209)
                Dim iCountRows As Integer = 0
                'Bôi đậm mầu các dòng Mục lớn: A, B, C, ... Hoặc I, II, III ...
                If Not (dgv_main.DataSource Is Nothing) Then
                    If dgv_main.Rows.Count > 1 Then
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            dgv_main.Rows(i).DefaultCellStyle.ForeColor = Color.Navy
                            Dim _STT As String = dgv_main.Item("STT", i).Value.ToString().Trim()
                            Dim sampleArray As New List(Of String)(New String() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ"})
                            If InStr(_STT, "CN") > 0 Or _STT = "I" Or _STT = "II" Or _STT = "III" Or _STT = "IV" Or _STT = "V" Or _STT = "VI" Or _STT = "VII" Or _STT = "VIII" Or _STT = "IX" Or _STT = "X" Or _STT = "XI" Or _STT = "XII" Or _STT = "XIII" Or _STT = "XIV" Or _STT = "XV" Or _STT = "XVI" Or _STT = "XVII" Or _STT = "XVIII" Or _STT = "XIX" Or _STT = "XX" Or _STT = "XXI" Or _STT = "XXII" Or _STT = "XXIII" Or _STT = "XXIV" Or _STT = "XXV" Or _STT = "XXVI" Or _STT = "XXVII" Or _STT = "XXVIII" Or _STT = "XXIX" Or _STT = "XXX" Or _STT = "XXXI" Or _STT = "XXXII" Or _STT = "XXIII" Or _STT = "XXXIV" Or _STT = "XXXV" Then
                                dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Bold)
                                dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Italic)
                                dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.Orange '.FromArgb(242, 249, 3) 'Color.FromArgb(97, 135, 214) '
                            Else
                                If sampleArray.Contains(_STT) Then          'If _STT = "A" Or _STT = "B" Or _STT = "C" Then
                                    dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Bold)
                                    dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(97, 135, 214)
                                Else
                                    iCountRows = iCountRows + 1
                                End If
                            End If


                            'If sampleArray.Contains(_STT) Then          'If _STT = "A" Or _STT = "B" Or _STT = "C" Then
                            '    dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Bold)
                            '    dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(97, 135, 214)
                            'Else
                            '    If InStr(_STT, "CN") > 0 Or _STT = "I" Or _STT = "II" Or _STT = "III" Or _STT = "IV" Or _STT = "V" Or _STT = "VI" Or _STT = "VII" Or _STT = "VIII" Or _STT = "IX" Or _STT = "X" Or _STT = "XI" Or _STT = "XII" Or _STT = "XIII" Or _STT = "XIV" Or _STT = "XV" Or _STT = "XVI" Or _STT = "XVII" Or _STT = "XVIII" Or _STT = "XIX" Or _STT = "XX" Or _STT = "XXI" Or _STT = "XXII" Or _STT = "XXIII" Or _STT = "XXIV" Or _STT = "XXV" Or _STT = "XXVI" Or _STT = "XXVII" Or _STT = "XXVIII" Or _STT = "XXIX" Or _STT = "XXX" Or _STT = "XXXI" Or _STT = "XXXII" Or _STT = "XXIII" Or _STT = "XXXIV" Or _STT = "XXXV" Then
                            '        dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Bold)
                            '        dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Italic)
                            '        dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.Orange '.FromArgb(242, 249, 3) 'Color.FromArgb(97, 135, 214) '
                            '    Else
                            '        iCountRows = iCountRows + 1
                            '    End If
                            'End If
                        Next
                    End If
                End If
                lbl_waitting.Text = ""
                Cursor = Cursors.Default
                lbl_tonghop.Text = " Tổng số cán bộ: " & iCountRows.ToString()
            End If
        Catch ex As Exception
            lbl_waitting.Text = "Error ..."
            Cursor = Cursors.Default
            MessageBox.Show("Lỗi truy vấn Danh sách cán bộ: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_xuatexcel_Click(sender As Object, e As EventArgs) Handles btn_xuatexcel.Click

        If (db_dscanbo Is Nothing Or db_dscanbo.Rows.Count <= 0) Then
            MessageBox.Show("Không có danh sách cán bộ để xuất file excel (Hãy click nút lệnh 'Tra cứu' trước). Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        Cursor = Cursors.WaitCursor
        lbl_waitting.Text = "Waitting for data export..."

        'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Dim _ThoiDiem As DateTime = dtpk_thoidiem.Value
        Dim _IndexBC As Int32 = cb_loaicb.SelectedIndex
        'Dim _LoaiCB As Byte = IIf(_IdIndex = 1, 2, IIf(chk_TrunguongQL.Checked = True, 1, 0))

        Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
        Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
        Dim sPath_Export As String = "C:\Temp\"
        Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
        sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")

        sFileExcelTemplate = sPathExcelTemplate + IIf(_IndexBC = 1, "SK_DanhSachCanBo_BVTV_01.xlsx", "SK_DanhSachCanBo_01.xlsx")
        sFileNameEx = "Sao ke danh sach " + IIf(_IndexBC = 1, "Bao ve, Tap vu thoi diem_", IIf(chk_TrunguongQL.Checked = True, "Can bo TW Quan ly thoi diem_", "Can bo thoi diem_")) + _ThoiDiem.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
        Dim iRowStart As Integer = 0, iRowA As Integer = 0

        If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
            Using excelEngine As ExcelEngine = New ExcelEngine()
                Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                Dim worksheet As IWorksheet = workbook.Worksheets(0)
                iRowStart = 2
                sColNameEnd = IIf(_IndexBC = 1, "AH", "AT")

                'Create Template Marker Processor
                Dim marker As ITemplateMarkersProcessor = workbook.CreateTemplateMarkersProcessor()

                Dim columnNames(db_dscanbo.Columns.Count) As String
                Dim i As Integer = 0
                For Each column As DataColumn In db_dscanbo.Columns
                    columnNames(i) = column.ColumnName
                    i += 1
                Next

                For Each c As DataColumn In db_dscanbo.Columns
                    Dim columnName As String = c.ColumnName
                    Dim columnData As EnumerableRowCollection(Of Object)
                    If Globals.IsNumeric(c) Then
                        columnData = db_dscanbo.AsEnumerable().[Select](Function(r) If(r.Field(Of Object)(columnName), 0))
                    ElseIf c.DataType Is GetType(Date) Then
                        columnData = db_dscanbo.AsEnumerable.[Select](Function(r) If(r.Field(Of Object)(columnName), ""))
                    Else
                        columnData = db_dscanbo.AsEnumerable.[Select](Function(r) If(("'" & r.Field(Of Object)(columnName)), ""))
                    End If

                    Dim paramDataArray As Object() = columnData.ToArray
                    marker.AddVariable(columnName, paramDataArray)
                Next

                marker.ApplyMarkers()
                workbook.Version = ExcelVersion.Excel2013

                'Căn chỉnh bôi đậm, nghiêng các dòng bản ghi và điền dữ liệu vào một số CELL
                If Not (db_dscanbo Is Nothing) Then
                    If (db_dscanbo.Rows.Count > 0) Then
                        For iTT As Integer = 0 To db_dscanbo.Rows.Count - 1
                            iRowA = (iRowStart + 1 + iTT)
                            Dim iKieuIn As Integer = CType(db_dscanbo.Rows(iTT)("KieuIn"), Integer)
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

    ''' <summary>
    ''' Hàm thực hiện định nghĩa lại các cột của Lưới dữ liệu hiển thị kết quả
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FixGrid_Columns()
        For Each _column As DataGridViewColumn In dgv_main.Columns
            _column.ReadOnly = True
            _column.SortMode = DataGridViewColumnSortMode.NotSortable
            Select Case _column.HeaderText.Trim()
                Case "STT"
                    _column.HeaderText = "STT"
                    _column.Width = 40
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                    _column.Frozen = True

                Case "HoTen"
                    _column.HeaderText = "Đơn vị/Phòng ban/Họ tên"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                    _column.Frozen = True
                    
                Case "MaCB"
                    _column.HeaderText = "Mã CB"
                    _column.Width = 60
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NgaySinh"
                    _column.HeaderText = "Ngày sinh"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "GioiTinh"
                    _column.HeaderText = "Giới tính"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_So"
                    _column.HeaderText = "Số CMND/Thẻ căn cước"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_NgayCap"
                    _column.HeaderText = "Ngày cấp"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_NoiCap"
                    _column.HeaderText = "Nơi cấp"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NS_DiaChi"
                    _column.HeaderText = "Nơi sinh"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NQ_DiaChi"
                    _column.HeaderText = "Nguyên quán"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "TTr_DiaChi"
                    _column.HeaderText = "Thường trú/Tạm trú"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "ThT_DiaChi"
                    _column.HeaderText = "Địa chỉ theo hộ khẩu"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DanToc"
                    _column.HeaderText = "Dân tộc"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "TonGiao"
                    _column.HeaderText = "Tôn giáo"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "TrinhDoCT"
                    _column.HeaderText = "Trình độ chính trị"
                    _column.Width = 70
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DienThoai_DD"
                    _column.HeaderText = "Số điện thoại"
                    _column.Width = 90
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "Email"
                    _column.HeaderText = "Địa chỉ e-mail"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "QDNS_CV_PB_ChMon"
                    _column.HeaderText = "Chức danh/Chuyên môn đảm nhận"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "CVTruocTuyenDung"
                    _column.HeaderText = "Công việc trước khi Tuyển dụng/Tiếp nhận"
                    _column.Width = 120
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NgayTiepNhan"
                    _column.HeaderText = "Ngày tiếp nhận"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NgayTuyenDung"
                    _column.HeaderText = "Ngày tuyển dụng"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DVTuyenDung"
                    _column.HeaderText = "Đơn vị tuyển dụng"
                    _column.Width = 120
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "QDNS_Ngay_BN_BNL"
                    _column.HeaderText = "Ngày Bổ nhiệm/Bổ nhiệm lại"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DV_NgayCThuc"
                    _column.HeaderText = "Ngày vào đảng chính thức"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DV_SoThe"
                    _column.HeaderText = "Số thẻ đảng"
                    _column.Width = 70
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DV_NoiCapThe"
                    _column.HeaderText = "Nơi cấp thẻ đảng"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_TrinhDo"
                    _column.HeaderText = "Trình độ"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_CoSoDT"
                    _column.HeaderText = "Trường đào tạo"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_HeDT"
                    _column.HeaderText = "Hệ đào tạo"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_ChuyenNganh"
                    _column.HeaderText = "Chuyên ngành đào tạo"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_NamTN"
                    _column.HeaderText = "Năm tốt nghiệp"
                    _column.Width = 45
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "LCB_NgachLuong"
                    _column.HeaderText = "Ngạch lương"
                    _column.Width = 150
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "LCB_BacLuong"
                    _column.HeaderText = "Bậc lương"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "LCB_HeSoLuong"
                    _column.HeaderText = "Hệ số lương/Tiền lương"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "LCB_NgayHuong"
                    _column.HeaderText = "Ngày hưởng lương"
                    _column.Width = 77
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "LCB_NgayLenDK"
                    _column.HeaderText = "Ngày nâng lương dự kiến"
                    _column.Width = 77
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "PCCV"
                    _column.HeaderText = "Phụ cấp chức vụ"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCKV"
                    _column.HeaderText = "Phụ cấp khu vực"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCTN"
                    _column.HeaderText = "Phụ cấp trách nhiệm"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCDH"
                    _column.HeaderText = "Phụ cấp độc hại"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCTH"
                    _column.HeaderText = "Phụ cấp thu hút"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCDang"
                    _column.HeaderText = "Phụ cấp đảng"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "HDLD_LoaiHD"
                    _column.HeaderText = "Loại HĐLĐ"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "HDLD_So"
                    _column.HeaderText = "Số HĐLĐ"
                    _column.Width = 110
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "HDLD_NgayHL"
                    _column.HeaderText = "Ngày hiệu lực HĐLĐ"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HDLD_NgayKy"
                    _column.HeaderText = "Ngày ký HĐLĐ"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HDLD_DVKyHD"
                    _column.HeaderText = "Đơn vị ký HĐLĐ"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case Else
                    _column.Visible = False
            End Select
        Next
    End Sub

End Class