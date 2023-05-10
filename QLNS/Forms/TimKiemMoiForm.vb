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

Public Class TimKiemMoiForm
    Private _ChiNhanhBLL As ChiNhanhBLL = New ChiNhanhBLL()
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo

    Private ARL_DonVi As ArrayList = New ArrayList
    Private ARL_PhongBan As ArrayList = New ArrayList
    Private ARL_ChucVu As ArrayList = New ArrayList
    Private ARL_DanToc As ArrayList = New ArrayList
    Private ARL_TonGiao As ArrayList = New ArrayList
    Private db_dscanbo As System.Data.DataTable = New System.Data.DataTable()

#Region "---> Events <---"
    Private Sub TimKiemMoiForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dNgayHienTai As DateTime = Globals.GetDateTime_ForServerDB()
        dtpk_thoidiem.Text = dNgayHienTai
        Dim sSQL As String = ""
        sSQL = "Select Id As Value,(Case When (Id_Goc > 5) Then ('  + ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display"
        sSQL = sSQL & " From ("
        sSQL = sSQL & "        Select Id,Id_Goc,Ma_So,Ten_Goi As Ten_Goi,Status,(Case When A.Id_Goc In (0,1) Then A.Ma_So Else (Select Distinct X.Ma_So From ChiNhanh X Where X.Id=A.Id_Goc And Status=1) End) MaCN From ChiNhanh A Where A.Status=1"
        If TRUCTHUOC = 0 Then   'Chạy tại Chi nhánh
            sSQL = sSQL & "      And Ma_So = '" & DONVI.Trim() & "' Or Id_Goc In (Select Id From ChiNhanh Where Ma_So = '" & DONVI.Trim() & "' And Status = 1) "
        End If
        sSQL = sSQL & "      ) As Z "
        sSQL = sSQL & "      Order By MaCN,(Case When MaCN = Ma_So Then 0 Else 1 End),Ma_So"

        ARL_DonVi = _Globals.Bind_ComBoBox(cb_donvi, sSQL, "---Danh sách đơn vị---")
        If (cb_donvi.Items.Count <> 0) Then
            cb_donvi.SelectedIndex = 0
        End If
        cb_donvi_SelectedIndexChanged(sender, e)

        ckb_toandonvi.Checked = False
        edt_macb.Text = ""
        edt_hoten.Text = ""
        ckb_gioitinh_nam.Checked = False
        ckb_gioitinh_nu.Checked = False
        ckb_ngaysinh.Checked = False
        dtpk_ngaysinh_bd.Text = dNgayHienTai.AddYears(-50)
        dtpk_ngaysinh_kt.Text = dNgayHienTai
        ckb_ngayvaonhcs.Checked = False
        dtpk_ngayvaonhcs_bd.Text = dNgayHienTai.AddYears(-20)
        dtpk_ngayvaonhcs_kt.Text = dNgayHienTai
        ckb_nghihuu.Checked = False
        dtpk_ngaytinhnghihuu.Text = dNgayHienTai
        ckb_chualamsobhxh.Checked = False

        ckb_phongban.Checked = False
        ckb_chucvu.Checked = False
        ckb_dantoc.Checked = False
        ckb_tongiao.Checked = False

        ckb_bonhiemlai.Checked = False
        dtpk_bonhiemlai_tungay.Text = DateTimeUtil.GetStartOfMonth(dNgayHienTai.Month, dNgayHienTai.Year)
        dtpk_bonhiemlai_denngay.Text = DateTimeUtil.GetEndOfMonth(dNgayHienTai.Month, dNgayHienTai.Year)

        ARL_ChucVu.Clear()
        cb_chucvu.Items.Clear()
        ARL_ChucVu = _Globals.Bind_ComBoBox(cb_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ người ký---")
        If (cb_chucvu.Items.Count <> 0) Then
            cb_chucvu.SelectedIndex = 0
        End If
        'Bind dữ liệu vào combobox - Dân tộc
        ARL_DanToc.Clear()
        cb_dantoc.Items.Clear()
        ARL_DanToc = _Globals.Bind_ComBoBox(cb_dantoc, clsHT_DanhMuc.Sql_Dantoc, "---Dân tộc---")
        If (cb_dantoc.Items.Count <> 0) Then
            cb_dantoc.SelectedIndex = 0
        End If
        'Bind dữ liệu vào combobox - Tôn giáo
        ARL_TonGiao.Clear()
        cb_tongiao.Items.Clear()
        ARL_TonGiao = _Globals.Bind_ComBoBox(cb_tongiao, clsHT_DanhMuc.Sql_Tongiao, "---Tôn giáo---")
        If (cb_tongiao.Items.Count <> 0) Then
            cb_tongiao.SelectedIndex = 0
        End If
    End Sub

    Private Sub cb_donvi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_donvi.SelectedIndexChanged
        Try
            ARL_PhongBan.Clear()
            cb_phongban.Items.Clear()
            If (cb_donvi.SelectedIndex > 0 And cb_donvi.Items.Count <> 0) Then
                Dim _DonviId As Int32 = CType(ARL_DonVi(IIf(cb_donvi.SelectedIndex > 0, cb_donvi.SelectedIndex, "0")), Int32)
                'Lấy mã chi nhánh từ Id chi nhánh
                Dim strSQL As String = ""
                strSQL = String.Format("Select * from ChiNhanh Where id = {0} and Status = 1", _DonviId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim tructhuocCurrent As String = _HS_CanBo.GetTrucThuoc(db.Rows(0)("ma_so").ToString().Trim())
                            strSQL = "Select Id,Ten_Phong As Ten_Goi From PhongBan Where Charindex('" & tructhuocCurrent.ToString & "',Truc_Thuoc) > 0 And Status = 1 Order By Ma_so"
                            ARL_PhongBan = _Globals.Bind_ComBoBox(cb_phongban, strSQL, "---Phòng ban---")
                            If (cb_phongban.Items.Count <> 0) Then
                                cb_phongban.SelectedIndex = 0
                            End If
                        End If
                    End If
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show("Điền dữ liệu danh sách Phòng ban: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_huybo_Click(sender As Object, e As EventArgs) Handles btn_huybo.Click
        Dim dNgayHienTai As DateTime = Globals.GetDateTime_ForServerDB()
        dtpk_thoidiem.Text = dNgayHienTai
        If (cb_donvi.Items.Count <> 0) Then
            cb_donvi.SelectedIndex = 0
        End If
        cb_donvi_SelectedIndexChanged(sender, e)

        ckb_toandonvi.Checked = False
        edt_macb.Text = ""
        edt_hoten.Text = ""
        ckb_gioitinh_nam.Checked = False
        ckb_gioitinh_nu.Checked = False
        ckb_ngaysinh.Checked = False
        dtpk_ngaysinh_bd.Text = dNgayHienTai.AddYears(-50)
        dtpk_ngaysinh_kt.Text = dNgayHienTai
        ckb_ngayvaonhcs.Checked = False
        dtpk_ngayvaonhcs_bd.Text = dNgayHienTai.AddYears(-20)
        dtpk_ngayvaonhcs_kt.Text = dNgayHienTai
        ckb_nghihuu.Checked = False
        dtpk_ngaytinhnghihuu.Text = dNgayHienTai
        ckb_chualamsobhxh.Checked = False

        ckb_phongban.Checked = False
        ckb_chucvu.Checked = False
        ckb_dantoc.Checked = False
        ckb_tongiao.Checked = False

        ckb_bonhiemlai.Checked = False
        dtpk_bonhiemlai_tungay.Text = DateTimeUtil.GetStartOfMonth(dNgayHienTai.Month, dNgayHienTai.Year)
        dtpk_bonhiemlai_denngay.Text = DateTimeUtil.GetEndOfMonth(dNgayHienTai.Month, dNgayHienTai.Year)

        If (cb_chucvu.Items.Count <> 0) Then
            cb_chucvu.SelectedIndex = 0
        End If
        If (cb_dantoc.Items.Count <> 0) Then
            cb_dantoc.SelectedIndex = 0
        End If
        If (cb_tongiao.Items.Count <> 0) Then
            cb_tongiao.SelectedIndex = 0
        End If
    End Sub

    Private Sub ckb_phongban_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_phongban.CheckedChanged
        If ckb_phongban.Checked Then
            cb_phongban.Enabled = True
        Else
            cb_phongban.Enabled = False
        End If
    End Sub

    Private Sub ckb_chucvu_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_chucvu.CheckedChanged
        If ckb_chucvu.Checked Then
            cb_chucvu.Enabled = True
        Else
            cb_chucvu.Enabled = False
        End If
    End Sub

    Private Sub ckb_dantoc_CheckedChanged(sender As Object, e As EventArgs)
        If ckb_dantoc.Checked Then
            cb_dantoc.Enabled = True
        Else
            cb_dantoc.Enabled = False
        End If
    End Sub

    Private Sub ckb_tongiao_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_tongiao.CheckedChanged
        If ckb_tongiao.Checked Then
            cb_tongiao.Enabled = True
        Else
            cb_tongiao.Enabled = False
        End If
    End Sub

    Private Sub ckb_ngaysinh_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_ngaysinh.CheckedChanged
        If ckb_ngaysinh.Checked Then
            dtpk_ngaysinh_bd.Enabled = True
            dtpk_ngaysinh_kt.Enabled = True
        Else
            dtpk_ngaysinh_bd.Enabled = False
            dtpk_ngaysinh_kt.Enabled = False
        End If
    End Sub

    Private Sub ckb_ngayvaonhcs_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_ngayvaonhcs.CheckedChanged
        If ckb_ngayvaonhcs.Checked Then
            dtpk_ngayvaonhcs_bd.Enabled = True
            dtpk_ngayvaonhcs_kt.Enabled = True
        Else
            dtpk_ngayvaonhcs_bd.Enabled = False
            dtpk_ngayvaonhcs_kt.Enabled = False
        End If
    End Sub

    Private Sub ckb_nghihuu_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_nghihuu.CheckedChanged
        If ckb_nghihuu.Checked Then
            dtpk_ngaytinhnghihuu.Enabled = True
        Else
            dtpk_ngaytinhnghihuu.Enabled = False
        End If
    End Sub

    Private Sub ckb_bonhiemlai_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_bonhiemlai.CheckedChanged
        If ckb_bonhiemlai.Checked Then
            dtpk_bonhiemlai_tungay.Enabled = True
            dtpk_bonhiemlai_denngay.Enabled = True
        Else
            dtpk_bonhiemlai_tungay.Enabled = False
            dtpk_bonhiemlai_denngay.Enabled = False
        End If
    End Sub

    Private Sub btn_tracuu_Click(sender As Object, e As EventArgs) Handles btn_tracuu.Click
        Try
            Cursor = Cursors.WaitCursor
            lbl_waitting.Text = "Waitting ..."

            'If ((cb_donvi Is Nothing) Or (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0)) Then
            '    MessageBox.Show("Bạn chưa chọn đơn vị cần tra cứu thông tin!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            '    ActiveControl = cb_donvi
            '    Return
            'End If

            Dim _DonVi_Id As Integer = CType(IIf(ARL_DonVi.Count > 0, ARL_DonVi(cb_donvi.SelectedIndex), "0"), Int32)
            Dim _ThoiDiem As String = dtpk_thoidiem.Value.ToString("dd/MM/yyyy") 'DateTimeUtil.getDatetime()
            Dim _IsALL As Byte = IIf(ckb_toandonvi.Checked = True, 1, 0)
            Dim _IdCanBo As String = ""
            Dim _MaCB As String = edt_macb.Text
            Dim _HoTen As String = edt_hoten.Text
            Dim _DonVi_Cd As String = ""
            Dim _DonVi_HT As String = ""
            Dim _PhongBan_Id As Integer = 0
            If ARL_PhongBan.Count <= 0 Then
                _PhongBan_Id = 0
            Else
                _PhongBan_Id = CType(IIf(ARL_PhongBan.Count > 0, ARL_PhongBan(cb_phongban.SelectedIndex), "0"), Int32)
            End If

            Dim _PhongBan_Cd As String = ""
            Dim _PhongBan_HT As String = ""
            Dim _ChucVu_Id As Integer = CType(IIf(ARL_ChucVu.Count > 0, ARL_ChucVu(cb_chucvu.SelectedIndex), "0"), Int32)
            Dim _ChucVu_Cd As String = ""
            Dim _ChucVu_HT As String = ""
            Dim _NgaySinh_BD As String = ""
            Dim _NgaySinh_KT As String = ""
            Dim _NgayBoNhiemLai_BD As String = ""
            Dim _NgayBoNhiemLai_KT As String = ""

            If ckb_ngaysinh.Checked Then
                _NgaySinh_BD = dtpk_ngaysinh_bd.Value.ToString("dd/MM/yyyy")
                _NgaySinh_KT = dtpk_ngaysinh_kt.Value.ToString("dd/MM/yyyy")
            End If
            Dim _GioiTinh_Cd As String = ""
            If (ckb_gioitinh_nam.Checked) Then
                _GioiTinh_Cd = "M"
            End If
            If (ckb_gioitinh_nu.Checked) Then
                _GioiTinh_Cd = "F"
            End If
            Dim _GioiTinh_HT As String = ""
            Dim _HonNhan_Cd As String = ""
            Dim _SoCMT As String = ""
            Dim _DanToc_IdTMP As Integer = CType(IIf(ARL_DanToc.Count > 0, ARL_DanToc(cb_dantoc.SelectedIndex), "0"), Int32)
            Dim _DanToc_Id As String = IIf(_DanToc_IdTMP <> 0, _DanToc_IdTMP.ToString(), "")
            Dim _TonGiao_IdTMP As Integer = CType(IIf(ARL_TonGiao.Count > 0, ARL_TonGiao(cb_tongiao.SelectedIndex), "0"), Int32)
            Dim _TonGiao_Id As String = IIf(_TonGiao_IdTMP <> 0, _TonGiao_IdTMP.ToString(), "")
            Dim _TWQuanLy As Byte = 0
            Dim _NghiHuu As Byte = 0
            If ckb_nghihuu.Checked Then
                If dtpk_ngaytinhnghihuu.Value.Year > dtpk_ngaytinhnghihuu_dennam.Value.Year Then
                    MessageBox.Show("Khoảng thời gian (năm) nghỉ hưu không hợp lệ, từ năm [" + dtpk_ngaytinhnghihuu.Value.ToString("yyyy") + "] lớn hơn đến năm [" + dtpk_ngaytinhnghihuu_dennam.Value.ToString("yyyy") + "]. Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    lbl_waitting.Text = ""
                    Cursor = Cursors.Default
                    Return
                End If
                _NghiHuu = 1
            End If
            Dim _NgayTinhNghiHuu As String = dtpk_ngaytinhnghihuu.Value.ToString("dd/MM/yyyy")
            Dim _NgayTinhNghiHuuDenNam As String = dtpk_ngaytinhnghihuu_dennam.Value.ToString("dd/MM/yyyy")
            Dim _SoBHXH As Byte = 0
            If ckb_chualamsobhxh.Checked Then
                _SoBHXH = 1
            End If
            Dim _NgayVaoNHCS_BD As String = ""
            Dim _NgayVaoNHCS_KT As String = ""
            If ckb_ngayvaonhcs.Checked Then
                _NgayVaoNHCS_BD = dtpk_ngayvaonhcs_bd.Value.ToString("dd/MM/yyyy")
                _NgayVaoNHCS_KT = dtpk_ngayvaonhcs_kt.Value.ToString("dd/MM/yyyy")
            End If

            If ckb_bonhiemlai.Checked Then
                _NgayBoNhiemLai_BD = dtpk_bonhiemlai_tungay.Value.ToString("dd/MM/yyyy")
                _NgayBoNhiemLai_KT = dtpk_bonhiemlai_denngay.Value.ToString("dd/MM/yyyy")
            End If

            Dim _FlagDL As Byte = 0
            lbl_tonghop.Text = ""
            dgv_main.AutoGenerateColumns = True
            dgv_main.Columns.Clear()
            dgv_main.ColumnHeadersHeight = 35
            dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            dgv_main.Columns().Clear()

            If Not (db_dscanbo Is Nothing) Then
                db_dscanbo.Reset()
            End If
            db_dscanbo = _HS_CanBo.GetHS_CanBo_Search(_ThoiDiem, _IdCanBo, _MaCB, _HoTen, _DonVi_Id, _DonVi_Cd, _DonVi_HT, _PhongBan_Id, _PhongBan_Cd, _PhongBan_HT, _ChucVu_Id, _ChucVu_Cd, _ChucVu_HT, _NgaySinh_BD, _NgaySinh_KT, _GioiTinh_Cd, _GioiTinh_HT, _HonNhan_Cd, _SoCMT, _DanToc_Id, _TonGiao_Id, _TWQuanLy, _NghiHuu, _NgayTinhNghiHuu, _NgayTinhNghiHuuDenNam, _SoBHXH, _NgayVaoNHCS_BD, _NgayVaoNHCS_KT, _NgayBoNhiemLai_BD, _NgayBoNhiemLai_KT, _IsALL, _FlagDL)
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


                    Next
                End If
            End If
            lbl_waitting.Text = ""
            Cursor = Cursors.Default
            lbl_tonghop.Text = " Kết quả tra cứu thời điểm " & _ThoiDiem & " là: " & iCountRows.ToString() & " bản ghi"
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

        Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
        Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
        Dim sPath_Export As String = "C:\Temp\"
        Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
        sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")

        sFileExcelTemplate = sPathExcelTemplate + "SK_DanhSachCanBo_02.xlsx"
        sFileNameEx = "Danh sach Can bo thoi diem_" + _ThoiDiem.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
        Dim iRowStart As Integer = 0, iRowA As Integer = 0

        If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
            Using excelEngine As ExcelEngine = New ExcelEngine()
                Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                Dim worksheet As IWorksheet = workbook.Worksheets(0)
                iRowStart = 2
                sColNameEnd = "AR"

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

    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        db_dscanbo.Dispose()
        Close()
    End Sub
#End Region

#Region "---> Functions <---"
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
                Case "MaCB"
                    _column.HeaderText = "Mã CB"
                    _column.Width = 60
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HoTen"
                    _column.HeaderText = "Họ tên"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                    _column.Frozen = True
                Case "GioiTinh_HT"
                    _column.HeaderText = "Giới tính"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NgaySinh_HT"
                    _column.HeaderText = "Ngày sinh"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NQ_DiaChi_HT"
                    _column.HeaderText = "Quê quán"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "TTr_DiaChi_HT"
                    _column.HeaderText = "Thường trú/Tạm trú"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "CMT_So"
                    _column.HeaderText = "Số CMND/Thẻ căn cước"
                    _column.Width = 95
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_NgayCap_HT"
                    _column.HeaderText = "Ngày cấp"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_NoiCap"
                    _column.HeaderText = "Nơi cấp"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "GhiChu_BS"
                    _column.HeaderText = "Đơn vị"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "Phong_Moi_HT"
                    _column.HeaderText = "Phòng ban"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "ChucVu_Moi_HT"
                    _column.HeaderText = "Chức vụ"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "LCB_HeSoLuong"
                    _column.HeaderText = "Hệ số lương"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "LCB_NgayHuongHT"
                    _column.HeaderText = "Ngày hưởng lương"
                    _column.Width = 77
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DanToc_HT"
                    _column.HeaderText = "Dân tộc"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "TonGiao_HT"
                    _column.HeaderText = "Tôn giáo"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "TrinhDoCT"
                    _column.HeaderText = "Trình độ chính trị"
                    _column.Width = 70
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "Ngay_NH_HT"
                    _column.HeaderText = "Ngày vào NH"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "Ngay_VBSP_HT"
                    _column.HeaderText = "Ngày vào NHCS"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DV_NgayCThuc_HT"
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
                Case "DienThoai_NR"
                    _column.HeaderText = "Điện nhà riêng"
                    _column.Width = 90
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DienThoai_DD"
                    _column.HeaderText = "Điện thoại di động"
                    _column.Width = 90
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "Email"
                    _column.HeaderText = "Địa chỉ e-mail"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NH_SoTK"
                    _column.HeaderText = "Số tài khoản"
                    _column.Width = 110
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NH_TenNH"
                    _column.HeaderText = "Ngân hàng"
                    _column.Width = 110
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "MaSoThue"
                    _column.HeaderText = "Mã số thuế"
                    _column.Width = 110
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "BHXH_SoSo"
                    _column.HeaderText = "Số sổ BHXH"
                    _column.Width = 80
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
                Case "HDLD_LoaiHD"
                    _column.HeaderText = "Loại HĐLĐ"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "HDLD_So"
                    _column.HeaderText = "Số HĐLĐ"
                    _column.Width = 110
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "HDLD_NgayHL_HT"
                    _column.HeaderText = "Ngày hiệu lực HĐLĐ"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HDLD_NgayKy_HT"
                    _column.HeaderText = "Ngày ký HĐLĐ"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HDLD_DVKyHD"
                    _column.HeaderText = "Đơn vị ký HĐLĐ"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NgayBoNhiemLai_HT"
                    _column.HeaderText = "Ngày bổ nhiệm lại"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NgayNghiHuuHT"
                    _column.HeaderText = "Ngày nghỉ hưu"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case Else
                    _column.Visible = False
            End Select
        Next
    End Sub

#End Region

End Class