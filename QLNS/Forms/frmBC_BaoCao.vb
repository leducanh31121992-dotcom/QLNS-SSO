Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports Syncfusion.XlsIO

Public Class frmBC_BaoCao

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    ''' <summary>
    ''' Biến lưu chỉ số xác định loại báo cáo
    ''' </summary>
    ''' <remarks></remarks>
    Public flagReport As Integer
    Private arrChiNhanh As ArrayList = New ArrayList()
    Private _Globals As Globals = New Globals()
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _HSCanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _Reports As clsHS_BaoCao = New clsHS_BaoCao()
    Private _HeThongBLL As HeThongBLL = New HeThongBLL()

    Private _Quarter() As String = {"--- Chọn quý ---", "Quý I", "Quý II", "Quý III", "Quý IV"}
    Private _Quarter1() As String = {"--- Chọn thời điểm báo cáo ---", "Quý I", "Quý II", "Quý III", "Quý IV", "Số lượt bổ nhiệm"}
    Private _Schedule() As String = {"--- Chọn kỳ báo cáo ---", "Kỳ 1 (thời điểm từ ngày 01/01 đến ngày 30/06)", "Kỳ 2 (thời điểm từ ngày 01/07 đến ngày 31/12)"}
#End Region

#Region "---> Functions and Events: Sự kiện chính <---"
    ''' <summary>
    ''' Hàm thực hiện reset controls: Khởi tạo giá trị ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls()
        If DONVI = gMaDonViTW Then
            lbl_giamdoc.Text = "Tổng giám đốc "
        Else
            lbl_giamdoc.Text = "Giám đốc "
        End If

        If cb_chinhanh.Items.Count > 0 Then cb_chinhanh.SelectedIndex = 0
        'cb_chinhanh_SelectedIndexChanged(Nothing, Nothing)
        If cb_kybc.Items.Count > 0 Then cb_kybc.SelectedIndex = 0
        Dim db_sysvar As DataTable = New DataTable()
        edt_giamdoc.Text = GIAMDOC
        edt_truongphong.Text = TRUONGHCTC
        db_sysvar = _HeThongBLL.GetListSysVarSearch(0, 0, DONVI, "", "", "", 1)
        If Not (db_sysvar Is Nothing) Then
            If (db_sysvar.Rows.Count > 0) Then
                edt_giamdoc.Text = db_sysvar.Rows(0)("GiamDoc").ToString()
                edt_truongphong.Text = db_sysvar.Rows(0)("TruongHCTC").ToString()
            End If
        End If
        edt_nguoilap.Text = MainForm.lblNguoiSuDung.Text
        edt_so_bc.Text = ""
        dtpk_ngaylap_bc.Text = DateTime.Now.ToShortDateString()

        If (flagReport <> 3) Then
            dtpk_nam_bc.Text = DateTime.Now.ToShortDateString()
        Else
            dtpk_nam_bc.Value = New DateTime(DateTime.Now.Year, 12, 31)
        End If
    End Sub

    Private Sub ShowHide_ExportExcel()
        'btn_exportexcel.Visible = False
        'Dim _ProvinceId As Integer = CType(IIf(arrChiNhanh.Count > 0, arrChiNhanh(cb_chinhanh.SelectedIndex), "0"), Integer)
        'If (flagReport = 2) Then ' And (_ProvinceId = 0)
        '    btn_exportexcel.Visible = True
        'End If
    End Sub

    ''' <summary>
    ''' Hàm kiểm tra dữ liệu hợp lệ khi báo cáo
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        If TRUCTHUOC <> 1 Then
            If (cb_chinhanh.SelectedIndex <= 0 And cb_chinhanh.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn chi nhánh cần báo cáo!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_chinhanh
                Return False
            End If
        End If
        'If (flagReport = 1 Or flagReport = 2) Then
        '    If (edt_so_bc.Text.Trim() = "") Then
        '        MessageBox.Show("Số báo cáo không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
        '        ActiveControl = edt_so_bc
        '        Return False
        '    End If
        'End If
        If (edt_nguoilap.Text.Trim() = "") Then
            MessageBox.Show("Người lập báo cáo không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_nguoilap
            Return False
        End If
        If (edt_truongphong.Text.Trim() = "") Then
            MessageBox.Show("Trưởng phòng hành chính tổ chức không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_truongphong
            Return False
        End If
        If (edt_giamdoc.Text.Trim() = "") Then
            MessageBox.Show("Giám đốc không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_giamdoc
            Return False
        End If
        If (flagReport = 4 Or flagReport = 2) Then
            If (cb_kybc.SelectedIndex <= 0 And cb_kybc.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn kỳ báo cáo!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_kybc
                Return False
            End If
            'ElseIf (flagReport = 1 Or flagReport = 2) Then
        ElseIf (flagReport = 1 Or flagReport = 11) Then
            If (cb_kybc.SelectedIndex <= 0 And cb_kybc.Items.Count <> 0) Then
                MessageBox.Show("Bạn chưa chọn quý báo cáo!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = cb_kybc
                Return False
            End If
        End If
        Return True
    End Function

    ''' <summary>
    ''' Hàm thiết lập trạng thái controls đối với các trường hợp báo cáo cụ thể
    ''' </summary>
    ''' <param name="_status">Chỉ số xác định các báo cáo</param>
    ''' <remarks></remarks>
    Private Sub SetStatus_Controls(ByVal _status As Byte)
        labTuThang.Visible = False
        numTuThang.Visible = False
        labDenThang.Visible = False
        numDenThang.Visible = False
        Select Case _status
            Case 1  'Nếu chọn báo cáo 1
                Me.Text = "Báo cáo 01: TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI HOẠT ĐỘNG"
                lbl_kybc.Text = "Quý "
                ' cb_kybc.Width = 197
                cb_kybc.Items.Clear()
                For i As Integer = 0 To _Quarter.Length - 1
                    cb_kybc.Items.Add(_Quarter(i).ToString())
                Next
            Case 2  'Nếu chọn báo cáo 2
                Me.Text = "Báo cáo 02: TÌNH HÌNH THỰC HIỆN CÔNG TÁC BỔ NHIỆM CÁN BỘ LÃNH ĐẠO, QUẢN LÝ"
                'Thay combo thành Quý báo cáo
                lbl_kybc.Text = "Quý "
                ' cb_kybc.Width = 197
                cb_kybc.Items.Clear()
                For i As Integer = 0 To _Quarter1.Length - 1
                    cb_kybc.Items.Add(_Quarter1(i).ToString())
                Next
            Case 3  'Nếu chọn báo cáo 3
                Me.Text = "Báo cáo 03: BÁO CÁO THỐNG KÊ SỐ LƯỢNG & CHẤT LƯỢNG CÁN BỘ"
                cb_kybc.Visible = False
                lbl_kybc.Visible = False
                lbl_sobc.Visible = False
                edt_so_bc.Visible = False
                'cb_chinhanh.Width = 221
                'lbl_giamdoc.Width = 129
                'edt_giamdoc.Width = 221
                lbl_nambc.Text = "Thời điểm báo cáo "
                lbl_nambc.Width = 112
                dtpk_nam_bc.Width = 100
                dtpk_nam_bc.CustomFormat = "dd/MM/yyyy"
                dtpk_nam_bc.ShowUpDown = False
                'pnl_2.Controls.Add(lbl_truongphong)
                'pnl_2.Controls.Add(edt_truongphong)
                'lbl_truongphong.BringToFront()
                'edt_truongphong.BringToFront()
                'lbl_truongphong.Width = 158
                'edt_truongphong.Width = 275
                'lbl_nguoilap.Width = 129
                'edt_nguoilap.Width = 221
                lbl_ngaylap.Width = 112
                dtpk_ngaylap_bc.Width = 100
            Case 4  'Nếu chọn báo cáo 4
                Me.Text = "Báo cáo 04: BÁO CÁO DANH SÁCH TUYỂN DỤNG & TIẾP NHẬN CÁN BỘ"
                lbl_kybc.Text = "Kỳ báo cáo "
                cb_kybc.Items.Clear()
                For i As Integer = 0 To _Schedule.Length - 1
                    cb_kybc.Items.Add(_Schedule(i).ToString())
                Next
                'cb_chinhanh.Width = 290
                lbl_nambc.Width = 150
                dtpk_nam_bc.Width = 95
                'cb_kybc.Width = 302
                'edt_giamdoc.Width = 201
                edt_so_bc.Width = 101
            Case 11  'Nếu chọn báo cáo 11 -Bao cao 01 ngoai nganh
                Me.Text = "Báo cáo 01: TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI HOẠT ĐỘNG"
                lbl_kybc.Text = "Quý "
                cb_kybc.Items.Clear()
                For i As Integer = 0 To _Quarter.Length - 1
                    cb_kybc.Items.Add(_Quarter(i).ToString())
                Next
        End Select
    End Sub

    Private Sub frmBC_ShowReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStatus_Controls(flagReport)
        Dim strSQL As String = ""
        If TRUCTHUOC = 1 Then
            strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And id_goc IN (0,1)"
            arrChiNhanh.Clear()
            cb_chinhanh.Items.Clear()
            arrChiNhanh = _Globals.Bind_ComBoBox(cb_chinhanh, strSQL, "--- Chọn Toàn quốc ---")
        Else
            strSQL = "Select id,ten_goi From ChiNhanh Where Status = 1 And ma_so = '" & DONVI.Trim & "' and id_goc <= 1"
            arrChiNhanh.Clear()
            cb_chinhanh.Items.Clear()
            arrChiNhanh = _Globals.Bind_ComBoBox(cb_chinhanh, strSQL, "--- Chọn chi nhánh ---")
        End If

        ResetAll_Controls()
        Dim _CodeCN As String = DONVI
        If (_CodeCN <> "") Then
            strSQL = String.Format("Select * from ChiNhanh Where ma_so = '{0}' And Status = 1 And id_goc <= 1", _CodeCN)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        cb_chinhanh.SelectedIndex = IIf(db.Rows(0)("id").ToString() <> "", CType(arrChiNhanh.IndexOf(db.Rows(0)("id").ToString()), Integer), 0)
                    End If
                End If
            End Using
        End If
        cb_chinhanh_SelectedIndexChanged(Nothing, Nothing)
        btn_reset.Visible = False
        'Me.ReportViewer1.RefreshReport()
    End Sub

    Private Sub btn_view_report_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_view_report.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            If (IsValid()) Then
                Dim _IdChiNhanh As Integer = 0
                Dim strSQL As String = ""
                _IdChiNhanh = CType(IIf(arrChiNhanh.Count > 0, arrChiNhanh(cb_chinhanh.SelectedIndex), "0"), Integer)
                'Lấy tên và mã chi nhánh từ Id chi nhánh
                Dim _Diadiem As String = ""
                Dim _Name As String = ""
                Dim _Code As String = ""
                If _IdChiNhanh = 0 Or _IdChiNhanh = 1 Or DONVI = "000100" Then
                    _Name = "BAN TỔ CHỨC CÁN BỘ"
                    _Code = gMaDonViTW
                    _Diadiem = "Hà Nội"
                Else
                    strSQL = String.Format("Select * from ChiNhanh Where Id = {0} And Status = 1", _IdChiNhanh)
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If _db.Rows.Count > 0 Then
                                _Name = _db.Rows(0)("ten_goi").ToString().Trim()
                                _Code = _db.Rows(0)("ma_so").ToString().Trim()
                            End If
                        End If
                    End Using
                    'Lấy dữ liệu phần địa danh
                    _Diadiem = IIf(DIABAN Is Nothing, "", DIABAN)
                End If
                Dim sSobc_HCTC As String
                Dim sLabHCTC As String
                Dim sGiamDoc_Label As String
                sGiamDoc_Label = IIf(DONVI = gMaDonViTW, "Tổng Giám đốc", "Giám đốc")
                sSobc_HCTC = IIf(DONVI = gMaDonViTW, "NHCS-TCCB", IIf(DONVI = "000196", "CNTT-TH", IIf(DONVI = "000197" Or DONVI = "000101", "NHCS-HCNS", "NHCS-HCTC")))
                sLabHCTC = IIf(DONVI = gMaDonViTW, "Giám đốc Ban TCCB", IIf(DONVI = "000196", "Trưởng phòng Tổng hợp", IIf(DONVI = "000197" Or DONVI = "000101", "Trưởng phòng HC-NS", "Trưởng phòng HC-TC")))

                'Lấy bảng dữ liệu cần xuất ra báo cáo
                Dim db As DataTable = New DataTable()
                Select Case flagReport
                    Case 1              'Báo cáo: Tình hình thực hiện lao động - Mạng lưới hoạt động
                        'Lấy các tháng trong quý
                        Dim _ThangT1 As Integer = 0 'Tháng thứ nhất trong quý
                        Dim _ThangT2 As Integer = 0 'Tháng thứ hai trong quý
                        Dim _ThangT3 As Integer = 0 'Tháng thứ 3 trong quý
                        _ThangT1 = CByte(cb_kybc.SelectedIndex) * 3 - 2
                        _ThangT2 = _ThangT1 + 1
                        _ThangT3 = _ThangT1 + 2
                        If _IdChiNhanh = 0 And _Code = gMaDonViTW Then
                            db = _Reports.GetAll_Report01(_IdChiNhanh, dtpk_nam_bc.Value.Year, _ThangT1, _ThangT2, _ThangT3, 0)
                        Else
                            db = _Reports.GetAll_Report01(_IdChiNhanh, dtpk_nam_bc.Value.Year, _ThangT1, _ThangT2, _ThangT3, 1)
                        End If

                        db.TableName = "BC_01"
                        rdt_reportBC1.SetDataSource(db)
                        'Thực hiện truyền các tham số cho báo cáo
                        rdt_reportBC1.SetParameterValue("ten_cn", _Name)
                        rdt_reportBC1.SetParameterValue("so_bc", IIf(edt_so_bc.Text.Trim() = "", "         ", edt_so_bc.Text.Trim()))
                        rdt_reportBC1.SetParameterValue("nam_bc", dtpk_nam_bc.Value.Year)
                        rdt_reportBC1.SetParameterValue("quy_bc", cb_kybc.SelectedItem.ToString())
                        rdt_reportBC1.SetParameterValue("_TT1", _ThangT1)
                        rdt_reportBC1.SetParameterValue("_TT2", _ThangT2)
                        rdt_reportBC1.SetParameterValue("_TT3", _ThangT3)
                        rdt_reportBC1.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rdt_reportBC1.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rdt_reportBC1.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rdt_reportBC1.SetParameterValue("pr_tinh", _Diadiem)
                        'rdt_reportBC1.SetParameterValue("giamdoc_label", lbl_giamdoc.Text.Trim())
                        rdt_reportBC1.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                        rdt_reportBC1.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                        rdt_reportBC1.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))

                        rdt_reportBC1.SetParameterValue("giamdoc_label", sGiamDoc_Label)
                        rdt_reportBC1.SetParameterValue("labHCTC", sLabHCTC)
                        rdt_reportBC1.SetParameterValue("sobc_hctc", sSobc_HCTC)

                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        'crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rdt_reportBC1
                    Case 2              'Báo cáo: Tình hình thực hiện công tác bổ nhiệm cán bộ lãnh đạo, Quản lý
                        InBC_02_TCCB(_IdChiNhanh, _Code, _Name, _Diadiem, sGiamDoc_Label, sLabHCTC, sSobc_HCTC)
                    Case 3              'Báo cáo: Thống kê số lượng và Chất lượng cán bộ
                        db = _Reports.GetAll_Report03(_IdChiNhanh, 1, dtpk_nam_bc.Value)
                        db.TableName = "BC_03"
                        rdt_reportBC3.SetDataSource(db)
                        'Thực hiện truyền các tham số cho báo cáo
                        rdt_reportBC3.SetParameterValue("ten_cn", _Name)
                        rdt_reportBC3.SetParameterValue("nam_bc", dtpk_nam_bc.Value.ToString("dd/MM/yyyy"))
                        rdt_reportBC3.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rdt_reportBC3.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rdt_reportBC3.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rdt_reportBC3.SetParameterValue("pr_tinh", _Diadiem)
                        rdt_reportBC3.SetParameterValue("giamdoc_label", lbl_giamdoc.Text.Trim())
                        rdt_reportBC3.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                        rdt_reportBC3.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                        rdt_reportBC3.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
                        rdt_reportBC3.SetParameterValue("HCTC_lab", sLabHCTC)

                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rdt_reportBC3
                    Case 4              'Báo cáo: Danh sách tuyển dụng, tiếp nhận cán bộ
                        If _IdChiNhanh = 0 And _Code = gMaDonViTW Then
                            db = _Reports.GetAll_Report04(_IdChiNhanh, dtpk_nam_bc.Value.Year, CByte(cb_kybc.SelectedIndex), 0)
                        Else
                            db = _Reports.GetAll_Report04(_IdChiNhanh, dtpk_nam_bc.Value.Year, CByte(cb_kybc.SelectedIndex), 1)
                        End If
                        db.TableName = "BC_04"
                        rdt_reportBC4.SetDataSource(db)
                        'Thực hiện truyền các tham số cho báo cáo
                        rdt_reportBC4.SetParameterValue("ten_cn", _Name)
                        If _IdChiNhanh = 0 And _Code = gMaDonViTW Then
                            rdt_reportBC4.SetParameterValue("c2_header", "Chi nhánh")
                        Else
                            rdt_reportBC4.SetParameterValue("c2_header", "Họ và tên")
                        End If
                        If cb_kybc.SelectedIndex = 1 Then
                            rdt_reportBC4.SetParameterValue("thoidiem_bc", "30/06")
                        Else
                            rdt_reportBC4.SetParameterValue("thoidiem_bc", "31/12")
                        End If
                        rdt_reportBC4.SetParameterValue("nam_bc", dtpk_nam_bc.Value.Year)
                        rdt_reportBC4.SetParameterValue("giamdoc_label", lbl_giamdoc.Text.Trim())
                        rdt_reportBC4.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
                        rdt_reportBC4.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rdt_reportBC4.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rdt_reportBC4.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rdt_reportBC4.SetParameterValue("pr_tinh", _Diadiem)
                        Dim _day As String = dtpk_ngaylap_bc.Value.Day.ToString()
                        If (_day.Length < 2) Then _day = "0" + _day
                        Dim _month As String = dtpk_ngaylap_bc.Value.Month.ToString()
                        If (_month.Length < 2) Then _month = "0" + _month
                        rdt_reportBC4.SetParameterValue("pr_ngay", _day)
                        rdt_reportBC4.SetParameterValue("pr_thang", _month)
                        rdt_reportBC4.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.Year)

                        rdt_reportBC4.SetParameterValue("HCTC_lab", sLabHCTC)
                        rdt_reportBC4.SetParameterValue("sobc_hctc", sSobc_HCTC)

                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rdt_reportBC4
                    Case 11              'Báo cáo Ngoai ngành: Tình hình thực hiện lao động - Mạng lưới hoạt động 
                        'Lấy các tháng trong quý
                        Dim _ThangT1 As Integer = 0 'Tháng thứ nhất trong quý
                        Dim _ThangT2 As Integer = 0 'Tháng thứ hai trong quý
                        Dim _ThangT3 As Integer = 0 'Tháng thứ 3 trong quý
                        _ThangT1 = CByte(cb_kybc.SelectedIndex) * 3 - 2
                        _ThangT2 = _ThangT1 + 1
                        _ThangT3 = _ThangT1 + 2
                        If _IdChiNhanh = 0 And _Code = gMaDonViTW Then
                            db = _Reports.GetAll_Report01_NN(_IdChiNhanh, dtpk_nam_bc.Value.Year, _ThangT1, _ThangT2, _ThangT3, 0)
                        Else
                            db = _Reports.GetAll_Report01_NN(_IdChiNhanh, dtpk_nam_bc.Value.Year, _ThangT1, _ThangT2, _ThangT3, 1)
                        End If

                        db.TableName = "BCNN_01"
                        rdt_reportBC1.SetDataSource(db)
                        'Thực hiện truyền các tham số cho báo cáo
                        rdt_reportBC1.SetParameterValue("ten_cn", _Name)
                        rdt_reportBC1.SetParameterValue("so_bc", IIf(edt_so_bc.Text.Trim() = "", "         ", edt_so_bc.Text.Trim()))
                        rdt_reportBC1.SetParameterValue("nam_bc", dtpk_nam_bc.Value.Year)
                        rdt_reportBC1.SetParameterValue("quy_bc", cb_kybc.SelectedItem.ToString())
                        rdt_reportBC1.SetParameterValue("_TT1", _ThangT1)
                        rdt_reportBC1.SetParameterValue("_TT2", _ThangT2)
                        rdt_reportBC1.SetParameterValue("_TT3", _ThangT3)
                        rdt_reportBC1.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rdt_reportBC1.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rdt_reportBC1.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rdt_reportBC1.SetParameterValue("pr_tinh", _Diadiem)
                        rdt_reportBC1.SetParameterValue("giamdoc_label", lbl_giamdoc.Text.Trim())
                        rdt_reportBC1.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                        rdt_reportBC1.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                        rdt_reportBC1.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rdt_reportBC1
                End Select
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("Xem báo cáo: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btn_close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_close.Click
        Close()
    End Sub

    Private Sub btn_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_reset.Click
        ResetAll_Controls()
    End Sub

    'Private Sub cb_chinhanh_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_chinhanh.SelectedIndexChanged
    '    If (cb_chinhanh.Items.Count <> 0) Then
    '        Dim _IdChiNhanh As Integer = 0
    '        If cb_chinhanh.SelectedIndex > 0 Then
    '            _IdChiNhanh = CType(arrChiNhanh(IIf(cb_chinhanh.SelectedIndex > 0, cb_chinhanh.SelectedIndex, "0")), Integer)
    '        Else
    '            _IdChiNhanh = 0
    '        End If
    '        If (_IdChiNhanh > 0) Then
    '            Dim strSQL As String = String.Format("Select * from ChiNhanh Where id = {0} And Status = 1", _IdChiNhanh)
    '            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
    '                If Not (db Is Nothing) Then
    '                    If (db.Rows.Count > 0) Then
    '                        If (db.Rows(0)("ma_so").ToString().Trim() = gMaDonViTW) Then
    '                            lbl_giamdoc.Text = "Tổng giám đốc "
    '                        Else
    '                            lbl_giamdoc.Text = "Giám đốc "
    '                        End If
    '                        edt_giamdoc.Text = GIAMDOC
    '                        edt_truongphong.Text = _HSCanBo.GetVarNam("TRUONGHC-TC")
    '                    End If
    '                End If
    '            End Using
    '        ElseIf (_IdChiNhanh = 0 And TRUCTHUOC = 1) Then
    '            lbl_giamdoc.Text = "Tổng giám đốc "
    '            edt_giamdoc.Text = GIAMDOC
    '            edt_truongphong.Text = _HSCanBo.GetVarNam("TRUONGHC-TC")
    '        End If
    '    End If
    'End Sub
#End Region

#Region "---> Events: Sự kiện ngoại lệ người dùng <---"
    Private Sub cb_chinhanh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_chinhanh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (flagReport <> 3) Then
                edt_so_bc.Focus()
            Else
                dtpk_nam_bc.Focus()
            End If
        End If
    End Sub

    Private Sub edt_so_bc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_so_bc.Text = edt_so_bc.Text.Trim().ToUpper()
            dtpk_nam_bc.Focus()
        End If
    End Sub

    Private Sub edt_so_bc_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Up) Then
            cb_chinhanh.Focus()
        End If
    End Sub

    Private Sub dtpk_nam_bc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (flagReport <> 3) Then
                cb_kybc.Focus()
            Else
                edt_giamdoc.Focus()
            End If
        End If
    End Sub

    Private Sub cb_kybc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_kybc.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_giamdoc.Focus()
        End If
    End Sub

    Private Sub edt_giamdoc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_giamdoc.Text = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_giamdoc.Text.Trim()))
            edt_truongphong.Focus()
        End If
    End Sub

    Private Sub edt_giamdoc_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Up) Then
            If (flagReport <> 3) Then
                cb_kybc.Focus()
            Else
                dtpk_nam_bc.Focus()
            End If
        End If
    End Sub

    Private Sub edt_truongphong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_truongphong.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_truongphong.Text = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_truongphong.Text.Trim()))
            edt_nguoilap.Focus()
        End If
    End Sub

    Private Sub edt_truongphong_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_truongphong.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_giamdoc.Focus()
        End If
    End Sub

    Private Sub edt_nguoilap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nguoilap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_nguoilap.Text = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_nguoilap.Text.Trim()))
            dtpk_ngaylap_bc.Focus()
        End If
    End Sub

    Private Sub edt_nguoilap_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_nguoilap.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_truongphong.Focus()
        End If
    End Sub

    Private Sub dtpk_ngaylap_bc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            btn_view_report.Focus()
        End If
    End Sub

    Private Sub edt_giamdoc_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        edt_giamdoc.Text = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_giamdoc.Text.Trim()))
    End Sub

    Private Sub edt_truongphong_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_truongphong.Leave
        edt_truongphong.Text = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_truongphong.Text.Trim()))
    End Sub

    Private Sub edt_nguoilap_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_nguoilap.Leave
        edt_nguoilap.Text = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_nguoilap.Text.Trim()))
    End Sub
#End Region

    Private Sub cb_kybc_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_kybc.SelectedValueChanged
        If flagReport = 2 Then
            If CByte(cb_kybc.SelectedIndex) = 5 Then
                labTuThang.Visible = True
                numTuThang.Visible = True
                labDenThang.Visible = True
                numDenThang.Visible = True
            Else
                labTuThang.Visible = False
                numTuThang.Visible = False
                labDenThang.Visible = False
                numDenThang.Visible = False
            End If
        End If
    End Sub

    Private Sub btn_exportexcel_Click(sender As Object, e As EventArgs) Handles btn_exportexcel.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim _IdChiNhanh As Integer = 0
            Dim strSQL As String = ""
            _IdChiNhanh = CType(IIf(arrChiNhanh.Count > 0, arrChiNhanh(cb_chinhanh.SelectedIndex), "0"), Integer)
            Dim db As DataTable = New DataTable()
            Dim db_begin As DateTime
            Dim db_end As DateTime
            'Lấy ngày bắt đầu báo cáo
            db_begin = clsHS_BaoCao.GetDate(CByte(cb_kybc.SelectedIndex), dtpk_nam_bc.Value.Year, True)
            'Lấy ngày kết thúc báo cáo
            db_end = clsHS_BaoCao.GetDate(CByte(cb_kybc.SelectedIndex), dtpk_nam_bc.Value.Year, False)
            db = _Reports.GetAll_Report02(_IdChiNhanh, db_begin, db_end, 0, 1, 2)

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("Xem báo cáo: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cb_chinhanh_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_chinhanh.SelectedIndexChanged
        ckb_ChiTietPGD.Visible = False
        ShowHide_ExportExcel()
        If flagReport = 2 Then
            If Not (cb_chinhanh Is Nothing) Then
                If cb_chinhanh.Items.Count > 0 And cb_chinhanh.SelectedIndex = 0 Then
                    ckb_ChiTietPGD.Visible = True
                End If
            End If
        End If
    End Sub

    Private Sub InBC_02_TCCB(pIdChiNhanh As Integer, pCode As String, pName As String, pDiaDiem As String, pGiamDocLabel As String, pLabHCTC As String, pSoBC_HCTC As String)
        Dim dNgayDB As DateTime
        Dim dNgayKT As DateTime
        Dim db_02TCCB As DataTable = New DataTable()
        Dim bIsFlagALL As Byte = IIf(ckb_ChiTietPGD.Checked = True, 0, 1)

        If CByte(cb_kybc.SelectedIndex) = 5 Then    'Số lượt bổ nhiệm trong khoảng Thời gian (Từ ngày .... Đến ngày)
            dNgayDB = DateTimeUtil.getDate("01/" & numTuThang.Value & "/" & dtpk_nam_bc.Value.Year)
            dNgayKT = DateTimeUtil.getDate("01/" & numDenThang.Value & "/" & dtpk_nam_bc.Value.Year)
            dNgayKT = dNgayKT.AddMonths(1).AddDays(-1)
        Else                                        'Số Lãnh đạo quản lý đến thời điểm cuối Quý được chọn
            dNgayDB = DateTimeUtil.MinSqlDateTime
            dNgayKT = clsHS_BaoCao.GetDate(CByte(cb_kybc.SelectedIndex), dtpk_nam_bc.Value.Year, False)
        End If
        db_02TCCB = _Reports.GetAll_Report02(pIdChiNhanh, dNgayDB, dNgayKT, CByte(cb_kybc.SelectedIndex), bIsFlagALL, 1)
        db_02TCCB.TableName = "BC02"

        Dim sThoiDiem As String = IIf(CByte(cb_kybc.SelectedIndex) = 5, "Số lượt bổ nhiệm: từ " & DateTimeUtil.getShortDate(dNgayDB) & " đến " & DateTimeUtil.getShortDate(dNgayKT), cb_kybc.SelectedItem.ToString() + "/" + dtpk_nam_bc.Value.Year.ToString())
        If pIdChiNhanh = 0 Then         'Số liệu của Toàn quốc
            rpt_BC02QT.SetDataSource(db_02TCCB)
            rpt_BC02QT.SetParameterValue("ten_cn", pName)
            rpt_BC02QT.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
            rpt_BC02QT.SetParameterValue("sobc_hctc", pSoBC_HCTC)
            rpt_BC02QT.SetParameterValue("pr_ThoiDiem", sThoiDiem)
            rpt_BC02QT.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
            rpt_BC02QT.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
            rpt_BC02QT.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
            rpt_BC02QT.SetParameterValue("pr_tinh", pDiaDiem)
            rpt_BC02QT.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
            rpt_BC02QT.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
            rpt_BC02QT.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
            rpt_BC02QT.SetParameterValue("giamdoc_label", pGiamDocLabel)
            rpt_BC02QT.SetParameterValue("labHCTC", pLabHCTC)
            crpv_main.Show()
            crpv_main.Zoom(95)
            crpv_main.DisplayGroupTree = False
            crpv_main.ReportSource = rpt_BC02QT
        ElseIf pIdChiNhanh = 1 Then         'Số liệu của Hội sở Chính
            rptBC02_HSC.SetDataSource(db_02TCCB)
            rptBC02_HSC.SetParameterValue("ten_cn", pName)
            rptBC02_HSC.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
            rptBC02_HSC.SetParameterValue("sobc_hctc", pSoBC_HCTC)
            rptBC02_HSC.SetParameterValue("pr_ThoiDiem", sThoiDiem)
            rptBC02_HSC.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
            rptBC02_HSC.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
            rptBC02_HSC.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
            rptBC02_HSC.SetParameterValue("pr_tinh", pDiaDiem)
            rptBC02_HSC.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
            rptBC02_HSC.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
            rptBC02_HSC.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
            crpv_main.Show()
            crpv_main.Zoom(95)
            crpv_main.DisplayGroupTree = False
            crpv_main.ReportSource = rptBC02_HSC
        Else
            If "000101,000196,000197".Contains(pCode) Then  'Nếu là TTCNTT; TTĐT; SGD
                Select Case pCode
                    Case "000196"    'Trung tâm Công nghệ thông tin
                        rdt_reportBC2CNTT.SetDataSource(db_02TCCB)
                        rdt_reportBC2CNTT.SetParameterValue("ten_cn", pName)
                        rdt_reportBC2CNTT.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
                        rdt_reportBC2CNTT.SetParameterValue("pr_ThoiDiem", sThoiDiem)
                        rdt_reportBC2CNTT.SetParameterValue("giamdoc_lable", lbl_giamdoc.Text.Trim())
                        rdt_reportBC2CNTT.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rdt_reportBC2CNTT.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rdt_reportBC2CNTT.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rdt_reportBC2CNTT.SetParameterValue("pr_tinh", pDiaDiem)
                        rdt_reportBC2CNTT.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                        rdt_reportBC2CNTT.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                        rdt_reportBC2CNTT.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
                        rdt_reportBC2CNTT.SetParameterValue("HCTC_lab", pLabHCTC)
                        rdt_reportBC2CNTT.SetParameterValue("sobc_hctc", pSoBC_HCTC)
                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rdt_reportBC2CNTT
                    Case "000197"    'Trung tâm Đào tạo
                        rdt_reportBC2TTDT.SetDataSource(db_02TCCB)
                        rdt_reportBC2TTDT.SetParameterValue("ma_cn", pCode)
                        rdt_reportBC2TTDT.SetParameterValue("ten_cn", pName)
                        rdt_reportBC2TTDT.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
                        rdt_reportBC2TTDT.SetParameterValue("pr_ThoiDiem", sThoiDiem)
                        rdt_reportBC2TTDT.SetParameterValue("giamdoc_lable", lbl_giamdoc.Text.Trim())
                        rdt_reportBC2TTDT.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rdt_reportBC2TTDT.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rdt_reportBC2TTDT.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rdt_reportBC2TTDT.SetParameterValue("pr_tinh", pDiaDiem)
                        rdt_reportBC2TTDT.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                        rdt_reportBC2TTDT.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                        rdt_reportBC2TTDT.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
                        rdt_reportBC2TTDT.SetParameterValue("HCTC_lab", pLabHCTC)
                        rdt_reportBC2TTDT.SetParameterValue("sobc_hctc", pSoBC_HCTC)
                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rdt_reportBC2TTDT
                    Case "000101"    ' Sở giao dịch
                        rptBC02SGD.SetDataSource(db_02TCCB)
                        rptBC02SGD.SetParameterValue("ma_cn", pCode)
                        rptBC02SGD.SetParameterValue("ten_cn", pName)
                        rptBC02SGD.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
                        rptBC02SGD.SetParameterValue("pr_ThoiDiem", sThoiDiem)
                        rptBC02SGD.SetParameterValue("giamdoc_lable", lbl_giamdoc.Text.Trim())
                        rptBC02SGD.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                        rptBC02SGD.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                        rptBC02SGD.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                        rptBC02SGD.SetParameterValue("pr_tinh", pDiaDiem)
                        rptBC02SGD.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                        rptBC02SGD.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                        rptBC02SGD.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
                        rptBC02SGD.SetParameterValue("HCTC_lab", pLabHCTC)
                        rptBC02SGD.SetParameterValue("sobc_hctc", pSoBC_HCTC)
                        crpv_main.Show()
                        crpv_main.Zoom(95)
                        crpv_main.DisplayGroupTree = False
                        crpv_main.ReportSource = rptBC02SGD

                End Select
            Else                                            'Nếu là các Chi nhánh NHCSXH Tỉnh/TP còn lại
                rdt_reportBC2.SetDataSource(db_02TCCB)
                rdt_reportBC2.SetParameterValue("title_bangd", "Ban Giám đốc")
                rdt_reportBC2.SetParameterValue("title_giamdoc", "Giám đốc")
                rdt_reportBC2.SetParameterValue("title_phogd", "Phó Giám đốc")
                rdt_reportBC2.SetParameterValue("title_debt", "Phòng KTKTNB")
                rdt_reportBC2.SetParameterValue("title_KTNQ", "Phòng KT-NQ")
                rdt_reportBC2.SetParameterValue("title_KHNV", "Phòng KH-NVTD")
                rdt_reportBC2.SetParameterValue("title_HCNS", "Phòng HC-TC")
                rdt_reportBC2.SetParameterValue("title_TH", "Phòng Tin học")
                rdt_reportBC2.SetParameterValue("ten_cn", pName)
                rdt_reportBC2.SetParameterValue("so_bc", edt_so_bc.Text.Trim())
                rdt_reportBC2.SetParameterValue("pr_ThoiDiem", sThoiDiem)
                rdt_reportBC2.SetParameterValue("giamdoc_lable", lbl_giamdoc.Text.Trim())
                rdt_reportBC2.SetParameterValue("nguoilap_bc", edt_nguoilap.Text.Trim())
                rdt_reportBC2.SetParameterValue("truongphong_hctc", edt_truongphong.Text.Trim())
                rdt_reportBC2.SetParameterValue("giamdoc", edt_giamdoc.Text.Trim())
                rdt_reportBC2.SetParameterValue("pr_tinh", pDiaDiem)
                rdt_reportBC2.SetParameterValue("pr_ngay", dtpk_ngaylap_bc.Value.ToString("dd"))
                rdt_reportBC2.SetParameterValue("pr_thang", dtpk_ngaylap_bc.Value.ToString("MM"))
                rdt_reportBC2.SetParameterValue("pr_nam", dtpk_ngaylap_bc.Value.ToString("yyyy"))
                rdt_reportBC2.SetParameterValue("labHCTC", pLabHCTC)
                rdt_reportBC2.SetParameterValue("sobc_hctc", pSoBC_HCTC)
                crpv_main.Show()
                crpv_main.Zoom(95)
                crpv_main.DisplayGroupTree = False
                crpv_main.ReportSource = rdt_reportBC2
            End If
        End If
    End Sub
End Class