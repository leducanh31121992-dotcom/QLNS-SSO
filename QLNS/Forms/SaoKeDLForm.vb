Imports System
Imports System.Data
Imports System.Data.SqlClient
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

Public Class SaoKeDLForm
    Private _ChiNhanhBLL As ChiNhanhBLL = New ChiNhanhBLL()
    Private _HsCanBoBLL As HsCanBoBLL = New HsCanBoBLL()
    Private _BaoCaoBLL As BaoCaoBLL = New BaoCaoBLL()
    Private db_dscanbo As System.Data.DataTable = New System.Data.DataTable()
    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        'vNFInfo.NumberDecimalDigits = 2
        'vNFInfo.NumberGroupSeparator = " "
    End Sub

#Region "---> Events <---"
    Private Sub SaoKeDLForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Bind danh sách các sao kê cần lấy theo yêu cầu
        cb_saoke.Items.Clear()
        cb_saoke.Items.Add("--- Chọn sao kê cần truy vấn ---")
        cb_saoke.Items.Add("01/SK - Sao kê Thống kê số lượng, chất lượng cán bộ")
        cb_saoke.Items.Add("02/SK - Sao kê Tình hình thực hiện lao động - Màng lưới lao động")

        If Not (cb_saoke Is Nothing) Then
            If cb_saoke.Items.Count > 0 Then
                cb_saoke.SelectedIndex = 0
            End If
        End If
        'Danh sách đơn vị trực thuộc NHCSXH cần lấy sao kê
        cb_donvi.DataSource = _ChiNhanhBLL.GetListComBo_ChiNhanh(3, IIf(TRUCTHUOC = 0, "", "---Danh sách đơn vị ---"))
        cb_donvi.DisplayMember = "Display"
        cb_donvi.ValueMember = "Value"
        If Not (cb_donvi Is Nothing) Then
            If cb_donvi.Items.Count > 0 Then
                cb_donvi.SelectedIndex = 0
            End If
        End If
        'Reset Controls
        dtpk_thoidiem.Text = Globals.GetDateTime_ForServerDB()
        rb_tonghop.Checked = True
        rb_chitiet.Checked = False

        lbl_waitting.Text = ""
        Dim LoaiSK As Integer = IIf(cb_saoke.SelectedIndex <= 0, 1, cb_saoke.SelectedIndex)
        _HsCanBoBLL.Create_Frame(dgv_main, 1)
        dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv_main.ColumnHeadersHeight = 30

        Me.dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        Me.dgv_main.ColumnHeadersHeight = Me.dgv_main.ColumnHeadersHeight * 3
        Me.dgv_main.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        AddHandler Me.dgv_main.CellPainting, AddressOf dgv_main_CellPainting

        AddHandler Me.dgv_main.Paint, AddressOf dgv_main_Paint
        AddHandler Me.dgv_main.Scroll, AddressOf dgv_main_Scroll
        AddHandler Me.dgv_main.ColumnWidthChanged, AddressOf dgv_main_ColumnWidthChanged

        'AddHandler Me.dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        'AddHandler Me.dgv_main.Paint, AddressOf dgv_main_Paint
        'AddHandler Me.dgv_main.Scroll, AddressOf dgv_main_Scroll
        'AddHandler Me.dgv_main.ColumnWidthChanged, AddressOf dgv_main_ColumnWidthChanged
    End Sub

    Private Sub btn_tracuu_Click(sender As Object, e As EventArgs) Handles btn_tracuu.Click
        Cursor = Cursors.WaitCursor
        lbl_waitting.Text = "Đang thực hiện tra cứu dữ liệu ..."
        Try
            If (Not (cb_saoke Is Nothing) And (cb_saoke.SelectedIndex >= 0) And (cb_saoke.Items.Count <> 0)) Then
                If (TRUCTHUOC <> 0) Then
                    'If cb_saoke.SelectedIndex <> 1 Then
                    '    If ((cb_donvi Is Nothing) Or (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0)) Then
                    '        MessageBox.Show("Bạn chưa chọn đơn vị cần sao kê dữ liệu. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    '        ActiveControl = cb_donvi
                    '        Return
                    '    End If
                    'End If
                    
                End If

                If (rb_chitiet.Checked = False And rb_tonghop.Checked = False And cb_saoke.SelectedIndex <> 1) Or (rb_chitiet.Checked = False And rb_tonghop.Checked = False And rb_tonghop_pgd.Checked = False And cb_saoke.SelectedIndex = 1) Then
                    MessageBox.Show("Bạn chưa chọn loại sao kê dữ liệu (Tổng hợp/Chi tiết). Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = rb_tonghop
                    Return
                End If

                Dim iFlag_TH_CT As Integer = 0
                iFlag_TH_CT = IIf(rb_tonghop.Checked = True, 1, 2)
                lbl_waitting.Text = "Waitting ..."
                Dim _KieuIn As Byte = 0
                Dim iCountRows As Integer = 0
                Select Case cb_saoke.SelectedIndex
                    Case 1      'Sao kê Số lượng, chất lượng cán bộ
                        Dim _ThoiDiem As DateTime = dtpk_thoidiem.Value
                        Dim _IdDonVi As Integer = 0
                        If ((cb_donvi Is Nothing) Or (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0)) Then
                            _IdDonVi = CInt(cb_donvi.SelectedValue)
                        End If

                        If rb_tonghop.Checked = True Or rb_tonghop_pgd.Checked = True Then
                            _HsCanBoBLL.Create_Frame(dgv_main, 1)
                            dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                            dgv_main.ColumnHeadersHeight = 50
                            dgv_main.Rows.Clear()
                            If Not (db_dscanbo Is Nothing) Then
                                db_dscanbo.Reset()
                            End If
                            If rb_tonghop_pgd.Checked = True Then
                                iFlag_TH_CT = 2
                            Else : iFlag_TH_CT = 1
                            End If
                            lbl_waitting.Text = "Waitting ..."
                            db_dscanbo = _HsCanBoBLL.GetData_SaoKe(5, _IdDonVi, _ThoiDiem, iFlag_TH_CT)
                            If Not (db_dscanbo Is Nothing) Then
                                If (db_dscanbo.Rows.Count > 0) Then
                                    For i As Integer = 0 To db_dscanbo.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_DonVi_HT").Value = db_dscanbo.Rows(i)("DonVi_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_KieuIn").Value = db_dscanbo.Rows(i)("KieuIn").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_KhuVuc_HT").Value = db_dscanbo.Rows(i)("KhuVuc_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DonVi_Id").Value = db_dscanbo.Rows(i)("DonVi_Id").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = db_dscanbo.Rows(i)("Id").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_STT").Value = db_dscanbo.Rows(i)("STT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DonVi_HT").Value = db_dscanbo.Rows(i)("DonVi_HT").ToString().Trim()
                                        Dim sColumnWidth As String = "TongSo_LD;TongSo_LD_Nu;TongSo_DangVien;DanToc_Tso;ChinhTri_CC;ChinhTri_CN;ChinhTri_TC;ChinhTri_SC;TrDoCM_TienSy;TrDoCM_ThacSy;TrDoCM_DaiHoc;TrDoCM_CaoDang;TrDoCM_CaoCapNH;TrDoCM_TrungCap;TrDoCM_SCKhac;NgoaiNgu_DH;NgoaiNgu_C;NgoaiNgu_B;NgoaiNgu_A;TinHoc_DH;TinHoc_C;TinHoc_B;TinHoc_A;TuoiDoi_30;TuoiDoi_31_35;TuoiDoi_36_40;TuoiDoi_41_45;TuoiDoi_46_50;TuoiDoi_51_55;TuoiDoi_56_62;LDao_BanGDNHCS;LDao_GDBanHSC;LDao_BanGDCN;LDao_TP_PP;LDao_BanGDPGD;ChMon_ToTruong_HSC;ChMon_ToTruong_PGD;ChMon_TinDung;ChMon_KeToan;ChMon_KTKTNB;ChMon_TinHoc;ChMon_HcNhanSu;ChMon_ThuQuy;ChMon_LaiXe;ChMon_NvKhac;BaoVe_TapVu"
                                        Dim ARL_ColWidth() As String = Globals.Splip_Strings(sColumnWidth, ";")
                                        For Each _Value As String In ARL_ColWidth
                                            If Not String.IsNullOrEmpty(_Value) Then
                                                If String.IsNullOrEmpty(db_dscanbo.Rows(i)(_Value).ToString()) Or db_dscanbo.Rows(i)(_Value).ToString().Trim() = "0" Then
                                                    dgv_main.Rows(i).Cells("cln_" + _Value).Value = ""
                                                Else
                                                    dgv_main.Rows(i).Cells("cln_" + _Value).Value = Double.Parse(db_dscanbo.Rows(i)(_Value).ToString(), Globals.cultureNum).ToString("N0", vNFInfo)
                                                End If
                                            End If
                                        Next
                                        _KieuIn = CType(db_dscanbo.Rows(i)("KieuIn"), Byte)
                                        _HsCanBoBLL.SetStyleRowGrid(i, dgv_main, _KieuIn, 0)
                                        iCountRows = iCountRows + 1
                                    Next
                                End If
                            End If
                        Else
                            _HsCanBoBLL.Create_Frame(dgv_main, 3)
                            dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                            dgv_main.ColumnHeadersHeight = 30
                            dgv_main.Rows.Clear()
                            If Not (db_dscanbo Is Nothing) Then
                                db_dscanbo.Reset()
                            End If
                            lbl_waitting.Text = "Waitting ..."
                            iFlag_TH_CT = 9
                            db_dscanbo = _HsCanBoBLL.GetData_SaoKe(5, _IdDonVi, _ThoiDiem, iFlag_TH_CT)
                            If Not (db_dscanbo Is Nothing) Then
                                If (db_dscanbo.Rows.Count > 0) Then
                                    For i As Integer = 0 To db_dscanbo.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        'dgv_main.Rows(i).Cells("cln_IdCanBo").Value = db_dscanbo.Rows(i)("IdCanBo").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_STT").Value = db_dscanbo.Rows(i)("STT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_ChiNhanh_HT").Value = db_dscanbo.Rows(i)("ChiNhanh_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DonVi_HT").Value = db_dscanbo.Rows(i)("DonVi_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_HoTen").Value = db_dscanbo.Rows(i)("HoTen").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_MaCB").Value = db_dscanbo.Rows(i)("MaCB").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_NgaySinh_HT").Value = db_dscanbo.Rows(i)("NgaySinh_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_GioiTinh_HT").Value = db_dscanbo.Rows(i)("GioiTinh_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DanToc_HT").Value = db_dscanbo.Rows(i)("DanToc_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_TrinhDoCT_HT").Value = db_dscanbo.Rows(i)("TrinhDoCT_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_PhongBan_HT").Value = db_dscanbo.Rows(i)("PhongBan_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_ChucVu_HT").Value = db_dscanbo.Rows(i)("ChucVu_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_ChuyenMon_HT").Value = db_dscanbo.Rows(i)("ChuyenMon_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DangVien").Value = db_dscanbo.Rows(i)("DangVien").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DT_TrinhDo_HT").Value = db_dscanbo.Rows(i)("DT_TrinhDo_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DT_NgoaiNgu_HT").Value = db_dscanbo.Rows(i)("DT_NgoaiNgu_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_DT_TinHoc_HT").Value = db_dscanbo.Rows(i)("DT_TinHoc_HT").ToString().Trim()
                                        dgv_main.Rows(i).Cells("cln_Loai_CB").Value = db_dscanbo.Rows(i)("Loai_CB").ToString().Trim()
                                        
                                        Dim sColumnWidth As String = "KieuIn;STT_HT;ThoiDiemDL;IdCanBo;DonVi_Id;DonVi_Cd;ChiNhanh_Id;ChiNhanh_Cd;TenChiNhanh_HT;PhongBan_Id;PhongBan_Cd;ChucVu_Id;ChucVu_Cd;ChuyenMon_Id;ChuyenMon_Cd;IdLoaiQD;LoaiQd_Cd;LoaiQd_HT;NgayHL;DanToc_Id;DanToc_Cd;NgaySinh;TuoiDoi;GioiTinh;TrinhDoCT_Id;TrinhDoCT_Cd;DangVien_Id;DangVien_SoThe;DangVien_HT;DT_TrinhDo_Id;DT_TrinhDo_Cd;DT_NgoaiNgu_Cd;DT_TinHoc_Cd;DonViALL"
                                        Dim ARL_ColWidth() As String = Globals.Splip_Strings(sColumnWidth, ";")
                                        For Each _Value As String In ARL_ColWidth
                                            If Not String.IsNullOrEmpty(_Value) Then
                                                If String.IsNullOrEmpty(db_dscanbo.Rows(i)(_Value).ToString()) Or db_dscanbo.Rows(i)(_Value).ToString().Trim() = "0" Then
                                                    dgv_main.Rows(i).Cells("cln_" + _Value).Value = ""
                                                Else
                                                    dgv_main.Rows(i).Cells("cln_" + _Value).Value = db_dscanbo.Rows(i)(_Value).ToString().Trim()
                                                End If
                                            End If
                                        Next
                                        _KieuIn = CType(db_dscanbo.Rows(i)("KieuIn"), Byte)
                                        _HsCanBoBLL.SetStyleRowGrid(i, dgv_main, _KieuIn, 0)
                                        iCountRows = iCountRows + 1
                                    Next
                                End If
                            End If
                        End If
                       
                        lbl_waitting.Text = ""
                        lbl_records.Text = IIf(iFlag_TH_CT = 1, " Tổng đơn vị: ", " Tổng số bản ghi: ") & iCountRows.ToString()

                        'dgv_main.AutoGenerateColumns = True
                        'dgv_main.Columns.Clear()
                        'dgv_main.ColumnHeadersHeight = 50
                        'dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                        'dgv_main.Columns().Clear()
                        'If Not (db_dscanbo Is Nothing) Then
                        '    db_dscanbo.Reset()
                        'End If

                        'lbl_waitting.Text = "Waitting ..."
                        'db_dscanbo = _HsCanBoBLL.GetData_SaoKe(cb_saoke.SelectedIndex, _IdDonVi, _ThoiDiem, iFlag_TH_CT)
                        ''Bind danh sách sao kê ra Lưới dữ liệu
                        'dgv_main.DataSource = db_dscanbo
                        'FixGrid_Columns_01(cb_saoke.SelectedIndex, iFlag_TH_CT)
                        'SetLayout_MangLuoi()
                        'dgv_main.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(204, 209, 209)
                        'Dim iCountRows As Integer = 0
                        ''Bôi đậm mầu các dòng Mục lớn: A, B, C, ... Hoặc I, II, III ...
                        'If Not (dgv_main.DataSource Is Nothing) Then
                        '    If dgv_main.Rows.Count > 1 Then
                        '        For i As Integer = 0 To dgv_main.Rows.Count - 1
                        '            dgv_main.Rows(i).DefaultCellStyle.ForeColor = Color.Navy
                        '            Dim _STT As String = dgv_main.Item("STT", i).Value.ToString().Trim()
                        '            Dim sampleArray As New List(Of String)(New String() {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX", "XXI", "XXII", "XXIII", "XXIV", "XXV", "XXVI", "XXVII", "XXVIII", "XXIX", "XXX", "XXXI", "XXXII", "XXIII", "XXXIV", "XXXV"})
                        '            If sampleArray.Contains(_STT) Or _STT = "" Then
                        '                dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Bold)
                        '                dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(97, 135, 214)
                        '            Else
                        '                iCountRows = iCountRows + 1
                        '            End If
                        '        Next
                        '    End If
                        'End If
                        'lbl_waitting.Text = ""
                        'lbl_records.Text = IIf(iFlag_TH_CT = 1, " Tổng đơn vị: ", " Tổng số bản ghi: ") & iCountRows.ToString()

                    Case 2      'Sao kê Tình hình thực hiện lao động - Màng lưới lao động
                        Dim _ThoiDiem As DateTime = dtpk_thoidiem.Value
                        Dim _IdDonVi As Integer = CInt(cb_donvi.SelectedValue)
                        If _IdDonVi = 0 Then
                            _HsCanBoBLL.Create_Frame(dgv_main, 2)
                        Else
                            _HsCanBoBLL.Create_Frame(dgv_main, 4)
                        End If

                        dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                        dgv_main.ColumnHeadersHeight = 30
                        dgv_main.Rows.Clear()
                        If Not (db_dscanbo Is Nothing) Then
                            db_dscanbo.Reset()
                        End If

                        lbl_waitting.Text = "Waitting ..."
                        'db_dscanbo = _HsCanBoBLL.GetData_SaoKe(cb_saoke.SelectedIndex, _IdDonVi, _ThoiDiem, iFlag_TH_CT)
                        db_dscanbo = _BaoCaoBLL.GetDLBaoCao_BaoCao_TinhHinhLD_01TCCB(DateTimeUtil.DateTimeToString(_ThoiDiem, "yyyy-MM-dd"), _IdDonVi, 2)

                        If Not (db_dscanbo Is Nothing) Then
                            If (db_dscanbo.Rows.Count > 0) Then
                                For i As Integer = 0 To db_dscanbo.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = db_dscanbo.Rows(i)("Id").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_DonVi_Id").Value = db_dscanbo.Rows(i)("DonVi_Id").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_KieuIn").Value = db_dscanbo.Rows(i)("KieuIn").ToString().Trim()

                                    dgv_main.Rows(i).Cells("cln_STT").Value = db_dscanbo.Rows(i)("STT").ToString().Trim()
                                    dgv_main.Rows(i).Cells("cln_DonVi_HT").Value = db_dscanbo.Rows(i)("DonVi_HT").ToString().Trim()
                                    Dim sColumnWidth As String = "Thang1_ThucHien_TC;Thang2_ThucHien_TC;Thang3_ThucHien_TC;Thang1_ThongBao_TC_CN;Thang2_ThongBao_TC_CN;Thang3_ThongBao_TC_CN;Thang1_DaiHan_TB;Thang1_NganHan_TB;Thang2_DaiHan_TB;Thang2_NganHan_TB;Thang3_DaiHan_TB;Thang3_NganHan_TB;Thang1_ThongBao_TC;Thang2_ThongBao_TC;Thang3_ThongBao_TC;DuNo_TH;DuNo_QH;DuNo_KH;Thang1_DaiHan_TBCN;Thang1_DaiHan_TH;Thang1_NganHan_TBCN;Thang1_NganHan_TH;Thang2_DaiHan_TBCN;Thang2_DaiHan_TH;Thang2_NganHan_TBCN;Thang2_NganHan_TH;Thang3_DaiHan_TBCN;Thang3_DaiHan_TH;Thang3_NganHan_TBCN;Thang3_NganHan_TH;So_ToTKVV;SoKH_DN;So_XaPhuong;So_DiemGD;TongDN"
                                    Dim ARL_ColWidth() As String = Globals.Splip_Strings(sColumnWidth, ";")
                                    For Each _Value As String In ARL_ColWidth
                                        If Not String.IsNullOrEmpty(_Value) Then
                                            If String.IsNullOrEmpty(db_dscanbo.Rows(i)(_Value).ToString()) Or db_dscanbo.Rows(i)(_Value).ToString().Trim() = "0" Then
                                                dgv_main.Rows(i).Cells("cln_" + _Value).Value = ""
                                            Else
                                                dgv_main.Rows(i).Cells("cln_" + _Value).Value = Double.Parse(db_dscanbo.Rows(i)(_Value).ToString(), Globals.cultureNum).ToString("N0", vNFInfo)
                                            End If
                                        End If
                                    Next
                                    _KieuIn = CType(db_dscanbo.Rows(i)("KieuIn"), Byte)
                                    _HsCanBoBLL.SetStyleRowGrid(i, dgv_main, _KieuIn, 0)
                                    iCountRows = iCountRows + 1
                                Next
                            End If
                        End If
                        lbl_waitting.Text = ""
                        lbl_records.Text = IIf(iFlag_TH_CT = 1, " Tổng đơn vị: ", " Tổng số bản ghi: ") & iCountRows.ToString()


                        'dgv_main.AutoGenerateColumns = True
                        'dgv_main.Columns.Clear()
                        'dgv_main.ColumnHeadersHeight = 50
                        'dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                        'dgv_main.Columns().Clear()
                        'If Not (db_dscanbo Is Nothing) Then
                        '    db_dscanbo.Reset()
                        'End If

                        'lbl_waitting.Text = "Waitting ..."
                        'db_dscanbo = _HsCanBoBLL.GetData_SaoKe(cb_saoke.SelectedIndex, _IdDonVi, _ThoiDiem, iFlag_TH_CT)
                        'dgv_main.DataSource = db_dscanbo
                        'FixGrid_Columns_02(cb_saoke.SelectedIndex, iFlag_TH_CT)
                        ''SetLayout_MangLuoi()
                        'dgv_main.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(204, 209, 209)
                        'Dim iCountRows As Integer = 0
                        ''Bôi đậm mầu các dòng Mục lớn: A, B, C, ... Hoặc I, II, III ...
                        'If Not (dgv_main.DataSource Is Nothing) Then
                        '    If dgv_main.Rows.Count > 1 Then
                        '        For i As Integer = 0 To dgv_main.Rows.Count - 1
                        '            dgv_main.Rows(i).DefaultCellStyle.ForeColor = Color.Navy
                        '            Dim _STT As String = dgv_main.Item("STT", i).Value.ToString().Trim()
                        '            Dim sampleArray As New List(Of String)(New String() {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX", "XXI", "XXII", "XXIII", "XXIV", "XXV", "XXVI", "XXVII", "XXVIII", "XXIX", "XXX", "XXXI", "XXXII", "XXIII", "XXXIV", "XXXV"})
                        '            If sampleArray.Contains(_STT) Or _STT = "" Then
                        '                dgv_main.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(dgv_main.Font, FontStyle.Bold)
                        '                dgv_main.Rows(i).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(97, 135, 214)
                        '            Else
                        '                iCountRows = iCountRows + 1
                        '            End If
                        '        Next
                        '    End If
                        'End If
                        'lbl_waitting.Text = ""
                        'lbl_records.Text = IIf(iFlag_TH_CT = 1, " Tổng đơn vị: ", " Tổng số bản ghi: ") & iCountRows.ToString()
                    Case Else

                End Select
            End If
            lbl_waitting.Text = ""
            Cursor = Cursors.Default
            dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            dgv_main.ColumnHeadersHeight = 80
        Catch ex As Exception
            lbl_waitting.Text = "Error ..."
            Cursor = Cursors.Default
            MessageBox.Show("Lỗi truy vấn Danh sách cán bộ: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_xuatexcel_Click(sender As Object, e As EventArgs) Handles btn_xuatexcel.Click
        If (db_dscanbo Is Nothing Or db_dscanbo.Rows.Count <= 0) Then
            MessageBox.Show("Không có danh sách để xuất file excel (Hãy click nút lệnh 'Tra cứu' trước). Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        Cursor = Cursors.WaitCursor
        lbl_waitting.Text = "Waitting for data export..."
        'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Dim iFlag_TH_CT As Integer = 0, iRowStart As Integer = 0
        iFlag_TH_CT = IIf(rb_tonghop.Checked = True, 1, 2)
        Dim _ProvinceName As String = ""
        Dim _IdDonVi As Integer = CInt(cb_donvi.SelectedValue)
        If _IdDonVi = 0 Or _IdDonVi = 1 Then
            _ProvinceName = "BAN TỔ CHỨC CÁN BỘ"
        Else
            _ProvinceName = SoftSqlHelper.GetString(String.Format("Select Top 1 IsNull(Upper(Ten_Goi),'') From ChiNhanh Where Id={0}", _IdDonVi), "")
        End If
        Dim _ThoiDiem As DateTime = dtpk_thoidiem.Value

        Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
        Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
        Dim sPath_Export As String = "C:\Temp\"
        Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
        sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")

        If cb_saoke.SelectedIndex = 1 Then          'Sao kê Số lượng, chất lượng cán bộ
            sFileExcelTemplate = sPathExcelTemplate + IIf(iFlag_TH_CT = 1, "BC_SoLuong_ChatLuong_CanBo_TH_08_TCCB_2022.xlsx", "SK_SoLuong_ChatLuong_CanBo_TH_08_TCCB_CT.xlsx")
            sFileNameEx = "08_TCCB_So luong chat luong can bo " + IIf(iFlag_TH_CT = 1, "Tong hop", "Chi tiet") + "_" + DateTime.Now.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
        ElseIf cb_saoke.SelectedIndex = 2 Then      'Sao kê Tình hình thực hiện lao động - Màng lưới lao động
            If _IdDonVi = 0 Then
                sFileExcelTemplate = sPathExcelTemplate + "SK_MangLuoiLaoDong_01_TCCB_TQ.xlsx"
            Else
                sFileExcelTemplate = sPathExcelTemplate + "SK_MangLuoiLaoDong_01_TCCB_CN.xlsx"
            End If

            sFileNameEx = "01_TCCB_Mang luoi lao dong_" + DateTime.Now.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
        End If

        If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
            Using excelEngine As ExcelEngine = New ExcelEngine()
                Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                Dim worksheet As IWorksheet = workbook.Worksheets(0)
                If cb_saoke.SelectedIndex = 1 Then          'Sao kê Số lượng, chất lượng cán bộ
                    iRowStart = 9
                    sColNameEnd = IIf(iFlag_TH_CT = 1, "AV", "Q")
                    worksheet.Range("A4").Text = IIf(iFlag_TH_CT = 1, "BÁO CÁO THỐNG KÊ SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ", "BÁO CÁO THỐNG KÊ CHI TIẾT SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ")
                    worksheet.Range("A5").Text = "THỜI ĐIỂM " + _ThoiDiem.ToString("dd/MM/yyyy")
                    worksheet.Range("A3").Text = _ProvinceName
                ElseIf cb_saoke.SelectedIndex = 2 Then      'Sao kê Tình hình thực hiện lao động - Màng lưới lao động
                    iRowStart = 12
                    sColNameEnd = "S"
                    worksheet.Range("A5").Text = "BÁO CÁO TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MÀNG LƯỚI HOẠT ĐỘNG"
                    worksheet.Range("A6").Text = db_dscanbo.Rows(0)("ThoiDiem_HT").ToString().ToUpper()
                    worksheet.Range("A3").Text = _ProvinceName
                End If

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
                Dim iRowA As Integer = 0
                If cb_saoke.SelectedIndex = 2 Then      'Sao kê Tình hình thực hiện lao động - Màng lưới lao động
                    worksheet.Range("D9:G9").Text = "Tháng " + (dtpk_thoidiem.Value.Month - 2).ToString()
                    worksheet.Range("H9:K9").Text = "Tháng " + (dtpk_thoidiem.Value.Month - 1).ToString()
                    worksheet.Range("L9:O9").Text = "Tháng " + (dtpk_thoidiem.Value.Month).ToString()
                End If

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

    Private Sub btn_huybo_Click(sender As Object, e As EventArgs) Handles btn_huybo.Click
        lbl_waitting.Text = ""
        dtpk_thoidiem.Text = Globals.GetDateTime_ForServerDB()
        rb_tonghop.Checked = True
        rb_chitiet.Checked = False

        If Not (cb_saoke Is Nothing) Then
            If cb_saoke.Items.Count > 0 Then
                cb_saoke.SelectedIndex = 0
            End If
        End If

        If Not (cb_donvi Is Nothing) Then
            If cb_donvi.Items.Count > 0 Then
                cb_donvi.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        If Not (db_dscanbo Is Nothing) Then
            db_dscanbo.Dispose()
        End If
        Close()
    End Sub

    Private Sub cb_saoke_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_saoke.SelectedIndexChanged
        rb_chitiet.Enabled = True
        rb_tonghop_pgd.Enabled = False
        If (Not (cb_saoke Is Nothing) And (cb_saoke.Items.Count <> 0)) Then
            If (cb_saoke.SelectedIndex = 2) Then
                rb_chitiet.Enabled = False
                rb_tonghop.Checked = True
            ElseIf cb_saoke.SelectedIndex = 1 Then
                rb_tonghop_pgd.Enabled = True
            End If
        End If
    End Sub
#End Region

#Region "---> Định nghĩa lưới dữ liệu <---"
    Private Sub FixGrid_Columns_01_Bo(ByVal _IndexSK As Byte, ByVal _FlagCTTH As Byte)
        dgv_main.ColumnHeadersDefaultCellStyle.ForeColor = Color.Navy
        dgv_main.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Tahoma", 8, FontStyle.Bold)

        For Each _column As DataGridViewColumn In dgv_main.Columns
            _column.ReadOnly = True
            _column.SortMode = DataGridViewColumnSortMode.NotSortable
            _column.Name = _column.HeaderText.Trim()
            If _IndexSK = 1 Then    'Sao kê Chi tiết/Tổng hợp: Số lượng, chất lượng cán bộ
                If _FlagCTTH = 1 Then   'Sao kê tổng hợp
                    Select Case _column.HeaderText.Trim()
                        Case "STT"
                            _column.HeaderText = "STT"
                            _column.Width = 40
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "Ten_CN"
                            _column.HeaderText = "Đơn vị"
                            _column.Width = 120
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "TongSo_LD"
                            _column.HeaderText = "Tổng số lao động có mặt"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TongSo_LD_Nu"
                            _column.HeaderText = "Nữ"
                            _column.Width = 45
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TongSo_DangVien"
                            _column.HeaderText = "Đảng viên"
                            _column.Width = 50
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "DanToc_TSo"
                            _column.HeaderText = "Dân tộc thiểu số"
                            _column.Width = 50
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChinhTri_CC"
                            _column.HeaderText = "Cao cấp" 'Trình độ chính trị - 
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChinhTri_TC"
                            _column.HeaderText = "Trung cấp"        'Trình độ chính trị - 
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChinhTri_SC"
                            _column.HeaderText = "Sơ cấp"           'Trình độ chính trị - 
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_TienSy"
                            _column.HeaderText = "Trình độ chuyên môn - Tiến sỹ"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_ThacSy"
                            _column.HeaderText = "Trình độ chuyên môn - Thạc sỹ"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_DaiHoc"
                            _column.HeaderText = "Trình độ chuyên môn - Đại học"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_CaoDang"
                            _column.HeaderText = "Trình độ chuyên môn - Cao đẳng"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_CaoCapNH"
                            _column.HeaderText = "Trình độ chuyên môn - Cao cấp NH"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_TrungCap"
                            _column.HeaderText = "Trình độ chuyên môn - Trung cấp"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TrDoCM_SCKhac"
                            _column.HeaderText = "Trình độ chuyên môn - Sơ cấp và khác"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "NgoaiNgu_DH"
                            _column.HeaderText = "Trình độ ngoại ngữ - Đại học"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "NgoaiNgu_A"
                            _column.HeaderText = "Trình độ ngoại ngữ - Bằng A"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "NgoaiNgu_B"
                            _column.HeaderText = "Trình độ ngoại ngữ - Bằng B"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "NgoaiNgu_C"
                            _column.HeaderText = "Trình độ ngoại ngữ - Bằng C"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TinHoc_DH"
                            _column.HeaderText = "Trình độ tin học - Đại học"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TinHoc_A"
                            _column.HeaderText = "Trình độ tin học - Bằng A"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TinHoc_B"
                            _column.HeaderText = "Trình độ tin học - Bằng B"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TinHoc_C"
                            _column.HeaderText = "Trình độ tin học - Bằng C"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_30"
                            _column.HeaderText = "Tuổi đời - Dưới 30"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_31_35"
                            _column.HeaderText = "Tuổi đời - Từ 31 đến 35"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_36_40"
                            _column.HeaderText = "Tuổi đời - Từ 36 đến 40"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_41_45"
                            _column.HeaderText = "Tuổi đời - Từ 41 đến 45"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_46_50"
                            _column.HeaderText = "Tuổi đời - Từ 46 đến 50"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_51_55"
                            _column.HeaderText = "Tuổi đời - Từ 51 đến 55"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TuoiDoi_55_60"
                            _column.HeaderText = "Tuổi đời - Từ 55 đến 60"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "LDao_BanGDCN"
                            _column.HeaderText = "Chức danh LĐ - Ban TGĐ/Ban GĐ Chi nhánh"
                            _column.Width = 80
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "LDao_TP_PP"
                            _column.HeaderText = "Chức danh LĐ - GĐ,PGĐ Ban/Trưởng, Phó phòng"
                            _column.Width = 80
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "LDao_BanGDPGD"
                            _column.HeaderText = "Chức danh LĐ - GĐ, PGĐ huyện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "LDao_ToTruong"
                            _column.HeaderText = "Chức danh LĐ - Tổ trưởng"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_TinDung"
                            _column.HeaderText = "Cán bộ CMNV - Tín dụng"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_KeToan"
                            _column.HeaderText = "Cán bộ CMNV - Kế toán"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_TinHoc"
                            _column.HeaderText = "Cán bộ CMNV - Tin học"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_ThuQuy"
                            _column.HeaderText = "Cán bộ CMNV - Thủ quỹ"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_KTKTNB"
                            _column.HeaderText = "Cán bộ CMNV - KTKTNB"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_HcNhanSu"
                            _column.HeaderText = "Cán bộ CMNV - Hành chính, Nhân sự"
                            _column.Width = 70
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_LaiXe"
                            _column.HeaderText = "Cán bộ CMNV - Lái xe"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "ChMon_NvKhac"
                            _column.HeaderText = "Cán bộ CMNV - Nhân viên khác"
                            _column.Width = 65
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "BaoVe_TapVu"
                            _column.HeaderText = "Bảo vệ, tạp vụ"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case Else
                            _column.Visible = False
                    End Select
                Else                    'Sao kê chi tiết
                    Select Case _column.HeaderText.Trim()
                        Case "STT"
                            _column.HeaderText = "STT"
                            _column.Width = 40
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "DonViALL"
                            _column.HeaderText = "Đơn vị công tác"
                            _column.Width = 200
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "HoTen"
                            _column.HeaderText = "Họ và tên"
                            _column.Width = 130
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "MaCB"
                            _column.HeaderText = "Mã CB"
                            _column.Width = 55
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "Ngay_Sinh"
                            _column.HeaderText = "Ngày sinh"
                            _column.Width = 70
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "GioiTinh"
                            _column.HeaderText = "Giới tính"
                            _column.Width = 45
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "DanToc"
                            _column.HeaderText = "Dân tộc"
                            _column.Width = 55
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "TrinhDoCT"
                            _column.HeaderText = "Trình độ chính trị"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "PhongBan"
                            _column.HeaderText = "Phòng ban"
                            _column.Width = 130
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "ChucVu"
                            _column.HeaderText = "Chức vụ"
                            _column.Width = 120
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "ChuyenMon"
                            _column.HeaderText = "Chuyên môn công tác theo QĐNS"
                            _column.Width = 110
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "DangVien"
                            _column.HeaderText = "Đảng viên"
                            _column.Width = 40
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "DT_TrinhDo"
                            _column.HeaderText = "Trình độ đào tạo"
                            _column.Width = 70
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "NgoaiNgu"
                            _column.HeaderText = "Trình độ ngoại ngữ"
                            _column.Width = 65
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "TinHoc"
                            _column.HeaderText = "Trình độ tin học"
                            _column.Width = 65
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "Loai_CB"
                            _column.HeaderText = "Loại cán bộ"
                            _column.Width = 120
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case Else
                            _column.Visible = False
                    End Select
                End If
            End If
        Next
    End Sub

    Private Sub FixGrid_Columns_02_Bo(ByVal _IndexSK As Byte, ByVal _FlagCTTH As Byte)
        dgv_main.ColumnHeadersDefaultCellStyle.ForeColor = Color.Navy
        dgv_main.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Tahoma", 8, FontStyle.Bold)

        For Each _column As DataGridViewColumn In dgv_main.Columns
            _column.ReadOnly = True
            _column.SortMode = DataGridViewColumnSortMode.NotSortable
            If _IndexSK = 2 Then    'Sao kê Chi tiết/Tổng hợp: Số lượng, chất lượng cán bộ
                If _FlagCTTH = 1 Then   'Sao kê tổng hợp
                    Select Case _column.HeaderText.Trim()
                        Case "STT"
                            _column.HeaderText = "STT"
                            _column.Width = 40
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "Ten_CN"
                            _column.HeaderText = "Đơn vị"
                            _column.Width = 120
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                        Case "TongSoCT_TWTB"
                            _column.HeaderText = "Tổng số chỉ tiêu TW thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang1_DH_TB"
                            _column.HeaderText = "Tháng 1 - Dài hạn - Thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang1_DH_TH"
                            _column.HeaderText = "Tháng 1 - Dài hạn - Thực hiện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang1_NH_TB"
                            _column.HeaderText = "Tháng 1 - Ngắn hạn - Thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang1_NH_TH"
                            _column.HeaderText = "Tháng 1 - Ngắn hạn - Thực hiện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang2_DH_TB"
                            _column.HeaderText = "Tháng 2 - Dài hạn - Thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang2_DH_TH"
                            _column.HeaderText = "Tháng 2 - Dài hạn - Thực hiện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang2_NH_TB"
                            _column.HeaderText = "Tháng 2 - Ngắn hạn - Thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang2_NH_TH"
                            _column.HeaderText = "Tháng 2 - Ngắn hạn - Thực hiện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang3_DH_TB"
                            _column.HeaderText = "Tháng 3 - Dài hạn - Thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang3_DH_TH"
                            _column.HeaderText = "Tháng 3 - Dài hạn - Thực hiện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang3_NH_TB"
                            _column.HeaderText = "Tháng 3 - Ngắn hạn - Thông báo"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "Thang3_NH_TH"
                            _column.HeaderText = "Tháng 3 - Ngắn hạn - Thực hiện"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"

                        Case "MLLD_SoHuyen"
                            _column.HeaderText = "Màng lưới hoạt động - Tổng số quận, huyện"
                            _column.Width = 70
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "MLLD_SoXa"
                            _column.HeaderText = "Màng lưới hoạt động - Tổng số xã phường"
                            _column.Width = 70
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "MLLD_SoDiemGD"
                            _column.HeaderText = "Màng lưới hoạt động - Tổng số điểm giao dịch"
                            _column.Width = 70
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case "TongDuNo"
                            _column.HeaderText = "Dư nợ (tỷ đồng)"
                            _column.Width = 60
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                            _column.DefaultCellStyle.Format = "##,0"
                        Case Else
                            _column.Visible = False
                    End Select
                Else                    'Sao kê chi tiết
                    Select Case _column.HeaderText.Trim()
                        Case "STT"
                            _column.HeaderText = "STT"
                            _column.Width = 40
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                        Case "Ten_CN"
                            _column.HeaderText = "Đơn vị"
                            _column.Width = 120
                            _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft

                        Case Else
                            _column.Visible = False
                    End Select
                End If
            End If
        Next
    End Sub
#End Region

#Region "---> Functions Merget DataGridView <---"
    Private Sub SetLayout_MangLuoi_Bo()
        AddHandler Me.dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        AddHandler Me.dgv_main.Paint, AddressOf dgv_main_Paint
        AddHandler Me.dgv_main.Scroll, AddressOf dgv_main_Scroll
        AddHandler Me.dgv_main.ColumnWidthChanged, AddressOf dgv_main_ColumnWidthChanged
    End Sub

    Private Sub dgv_main_ColumnWidthChanged(sender As Object, e As System.Windows.Forms.DataGridViewColumnEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_main.DisplayRectangle
        rtHeader.Height = Me.dgv_main.ColumnHeadersHeight / 2
        Me.dgv_main.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_main_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_main.DisplayRectangle
        rtHeader.Height = Me.dgv_main.ColumnHeadersHeight / 2
        Me.dgv_main.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_main_CellPainting(sender As Object, e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then
            Dim r2 As Rectangle = e.CellBounds
            r2.Y += e.CellBounds.Height / 3
            r2.Height = e.CellBounds.Height / 3
            e.PaintBackground(r2, True)
            e.PaintContent(r2)
            e.Handled = True
        End If
    End Sub

    Private Sub dgv_main_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim format As StringFormat = New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center
        Dim rColumAllHeight As Integer = 0

        'Chỉ lặp những cột nào cần merge
        If ((cb_saoke.SelectedIndex = 1 Or cb_saoke.SelectedIndex = 0) And rb_tonghop.Checked = True) Then               'BÁO CÁO THỐNG KÊ SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 8 Then       'Trình độ chính trị
                    Dim rTrinhDoCT As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width

                    rTrinhDoCT.X += 1
                    rTrinhDoCT.Y += 1
                    rTrinhDoCT.Width = rTrinhDoCT.Width + w1 + w2 + w3 - 1
                    rTrinhDoCT.Height = rTrinhDoCT.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rTrinhDoCT.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTrinhDoCT)
                    e.Graphics.DrawRectangle(Pens.Silver, rTrinhDoCT)
                    e.Graphics.DrawString("Trình độ chính trị", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTrinhDoCT, format)
                End If
                If j = 12 Then       'Trình độ chuyên môn
                    Dim rTdChuyenMon As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                    Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width
                    rTdChuyenMon.X += 1
                    rTdChuyenMon.Y += 1
                    rTdChuyenMon.Width = rTdChuyenMon.Width + w1 + w2 + w3 + w4 + w5 + w6 - 1
                    rTdChuyenMon.Height = rTdChuyenMon.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rTdChuyenMon.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTdChuyenMon)
                    e.Graphics.DrawRectangle(Pens.Silver, rTdChuyenMon)
                    e.Graphics.DrawString("Trình độ chuyên môn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTdChuyenMon, format)
                End If

                If j = 19 Then       'Trình độ ngoại ngữ 
                    Dim rTdNgoaiNgu As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rTdNgoaiNgu.X += 1
                    rTdNgoaiNgu.Y += 1
                    rTdNgoaiNgu.Width = rTdNgoaiNgu.Width + w1 + w2 + w3 - 1
                    rTdNgoaiNgu.Height = rTdNgoaiNgu.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rTdNgoaiNgu.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTdNgoaiNgu)
                    e.Graphics.DrawRectangle(Pens.Silver, rTdNgoaiNgu)
                    e.Graphics.DrawString("Trình độ ngoại ngữ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTdNgoaiNgu, format)
                End If
                If j = 23 Then       'Trình độ tin học
                    Dim rTdTinHoc As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rTdTinHoc.X += 1
                    rTdTinHoc.Y += 1
                    rTdTinHoc.Width = rTdTinHoc.Width + w1 + w2 + w3 - 1
                    rTdTinHoc.Height = rTdTinHoc.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rTdTinHoc.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTdTinHoc)
                    e.Graphics.DrawRectangle(Pens.Silver, rTdTinHoc)
                    e.Graphics.DrawString("Trình độ tin học", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTdTinHoc, format)
                End If
                If j = 27 Then       'Tuổi đời 
                    Dim rTuoiDoi As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                    Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width
                    rTuoiDoi.X += 1
                    rTuoiDoi.Y += 1
                    rTuoiDoi.Width = rTuoiDoi.Width + w1 + w2 + w3 + w4 + w5 + w6 - 1
                    rTuoiDoi.Height = rTuoiDoi.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rTuoiDoi.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTuoiDoi)
                    e.Graphics.DrawRectangle(Pens.Silver, rTuoiDoi)
                    e.Graphics.DrawString("Tuổi đời", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTuoiDoi, format)
                End If
                If j = 34 Then       ''Chức danh Lãnh đạo
                    Dim rCvLanhDao As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width

                    rCvLanhDao.X += 1
                    rCvLanhDao.Y += 1
                    rCvLanhDao.Width = rCvLanhDao.Width + w1 + w2 + w3 + w4 - 1
                    rCvLanhDao.Height = rCvLanhDao.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rCvLanhDao.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rCvLanhDao)
                    e.Graphics.DrawRectangle(Pens.Silver, rCvLanhDao)
                    e.Graphics.DrawString("Chức danh Lãnh đạo", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rCvLanhDao, format)
                End If

                If j = 39 Then       'Cán bộ CMNV
                    Dim rCanBo As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                    Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                    Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width
                    Dim w7 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 7, -1, True).Width
                    Dim w8 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 8, -1, True).Width
                    Dim w9 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 9, -1, True).Width
                    rCanBo.X += 1
                    rCanBo.Y += 1
                    rCanBo.Width = rCanBo.Width + w1 + w2 + w3 + w4 + w5 + w6 + w7 + w8 + w9 - 1
                    rCanBo.Height = rCanBo.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rCanBo.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rCanBo)
                    e.Graphics.DrawRectangle(Pens.Silver, rCanBo)
                    e.Graphics.DrawString("Cán bộ CMNV", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rCanBo, format)
                End If
            Next
        ElseIf (cb_saoke.SelectedIndex = 2) Then       'Báo cáo mạng lưới lao động
            Dim _ThoiDiemBC As DateTime = dtpk_thoidiem.Value
            Dim sThang01 As String = "Tháng " + _ThoiDiemBC.AddMonths(-2).Month.ToString("D2")
            Dim sThang02 As String = "Tháng " + _ThoiDiemBC.AddMonths(-1).Month.ToString("D2")
            Dim sThang03 As String = "Tháng " + _ThoiDiemBC.Month.ToString("D2")
            For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
                If j = 5 Then       'Tháng 1 - Dài Hạn và Ngắn hạn
                    Dim rThang1 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rThang1.X += 1
                    rThang1.Y += 1
                    rThang1.Width = rThang1.Width + w1 + w2 + w3 - 1
                    rThang1.Height = rThang1.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rThang1.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThang1)
                    e.Graphics.DrawRectangle(Pens.Silver, rThang1)
                    e.Graphics.DrawString(sThang01, Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThang1, format)
                End If

                If j = 5 Then       'Tháng 1 - Dài Hạn
                    Dim rThangDH1 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rThangDH1.X += 1
                    rThangDH1.Y += rColumAllHeight + 1
                    rThangDH1.Width = rThangDH1.Width + w1 - 1
                    rThangDH1.Height = rThangDH1.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThangDH1)
                    e.Graphics.DrawRectangle(Pens.Silver, rThangDH1)
                    e.Graphics.DrawString("Dài hạn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThangDH1, format)
                End If
                If j = 7 Then       'Tháng 1 - Ngắn hạn
                    Dim rThangNH1 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rThangNH1.X += 1
                    rThangNH1.Y += rColumAllHeight + 1
                    rThangNH1.Width = rThangNH1.Width + w1 - 1
                    rThangNH1.Height = rThangNH1.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThangNH1)
                    e.Graphics.DrawRectangle(Pens.Silver, rThangNH1)
                    e.Graphics.DrawString("Ngắn hạn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThangNH1, format)
                End If
                '===============================================================================================================================
                If j = 9 Then       'Tháng 2 - Dài Hạn và Ngắn hạn
                    Dim rThang2 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rThang2.X += 1
                    rThang2.Y += 1
                    rThang2.Width = rThang2.Width + w1 + w2 + w3 - 1
                    rThang2.Height = rThang2.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rThang2.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThang2)
                    e.Graphics.DrawRectangle(Pens.Silver, rThang2)
                    e.Graphics.DrawString(sThang02, Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThang2, format)
                End If

                If j = 9 Then       'Tháng 2 - Dài Hạn
                    Dim rThangDH2 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rThangDH2.X += 1
                    rThangDH2.Y += rColumAllHeight + 1
                    rThangDH2.Width = rThangDH2.Width + w1 - 1
                    rThangDH2.Height = rThangDH2.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThangDH2)
                    e.Graphics.DrawRectangle(Pens.Silver, rThangDH2)
                    e.Graphics.DrawString("Dài hạn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThangDH2, format)
                End If
                If j = 11 Then       'Tháng 2 - Ngắn hạn
                    Dim rThangNH2 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rThangNH2.X += 1
                    rThangNH2.Y += rColumAllHeight + 1
                    rThangNH2.Width = rThangNH2.Width + w1 - 1
                    rThangNH2.Height = rThangNH2.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThangNH2)
                    e.Graphics.DrawRectangle(Pens.Silver, rThangNH2)
                    e.Graphics.DrawString("Ngắn hạn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThangNH2, format)
                End If


                '===============================================================================================================================
                If j = 13 Then       'Tháng 3 - Dài Hạn và Ngắn hạn
                    Dim rThang3 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                    rThang3.X += 1
                    rThang3.Y += 1
                    rThang3.Width = rThang3.Width + w1 + w2 + w3 - 1
                    rThang3.Height = rThang3.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rThang3.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThang3)
                    e.Graphics.DrawRectangle(Pens.Silver, rThang3)
                    e.Graphics.DrawString(sThang03, Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThang3, format)
                End If

                If j = 13 Then       'Tháng 3 - Dài Hạn
                    Dim rThangDH3 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rThangDH3.X += 1
                    rThangDH3.Y += rColumAllHeight + 1
                    rThangDH3.Width = rThangDH3.Width + w1 - 1
                    rThangDH3.Height = rThangDH3.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThangDH3)
                    e.Graphics.DrawRectangle(Pens.Silver, rThangDH3)
                    e.Graphics.DrawString("Dài hạn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThangDH3, format)
                End If
                If j = 15 Then       'Tháng 3 - Ngắn hạn
                    Dim rThangNH3 As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    rThangNH3.X += 1
                    rThangNH3.Y += rColumAllHeight + 1
                    rThangNH3.Width = rThangNH3.Width + w1 - 1
                    rThangNH3.Height = rThangNH3.Height / 3 - 2              ' /3 la vi 3 lop
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rThangNH3)
                    e.Graphics.DrawRectangle(Pens.Silver, rThangNH3)
                    e.Graphics.DrawString("Ngắn hạn", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                        New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rThangNH3, format)
                End If

                If j = 17 Then       'Màng lưới hoạt động 
                    Dim rCvLanhDao As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                    Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                    Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                    Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width

                    rCvLanhDao.X += 1
                    rCvLanhDao.Y += 1
                    rCvLanhDao.Width = rCvLanhDao.Width + w1 + w2 + w3 - 1
                    rCvLanhDao.Height = rCvLanhDao.Height / 3 - 2 ' /3 la vi 3 lop
                    rColumAllHeight = rCvLanhDao.Height
                    e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rCvLanhDao)
                    e.Graphics.DrawRectangle(Pens.Silver, rCvLanhDao)
                    e.Graphics.DrawString("Màng lưới hoạt động", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                                          New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rCvLanhDao, format)
                End If
            Next
        End If


    End Sub

    Dim r1 As System.Drawing.Rectangle, r2 As System.Drawing.Rectangle, r3 As System.Drawing.Rectangle
    Private Sub Merge(ByVal _title As String, ByVal _start As Integer, ByVal _end As Integer, ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim format As StringFormat = New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center

        r1 = dgv_main.GetCellDisplayRectangle(_start, -1, True)
        r2 = dgv_main.GetCellDisplayRectangle(_end - 1, -1, True)
        r3 = r1

        r3.Height = r1.Height / 2
        r3.Y = r1.Y + 2         'Top
        r3.X = r1.X + 2         'Left
        r3.Width = r2.Width + r2.X - r1.X
        r2 = dgv_main.GetCellDisplayRectangle(_end, -1, True)
        r3.Width = r3.Width + r2.Width - 4

        e.Graphics.FillRectangle(New SolidBrush(dgv_main.ColumnHeadersDefaultCellStyle.BackColor), r3.X, r3.Y, r3.Width, r3.Height)
        e.Graphics.DrawLine(New Pen(dgv_main.GridColor, 1), r3.X, r3.Bottom, r3.X + r3.Width, r3.Bottom)
        e.Graphics.DrawString(_title, dgv_main.ColumnHeadersDefaultCellStyle.Font, New SolidBrush(dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), r3, format)
    End Sub
#End Region

    
End Class
