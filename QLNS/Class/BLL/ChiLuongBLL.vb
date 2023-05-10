Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Syncfusion.XlsIO
Imports Syncfusion.ExcelToPdfConverter
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Collections
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Public Class ChiLuongBLL

    ''' <summary>
    ''' Hàm thiết lập định dạng hiển thị của Cột dữ liệu trong lưới DataGridView
    ''' </summary>
    ''' <param name="pNameDGV">Tên DataGridView</param>
    ''' <param name="pListColumns">Danh sách tên các cột của DataGridView</param>
    ''' <param name="pDelimiter">Ký tự phân cách tên các cột</param>
    ''' <param name="pAlignment">Chỉ số xác định Style Alignment của Cột. 1 - Trái; 2 - Giữa; 3 - Phải</param>
    ''' <remarks></remarks>
    Private Sub FormatStyleAlignment_ColumnDataGridView(ByVal pNameDGV As DataGridView, ByVal pListColumns As String, ByVal pDelimiter As String, pAlignment As Byte)
        Dim ARL_ColRight() As String = Globals.Splip_Strings(pListColumns, pDelimiter)
        For Each _ValueNull As String In ARL_ColRight
            If Not String.IsNullOrEmpty(_ValueNull) Then
                If pAlignment = 1 Then
                    pNameDGV.Columns("cln_" + _ValueNull).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                ElseIf pAlignment = 2 Then
                    pNameDGV.Columns("cln_" + _ValueNull).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ElseIf pAlignment = 3 Then
                    pNameDGV.Columns("cln_" + _ValueNull).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                End If
            End If
        Next
    End Sub

    Public Function FormatMoneyInTextbox(ByVal vObjTextBox As TextBox) As TextBox
        Dim Text As String
        Dim selStart, i As Integer
        Dim commaCount_Before As Integer = 0
        Dim commaCount_After As Integer = 0
        If vObjTextBox.Text.Trim = "" Then Return vObjTextBox
        Text = vObjTextBox.Text
        selStart = vObjTextBox.SelectionStart
        commaCount_Before = 0
        commaCount_After = 0
        For i = 0 To Text.Length - 1
            If Text.Substring(i, 1) = "," Then commaCount_Before += 1
        Next
        vObjTextBox.Text = formatMoney(Text)
        Text = vObjTextBox.Text
        For i = 0 To Text.Length - 1
            If Text.Substring(i, 1) = "," Then commaCount_After += 1
        Next
        vObjTextBox.SelectionStart = selStart + (commaCount_After - commaCount_Before)
        Return vObjTextBox
    End Function

    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal state As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        dgv_name.EnableHeadersVisualStyles = False
        dgv_name.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray 'Color.FromArgb(178, 219, 255)
        dgv_name.ColumnHeadersDefaultCellStyle.ForeColor = Color.Blue 'Color.FromArgb(16, 37, 127)
        Dim sColumnVisible As String = "", sColumnRight As String = "", sColumnTMP = ""
        Select Case state
            Case 1 ' Chi lương tháng - Kỳ 1
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice) '1
                dgv_name.Columns.Add("cln_STT", "STT") '2
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên") '3
                dgv_name.Columns.Add("cln_ThongTin_HT", "Chức vụ, vị trí công tác")                                                 '4
                dgv_name.Columns.Add("cln_LCB_HeSo", "Hệ số lương cấp bậc" + vbNewLine)                                             '5
                dgv_name.Columns.Add("cln_PhCap_ChVu_HeSo", "CV" + vbNewLine)                                                       '6
                dgv_name.Columns.Add("cln_PhCap_TrNhiem_HeSo", "TN" + vbNewLine)                                                    '7
                dgv_name.Columns.Add("cln_PhCap_DocHai_HeSo", "ĐH" + vbNewLine)                                                     '8
                dgv_name.Columns.Add("cln_SoNgay_KhongLV", "Số ngày không làm việc trong tháng tại đơn vị (Không lương, PhC các loại)")                            '9
                dgv_name.Columns.Add("cln_MucHuong", "Tỷ lệ hưởng lương theo HĐLĐ")                                                 '10
                dgv_name.Columns.Add("cln_PhCap_ThuHut_SoTien", "Thu hút" + vbNewLine + vbNewLine + vbNewLine)                      '11
                dgv_name.Columns.Add("cln_PhCap_KhuVuc_SoTien", "Khu vực" + vbNewLine + vbNewLine + vbNewLine)                      '12
                dgv_name.Columns.Add("cln_Luong_V1", "Tổng tiền lương V1 (100%)")                                                   '13
                dgv_name.Columns.Add("cln_Luong_V1_TamUng_K1", "Tổng tiền lương V1 được tạm ứng (80%)")                             '14
                dgv_name.Columns.Add("cln_BHXH_CB_SoTien", "Trích BHXH (8%)" + vbNewLine)                                           '15
                dgv_name.Columns.Add("cln_BHYT_CB_SoTien", "Trích BHYT (1.5%)" + vbNewLine)                                         '16
                dgv_name.Columns.Add("cln_BHTN_CB_SoTien", "Trích BHTN (1%)" + vbNewLine)                                           '17
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac", "Các khoản trích nộp khác" + vbNewLine)                                  '18
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac_GhiChu", "Ghi chú trích nộp khác" + vbNewLine + vbNewLine + vbNewLine)     '19
                dgv_name.Columns.Add("cln_Tru_CacKhoan_Tong", "Tổng các khoản phải trừ" + vbNewLine)                                '20
                dgv_name.Columns.Add("cln_Tru_CacKhoan_GhiChu", "Ghi chú các khoản phải trừ" + vbNewLine)                           '21
                dgv_name.Columns.Add("cln_Luong_ThucLinh", "Tiền lương V1 thực lĩnh kỳ này")                                        '22
                dgv_name.Columns.Add("cln_GhiChu_ChiLuong", "Ghi chú")                                                              '23
                dgv_name.Columns.Add("cln_SoTK_NHCS", "Số tài khoản NHCS")                                                          '24
                sColumnVisible = "MaCB;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;Loai_CB;LCB_BacLuong_Id;LCB_NgHuong;KieuIn;Luong_TTV_Ma;Luong_TTV;Luong_CSo;BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TrangThai;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_NgHuong;PhCap_ThuHut_HeSo;PhCap_KhuVuc_HeSo;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_V1_TamUng_K2;BHXH_DV_SoTien;BHYT_DV_SoTien;BHTN_DV_SoTien;Tinh_Thue_TNCN;SoNguoi_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;DPCD_CB_SoTien;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;SoNgayLViec_Thang;Luong_V1_ConLai;Luong_V1_ThucTra;Luong_V2_TamUng;MucTamUng_V2;HeSoLuong_V2;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc;Loai_CB;Loai_CB_HT;LCB_Tong_HeSo"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next
                dgv_name.Columns("cln_IdCanBo").Width = 50                              '2
                dgv_name.Columns("cln_STT").Width = 35                              '2
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_HoTen").Width = 150                           '3
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThongTin_HT").Width = 120                     '4
                dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_LCB_HeSo").Width = 55                         '5
                dgv_name.Columns("cln_PhCap_ChVu_HeSo").Width = 55                  '6
                dgv_name.Columns("cln_PhCap_TrNhiem_HeSo").Width = 55               '7
                dgv_name.Columns("cln_PhCap_DocHai_HeSo").Width = 55                '8
                dgv_name.Columns("cln_SoNgay_KhongLV").Width = 80
                dgv_name.Columns("cln_SoNgay_KhongLV").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_MucHuong").Width = 60
                dgv_name.Columns("cln_MucHuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_PhCap_ThuHut_SoTien").Width = 80                '9
                dgv_name.Columns("cln_PhCap_KhuVuc_SoTien").Width = 80                '10
                dgv_name.Columns("cln_Luong_V1").Width = 100                           '11
                dgv_name.Columns("cln_Luong_V1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Luong_V1_TamUng_K1").Width = 100                 '12
                dgv_name.Columns("cln_Luong_V1_TamUng_K1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_BHXH_CB_SoTien").Width = 80                       '13
                dgv_name.Columns("cln_BHYT_CB_SoTien").Width = 75                       '14
                dgv_name.Columns("cln_BHTN_CB_SoTien").Width = 75                       '15
                dgv_name.Columns("cln_Tru_Khoan_Khac").Width = 90                       '16
                dgv_name.Columns("cln_Tru_Khoan_Khac_GhiChu").Width = 150                       '17
                dgv_name.Columns("cln_Tru_CacKhoan_Tong").Width = 95                       '18
                dgv_name.Columns("cln_Tru_CacKhoan_GhiChu").Width = 150                       '17
                dgv_name.Columns("cln_Luong_ThucLinh").Width = 95                    '19
                dgv_name.Columns("cln_Luong_ThucLinh").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GhiChu_ChiLuong").Width = 200                      '20
                dgv_name.Columns("cln_GhiChu_ChiLuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTK_NHCS").Width = 130                            '21
                dgv_name.Columns("cln_SoTK_NHCS").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                ''Căn chỉnh tiêu đề
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_MucHuong").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoNgay_KhongLV").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                sColumnRight = "LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1;Luong_V1_TamUng_K1;BHXH_CB_SoTien;BHYT_CB_SoTien;BHTN_CB_SoTien;Tru_Khoan_Khac;Tru_CacKhoan_Tong;Luong_ThucLinh"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.Columns("cln_IdCanBo").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                    dgv_name.Columns(i).DefaultCellStyle.BackColor = Color.Gainsboro
                Next

                'Để column check có thể edit
                sColumnTMP = "Choice;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu;Tru_CacKhoan_GhiChu;GhiChu_ChiLuong;SoNgay_KhongLV;MucHuong"
                ARL_Cols = Globals.Splip_Strings(sColumnTMP, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).ReadOnly = False
                        dgv_name.Columns("cln_" + _Value).DefaultCellStyle.BackColor = Color.White
                    End If
                Next
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray
            Case 2 ' Chi lương tháng - Kỳ 2
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice) '1
                dgv_name.Columns.Add("cln_STT", "STT") '2
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên") '3
                dgv_name.Columns.Add("cln_ThongTin_HT", "Chức vụ, vị trí công tác") '4
                dgv_name.Columns.Add("cln_LCB_HeSo", "Hệ số lương cấp bậc" + vbNewLine) '5
                dgv_name.Columns.Add("cln_PhCap_ChVu_HeSo", "CV" + vbNewLine) '6
                dgv_name.Columns.Add("cln_PhCap_TrNhiem_HeSo", "TN" + vbNewLine) '7
                dgv_name.Columns.Add("cln_PhCap_DocHai_HeSo", "ĐH" + vbNewLine) '8
                dgv_name.Columns.Add("cln_SoNgay_KhongLV", "Số ngày không làm việc trong tháng tại đơn vị (Không lương, PhC các loại)")                            '9
                dgv_name.Columns.Add("cln_MucHuong", "Tỷ lệ hưởng lương theo HĐLĐ")                                                 '10

                dgv_name.Columns.Add("cln_PhCap_ThuHut_SoTien", "Thu hút" + vbNewLine + vbNewLine) '11
                dgv_name.Columns.Add("cln_PhCap_KhuVuc_SoTien", "Khu vực" + vbNewLine + vbNewLine) '12
                dgv_name.Columns.Add("cln_Luong_V1_TamUng_K2", "Tiền lương V1 kỳ này (20%)" + vbNewLine)                            '13
                dgv_name.Columns.Add("cln_SoNgayNghi_TruLuong", "Số ngày nghỉ không lương, PhC các loại") '14
                dgv_name.Columns.Add("cln_SoTienNghi_TruLuong", "TL, PhC các loại ngày nghỉ không lương phải trừ") '15
                dgv_name.Columns.Add("cln_Luong_V1_ConLai", "Tiền lương V1 còn lại" + vbNewLine + vbNewLine) '16
                dgv_name.Columns.Add("cln_Luong_V1_ThucTra", "Tiền lương V1 (không bao gồm phụ cấp TH, KV, TL nghỉ không lương)") '17
                dgv_name.Columns.Add("cln_MucTamUng_V2", "Mức tạm ứng (%)" + vbNewLine) '18  HeSoLuong_V2
                dgv_name.Columns.Add("cln_Luong_V2_TamUng", "Tiền lương V2 tạm ứng" + vbNewLine) '19

                dgv_name.Columns.Add("cln_TNTT_SoTien_NopThue", "Tạm khấu trừ thuế TNCN" + vbNewLine) '20
                dgv_name.Columns.Add("cln_DPCD_CB_SoTien", "1% đoàn phí công đoàn" + vbNewLine) '21
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac", "Quỹ XH, tình nghĩa, các khoản trích nộp khác" + vbNewLine) '22
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac_GhiChu", "Ghi chú trích nộp khác" + vbNewLine + vbNewLine) '23

                dgv_name.Columns.Add("cln_Luong_ThucLinh", "Tổng số tiền thực lĩnh kỳ này") '24
                dgv_name.Columns.Add("cln_GhiChu_ChiLuong", "Ghi chú") '25
                dgv_name.Columns.Add("cln_SoTK_NHCS", "Số tài khoản thụ hưởng") '26
                dgv_name.Columns.Add("cln_SoNguoi_PhuThuoc", "Số người phụ thuộc") '26
                dgv_name.Columns.Add("cln_TNTT_SoTien", "Thu nhập tính thuế")

                sColumnVisible = "MaCB;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;Loai_CB;LCB_BacLuong_Id;LCB_NgHuong;KieuIn;Luong_TTV_Ma;Luong_TTV;Luong_CSo;BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TrangThai;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_NgHuong;PhCap_ThuHut_HeSo;PhCap_KhuVuc_HeSo;BHXH_DV_SoTien;BHYT_DV_SoTien;BHTN_DV_SoTien;TNTT_TyLe;DPCD_LD_NN_SoTien;SoNgayLViec_Thang;Luong_V1;Luong_V1_TamUng_K1;Tru_CacKhoan_Tong;Tru_CacKhoan_GhiChu;BHXH_CB_SoTien;BHYT_CB_SoTien;BHTN_CB_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tinh_Thue_TNCN;HeSoLuong_V2;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc;Loai_CB;Loai_CB_HT;LCB_Tong_HeSo"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_IdCanBo").Width = 50                              '0
                dgv_name.Columns("cln_STT").Width = 35                              '2
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_HoTen").Width = 150                           '3
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThongTin_HT").Width = 120                     '4
                dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_LCB_HeSo").Width = 55                         '5
                dgv_name.Columns("cln_PhCap_ChVu_HeSo").Width = 55                  '6
                dgv_name.Columns("cln_PhCap_TrNhiem_HeSo").Width = 55               '7
                dgv_name.Columns("cln_PhCap_DocHai_HeSo").Width = 55                '8
                dgv_name.Columns("cln_SoNgay_KhongLV").Width = 80
                dgv_name.Columns("cln_SoNgay_KhongLV").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_MucHuong").Width = 60
                dgv_name.Columns("cln_MucHuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_PhCap_ThuHut_SoTien").Width = 80                '9
                dgv_name.Columns("cln_PhCap_KhuVuc_SoTien").Width = 80                '10
                dgv_name.Columns("cln_Luong_V1_TamUng_K2").Width = 95                       '11
                dgv_name.Columns("cln_SoNgayNghi_TruLuong").Width = 70                      '12
                dgv_name.Columns("cln_SoTienNghi_TruLuong").Width = 90                      '13
                dgv_name.Columns("cln_Luong_V1_ConLai").Width = 100                         '14
                dgv_name.Columns("cln_Luong_V1_ThucTra").Width = 100                        '15
                dgv_name.Columns("cln_MucTamUng_V2").Width = 55                                '16
                dgv_name.Columns("cln_Luong_V2_TamUng").Width = 95                          '17
                dgv_name.Columns("cln_TNTT_SoTien_NopThue").Width = 90                      '18
                dgv_name.Columns("cln_DPCD_CB_SoTien").Width = 80                           '19
                dgv_name.Columns("cln_Tru_Khoan_Khac").Width = 90                           '20
                dgv_name.Columns("cln_Tru_Khoan_Khac_GhiChu").Width = 150                   '21
                dgv_name.Columns("cln_Luong_ThucLinh").Width = 110                    '22
                dgv_name.Columns("cln_Luong_ThucLinh").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GhiChu_ChiLuong").Width = 200                      '23
                dgv_name.Columns("cln_GhiChu_ChiLuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTK_NHCS").Width = 130                            '24
                dgv_name.Columns("cln_SoTK_NHCS").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoNguoi_PhuThuoc").Width = 60
                dgv_name.Columns("cln_SoNguoi_PhuThuoc").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_TNTT_SoTien").Width = 99
                dgv_name.Columns("cln_TNTT_SoTien").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                ''Căn chỉnh tiêu đề
                FormatStyleAlignment_ColumnDataGridView(dgv_name, "STT;SoNgay_KhongLV;MucHuong", ";", 2)
                sColumnRight = "LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1_TamUng_K2;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_V1_ConLai;Luong_V1_ThucTra;MucTamUng_V2;Luong_V2_TamUng;TNTT_SoTien_NopThue;DPCD_CB_SoTien;Tru_Khoan_Khac;Luong_ThucLinh;SoNguoi_PhuThuoc;TNTT_SoTien"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.Columns("cln_IdCanBo").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                    dgv_name.Columns(i).DefaultCellStyle.BackColor = Color.Gainsboro
                Next

                'Để column check có thể edit
                sColumnTMP = "Choice;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu;Tru_CacKhoan_GhiChu;GhiChu_ChiLuong;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;TNTT_SoTien_NopThue;MucTamUng_V2;DPCD_CB_SoTien;SoNgay_KhongLV;MucHuong"
                ARL_Cols = Globals.Splip_Strings(sColumnTMP, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).ReadOnly = False
                        dgv_name.Columns("cln_" + _Value).DefaultCellStyle.BackColor = Color.White
                    End If
                Next

                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray

            Case 3 ' Chi lương tháng - Lao động Ngắn hạn (Bảo vệ/Tạp vụ)
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                          '1
                dgv_name.Columns.Add("cln_STT", "STT")                                          '2
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên")                                  '3
                dgv_name.Columns.Add("cln_ThongTin_HT", "Vị trí công việc" + vbNewLine + vbNewLine)                     '4
                dgv_name.Columns.Add("cln_Loai_HDNH_DB", "Định biên" + vbNewLine + vbNewLine)                               '5
                dgv_name.Columns.Add("cln_Loai_HDNH_PT", "Phụ trợ" + vbNewLine + vbNewLine)                                   '6
                dgv_name.Columns.Add("cln_Luong_V1", "Tổng số lương được hưởng")                '7
                dgv_name.Columns.Add("cln_BHXH_CB_SoTien", "Trích BHXH (8%)" + vbNewLine + vbNewLine)                   '8
                dgv_name.Columns.Add("cln_BHYT_CB_SoTien", "Trích BHYT (1.5%)" + vbNewLine + vbNewLine)                 '9
                dgv_name.Columns.Add("cln_BHTN_CB_SoTien", "Trích BHTN (1%)" + vbNewLine + vbNewLine)                   '10
                dgv_name.Columns.Add("cln_DPCD_CB_SoTien", "Đoàn phí CĐ  1%" + vbNewLine + vbNewLine)                   '1
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac", "Các khoản trích nộp khác" + vbNewLine + vbNewLine)          '12
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac_GhiChu", "Ghi chú trích nộp khác" + vbNewLine + vbNewLine + vbNewLine)     '13
                dgv_name.Columns.Add("cln_Tru_CacKhoan_Tong", "Tổng các khoản phải trừ" + vbNewLine + vbNewLine)        '14
                dgv_name.Columns.Add("cln_Tru_CacKhoan_GhiChu", "Ghi chú các khoản phải trừ" + vbNewLine + vbNewLine)   '15
                dgv_name.Columns.Add("cln_LamDem_SoGio", "Số giờ" + vbNewLine + vbNewLine + vbNewLine)                              '16
                dgv_name.Columns.Add("cln_LamDem_SoTienPC", "Số tiền phụ cấp" + vbNewLine + vbNewLine + vbNewLine)                  '17
                dgv_name.Columns.Add("cln_Luong_ThucLinh", "Tổng số tiền thực lĩnh")            '18
                dgv_name.Columns.Add("cln_GhiChu_ChiLuong", "Ghi chú")                          '19
                dgv_name.Columns.Add("cln_SoTK_NHCS", "Số tài khoản NHCS")                      '20
                sColumnVisible = "MaCB;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;Loai_CB;Loai_CB_HT;SoNgay_KhongLV;MucHuong;LCB_BacLuong_Id;LCB_NgHuong;KieuIn;Luong_TTV_Ma;Luong_TTV;Luong_CSo;BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TrangThai;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_NgHuong;PhCap_ThuHut_HeSo;PhCap_KhuVuc_HeSo;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_V1_TamUng_K2;BHXH_DV_SoTien;BHYT_DV_SoTien;BHTN_DV_SoTien;Tinh_Thue_TNCN;SoNguoi_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;SoNgayLViec_Thang;Luong_V1_ConLai;Luong_V1_ThucTra;Luong_V2_TamUng;MucTamUng_V2;HeSoLuong_V2;IsBHXH;IsBHYT;IsBHTN;IsDPCD;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1_TamUng_K1;DonGia_LamDem_Gio;Loai_HDNH;PhCap_KhuVuc_ST_Goc;PhCap_ThuHut_ST_Goc"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_IdCanBo").Width = 50                              '2
                dgv_name.Columns("cln_STT").Width = 35                              '2
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_HoTen").Width = 150                           '3
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThongTin_HT").Width = 110                     '4
                dgv_name.Columns("cln_Loai_HDNH_DB").Width = 40
                dgv_name.Columns("cln_Loai_HDNH_PT").Width = 40

                'dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Luong_V1").Width = 100                           '11
                dgv_name.Columns("cln_Luong_V1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_BHXH_CB_SoTien").Width = 80
                dgv_name.Columns("cln_BHYT_CB_SoTien").Width = 80
                dgv_name.Columns("cln_BHTN_CB_SoTien").Width = 80
                dgv_name.Columns("cln_DPCD_CB_SoTien").Width = 80
                dgv_name.Columns("cln_Tru_Khoan_Khac").Width = 100
                dgv_name.Columns("cln_Tru_Khoan_Khac_GhiChu").Width = 110
                dgv_name.Columns("cln_Tru_CacKhoan_Tong").Width = 85
                dgv_name.Columns("cln_Tru_CacKhoan_GhiChu").Width = 110
                dgv_name.Columns("cln_LamDem_SoGio").Width = 60
                dgv_name.Columns("cln_LamDem_SoTienPC").Width = 100

                dgv_name.Columns("cln_Luong_ThucLinh").Width = 95
                dgv_name.Columns("cln_Luong_ThucLinh").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GhiChu_ChiLuong").Width = 150
                dgv_name.Columns("cln_GhiChu_ChiLuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTK_NHCS").Width = 130                            '21
                dgv_name.Columns("cln_SoTK_NHCS").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                ''Căn chỉnh tiêu đề
                sColumnTMP = "LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1;Luong_V1_TamUng_K1;BHXH_CB_SoTien;BHYT_CB_SoTien;BHTN_CB_SoTien;DPCD_CB_SoTien;Tru_Khoan_Khac;Tru_CacKhoan_Tong;Luong_ThucLinh;LamDem_SoGio;LamDem_SoTienPC"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnTMP, ";", 3)
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Loai_HDNH_DB").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Loai_HDNH_PT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_IdCanBo").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).DefaultCellStyle.BackColor = Color.Gainsboro
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Để column check có thể edit
                sColumnTMP = "Choice;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu;Tru_CacKhoan_GhiChu;LamDem_SoGio;GhiChu_ChiLuong"
                ARL_Cols = Globals.Splip_Strings(sColumnTMP, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).ReadOnly = False
                        dgv_name.Columns("cln_" + _Value).DefaultCellStyle.BackColor = Color.White
                    End If
                Next

                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray

            Case 5 ' Chi lương Cho cán bộ Tập sự (Lương trọn gói) - Cán bộ tập sự (Đào tạo, tập nghề)
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                      '1
                dgv_name.Columns.Add("cln_STT", "STT")                                      '2
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên")                              '3
                dgv_name.Columns.Add("cln_ThongTin_HT", "Chức vụ, vị trí công tác")         '4
                dgv_name.Columns.Add("cln_LCB_HeSo", "HSL cấp bậc" + vbNewLine)                                         '5
                dgv_name.Columns.Add("cln_PhCap_ChVu_HeSo", "Chức vụ" + vbNewLine)                                      '6
                dgv_name.Columns.Add("cln_PhCap_TrNhiem_HeSo", "Trách nhiệm" + vbNewLine)                               '7
                dgv_name.Columns.Add("cln_PhCap_DocHai_HeSo", "Độc hại" + vbNewLine)                                    '8
                dgv_name.Columns.Add("cln_Luong_V1", "Tổng số lương được hưởng")                                       '9
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac", "Các khoản trích nộp khác" + vbNewLine + vbNewLine)          '10
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac_GhiChu", "Ghi chú trích nộp khác" + vbNewLine + vbNewLine)     '11
                dgv_name.Columns.Add("cln_Luong_ThucLinh", "Tiền lương thực lĩnh")                                                  '12
                dgv_name.Columns.Add("cln_GhiChu_ChiLuong", "Ghi chú")                                  '13
                dgv_name.Columns.Add("cln_SoTK_NHCS", "Số tài khoản NHCS")                              '14
                sColumnVisible = "KieuIn;MaCB;NgaySinh;GioiTinh;ChiNhanh_Cd;DonVi_Cd;PhongBan_Cd;ChucVu_Cd;ChuyenMon_Cd;LoaiQD_Cd;LoaiHDLD_Cd;LCB_Tong_HeSo;LCB_BacLuong_Id;LCB_NgHuong;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_HeSo;PhCap_ThuHut_ST_Goc;PhCap_ThuHut_SoTien;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_HeSo;PhCap_KhuVuc_ST_Goc;PhCap_KhuVuc_SoTien;PhCap_KhuVuc_NgHuong;SoNgay_KhongLV;SoNgayLViec_Thang;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_TTV_Ma;Luong_TTV;Luong_CSo;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;BHXH_CB;BHXH_CB_SoTien;BHXH_DV;BHXH_DV_SoTien;BHYT_CB;BHYT_CB_SoTien;BHYT_DV;BHYT_DV_SoTien;BHTN_CB;BHTN_CB_SoTien;BHTN_DV;BHTN_DV_SoTien;DPCD_CB;DPCD_CB_SoTien;DPCD_LD_NN;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tru_CacKhoan_Tong;Tru_CacKhoan_GhiChu;Tinh_Thue_TNCN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;SoNguoi_PhuThuoc;Luong_V1_ConLai;Luong_V1_ThucTra;MucTamUng_V2;HeSoLuong_V2;Luong_V2_TamUng;MucHuong;Loai_CB;Loai_CB_HT;TrangThai;TrangThai_HT;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;Loai_HDNH_DB;Loai_HDNH_PT"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next
                dgv_name.Columns("cln_STT").Width = 50
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_HoTen").Width = 180
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThongTin_HT").Width = 140
                dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_LCB_HeSo").Width = 65
                dgv_name.Columns("cln_PhCap_ChVu_HeSo").Width = 60
                dgv_name.Columns("cln_PhCap_TrNhiem_HeSo").Width = 60
                dgv_name.Columns("cln_PhCap_DocHai_HeSo").Width = 60
                dgv_name.Columns("cln_Luong_V1").Width = 130
                dgv_name.Columns("cln_Luong_V1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_Tru_Khoan_Khac").Width = 100
                dgv_name.Columns("cln_Tru_Khoan_Khac_GhiChu").Width = 110
                dgv_name.Columns("cln_Luong_ThucLinh").Width = 130
                dgv_name.Columns("cln_Luong_ThucLinh").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTK_NHCS").Width = 140                            '21
                dgv_name.Columns("cln_SoTK_NHCS").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GhiChu_ChiLuong").Width = 200                      '20
                dgv_name.Columns("cln_GhiChu_ChiLuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                sColumnRight = "LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;Luong_V1;Luong_ThucLinh;Tru_Khoan_Khac;Tru_CacKhoan_Tong"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.Columns("cln_IdCanBo").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Để column check có thể edit
                sColumnTMP = "Choice;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu;GhiChu_ChiLuong"
                ARL_Cols = Globals.Splip_Strings(sColumnTMP, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).ReadOnly = False
                        dgv_name.Columns("cln_" + _Value).DefaultCellStyle.BackColor = Color.White
                    End If
                Next
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray

            Case 10 '01/TL -  TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - TIỀN LƯƠNG
                dgv_name.Columns.Add("cln_TongHopId", "TongHopId")                              '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                              '1
                dgv_name.Columns.Add("cln_STT", "TT")                                               '2
                dgv_name.Columns.Add("cln_ThongTin_HT", "Tên đơn vị/phòng ban")                     '3
                dgv_name.Columns.Add("cln_PhanLoai_HT", "Phân loại")                                '4
                dgv_name.Columns.Add("cln_LaoDong_HT", "Lao động")                                  '5
                dgv_name.Columns.Add("cln_KyBC", "Kỳ báo cáo")                                      '6

                dgv_name.Columns.Add("cln_SoLD", "Số LĐ" + vbNewLine + vbNewLine + vbNewLine)                                           '7
                dgv_name.Columns.Add("cln_LCB_Tong_HeSo", "Tổng" + vbNewLine + vbNewLine + vbNewLine)                                   '8
                dgv_name.Columns.Add("cln_LCB_HeSo", "Cấp bậc" + vbNewLine + vbNewLine)                                     '9
                dgv_name.Columns.Add("cln_PhCap_ChVu_HeSo", "CV" + vbNewLine)                                   '10
                dgv_name.Columns.Add("cln_PhCap_TrNhiem_HeSo", "TN" + vbNewLine)                                '11
                dgv_name.Columns.Add("cln_PhCap_DocHai_HeSo", "ĐH" + vbNewLine)                                 '12
                dgv_name.Columns.Add("cln_Luong_TTV", "Lương tối thiểu vùng" + vbNewLine + vbNewLine + vbNewLine)                       '13
                dgv_name.Columns.Add("cln_TL_PhuCap", "Tiền lương và phụ cấp" + vbNewLine + vbNewLine + vbNewLine)                      '14
                dgv_name.Columns.Add("cln_PhCap_ThuHut_SoTien", "Thu hút" + vbNewLine + vbNewLine + vbNewLine)                          '15
                dgv_name.Columns.Add("cln_PhCap_KhuVuc_SoTien", "Khu vực" + vbNewLine + vbNewLine + vbNewLine)                          '16
                dgv_name.Columns.Add("cln_Luong_V1", "Tổng tiền lương V1" + vbNewLine + vbNewLine + vbNewLine)                          '17
                dgv_name.Columns.Add("cln_Luong_V2_TamUng", "TẠM ỨNG TIỀN LƯƠNG V2" + vbNewLine + vbNewLine + vbNewLine)                '18
                dgv_name.Columns.Add("cln_TongLuong_TamUng", "TỔNG TIỀN LƯƠNG TẠM ỨNG V1+V2" + vbNewLine + vbNewLine + vbNewLine)       '19

                dgv_name.Columns.Add("cln_Loai_HDNH_DB_SoLD", "Số LĐ" + vbNewLine)                              '20
                dgv_name.Columns.Add("cln_Loai_HDNH_DB_SoTien", "Số tiền" + vbNewLine)                          '21
                dgv_name.Columns.Add("cln_Loai_HDNH_PT_SoLD", "Số LĐ" + vbNewLine)                              '22
                dgv_name.Columns.Add("cln_Loai_HDNH_PT_SoTien", "Số tiền" + vbNewLine)                          '23
                dgv_name.Columns.Add("cln_Tong_TienCong", "Tổng cộng tiền công" + vbNewLine + vbNewLine)                    '24
                dgv_name.Columns.Add("cln_LamDem_SoTienPC", "Tiền Phụ cấp làm đêm" + vbNewLine + vbNewLine)                 '25

                dgv_name.Columns.Add("cln_Tong_TL_TC_TamChi", "TỔNG CỘNG TIỀN LƯƠNG, TIỀN CÔNG TẠM CHI" + vbNewLine + vbNewLine + vbNewLine)                '26
                dgv_name.Columns.Add("cln_LdNgViec_CoBHXH_SoLD", "Số LĐ")                                             '27
                dgv_name.Columns.Add("cln_LdNgViec_CoBHXH_HSLPC", "HSL và PC")                                        '28
                dgv_name.Columns.Add("cln_LdNgViec_KoBHXH_SoLD", "Số LĐ")                                             '29
                dgv_name.Columns.Add("cln_LdNgViec_KoBHXH_HSLPC", "HSL và PC")                                        '30
                dgv_name.Columns.Add("cln_GhiChu_CL", "Ghi chú" + vbNewLine + vbNewLine + vbNewLine + vbNewLine)                                                        '31
                dgv_name.Columns.Add("cln_TrangThai_HT", "Trạng thái" + vbNewLine + vbNewLine + vbNewLine)                                                        '31
                dgv_name.Columns.Add("cln_KieuIn", "KieuIn")
                sColumnVisible = "OrderNo;ChiNhanh_Cd;ChiNhanh_HT;DonVi_CL_Cd;DonVi_CL_HT;PhongBan_CL_Cd;PhongBan_HT;NgayBC;NamBC;ThangBC;PhanLoai_Cd;LaoDong_Cd;Luong_TTV_Ma;Luong_TTV_HT;Luong_CSo;BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN;Tinh_Thue_TNCN;Tinh_Thue_TNCN_HT;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;SoNgayLViec_Thang;MucTamUng_V2;HeSoLuong_V2;TrangThai;SoNgay_KhongLV;MucHuong"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_TongHopId").Width = 10
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_ThongTin_HT").Width = 140
                dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                dgv_name.Columns("cln_PhanLoai_HT").Width = 105
                dgv_name.Columns("cln_PhanLoai_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_PhanLoai_HT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_LaoDong_HT").Width = 60
                dgv_name.Columns("cln_LaoDong_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_LaoDong_HT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_KyBC").Width = 35
                dgv_name.Columns("cln_KyBC").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_KyBC").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_SoLD").Width = 65
                dgv_name.Columns("cln_LCB_Tong_HeSo").Width = 75
                dgv_name.Columns("cln_LCB_HeSo").Width = 70
                dgv_name.Columns("cln_PhCap_ChVu_HeSo").Width = 70
                dgv_name.Columns("cln_PhCap_TrNhiem_HeSo").Width = 60
                dgv_name.Columns("cln_PhCap_DocHai_HeSo").Width = 60
                dgv_name.Columns("cln_Luong_TTV").Width = 75
                dgv_name.Columns("cln_TL_PhuCap").Width = 100
                dgv_name.Columns("cln_PhCap_ThuHut_SoTien").Width = 90
                dgv_name.Columns("cln_PhCap_KhuVuc_SoTien").Width = 90
                dgv_name.Columns("cln_Luong_V1").Width = 110
                dgv_name.Columns("cln_Luong_V2_TamUng").Width = 110
                dgv_name.Columns("cln_TongLuong_TamUng").Width = 110
                dgv_name.Columns("cln_Loai_HDNH_DB_SoLD").Width = 55
                dgv_name.Columns("cln_Loai_HDNH_DB_SoTien").Width = 90
                dgv_name.Columns("cln_Loai_HDNH_PT_SoLD").Width = 55
                dgv_name.Columns("cln_Loai_HDNH_PT_SoTien").Width = 90
                dgv_name.Columns("cln_Tong_TienCong").Width = 100
                dgv_name.Columns("cln_LamDem_SoTienPC").Width = 90
                dgv_name.Columns("cln_Tong_TL_TC_TamChi").Width = 110
                dgv_name.Columns("cln_LdNgViec_CoBHXH_SoLD").Width = 55
                dgv_name.Columns("cln_LdNgViec_CoBHXH_HSLPC").Width = 90
                dgv_name.Columns("cln_LdNgViec_KoBHXH_SoLD").Width = 55
                dgv_name.Columns("cln_LdNgViec_KoBHXH_HSLPC").Width = 90
                dgv_name.Columns("cln_GhiChu_CL").Width = 160
                dgv_name.Columns("cln_TrangThai_HT").Width = 70

                sColumnRight = "SoLD;LCB_Tong_HeSo;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;Luong_TTV;TL_PhuCap;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1;Luong_V2_TamUng;TongLuong_TamUng;Loai_HDNH_DB_SoLD;Loai_HDNH_DB_SoTien;Loai_HDNH_PT_SoLD;Loai_HDNH_PT_SoTien;Tong_TienCong;LamDem_SoTienPC;Tong_TL_TC_TamChi;LdNgViec_CoBHXH_SoLD;LdNgViec_CoBHXH_HSLPC;LdNgViec_KoBHXH_SoLD;LdNgViec_KoBHXH_HSLPC"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.Columns("cln_TongHopId").Visible = False
                dgv_name.Columns("cln_KieuIn").Visible = False

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_Choice").ReadOnly = False
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_ThongTin_HT").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_ThongTin_HT").DefaultCellStyle.BackColor = Color.LightGray

            Case 4 ' Chi lương Thưởng (Tạm ứng thưởng)
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                      '1
                dgv_name.Columns.Add("cln_STT", "STT")                                      '2
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên")                              '3
                dgv_name.Columns.Add("cln_ThongTin_HT", "Chức vụ, vị trí công tác")         '4
                dgv_name.Columns.Add("cln_LCB_Tong_HeSo", "Tổng hệ số" + vbNewLine + vbNewLine)                        '5
                dgv_name.Columns.Add("cln_LCB_HeSo", "HSL cấp bậc" + vbNewLine)                 '6
                dgv_name.Columns.Add("cln_PhCap_ChVu_HeSo", "Chức vụ" + vbNewLine)                           '7
                dgv_name.Columns.Add("cln_PhCap_TrNhiem_HeSo", "Trách nhiệm" + vbNewLine)                        '8
                dgv_name.Columns.Add("cln_PhCap_DocHai_HeSo", "Độc hại" + vbNewLine)                         '9
                dgv_name.Columns.Add("cln_Luong_V1", "Tổng tiền lương V1 (100%)")           '10
                dgv_name.Columns.Add("cln_Luong_V2_TamUng", "Tiền lương V2 (Hệ số 0,8; mức 100%)")      '11
                dgv_name.Columns.Add("cln_Luong_ThucLinh", "Tiền lương thực lĩnh (Tổng cộng V1+V2)")    '12
                dgv_name.Columns.Add("cln_GhiChu_ChiLuong", "Ghi chú")                                  '13
                dgv_name.Columns.Add("cln_SoTK_NHCS", "Số tài khoản NHCS")                              '14
                dgv_name.Columns.Add("cln_Loai_CB_HT", "Phân loại cán bộ")                              '15
                sColumnVisible = "KieuIn;TongHopId;ChiTietId;MaCB;NgaySinh;GioiTinh;ChiNhanh_Cd;ChiNhanh_HT;DonVi_Cd;DonVi_HT;PhongBan_Cd;PhongBan_HT;ChucVu_Cd;ChucVu_HT;ChuyenMon_Cd;ChuyenMon_HT;LoaiQD_Cd;LoaiQD_HT;LoaiHDLD_Cd;LoaiHDLD_HT;LCB_BacLuong_Id;LCB_BacLuong_HT;LCB_NgHuong;PhCap_ChVu_NgHuong;PhCap_TrNhiem_NgHuong;PhCap_DocHai_NgHuong;PhCap_ThuHut_HeSo;PhCap_ThuHut_ST_Goc;PhCap_ThuHut_SoTien;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_HeSo;PhCap_KhuVuc_ST_Goc;PhCap_KhuVuc_SoTien;PhCap_KhuVuc_NgHuong;SoNgay_KhongLV;SoNgayLViec_Thang;SoNgayNghi_TruLuong;SoTienNghi_TruLuong;Luong_TTV_Ma;Luong_TTV;Luong_CSo;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;BHXH_CB;BHXH_CB_SoTien;BHXH_DV;BHXH_DV_SoTien;BHYT_CB;BHYT_CB_SoTien;BHYT_DV;BHYT_DV_SoTien;BHTN_CB;BHTN_CB_SoTien;BHTN_DV;BHTN_DV_SoTien;DPCD_CB;DPCD_CB_SoTien;DPCD_LD_NN;DPCD_LD_NN_SoTien;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tru_CacKhoan_Tong;Tru_CacKhoan_GhiChu;Tinh_Thue_TNCN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;SoNguoi_PhuThuoc;Luong_V1_ConLai;Luong_V1_ThucTra;MucTamUng_V2;HeSoLuong_V2;MucHuong;Loai_CB;TrangThai;TrangThai_HT;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;Loai_HDNH_DB;Loai_HDNH_PT"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next
                dgv_name.Columns("cln_STT").Width = 45
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_HoTen").Width = 180
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_ThongTin_HT").Width = 140
                dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_LCB_Tong_HeSo").Width = 65
                dgv_name.Columns("cln_LCB_HeSo").Width = 61
                dgv_name.Columns("cln_PhCap_ChVu_HeSo").Width = 60
                dgv_name.Columns("cln_PhCap_TrNhiem_HeSo").Width = 60
                dgv_name.Columns("cln_PhCap_DocHai_HeSo").Width = 60
                dgv_name.Columns("cln_Luong_V1").Width = 120
                dgv_name.Columns("cln_Luong_V1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_Luong_V2_TamUng").Width = 120
                dgv_name.Columns("cln_Luong_V2_TamUng").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_Luong_ThucLinh").Width = 130
                dgv_name.Columns("cln_Luong_ThucLinh").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GhiChu_ChiLuong").Width = 150                      '20
                dgv_name.Columns("cln_GhiChu_ChiLuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTK_NHCS").Width = 140                            '21
                dgv_name.Columns("cln_SoTK_NHCS").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_Loai_CB_HT").Width = 130                            '21
                dgv_name.Columns("cln_Loai_CB_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

                sColumnRight = "LCB_Tong_HeSo;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;Luong_V1;Luong_V2_TamUng;Luong_ThucLinh;MucTamUng_V2;HeSoLuong_V2"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.Columns("cln_IdCanBo").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Để column check có thể edit
                dgv_name.Columns("cln_Choice").ReadOnly = False

                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray

            Case 6 ' Truy lĩnh lương
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                      '1
                dgv_name.Columns.Add("cln_STT", "STT")                                      '2
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên")                              '3
                dgv_name.Columns.Add("cln_ThongTin_HT", "Chức vụ, vị trí công tác")         '4
                dgv_name.Columns.Add("cln_PhCap_ThuHut_ST_Goc", "Hệ số cũ" + vbNewLine + vbNewLine)                        '5
                dgv_name.Columns.Add("cln_LCB_HeSo", "Hệ số mới" + vbNewLine + vbNewLine)                 '6
                dgv_name.Columns.Add("cln_PhCap_KhuVuc_ST_Goc", "Chênh lệch hệ số" + vbNewLine)                 '7
                dgv_name.Columns.Add("cln_SoNgayNghi_TruLuong", "Số ngày được truy lĩnh" + vbNewLine)                           '8
                dgv_name.Columns.Add("cln_Luong_V1", "Tiền lương V1" + vbNewLine + vbNewLine)           '9         '(8) = (((Chênh lệch hệ số)*LTTV)/(Số ngày thực tế trong quý))*3*(Số ngày được truy lĩnh)
                dgv_name.Columns.Add("cln_Luong_V2_TamUng", "Tiền lương V2" + vbNewLine + vbNewLine)      '10
                dgv_name.Columns.Add("cln_Luong_V1_ThucTra", "Tổng tiền lương được truy lĩnh" + vbNewLine)      '11
                dgv_name.Columns.Add("cln_BHXH_CB_SoTien", "Trích BHXH (8%)" + vbNewLine)                                           '12
                dgv_name.Columns.Add("cln_BHYT_CB_SoTien", "Trích BHYT (1.5%)" + vbNewLine)                                         '13
                dgv_name.Columns.Add("cln_BHTN_CB_SoTien", "Trích BHTN (1%)" + vbNewLine)                                           '14
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac", "Các khoản trừ khác" + vbNewLine + vbNewLine)                                  '15
                dgv_name.Columns.Add("cln_Tru_Khoan_Khac_GhiChu", "Ghi chú trừ khác" + vbNewLine + vbNewLine + vbNewLine)     '16
                dgv_name.Columns.Add("cln_Tru_CacKhoan_Tong", "Tổng cộng phải trừ" + vbNewLine + vbNewLine)                               '17
                dgv_name.Columns.Add("cln_Luong_ThucLinh", "Thực lĩnh")                                        '18
                dgv_name.Columns.Add("cln_SoTK_NHCS", "Số tài khoản NHCS")                                                          '19
                dgv_name.Columns.Add("cln_GhiChu_ChiLuong", "Ghi chú")                                                          '20

                sColumnVisible = "KieuIn;TongHopId;ChiTietId;MaCB;NgaySinh;GioiTinh;ChiNhanh_Cd;ChiNhanh_HT;DonVi_Cd;DonVi_HT;PhongBan_Cd;PhongBan_HT;ChucVu_Cd;ChucVu_HT;ChuyenMon_Cd;ChuyenMon_HT;LoaiQD_Cd;LoaiQD_HT;LoaiHDLD_Cd;LoaiHDLD_HT;LCB_Tong_HeSo;LCB_BacLuong_Id;LCB_BacLuong_HT;LCB_NgHuong;PhCap_ChVu_HeSo;PhCap_ChVu_NgHuong;PhCap_TrNhiem_HeSo;PhCap_TrNhiem_NgHuong;PhCap_DocHai_HeSo;PhCap_DocHai_NgHuong;PhCap_ThuHut_HeSo;PhCap_ThuHut_SoTien;PhCap_ThuHut_NgHuong;PhCap_KhuVuc_HeSo;PhCap_KhuVuc_SoTien;PhCap_KhuVuc_NgHuong;SoNgay_KhongLV;SoNgayLViec_Thang;SoTienNghi_TruLuong;Luong_TTV_Ma;Luong_TTV;Luong_CSo;Luong_V1_TamUng_K1;Luong_V1_TamUng_K2;BHXH_CB;BHXH_DV;BHXH_DV_SoTien;BHYT_CB;BHYT_DV;BHYT_DV_SoTien;BHTN_CB;BHTN_DV;BHTN_DV_SoTien;DPCD_CB;DPCD_CB_SoTien;DPCD_LD_NN;DPCD_LD_NN_SoTien;Tru_UngHo_Khac;Tru_UngHo_Khac_GhiChu;Tru_CacKhoan_GhiChu;Tinh_Thue_TNCN;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;TNTT_SoTien;TNTT_TyLe;TNTT_SoTien_NopThue;SoNguoi_PhuThuoc;Luong_V1_ConLai;MucTamUng_V2;HeSoLuong_V2;MucHuong;Loai_CB;Loai_CB_HT;TrangThai;TrangThai_HT;IsBHXH;IsBHYT;IsBHTN;IsDPCD;DonGia_LamDem_Gio;LamDem_SoGio;LamDem_SoTienPC;Loai_HDNH;Loai_HDNH_DB;Loai_HDNH_PT"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_STT").Width = 45
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_HoTen").Width = 170
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ThongTin_HT").Width = 110
                dgv_name.Columns("cln_ThongTin_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_PhCap_ThuHut_ST_Goc").Width = 60
                dgv_name.Columns("cln_LCB_HeSo").Width = 60
                dgv_name.Columns("cln_PhCap_KhuVuc_ST_Goc").Width = 60
                dgv_name.Columns("cln_SoNgayNghi_TruLuong").Width = 50
                dgv_name.Columns("cln_Luong_V1").Width = 95
                dgv_name.Columns("cln_Luong_V2_TamUng").Width = 95
                dgv_name.Columns("cln_Luong_V1_ThucTra").Width = 95
                dgv_name.Columns("cln_BHXH_CB_SoTien").Width = 80                       '13
                dgv_name.Columns("cln_BHYT_CB_SoTien").Width = 75                       '14
                dgv_name.Columns("cln_BHTN_CB_SoTien").Width = 75                       '15
                dgv_name.Columns("cln_Tru_Khoan_Khac").Width = 90                       '16
                dgv_name.Columns("cln_Tru_Khoan_Khac_GhiChu").Width = 150                       '17
                dgv_name.Columns("cln_Tru_CacKhoan_Tong").Width = 95                       '18
                dgv_name.Columns("cln_Luong_ThucLinh").Width = 95                    '19
                dgv_name.Columns("cln_Luong_ThucLinh").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTK_NHCS").Width = 130                            '21
                dgv_name.Columns("cln_SoTK_NHCS").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GhiChu_ChiLuong").Width = 200                      '20
                dgv_name.Columns("cln_GhiChu_ChiLuong").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                sColumnRight = "PhCap_ThuHut_ST_Goc;LCB_HeSo;PhCap_KhuVuc_ST_Goc;SoNgayNghi_TruLuong;Luong_V1;Luong_V2_TamUng;Luong_V1_ThucTra;BHXH_CB_SoTien;BHYT_CB_SoTien;BHTN_CB_SoTien;Tru_Khoan_Khac;Tru_CacKhoan_Tong;Luong_ThucLinh"
                FormatStyleAlignment_ColumnDataGridView(dgv_name, sColumnRight, ";", 3)

                dgv_name.Columns("cln_IdCanBo").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Để column check có thể edit
                'Để column check có thể edit
                sColumnTMP = "Choice;SoNgayNghi_TruLuong;Tru_Khoan_Khac;Tru_Khoan_Khac_GhiChu"
                ARL_Cols = Globals.Splip_Strings(sColumnTMP, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).ReadOnly = False
                        dgv_name.Columns("cln_" + _Value).DefaultCellStyle.BackColor = Color.White
                    End If
                Next
                dgv_name.Columns("cln_Choice").ReadOnly = False
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray
        End Select
    End Sub

    Public Sub Fill_Tree(ByRef pTv_Name As TreeView, ByVal pNamBC As Integer, ByVal pThangBC As Integer, Optional ByVal pNotAll As Boolean = False)
        Dim _HeThongBLL As HeThongBLL = New HeThongBLL()
        Dim db_parent As DataTable
        Dim i As Integer
        Dim parentNode As TreeNode
        Dim childNode As TreeNode
        Dim grandchildNode As TreeNode
        db_parent = _HeThongBLL.GetDanhMuc_TreeViews(1, Cap_Nd, pNamBC, pThangBC, DONVI)
        If Not (db_parent Is Nothing) Then
            If db_parent.Rows.Count > 0 Then
                For i = 0 To db_parent.Rows.Count - 1
                    If Globals.Group = "S03" Or Globals.Group = "S04" Then
                        If db_parent.Rows(i)("MaSo_01").ToString() = "000199" Or db_parent.Rows(i)("MaSo_01").ToString() = "000100" Then
                            If db_parent.Rows(i)("Cap_DM").ToString() = "1" Then
                                parentNode = pTv_Name.Nodes.Add(db_parent.Rows(i)("Ma_So").ToString(), db_parent.Rows(i)("Ten_Goi").ToString() + " " + db_parent.Rows(i)("GhiChu").ToString())
                                parentNode.Tag = db_parent.Rows(i)("Cap_DM").ToString() + "_" + db_parent.Rows(i)("Khoa").ToString() + "_" + db_parent.Rows(i)("Ma_So").ToString()
                            ElseIf db_parent.Rows(i)("Cap_DM").ToString() = "2" Then
                                childNode = parentNode.Nodes.Add(db_parent.Rows(i)("Ma_So").ToString(), db_parent.Rows(i)("Ten_Goi").ToString() + " " + db_parent.Rows(i)("GhiChu").ToString())
                                childNode.Tag = db_parent.Rows(i)("Cap_DM").ToString() + "_" + db_parent.Rows(i)("Khoa").ToString() + "_" + db_parent.Rows(i)("Ma_So").ToString()
                            ElseIf db_parent.Rows(i)("Cap_DM").ToString() = "3" Then
                                grandchildNode = childNode.Nodes.Add(db_parent.Rows(i)("Ma_So").ToString(), db_parent.Rows(i)("Ten_Goi").ToString() + " " + db_parent.Rows(i)("GhiChu").ToString())
                                grandchildNode.Tag = db_parent.Rows(i)("Cap_DM").ToString() + "_" + db_parent.Rows(i)("Khoa").ToString() + "_" + db_parent.Rows(i)("Ma_So").ToString()
                            End If
                        End If
                    Else
                        If db_parent.Rows(i)("Cap_DM").ToString() = "1" Then
                            parentNode = pTv_Name.Nodes.Add(db_parent.Rows(i)("Ma_So").ToString(), db_parent.Rows(i)("Ten_Goi").ToString() + " " + db_parent.Rows(i)("GhiChu").ToString())
                            parentNode.Tag = db_parent.Rows(i)("Cap_DM").ToString() + "_" + db_parent.Rows(i)("Khoa").ToString() + "_" + db_parent.Rows(i)("Ma_So").ToString()
                        ElseIf db_parent.Rows(i)("Cap_DM").ToString() = "2" Then
                            childNode = parentNode.Nodes.Add(db_parent.Rows(i)("Ma_So").ToString(), db_parent.Rows(i)("Ten_Goi").ToString() + " " + db_parent.Rows(i)("GhiChu").ToString())
                            childNode.Tag = db_parent.Rows(i)("Cap_DM").ToString() + "_" + db_parent.Rows(i)("Khoa").ToString() + "_" + db_parent.Rows(i)("Ma_So").ToString()
                        ElseIf db_parent.Rows(i)("Cap_DM").ToString() = "3" Then
                            grandchildNode = childNode.Nodes.Add(db_parent.Rows(i)("Ma_So").ToString(), db_parent.Rows(i)("Ten_Goi").ToString() + " " + db_parent.Rows(i)("GhiChu").ToString())
                            grandchildNode.Tag = db_parent.Rows(i)("Cap_DM").ToString() + "_" + db_parent.Rows(i)("Khoa").ToString() + "_" + db_parent.Rows(i)("Ma_So").ToString()
                        End If
                    End If
                Next
            End If
        End If
        If Not pNotAll AndAlso pTv_Name.Nodes.Count > 0 Then pTv_Name.Nodes(0).ExpandAll()
    End Sub

    Public Function GetTinhThue_TNCN(ByVal pTNTT_SoTien As Double, ByVal pNgayTinh As DateTime) As Double
        Dim dResult As Double = 0

        If pNgayTinh >= DateTimeUtil.StringToDateTime("01/01/2019", "dd/MM/yyyy") Then
            If pTNTT_SoTien > 0 And pTNTT_SoTien <= 5000000 Then
                dResult = pTNTT_SoTien * 0.05
            ElseIf pTNTT_SoTien > 5000000 And pTNTT_SoTien <= 10000000 Then
                dResult = pTNTT_SoTien * 0.1 - 250000
            ElseIf pTNTT_SoTien > 10000000 And pTNTT_SoTien <= 18000000 Then
                dResult = pTNTT_SoTien * 0.15 - 750000
            ElseIf pTNTT_SoTien > 18000000 And pTNTT_SoTien <= 32000000 Then
                dResult = pTNTT_SoTien * 0.2 - 1650000
            ElseIf pTNTT_SoTien > 32000000 And pTNTT_SoTien <= 52000000 Then
                dResult = pTNTT_SoTien * 0.25 - 3250000
            ElseIf pTNTT_SoTien > 52000000 And pTNTT_SoTien <= 80000000 Then
                dResult = pTNTT_SoTien * 0.3 - 5850000
            ElseIf pTNTT_SoTien > 80000000 Then
                dResult = pTNTT_SoTien * 0.35 - 9850000
            Else : dResult = 0
            End If
        End If
        Return Math.Round(dResult)
    End Function

    Public Sub SetStyleRowGrid(ByVal iRow As Integer, ByVal dgv_Name As DataGridView, ByVal iKieuIn As Byte, ByVal iValueNote As Byte)
        dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.Navy
        If iKieuIn = 0 Or iKieuIn = 1 Or iKieuIn = Nothing Then
            dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Bold)
            dgv_Name.Rows(iRow).DefaultCellStyle.BackColor = System.Drawing.Color.Azure
            dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = System.Drawing.Color.MediumBlue
        Else
            If iKieuIn = 4 Then
                dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Italic)
                dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.OrangeRed
            Else
                If iKieuIn = 2 Then
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Italic)
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Bold)
                ElseIf iValueNote = 1 Then
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Regular)
                    dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.Maroon
                ElseIf iValueNote = 2 Then
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Regular)
                    dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.Green
                Else
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Regular)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Tìm kiếm lấy danh sách Bảng kê Chi lương chi tiết theo các điều kiện truyền vào
    ''' </summary>
    ''' <param name="_LoaiChi_Cd">Mã Loại chi lương. Select ValueKey,ValueDesc From SysListValue Where ListKey='LOAICHILUONG'</param>
    ''' <param name="_ThoiDiem">Thời điểm báo cáo. Là căn cứ lấy DS Nhân sự theo QĐ nhân sự thời điểm đó. Đinh dang dd/MM/yyyy</param>
    ''' <param name="_NamBC">Năm báo cáo</param>
    ''' <param name="_ThangBC">Tháng báo cáo</param>
    ''' <param name="_KyBC">Kỳ báo cáo. Nếu là lương bổ sung Kỳ nhận giá trị 0. Không có điều kiện thì Kỳ 99</param>
    ''' <param name="_DonVi_Cd">Đơn vị báo cáo</param>
    ''' <param name="_PhongBan_Cd">Phòng ban báo cáo</param>
    ''' <param name="_IsAll">1 Lấy tất cả đơn vị trực thuộc của Đơn vị @_DonVi_Cd. 0 - Lấy duy nhất theo POS</param>
    ''' <param name="_FlagCall">1 - Tạo mới Chi lương</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChiLuongSearch(ByVal _LoaiChi_Cd As String, ByVal _ThoiDiem As String, ByVal _NamBC As Integer, ByVal _ThangBC As Byte, ByVal _KyBC As Byte, ByVal _DonVi_Cd As String, ByVal _PhongBan_Cd As String, ByVal _IsAll As Byte, ByVal _FlagCall As Byte) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("ChiLuong_GetSearch", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThoiDiem", _ThoiDiem, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NamBC", _NamBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThangBC", _ThangBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_KyBC", _KyBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonVi_Cd", _DonVi_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhongBan_Cd", _PhongBan_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_FlagCall", _FlagCall, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IsAll", _IsAll, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "DsChiLuong_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("DsChiLuong_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm lấy danh sách Chi lương Bảng kê chi tiết theo điều kiện
    ''' </summary>
    ''' <param name="_TongHopId"></param>
    ''' <param name="_ChiTietId"></param>
    ''' <param name="_ThoiDiem">Ngày báo cáo. Định dạng dd/MM/yyyy</param>
    ''' <param name="_NamBC"></param>
    ''' <param name="_ThangBC"></param>
    ''' <param name="_KyBC">Nếu lấy tất truyền vào là 99</param>
    ''' <param name="_DonVi_Cd"></param>
    ''' <param name="_PhongBan_Cd"></param>
    ''' <param name="_IsAll">1 - Lấy tất cả chi nhánh; 0 - Lấy duy nhất đơn vị</param>
    ''' <param name="_FlagCall"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChiLuong_BangKeCT_GetSearch(ByVal _TongHopId As Long, ByVal _ChiTietId As Long, ByVal _ThoiDiem As String, ByVal _NamBC As Integer, ByVal _ThangBC As Byte, ByVal _KyBC As Byte, ByVal _DonVi_Cd As String, ByVal _PhongBan_Cd As String, ByVal _IsAll As Byte, ByVal _FlagCall As Byte) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("ChiLuong_BangKeCT_GetSearch", connection)

            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TongHopId", _TongHopId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChiTietId", _ChiTietId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThoiDiem", _ThoiDiem, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NamBC", _NamBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThangBC", _ThangBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_KyBC", _KyBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonVi_Cd", _DonVi_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhongBan_Cd", _PhongBan_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_FlagCall", _FlagCall, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IsAll", _IsAll, ParameterDirection.Input))

            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "DsChiLuong_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("DsChiLuong_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm lấy danh sách Chi lương Tổng hợp theo điều kiện truyền vào
    ''' </summary>
    ''' <param name="pDonViCd">Đơn vị cần lấy</param>
    ''' <param name="pPhongBanCd"></param>
    ''' <param name="pNgayBC">Ngày báo cáo. Định dạng dd/MM/yyyy</param>
    ''' <param name="pNamBC"></param>
    ''' <param name="pThangBC"></param>
    ''' <param name="pKyBC">Nếu lấy tất truyền vào là 99</param>
    ''' <param name="pPhanLoai_Cd"></param>
    ''' <param name="pLaoDong_Cd"></param>
    ''' <param name="pTongHopId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChiLuong_TongHop_GetSearch(ByVal pDonViCd As String, ByVal pPhongBanCd As String, ByVal pNgayBC As String, ByVal pNamBC As Integer, ByVal pThangBC As Integer, ByVal pKyBC As Byte, ByVal pPhanLoai_Cd As String, ByVal pLaoDong_Cd As String, ByVal pTongHopId As Long) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim _command As SqlCommand = New SqlCommand("ChiLuong_TongHop_GetSearch", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViCd", pDonViCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhongBanCd", pPhongBanCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNgayBC", pNgayBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNamBC", pNamBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThangBC", pThangBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pKyBC", pKyBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhanLoai_Cd", pPhanLoai_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pLaoDong_Cd", pLaoDong_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pTongHopId", pTongHopId, ParameterDirection.Input))

            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "ChiLuong_TongHop_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("ChiLuong_TongHop_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm lấy Số liệu Báo cáo Tình hình thực hiện Lao động - Tiền lương 01/TL hàng tháng
    ''' </summary>
    ''' <param name="pDonViCd">Mã đơn vị</param>
    ''' <param name="pPhongBanCd">Mã phòng ban</param>
    ''' <param name="pNgayBC">Ngày báo cáo. Giá trị định dạng: dd/MM/yyyy</param>
    ''' <param name="pNamBC">Năm báo cáo</param>
    ''' <param name="pThangBC">Tháng báo cáo</param>
    ''' <param name="pKyBC">Kỳ báo cáo. Nếu lấy không tính điều kiện Kỳ BC thì truyền: 99</param>
    ''' <param name="pPhanLoai_Cd"></param>
    ''' <param name="pLaoDong_Cd"></param>
    ''' <param name="pTongHopId"></param>
    ''' <param name="pFlagCall">1 - Lấy chi tiết lên Form; 2 - Lấy tổng hợp ra báo cáo 01/TL</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChiLuong_TongHop_Thang_BC01TL(ByVal pChiNhanhCd As String, ByVal pDonViCd As String, ByVal pPhongBanCd As String, ByVal pNgayBC As String, ByVal pNamBC As Integer, ByVal pThangBC As Integer, ByVal pKyBC As Byte, ByVal pPhanLoai_Cd As String, ByVal pLaoDong_Cd As String, ByVal pTongHopId As Long, ByVal pFlagCall As Byte) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim _command As SqlCommand = New SqlCommand("ChiLuong_TongHop_Thang_BC01TL", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pChiNhanhCd", pChiNhanhCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViCd", pDonViCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhongBanCd", pPhongBanCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNgayBC", pNgayBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNamBC", pNamBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThangBC", pThangBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pKyBC", pKyBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhanLoai_Cd", pPhanLoai_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pLaoDong_Cd", pLaoDong_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pTongHopId", pTongHopId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pFlagCall", pFlagCall, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "ChiLuong_TongHop_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("ChiLuong_TongHop_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    Public Function Insert_Update_ChiLuong_TongHop(ByVal _ChiLuong_TongHop As ChiLuong_TongHop) As Long
        Dim _Ret As Long = 0
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "ChiLuong_TongHop_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TongHopId", _ChiLuong_TongHop.TongHopId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonVi_CL_Cd", _ChiLuong_TongHop.DonVi_CL_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhongBan_CL_Cd", _ChiLuong_TongHop.PhongBan_CL_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgayBC", _ChiLuong_TongHop.NgayBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NamBC", _ChiLuong_TongHop.NamBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThangBC", _ChiLuong_TongHop.ThangBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_KyBC", _ChiLuong_TongHop.KyBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhanLoai_Cd", _ChiLuong_TongHop.PhanLoai_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LaoDong_Cd", _ChiLuong_TongHop.LaoDong_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_TTV_Ma", _ChiLuong_TongHop.Luong_TTV_Ma, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_TTV", _ChiLuong_TongHop.Luong_TTV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_CSo", _ChiLuong_TongHop.Luong_CSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHXH_CB", _ChiLuong_TongHop.BHXH_CB, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHXH_DV", _ChiLuong_TongHop.BHXH_DV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHYT_CB", _ChiLuong_TongHop.BHYT_CB, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHYT_DV", _ChiLuong_TongHop.BHYT_DV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHTN_CB", _ChiLuong_TongHop.BHTN_CB, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHTN_DV", _ChiLuong_TongHop.BHTN_DV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DPCD_CB", _ChiLuong_TongHop.DPCD_CB, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DPCD_LD_NN", _ChiLuong_TongHop.DPCD_LD_NN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tinh_Thue_TNCN", _ChiLuong_TongHop.Tinh_Thue_TNCN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TNTT_Muc_BanThan", _ChiLuong_TongHop.TNTT_Muc_BanThan, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TNTT_Muc_PhuThuoc", _ChiLuong_TongHop.TNTT_Muc_PhuThuoc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoNgayLViec_Thang", _ChiLuong_TongHop.SoNgayLViec_Thang, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_MucTamUng_V2", _ChiLuong_TongHop.MucTamUng_V2, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_HeSoLuong_V2", _ChiLuong_TongHop.HeSoLuong_V2, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GhiChu_CL", _ChiLuong_TongHop.GhiChu_CL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TrangThai", _ChiLuong_TongHop.TrangThai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CreatedBy", _ChiLuong_TongHop.CreatedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _ChiLuong_TongHop.ModifiedBy, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _ChiLuong_TongHop.IdOutPut = CType(_command.Parameters("@_IdOutPut").Value.ToString(), Long)
                        _Ret = _ChiLuong_TongHop.IdOutPut
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Cập nhật chi lương tổng hợp (ChiLuong_TongHop)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = 0
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    Public Function Insert_Update_ChiLuong_BangKeCT(ByVal _ChiLuong_BangKeCT As ChiLuong_BangKeCT) As Long
        Dim _Ret As Long = 0
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "ChiLuong_BangKeCT_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.BigInt).Direction = ParameterDirection.Output

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChiTietId", _ChiLuong_BangKeCT.ChiTietId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TongHopId", _ChiLuong_BangKeCT.TongHopId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo", _ChiLuong_BangKeCT.IdCanBo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_MaCB", _ChiLuong_BangKeCT.MaCB, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_HoTen", _ChiLuong_BangKeCT.HoTen, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThongTin_HT", _ChiLuong_BangKeCT.ThongTin_HT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoTK_NHCS", _ChiLuong_BangKeCT.SoTK_NHCS, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChiNhanh_Cd", _ChiLuong_BangKeCT.ChiNhanh_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonVi_Cd", _ChiLuong_BangKeCT.DonVi_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhongBan_Cd", _ChiLuong_BangKeCT.PhongBan_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChucVu_Cd", _ChiLuong_BangKeCT.ChucVu_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChuyenMon_Cd", _ChiLuong_BangKeCT.ChuyenMon_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LoaiQD_Cd", _ChiLuong_BangKeCT.LoaiQD_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LoaiHDLD_Cd", _ChiLuong_BangKeCT.LoaiHDLD_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LCB_BacLuong_Id", _ChiLuong_BangKeCT.LCB_BacLuong_Id, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LCB_HeSo", _ChiLuong_BangKeCT.LCB_HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LCB_NgHuong", _ChiLuong_BangKeCT.LCB_NgHuong, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_ChVu_HeSo", _ChiLuong_BangKeCT.PhCap_ChVu_HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_ChVu_NgHuong", _ChiLuong_BangKeCT.PhCap_ChVu_NgHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_TrNhiem_HeSo", _ChiLuong_BangKeCT.PhCap_TrNhiem_HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_TrNhiem_NgHuong", _ChiLuong_BangKeCT.PhCap_TrNhiem_NgHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_DocHai_HeSo", _ChiLuong_BangKeCT.PhCap_DocHai_HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_DocHai_NgHuong", _ChiLuong_BangKeCT.PhCap_DocHai_NgHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_ThuHut_HeSo", _ChiLuong_BangKeCT.PhCap_ThuHut_HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_ThuHut_ST_Goc", _ChiLuong_BangKeCT.PhCap_ThuHut_ST_Goc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_ThuHut_SoTien", _ChiLuong_BangKeCT.PhCap_ThuHut_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_ThuHut_NgHuong", _ChiLuong_BangKeCT.PhCap_ThuHut_NgHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_KhuVuc_HeSo", _ChiLuong_BangKeCT.PhCap_KhuVuc_HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_KhuVuc_ST_Goc", _ChiLuong_BangKeCT.PhCap_KhuVuc_ST_Goc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_KhuVuc_SoTien", _ChiLuong_BangKeCT.PhCap_KhuVuc_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhCap_KhuVuc_NgHuong", _ChiLuong_BangKeCT.PhCap_KhuVuc_NgHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoNgay_KhongLV", _ChiLuong_BangKeCT.SoNgay_KhongLV, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoNgayNghi_TruLuong", _ChiLuong_BangKeCT.SoNgayNghi_TruLuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoTienNghi_TruLuong", _ChiLuong_BangKeCT.SoTienNghi_TruLuong, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_V1", _ChiLuong_BangKeCT.Luong_V1, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_V1_TamUng_K1", _ChiLuong_BangKeCT.Luong_V1_TamUng_K1, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_V1_TamUng_K2", _ChiLuong_BangKeCT.Luong_V1_TamUng_K2, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHXH_CB_SoTien", _ChiLuong_BangKeCT.BHXH_CB_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHXH_DV_SoTien", _ChiLuong_BangKeCT.BHXH_DV_SoTien, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHYT_CB_SoTien", _ChiLuong_BangKeCT.BHYT_CB_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHYT_DV_SoTien", _ChiLuong_BangKeCT.BHYT_DV_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHTN_CB_SoTien", _ChiLuong_BangKeCT.BHTN_CB_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHTN_DV_SoTien", _ChiLuong_BangKeCT.BHTN_DV_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DPCD_CB_SoTien", _ChiLuong_BangKeCT.DPCD_CB_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DPCD_LD_NN", _ChiLuong_BangKeCT.DPCD_LD_NN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DPCD_LD_NN_SoTien", _ChiLuong_BangKeCT.DPCD_LD_NN_SoTien, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tru_Khoan_Khac", _ChiLuong_BangKeCT.Tru_Khoan_Khac, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tru_Khoan_Khac_GhiChu", _ChiLuong_BangKeCT.Tru_Khoan_Khac_GhiChu, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tru_UngHo_Khac", _ChiLuong_BangKeCT.Tru_UngHo_Khac, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tru_UngHo_Khac_GhiChu", _ChiLuong_BangKeCT.Tru_UngHo_Khac_GhiChu, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tru_CacKhoan_Tong", _ChiLuong_BangKeCT.Tru_CacKhoan_Tong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tru_CacKhoan_GhiChu", _ChiLuong_BangKeCT.Tru_CacKhoan_GhiChu, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoNguoi_PhuThuoc", _ChiLuong_BangKeCT.SoNguoi_PhuThuoc, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TNTT_SoTien", _ChiLuong_BangKeCT.TNTT_SoTien, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TNTT_TyLe", _ChiLuong_BangKeCT.TNTT_TyLe, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TNTT_SoTien_NopThue", _ChiLuong_BangKeCT.TNTT_SoTien_NopThue, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_V1_ConLai", _ChiLuong_BangKeCT.Luong_V1_ConLai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_V1_ThucTra", _ChiLuong_BangKeCT.Luong_V1_ThucTra, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_MucHuong", _ChiLuong_BangKeCT.MucHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_MucTamUng_V2", _ChiLuong_BangKeCT.MucTamUng_V2, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_V2_TamUng", _ChiLuong_BangKeCT.Luong_V2_TamUng, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Luong_ThucLinh", _ChiLuong_BangKeCT.Luong_ThucLinh, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GhiChu_ChiLuong", _ChiLuong_BangKeCT.GhiChu_ChiLuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Loai_CB", _ChiLuong_BangKeCT.Loai_CB, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TrangThai", _ChiLuong_BangKeCT.TrangThai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IsBHXH", _ChiLuong_BangKeCT.IsBHXH, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IsBHYT", _ChiLuong_BangKeCT.IsBHYT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IsBHTN", _ChiLuong_BangKeCT.IsBHTN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IsDPCD", _ChiLuong_BangKeCT.IsDPCD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonGia_LamDem_Gio", _ChiLuong_BangKeCT.DonGia_LamDem_Gio, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LamDem_SoGio", _ChiLuong_BangKeCT.LamDem_SoGio, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LamDem_SoTienPC", _ChiLuong_BangKeCT.LamDem_SoTienPC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Loai_HDNH", _ChiLuong_BangKeCT.Loai_HDNH, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CreatedBy", _ChiLuong_BangKeCT.CreatedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _ChiLuong_BangKeCT.ModifiedBy, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _ChiLuong_BangKeCT.IdOutPut = CType(_command.Parameters("@_IdOutPut").Value.ToString(), Long)
                        _Ret = _ChiLuong_BangKeCT.IdOutPut
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Cập nhật chi lương chi tiết (ChiLuong_BangKeCT)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = 0
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    ''' <summary>
    ''' Hàm thực hiện xóa bản ghi Chi lương Tổng hợp
    ''' </summary>
    ''' <param name="_TongHopId">Chỉ số xác định bản ghi. Với trường hợp _FlagDelete=3 thì _TongHopId=0</param>
    ''' <param name="_ModifiedBy">Người thực hiện</param>
    ''' <param name="_PhanLoaiCd">_FlagDelete=3 Thì truyền Phân loại chi lương</param>
    ''' <param name="_LaoDong_Cd">_FlagDelete=3 Thì truyền Loại cán bộ</param>
    ''' <param name="_DonViCd">_FlagDelete=3 Thì truyền Mã đơn vị</param>
    ''' <param name="_NamBC">_FlagDelete=3 Thì truyền Năm báo cáo</param>
    ''' <param name="_ThangBC">_FlagDelete=3 Thì truyền Tháng báo cáo</param>
    ''' <param name="_KyBC">_FlagDelete=3 Thì truyền Kỳ báo cáo</param>
    ''' <param name="_NgayBC">_FlagDelete=3 Thì truyền Ngày báo cáo định dạng dd/MM/yyyy (Hoặc có thể để trống sẽ bỏ qua đk NgayBC khi xóa)</param>
    ''' <param name="_FlagDelete">Cờ xác định thực hiện: 1 - Xóa hẳn dữ liệu theo _TongHopId; 2 - Đánh dấu xóa bản ghi theo _TongHopId; 
    '''                                                  3 Xóa hẳn theo điều kiện khác
    '''                                                  4 Xóa hẳn theo điều kiện khác và xóa luôn chi lương chi tiết
    '''                                                  5: Update trạng thái Chi lương tổng hợp và chi lương chi tiết
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete_UpdateStatus_ChiLuong_TongHop(ByVal _TongHopId As Long, ByVal _TrangThai As Byte, ByVal _ModifiedBy As String, ByVal _PhanLoaiCd As String, ByVal _LaoDong_Cd As String, ByVal _DonViCd As String, ByVal _PhongBanCd As String, ByVal _NamBC As Integer, _ThangBC As Integer, ByVal _KyBC As Byte, ByVal _NgayBC As String, _FlagDelete As Byte) As Boolean
        Dim _Ret As Boolean = False
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "ChiLuong_TongHop_Delete_UpdateStatus"
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TongHopId", _TongHopId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TrangThai", _TrangThai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _ModifiedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhanLoaiCd", _PhanLoaiCd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LaoDong_Cd", _LaoDong_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonViCd", _DonViCd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhongBan_CL_Cd", _PhongBanCd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NamBC", _NamBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ThangBC", _ThangBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_KyBC", _KyBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgayBC", _NgayBC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_FlagDelete", _FlagDelete, ParameterDirection.Input))
                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Xóa dữ liệu Chi lương tổng hợp (ChiLuong_TongHop)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = False
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    ''' <summary>
    ''' Hàm Xóa/Đánh dấu xóa bảng dữ liệu Chi lương chi tiết (ChiLuong_BangKeCT)
    ''' </summary>
    ''' <param name="_ChiTietId">Chỉ số xác định Id Chi lương chi tiết</param>
    ''' <param name="_TongHopId">Chỉ số xác định Id Chi lương tổng hợp</param>
    ''' <param name="_ModifiedBy">Người thực hiện</param>
    ''' <param name="_FlagDelete">Cờ xác định thực hiện: 1 - Xóa hẳn dữ liệu; 2 - Đánh dấu xóa bản ghi; 3 - Xóa hẳn dữ liệu theo TongHopId; 4 - Đánh dấu xóa bản ghi theo TongHopId</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete_ChiLuong_BangKeCT(ByVal _ChiTietId As Long, ByVal _TongHopId As Long, ByVal _ModifiedBy As String, _FlagDelete As Byte) As Boolean
        Dim _Ret As Boolean = False

        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "ChiLuong_BangKeCT_Delete"
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChiTietId", _ChiTietId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TongHopId", _TongHopId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _ModifiedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_FlagDelete", _FlagDelete, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Xóa dữ liệu Chi lương tổng hợp (ChiLuong_TongHop)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = False
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    Public Function ExportExcel_TH_LuongThang_01TL(ByVal pDonVi_TK_Cd As String, ByVal pProvinceName As String, ByVal pBranchName As String, ByVal pNamBC As Integer, ByVal pThangBC As Integer, pLoaiCLCd As String, pKyBC As Byte, pLaoDong As String, db_source As DataTable) As String
        Dim _ResultFile As String = ""
        Try
            Dim _HeThongBLL As HeThongBLL = New HeThongBLL()
            Dim _NgayHeThong As DateTime = Globals.GetDateTime_ForServerDB
            Dim iRowStart As Integer = 0
            Dim _LapBieu_1 = "", _LapBieu_2 = "", _HanhChinhTC_1 = "", _HanhChinhTC_2 = "", _KeToanNQ_1 = "", _KeToanNQ_2 = "", _GiamDoc_1 = "", _GiamDoc_2 = ""
            Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
            Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
            Dim sPath_Export As String = "C:\Temp\"
            Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
            sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")
            sFileExcelTemplate = sPathExcelTemplate + "BC_LuongThang_01TL.xlsx"
            sFileNameEx = "BC_LuongThang_01TL_" + "Thang_" + pThangBC.ToString("D2") + pNamBC.ToString("D4") + "_" + _NgayHeThong.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
            sColNameEnd = "AA"
            _LapBieu_1 = "B"
            _LapBieu_2 = "D"
            _HanhChinhTC_1 = "H"
            _HanhChinhTC_2 = "L"
            _KeToanNQ_1 = "O"
            _KeToanNQ_2 = "S"
            _GiamDoc_1 = "V"
            _GiamDoc_2 = "AA"
            iRowStart = 14

            If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
                Using excelEngine As ExcelEngine = New ExcelEngine()
                    Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                    Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                    Dim worksheet As IWorksheet = workbook.Worksheets(0)
                    worksheet.Range("A1").Text = pProvinceName.ToUpper()
                    worksheet.Range("A2").Text = pBranchName.ToUpper()
                    If pThangBC <= 0 Or pThangBC > 12 Then
                        worksheet.Range("A5").Text = "Năm " + pNamBC.ToString("D4")
                    Else
                        worksheet.Range("A5").Text = "Tháng " + pThangBC.ToString("D2") + " năm " + pNamBC.ToString("D4")
                    End If

                    'Create Template Marker Processor
                    Dim marker As ITemplateMarkersProcessor = workbook.CreateTemplateMarkersProcessor()
                    Dim columnNames(db_source.Columns.Count) As String
                    Dim i As Integer = 0
                    For Each column As DataColumn In db_source.Columns
                        columnNames(i) = column.ColumnName
                        i += 1
                    Next

                    For Each c As DataColumn In db_source.Columns
                        Dim columnName As String = c.ColumnName
                        Dim columnData As EnumerableRowCollection(Of Object)
                        If Globals.IsNumeric(c) Then
                            columnData = db_source.AsEnumerable().[Select](Function(r) If(r.Field(Of Object)(columnName), 0))
                        ElseIf c.DataType Is GetType(Date) Then
                            columnData = db_source.AsEnumerable.[Select](Function(r) If(r.Field(Of Object)(columnName), ""))
                        Else
                            columnData = db_source.AsEnumerable.[Select](Function(r) If(("'" & r.Field(Of Object)(columnName)), ""))
                        End If

                        Dim paramDataArray As Object() = columnData.ToArray
                        marker.AddVariable(columnName, paramDataArray)
                    Next
                    marker.ApplyMarkers()
                    workbook.Version = ExcelVersion.Excel2013

                    'Căn chỉnh bôi đậm, nghiêng các dòng bản ghi và điền dữ liệu vào một số CELL
                    Dim iRowA As Integer = 0
                    If Not (db_source Is Nothing) Then
                        If (db_source.Rows.Count > 0) Then
                            For iTT As Integer = 0 To db_source.Rows.Count - 1
                                iRowA = (iRowStart + iTT)
                                Dim iKieuIn As Integer = CType(db_source.Rows(iTT)("KieuIn"), Integer)
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
                            'Xuất đoạn cuối: Lập biểu/Kiểm soát/Giám đốc
                            Dim db_sysvar As DataTable = New DataTable()
                            db_sysvar = _HeThongBLL.GetListSysVarSearch(0, 0, pDonVi_TK_Cd, "", "", "", 1)
                            If Not (db_sysvar Is Nothing) Then
                                If (db_sysvar.Rows.Count > 0) Then
                                    iRowA = iRowA + 1
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).Text = db_sysvar.Rows(0)("DiaBan").ToString().Trim() + ", ngày " + _NgayHeThong.Day.ToString("D2") + " tháng " + _NgayHeThong.Month.ToString("D2") + " năm " + _NgayHeThong.Year.ToString("D4") + "        "
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).RowHeight = 23

                                    iRowA = iRowA + 1
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Text = "NGƯỜI LẬP BIỂU"
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range(_HanhChinhTC_1 + iRowA.ToString() + ":" + _HanhChinhTC_2 + iRowA.ToString()).Text = "TP. HÀNH CHÍNH-TỔ CHỨC"
                                    worksheet.Range(_HanhChinhTC_1 + iRowA.ToString() + ":" + _HanhChinhTC_2 + iRowA.ToString()).Merge(True)

                                    worksheet.Range(_KeToanNQ_1 + iRowA.ToString() + ":" + _KeToanNQ_2 + iRowA.ToString()).Text = "TP. KẾ TOÁN-NGÂN QUỸ"
                                    worksheet.Range(_KeToanNQ_1 + iRowA.ToString() + ":" + _KeToanNQ_2 + iRowA.ToString()).Merge(True)

                                    If (IsNothing(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString()) Or String.IsNullOrEmpty(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString())) Then
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = IIf((DONVI = "000100" Or DONVI = "000199"), "TỔNG GIÁM ĐỐC", "GIÁM ĐỐC")
                                    Else
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = IIf((DONVI = "000100" Or DONVI = "000199"), "KT. TỔNG GIÁM ĐỐC" + vbNewLine + "PHÓ TỔNG GIÁM ĐỐC", "KT. GIÁM ĐỐC" + vbNewLine + "PHÓ GIÁM ĐỐC")
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).RowHeight = 27
                                    End If
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter

                                    iRowA = iRowA + 1
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Text = "(Ký ghi rõ họ tên)"
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range(_HanhChinhTC_1 + iRowA.ToString() + ":" + _HanhChinhTC_2 + iRowA.ToString()).Text = "(Ký ghi rõ họ tên)"
                                    worksheet.Range(_HanhChinhTC_1 + iRowA.ToString() + ":" + _HanhChinhTC_2 + iRowA.ToString()).Merge(True)

                                    worksheet.Range(_KeToanNQ_1 + iRowA.ToString() + ":" + _KeToanNQ_2 + iRowA.ToString()).Text = "(Ký ghi rõ họ tên)"
                                    worksheet.Range(_KeToanNQ_1 + iRowA.ToString() + ":" + _KeToanNQ_2 + iRowA.ToString()).Merge(True)

                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = "(Ký tên, đóng dấu)"
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 10
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter

                                    'Điền họ tên của: Lập biểu; Kiểm soát; Lãnh đạo
                                    iRowA = iRowA + 6
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Text = MainForm.lblNguoiSuDung.Text
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Merge(True)

                                    '
                                    worksheet.Range(_HanhChinhTC_1 + iRowA.ToString() + ":" + _HanhChinhTC_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("TruongHCTC").ToString()
                                    worksheet.Range(_HanhChinhTC_1 + iRowA.ToString() + ":" + _HanhChinhTC_2 + iRowA.ToString()).Merge(True)

                                    worksheet.Range(_KeToanNQ_1 + iRowA.ToString() + ":" + _KeToanNQ_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("KeToanTruong").ToString()
                                    worksheet.Range(_KeToanNQ_1 + iRowA.ToString() + ":" + _KeToanNQ_2 + iRowA.ToString()).Merge(True)

                                    If (IsNothing(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString()) Or String.IsNullOrEmpty(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString())) Then
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("GiamDoc").ToString()
                                    Else
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("PhoGiamDoc").ToString()
                                    End If
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).RowHeight = 20
                                End If
                            End If
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
                    _ResultFile = file_name
                End Using
            Else
                MessageBox.Show("Không lấy được file excel mẫu. Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi xuất báo cáo ra excel: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi xuất báo cáo ra excel ChiLuongBLL\ExportExcel_TH_LuongThang_01TL: " + ex.Message)
            Return ""
        End Try
        Return _ResultFile
    End Function

    Public Function ExportExcel_CT_LuongThang(ByVal pDonVi_TK_Cd As String, ByVal pProvinceName As String, ByVal pBranchName As String, ByVal pNgayBC As DateTime, pLoaiCLCd As String, pKyBC As Byte, pLaoDong As String, db_source As DataTable) As String
        Dim _ResultFile As String = ""
        Try
            Dim _HeThongBLL As HeThongBLL = New HeThongBLL()
            Dim _NgayHeThong As DateTime = Globals.GetDateTime_ForServerDB
            Dim iRowStart As Integer = 0
            Dim _LapBieu_1 = "", _LapBieu_2 = "", _KiemSoat_1 = "", _KiemSoat_2 = "", _GiamDoc_1 = "", _GiamDoc_2 = ""
            Dim _ThoiDiem As DateTime = pNgayBC
            Dim sRootPath As String = System.Windows.Forms.Application.StartupPath
            Dim sPathExcelTemplate As String = sRootPath + Globals.GetAppSetting("File_Excel_Template")
            Dim sPath_Export As String = "C:\Temp\"
            Dim sFileExcelTemplate As String = "", sFileNameEx As String = "", sAutoNumber As String = "", sColNameEnd As String = ""
            sAutoNumber = SoftSqlHelper.GetString(String.Format("Select Cast(Abs(Checksum(NewID())) As Varchar(10))"), "")
            If pLoaiCLCd = "01" Then          'Chi lương tháng
                If pLaoDong = "1" Then               ' Lao động dài hạn
                    sFileExcelTemplate = sPathExcelTemplate + IIf(pKyBC = 1, "BC_LuongThang_DaiHan_Ky_01.xlsx", "BC_LuongThang_DaiHan_Ky_02.xlsx")
                    sFileNameEx = "BC_LuongThang_DaiHan_Ky_" + IIf(pKyBC = 1, "01", "02") + "_Thang_" + pNgayBC.Month.ToString("D2") + pNgayBC.Year.ToString("D4") + "_" + _NgayHeThong.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                    sColNameEnd = IIf(pKyBC = 1, "R", "U")
                    _LapBieu_1 = "B"
                    _LapBieu_2 = "C"
                    _KiemSoat_1 = IIf(pKyBC = 1, "J", "K")
                    _KiemSoat_2 = IIf(pKyBC = 1, "L", "M")
                    _GiamDoc_1 = IIf(pKyBC = 1, "P", "S")
                    _GiamDoc_2 = IIf(pKyBC = 1, "R", "U")
                    iRowStart = 12
                ElseIf pLaoDong = "2" Then           ' Lao động ngắn hạn
                    sFileExcelTemplate = sPathExcelTemplate + IIf(pKyBC = 1, "BC_LuongThang_NganHan.xlsx", "BC_LuongThang_NganHan.xlsx")
                    sFileNameEx = "BC_LuongThang_NganHan_" + "Thang_" + pNgayBC.Month.ToString("D2") + pNgayBC.Year.ToString("D4") + "_" + _NgayHeThong.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                    sColNameEnd = IIf(pKyBC = 1, "P", "P")
                    _LapBieu_1 = "B"
                    _LapBieu_2 = "C"
                    _KiemSoat_1 = IIf(pKyBC = 1, "H", "H")
                    _KiemSoat_2 = IIf(pKyBC = 1, "J", "J")
                    _GiamDoc_1 = IIf(pKyBC = 1, "N", "N")
                    _GiamDoc_2 = IIf(pKyBC = 1, "P", "P")
                    iRowStart = 12
                ElseIf pLaoDong = "3" Then           ' Lao động tập sự
                    sFileExcelTemplate = sPathExcelTemplate + IIf(pKyBC = 1, "BC_LuongThang_TapSu_Ky_01.xlsx", "BC_LuongThang_TapSu_Ky_02.xlsx")
                    sFileNameEx = "BC_LuongThang_TapSu_" + "Thang_" + pNgayBC.Month.ToString("D2") + pNgayBC.Year.ToString("D4") + "_" + _NgayHeThong.ToString("ddMMyyyy") + "_" + sAutoNumber + ".xlsx"
                    sColNameEnd = IIf(pKyBC = 1, "R", "U")
                    iRowStart = 12
                End If
            End If

            If (Globals.FileDao.IsFile(sFileExcelTemplate)) Then
                Using excelEngine As ExcelEngine = New ExcelEngine()
                    Dim streamRead As Stream = File.OpenRead(sFileExcelTemplate)
                    Dim workbook As IWorkbook = excelEngine.Excel.Workbooks.Open(streamRead, ExcelOpenType.Automatic)
                    Dim worksheet As IWorksheet = workbook.Worksheets(0)
                    worksheet.Range("A1").Text = pProvinceName
                    worksheet.Range("A2").Text = pBranchName
                    worksheet.Range("A5").Text = "Tháng " + _ThoiDiem.Month.ToString("D2") + " năm " + _ThoiDiem.Year.ToString("D4")

                    'Create Template Marker Processor
                    Dim marker As ITemplateMarkersProcessor = workbook.CreateTemplateMarkersProcessor()
                    Dim columnNames(db_source.Columns.Count) As String
                    Dim i As Integer = 0
                    For Each column As DataColumn In db_source.Columns
                        columnNames(i) = column.ColumnName
                        i += 1
                    Next

                    For Each c As DataColumn In db_source.Columns
                        Dim columnName As String = c.ColumnName
                        Dim columnData As EnumerableRowCollection(Of Object)
                        If Globals.IsNumeric(c) Then
                            columnData = db_source.AsEnumerable().[Select](Function(r) If(r.Field(Of Object)(columnName), 0))
                        ElseIf c.DataType Is GetType(Date) Then
                            columnData = db_source.AsEnumerable.[Select](Function(r) If(r.Field(Of Object)(columnName), ""))
                        Else
                            columnData = db_source.AsEnumerable.[Select](Function(r) If(("'" & r.Field(Of Object)(columnName)), ""))
                        End If

                        Dim paramDataArray As Object() = columnData.ToArray
                        marker.AddVariable(columnName, paramDataArray)
                    Next
                    marker.ApplyMarkers()
                    workbook.Version = ExcelVersion.Excel2013

                    'Căn chỉnh bôi đậm, nghiêng các dòng bản ghi và điền dữ liệu vào một số CELL
                    Dim iRowA As Integer = 0
                    If Not (db_source Is Nothing) Then
                        If (db_source.Rows.Count > 0) Then
                            For iTT As Integer = 0 To db_source.Rows.Count - 1
                                iRowA = (iRowStart + iTT)
                                Dim iKieuIn As Integer = CType(db_source.Rows(iTT)("KieuIn"), Integer)
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
                            'Xuất đoạn cuối: Lập biểu/Kiểm soát/Giám đốc
                            Dim db_sysvar As DataTable = New DataTable()
                            db_sysvar = _HeThongBLL.GetListSysVarSearch(0, 0, pDonVi_TK_Cd, "", "", "", 1)
                            If Not (db_sysvar Is Nothing) Then
                                If (db_sysvar.Rows.Count > 0) Then
                                    iRowA = iRowA + 1
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).Text = db_sysvar.Rows(0)("DiaBan").ToString().Trim() + ", ngày " + _NgayHeThong.Day.ToString("D2") + " thang " + _NgayHeThong.Month.ToString("D2") + " năm " + _NgayHeThong.Year.ToString("D4") + "        "
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight
                                    worksheet.Range(sColNameEnd + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).RowHeight = 23

                                    iRowA = iRowA + 1
                                    '_LapBieu_1 = "B", _LapBieu_2 = "C", _KiemSoat_1 = "J", _KiemSoat_2 = "L", _GiamDoc_1 = "P", _GiamDoc_2 = "R"
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Text = "NGƯỜI LẬP BIỂU"
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range(_KiemSoat_1 + iRowA.ToString() + ":" + _KiemSoat_2 + iRowA.ToString()).Text = "KIỂM SOÁT"
                                    worksheet.Range(_KiemSoat_1 + iRowA.ToString() + ":" + _KiemSoat_2 + iRowA.ToString()).Merge(True)
                                    If (IsNothing(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString()) Or String.IsNullOrEmpty(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString())) Then
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = IIf((DONVI = "000100" Or DONVI = "000199"), "TỔNG GIÁM ĐỐC", "GIÁM ĐỐC")
                                    Else
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = IIf((DONVI = "000100" Or DONVI = "000199"), "KT. TỔNG GIÁM ĐỐC" + vbNewLine + "PHÓ TỔNG GIÁM ĐỐC", "KT. GIÁM ĐỐC" + vbNewLine + "PHÓ GIÁM ĐỐC")
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).RowHeight = 27
                                    End If
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter

                                    iRowA = iRowA + 1
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Text = "(Ký ghi rõ họ tên)"
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range(_KiemSoat_1 + iRowA.ToString() + ":" + _KiemSoat_2 + iRowA.ToString()).Text = "(Ký ghi rõ họ tên)"
                                    worksheet.Range(_KiemSoat_1 + iRowA.ToString() + ":" + _KiemSoat_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = "(Ký tên, đóng dấu)"
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 10
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Italic = True
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter

                                    'Điền họ tên của: Lập biểu; Kiểm soát; Lãnh đạo
                                    iRowA = iRowA + 6
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("TruongHCTC").ToString()
                                    worksheet.Range(_LapBieu_1 + iRowA.ToString() + ":" + _LapBieu_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range(_KiemSoat_1 + iRowA.ToString() + ":" + _KiemSoat_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("KeToanTruong").ToString()
                                    worksheet.Range(_KiemSoat_1 + iRowA.ToString() + ":" + _KiemSoat_2 + iRowA.ToString()).Merge(True)
                                    If (IsNothing(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString()) Or String.IsNullOrEmpty(db_sysvar.Rows(0)("IdCanBo_PhoGD").ToString())) Then
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("GiamDoc").ToString()
                                    Else
                                        worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Text = db_sysvar.Rows(0)("PhoGiamDoc").ToString()
                                    End If
                                    worksheet.Range(_GiamDoc_1 + iRowA.ToString() + ":" + _GiamDoc_2 + iRowA.ToString()).Merge(True)
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.FontName = "Times New Roman"
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Size = 11
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.Font.Bold = True
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter
                                    worksheet.Range("A" + iRowA.ToString() + ":" + sColNameEnd + iRowA.ToString()).RowHeight = 20
                                End If
                            End If
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
                    _ResultFile = file_name
                End Using
            Else
                MessageBox.Show("Không lấy được file excel mẫu. Vui lòng kiểm tra lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi xuất báo cáo ra excel: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi xuất báo cáo ra excel ChiLuongBLL\ExportExcel_CT_LuongThang: " + ex.Message)
            Return ""
        End Try
        Return _ResultFile
    End Function



    'Xuất file Word Quyết định nâng bậc lương....

End Class
