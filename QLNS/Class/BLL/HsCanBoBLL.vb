Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class HsCanBoBLL
    Public Sub SetStyleRowGrid(ByVal iRow As Integer, ByVal dgv_Name As DataGridView, ByVal iKieuIn As Byte, ByVal iValueNote As Byte)
        dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.Navy
        If iKieuIn = 0 Or iKieuIn = 1 Or iKieuIn = Nothing Then
            dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Bold)
            dgv_Name.Rows(iRow).DefaultCellStyle.BackColor = System.Drawing.Color.Orange
            dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = System.Drawing.Color.GhostWhite
        Else
            If iKieuIn = 4 Then
                dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Italic)
                dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.OrangeRed
            Else
                If iKieuIn = 2 Then
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Bold)
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Italic)
                ElseIf iValueNote = 1 Then
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Regular)
                    dgv_Name.Rows(iRow).DefaultCellStyle.ForeColor = Color.Maroon
                Else
                    dgv_Name.Rows(iRow).DefaultCellStyle.Font = New System.Drawing.Font(dgv_Name.Font, FontStyle.Regular)
                End If
            End If
        End If
    End Sub
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal state As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()

        dgv_name.EnableHeadersVisualStyles = False
        dgv_name.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray 'Color.FromArgb(178, 219, 255)
        dgv_name.ColumnHeadersDefaultCellStyle.ForeColor = Color.Blue 'Color.FromArgb(16, 37, 127)
        Dim sColumnVisible As String = ""

        Select Case state
            Case 1 ' BÁO CÁO THỐNG KÊ SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ
                dgv_name.Columns.Add("cln_DonVi_Cd", "Mã số") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                          '1
                dgv_name.Columns.Add("cln_STT", "STT")                                          '2
                dgv_name.Columns.Add("cln_DonVi_HT", "Đơn vị")                                    '3
                dgv_name.Columns.Add("cln_TongSo_LD", "Tổng số lao động có mặt")                '4
                dgv_name.Columns.Add("cln_TongSo_LD_Nu", "Nữ")                                  '5
                dgv_name.Columns.Add("cln_TongSo_DangVien", "Đảng viên")                        '6
                dgv_name.Columns.Add("cln_DanToc_Tso", "Dân tộc thiểu số")                      '7
                dgv_name.Columns.Add("cln_ChinhTri_CC", "Cao cấp")                              '8      Trình độ chính trị
                dgv_name.Columns.Add("cln_ChinhTri_CN", "Cử nhân")                              '9
                dgv_name.Columns.Add("cln_ChinhTri_TC", "Trung cấp")                            '10
                dgv_name.Columns.Add("cln_ChinhTri_SC", "Sơ cấp")                               '11
                
                dgv_name.Columns.Add("cln_TrDoCM_TienSy", "Tiến sỹ")                            '12      Trình độ chuyên môn
                dgv_name.Columns.Add("cln_TrDoCM_ThacSy", "Thạc sỹ")                            '13
                dgv_name.Columns.Add("cln_TrDoCM_DaiHoc", "Đại học")                            '14
                dgv_name.Columns.Add("cln_TrDoCM_CaoDang", "Cao đẳng")                          '15
                dgv_name.Columns.Add("cln_TrDoCM_CaoCapNH", "Cao cấp NH")                       '16
                dgv_name.Columns.Add("cln_TrDoCM_TrungCap", "Trung cấp")                        '17
                dgv_name.Columns.Add("cln_TrDoCM_SCKhac", "Sơ cấp và khác")                     '18

                dgv_name.Columns.Add("cln_NgoaiNgu_DH", "Đại học")                              '19     Trình độ ngoại ngữ 
                dgv_name.Columns.Add("cln_NgoaiNgu_C", "Bằng C")                                '20
                dgv_name.Columns.Add("cln_NgoaiNgu_B", "Bằng B")                                '21
                dgv_name.Columns.Add("cln_NgoaiNgu_A", "Bằng A")                                '22

                dgv_name.Columns.Add("cln_TinHoc_DH", "Đại học")                                '23     Trình độ tin học
                dgv_name.Columns.Add("cln_TinHoc_C", "Bằng C")                                  '24
                dgv_name.Columns.Add("cln_TinHoc_B", "Bằng B")                                  '25
                dgv_name.Columns.Add("cln_TinHoc_A", "Bằng A")                                  '26

                dgv_name.Columns.Add("cln_TuoiDoi_30", "Dưới 30")                               '27         Tuổi đời
                dgv_name.Columns.Add("cln_TuoiDoi_31_35", "Từ 31 đến 35")                       '28
                dgv_name.Columns.Add("cln_TuoiDoi_36_40", "Từ 36 đến 40")                       '29
                dgv_name.Columns.Add("cln_TuoiDoi_41_45", "Từ 41 đến 45")                       '30
                dgv_name.Columns.Add("cln_TuoiDoi_46_50", "Từ 46 đến 50")                       '31
                dgv_name.Columns.Add("cln_TuoiDoi_51_55", "Từ 51 đến 55")                       '32
                dgv_name.Columns.Add("cln_TuoiDoi_56_62", "Từ 55 đến 62")                       '33

                dgv_name.Columns.Add("cln_LDao_BanGDNHCS", "Ban lãnh đạo NHCSXH")                     '34     'Chức danh Lãnh đạo
                dgv_name.Columns.Add("cln_LDao_GDBanHSC", "Giám đốc các Ban tại HSC và tương đương")   '35
                dgv_name.Columns.Add("cln_LDao_BanGDCN", "Ban Giám đốc CN cấp tỉnh")                  '36
                dgv_name.Columns.Add("cln_LDao_TP_PP", "Trưởng, Phó phòng CMNV hoặc tương đương")   '37
                dgv_name.Columns.Add("cln_LDao_BanGDPGD", "Ban Giám đốc cấp huyện")                 '38

                dgv_name.Columns.Add("cln_ChMon_ToTruong_HSC", "Tổ trưởng Ban CMNV tại HSC")             '39     Cán bộ CMNV
                dgv_name.Columns.Add("cln_ChMon_ToTruong_PGD", "Tổ trưởng nghiệp vụ PGD cấp huyện")      '40
                dgv_name.Columns.Add("cln_ChMon_TinDung", "Tín dụng")                               '41 
                dgv_name.Columns.Add("cln_ChMon_KeToan", "Kế toán")                                 '42
                dgv_name.Columns.Add("cln_ChMon_KTKTNB", "KT KTNB")                                 '43
                dgv_name.Columns.Add("cln_ChMon_TinHoc", "Tin học")                                 '44
                dgv_name.Columns.Add("cln_ChMon_HcNhanSu", "HC-TC")                                 '45
                dgv_name.Columns.Add("cln_ChMon_ThuQuy", "Thủ quỹ")                                 '46
                dgv_name.Columns.Add("cln_ChMon_LaiXe", "Lái xe")                                   '47
                dgv_name.Columns.Add("cln_ChMon_NvKhac", "Nhân viên khác")                          '48
                dgv_name.Columns.Add("cln_BaoVe_TapVu", "Bảo vệ, tạp vụ")                           '49     'Bảo vệ, tạp vụ

                sColumnVisible = "KieuIn;KhuVuc_HT;DonVi_Id;Id"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_DonVi_Cd").Width = 50
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_HT").Width = 130                           '3
                dgv_name.Columns("cln_DonVi_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TongSo_LD").Width = 60                           '11
                dgv_name.Columns("cln_TongSo_LD").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TongSo_LD_Nu").Width = 55                           '11
                dgv_name.Columns("cln_TongSo_LD_Nu").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TongSo_DangVien").Width = 60                           '11
                dgv_name.Columns("cln_TongSo_DangVien").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DanToc_Tso").Width = 55                           '11
                dgv_name.Columns("cln_DanToc_Tso").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                Dim sColumnWidth As String = "" 'ChinhTri_CC;ChinhTri_TC;ChinhTri_SC;TrDoCM_TienSy;TrDoCM_ThacSy;TrDoCM_DaiHoc;TrDoCM_CaoDang;TrDoCM_CaoCapNH;TrDoCM_TrungCap;TrDoCM_SCKhac;NgoaiNgu_DH;NgoaiNgu_A;NgoaiNgu_B;NgoaiNgu_C;TinHoc_DH;TinHoc_A;TinHoc_B;TinHoc_C;TuoiDoi_30;TuoiDoi_31_35;TuoiDoi_36_40;TuoiDoi_41_45;TuoiDoi_46_50;TuoiDoi_51_55;TuoiDoi_55_60;LDao_BanGDCN;LDao_TP_PP;LDao_BanGDPGD;LDao_ToTruong;ChMon_TinDung;ChMon_KeToan;ChMon_TinHoc;ChMon_ThuQuy;ChMon_KTKTNB;ChMon_HcNhanSu;ChMon_LaiXe;ChMon_NvKhac;BaoVe_TapVu
                sColumnWidth = "ChinhTri_CC;ChinhTri_CN;ChinhTri_TC;ChinhTri_SC;TrDoCM_TienSy;TrDoCM_ThacSy;TrDoCM_DaiHoc;TrDoCM_CaoDang;TrDoCM_CaoCapNH;TrDoCM_TrungCap;TrDoCM_SCKhac;NgoaiNgu_DH;NgoaiNgu_C;NgoaiNgu_B;NgoaiNgu_A;TinHoc_DH;TinHoc_C;TinHoc_B;TinHoc_A;TuoiDoi_30;TuoiDoi_31_35;TuoiDoi_36_40;TuoiDoi_41_45;TuoiDoi_46_50;TuoiDoi_51_55;TuoiDoi_56_62;LDao_BanGDNHCS;LDao_GDBanHSC;LDao_BanGDCN;LDao_TP_PP;LDao_BanGDPGD;ChMon_ToTruong_HSC;ChMon_ToTruong_PGD;ChMon_TinDung;ChMon_KeToan;ChMon_KTKTNB;ChMon_TinHoc;ChMon_HcNhanSu;ChMon_ThuQuy;ChMon_LaiXe;ChMon_NvKhac;BaoVe_TapVu"
                Dim ARL_ColWidth() As String = Globals.Splip_Strings(sColumnWidth, ";")
                For Each _Value As String In ARL_ColWidth
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).Width = 50
                        dgv_name.Columns("cln_" + _Value).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                    End If
                Next
                dgv_name.Columns("cln_LDao_BanGDCN").Width = 60
                dgv_name.Columns("cln_LDao_TP_PP").Width = 75
                dgv_name.Columns("cln_ChMon_HcNhanSu").Width = 50

                dgv_name.Columns("cln_TongSo_LD").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_TongSo_LD_Nu").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_TongSo_DangVien").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_DanToc_Tso").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_Cd").Visible = False
                dgv_name.Columns("cln_Choice").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_DonVi_HT").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_DonVi_HT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                dgv_name.ColumnHeadersHeight = 30

            Case 2 ' TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MÀNG LƯỚI HOẠT 
                dgv_name.Columns.Add("cln_Id", "Id") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                          '1
                dgv_name.Columns.Add("cln_STT", "STT")                                          '2
                dgv_name.Columns.Add("cln_DonVi_HT", "Đơn vị")                                    '3
                dgv_name.Columns.Add("cln_Thang3_ThongBao_TC_CN", "Tổng số chỉ tiêu TW thông báo")      '4
                dgv_name.Columns.Add("cln_Thang1_DaiHan_TBCN", "Số TB")                               '5          Tháng 1 - Dài Hạn
                dgv_name.Columns.Add("cln_Thang1_DaiHan_TH", "Số TH")                               '6
                dgv_name.Columns.Add("cln_Thang1_NganHan_TBCN", "Số TB")                               '7          Tháng 1 - Ngắn hạn
                dgv_name.Columns.Add("cln_Thang1_NganHan_TH", "Số TH")                               '8

                dgv_name.Columns.Add("cln_Thang2_DaiHan_TBCN", "Số TB")                               '9          Tháng 2 - Dài Hạn
                dgv_name.Columns.Add("cln_Thang2_DaiHan_TH", "Số TH")                               '10
                dgv_name.Columns.Add("cln_Thang2_NganHan_TBCN", "Số TB")                               '11         Tháng 2 - Ngắn hạn
                dgv_name.Columns.Add("cln_Thang2_NganHan_TH", "Số TH")                               '12

                dgv_name.Columns.Add("cln_Thang3_DaiHan_TBCN", "Số TB")                               '13         Tháng 3 - Dài Hạn
                dgv_name.Columns.Add("cln_Thang3_DaiHan_TH", "Số TH")                               '14
                dgv_name.Columns.Add("cln_Thang3_NganHan_TBCN", "Số TB")                               '15         Tháng 3 - Ngắn hạn
                dgv_name.Columns.Add("cln_Thang3_NganHan_TH", "Số TH")                               '16

                dgv_name.Columns.Add("cln_SoKH_DN", "Số khách hàng")                 '17         Màng lưới hoạt động 
                dgv_name.Columns.Add("cln_So_ToTKVV", "Số Tổ TK-VV")                 '18         Màng lưới hoạt động 
                dgv_name.Columns.Add("cln_So_XaPhuong", "Số xã, phường")                      '19
                dgv_name.Columns.Add("cln_So_DiemGD", "Số điểm giao dịch")             '20
                dgv_name.Columns.Add("cln_TongDN", "Dư nợ (tỷ đồng)")                         '21

                sColumnVisible = "GhiChu;TT_Dong;STT_KV;STT_CN;STT_POS;STT_LUONG;ThoiDiem;ThoiDiem_HT;KhuVuc_Cd;KhuVuc_HT;DonVi_Id;DonVi_Cd;ChiNhanh_Id;ChiNhanh_Cd;ChiNhanh_HT;TenChiNhanh_HT;PhongBan_Cd;PhongBan_HT;Luong_TTV_Ma;Luong_TTV_HT;KieuIn;Thang1_ThucHien_TC;Thang2_ThucHien_TC;Thang3_ThucHien_TC;Thang1_ThongBao_TC_CN;Thang2_ThongBao_TC_CN;Thang1_DaiHan_TB;Thang1_NganHan_TB;Thang2_DaiHan_TB;Thang2_NganHan_TB;Thang3_DaiHan_TB;Thang3_NganHan_TB;Thang1_ThongBao_TC;Thang2_ThongBao_TC;Thang3_ThongBao_TC;DuNo_TH;DuNo_QH;DuNo_KH"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_HT").Width = 140
                dgv_name.Columns("cln_DonVi_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Thang3_ThongBao_TC_CN").Width = 100
                dgv_name.Columns("cln_Thang3_ThongBao_TC_CN").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TongDN").Width = 100
                dgv_name.Columns("cln_TongDN").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                
                Dim sColumnWidth As String = "Thang1_DaiHan_TBCN;Thang1_DaiHan_TH;Thang1_NganHan_TBCN;Thang1_NganHan_TH;Thang2_DaiHan_TBCN;Thang2_DaiHan_TH;Thang2_NganHan_TBCN;Thang2_NganHan_TH;Thang3_DaiHan_TBCN;Thang3_DaiHan_TH;Thang3_NganHan_TBCN;Thang3_NganHan_TH"
                Dim ARL_ColWidth() As String = Globals.Splip_Strings(sColumnWidth, ";")
                For Each _Value As String In ARL_ColWidth
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).Width = 60
                        dgv_name.Columns("cln_" + _Value).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                    End If
                Next

                dgv_name.Columns("cln_SoKH_DN").Width = 90
                dgv_name.Columns("cln_So_ToTKVV").Width = 80
                dgv_name.Columns("cln_So_XaPhuong").Width = 70
                dgv_name.Columns("cln_So_DiemGD").Width = 85
                dgv_name.Columns("cln_TongDN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_TongDN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_SoKH_DN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_So_XaPhuong").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_So_DiemGD").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_So_ToTKVV").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Thang3_ThongBao_TC_CN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Choice").Visible = False
                dgv_name.Columns("cln_Id").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_DonVi_HT").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_DonVi_HT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                dgv_name.ColumnHeadersHeight = 60

            Case 3          ' BÁO CÁO THỐNG KÊ SỐ LƯỢNG VÀ CHẤT LƯỢNG CÁN BỘ - CHI TIẾT
                'dgv_name.Columns.Add("cln_IdCanBo", "Mã số") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                          '1
                dgv_name.Columns.Add("cln_STT", "STT")                                          '2
                dgv_name.Columns.Add("cln_ChiNhanh_HT", "Chi nhánh")                            '2
                dgv_name.Columns.Add("cln_DonVi_HT", "Đơn vị công tác")                         '3
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên")
                dgv_name.Columns.Add("cln_MaCB", "Mã CB")
                dgv_name.Columns.Add("cln_NgaySinh_HT", "Ngày sinh")
                dgv_name.Columns.Add("cln_GioiTinh_HT", "Giới tính")
                dgv_name.Columns.Add("cln_DanToc_HT", "Dân tộc")
                dgv_name.Columns.Add("cln_TrinhDoCT_HT", "Trình độ chính trị")
                dgv_name.Columns.Add("cln_PhongBan_HT", "Phòng ban")
                dgv_name.Columns.Add("cln_ChucVu_HT", "Chức vụ")
                dgv_name.Columns.Add("cln_ChuyenMon_HT", "Chuyên môn công tác theo QĐNS")
                dgv_name.Columns.Add("cln_DangVien", "Đảng viên")
                dgv_name.Columns.Add("cln_DT_TrinhDo_HT", "Trình độ đào tạo")
                dgv_name.Columns.Add("cln_DT_NgoaiNgu_HT", "Trình độ ngoại ngữ")
                dgv_name.Columns.Add("cln_DT_TinHoc_HT", "Trình độ tin học")
                dgv_name.Columns.Add("cln_Loai_CB", "Loại cán bộ")
                sColumnVisible = "KieuIn;STT_HT;ThoiDiemDL;IdCanBo;DonVi_Id;DonVi_Cd;ChiNhanh_Id;ChiNhanh_Cd;TenChiNhanh_HT;PhongBan_Id;PhongBan_Cd;ChucVu_Id;ChucVu_Cd;ChuyenMon_Id;ChuyenMon_Cd;IdLoaiQD;LoaiQd_Cd;LoaiQd_HT;NgayHL;DanToc_Id;DanToc_Cd;NgaySinh;TuoiDoi;GioiTinh;TrinhDoCT_Id;TrinhDoCT_Cd;DangVien_Id;DangVien_SoThe;DangVien_HT;DT_TrinhDo_Id;DT_TrinhDo_Cd;DT_NgoaiNgu_Cd;DT_TinHoc_Cd;DonViALL"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_STT").Width = 40
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ChiNhanh_HT").Width = 80
                dgv_name.Columns("cln_ChiNhanh_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_HT").Width = 180
                dgv_name.Columns("cln_DonVi_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_HoTen").Width = 140
                dgv_name.Columns("cln_HoTen").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_MaCB").Width = 55
                dgv_name.Columns("cln_MaCB").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NgaySinh_HT").Width = 73
                dgv_name.Columns("cln_NgaySinh_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GioiTinh_HT").Width = 45
                dgv_name.Columns("cln_GioiTinh_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DanToc_HT").Width = 55
                dgv_name.Columns("cln_DanToc_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TrinhDoCT_HT").Width = 60
                dgv_name.Columns("cln_TrinhDoCT_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_PhongBan_HT").Width = 140
                dgv_name.Columns("cln_PhongBan_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ChucVu_HT").Width = 140
                dgv_name.Columns("cln_ChucVu_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_ChuyenMon_HT").Width = 140
                dgv_name.Columns("cln_ChuyenMon_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DangVien").Width = 40
                dgv_name.Columns("cln_DangVien").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DT_TrinhDo_HT").Width = 90
                dgv_name.Columns("cln_DT_TrinhDo_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DT_NgoaiNgu_HT").Width = 65
                dgv_name.Columns("cln_DT_NgoaiNgu_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DT_TinHoc_HT").Width = 65
                dgv_name.Columns("cln_DT_TinHoc_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Loai_CB").Width = 125
                dgv_name.Columns("cln_Loai_CB").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_MaCB").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NgaySinh_HT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GioiTinh_HT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DanToc_HT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DangVien").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DangVien").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

                dgv_name.Columns("cln_IdCanBo").Visible = False
                dgv_name.Columns("cln_Choice").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_ChiNhanh_HT").Frozen = True
                dgv_name.Columns("cln_DonVi_HT").Frozen = True
                dgv_name.Columns("cln_HoTen").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_ChiNhanh_HT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_DonVi_HT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_HoTen").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                dgv_name.ColumnHeadersHeight = 30

            Case 4 ' TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MÀNG LƯỚI HOẠT => Như với Cờ 2 nhưng áp dụng cho Chi nhánh sao kê các đơn vị PGD
                dgv_name.Columns.Add("cln_Id", "Id") '0
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)                          '1
                dgv_name.Columns.Add("cln_STT", "STT")                                          '2
                dgv_name.Columns.Add("cln_DonVi_HT", "Đơn vị")                                    '3
                dgv_name.Columns.Add("cln_Thang3_ThongBao_TC", "Tổng số chỉ tiêu TW thông báo")      '4
                dgv_name.Columns.Add("cln_Thang1_DaiHan_TB", "Số TB")                               '5          Tháng 1 - Dài Hạn
                dgv_name.Columns.Add("cln_Thang1_DaiHan_TH", "Số TH")                               '6
                dgv_name.Columns.Add("cln_Thang1_NganHan_TB", "Số TB")                               '7          Tháng 1 - Ngắn hạn
                dgv_name.Columns.Add("cln_Thang1_NganHan_TH", "Số TH")                               '8

                dgv_name.Columns.Add("cln_Thang2_DaiHan_TB", "Số TB")                               '9          Tháng 2 - Dài Hạn
                dgv_name.Columns.Add("cln_Thang2_DaiHan_TH", "Số TH")                               '10
                dgv_name.Columns.Add("cln_Thang2_NganHan_TB", "Số TB")                               '11         Tháng 2 - Ngắn hạn
                dgv_name.Columns.Add("cln_Thang2_NganHan_TH", "Số TH")                               '12

                dgv_name.Columns.Add("cln_Thang3_DaiHan_TB", "Số TB")                               '13         Tháng 3 - Dài Hạn
                dgv_name.Columns.Add("cln_Thang3_DaiHan_TH", "Số TH")                               '14
                dgv_name.Columns.Add("cln_Thang3_NganHan_TB", "Số TB")                               '15         Tháng 3 - Ngắn hạn
                dgv_name.Columns.Add("cln_Thang3_NganHan_TH", "Số TH")                               '16

                dgv_name.Columns.Add("cln_SoKH_DN", "Số khách hàng")                 '17         Màng lưới hoạt động 
                dgv_name.Columns.Add("cln_So_ToTKVV", "Số Tổ TK-VV")                 '18         Màng lưới hoạt động 
                dgv_name.Columns.Add("cln_So_XaPhuong", "Số xã, phường")                      '19
                dgv_name.Columns.Add("cln_So_DiemGD", "Số điểm giao dịch")             '20
                dgv_name.Columns.Add("cln_TongDN", "Dư nợ (tỷ đồng)")                         '21

                sColumnVisible = "GhiChu;TT_Dong;STT_KV;STT_CN;STT_POS;STT_LUONG;ThoiDiem;ThoiDiem_HT;KhuVuc_Cd;KhuVuc_HT;DonVi_Id;DonVi_Cd;ChiNhanh_Id;ChiNhanh_Cd;ChiNhanh_HT;TenChiNhanh_HT;PhongBan_Cd;PhongBan_HT;Luong_TTV_Ma;Luong_TTV_HT;KieuIn;Thang1_DaiHan_TBCN;Thang1_NganHan_TBCN;Thang2_DaiHan_TBCN;Thang2_NganHan_TBCN;Thang3_DaiHan_TBCN;Thang3_NganHan_TBCN;Thang1_ThucHien_TC;Thang2_ThucHien_TC;Thang3_ThucHien_TC;Thang1_ThongBao_TC_CN;Thang2_ThongBao_TC_CN;Thang3_ThongBao_TC_CN;Thang1_ThongBao_TC;Thang2_ThongBao_TC;DuNo_TH;DuNo_QH;DuNo_KH"
                Dim ARL_Cols() As String = Globals.Splip_Strings(sColumnVisible, ";")
                For Each _Value As String In ARL_Cols
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns.Add("cln_" + _Value, _Value)
                        dgv_name.Columns("cln_" + _Value).Visible = False
                    End If
                Next

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_STT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_HT").Width = 140
                dgv_name.Columns("cln_DonVi_HT").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Thang3_ThongBao_TC").Width = 100
                dgv_name.Columns("cln_Thang3_ThongBao_TC").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TongDN").Width = 100
                dgv_name.Columns("cln_TongDN").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                Dim sColumnWidth As String = "Thang1_DaiHan_TB;Thang1_DaiHan_TH;Thang1_NganHan_TB;Thang1_NganHan_TH;Thang2_DaiHan_TB;Thang2_DaiHan_TH;Thang2_NganHan_TB;Thang2_NganHan_TH;Thang3_DaiHan_TB;Thang3_DaiHan_TH;Thang3_NganHan_TB;Thang3_NganHan_TH"
                Dim ARL_ColWidth() As String = Globals.Splip_Strings(sColumnWidth, ";")
                For Each _Value As String In ARL_ColWidth
                    If Not String.IsNullOrEmpty(_Value) Then
                        dgv_name.Columns("cln_" + _Value).Width = 60
                        dgv_name.Columns("cln_" + _Value).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                    End If
                Next

                dgv_name.Columns("cln_SoKH_DN").Width = 90
                dgv_name.Columns("cln_So_ToTKVV").Width = 80
                dgv_name.Columns("cln_So_XaPhuong").Width = 70
                dgv_name.Columns("cln_So_DiemGD").Width = 85
                dgv_name.Columns("cln_TongDN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_TongDN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_SoKH_DN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_So_XaPhuong").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_So_DiemGD").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_So_ToTKVV").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Thang3_ThongBao_TC").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Choice").Visible = False
                dgv_name.Columns("cln_Id").Visible = False
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_STT").Frozen = True
                dgv_name.Columns("cln_DonVi_HT").Frozen = True
                dgv_name.Columns("cln_Choice").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_STT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.Columns("cln_DonVi_HT").DefaultCellStyle.BackColor = Color.LightGray
                dgv_name.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
                dgv_name.ColumnHeadersHeight = 60
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Lấy dữ liệu Danh sách cán bộ truy vấn
    ''' </summary>
    ''' <param name="_CanBoTWQL">Chỉ số xác định lấy toàn bộ cán bộ. Giá trị: 0 - Tất cả; 1 - Cán bộ do TW Quản lý; 2 - Cán bộ Bảo vệ, Tạp vụ</param>
    ''' <param name="_IdDonVi">Chỉ số xác định đơn vị cần lấy. Nếu </param>
    ''' <param name="_ThoiDiem">Thời điểm cần lấy dữ liệu</param>
    ''' <returns>Danh sách cán bộ trả về</returns>
    ''' <remarks></remarks>
    Public Function GetDsCanBos(ByVal _LoaiCB As Byte, ByVal _IdDonVi As Integer, ByVal _ThoiDiem As DateTime) As DataTable
        Try
            Dim ds_ret As DataSet = New DataSet()
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                If (_LoaiCB = 0) Then
                    Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("GetListCanBo_NEW", conn_obj)
                        _command.CommandType = CommandType.StoredProcedure
                        _command.Parameters.Add(New SqlParameter("@_IdDonVi", SqlDbType.Int))
                        _command.Parameters("@_IdDonVi").Value = _IdDonVi

                        _command.Parameters.Add(New SqlParameter("@_ThoiDiem", SqlDbType.DateTime))
                        _command.Parameters("@_ThoiDiem").Value = _ThoiDiem

                        Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                            _sqldap.Fill(ds_ret, "Tbl_DsCanBo")
                            If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                Return Nothing
                            End If
                        End Using
                    End Using
                Else
                    If (_LoaiCB = 1) Then
                        Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("GetListCanBo_TWQL", conn_obj)
                            _command.CommandType = CommandType.StoredProcedure

                            _command.Parameters.Add(New SqlParameter("@_Id_Hsc_DonVi", SqlDbType.Int))
                            _command.Parameters("@_Id_Hsc_DonVi").Value = _IdDonVi

                            _command.Parameters.Add(New SqlParameter("@_ThoiDiem", SqlDbType.DateTime))
                            _command.Parameters("@_ThoiDiem").Value = _ThoiDiem

                            Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                                _sqldap.Fill(ds_ret, "Tbl_DsCanBo")
                                If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                    Return Nothing
                                End If
                            End Using
                        End Using
                    Else
                        Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("GetListCanBo_BVTS", conn_obj)
                            _command.CommandType = CommandType.StoredProcedure
                            _command.Parameters.Add(New SqlParameter("@_IdDonVi", SqlDbType.Int))
                            _command.Parameters("@_IdDonVi").Value = _IdDonVi

                            _command.Parameters.Add(New SqlParameter("@_ThoiDiem", SqlDbType.DateTime))
                            _command.Parameters("@_ThoiDiem").Value = _ThoiDiem

                            Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                                _sqldap.Fill(ds_ret, "Tbl_DsCanBo")
                                If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                    Return Nothing
                                End If
                            End Using
                        End Using
                    End If
                End If
            End Using
            Return ds_ret.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Lỗi truy vấn Danh sách cán bộ: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function
    Public Function GetDsCanBos_OLD(ByVal _CanBoTWQL As Byte, ByVal _IdDonVi As Integer, ByVal _ThoiDiem As DateTime) As DataTable
        Try
            Dim db_reult As DataTable = New DataTable()
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()

                If (_CanBoTWQL = 0) Then
                    Dim sqlParams(2) As SqlClient.SqlParameter

                    sqlParams(0) = New SqlClient.SqlParameter("@_IdDonVi", SqlDbType.Int)
                    sqlParams(0).Value = CType(_IdDonVi, Int32)

                    sqlParams(1) = New SqlClient.SqlParameter("@_ThoiDiem", SqlDbType.DateTime)
                    sqlParams(1).Value = CType(_ThoiDiem, DateTime)
                    db_reult = SoftSqlHelper.ExecuteForSQL(conn_obj, CommandType.StoredProcedure, "GetListCanBo_NEW", sqlParams)
                Else
                    db_reult = SoftSqlHelper.ExecuteForSQL(conn_obj, CommandType.StoredProcedure, "GetListCanBo_TWQL", SoftSqlHelper.CreateParameter("@_ThoiDiem", _ThoiDiem, ParameterDirection.Input))
                End If
            End Using
            Return db_reult
        Catch ex As Exception
            MessageBox.Show("Lỗi truy vấn Danh sách cán bộ: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện lấy sao kê dữ liệu theo yêu cầu
    ''' </summary>
    ''' <param name="_MaSK">Mã sao kê cần lấy. Giá trị quy ước:
    '''                          1 - Sao kê Số lượng, chất lượng cán bộ
    '''                          2 - Sao kê Tình hình thực hiện lao động - Màng lưới lao động
    '''                          5 - Sao kê Số lượng, chất lượng cán bộ 08-BC-TCCB. Áp dụng từ 11/2022
    ''' </param>
    ''' <param name="_IdDonVi">Chỉ số xác định đơn vị cần lấy sao kê</param>
    ''' <param name="_ThoiDiem">Thời điểm lấy sao kê</param>
    ''' <param name="_LoaiSK">Loại sao kê. Giá trị 1 - Tổng hợp; 2 - Chi tiết</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetData_SaoKe(ByVal _MaSK As Integer, ByVal _IdDonVi As Integer, ByVal _ThoiDiem As DateTime, ByVal _LoaiSK As Integer) As DataTable
        Try
            Dim ds_ret As DataSet = New DataSet()
            Select Case _MaSK
                Case 1      'Sao kê Số lượng, chất lượng cán bộ
                    Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                        Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("GetTongHopSL_01", conn_obj)
                            _command.CommandType = CommandType.StoredProcedure
                            _command.Parameters.Add(New SqlParameter("@_UnitId", SqlDbType.Int))
                            _command.Parameters("@_UnitId").Value = _IdDonVi

                            _command.Parameters.Add(New SqlParameter("@_ThoiDiem", SqlDbType.DateTime))
                            _command.Parameters("@_ThoiDiem").Value = _ThoiDiem

                            _command.Parameters.Add(New SqlParameter("@_LoaiSK", SqlDbType.Int))
                            _command.Parameters("@_LoaiSK").Value = _LoaiSK

                            Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                                _sqldap.Fill(ds_ret, "Tbl_DsCanBo")
                                If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                    Return Nothing
                                End If
                            End Using
                        End Using
                    End Using
                Case 5      'Sao kê Số lượng, chất lượng cán bộ (Mới áp dụng từ 11/2022)
                    Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                        Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("BaoCao_SoLuongChatLuongCB_08", conn_obj)
                            _command.CommandType = CommandType.StoredProcedure
                            _command.Parameters.Add(New SqlParameter("@pDonViId", SqlDbType.Int))
                            _command.Parameters("@pDonViId").Value = _IdDonVi

                            _command.Parameters.Add(New SqlParameter("@pThoiDiem", SqlDbType.VarChar, 10))
                            _command.Parameters("@pThoiDiem").Value = _ThoiDiem.ToString("yyyy-MM-dd")

                            _command.Parameters.Add(New SqlParameter("@pFlagCall", SqlDbType.Int))
                            _command.Parameters("@pFlagCall").Value = _LoaiSK

                            Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                                _sqldap.Fill(ds_ret, "Tbl_DsCanBo")
                                If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                    Return Nothing
                                End If
                            End Using
                        End Using
                    End Using
                Case 2      'Sao kê Tình hình thực hiện lao động - Màng lưới lao động
                    If _LoaiSK = 1 Then     'Trường hợp tổng hợp
                        Dim _ProceName As String = ""
                        If _IdDonVi = 1 Then
                            _ProceName = "GetTongHopSL_02_TQ"
                        Else
                            _ProceName = "GetTongHopSL_02_CN"
                        End If
                        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand(_ProceName, conn_obj)
                                _command.CommandType = CommandType.StoredProcedure
                                _command.Parameters.Add(New SqlParameter("@_UnitId", SqlDbType.Int))
                                _command.Parameters("@_UnitId").Value = _IdDonVi

                                _command.Parameters.Add(New SqlParameter("@_ThoiDiem", SqlDbType.DateTime))
                                _command.Parameters("@_ThoiDiem").Value = _ThoiDiem

                                _command.Parameters.Add(New SqlParameter("@_LoaiSK", SqlDbType.Int))
                                _command.Parameters("@_LoaiSK").Value = _LoaiSK

                                Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                                    _sqldap.Fill(ds_ret, "Tbl_DsCanBo")
                                    If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                        Return Nothing
                                    End If
                                End Using
                            End Using
                        End Using
                    Else                    'Trường hợp chi tiết

                    End If
                Case Else

            End Select

            Return ds_ret.Tables(0)

        Catch ex As Exception
            MessageBox.Show("Lỗi Sao kê dữ liệu: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

End Class
