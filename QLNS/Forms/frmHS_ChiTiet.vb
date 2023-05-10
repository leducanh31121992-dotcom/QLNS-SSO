Public Class frmHS_ChiTiet

#Region "---> Defined parametter and properties <---"
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()
    'Khai báo và khởi tạo thuộc tính của các đối tượng
    Private _IdCanBo As String
    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal value As String)
            _IdCanBo = value
        End Set
    End Property

    'Khai báo các biến để di chuyển bản ghi
    Private _Records As Int16    'Tổng số bản ghi tìm thấy
    Public Property Records() As Int16
        Get
            Return _Records
        End Get
        Set(ByVal value As Int16)
            _Records = value
        End Set
    End Property

    Private _RecordCurrent As Byte    'Số thứ tự bản ghi select
    Public Property RecordCurrent() As Byte
        Get
            Return _RecordCurrent
        End Get
        Set(ByVal value As Byte)
            _RecordCurrent = value
        End Set
    End Property

    ''' <summary>
    ''' Mảng lưu danh sách các id của hồ sơ tìm được
    ''' </summary>
    ''' <remarks></remarks>
    Public arr_RecordId As ArrayList = New ArrayList()
    'Biến lưu đơn vị Hiện tại của Cán bộ (Load thông tin quyết định nhân sự mới đây nhất của cán bộ)
    Public _IdDonviHT As Integer = 0
    Private arr_Ut_Banthan As ArrayList = New ArrayList
#End Region

#Region "---> Functions main <---"
    ''' <summary>
    ''' Hàm load thông tin Chi tiết về Hồ sơ cán bộ
    ''' </summary>
    ''' <param name="_Id">Id cán bộ</param>
    ''' <remarks></remarks>
    Private Sub Load_Infor_PersonelFile(ByVal _Id As String)
        Try
            ResetControls()
            Dim dr As DataRow
            dr = _HS_CanBo.GetHuman_ForCode(_Id)
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    'Chi tiết: Thông tin chung về cán bộ
                    lbl_macb.Text = dr("MaCB").ToString().Trim()
                    lbl_hoten.Text = dr("HoTen").ToString().Trim()
                    lbl_ten_tg.Text = dr("TenThuongGoi").ToString().Trim()
                    lbl_bidanh.Text = dr("BiDanh").ToString().Trim()
                    lbl_ten_tg.Text = dr("TenThuongGoi").ToString().Trim()
                    If (dr("GioiTinh").ToString() <> "") Then lbl_gioitinh.Text = IIf(dr("GioiTinh").ToString() = False, " Nam", " Nữ")
                    If dr("NgaySinh").ToString() <> "" Then lbl_ngaysinh.Text = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_donvi.Text = IIf(dr("IdDonVi").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh Where id = {0} and Status = 1", CType(dr("IdDonVi").ToString(), Int32))), "")
                    lbl_quoctich.Text = IIf(dr("IdQuocTich").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from QuocGia Where id = {0} and Status = 1", CType(dr("IdQuocTich").ToString(), Int32))), "")
                    lbl_dantoc.Text = IIf(dr("IdDanToc").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 21 and Status = 1", CType(dr("IdDanToc").ToString(), Int32))), "")
                    lbl_tongiao.Text = IIf(dr("IdTonGiao").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 24 and Status = 1", CType(dr("IdTonGiao").ToString(), Int32))), "")
                    lbl_socmt.Text = dr("CMT_So").ToString().Trim()
                    If dr("CMT_NgayCap").ToString() <> "" Then lbl_ngcap_cmt.Text = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_noicap_cmt.Text = dr("CMT_NoiCap").ToString().Trim()
                    lbl_honhan.Text = dr("HonNhan_HT").ToString().Trim()
                    '_HS_CanBo.Show_Picture(dr("AnhThe").ToString(), picbx_main)
                    My_LoadImage(picbx_main, dr("AnhThe"))

                    '---> Tab - Chi tiết Sơ yếu lý lịch
                    If (dr("IdNS_Tinh").ToString() <> "" And dr("IdNS_Tinh").ToString() <> "0") Then
                        lbl_ns_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc = 0 and Status = 1", CType(dr("IdNS_Tinh").ToString(), Int32)))
                    End If
                    If (dr("IdNS_Huyen").ToString() <> "" And dr("IdNS_Huyen").ToString() <> "0") Then
                        lbl_ns_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdNS_Huyen").ToString(), Int32)))
                    End If
                    lbl_ns_diachi.Text = dr("NS_DChi").ToString().Trim()

                    If (dr("IdNQ_Tinh").ToString() <> "" And dr("IdNQ_Tinh").ToString() <> "0") Then
                        lbl_nq_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc = 0 and Status = 1", CType(dr("IdNQ_Tinh").ToString(), Int32)))
                    End If
                    If (dr("IdNQ_Huyen").ToString() <> "" And dr("IdNQ_Huyen").ToString() <> "0") Then
                        lbl_nq_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdNQ_Huyen").ToString(), Int32)))
                    End If
                    lbl_nq_diachi.Text = dr("NQ_DChi").ToString().Trim()

                    If (dr("IdThT_Tinh").ToString() <> "" And dr("IdThT_Tinh").ToString() <> "0") Then
                        lbl_tt_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc = 0 and Status = 1", CType(dr("IdThT_Tinh").ToString(), Int32)))
                    End If
                    If (dr("IdThT_Huyen").ToString() <> "" And dr("IdThT_Huyen").ToString() <> "0") Then
                        lbl_tt_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdThT_Huyen").ToString(), Int32)))
                    End If
                    lbl_tt_diachi.Text = dr("ThT_Diachi").ToString().Trim()
                    lbl_tt_dienthoai.Text = dr("ThT_Dienthoai").ToString().Trim()

                    If (dr("IdTTr_Tinh").ToString() <> "" And dr("IdTTr_Tinh").ToString() <> "0") Then
                        lbl_ttr_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc = 0 and Status = 1", CType(dr("IdTTr_Tinh").ToString(), Int32)))
                    End If
                    If (dr("IdTTr_Huyen").ToString() <> "" And dr("IdTTr_Huyen").ToString() <> "0") Then
                        lbl_ttr_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdTTr_Huyen").ToString(), Int32)))
                    End If
                    lbl_ttr_diachi.Text = dr("TTr_Diachi").ToString().Trim()
                    lbl_ttr_dienthoai.Text = dr("TTr_Dienthoai").ToString().Trim()

                    lbl_dt_coquan.Text = dr("DienThoai_CQ").ToString().Trim()
                    lbl_dt_nharieng.Text = dr("DienThoai_NR").ToString().Trim()
                    lbl_dt_didong.Text = dr("DienThoai_DD").ToString().Trim()
                    lbl_sofax.Text = dr("SoFax").ToString().Trim()
                    lbl_email.Text = dr("Emai").ToString().Trim()
                    lbl_nhommau.Text = dr("NhomMau").ToString().Trim()

                    '---> Tab - Chi tiết Bản thân mối quan hệ
                    lbl_tp_giadinh.Text = IIf(dr("IdThanhPhanGD").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 13 and Status = 1", CType(dr("IdThanhPhanGD").ToString(), Int32))), "")
                    lbl_ut_giadinh.Text = IIf(dr("IdUT_GDinh").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 19 and Status = 1", CType(dr("IdUT_GDinh").ToString(), Int32))), "")
                    ' 'Load dữ liệu Ưu tiên bản thân
                    If (dr("IdUT_BThan").ToString() <> "") Then
                        _Globals.SetItems_CheckedListBox(clb_ut_banthan, dr("IdUT_BThan").ToString(), arr_Ut_Banthan)
                    End If

                    lbl_dacdiem_banthan.Text = dr("DacDiem_BT").ToString().Trim()
                    lbl_quanhe_ncngoai.Text = dr("QuanHe_Nguoi_NN").ToString().Trim()

                    '---> Tab - Chi tiết Thông tin khác của hồ sơ
                    lbl_tdvh.Text = IIf(dr("IdTrinhDoVH").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 6 and Status = 1", CType(dr("IdTrinhDoVH").ToString(), Int32))), "")
                    lbl_tdct.Text = IIf(dr("IdTrinhDoCT").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 36 and Status = 1", CType(dr("IdTrinhDoCT").ToString(), Int32))), "")

                    If dr("Ngay_ThamNien").ToString() <> "" Then lbl_ng_thamnien.Text = IIf(CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    If dr("Ngay_NH").ToString() <> "" Then lbl_ng_vaonganh.Text = IIf(CType(dr("Ngay_NH").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_NH").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    If dr("Ngay_VBSP").ToString() <> "" Then lbl_ng_vaonhcsxh.Text = IIf(CType(dr("Ngay_VBSP").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_VBSP").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    If dr("Ngay_BienChe").ToString() <> "" Then lbl_ng_bienche_nh.Text = IIf(CType(dr("Ngay_BienChe").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_BienChe").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    If dr("Ngay_CQ").ToString() <> "" Then lbl_ng_vao_cq.Text = IIf(CType(dr("Ngay_CQ").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_CQ").ToString(), DateTime).ToString("dd-MM-yyyy"), "")

                    If dr("CM_Ngay").ToString() <> "" Then lbl_ng_cachmang.Text = IIf(CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_tochuc_cm.Text = dr("CM_ToChuc").ToString().Trim()

                    lbl_sotruong_ct.Text = dr("SoTruong_CT").ToString().Trim()
                    lbl_congviec_lamlau.Text = dr("CV_Lau").ToString().Trim()

                    lbl_nh_makh.Text = dr("NH_MaKH").ToString().Trim()
                    lbl_nh_sotk.Text = dr("NH_SoTK").ToString().Trim()
                    lbl_nh_tennh.Text = dr("NH_TenNH").ToString().Trim()
                    lbl_ms_thue.Text = dr("MaSoThue").ToString().Trim()

                    lbl_bhxh_soso.Text = dr("BHXH_SoSo").ToString().Trim()
                    If dr("BHXH_NgaySo").ToString() <> "" Then lbl_bhxh_ngaylam.Text = IIf(CType(dr("BHXH_NgaySo").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("BHXH_NgaySo").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    If dr("BHXH_NgayBatDau").ToString() <> "" Then lbl_bhxh_ngaydong.Text = IIf(CType(dr("BHXH_NgayBatDau").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("BHXH_NgayBatDau").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_bhxh_noilam.Text = dr("BHXH_NoiLam").ToString().Trim()
                    lbl_ghichu.Text = dr("GhiChu").ToString().Trim()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị chi tiết thông tin hồ sơ cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Load thông tin về quyết định nhân sự gần đây nhất của một cán bộ
    ''' </summary>
    ''' <param name="_CanBoId">Id cán bộ</param>
    ''' <remarks></remarks>
    Private Sub Load_Infor_Decision(ByVal _CanBoId As String)
        Try
            Using db As DataTable = _HS_CanBo.GetAll_Decision(_IdCanBo, _IdDonviHT, 0)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        lbl_qd_soqd.Text = db.Rows(0)("So_QD").ToString().Trim()
                        lbl_qd_loaiqd.Text = IIf(db.Rows(0)("IdLoaiQD").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 15 and Status = 1", CType(db.Rows(0)("IdLoaiQD").ToString(), Int32))), "")

                        If db.Rows(0)("NgayKy_QD").ToString() <> "" Then lbl_qd_ngayky.Text = IIf(CType(db.Rows(0)("NgayKy_QD").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(0)("NgayKy_QD").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        If db.Rows(0)("NgayHL").ToString() <> "" Then lbl_qd_ngayhl.Text = IIf(CType(db.Rows(0)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(0)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")

                        lbl_qd_nguoiky.Text = db.Rows(0)("NguoiKy_QD").ToString().Trim()
                        lbl_qd_chucvu.Text = IIf(db.Rows(0)("IdCV_Nguoiky_QD").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} and id_goc = 14 and Status = 1", CType(db.Rows(0)("IdCV_Nguoiky_QD").ToString(), Int32))), "")
                        lbl_qd_donvi_moi.Text = IIf(db.Rows(0)("IdDonvi_Moi").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from ChiNhanh Where id = {0} and Status = 1", CType(db.Rows(0)("IdDonvi_Moi").ToString(), Int32))), "")
                        lbl_qd_phongban_moi.Text = IIf(db.Rows(0)("IdPhong_Moi").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_phong from PhongBan Where id = {0} and Status = 1", CType(db.Rows(0)("IdPhong_Moi").ToString(), Int32))), "")
                        lbl_qd_chucvu_moi.Text = IIf(db.Rows(0)("IdChucvu_Moi").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} and id_goc = 14 and Status = 1", CType(db.Rows(0)("IdChucvu_Moi").ToString(), Int32))), "")
                        lbl_qd_chmon_moi.Text = IIf(db.Rows(0)("IdChuyenMon_Moi").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} and id_goc = 12 and Status = 1", CType(db.Rows(0)("IdChuyenMon_Moi").ToString(), Int32))), "")

                        lbl_qd_ghichu.Text = db.Rows(0)("GhiChu").ToString().Trim()
                    End If
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Hiển thị chi tiết thông tin Quyết định nhân sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Load thông tin chi tiết về các hồ sơ khác của cán bộ
    ''' </summary>
    ''' <param name="_HumanId">Id cán bộ</param>
    ''' <remarks></remarks>
    Private Sub Load_Infor_HsKhac(ByVal _HumanId As String)
        Try
            If (_IdCanBo <> "") Then
                Dim strSQL As String = ""
                'Fill data - Hồ sơ công tác của cán bộ
                dgv_congtac.Rows.Clear()
                'strSQL = String.Format("Select * from HS_CongTac Where IdCanBo = '{0}' Order by TuNgay Asc", _HumanId)
                'Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                Using db As DataTable = listQDNhansu(_IdCanBo, 2)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim row As Integer = 0
                            For i As Integer = 0 To db.Rows.Count - 1
                                'Dim ThoiGian As String = ""
                                'If i = 0 Then
                                '    dgv_congtac.Rows.Add()
                                '    dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                '    ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                '    Dim j As Integer = i + 1
                                '    While j < db.Rows.Count
                                '        Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                '            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516", "1532", "1534", "1535"
                                '                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                '                i = j - 1
                                '                Exit While
                                '            Case Else
                                '                j = j + 1
                                '        End Select
                                '    End While
                                '    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                '    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                '    row = row + 1
                                'Else
                                '    If Not (db.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
                                '            Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
                                '            Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
                                '        dgv_congtac.Rows.Add()
                                '        dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                '        ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                '        If Not (db.Rows(i)("DenNgay") Is DBNull.Value) Then
                                '            If Not (CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                '                ThoiGian = ThoiGian.Substring(3) & " - " & CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                '                dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                '                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                '            Else
                                '                If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                                '                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                '                Else
                                '                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                '                End If
                                '                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                '                Dim j As Integer = i + 1
                                '                While j < db.Rows.Count
                                '                    Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                '                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516", "1532", "1534", "1535"
                                '                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                '                            i = j - 1
                                '                            Exit While
                                '                        Case Else
                                '                            j = j + 1
                                '                    End Select
                                '                End While
                                '                dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                '            End If
                                '        Else
                                '            If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                                '                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                '            Else
                                '                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                '            End If
                                '            dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                '            Dim j As Integer = i + 1
                                '            While j < db.Rows.Count
                                '                Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                '                    Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1516", "1532", "1534", "1535"
                                '                        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                '                        i = j - 1
                                '                        Exit While
                                '                    Case Else
                                '                        j = j + 1
                                '                End Select
                                '            End While
                                '            dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                '        End If
                                '        row = row + 1
                                '    End If
                                'End If
                                Dim ThoiGian As String = ""
                                If CInt(db.Rows(i)("IsQD_NHCS")) = 0 Then
                                    If i = 0 Then
                                        dgv_congtac.Rows.Add()
                                        dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                        ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        Dim j As Integer = i + 1
                                        While j < db.Rows.Count
                                            Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                                Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    i = j - 1
                                                    Exit While
                                                Case Else
                                                    j = j + 1
                                            End Select
                                        End While
                                        dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                        row = row + 1
                                    Else
                                        If Not (db.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
                                                Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
                                                Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
                                            dgv_congtac.Rows.Add()
                                            dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                            ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                            If Not (db.Rows(i)("DenNgay") Is DBNull.Value) Then
                                                If Not (CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                                    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                Else
                                                    If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                                                        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                                    Else
                                                        dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                    End If
                                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                    Dim j As Integer = i + 1
                                                    While j < db.Rows.Count
                                                        Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                                            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                                i = j - 1
                                                                Exit While
                                                            Case Else
                                                                j = j + 1
                                                        End Select
                                                    End While
                                                    dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                                End If
                                            Else
                                                If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1508" Or db.Rows(i)("MaLoaiQD").ToString().Trim() = "1509" Then
                                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = "Nhân viên " & IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                                Else
                                                    dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                End If
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                                Dim j As Integer = i + 1
                                                While j < db.Rows.Count
                                                    Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            i = j - 1
                                                            Exit While
                                                        Case Else
                                                            j = j + 1
                                                    End Select
                                                End While
                                                dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                            End If
                                            row = row + 1
                                        End If
                                    End If
                                Else
                                    If db.Rows(i)("MaLoaiQD").ToString().Trim() <> "1519" Then
                                        dgv_congtac.Rows.Add()
                                        dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                        ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        Dim j As Integer = i + 1
                                        While j < db.Rows.Count
                                            Select Case db.Rows(j)("MaLoaiQD").ToString().Trim()
                                                Case "1534", "1535"
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    Exit While
                                                Case "1518"
                                                    'Bỏ qua nếu QĐ tiếp theo là QĐ kiêm nhiệm
                                                    Exit While
                                                Case "1519"
                                                    'Lấy mốc hết thời gian kiêm nhiệm
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    j = j + 1
                                                    Exit While
                                                Case Else
                                                    'TH không đọc tiếp nếu: Nếu đơn vị, phòng ban, chức vụ cũ; hoặc QĐ tiếp theo là QĐ kiêm nhiệm
                                                    If (CInt(db.Rows(i)("IDDonVi_Moi")) = CInt(db.Rows(j)("IDDonVi_Moi")) And CInt(db.Rows(i)("IdPhong_Moi")) = CInt(db.Rows(j)("IdPhong_Moi")) And CInt(db.Rows(i)("IDChucVu_Moi")) = CInt(db.Rows(j)("IDChucVu_Moi"))) Then
                                                        j = j + 1
                                                    Else
                                                        'Nếu QĐ trước đó la QD Kiêm nhiệm mà QĐ tiếp theo ko phải là QĐ Thôi kiêm nhiệm thì bỏ qua
                                                        If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                            Exit While
                                                        Else
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            i = j - 1
                                                            Exit While
                                                        End If
                                                    End If
                                            End Select
                                        End While
                                        dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                        If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1534" Then
                                            dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("LoaiQD").ToString() <> "", db.Rows(i)("LoaiQD").ToString(), "")
                                        Else
                                            If db.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = db.Rows(i)("LoaiQD").ToString() & " " & IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString().Replace("(CN cấp I & tương đương)", ""), "")
                                            Else
                                                dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", ""), "")
                                            End If
                                        End If
                                        row = row + 1
                                        If j = db.Rows.Count Then Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End Using

                'Fill data - Hồ sơ xuất ngoại của cán bộ ra lưới dữ liệu
                dgv_xuatngoai.Rows.Clear()
                strSQL = String.Format("Select a.*,b.ten_goi As Nuoc_Den, c.ten_goi As Chuc_Vu from HS_XuatNgoai As a, QuocGia As b, DanhMuc As c Where (a.IdNuocDen = b.Id) and (c.Id = a.IdCV_NguoiKy_QD and c.id_goc = 14) And a.IdCanBo = '{0}' Order by TuNgay Asc", _HumanId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                dgv_xuatngoai.Rows.Add()
                                dgv_xuatngoai.Rows(i).Cells(0).Value = CType(i + 1, String)
                                dgv_xuatngoai.Rows(i).Cells(1).Value = IIf(db.Rows(i)("IdXuatNgoai").ToString() <> "", db.Rows(i)("IdXuatNgoai").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells(2).Value = IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_xuatngoai.Rows(i).Cells(3).Value = IIf(db.Rows(i)("DenNgay").ToString().Trim() <> "", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_xuatngoai.Rows(i).Cells(4).Value = IIf(db.Rows(i)("Nuoc_Den").ToString() <> "", db.Rows(i)("Nuoc_Den").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells(5).Value = IIf(db.Rows(i)("MucDich").ToString() <> "", db.Rows(i)("MucDich").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells(6).Value = IIf(db.Rows(i)("SoQD").ToString() <> "", db.Rows(i)("SoQD").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells(7).Value = IIf(db.Rows(i)("NgayKy_QD").ToString().Trim() <> "", CType(db.Rows(i)("NgayKy_QD").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_xuatngoai.Rows(i).Cells(8).Value = IIf(db.Rows(i)("NguoiKy_QD").ToString() <> "", db.Rows(i)("NguoiKy_QD").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells(9).Value = IIf(db.Rows(i)("Chuc_Vu").ToString() <> "", db.Rows(i)("Chuc_Vu").ToString(), "")
                                dgv_xuatngoai.Rows(i).Cells(10).Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Hồ sơ Hộ chiếu của cán bộ
                dgv_hochieu.Rows.Clear()
                strSQL = String.Format("Select * from HS_HoChieu Where IdCanBo = '{0}' Order by NgayCap Asc", _HumanId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                dgv_hochieu.Rows.Add()
                                dgv_hochieu.Rows(i).Cells(0).Value = CType(i + 1, String)
                                dgv_hochieu.Rows(i).Cells(1).Value = IIf(db.Rows(i)("IdHoChieu").ToString() <> "", db.Rows(i)("IdHoChieu").ToString(), "")
                                dgv_hochieu.Rows(i).Cells(2).Value = IIf(db.Rows(i)("So_HoChieu").ToString() <> "", db.Rows(i)("So_HoChieu").ToString(), "")
                                If db.Rows(i)("Loai_HC").ToString() = "1" Then
                                    dgv_hochieu.Rows(i).Cells(3).Value = "Phổ thông"
                                ElseIf db.Rows(i)("Loai_HC").ToString() = "2" Then
                                    dgv_hochieu.Rows(i).Cells(3).Value = "Công vụ"
                                ElseIf db.Rows(i)("Loai_HC").ToString() = "3" Then
                                    dgv_hochieu.Rows(i).Cells(3).Value = "Ngoại giao"
                                Else : dgv_hochieu.Rows(i).Cells(3).Value = ""
                                End If

                                dgv_hochieu.Rows(i).Cells(4).Value = IIf(db.Rows(i)("NgayCap").ToString().Trim() <> "", CType(db.Rows(i)("NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_hochieu.Rows(i).Cells(5).Value = IIf(db.Rows(i)("NoiCap").ToString() <> "", db.Rows(i)("NoiCap").ToString(), "")
                                dgv_hochieu.Rows(i).Cells(6).Value = IIf(db.Rows(i)("NgayHH").ToString().Trim() <> "", CType(db.Rows(i)("NgayHH").ToString(), DateTime).ToString("dd-MM-yyyy"), "")

                                If db.Rows(i)("TinhTrang").ToString() = "1" Then
                                    dgv_hochieu.Rows(i).Cells(7).Value = "Còn hiệu lực"
                                ElseIf db.Rows(i)("TinhTrang").ToString() = "2" Then
                                    dgv_hochieu.Rows(i).Cells(7).Value = "Hết hạn"
                                ElseIf db.Rows(i)("TinhTrang").ToString() = "3" Then
                                    dgv_hochieu.Rows(i).Cells(7).Value = "Mất"
                                Else : dgv_hochieu.Rows(i).Cells(7).Value = ""
                                End If

                                dgv_hochieu.Rows(i).Cells(8).Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Hồ sơ lực lượng vũ trang

                dgv_llvt.Rows.Clear()
                strSQL = String.Format("Select a.*,b.ten_goi as Phan_Loai from HS_LLVT as a,DanhMuc as b Where (a.IdLoaiLLVT = b.id and b.id_goc = 41) And IdCanBo = '{0}' Order by TuNgay Asc", _HumanId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                dgv_llvt.Rows.Add()
                                dgv_llvt.Rows(i).Cells(0).Value = CType(i + 1, String)
                                dgv_llvt.Rows(i).Cells(1).Value = IIf(db.Rows(i)("IdHSLLVT").ToString() <> "", db.Rows(i)("IdHSLLVT").ToString(), "")
                                dgv_llvt.Rows(i).Cells(2).Value = IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_llvt.Rows(i).Cells(3).Value = IIf(db.Rows(i)("DenNgay").ToString().Trim() <> "", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_llvt.Rows(i).Cells(4).Value = IIf(db.Rows(i)("Phan_Loai").ToString() <> "", db.Rows(i)("Phan_Loai").ToString(), "")
                                dgv_llvt.Rows(i).Cells(5).Value = IIf(db.Rows(i)("IdQuanHam").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 2 and Status = 1", CType(db.Rows(i)("IdQuanHam").ToString(), Int32))), "")
                                dgv_llvt.Rows(i).Cells(6).Value = IIf(db.Rows(i)("ChucVu").ToString() <> "", db.Rows(i)("ChucVu").ToString(), "")
                                dgv_llvt.Rows(i).Cells(7).Value = IIf(db.Rows(i)("DonVi").ToString() <> "", db.Rows(i)("DonVi").ToString(), "")
                                dgv_llvt.Rows(i).Cells(8).Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Hồ sơ cũ của cán bộ
                'dgv_hosocu.Rows.Clear()
                'strSQL = String.Format("Select * From HS_Cu Where IdCanBo = '{0}' Order by TuThang Asc", _HumanId)
                'Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '    If Not (db Is Nothing) Then
                '        If (db.Rows.Count > 0) Then
                '            For i As Int32 = 0 To db.Rows.Count - 1
                '                dgv_hosocu.Rows.Add()
                '                dgv_hosocu.Rows(i).Cells(0).Value = CType(i + 1, String)
                '                dgv_hosocu.Rows(i).Cells(1).Value = IIf(db.Rows(i)("IdHSCu").ToString() <> "", db.Rows(i)("IdHSCu").ToString(), "")
                '                dgv_hosocu.Rows(i).Cells(2).Value = IIf(db.Rows(i)("TuThang").ToString().Trim() <> "", CType(db.Rows(i)("TuThang").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                '                dgv_hosocu.Rows(i).Cells(3).Value = IIf(db.Rows(i)("DenThang").ToString().Trim() <> "", CType(db.Rows(i)("DenThang").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                '                dgv_hosocu.Rows(i).Cells(4).Value = IIf(db.Rows(i)("DiaChi").ToString() <> "", db.Rows(i)("DiaChi").ToString(), "")
                '                dgv_hosocu.Rows(i).Cells(5).Value = IIf(db.Rows(i)("NgheNghiep").ToString() <> "", db.Rows(i)("NgheNghiep").ToString(), "")
                '                dgv_hosocu.Rows(i).Cells(6).Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                '            Next
                '        End If
                '    End If
                'End Using

                'Fill data - Hồ sơ học hàm của cán bộ
                dgv_hocham.Rows.Clear()
                strSQL = String.Format("Select a.IdCB_HocHam,b.ten_goi as Hoc_Ham From CB_HocHam a, DanhMuc b Where (a.IdHocHam = b.id and b.id_goc = 26) and IdCanBo = '{0}' Order by IdHocHam Asc", _HumanId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Int32 = 0 To db.Rows.Count - 1
                                dgv_hocham.Rows.Add()
                                dgv_hocham.Rows(i).Cells(0).Value = CType(i + 1, String)
                                dgv_hocham.Rows(i).Cells(1).Value = IIf(db.Rows(i)("IdCB_HocHam").ToString() <> "", db.Rows(i)("IdCB_HocHam").ToString(), "")
                                dgv_hocham.Rows(i).Cells(2).Value = IIf(db.Rows(i)("Hoc_Ham").ToString() <> "", db.Rows(i)("Hoc_Ham").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Hồ sơ học vị của cán bộ
                dgv_hocvi.Rows.Clear()
                strSQL = String.Format("Select a.IdCB_HocVi,b.ten_goi as Hoc_Vi From CB_HocVi a, DanhMuc b Where (a.IdHocVi = b.id and b.id_goc = 20) and IdCanBo = '{0}' Order by IdHocVi Asc", _HumanId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Int32 = 0 To db.Rows.Count - 1
                                dgv_hocvi.Rows.Add()
                                dgv_hocvi.Rows(i).Cells(0).Value = CType(i + 1, String)
                                dgv_hocvi.Rows(i).Cells(1).Value = IIf(db.Rows(i)("IdCB_HocVi").ToString() <> "", db.Rows(i)("IdCB_HocVi").ToString(), "")
                                dgv_hocvi.Rows(i).Cells(2).Value = IIf(db.Rows(i)("Hoc_Vi").ToString() <> "", db.Rows(i)("Hoc_Vi").ToString(), "")
                            Next
                        End If
                    End If
                End Using

            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị chi tiết thông tin hồ sơ khác: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Function - Reset controls - Thiết lập trạng thái ban đầu cho các controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetControls()
        lbl_macb.Text = ""
        lbl_hoten.Text = ""
        lbl_ten_tg.Text = ""
        lbl_bidanh.Text = ""
        lbl_gioitinh.Text = ""
        lbl_ngaysinh.Text = ""
        lbl_donvi.Text = ""
        lbl_quoctich.Text = ""
        lbl_dantoc.Text = ""
        lbl_socmt.Text = ""
        lbl_ngcap_cmt.Text = ""
        lbl_noicap_cmt.Text = ""
        If Not (picbx_main.Image Is Nothing) Then
            picbx_main.Image.Dispose()
            picbx_main.Image = Nothing
        End If
        Dim strdefault As String = System.IO.Path.Combine(Application.StartupPath, "ANH_HOSO") + "//"
        strdefault = String.Format("{0}{1}", strdefault, "default.jpg")
        If (System.IO.File.Exists(strdefault)) Then
            picbx_main.Image = Image.FromFile(strdefault)
            picbx_main.SizeMode = PictureBoxSizeMode.StretchImage
        End If

        'Tab - Sơ yếu lý lịch
        lbl_ns_tinh.Text = ""
        lbl_ns_huyen.Text = ""
        lbl_ns_diachi.Text = ""

        lbl_nq_tinh.Text = ""
        lbl_nq_huyen.Text = ""
        lbl_nq_diachi.Text = ""

        lbl_tt_tinh.Text = ""
        lbl_tt_huyen.Text = ""
        lbl_tt_diachi.Text = ""
        lbl_tt_dienthoai.Text = ""

        lbl_ttr_tinh.Text = ""
        lbl_ttr_huyen.Text = ""
        lbl_ttr_diachi.Text = ""
        lbl_ttr_dienthoai.Text = ""

        lbl_dt_didong.Text = ""
        lbl_dt_coquan.Text = ""
        lbl_dt_nharieng.Text = ""
        lbl_sofax.Text = ""
        lbl_email.Text = ""
        lbl_nhommau.Text = ""
        lbl_tp_giadinh.Text = ""
        lbl_ut_giadinh.Text = ""
        _Globals.ResetItems_CheckedListBox(clb_ut_banthan)
        lbl_dacdiem_banthan.Text = ""
        lbl_quanhe_ncngoai.Text = ""
        lbl_ghichu.Text = ""

        'Tab - Thông tin khác của Hồ sơ
        lbl_tdvh.Text = ""
        lbl_tdct.Text = ""
        lbl_ng_thamnien.Text = ""
        lbl_ng_vaonganh.Text = ""
        lbl_ng_vaonhcsxh.Text = ""
        lbl_ng_bienche_nh.Text = ""
        lbl_ng_vao_cq.Text = ""
        lbl_ng_cachmang.Text = ""
        lbl_tochuc_cm.Text = ""
        lbl_sotruong_ct.Text = ""
        lbl_congviec_lamlau.Text = ""

        lbl_nh_makh.Text = ""
        lbl_nh_sotk.Text = ""
        lbl_nh_tennh.Text = ""
        lbl_ms_thue.Text = ""

        lbl_bhxh_soso.Text = ""
        lbl_bhxh_ngaylam.Text = ""
        lbl_bhxh_ngaydong.Text = ""
        lbl_bhxh_noilam.Text = ""

        'Tab - Quyết định Nhân sự của cán bộ
        lbl_qd_soqd.Text = ""
        lbl_qd_loaiqd.Text = ""
        lbl_qd_ngayky.Text = ""
        lbl_qd_ngayhl.Text = ""
        'lbl_qd_ngay_bntt.Text = ""
        'lbl_qd_ngaytl.Text = ""
        lbl_qd_nguoiky.Text = ""
        lbl_qd_chucvu.Text = ""

        'lbl_qd_donvi_cu.Text = ""
        'lbl_qd_phongban_cu.Text = ""
        'lbl_qd_chucvu_cu.Text = ""
        'lbl_qd_chmon_cu.Text = ""

        lbl_qd_donvi_moi.Text = ""
        lbl_qd_phongban_moi.Text = ""
        lbl_qd_chucvu_moi.Text = ""
        lbl_qd_chmon_moi.Text = ""
        lbl_qd_ghichu.Text = ""
        lbl_honhan.Text = ""
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill số bản ghi tìm được vào combobox
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Fill_Record()
        cb_records.Items.Clear()
        For i As Byte = 1 To _Records
            cb_records.Items.Add(i)
        Next
        If (cb_records.Items.Count > 0) Then
            cb_records.SelectedIndex = _RecordCurrent
        End If
    End Sub
#End Region

#Region "---> Events main <---"
    Private Sub frmHS_ChiTiet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If tctrl_main.TabPages.Contains(tp_hs_cu) Then tctrl_main.TabPages.Remove(tp_hs_cu)
        If tctrl_hscanbo.TabPages.Contains(tp_hs_qdnhansu) Then tctrl_hscanbo.TabPages.Remove(tp_hs_qdnhansu)
        _HS_CanBo.Create_Frame(dgv_congtac, 7)
        _HS_CanBo.Create_Frame(dgv_xuatngoai, 8)
        _HS_CanBo.Create_Frame(dgv_hochieu, 9)
        _HS_CanBo.Create_Frame(dgv_llvt, 10)
        '_HS_CanBo.Create_Frame(dgv_hosocu, 11)
        _HS_CanBo.Create_Frame(dgv_hocham, 12)
        _HS_CanBo.Create_Frame(dgv_hocvi, 13)
        'clb_ut_banthan
        'Load danh sách ưu tiên bản thân
        clb_ut_banthan.Items.Clear()
        arr_Ut_Banthan.Clear()
        arr_Ut_Banthan = _Globals.FillData_CheckedListBox(clb_ut_banthan, clsHT_DanhMuc.Sql_Ut_banthan)
        clb_ut_banthan.Enabled = False
        Fill_Record()
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub frmHS_ChiTiet_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Dim _totalRemain As Int32 = pnl_ttc_1.Width + pnl_ttc_2.Width + pnl_pic.Width + 16
        pnl_ttc_1.Width = (_totalRemain - (pnl_pic.Width + 16)) / 2
        pnl_ttc_2.Width = (_totalRemain - (pnl_pic.Width + 16)) / 2
    End Sub

    'Private Sub dgv_congtac_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_congtac.CellClick
    '    lbl_ct_tungay.Text = ""
    '    lbl_ct_denngay.Text = ""
    '    lbl_ct_chucvu.Text = ""
    '    lbl_ct_diachi.Text = ""
    '    lbl_ct_lydonv.Text = ""
    '    lbl_ct_ghichu.Text = ""
    '    If (dgv_congtac.Rows.Count > 0) Then
    '        If ((dgv_congtac.CurrentRow.Cells(1).Value IsNot Nothing) And (dgv_congtac.CurrentRow.Cells(1).Value.ToString() <> "")) Then
    '            lbl_ct_tungay.Text = dgv_congtac.CurrentRow.Cells("cln_Tungay").Value.ToString().Trim()
    '            lbl_ct_denngay.Text = dgv_congtac.CurrentRow.Cells("cln_Denngay").Value.ToString().Trim()
    '            lbl_ct_chucvu.Text = dgv_congtac.CurrentRow.Cells("cln_Chucvu").Value.ToString().Trim()
    '            lbl_ct_diachi.Text = dgv_congtac.CurrentRow.Cells("cln_Diachi").Value.ToString().Trim()
    '            lbl_ct_lydonv.Text = dgv_congtac.CurrentRow.Cells("cln_Lydo").Value.ToString().Trim()
    '            lbl_ct_ghichu.Text = dgv_congtac.CurrentRow.Cells("cln_Ghichu").Value.ToString().Trim()
    '        End If
    '    End If
    'End Sub

    'Private Sub dgv_congtac_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_congtac.KeyUp
    '    dgv_congtac_CellClick(sender, Nothing)
    'End Sub

    Private Sub dgv_xuatngoai_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_xuatngoai.CellClick
        lbl_xn_tungay.Text = ""
        lbl_xn_denngay.Text = ""
        lbl_xn_nuocden.Text = ""
        lbl_xn_mucdich.Text = ""
        lbl_xn_soqd.Text = ""
        lbl_xn_ngayky.Text = ""
        lbl_xn_nguoiky.Text = ""
        lbl_xn_chucvu.Text = ""
        lbl_xn_ghichu.Text = ""
        If (dgv_xuatngoai.Rows.Count > 0) Then
            If ((dgv_xuatngoai.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_xuatngoai.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                lbl_xn_tungay.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Tungay").Value.ToString().Trim()
                lbl_xn_denngay.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Denngay").Value.ToString().Trim()
                lbl_xn_nuocden.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Quocgia").Value.ToString().Trim()
                lbl_xn_mucdich.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Mucdich").Value.ToString().Trim()
                lbl_xn_soqd.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Soqd").Value.ToString().Trim()
                lbl_xn_ngayky.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Ngayky").Value.ToString().Trim()
                lbl_xn_nguoiky.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Nguoiky").Value.ToString().Trim()
                lbl_xn_chucvu.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Chucvu").Value.ToString().Trim()
                lbl_xn_ghichu.Text = dgv_xuatngoai.CurrentRow.Cells("cln_Ghichu").Value.ToString().Trim()
            End If
        End If
    End Sub

    Private Sub dgv_xuatngoai_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_xuatngoai.KeyUp
        dgv_xuatngoai_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_hochieu_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hochieu.CellClick
        lbl_hc_sohc.Text = ""
        lbl_hc_loaihc.Text = ""
        lbl_hc_ngaycap.Text = ""
        lbl_hc_noicap.Text = ""
        lbl_hc_ngayhh.Text = ""
        lbl_hc_tinhtrang.Text = ""
        lbl_hc_ghichu.Text = ""

        If (dgv_hochieu.Rows.Count > 0) Then
            If ((dgv_hochieu.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hochieu.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                lbl_hc_sohc.Text = dgv_hochieu.CurrentRow.Cells("cln_Sohc").Value.ToString().Trim()
                lbl_hc_loaihc.Text = dgv_hochieu.CurrentRow.Cells("cln_Loaihc").Value.ToString().Trim()
                lbl_hc_ngaycap.Text = dgv_hochieu.CurrentRow.Cells("cln_Ngaycap").Value.ToString().Trim()
                lbl_hc_noicap.Text = dgv_hochieu.CurrentRow.Cells("cln_Noicap").Value.ToString().Trim()
                lbl_hc_ngayhh.Text = dgv_hochieu.CurrentRow.Cells("cln_Ngayhh").Value.ToString().Trim()
                lbl_hc_tinhtrang.Text = dgv_hochieu.CurrentRow.Cells("cln_Tinhtrang").Value.ToString().Trim()
                lbl_hc_ghichu.Text = dgv_hochieu.CurrentRow.Cells("cln_Ghichu").Value.ToString().Trim()
            End If
        End If
    End Sub

    Private Sub dgv_hochieu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hochieu.KeyUp
        dgv_hochieu_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_llvt_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_llvt.CellClick
        lbl_llvt_tungay.Text = ""
        lbl_llvt_denngay.Text = ""
        lbl_llvt_phanloai.Text = ""
        lbl_llvt_quanham.Text = ""
        lbl_llvt_chucvu.Text = ""
        lbl_llvt_donvi.Text = ""
        lbl_llvt_ghichu.Text = ""

        If (dgv_llvt.Rows.Count > 0) Then
            If ((dgv_llvt.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_llvt.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                lbl_llvt_tungay.Text = dgv_llvt.CurrentRow.Cells("cln_Tungay").Value.ToString().Trim()
                lbl_llvt_denngay.Text = dgv_llvt.CurrentRow.Cells("cln_Denngay").Value.ToString().Trim()
                lbl_llvt_phanloai.Text = dgv_llvt.CurrentRow.Cells("cln_Phanloai").Value.ToString().Trim()
                lbl_llvt_quanham.Text = dgv_llvt.CurrentRow.Cells("cln_Quanham").Value.ToString().Trim()
                lbl_llvt_chucvu.Text = dgv_llvt.CurrentRow.Cells("cln_Chucvu").Value.ToString().Trim()
                lbl_llvt_donvi.Text = dgv_llvt.CurrentRow.Cells("cln_Donvi").Value.ToString().Trim()
                lbl_llvt_ghichu.Text = dgv_llvt.CurrentRow.Cells("cln_Ghichu").Value.ToString().Trim()
            End If
        End If
    End Sub

    Private Sub dgv_llvt_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_llvt.KeyUp
        dgv_llvt_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_hosocu_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hosocu.CellClick
        lbl_hscu_tuthang.Text = ""
        lbl_hscu_denthang.Text = ""
        lbl_hscu_diachi.Text = ""
        lbl_hscu_nghenghiep.Text = ""
        lbl_hscu_ghichu.Text = ""

        If (dgv_hosocu.Rows.Count > 0) Then
            If ((dgv_hosocu.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hosocu.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                lbl_hscu_tuthang.Text = dgv_hosocu.CurrentRow.Cells("cln_Tuthang").Value.ToString().Trim()
                lbl_hscu_denthang.Text = dgv_hosocu.CurrentRow.Cells("cln_Denthang").Value.ToString().Trim()
                lbl_hscu_diachi.Text = dgv_hosocu.CurrentRow.Cells("cln_Diachi").Value.ToString().Trim()
                lbl_hscu_nghenghiep.Text = dgv_hosocu.CurrentRow.Cells("cln_Nghenghiep").Value.ToString().Trim()
                lbl_hscu_ghichu.Text = dgv_hosocu.CurrentRow.Cells("cln_Ghichu").Value.ToString().Trim()
            End If
        End If
    End Sub

    Private Sub dgv_hosocu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hosocu.KeyUp
        dgv_hosocu_CellClick(sender, Nothing)
    End Sub

    Private Sub frmHS_ChiTiet_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub tctrl_main_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tctrl_main.SelectedIndexChanged
        Select Case tctrl_main.SelectedTab.Name
            'Case "tp_congtac"
            '    dgv_congtac_CellClick(sender, Nothing)
            Case "tp_xuatngoai"
                dgv_xuatngoai_CellClick(sender, Nothing)
            Case "tp_hochieu"
                dgv_hochieu_CellClick(sender, Nothing)
            Case "tp_llvt"
                dgv_llvt_CellClick(sender, Nothing)
            Case "tp_hs_cu"
                dgv_hosocu_CellClick(sender, Nothing)
            Case Else
        End Select
    End Sub

    Private Sub cb_records_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_records.SelectedIndexChanged
        If (cb_records.Items.Count <> 0) Then
            _IdCanBo = arr_RecordId(cb_records.SelectedIndex).ToString()
            Load_Infor_PersonelFile(_IdCanBo)
            'Load_Infor_Decision(_IdCanBo)
            Load_Infor_HsKhac(_IdCanBo)
        End If
    End Sub

    Private Sub btn_first_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_first.Click
        If (cb_records.SelectedIndex = 0) Then
            Return
        Else
            Try
                cb_records.SelectedIndex = 0
                _IdCanBo = arr_RecordId(cb_records.SelectedIndex).ToString()
            Catch ex As Exception
                MessageBox.Show("Di chuyển về bản ghi đầu: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub btn_previous_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_previous.Click
        If (cb_records.SelectedIndex = 0) Then
            Return
        Else
            Try
                cb_records.SelectedIndex -= 1
                _IdCanBo = arr_RecordId(cb_records.SelectedIndex).ToString()
            Catch ex As Exception
                MessageBox.Show("Di chuyển về bản phía trước: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub btn_next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_next.Click
        If (cb_records.SelectedIndex = cb_records.Items.Count - 1) Then
            Return
        Else
            Try
                cb_records.SelectedIndex += 1
                _IdCanBo = arr_RecordId(cb_records.SelectedIndex).ToString()
            Catch ex As Exception
                MessageBox.Show("Di chuyển về bản phía sau: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub btn_last_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_last.Click
        If (cb_records.SelectedIndex = cb_records.Items.Count - 1) Then
            Return
        Else
            Try
                cb_records.SelectedIndex = cb_records.Items.Count - 1
                _IdCanBo = arr_RecordId(cb_records.SelectedIndex).ToString()
            Catch ex As Exception
                MessageBox.Show("Di chuyển về bản cuối cùng: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub frmHS_ChiTiet_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If (picbx_main.Image IsNot Nothing) Then
            picbx_main.Image.Dispose()
            picbx_main.Image = Nothing
            Dim strdefault As String = System.IO.Path.Combine(Application.StartupPath, "ANH_HOSO") + "//"
            strdefault = String.Format("{0}{1}", strdefault, "default.jpg")
            If (System.IO.File.Exists(strdefault)) Then
                picbx_main.Image = Image.FromFile(strdefault)
                picbx_main.SizeMode = PictureBoxSizeMode.StretchImage
            End If
        End If
    End Sub
#End Region
    
   
End Class