Public Class frmHS_NghiHuu_ChuyenCT_ChiTiet

#Region "---> Defined parametter and properties <---"
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()
    'Khai báo và kh?i t?o thu?c tính c?a các ??i t??ng
    Private _IdCanBo As String
    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal value As String)
            _IdCanBo = value
        End Set
    End Property

    'Khai báo các bi?n ?? di chuy?n b?n ghi
    Private _Records As Int16    'T?ng s? b?n ghi tìm th?y
    Public Property Records() As Int16
        Get
            Return _Records
        End Get
        Set(ByVal value As Int16)
            _Records = value
        End Set
    End Property

    Private _RecordCurrent As Byte    'S? th? t? b?n ghi select
    Public Property RecordCurrent() As Byte
        Get
            Return _RecordCurrent
        End Get
        Set(ByVal value As Byte)
            _RecordCurrent = value
        End Set
    End Property

    ''' <summary>
    ''' M?ng l?u danh sách các id c?a h? s? tìm ???c
    ''' </summary>
    ''' <remarks></remarks>
    Public arr_RecordId As ArrayList = New ArrayList()
    'Bi?n l?u ??n v? Hi?n t?i c?a Cán b? (Load thông tin quy?t ??nh nhân s? m?i ?ây nh?t c?a cán b?)
    Public _IdDonviHT As Integer = 0
    Private arr_Ut_Banthan As ArrayList = New ArrayList
#End Region

#Region "---> Functions main <---"

    'Mảng danh sách cần cho tab - Quyết định
    Private arrQd_LoaiQd As ArrayList = New ArrayList()
    Private arrQd_Chucvu As ArrayList = New ArrayList()
    Private arrQd_Chucvu_moi As ArrayList = New ArrayList()
    Private arrQd_ChMon_moi As ArrayList = New ArrayList()
    Private arrQd_Phongban_moi As ArrayList = New ArrayList()
    Private arrQd_Donvi_moi As ArrayList = New ArrayList()
    Private arrQd_LyDoTV As ArrayList = New ArrayList()
    Private curr_IDQD As String = ""
    Private curr_table_key As String = ""
    Private curr_table_name As String = ""

    ''' <summary>
    ''' Hàm load thông tin Chi tiết về Hồ sơ cán bộ
    ''' </summary>
    ''' <param name="_Id">Id cán b?</param>
    ''' <remarks></remarks>
    Private Sub Load_Infor_PersonelFile(ByVal _Id As String)
        Try
            ResetControls()
            Dim dr As DataRow
            dr = _HS_CanBo.GetHuman_ForCode(_Id)
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then

                    lbl_macb.Text = dr("MaCB").ToString().Trim()
                    lbl_hoten.Text = dr("HoTen").ToString().Trim()
                    If (dr("GioiTinh").ToString() <> "") Then lbl_gioitinh.Text = IIf(dr("GioiTinh").ToString() = False, " Nam", " N?")
                    If dr("NgaySinh").ToString() <> "" Then lbl_ngaysinh.Text = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_dantoc.Text = IIf(dr("IdDanToc").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 21 and Status = 1", CType(dr("IdDanToc").ToString(), Int32))), "")
                    lbl_tongiao.Text = IIf(dr("IdTonGiao").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 24 and Status = 1", CType(dr("IdTonGiao").ToString(), Int32))), "")
                    lbl_socmt.Text = dr("CMT_So").ToString().Trim()
                    If dr("CMT_NgayCap").ToString() <> "" Then lbl_ngcap_cmt.Text = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    lbl_noicap_cmt.Text = dr("CMT_NoiCap").ToString().Trim()

                    'My_LoadImage(picbx_main, dr("AnhThe").ToString())
                    My_LoadImage(picbx_main, dr("AnhThe"))

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
                    If (dr("IdTTr_Tinh").ToString() <> "" And dr("IdTTr_Tinh").ToString() <> "0") Then
                        lbl_ttr_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc = 0 and Status = 1", CType(dr("IdTTr_Tinh").ToString(), Int32)))
                    End If
                    If (dr("IdTTr_Huyen").ToString() <> "" And dr("IdTTr_Huyen").ToString() <> "0") Then
                        lbl_ttr_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdTTr_Huyen").ToString(), Int32)))
                    End If
                    lbl_ttr_diachi.Text = dr("TTr_Diachi").ToString().Trim()

                    lbl_dt_coquan.Text = dr("DienThoai_CQ").ToString().Trim()
                    lbl_dt_nharieng.Text = dr("DienThoai_NR").ToString().Trim()
                    lbl_dt_didong.Text = dr("DienThoai_DD").ToString().Trim()
                    lbl_email.Text = dr("Emai").ToString().Trim()

                    lbl_tp_giadinh.Text = IIf(dr("IdThanhPhanGD").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 13 and Status = 1", CType(dr("IdThanhPhanGD").ToString(), Int32))), "")
                    lbl_ut_giadinh.Text = IIf(dr("IdUT_GDinh").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 19 and Status = 1", CType(dr("IdUT_GDinh").ToString(), Int32))), "")

                    If (dr("IdUT_BThan").ToString() <> "") Then
                        _Globals.SetItems_CheckedListBox(clb_ut_banthan, dr("IdUT_BThan").ToString(), arr_Ut_Banthan)
                    End If

                    lbl_dacdiem_banthan.Text = dr("DacDiem_BT").ToString().Trim()
                    lbl_quanhe_ncngoai.Text = dr("QuanHe_Nguoi_NN").ToString().Trim()

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
                    lbl_ghichu.Text = dr("GhiChu").ToString().Trim()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị chi tiết thông tin hồ sơ cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Load thông tin về quyết định nhân sự và thôi việc của một cán bộ
    ''' </summary>
    ''' <param name="_CanBoId">Id cán bộ</param>
    ''' <remarks></remarks>
    Private Sub Load_Infor_Decision(ByVal _CanBoId As String)
        Try
            dgv_QuyetDinh.Rows.Clear()
            Using db As DataTable = listQDNhansu(_CanBoId, 3)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        Dim _rows As Int32 = db.Rows.Count - 1
                        Dim i As Integer
                        'Lấy thông tin QĐ chuyển công tác or QĐ chấm dứt hợp đồng (thôi việc) cuối cùng
                        curr_IDQD = db.Rows(0)("IdQD").ToString()
                        curr_table_key = db.Rows(0)("table_key").ToString()
                        curr_table_name = db.Rows(0)("table_name").ToString()
                        For i = 0 To _rows
                            dgv_QuyetDinh.Rows.Add()
                            dgv_QuyetDinh.Rows(i).Cells(0).Value = db.Rows(i)("IdQD").ToString()
                            dgv_QuyetDinh.Rows(i).Cells(1).Value = CType(i + 1, String)
                            dgv_QuyetDinh.Rows(i).Cells(2).Value = IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                            dgv_QuyetDinh.Rows(i).Cells(3).Value = IIf(db.Rows(i)("So_QD").ToString() <> "", db.Rows(i)("So_QD").ToString(), "")
                            dgv_QuyetDinh.Rows(i).Cells(4).Value = IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                            dgv_QuyetDinh.Rows(i).Cells(5).Value = IIf(db.Rows(i)("LoaiQD").ToString() <> "", db.Rows(i)("LoaiQD").ToString(), "")
                            dgv_QuyetDinh.Rows(i).Cells(6).Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                        Next
                    End If
                End If
            End Using

            If curr_table_name.ToUpper = "QDNHANSU" Then
                'Load form và thông tin hồ sơ của cán bộ chuyển đơn vị 
                FillData_ComboBox(1)
                bntExportHSCB.Visible = True
                bntExportHSCB.Enabled = True
                Dim m_QDNhansu As QDNhanSu = New QDNhanSu
                m_QDNhansu = m_QDNhansu.getRecord(curr_IDQD)
                edt_qd_soqd.Text = m_QDNhansu.So_QD
                cb_qd_loaiqd.SelectedIndex = IIf(m_QDNhansu.IdLoaiQD.ToString() <> "", CType(arrQd_LoaiQd.IndexOf(m_QDNhansu.IdLoaiQD.ToString()), Int32), 0)
                dtpk_qd_ngayky.Value = m_QDNhansu.NgayKy_QD
                dtpk_qd_ngayhl.Value = m_QDNhansu.NgayHL
                edt_qd_nguoiky.Text = m_QDNhansu.NguoiKy_QD
                cb_qd_chucvu.SelectedIndex = IIf(m_QDNhansu.idCV_Nguoiky_QD.ToString() <> "", CType(arrQd_Chucvu.IndexOf(m_QDNhansu.idCV_Nguoiky_QD.ToString()), Int32), 0)
                cb_qd_donvi_moi.SelectedIndex = IIf(m_QDNhansu.IdDonvi_Moi.ToString() <> "", CType(arrQd_Donvi_moi.IndexOf(m_QDNhansu.IdDonvi_Moi.ToString()), Int32), 0)
                cb_qd_phongban_moi.SelectedIndex = IIf(m_QDNhansu.IdPhong_Moi.ToString() <> "", CType(arrQd_Phongban_moi.IndexOf(m_QDNhansu.IdPhong_Moi.ToString()), Int32), 0)
                cb_qd_chucvu_moi.SelectedIndex = IIf(m_QDNhansu.IdChucvu_Moi.ToString() <> "", CType(arrQd_Chucvu_moi.IndexOf(m_QDNhansu.IdChucvu_Moi.ToString()), Int32), 0)
                cb_qd_cm_moi.SelectedIndex = IIf(m_QDNhansu.IdChuyenMon_Moi.ToString() <> "", CType(arrQd_ChMon_moi.IndexOf(m_QDNhansu.IdChuyenMon_Moi.ToString()), Int32), 0)
                txtDVraQD.Text = m_QDNhansu.DVraQD
                edt_qd_ghichu.Text = m_QDNhansu.GhiChu
            Else
                'Load form và thông tin cán bộ chấm dứt hợp đồng lao động (nghỉ hưu)
                FillData_ComboBox(2)
                bntExportHSCB.Visible = False
                Dim m_QDThoiViec As CBThoiViec = New CBThoiViec
                m_QDThoiViec = m_QDThoiViec.getRecord(curr_IDQD)
                edt_qd_soqd.Text = m_QDThoiViec.So_QD
                cb_qd_loaiqd.SelectedIndex = IIf(m_QDThoiViec.IdLoaiQD.ToString() <> "", CType(arrQd_LoaiQd.IndexOf(m_QDThoiViec.IdLoaiQD.ToString()), Int32), 0)
                edt_qd_nguoiky.Text = m_QDThoiViec.NguoiKy_QD
                dtpk_qd_ngayky.Value = m_QDThoiViec.NgayKy_QD
                dtpk_qd_ngayhl.Value = m_QDThoiViec.Ngay_HL
                cboLyDo_TV.SelectedIndex = IIf(m_QDThoiViec.IdLyDo.ToString() <> "", CType(arrQd_LyDoTV.IndexOf(m_QDThoiViec.IdLyDo.ToString()), Int32), 0)
                txtDVraQD.Text = m_QDThoiViec.DVraQD
                cb_qd_chucvu.SelectedIndex = IIf(m_QDThoiViec.IdCV_NguoKy_QD.ToString() <> "", CType(arrQd_Chucvu.IndexOf(m_QDThoiViec.IdCV_NguoKy_QD.ToString()), Int32), 0)
                edt_qd_ghichu.Text = m_QDThoiViec.GhiChu
            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị chi tiết thông tin Quyết định: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện Load thông tin chi tiết về các hồ sơ khác của cán bộ
    ''' </summary>
    ''' <param name="_HumanId">Id cán bo</param>
    ''' <remarks></remarks>
    Private Sub Load_Infor_HsKhac(ByVal _HumanId As String)
        Try
            If (_IdCanBo <> "") Then
                Dim strSQL As String = ""
                'Fill data - Quá trình công tác của cán bộ
                dgv_congtac.Rows.Clear()
                Using db As DataTable = listQDNhansu(_IdCanBo, 2)
                    If (db.Rows.Count > 0) Then
                        Dim row As Integer = 0
                        For i As Integer = 0 To db.Rows.Count - 1
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
                                'dgv_congtac.Rows.Add()
                                'dgv_congtac.Rows(row).Cells("cln_STT").Value = row + 1
                                'ThoiGian = "Từ " & IIf(db.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                'Dim j As Integer = i + 1
                                'While j < db.Rows.Count
                                '    If (CInt(db.Rows(i)("IDDonVi_Moi")) = CInt(db.Rows(j)("IDDonVi_Moi")) And CInt(db.Rows(i)("IdPhong_Moi")) = CInt(db.Rows(j)("IdPhong_Moi")) And CInt(db.Rows(i)("IDChucVu_Moi")) = CInt(db.Rows(j)("IDChucVu_Moi"))) Then
                                '        j = j + 1
                                '    Else
                                '        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                '        i = j - 1
                                '        Exit While
                                '    End If
                                'End While
                                'dgv_congtac.Rows(row).Cells("cln_Tungay").Value = ThoiGian
                                'dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                                'row = row + 1
                                'If j = db.Rows.Count Then Exit For
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
                End Using

                'Fill data - Lương
                dgv_Luong.Rows.Clear()
                Using db As DataTable = listQDLuong(_HumanId)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            Dim i As Integer
                            For i = 0 To _rows
                                dgv_Luong.Rows.Add()
                                dgv_Luong.Rows(i).Cells(0).Value = CType(i + 1, String)
                                dgv_Luong.Rows(i).Cells(1).Value = IIf(db.Rows(i)("Ngay_Huong").ToString().Trim() <> "", CType(db.Rows(i)("Ngay_Huong").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                dgv_Luong.Rows(i).Cells(2).Value = IIf(db.Rows(i)("SoQD").ToString() <> "", db.Rows(i)("SoQD").ToString(), "")
                                dgv_Luong.Rows(i).Cells(3).Value = IIf(db.Rows(i)("DVraQD").ToString() <> "", db.Rows(i)("DVraQD").ToString(), "")
                                dgv_Luong.Rows(i).Cells(4).Value = IIf(db.Rows(i)("LoaiQD").ToString() <> "", db.Rows(i)("LoaiQD").ToString(), "")
                                dgv_Luong.Rows(i).Cells(5).Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Qua trinh sinh hoat Dang
                dgv_Dang.Rows.Clear()
                Using db As DataTable = _SqlHelper.SelectDBRows("SELECT t2.TuNgay, t2.DenNgay, (case t2.IskiemNhiem when 1 then N'Kiêm nhiệm' else N'' end ) as KhiemNhiem, (Select ten_goi From DanhMuc Where id_goc=16 and id=t2.IdCVDang) as CVDang, t2.Chibo FROM HS_dangvien t1 join HS_Dang t2 on t1.IdDangVien=t2.IdDangVien WHERE t1.idcanbo='" & _HumanId & "'  order by TuNgay ")
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                dgv_Dang.Rows.Add()
                                dgv_Dang.Rows(i).Cells(0).Value = CType(i + 1, String)
                                Dim tg As String = ""
                                tg += IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                If (CType(db.Rows(i)("DenNgay").ToString().Trim(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                    tg = "Từ " & tg
                                Else
                                    tg += " - " & IIf(db.Rows(i)("DenNgay").ToString().Trim() <> "", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                End If
                                dgv_Dang.Rows(i).Cells(1).Value = tg
                                dgv_Dang.Rows(i).Cells(2).Value = IIf(db.Rows(i)("KhiemNhiem").ToString() <> "", db.Rows(i)("KhiemNhiem").ToString(), "")
                                dgv_Dang.Rows(i).Cells(3).Value = IIf(db.Rows(i)("CVDang").ToString() <> "", db.Rows(i)("CVDang").ToString(), "")
                                dgv_Dang.Rows(i).Cells(4).Value = IIf(db.Rows(i)("Chibo").ToString() <> "", db.Rows(i)("Chibo").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Quá trình đào tạo văn bằng chứng chỉ
                dgv_DTVBCC.Rows.Clear()
                strSQL = "SELECT IdDTVBCC, TuNgay, DenNgay, NamTN, (SELECT ten_goi FROM DanhMuc WHERE id=IdHinhThucDT and id_goc=8) as HinhThucDT, " & _
                          " (SELECT ten_goi FROM DanhMuc WHERE id=IdTrinhDo and id_goc=38) as TrinhDo, (SELECT ten_goi FROM DanhMuc WHERE id=IdChuyenNganhDT and id_goc=11) as ChuyenNganhDT, NganhHoc, CoSo_DT" & _
                          " FROM HS_DTVBCC WHERE idCanbo='" & _HumanId & "' order by Tungay desc, NamTN desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                dgv_DTVBCC.Rows.Add()
                                dgv_DTVBCC.Rows(i).Cells(0).Value = CType(i + 1, String)
                                Dim tg As String = ""
                                If db.Rows(i)("TuNgay") Is DBNull.Value Then
                                    tg = db.Rows(i)("NamTN").ToString().Trim()
                                Else
                                    tg += IIf(db.Rows(i)("TuNgay").ToString().Trim() <> "", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                    If db.Rows(i)("DenNgay") Is DBNull.Value Then
                                        tg = "Từ " & tg
                                    Else
                                        tg += " - " & IIf(db.Rows(i)("DenNgay") Is DBNull.Value, "", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"))
                                    End If
                                End If
                                dgv_DTVBCC.Rows(i).Cells(1).Value = tg
                                dgv_DTVBCC.Rows(i).Cells(2).Value = IIf(db.Rows(i)("HinhThucDT").ToString() <> "", db.Rows(i)("HinhThucDT").ToString(), "")
                                dgv_DTVBCC.Rows(i).Cells(3).Value = IIf(db.Rows(i)("TrinhDo").ToString() <> "", db.Rows(i)("TrinhDo").ToString(), "")
                                dgv_DTVBCC.Rows(i).Cells(4).Value = IIf(db.Rows(i)("ChuyenNganhDT").ToString() <> "", db.Rows(i)("ChuyenNganhDT").ToString(), "")
                                dgv_DTVBCC.Rows(i).Cells(5).Value = IIf(db.Rows(i)("NganhHoc").ToString() <> "", db.Rows(i)("NganhHoc").ToString(), "")
                                dgv_DTVBCC.Rows(i).Cells(6).Value = IIf(db.Rows(i)("CoSo_DT").ToString() <> "", db.Rows(i)("CoSo_DT").ToString(), "")
                            Next
                        End If
                    End If
                End Using

                'Fill data - Hồ sơ Gia đình cán bộ
                '
                dgv_gdcb.Rows.Clear()
                strSQL = String.Format("SELECT a.*,b.ten_goi as QuanHe FROM  HS_GDCB a, DanhMuc b WHERE (IdCanBo = '{0}') And (a.IdQuanHe = b.Id And b.id_goc = 23 And b.Status = 1) Order by HoTen,QuanHe Asc", _HumanId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_gdcb.Rows.Add()
                                dgv_gdcb.Rows(i).Cells("cln_Id").Value = db.Rows(i)("IdTVien").ToString().Trim()
                                dgv_gdcb.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                                dgv_gdcb.Rows(i).Cells("cln_HoTen").Value = db.Rows(i)("HoTen").ToString().Trim()
                                If (db.Rows(i)("GioiTinh").ToString().Trim() = "") Then
                                    dgv_gdcb.Rows(i).Cells("cln_GioiTinh").Value = ""
                                Else
                                    dgv_gdcb.Rows(i).Cells("cln_GioiTinh").Value = IIf(db.Rows(i)("GioiTinh").ToString().Trim() = False, "Nam", "Nữ")
                                End If

                                dgv_gdcb.Rows(i).Cells("cln_NamSinh").Value = db.Rows(i)("NamSinh").ToString().Trim()
                                'Lấy tên gọi của danh mục Quan hệ từ chỉ số
                                dgv_gdcb.Rows(i).Cells("cln_QuanHe").Value = db.Rows(i)("QuanHe").ToString().Trim()
                                If (db.Rows(i)("ConMat").ToString().Trim() = "") Then
                                    dgv_gdcb.Rows(i).Cells("cln_TinhTrang").Value = ""
                                Else
                                    dgv_gdcb.Rows(i).Cells("cln_TinhTrang").Value = IIf(db.Rows(i)("ConMat").ToString().Trim() = False, "Còn sống", "Đã mất")
                                End If
                                dgv_gdcb.Rows(i).Cells("cln_DienThoai").Value = db.Rows(i)("DienThoai").ToString().Trim()
                                dgv_gdcb.Rows(i).Cells("cln_DiaChi").Value = db.Rows(i)("DiaChi").ToString().Trim()
                                dgv_gdcb.Rows(i).Cells("cln_NgheNghiep").Value = db.Rows(i)("NgheNghiep").ToString().Trim()
                            Next
                        End If
                    End If
                End Using

            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị thông tin hồ sơ cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Function - Reset controls - Thi?t l?p tr?ng thái ban ??u cho các controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetControls()
        lbl_macb.Text = ""
        lbl_hoten.Text = ""
        lbl_gioitinh.Text = ""
        lbl_ngaysinh.Text = ""
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

        'Tab - S? y?u lý l?ch
        lbl_ns_tinh.Text = ""
        lbl_ns_huyen.Text = ""
        lbl_ns_diachi.Text = ""

        lbl_nq_tinh.Text = ""
        lbl_nq_huyen.Text = ""
        lbl_nq_diachi.Text = ""

        lbl_tt_tinh.Text = ""
        lbl_tt_huyen.Text = ""
        lbl_tt_diachi.Text = ""

        lbl_ttr_tinh.Text = ""
        lbl_ttr_huyen.Text = ""
        lbl_ttr_diachi.Text = ""

        lbl_dt_didong.Text = ""
        lbl_dt_coquan.Text = ""
        lbl_dt_nharieng.Text = ""
        lbl_email.Text = ""
        lbl_tp_giadinh.Text = ""
        lbl_ut_giadinh.Text = ""
        _Globals.ResetItems_CheckedListBox(clb_ut_banthan)
        lbl_dacdiem_banthan.Text = ""
        lbl_quanhe_ncngoai.Text = ""
        lbl_ghichu.Text = ""

        'Tab - Thông tin khác c?a H? s?
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

    End Sub

    ''' <summary>
    ''' fill combobox
    ''' </summary>
    ''' <param name="vKind">1: fill cac combobox trong QD nhan su
    '''                     2: fill cac combobox trong QD thoi viec  
    ''' </param>
    ''' <remarks></remarks>
    Private Sub FillData_ComboBox(ByVal vKind As Integer)

        arrQd_Chucvu.Clear()
        cb_qd_chucvu.Items.Clear()
        arrQd_Chucvu = _Globals.Bind_ComBoBox(cb_qd_chucvu, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ người ký---")

        If vKind = 1 Then
            ' QD Nhan su
            Panel37.Visible = True
            Panel39.Visible = True
            pnlLyDoTV.Visible = False

            arrQd_LoaiQd.Clear()
            cb_qd_loaiqd.Items.Clear()
            arrQd_LoaiQd = _Globals.Bind_ComBoBox(cb_qd_loaiqd, clsHT_DanhMuc.Sql_LoaiQd_Ns, "---Loại quyết định---")

            arrQd_Donvi_moi.Clear()
            cb_qd_donvi_moi.Items.Clear()
            arrQd_Donvi_moi = _Globals.Bind_ComBoBox(cb_qd_donvi_moi, " SELECT id as Value, (CASE WHEN (id_goc>5) Then ('  '+ten_goi) ELSE ten_goi END) as Display FROM chinhanh Where Status = 1 Order by id_goc,Ma_so", "---Đơn vị mới---")

            arrQd_Chucvu_moi.Clear()
            cb_qd_chucvu_moi.Items.Clear()
            arrQd_Chucvu_moi = _Globals.Bind_ComBoBox(cb_qd_chucvu_moi, clsHT_DanhMuc.Sql_Chucvu, "---Chức vụ---")

            arrQd_ChMon_moi.Clear()
            cb_qd_cm_moi.Items.Clear()
            arrQd_ChMon_moi = _Globals.Bind_ComBoBox(cb_qd_cm_moi, clsHT_DanhMuc.Sql_Chuyenmon, "---Chuyên môn nghiệp vụ---")
        Else
            ' QD thoi viec  
            Panel37.Visible = False
            Panel39.Visible = False
            pnlLyDoTV.Visible = True

            arrQd_LoaiQd.Clear()
            cb_qd_loaiqd.Items.Clear()
            arrQd_LoaiQd = _Globals.Bind_ComBoBox(cb_qd_loaiqd, "Select id,ten_goi from DanhMuc Where id_goc = 37 and Status = 1 Order by ten_goi Asc", "---Loại quyết định---")

            arrQd_LyDoTV.Clear()
            cboLyDo_TV.Items.Clear()
            arrQd_LyDoTV = _Globals.Bind_ComBoBox(cboLyDo_TV, "Select id,ten_goi from DanhMuc Where id_goc = 22 and Status = 1 Order by ten_goi Asc", "---Lý do thôi việc---")
        End If
    End Sub

    Private Sub blankTabFrm()
        labAlert.Text = ""
        curr_IDQD = ""
        edt_qd_soqd.Text = ""
        dtpk_qd_ngayky.Value = Now
        dtpk_qd_ngayhl.Value = Now
        edt_qd_nguoiky.Text = ""
        txtDVraQD.Text = ""
        edt_qd_ghichu.Text = ""
    End Sub

    Private Sub updateQuyetDinh()
        If curr_table_name.ToUpper = "QDNHANSU" Then
            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
            m_QDNhanSu.IdQDNhanSu = curr_IDQD
            m_QDNhanSu.IdCanBo = _IdCanBo
            m_QDNhanSu.So_QD = edt_qd_soqd.Text.Trim.ToString()
            m_QDNhanSu.NgayKy_QD = dtpk_qd_ngayky.Value
            m_QDNhanSu.IdLoaiQD = CType(IIf(arrQd_LoaiQd.Count > 0, arrQd_LoaiQd(cb_qd_loaiqd.SelectedIndex), "0"), Int32)
            m_QDNhanSu.NgayHL = dtpk_qd_ngayhl.Value
            m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
            m_QDNhanSu.NgayThoiLuong = DateTime.MinValue
            m_QDNhanSu.NguoiKy_QD = standardizeName(edt_qd_nguoiky.Text.Trim.ToString())
            m_QDNhanSu.idCV_Nguoiky_QD = CType(IIf(arrQd_Chucvu.Count > 0, arrQd_Chucvu(cb_qd_chucvu.SelectedIndex), "0"), Int32)
            m_QDNhanSu.IdDonvi_Cu = 0
            m_QDNhanSu.IdPhong_Cu = 0
            m_QDNhanSu.IdChucvu_Cu = 0
            m_QDNhanSu.IdChuyenMon_Cu = 0
            m_QDNhanSu.IdDonvi_Moi = CType(IIf(arrQd_Donvi_moi.Count > 0, arrQd_Donvi_moi(cb_qd_donvi_moi.SelectedIndex), "0"), Int32)
            If (cb_qd_phongban_moi.Items.Count <> 0) Then
                m_QDNhanSu.IdPhong_Moi = CType(IIf(arrQd_Phongban_moi.Count > 0, arrQd_Phongban_moi(cb_qd_phongban_moi.SelectedIndex), "0"), Int32)
            Else : m_QDNhanSu.IdPhong_Moi = 0
            End If
            m_QDNhanSu.IdChucvu_Moi = CType(IIf(arrQd_Chucvu_moi.Count > 0, arrQd_Chucvu_moi(cb_qd_chucvu_moi.SelectedIndex), "0"), Int32)
            m_QDNhanSu.IdChuyenMon_Moi = CType(IIf(arrQd_ChMon_moi.Count > 0, arrQd_ChMon_moi(cb_qd_cm_moi.SelectedIndex), "0"), Int32)
            m_QDNhanSu.GhiChu = standardizeString(edt_qd_ghichu.Text.Trim.ToString())
            m_QDNhanSu.Active = 1
            m_QDNhanSu.IsKiemNhiem = 0
            m_QDNhanSu.IsQD_NHCS = 1
            m_QDNhanSu.DVraQD = getDonvi(IdDONVI)
            m_QDNhanSu.CV_NguoiKy_QD = ""
            m_QDNhanSu.LoaiQD = ""
            m_QDNhanSu.DenNgay = DateTime.MinValue
            m_QDNhanSu.NoiDung = ""
            If m_QDNhanSu.IdQDNhanSu = "" Then
                m_QDNhanSu.Add()
            Else
                m_QDNhanSu.Update()
            End If
        Else
            Dim m_QDThoiViec As CBThoiViec = New CBThoiViec
            m_QDThoiViec.IdCBThoiViec = curr_IDQD
            m_QDThoiViec.IdCanBo = _IdCanBo
            m_QDThoiViec.IdLoaiQD = CType(IIf(arrQd_LoaiQd.Count > 0, arrQd_LoaiQd(cb_qd_loaiqd.SelectedIndex), "0"), Int32)
            m_QDThoiViec.IdLyDo = CType(IIf(arrQd_LyDoTV.Count > 0, arrQd_LyDoTV(cboLyDo_TV.SelectedIndex), "0"), Int32)
            m_QDThoiViec.NgayKy_QD = dtpk_qd_ngayky.Value
            m_QDThoiViec.Ngay_HL = dtpk_qd_ngayhl.Value
            m_QDThoiViec.NguoiKy_QD = standardizeName(edt_qd_nguoiky.Text.Trim.ToString())
            m_QDThoiViec.TroCap_ThoiViec = 0
            m_QDThoiViec.TroCap_Khac = 0
            m_QDThoiViec.SoTien_BoiThuong = 0
            m_QDThoiViec.SoTien_ThuHoi = 0
            m_QDThoiViec.So_QD = edt_qd_soqd.Text.Trim.ToString()
            m_QDThoiViec.DVraQD = standardizeString(txtDVraQD.Text.Trim)
            m_QDThoiViec.IsQD_NHCS = 1
            m_QDThoiViec.IdCV_NguoKy_QD = CType(IIf(arrQd_Chucvu.Count > 0, arrQd_Chucvu(cb_qd_chucvu.SelectedIndex), "0"), Int32)
            m_QDThoiViec.CV_NguoiKy_QD = ""
            m_QDThoiViec.GhiChu = standardizeString(edt_qd_ghichu.Text.Trim.ToString())
            If m_QDThoiViec.IdCBThoiViec = "" Then
                m_QDThoiViec.Add()
            Else
                m_QDThoiViec.Update()
            End If
        End If
    End Sub

#End Region

#Region "---> Events main <---"

    Private Sub frmHS_NghiHuu_ChuyenCT_ChiTiet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _HS_CanBo.Create_Frame(dgv_congtac, 7)
        _HS_CanBo.Create_Frame(dgv_Luong, 22)
        _HS_CanBo.Create_Frame(dgv_Dang, 20)
        _HS_CanBo.Create_Frame(dgv_DTVBCC, 21)
        _HS_CanBo.Create_Frame(dgv_gdcb, 14)
        _HS_CanBo.Create_Frame(dgv_QuyetDinh, 19)
        dgv_gdcb.Columns("cln_Choice").Visible = False
        dgv_gdcb.Columns("cln_DienThoai").Visible = False
        clb_ut_banthan.Items.Clear()
        arr_Ut_Banthan.Clear()
        arr_Ut_Banthan = _Globals.FillData_CheckedListBox(clb_ut_banthan, clsHT_DanhMuc.Sql_Ut_banthan)
        _IdCanBo = arr_RecordId(_RecordCurrent).ToString()
        bntNew.Visible = False
        bntAdd.Visible = False
        bntDelete.Visible = False
        bntExportHSCB.Visible = False
        Load_Infor_PersonelFile(_IdCanBo)
        Load_Infor_HsKhac(_IdCanBo)
        Load_Infor_Decision(_IdCanBo)
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub frmHS_NghiHuu_ChuyenCT_ChiTiet_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Dim _totalRemain As Int32 = pnl_ttc_1.Width + pnl_ttc_2.Width + pnl_pic.Width + 16
        pnl_ttc_1.Width = (_totalRemain - (pnl_pic.Width + 16)) / 2
        pnl_ttc_2.Width = (_totalRemain - (pnl_pic.Width + 16)) / 2
    End Sub

    Private Sub frmHS_NghiHuu_ChuyenCT_ChiTiet_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

        If curr_table_name.ToUpper = "QDNHANSU" Then
            bntExportHSCB.Visible = True
            bntExportHSCB.Enabled = True
        Else
            bntExportHSCB.Visible = False
        End If

        Select Case tctrl_main.SelectedTab.Name
            Case "tp_QDNhansu"
                If curr_table_name.ToUpper = "QDNHANSU" Then
                    bntNew.Visible = True
                    bntNew.Enabled = True
                Else
                    bntNew.Visible = False
                End If
                bntAdd.Visible = True
                bntDelete.Visible = True
                bntAdd.Enabled = True
                bntDelete.Enabled = True
            Case Else
                bntNew.Visible = False
                bntAdd.Visible = False
                bntDelete.Visible = False
                'Case "tp_xuatngoai"
                '    dgv_xuatngoai_CellClick(sender, Nothing)
                'Case "tp_hochieu"
                '    dgv_hochieu_CellClick(sender, Nothing)
                'Case "tp_llvt"
                '    dgv_llvt_CellClick(sender, Nothing)
        End Select
    End Sub

    Private Sub frmHS_NghiHuu_ChuyenCT_ChiTiet_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
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

    Private Sub cb_qd_donvi_moi_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_qd_donvi_moi.SelectedIndexChanged
        Try
            arrQd_Phongban_moi.Clear()
            cb_qd_phongban_moi.Items.Clear()
            If (cb_qd_donvi_moi.SelectedIndex > 0 And cb_qd_donvi_moi.Items.Count <> 0) Then
                Dim _DonviId As Int32 = CType(arrQd_Donvi_moi(IIf(cb_qd_donvi_moi.SelectedIndex > 0, cb_qd_donvi_moi.SelectedIndex, "0")), Int32)
                'Lấy mã chi nhánh từ Id chi nhánh
                Dim strSQL As String = ""
                strSQL = String.Format("Select * from ChiNhanh Where id = {0} and Status = 1", _DonviId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim tructhuocCurrent As String = _HS_CanBo.GetTrucThuoc(db.Rows(0)("ma_so").ToString().Trim())
                            strSQL = "Select id,ten_phong as ten_goi From PhongBan Where Charindex('" & tructhuocCurrent.ToString & "',truc_thuoc) > 0 and Status = 1 Order by Ma_so"
                            arrQd_Phongban_moi = _Globals.Bind_ComBoBox(cb_qd_phongban_moi, strSQL, "---Phòng ban---")
                        End If
                    End If
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show("Điền dữ liệu danh sách Phòng ban: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bnt2C_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt2C.Click
        Dim HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
        HS_CanBo.ExportWord_2C(_IdCanBo)
    End Sub

    Private Sub bntLLCN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntLLCN.Click
        Dim HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
        HS_CanBo.ExportWord_CurriculumVitae(_IdCanBo)
    End Sub

    Private Sub bntAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Try
            updateQuyetDinh()
            labAlert.Text = "Ghi dữ liệu thành công!"
        Catch ex As Exception
            labAlert.Text = "Ghi dữ liệu không thành công!"
        End Try
    End Sub

    Private Sub bntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew.Click
        blankTabFrm()
    End Sub

    Private Sub bntDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete.Click
        If curr_IDQD = "" Then
            MessageBox.Show("Thao tác bị hủy bỏ: Bạn đang thực hiện thêm mới quyết định -> không xóa được." & vbCr & "Nếu muốn xóa quyết định Thôi việc/Chuyển công tác hiện tại thực hiện:" & vbCr & " đóng và vào lại giao diện.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Exit Sub
        End If
        If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
            Try
                If curr_table_name.ToUpper = "QDNHANSU" Then
                    Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                    m_QDNhanSu.IdQDNhanSu = curr_IDQD
                    m_QDNhanSu.Delete()
                Else
                    Dim m_QDThoiViec As CBThoiViec = New CBThoiViec
                    m_QDThoiViec.IdCBThoiViec = curr_IDQD
                    m_QDThoiViec.Delete()
                End If
                MessageBox.Show("Xoá dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Close()
            Catch ex As Exception
                MessageBox.Show("Có lỗi trong quá trình xóa dữ liệu: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub bntExportHSCB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExportHSCB.Click
        Dim frmEx As frmExportHSCB = New frmExportHSCB
        Dim QDNhanSufinalNHCS As QDNhanSu = New QDNhanSu
        Dim dbconn As DBAccess = New DBAccess
        Dim IDCN As Integer
        QDNhanSufinalNHCS = QDNhanSufinalNHCS.getFinalRecord(IdCanBo, Nothing, True)
        frmEx.IDCB_Moved = QDNhanSufinalNHCS.IdCanBo
        IDCN = dbconn.getString("SELECT id FROM ChiNhanh WHERE (id=" & QDNhanSufinalNHCS.IdDonvi_Moi & " and id_goc=1) or (id=" & QDNhanSufinalNHCS.IdDonvi_Moi & " and id_goc=0) or (id in (SELECT id_goc FROM Chinhanh WHERE id=" & QDNhanSufinalNHCS.IdDonvi_Moi & ") and id_goc=1)")
        frmEx.IDCN_Moved = IDCN
        frmEx.TG_Chuyen = QDNhanSufinalNHCS.NgayHL.Date
        frmEx.ShowDialog()
    End Sub

#End Region

End Class