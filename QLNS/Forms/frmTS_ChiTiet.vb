Public Class frmTS_ChiTiet

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
    Private _Labour As clsHS_Hdld = New clsHS_Hdld()
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

    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

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
#End Region

#Region "---> Functions: Các hàm chính <---"
    ''' <summary>
    ''' Function - Reset controls - Thiết lập trạng thái ban đầu cho các controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls()
        lbl_macb.Text = ""
        lbl_hoten.Text = ""
        lbl_ten_tg.Text = ""
        lbl_bidanh.Text = ""
        lbl_gioitinh.Text = ""
        lbl_ngaysinh.Text = ""
        lbl_donvi.Text = ""
        lbl_phongban.Text = ""
        lbl_quoctich.Text = ""
        lbl_dantoc.Text = ""
        lbl_tongiao.Text = ""
        lbl_cmt_so.Text = ""
        lbl_cmt_ngaycap.Text = ""
        lbl_cmt_noicap.Text = ""
        lbl_honnhan.Text = ""
        If (picbx_main.Image IsNot Nothing) Then
            picbx_main.Image.Dispose()
            picbx_main.Image = Nothing
        End If

        lbl_ns_tinh.Text = ""
        lbl_ns_huyen.Text = ""
        lbl_ns_xa.Text = ""
        lbl_ns_thon.Text = ""
        lbl_ns_diachi.Text = ""

        lbl_nq_tinh.Text = ""
        lbl_nq_huyen.Text = ""
        lbl_nq_xa.Text = ""
        lbl_nq_thon.Text = ""
        lbl_nq_diachi.Text = ""

        lbl_tt_tinh.Text = ""
        lbl_tt_huyen.Text = ""
        lbl_tt_xa.Text = ""
        lbl_tt_thon.Text = ""
        lbl_tt_diachi.Text = ""
        lbl_tt_dienthoai.Text = ""

        lbl_ttr_tinh.Text = ""
        lbl_ttr_huyen.Text = ""
        lbl_ttr_xa.Text = ""
        lbl_ttr_thon.Text = ""
        lbl_ttr_diachi.Text = ""
        lbl_ttr_dienthoai.Text = ""

        lbl_dt_didong.Text = ""
        lbl_dt_coquan.Text = ""
        lbl_dt_nharieng.Text = ""
        lbl_sofax.Text = ""
        lbl_email.Text = ""
        lbl_nhommau.Text = ""
        lbl_tp_giadinh.Text = ""
        lbl_td_vanhoa.Text = ""
        lbl_td_chinhtri.Text = ""
        lbl_hocham.Text = ""
        lbl_ut_giadinh.Text = ""
        lbl_ut_banthan.Text = ""
        lbl_hocvi.Text = ""
        lbl_trd_chmon.Text = ""
        lbl_trd_ngngu.Text = ""
        lbl_trd_tinhoc.Text = ""
        lbl_chng_daotao.Text = ""

        lbl_cm_ngat_tg.Text = ""
        lbl_cm_tochuc.Text = ""

        lbl_dv_ngayvao.Text = ""
        lbl_dv_ngay_ct.Text = ""
        lbl_dv_ngayra.Text = ""
        lbl_dv_noi_kn.Text = ""
        lbl_dv_sothe.Text = ""
        lbl_dv_nguoi_gt.Text = ""
        lbl_dv_lydo.Text = ""

        lbl_tk_makh.Text = ""
        lbl_tk_sotk.Text = ""
        lbl_tk_tennh.Text = ""
        lbl_ms_thue.Text = ""
        lbl_bhxh_soso.Text = ""
        lbl_bhxh_ngaylam.Text = ""
        lbl_bhxh_ngaydong.Text = ""
        lbl_bhxh_noilam.Text = ""
        lbl_chng_daotao.Text = ""

        rb_hdld.Checked = False
        rb_qdtamtuyen.Checked = False
        lbl_hdld_sohd.Text = ""
        lbl_hdld_ngayky.Text = ""
        lbl_hdld_nguoiky.Text = ""
        lbl_hdld_chucvu.Text = ""

        lbl_tg_tungay.Text = ""
        lbl_tg_denngay.Text = ""
        lbl_loaihinh_cv.Text = ""
        lbl_gio_batdau.Text = ""
        lbl_gio_ketthuc.Text = ""
        lbl_ht_traluong.Text = ""
        lbl_cdld_bangluong.Text = ""
        lbl_cdld_bangluong_code.Text = ""
        lbl_cdld_ngachluong.Text = ""
        lbl_cdld_ngachluong_code.Text = ""
        lbl_cdld_bacluong.Text = ""
        lbl_cdld_hsluong.Text = ""
        lbl_cdld_tlhuong.Text = ""
        lbl_cv_damnhan.Text = ""
        lbl_st_congtac.Text = ""
        lbl_cv_lamlau.Text = ""
        lbl_ghichu.Text = ""
        ckb_choice_bhxh.Checked = False
        ckb_choice_bhyt.Checked = False

    End Sub

    ''' <summary>
    ''' Function fill data in controls - Hàm thực hiện fill dữ liệu theo id cán bộ truyền vào
    ''' </summary>
    ''' <param name="_IdCanBo">Id cán bộ</param>
    ''' <remarks></remarks>
    Private Sub Fill_Data(ByVal _IdCanBo As String)
        ResetAll_Controls()
        Dim dr As DataRow
        dr = _HS_CanBo.GetTrainee(0, _IdCanBo, "", "", CType("1", Byte), CType("1", Byte))
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                lbl_macb.Text = dr("MaCB").ToString()
                lbl_hoten.Text = dr("HoTen").ToString()
                lbl_ten_tg.Text = dr("TenThuongGoi").ToString()
                lbl_bidanh.Text = dr("BiDanh").ToString()
                lbl_ten_tg.Text = dr("TenThuongGoi").ToString()
                lbl_gioitinh.Text = dr("GioiTinh_HT").ToString()
                lbl_ngaysinh.Text = dr("NgaySinh_HT").ToString()
                lbl_donvi.Text = dr("ChiNhanh_HT").ToString()
                lbl_phongban.Text = dr("DonVi_HT").ToString()
                'Lấy thông tin Phòng ban
                lbl_quoctich.Text = dr("QuocTich_HT").ToString()        'IIf(dr("IdQuocTich").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from QuocGia Where id = {0} and Status = 1", CType(dr("IdQuocTich").ToString(), Integer))), "")
                lbl_dantoc.Text = dr("DanToc_HT").ToString()        'IIf(dr("IdDanToc").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 21 and Status = 1", CType(dr("IdDanToc").ToString(), Integer))), "")
                lbl_tongiao.Text = dr("TonGiao_HT").ToString()        'IIf(dr("IdTonGiao").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 24 and Status = 1", CType(dr("IdTonGiao").ToString(), Integer))), "")

                lbl_cmt_so.Text = dr("CMT_So").ToString().Trim()
                If dr("CMT_NgayCap").ToString() <> "" Then lbl_cmt_ngaycap.Text = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                lbl_cmt_noicap.Text = dr("CMT_NoiCap").ToString().Trim()

                Dim ojb As Object = dr("AnhThe")

                If Not ojb Is Nothing Or Len(ojb.ToString()) > 0 Then
                    My_LoadImage(picbx_main, dr("AnhThe"))

                End If

                If (dr("IdNS_Tinh").ToString() <> "" And dr("IdNS_Tinh").ToString() <> "0") Then
                    lbl_ns_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Id={0} Order By TrangThai", CType(dr("IdNS_Tinh").ToString(), Int32)))
                End If
                If (dr("IdNS_Huyen").ToString() <> "" And dr("IdNS_Huyen").ToString() <> "0") Then
                    'lbl_ns_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdNS_Huyen").ToString(), Integer)))
                End If
                If (dr("IdNS_Xa").ToString() <> "" And dr("IdNS_Xa").ToString() <> "0") Then
                    lbl_ns_xa.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon='00' And Id = {0} Order By TrangThai", CType(dr("IdNS_Xa").ToString(), Int32)))
                End If
                If (dr("IdNS_Thon").ToString() <> "" And dr("IdNS_Thon").ToString() <> "0") Then
                    lbl_ns_thon.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon<>'00' And Id = {0} Order By TrangThai", CType(dr("IdNS_Thon").ToString(), Int32)))
                End If
                lbl_ns_diachi.Text = dr("NS_DiaChi").ToString().Trim()


                If (dr("IdNQ_Tinh").ToString() <> "" And dr("IdNQ_Tinh").ToString() <> "0") Then
                    lbl_nq_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Id={0} Order By TrangThai", CType(dr("IdNQ_Tinh").ToString(), Int32)))
                End If
                If (dr("IdNQ_Huyen").ToString() <> "" And dr("IdNQ_Huyen").ToString() <> "0") Then
                    'lbl_nq_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdNQ_Huyen").ToString(), Integer)))
                End If
                If (dr("IdNQ_Xa").ToString() <> "" And dr("IdNQ_Xa").ToString() <> "0") Then
                    lbl_nq_xa.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon='00' And Id = {0} Order By TrangThai", CType(dr("IdNQ_Xa").ToString(), Int32)))
                End If
                If (dr("IdNQ_Thon").ToString() <> "" And dr("IdNQ_Thon").ToString() <> "0") Then
                    lbl_nq_thon.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon<>'00' And Id = {0} Order By TrangThai", CType(dr("IdNQ_Thon").ToString(), Int32)))
                End If
                lbl_nq_diachi.Text = dr("NQ_DiaChi").ToString().Trim()


                If (dr("IdThT_Tinh").ToString() <> "" And dr("IdThT_Tinh").ToString() <> "0") Then
                    lbl_tt_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Id={0} Order By TrangThai", CType(dr("IdThT_Tinh").ToString(), Int32)))
                End If
                If (dr("IdThT_Huyen").ToString() <> "" And dr("IdThT_Huyen").ToString() <> "0") Then
                    'lbl_tt_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdThT_Huyen").ToString(), Integer)))
                End If
                If (dr("IdThT_Xa").ToString() <> "" And dr("IdThT_Xa").ToString() <> "0") Then
                    lbl_tt_xa.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon='00' And Id = {0} Order By TrangThai", CType(dr("IdThT_Xa").ToString(), Int32)))
                End If
                If (dr("IdThT_Thon").ToString() <> "" And dr("IdThT_Thon").ToString() <> "0") Then
                    lbl_tt_thon.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon<>'00' And Id = {0} Order By TrangThai", CType(dr("IdThT_Thon").ToString(), Int32)))
                End If
                lbl_tt_diachi.Text = dr("ThT_Diachi").ToString().Trim()
                lbl_tt_dienthoai.Text = dr("ThT_Dienthoai").ToString().Trim()

                If (dr("IdTTr_Tinh").ToString() <> "" And dr("IdTTr_Tinh").ToString() <> "0") Then
                    lbl_ttr_tinh.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Id={0} Order By TrangThai", CType(dr("IdTTr_Tinh").ToString(), Int32)))
                End If
                If (dr("IdTTr_Huyen").ToString() <> "" And dr("IdTTr_Huyen").ToString() <> "0") Then
                    'lbl_ttr_huyen.Text = _HS_CanBo.GetNameByCode(String.Format("Select id, ten_goi from DiaDanh Where id = {0} and id_goc != 0 and Status = 1", CType(dr("IdTTr_Huyen").ToString(), Integer)))
                End If
                If (dr("IdTTr_Xa").ToString() <> "" And dr("IdTTr_Xa").ToString() <> "0") Then
                    lbl_ttr_xa.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon='00' And Id = {0} Order By TrangThai", CType(dr("IdTTr_Xa").ToString(), Int32)))
                End If
                If (dr("IdTTr_Thon").ToString() <> "" And dr("IdTTr_Thon").ToString() <> "0") Then
                    lbl_ttr_thon.Text = _HS_CanBo.GetNameByCode(String.Format("Select Top 1 Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon<>'00' And Id = {0} Order By TrangThai", CType(dr("IdTTr_Thon").ToString(), Int32)))
                End If
                lbl_ttr_diachi.Text = dr("TTr_Diachi").ToString().Trim()
                lbl_ttr_dienthoai.Text = dr("TTr_Dienthoai").ToString().Trim()
                lbl_dt_didong.Text = dr("DienThoai_DD").ToString().Trim()
                lbl_dt_coquan.Text = dr("DienThoai_CQ").ToString().Trim()
                lbl_dt_nharieng.Text = dr("DienThoai_NR").ToString().Trim()
                lbl_sofax.Text = dr("SoFax").ToString().Trim()
                lbl_email.Text = dr("Email").ToString().Trim()
                lbl_nhommau.Text = dr("NhomMau").ToString().Trim()
                lbl_tp_giadinh.Text = dr("ThanhPhan_GD_HT").ToString() 'IIf(dr("IdThanhPhan_GD").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 13 and Status = 1", CType(dr("IdThanhPhan_GD").ToString(), Integer))), "")
                lbl_td_vanhoa.Text = dr("TrinhDoVH_HT").ToString()              'IIf(dr("IdTrinhDoVH").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 6 and Status = 1", CType(dr("IdTrinhDoVH").ToString(), Integer))), "")
                lbl_td_chinhtri.Text = dr("TrinhDoCT_HT").ToString()            'IIf(dr("IdTrinhDoCT").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 36 and Status = 1", CType(dr("IdTrinhDoCT").ToString(), Integer))), "")
                lbl_hocham.Text = dr("HocHam_HT").ToString()            'IIf(dr("IdHocHam").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 26 and Status = 1", CType(dr("IdHocHam").ToString(), Integer))), "")
                lbl_hocvi.Text = dr("HocVi_HT").ToString()            ' IIf(dr("IdHocVi").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 20 and Status = 1", CType(dr("IdHocVi").ToString(), Integer))), "")
                'If dr("IdTrdChMon").ToString().Trim() <> "" Then
                '    lbl_trd_chmon.Text = IIf(dr("IdTrdChMon").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 38 and Status = 1", CType(dr("IdTrdChMon").ToString(), Integer))), "")
                'End If
                lbl_trd_chmon.Text = dr("TrdChMon_HT").ToString()

                'If dr("IdTrdNgoaiNgu").ToString().Trim() <> "" Then
                '    lbl_trd_ngngu.Text = IIf(dr("IdTrdNgoaiNgu").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 38 and Status = 1", CType(dr("IdTrdNgoaiNgu").ToString(), Integer))), "")
                'End If
                lbl_trd_ngngu.Text = dr("TrdNgoaiNgu_HT").ToString()

                'If dr("IdTrdTinHoc").ToString().Trim() <> "" Then
                '    lbl_trd_tinhoc.Text = IIf(dr("IdTrdTinHoc").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 38 and Status = 1", CType(dr("IdTrdTinHoc").ToString(), Integer))), "")
                'End If
                lbl_trd_tinhoc.Text = dr("TrdTinHoc_HT").ToString()
                lbl_ut_giadinh.Text = dr("UT_GD_HT").ToString() 'IIf(dr("IdUT_GD").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 19 and Status = 1", CType(dr("IdUT_GD").ToString(), Integer))), "")
                If (dr("IdUT_BT").ToString() <> "") Then
                    Dim strVal As String = ""
                    Dim arrId As ArrayList = New ArrayList()
                    arrId = Globals.Splip_Strings(dr("IdUT_BT").ToString().Trim())
                    For i As Byte = 0 To arrId.Count - 1
                        If (i > 1 And i Mod 2 = 0) Then
                            strVal += vbCrLf
                        End If
                        strVal += "- " + IIf(arrId(i).ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 18 and Status = 1", CType(arrId(i).ToString(), Integer))), "") + ";  "
                    Next
                    strVal = strVal.Trim()
                    If (strVal <> "") Then
                        If (strVal.Substring(0, 1) = "-") Then
                            strVal = strVal.Substring(1).Trim()
                        End If
                        If (strVal.Substring(0, 1) = ";") Then
                            strVal = strVal.Substring(1).Trim()
                        End If
                    End If
                    While strVal.EndsWith(";") 'Bỏ dấu chấm phẩy ở cuối chuỗi đi
                        strVal = strVal.Substring(0, strVal.Length - 1)
                    End While
                    If (strVal <> "") Then strVal = "- " + strVal
                    lbl_ut_banthan.Text = strVal
                End If

                If dr("CM_Ngay").ToString() <> "" Then lbl_cm_ngat_tg.Text = CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd-MM-yyyy")
                lbl_cm_tochuc.Text = dr("CM_ToChuc").ToString().Trim()

                If dr("Dang_NgayVao").ToString() <> "" Then lbl_dv_ngayvao.Text = CType(dr("Dang_NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy")
                If dr("Dang_NgayChTh").ToString() <> "" Then lbl_dv_ngay_ct.Text = CType(dr("Dang_NgayChTh").ToString(), DateTime).ToString("dd-MM-yyyy")
                If dr("Dang_NgayRa").ToString() <> "" Then lbl_dv_ngayra.Text = CType(dr("Dang_NgayChTh").ToString(), DateTime).ToString("dd-MM-yyyy")
                lbl_dv_noi_kn.Text = dr("Dang_NoiKN").ToString().Trim()
                lbl_dv_sothe.Text = dr("Dang_SoThe").ToString().Trim()
                lbl_dv_nguoi_gt.Text = dr("Dang_NGT").ToString().Trim()
                lbl_dv_lydo.Text = dr("Dang_LyDoRa").ToString().Trim()

                lbl_tk_makh.Text = dr("NH_MSKH").ToString().Trim()
                lbl_tk_sotk.Text = dr("NH_SHTK").ToString().Trim()
                lbl_tk_tennh.Text = dr("NH_Ten_NH").ToString().Trim()
                lbl_ms_thue.Text = dr("MaSoThue").ToString().Trim()

                lbl_bhxh_soso.Text = dr("BHXH_SoSo").ToString().Trim()

                If dr("BHXH_NgayLam").ToString() <> "" Then lbl_bhxh_ngaylam.Text = CType(dr("BHXH_NgayLam").ToString(), DateTime).ToString("dd-MM-yyyy")
                If dr("BHXH_NgayDong").ToString() <> "" Then lbl_bhxh_ngaydong.Text = CType(dr("BHXH_NgayDong").ToString(), DateTime).ToString("dd-MM-yyyy")
                lbl_bhxh_noilam.Text = dr("BHXH_NoiLam").ToString().Trim()

                lbl_chng_daotao.Text = IIf(dr("IdChuyenNganhDT").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 11 and Status = 1", CType(dr("IdChuyenNganhDT").ToString(), Integer))), "")

                'View thông tin liên quan đến hợp đồng lao động của Cán bộ Tập sự
                Dim strSQL As String = String.Format("Select A.*,B.ten_goi As ChucVu,dbo.GetValForFieldName('LOAIHINHCONGVIEC',A.Loai,2) Loai_HT From HSCB_TS_HDLD A, DanhMuc B Where IdCanbo = '{0}' and (B.id = A.IdCV_Nguoi_QD and B.id_goc = 14 and B.Status = 1) Order by TuNgay DESC", _IdCanBo)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (db.Rows(0)("Status").ToString() <> "") Then
                                If db.Rows(0)("Status").ToString() <> "True" Then
                                    rb_hdld.Checked = True
                                    rb_hdld_CheckedChanged(Nothing, Nothing)
                                Else
                                    rb_qdtamtuyen.Checked = True
                                    rb_qdtamtuyen_CheckedChanged(Nothing, Nothing)
                                End If
                            End If

                            lbl_hdld_sohd.Text = db.Rows(0)("SoQD").ToString().Trim()
                            If db.Rows(0)("NgayQD").ToString() <> "" Then lbl_hdld_ngayky.Text = CType(db.Rows(0)("NgayQD").ToString(), DateTime).ToString("dd-MM-yyyy")
                            lbl_hdld_nguoiky.Text = db.Rows(0)("NguoiQD").ToString().Trim()
                            lbl_hdld_chucvu.Text = db.Rows(0)("ChucVu").ToString().Trim()

                            If db.Rows(0)("TuNgay").ToString() <> "" Then lbl_tg_tungay.Text = CType(db.Rows(0)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                            If db.Rows(0)("DenNgay").ToString() <> "" Then lbl_tg_denngay.Text = CType(db.Rows(0)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")

                            lbl_gio_batdau.Text = db.Rows(0)("TuGio").ToString().Trim()
                            lbl_gio_ketthuc.Text = db.Rows(0)("DenGio").ToString().Trim()
                            lbl_loaihinh_cv.Text = db.Rows(0)("Loai_HT").ToString().Trim() 'clsHS_CanBo.GetLoaiHinh(CType(db.Rows(0)("Loai").ToString().Trim(), Byte))
                            lbl_ht_traluong.Text = IIf(db.Rows(0)("IdHT_TraLuong").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select id,ten_goi from DanhMuc Where id = {0} And id_goc = 40 and Status = 1", CType(db.Rows(0)("IdHT_TraLuong").ToString(), Integer))), "")

                            Dim _BangLuongId As Integer = 0
                            Dim _NgachLuongId As Integer = 0
                            Dim _BacLuongId As Integer = 0
                            _BacLuongId = IIf(db.Rows(0)("IdBacLuong").ToString() <> "0", CType(db.Rows(0)("IdBacLuong").ToString(), Integer), 0) ' Id bậc lương
                            If (_BacLuongId > 0) Then
                                'Thực hiện hiển thị dữ liệu và cho hiện các controls
                                lbl_ht_huongluong.Text = "Thang - bảng lương"
                                pnl_cdld_4.Visible = True
                                lbl_cdld_div_2.Visible = True
                                pnl_cdld_3.Visible = True
                                lbl_cdld_div_1.Visible = True
                                lbl_cdld_div_1.BringToFront()
                                pnl_cdld_3.BringToFront()
                                lbl_cdld_div_2.BringToFront()
                                pnl_cdld_4.BringToFront()
                                lbl_bangluong_div.Visible = True
                                lbl_cdld_bangluong_code.Visible = True
                                lbl_cdld_bangluong_code.BringToFront()
                                lbl_title_bangluong.Text = "Bảng lương  "
                                pnl_hdld_8.Height = 95
                                lbl_cdld_bangluong.Width = 450
                                _NgachLuongId = _Labour.GetIdNgachLuong(_BacLuongId)
                                If (_NgachLuongId > 0) Then
                                    _BangLuongId = _Labour.GetIdBangLuong(_NgachLuongId)
                                End If
                                'Lấy thông tin mô tả của Bảng lương, Ngạch lương, Bậc lương
                                If (_BangLuongId > 0) Then
                                    strSQL = String.Format("Select * From BangLuong Where IdBangLuong = {0}", _BangLuongId)
                                    Using db_bluong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                        If Not (db_bluong Is Nothing) Then
                                            If (db_bluong.Rows.Count > 0) Then
                                                lbl_cdld_bangluong.Text = db_bluong.Rows(0)("MoTa").ToString().Trim()
                                                lbl_cdld_bangluong_code.Text = db_bluong.Rows(0)("BangLuong").ToString().Trim()
                                            End If
                                        End If
                                    End Using
                                Else
                                    lbl_cdld_bangluong.Text = ""
                                    lbl_cdld_bangluong_code.Text = ""
                                End If

                                If (_NgachLuongId > 0) Then
                                    strSQL = String.Format("Select * From NgachLuong Where IdNgachLuong = {0}", _NgachLuongId)
                                    Using db_ngluong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                        If Not (db_ngluong Is Nothing) Then
                                            If (db_ngluong.Rows.Count > 0) Then
                                                lbl_cdld_ngachluong.Text = db_ngluong.Rows(0)("MoTa").ToString().Trim()
                                                lbl_cdld_ngachluong_code.Text = db_ngluong.Rows(0)("NgachLuong").ToString().Trim()
                                            End If
                                        End If
                                    End Using
                                Else
                                    lbl_cdld_ngachluong.Text = ""
                                    lbl_cdld_ngachluong_code.Text = ""
                                End If

                                If (_BacLuongId > 0) Then
                                    strSQL = String.Format("Select * From BacLuong Where IdBacLuong = {0}", _BacLuongId)
                                    Using db_bluong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                        If Not (db_bluong Is Nothing) Then
                                            If (db_bluong.Rows.Count > 0) Then
                                                lbl_cdld_bacluong.Text = db_bluong.Rows(0)("MoTa").ToString().Trim()
                                            End If
                                        End If
                                    End Using
                                Else
                                    lbl_cdld_bacluong.Text = ""
                                End If
                                lbl_cdld_hsluong.Text = IIf(db.Rows(0)("Heso").ToString().Trim() <> "", CType(db.Rows(0)("Heso"), Double).ToString("N2"), "0")
                                lbl_cdld_tlhuong.Text = IIf(db.Rows(0)("TyleHuong").ToString().Trim() <> "", CType(db.Rows(0)("TyleHuong"), Double).ToString("N2"), "0")
                            Else
                                lbl_ht_huongluong.Text = "Lương trọn gói"
                                lbl_cdld_bangluong.Text = ""
                                lbl_cdld_ngachluong.Text = ""
                                lbl_cdld_bacluong.Text = ""
                                lbl_cdld_hsluong.Text = ""
                                lbl_cdld_tlhuong.Text = ""

                                lbl_bangluong_div.Visible = False
                                lbl_cdld_bangluong_code.Visible = False
                                lbl_cdld_div_1.Visible = False
                                pnl_cdld_3.Visible = False
                                lbl_cdld_div_2.Visible = False
                                pnl_cdld_4.Visible = False

                                lbl_title_bangluong.Text = "Tiền lương  "
                                pnl_hdld_8.Height = 46
                                lbl_cdld_bangluong.Width = 283
                                lbl_cdld_bangluong.Text = IIf(db.Rows(0)("Tien_Luong").ToString().Trim() <> "", Double.Parse(db.Rows(0)("Tien_Luong").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0") + "  VNĐ"
                            End If
                            If rb_hdld.Checked = True Then
                                ckb_choice_bhxh.Checked = IIf(db.Rows(0)("BHXH").ToString() = True, True, False)
                                ckb_choice_bhyt.Checked = IIf(db.Rows(0)("BHYT").ToString() = True, True, False)
                            End If
                            lbl_cv_damnhan.Text = db.Rows(0)("CongViec").ToString().Trim()
                        End If
                    End If
                End Using
                'Load thông tin về Lương -> Từ bậc lương
                lbl_st_congtac.Text = dr("SoTruong_CT").ToString().Trim()
                lbl_cv_lamlau.Text = dr("CV_Lau").ToString().Trim()
                lbl_ghichu.Text = dr("GhiChu").ToString().Trim()
                lbl_honnhan.Text = dr("HonNhan_HT").ToString().Trim()
            End If
        End If
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
            cb_records.SelectedIndex = Convert.ToInt32(_RecordCurrent)
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmTS_ChiTiet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Fill danh sách items ưu tiên bản thân
        Fill_Record()
    End Sub

    Private Sub cb_records_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_records.SelectedIndexChanged
        If (cb_records.Items.Count <> 0) Then
            _IdCanBo = arr_RecordId(cb_records.SelectedIndex).ToString()
            Fill_Data(_IdCanBo)
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

    Private Sub frmTS_ChiTiet_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Close()
        End If
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub frmTS_ChiTiet_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
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

    Private Sub rb_hdld_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_hdld.CheckedChanged
        If rb_hdld.Checked = True Then
            pnl_hdld_11.Visible = True
            pnl_hdld_12.Visible = True
            lbl_hdld_sohd_qdtt.Text = "Số hợp đồng  "
        End If
    End Sub

    Private Sub rb_qdtamtuyen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb_qdtamtuyen.CheckedChanged
        If rb_qdtamtuyen.Checked = True Then
            pnl_hdld_11.Visible = False
            pnl_hdld_12.Visible = False
            lbl_hdld_sohd_qdtt.Text = "Số quyết định  "
        End If
    End Sub
#End Region

End Class