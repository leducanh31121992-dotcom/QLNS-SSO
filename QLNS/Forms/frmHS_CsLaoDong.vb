Public Class frmHS_CsLaoDong

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - The profile of the accident <---"
    Private _HS_CsLaodong As clsHS_CsLaodong = New clsHS_CsLaodong()
    Private _PersonnelFile As clsHS_CanBo = New clsHS_CanBo()
    Private _ListDocument As clsHT_DanhMuc = New clsHT_DanhMuc()
    Private _Globals As Globals = New Globals()
    Private _SqlHelper As DBAccess = New DBAccess()
    'Khai báo mảng lưu danh sách các bản ghi - của dữ liệu fill ra combobox
    Private arr_Tengoi As ArrayList = New ArrayList 'Tên gọi vụ tai nạn
    Private arr_LoaiNP As ArrayList = New ArrayList

    'Khai báo các biến lưu id của các hồ sơ
    Private strCodeId As String = ""
    'Biến lưu id cán bộ khi select trên cây dữ liệu
    Private _IdCanBo As String = ""
    Private strNode As String = ""
    Dim obj_hs_vutainan As frmHS_VuTaiNan
    Private vNFInfo As System.Globalization.NumberFormatInfo
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

    'Khai báo biến sử dụng khi được gọi từ Hồ sơ cán bộ
    Public HumanId As String = ""       'Giá trị id của cán bộ
    Public FlagShow As Boolean = False  'Giá trị True là được gọi từ Hồ sơ cán bộ. False: Giá trị mặc định
    Public TabVal As Byte = 0           'Giá trị xác định kem gọi Tab nào ra thực hiện
    Private _IdDonviHT As Integer = IdDONVI 'Biến dùng để xem chi tiết Hồ sơ cán bộ
    Public TagNode As String = ""       'Giá trị lưu lại tag đơn vị đang select bên hồ sơ cán bộ
#End Region

#Region "---> Functions: Các hàm dùng chung <---"
    ''' <summary>
    ''' Hàm thực hiện Reset các controls về trạng thái ban đầu
    ''' </summary>
    ''' <param name="state">Chỉ số xác định Tab đang thao tác
    '''        0 - Toàn bộ controls
    '''        1 - Hồ sơ BHYT  
    '''        2 - Hồ sơ khám sức khoẻ
    '''        3 - HS nghỉ phép
    '''        4 - Hồ sơ cán bộ tai nạn
    ''' </param>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls(ByVal state As Byte)
        Select Case state
            Case 0  'Reset - Toàn bộ controls
                lbl_macb.Text = ""
                lbl_hoten.Text = ""
                lbl_ngaysinh.Text = ""
                lbl_gioitinh.Text = ""
                lbl_so_cmt.Text = ""

                strCodeId = ""
                _IdCanBo = ""
                'Tab2:  Reset - Hồ sơ bảo hiểm Y tế
                dgv_bhyt.Rows.Clear()
                'Thiết lập Reset các control - trên tab - Bảo hiểm y tế
                edt_bhyt_sothe.Text = ""
                dtpk_bhyt_ngaycap.Text = DateTime.Now.ToShortDateString()
                edt_bhyt_noicap.Text = ""
                edt_bhyt_noi_dkkb.Text = ""
                dtpk_bhyt_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_bhyt_denngay.Text = DateTime.Now.ToShortDateString()
                edt_bhyt_ghichu.Text = ""

                'Tab3:  Reset - Hồ sơ Khám sức khoẻ của cán bộ
                dgv_hoso_khamsk.Rows.Clear()
                'Thiết lập Reset các control - trên tab - Hồ sơ khám sức khoẻ
                dtpk_suckhoe_nam.Text = DateTime.Now.ToShortDateString()
                edt_suckhoe_dotkham.Value = 0
                edt_suckhoe_noidung.Text = ""
                edt_suckhoe_noikham.Text = ""
                dtpk_suckhoe_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_suckhoe_denngay.Text = DateTime.Now.ToShortDateString()
                edt_suckhoe_cannang.Text = "0"
                edt_suckhoe_chieucao.Text = "0"
                cb_suckhoe_loaisk.SelectedIndex = 0
                edt_suckhoe_ketluan.Text = ""
                edt_suckhoe_ghichu.Text = ""

                'Tab4:  Reset - Hồ sơ Nghỉ phép của cán bộ
                dgv_hsnghiphep.Rows.Clear()
                'Thiết lập Reset các control - trên tab - Hồ sơ Nghỉ phép của cán bộ
                dtpk_hsnghiphep_namnghi.Text = DateTime.Now.ToShortDateString()
                dtpk_hsnghiphep_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_hsnghiphep_denngay.Text = DateTime.Now.ToShortDateString()
                edt_hsnghiphep_songay.Text = Globals.GetCountWorkDays(dtpk_hsnghiphep_tungay.Value, dtpk_hsnghiphep_denngay.Value).ToString().Trim()
                edt_hsnghiphep_tyle.Text = "0"
                cb_hsnghiphep_loai_ngphep.SelectedIndex = 0
                cb_hsnghiphep_loai_ngphep_SelectedIndexChanged(Nothing, Nothing)
                edt_hsnghiphep_noinghi.Text = ""
                edt_hsnghiphep_lydo.Text = ""
                edt_hsnghiphep_tientrcap.Text = "0"
                edt_hsnghiphep_nguoiky.Text = ""
                edt_hsnghiphep_ghichu.Text = ""

                'Tab5:  Reset - Hồ sơ Cán bộ Tai nạn
                dgv_cbtainan.Rows.Clear()
                'Thiết lập Reset các control - trên tab - Hồ sơ cán bộ tai nạn
                cb_cbtainan_tengoi.SelectedIndex = 0
                cb_cbtainan_mucdo.SelectedIndex = 0
                edt_cbtainan_songay.Text = "0"
                edt_cbtainan_sotien_th.Text = "0"
                edt_cbtainan_tt_thuongtat.Text = ""
                edt_cbtainan_noi_ppdt.Text = ""
                edt_cbtainan_ghichu.Text = ""
            Case 1  'Reset - Hồ sơ bảo hiểm Y tế
                strCodeId = ""
                edt_bhyt_sothe.Text = ""
                dtpk_bhyt_ngaycap.Text = DateTime.Now.ToShortDateString()
                edt_bhyt_noicap.Text = ""
                edt_bhyt_noi_dkkb.Text = ""
                dtpk_bhyt_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_bhyt_denngay.Text = DateTime.Now.ToShortDateString()
                edt_bhyt_ghichu.Text = ""
            Case 2  'Reset - Hồ sơ Khám sức khoẻ của cán bộ
                strCodeId = ""
                dtpk_suckhoe_nam.Text = DateTime.Now.ToShortDateString()
                edt_suckhoe_dotkham.Value = 0
                edt_suckhoe_noidung.Text = ""
                edt_suckhoe_noikham.Text = ""
                dtpk_suckhoe_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_suckhoe_denngay.Text = DateTime.Now.ToShortDateString()
                edt_suckhoe_cannang.Text = "0"
                edt_suckhoe_chieucao.Text = "0"
                cb_suckhoe_loaisk.SelectedIndex = 0
                edt_suckhoe_ketluan.Text = ""
                edt_suckhoe_ghichu.Text = ""
            Case 3  'Reset - Hồ sơ Nghỉ phép của cán bộ
                strCodeId = ""
                dtpk_hsnghiphep_namnghi.Text = DateTime.Now.ToShortDateString()
                dtpk_hsnghiphep_tungay.Text = DateTime.Now.ToShortDateString()
                dtpk_hsnghiphep_denngay.Text = DateTime.Now.ToShortDateString()
                edt_hsnghiphep_songay.Text = Globals.GetCountWorkDays(dtpk_hsnghiphep_tungay.Value, dtpk_hsnghiphep_denngay.Value).ToString().Trim()
                cb_hsnghiphep_loai_ngphep.SelectedIndex = 0
                cb_hsnghiphep_loai_ngphep_SelectedIndexChanged(Nothing, Nothing)
                edt_hsnghiphep_noinghi.Text = ""
                edt_hsnghiphep_lydo.Text = ""
                edt_hsnghiphep_tientrcap.Text = "0"
                edt_hsnghiphep_nguoiky.Text = ""
                edt_hsnghiphep_ghichu.Text = ""
            Case 4  'Reset - Hồ sơ Cán bộ Tai nạn
                strCodeId = ""
                cb_cbtainan_tengoi.SelectedIndex = 0
                cb_cbtainan_mucdo.SelectedIndex = 0
                edt_cbtainan_songay.Text = "0"
                edt_cbtainan_sotien_th.Text = "0"
                edt_cbtainan_tt_thuongtat.Text = ""
                edt_cbtainan_noi_ppdt.Text = ""
                edt_cbtainan_ghichu.Text = ""
            Case Else
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu - Các hồ sơ theo Id cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Chuỗi id xác định cán bộ</param>
    ''' <remarks></remarks>
    Private Sub FillAll_Documents(ByVal _IdCanBo As String)
        'Thực hiện fill dữ liệu hồ sơ - Bảo hiểm Y tế của Cán bộ
        Using db As DataTable = _HS_CsLaodong.GetAll(_IdCanBo, 2)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_bhyt.Rows.Add()
                        dgv_bhyt.Rows(i).Cells("cln_Code").Value = db.Rows(i)("IdBHYT").ToString().Trim()
                        dgv_bhyt.Rows(i).Cells("cln_SoHieu").Value = db.Rows(i)("SoThe").ToString().Trim()
                        If (db.Rows(i)("NgayCap").ToString().Trim() <> "") Then
                            dgv_bhyt.Rows(i).Cells("cln_Ngaycap").Value = Format(CType(db.Rows(i)("NgayCap"), DateTime), "dd-MM-yyyy")
                        Else
                            dgv_bhyt.Rows(i).Cells("cln_Ngaycap").Value = ""
                        End If
                        dgv_bhyt.Rows(i).Cells("cln_Noicap").Value = db.Rows(i)("NoiCap").ToString().Trim()
                        dgv_bhyt.Rows(i).Cells("cln_Noi_Dk").Value = db.Rows(i)("NoiDangKy_KCB").ToString().Trim()
                        If (db.Rows(i)("TuNgay").ToString().Trim() <> "") Then
                            dgv_bhyt.Rows(i).Cells("cln_Tungay").Value = Format(CType(db.Rows(i)("TuNgay"), DateTime), "dd-MM-yyyy")
                        Else
                            dgv_bhyt.Rows(i).Cells("cln_Tungay").Value = ""
                        End If

                        If (db.Rows(i)("DenNgay").ToString().Trim() <> "") Then
                            dgv_bhyt.Rows(i).Cells("cln_Denngay").Value = Format(CType(db.Rows(i)("DenNgay"), DateTime), "dd-MM-yyyy")
                        Else
                            dgv_bhyt.Rows(i).Cells("cln_Denngay").Value = ""
                        End If
                        dgv_bhyt.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                    Next
                End If
            End If
        End Using

        'Thực hiện fill dữ liệu hồ sơ - Hồ sơ khám sức khoẻ của Cán bộ
        Using db As DataTable = _HS_CsLaodong.GetAll(_IdCanBo, 3)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_hoso_khamsk.Rows.Add()
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Code").Value = db.Rows(i)("IdKhamChua").ToString().Trim()
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Namkham").Value = db.Rows(i)("Nam").ToString().Trim()
                        dgv_hoso_khamsk.Rows(i).Cells("cln_DotKham").Value = db.Rows(i)("Dot").ToString().Trim()
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Noidung").Value = db.Rows(i)("NoiDung").ToString().Trim()
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Noikham").Value = db.Rows(i)("NoiKham").ToString().Trim()

                        If (db.Rows(i)("TuNgay").ToString().Trim() <> "") Then
                            dgv_hoso_khamsk.Rows(i).Cells("cln_Tungay").Value = Format(CType(db.Rows(i)("TuNgay"), DateTime), "dd-MM-yyyy")
                        Else
                            dgv_hoso_khamsk.Rows(i).Cells("cln_Tungay").Value = ""
                        End If

                        If (db.Rows(i)("DenNgay").ToString().Trim() <> "") Then
                            dgv_hoso_khamsk.Rows(i).Cells("cln_Denngay").Value = Format(CType(db.Rows(i)("DenNgay"), DateTime), "dd-MM-yyyy")
                        Else
                            dgv_hoso_khamsk.Rows(i).Cells("cln_Denngay").Value = ""
                        End If
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Cannang").Value = IIf(db.Rows(i)("CanNang").ToString().Trim() <> "", CType(db.Rows(i)("CanNang"), Double).ToString("N2"), "0")
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Chieucao").Value = IIf(db.Rows(i)("ChieuCao").ToString().Trim() <> "", CType(db.Rows(i)("ChieuCao"), Double).ToString("N2"), "0")

                        If (db.Rows(i)("LoaiSucKhoe").ToString().Trim() = "1") Then
                            dgv_hoso_khamsk.Rows(i).Cells("cln_LoaiSk").Value = "Loại A"
                        ElseIf (db.Rows(i)("LoaiSucKhoe").ToString().Trim() = "2") Then
                            dgv_hoso_khamsk.Rows(i).Cells("cln_LoaiSk").Value = "Loại B"
                        ElseIf (db.Rows(i)("LoaiSucKhoe").ToString().Trim() = "3") Then
                            dgv_hoso_khamsk.Rows(i).Cells("cln_LoaiSk").Value = "Loại C"
                        Else
                            dgv_hoso_khamsk.Rows(i).Cells("cln_LoaiSk").Value = ""
                        End If

                        dgv_hoso_khamsk.Rows(i).Cells("cln_Ketluan").Value = db.Rows(i)("KetLuan").ToString().Trim()
                        dgv_hoso_khamsk.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                    Next
                End If
            End If
        End Using

        'Thực hiện fill dữ liệu hồ sơ - Hồ sơ Nghỉ phép của Cán bộ
        Using db As DataTable = _HS_CsLaodong.HS_NgPh_GetSearch("", "", _IdCanBo, "", 0, 0, "", "", 0)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_hsnghiphep.Rows.Add()
                        dgv_hsnghiphep.Rows(i).Cells("cln_Code").Value = db.Rows(i)("IdNghiPhep").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_NamNghi").Value = db.Rows(i)("Nam").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_Tungay").Value = db.Rows(i)("TuNgay_HT").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_Denngay").Value = db.Rows(i)("DenNgay_HT").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_SoNgay").Value = db.Rows(i)("SoNgay").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_LoaiNgPhep").Value = db.Rows(i)("LoaiNghi_HT").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_Noi_NgPhep").Value = db.Rows(i)("NoiNghi").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_Lydonghi").Value = db.Rows(i)("LyDo").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_Nguoiky").Value = db.Rows(i)("NguoiKy_Nghi").ToString().Trim()
                        dgv_hsnghiphep.Rows(i).Cells("cln_TyLe").Value = IIf(db.Rows(i)("TyleHuong").ToString() <> "", CType(db.Rows(i)("TyleHuong"), Double).ToString("N2"), "0")
                        dgv_hsnghiphep.Rows(i).Cells("cln_TienTC").Value = IIf(db.Rows(i)("TienTroCap").ToString() <> "", Double.Parse(db.Rows(i)("TienTroCap").ToString(), Globals.cultureNum).ToString("N", vNFInfo), "0")
                        'dgv_main.Rows(i).Cells("cln_LCB_HeSo").Value = Double.Parse(db_luong.Rows(i)("LCB_HeSo").ToString(), Globals.cultureNum).ToString("N2", vNFInfo)
                        dgv_hsnghiphep.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                    Next
                End If
            End If
        End Using

        'Thực hiện fill dữ liệu hồ sơ - Hồ sơ Cán bộ Tai nạn
        Using db As DataTable = _HS_CsLaodong.GetAll(_IdCanBo, 5)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_cbtainan.Rows.Add()
                        dgv_cbtainan.Rows(i).Cells("cln_Code").Value = db.Rows(i)("IdCBTaiNan").ToString().Trim()
                        dgv_cbtainan.Rows(i).Cells("cln_Tengoi").Value = db.Rows(i)("Ten_Goi").ToString().Trim()
                        If (db.Rows(i)("MucDo_TN").ToString() <> "") Then
                            If (CType(db.Rows(i)("MucDo_TN").ToString(), Byte) > 0) Then
                                If (CType(db.Rows(i)("MucDo_TN").ToString(), Byte) = 1) Then
                                    dgv_cbtainan.Rows(i).Cells("cln_Mucdo").Value = "Nhẹ"
                                ElseIf (CType(db.Rows(i)("MucDo_TN").ToString(), Byte) = 2) Then
                                    dgv_cbtainan.Rows(i).Cells("cln_Mucdo").Value = "Nặng"
                                ElseIf (CType(db.Rows(i)("MucDo_TN").ToString(), Byte) = 3) Then
                                    dgv_cbtainan.Rows(i).Cells("cln_Mucdo").Value = "Chết"
                                End If
                            Else
                                dgv_cbtainan.Rows(i).Cells("cln_Mucdo").Value = ""
                            End If
                        Else
                            dgv_cbtainan.Rows(i).Cells("cln_Mucdo").Value = ""
                        End If
                        dgv_cbtainan.Rows(i).Cells("cln_Songay").Value = db.Rows(i)("SoNgayNghi").ToString().Trim()
                        dgv_cbtainan.Rows(i).Cells("cln_Chiphi").Value = IIf(db.Rows(i)("ChiPhi").ToString() <> "", Double.Parse(db.Rows(i)("ChiPhi").ToString(), Globals.cultureNum).ToString("N", vNFInfo), "0")
                        dgv_cbtainan.Rows(i).Cells("cln_TTThuongtat").Value = db.Rows(i)("TT_ThuongTat").ToString().Trim()
                        dgv_cbtainan.Rows(i).Cells("cln_Noi_PPdieutri").Value = db.Rows(i)("Noi_PP_DieuTri").ToString().Trim()
                        dgv_cbtainan.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                    Next
                End If
            End If
        End Using

    End Sub

    ''' <summary>
    ''' Thực hiện reset lại sau khi gọi Hồ sơ các vụ tai nạn sau
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad_Main()
        obj_hs_vutainan.Dispose()
        cb_cbtainan_tengoi.Items.Clear()
        Dim strSQL As String = ""
        strSQL = String.Format("Select IdVuTaiNan, TenVu_TN From HS_VuTaiNan Order by TenVu_TN")
        arr_Tengoi.Clear()
        cb_cbtainan_tengoi.Items.Clear()
        arr_Tengoi = _Globals.Bind_ComBoBox(cb_cbtainan_tengoi, strSQL, "---Tên vụ tai nạn---")
        cb_cbtainan_tengoi.SelectedIndex = CType(obj_hs_vutainan.Index, Integer)
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra dữ liệu cập nhật (thêm - sửa).
    ''' </summary>
    ''' <returns>True: Hợp lệ. fasle - Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Dim strSQL As String = ""
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"      'Hồ sơ bảo hiểm y tế
                If (edt_bhyt_sothe.Text.Trim() = "") Then
                    MessageBox.Show("Số thể bảo hiểm y tế không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_bhyt_sothe
                    Return False
                End If
                If (edt_bhyt_noicap.Text.Trim() = "") Then
                    MessageBox.Show("Nơi cấp thẻ bảo hiểm y tế không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_bhyt_noicap
                    Return False
                End If

                If (edt_bhyt_noi_dkkb.Text.Trim() = "") Then
                    MessageBox.Show("Nơi đăng ký khám chữa bệnh không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_bhyt_noi_dkkb
                    Return False
                End If
                If dtpk_bhyt_tungay.Value >= dtpk_bhyt_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_bhyt_denngay
                    Return False
                End If

                'Bắt điều kiện nhập dữ liệu khoảng thời gian không cho đan xen lẫn nhau
                If (strCodeId = "") Then
                    strSQL = String.Format("Select * From HS_BHYT Where IdCanBo = '{0}' Order by TuNgay Asc", _IdCanBo)
                Else
                    strSQL = String.Format("Select * From HS_BHYT Where IdCanBo = '{0}' And IdBHYT <> '{1}' Order by TuNgay Asc", _IdCanBo, strCodeId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_bhyt_tungay.Value And dtpk_bhyt_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_bhyt_denngay.Value And dtpk_bhyt_denngay.Value <= _DenNgay) Or (dtpk_bhyt_tungay.Value <= _TuNgay And dtpk_bhyt_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ bảo hiểm y tế không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_bhyt_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using
            Case "tp_suckhoe"      'Hồ sơ Khám sức khoẻ của cán bộ
                If (edt_suckhoe_dotkham.Value <= 0) Then
                    MessageBox.Show("Bạn chưa chọn đợt khám sức khoẻ trong năm!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_suckhoe_dotkham
                    Return False
                End If
                If dtpk_suckhoe_tungay.Value > dtpk_suckhoe_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu không thể lớn hơn ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_suckhoe_denngay
                    Return False
                End If
                If (edt_suckhoe_ketluan.Text.Trim() = "") Then
                    MessageBox.Show("Kết luận đợt khám sức khoẻ không thể để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_suckhoe_ketluan
                    Return False
                End If
                If (cb_suckhoe_loaisk.SelectedIndex <= 0) Then
                    MessageBox.Show("Bạn chưa chọn loại sức khoẻ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_suckhoe_loaisk
                    Return False
                End If
                If (edt_suckhoe_cannang.Text.Trim() <> "") Then
                    If (CType(edt_suckhoe_cannang.Text.Trim().Replace(" ", ""), Double) > 150) Then
                        MessageBox.Show(String.Format("Giá trị cân nặng trong hồ sơ khám sức khoẻ không hợp lệ!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_suckhoe_cannang
                        Return False
                    End If
                End If

                If (edt_suckhoe_chieucao.Text.Trim() <> "") Then
                    If (CType(edt_suckhoe_chieucao.Text.Trim().Replace(" ", ""), Double) > 250) Then
                        MessageBox.Show(String.Format("Giá trị chiều cao trong hồ sơ khám sức khoẻ không hợp lệ!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_suckhoe_chieucao
                        Return False
                    End If
                End If

                If (dtpk_suckhoe_nam.Value.Year <> dtpk_suckhoe_tungay.Value.Year Or dtpk_suckhoe_nam.Value.Year <> dtpk_suckhoe_denngay.Value.Year) Then
                    MessageBox.Show("Năm khám không phù hợp với khoảng thời gian trong hồ sơ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_suckhoe_nam
                    Return False
                End If

                'Bắt trùng đợi khám trong cùng một năm - của cán bộ
                If (edt_suckhoe_dotkham.Value > 0) Then
                    If (strCodeId <> "") Then
                        strSQL = String.Format("Select * from HS_SucKhoe Where Dot = {0} And Nam = {1} And IdCanBo = '{2}' And IdKhamChua <> '{3}'", CType(edt_suckhoe_dotkham.Value, Int16), dtpk_suckhoe_nam.Value.Year, _IdCanBo, strCodeId)
                    Else
                        strSQL = String.Format("Select * from HS_SucKhoe Where Dot = {0} And Nam = {1} And IdCanBo = '{2}'", CType(edt_suckhoe_dotkham.Value, Int16), dtpk_suckhoe_nam.Value.Year, _IdCanBo)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                MessageBox.Show(String.Format("Đợt khám sức khoẻ trong năm đã tồn tại. Bạn vui lòng kiểm tra lại!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_suckhoe_dotkham
                                Return False
                            End If
                        End If
                    End Using
                End If

                'Bắt điều kiện nhập dữ liệu khoảng thời gian không cho trùng hoặc đan xen lẫn nhau
                If (strCodeId = "") Then
                    strSQL = String.Format("Select * From HS_SucKhoe Where IdCanBo = '{0}' Order by TuNgay Asc", _IdCanBo)
                Else
                    strSQL = String.Format("Select * From HS_SucKhoe Where IdCanBo = '{0}' And IdKhamChua <> '{1}' Order by TuNgay Asc", _IdCanBo, strCodeId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_suckhoe_tungay.Value And dtpk_suckhoe_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_suckhoe_denngay.Value And dtpk_suckhoe_denngay.Value <= _DenNgay) Or (dtpk_suckhoe_tungay.Value <= _TuNgay And dtpk_suckhoe_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ khám sức khoẻ không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_suckhoe_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using

            Case "tp_nghiphep"      'Hồ sơ Nghỉ phép của cán bộ
                If (dtpk_hsnghiphep_tungay.Value.Year > dtpk_hsnghiphep_namnghi.Value.Year + 1 Or dtpk_hsnghiphep_denngay.Value.Year > dtpk_hsnghiphep_namnghi.Value.Year + 1) Then
                    MessageBox.Show("Khoảng thời gian nghỉ so với năm tính phép không hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_hsnghiphep_denngay
                    Return False
                End If
                If (edt_hsnghiphep_songay.Text.Trim() = "" Or CType(edt_hsnghiphep_songay.Text.Trim().Replace(" ", ""), Int16) <= 0) Then
                    MessageBox.Show("Số ngày nghỉ phép không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_hsnghiphep_songay
                    Return False
                End If
                If (cb_hsnghiphep_loai_ngphep.SelectedIndex <= 0 And cb_hsnghiphep_loai_ngphep.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn loại nghỉ phép trong hồ sơ nghỉ phép cần cập nhật!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_hsnghiphep_loai_ngphep
                    Return False
                End If
                If dtpk_hsnghiphep_tungay.Value > dtpk_hsnghiphep_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu không thể lớn hơn ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_suckhoe_denngay
                    Return False
                End If

                If (edt_hsnghiphep_tyle.Text.Trim() = "" Or CType(edt_hsnghiphep_tyle.Text.Trim().Replace(" ", ""), Double) < 0) Then
                    MessageBox.Show("Tỷ lệ hưởng lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_hsnghiphep_tyle
                    Return False
                End If
                If (edt_hsnghiphep_nguoiky.Text.Trim() = "") Then
                    MessageBox.Show("Người ký quyết định nghỉ phép không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_hsnghiphep_nguoiky
                    Return False
                End If

                'Bắt điều kiện Không cho khoảng thời gian Trùng Hoặc Đan sen lẫn nhau

                If (strCodeId = "") Then
                    strSQL = String.Format("Select * From HS_NgPh Where IdCanBo = '{0}' Order by TuNgay Asc", _IdCanBo)
                Else
                    strSQL = String.Format("Select * From HS_NgPh Where IdCanBo = '{0}' And IdNghiPhep <> '{1}' Order by TuNgay Asc", _IdCanBo, strCodeId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_hsnghiphep_tungay.Value And dtpk_hsnghiphep_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_hsnghiphep_denngay.Value And dtpk_hsnghiphep_denngay.Value <= _DenNgay) Or (dtpk_hsnghiphep_tungay.Value <= _TuNgay And dtpk_hsnghiphep_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ phỉ phép không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian trong hồ sơ không thể trùng hoặc đan xen lẫn nhau.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_hsnghiphep_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using
            Case "tp_cbtainan"      'Hồ sơ bảo hiểm xã hội
                If (cb_cbtainan_tengoi.SelectedIndex <= 0 And cb_cbtainan_tengoi.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn tên vụ tai nạn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_cbtainan_tengoi
                    Return False
                End If
                If (cb_cbtainan_mucdo.SelectedIndex <= 0 And cb_cbtainan_mucdo.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn mức độ tai nạn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_cbtainan_mucdo
                    Return False
                End If
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện check quyền thành viên thao tác với module
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";119;120;121;122;", _roles)) Then       ' Hồ sơ - Bảo hiểm Y tế
            If tctrl_main.TabPages.Contains(tp_bhyt) Then tctrl_main.TabPages.Remove(tp_bhyt)
        End If
        If Not (Globals.IsIntersect(";123;124;125;126;", _roles)) Then       ' Hồ sơ - Khám sức khoẻ
            If tctrl_main.TabPages.Contains(tp_suckhoe) Then tctrl_main.TabPages.Remove(tp_suckhoe)
        End If
        If Not (Globals.IsIntersect(";127;128;129;130;", _roles)) Then       ' Hồ sơ - Nghỉ phép
            If tctrl_main.TabPages.Contains(tp_nghiphep) Then tctrl_main.TabPages.Remove(tp_nghiphep)
        End If
        If Not (Globals.IsIntersect(";131;132;133;134;", _roles)) Then       ' Hồ sơ - Cán bộ tai nạn
            If tctrl_main.TabPages.Contains(tp_cbtainan) Then tctrl_main.TabPages.Remove(tp_cbtainan)
        End If
        If (Globals.Roles.IndexOf(";132;") < 0 And Globals.Roles.IndexOf(";133;") < 0) Then
            btn_showdetail.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện select row trên lưới dữ liệu --> Fill dữ liệu vào các controls
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi (Id record) đang select</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String)
        Dim dr As DataRow
        strCodeId = _Idnew
        'Xét trường hợp thao tác với từng hồ sơ
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"      'Hồ sơ Bảo hiểm y tế
                dr = _HS_CsLaodong.GetRecord(strCodeId, 2)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        edt_bhyt_sothe.Text = dr("SoThe").ToString().Trim()
                        If dr("NgayCap").ToString().Trim() <> "" Then
                            dtpk_bhyt_ngaycap.Value = CType(dr("NgayCap").ToString(), DateTime)
                        End If
                        edt_bhyt_noicap.Text = dr("NoiCap").ToString().Trim()
                        edt_bhyt_noi_dkkb.Text = dr("NoiDangKy_KCB").ToString().Trim()
                        If dr("TuNgay").ToString().Trim() <> "" Then
                            dtpk_bhyt_tungay.Value = CType(dr("TuNgay").ToString(), DateTime)
                        End If
                        If dr("DenNgay").ToString().Trim() <> "" Then
                            dtpk_bhyt_denngay.Value = CType(dr("DenNgay").ToString(), DateTime)
                        End If
                        edt_bhyt_ghichu.Text = dr("GhiChu").ToString().Trim()
                    End If
                End If
            Case "tp_suckhoe"   'Hồ sơ khám sức khoẻ
                dr = _HS_CsLaodong.GetRecord(strCodeId, 3)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        'Khai báo một biến datetime
                        Dim currentDate As DateTime = New DateTime(CType(dr("Nam").ToString().Trim(), Integer), 1, 1)
                        dtpk_suckhoe_nam.Value = currentDate
                        edt_suckhoe_dotkham.Value = CType(dr("Dot").ToString().Trim(), Byte)
                        edt_suckhoe_noidung.Text = dr("NoiDung").ToString().Trim()
                        edt_suckhoe_noikham.Text = dr("NoiKham").ToString().Trim()

                        If dr("TuNgay").ToString().Trim() <> "" Then
                            dtpk_suckhoe_tungay.Value = CType(dr("TuNgay").ToString(), DateTime)
                        End If
                        If dr("DenNgay").ToString().Trim() <> "" Then
                            dtpk_suckhoe_denngay.Value = CType(dr("DenNgay").ToString(), DateTime)
                        End If
                        edt_suckhoe_cannang.Text = IIf(dr("CanNang").ToString().Trim() <> "", CType(dr("CanNang"), Double).ToString("N2"), "0")
                        edt_suckhoe_chieucao.Text = IIf(dr("ChieuCao").ToString().Trim() <> "", CType(dr("ChieuCao"), Double).ToString("N2"), "0")
                        cb_suckhoe_loaisk.SelectedIndex = CType(dr("LoaiSucKhoe").ToString().Trim(), Byte)
                        edt_suckhoe_ketluan.Text = dr("KetLuan").ToString().Trim()
                        edt_suckhoe_ghichu.Text = dr("GhiChu").ToString().Trim()
                    End If
                End If
            Case "tp_nghiphep"  'Hồ sơ nghỉ phép
                Using db_row As DataTable = _HS_CsLaodong.HS_NgPh_GetSearch("", "", "", strCodeId, 0, 0, "", "", 0)
                    If Not (db_row Is Nothing) Then
                        If (db_row.Rows.Count > 0) Then
                            'Khai báo một biến datetime
                            Dim currentDate As DateTime = New DateTime(CType(db_row.Rows(0)("Nam").ToString().Trim(), Integer), 1, 1)
                            dtpk_hsnghiphep_namnghi.Value = currentDate
                            If db_row.Rows(0)("TuNgay").ToString().Trim() <> "" Then
                                dtpk_hsnghiphep_tungay.Value = CType(db_row.Rows(0)("TuNgay").ToString(), DateTime)
                            End If
                            If db_row.Rows(0)("DenNgay").ToString().Trim() <> "" Then
                                dtpk_hsnghiphep_denngay.Value = CType(db_row.Rows(0)("DenNgay").ToString(), DateTime)
                            End If
                            edt_hsnghiphep_songay.Text = CType(db_row.Rows(0)("SoNgay").ToString().Trim(), Int16)
                            cb_hsnghiphep_loai_ngphep.SelectedIndex = CType(arr_LoaiNP.IndexOf(db_row.Rows(0)("IdLoaiNghi").ToString()), Integer)
                            edt_hsnghiphep_noinghi.Text = db_row.Rows(0)("NoiNghi").ToString().Trim()
                            edt_hsnghiphep_lydo.Text = db_row.Rows(0)("LyDo").ToString().Trim()
                            edt_hsnghiphep_tyle.Text = db_row.Rows(0)("TyleHuong").ToString().Trim()
                            edt_hsnghiphep_tientrcap.Text = db_row.Rows(0)("TienTroCap").ToString().Trim()
                            edt_hsnghiphep_nguoiky.Text = db_row.Rows(0)("NguoiKy_Nghi").ToString().Trim()
                            edt_hsnghiphep_ghichu.Text = db_row.Rows(0)("GhiChu").ToString().Trim()
                        End If
                    End If
                End Using

            Case "tp_cbtainan"  'Hồ sơ cán bộ tai nạn
                dr = _HS_CsLaodong.GetRecord(strCodeId, 5)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        cb_cbtainan_tengoi.SelectedIndex = CType(arr_Tengoi.IndexOf(dr("IdVuTN").ToString()), Integer)
                        cb_cbtainan_mucdo.SelectedIndex = CType(dr("MucDo_TN").ToString().Trim(), Byte)
                        edt_cbtainan_songay.Text = dr("SoNgayNghi").ToString().Trim()
                        edt_cbtainan_sotien_th.Text = dr("ChiPhi").ToString().Trim()
                        edt_cbtainan_tt_thuongtat.Text = dr("TT_ThuongTat").ToString().Trim()
                        edt_cbtainan_noi_ppdt.Text = dr("Noi_PP_DieuTri").ToString().Trim()
                        edt_cbtainan_ghichu.Text = dr("GhiChu").ToString().Trim()
                    End If
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Hàm trả về tên tab hồ sơ hiện tại đang thao tác
    ''' </summary>
    ''' <returns>Tên hồ sơ đang thao tác</returns>
    ''' <remarks></remarks>
    Private Function GetName() As String
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhxh"
                Return "hồ sơ bảo hiểm xã hội"
            Case "tp_bhyt"
                Return "hồ sơ bảo hiểm y tế"
            Case "tp_suckhoe"
                Return "hồ sơ khám sức khoẻ"
            Case "tp_nghiphep"
                Return "hồ sơ nghỉ phép"
            Case "tp_cbtainan"
                Return "hồ sơ cán bộ tai nạn"
            Case Else
                Return ""
        End Select
    End Function
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHS_CsLaoDong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Fill data ra cây dữ liệu Treeview - Cơ cấu tổ chức
        _PersonnelFile.Fill_Tree(tv_main)
        'Create frame grid - Định nghĩa lưới dữ liệu
        _HS_CsLaodong.Create_Frame(dgv_bhyt, 2)
        _HS_CsLaodong.Create_Frame(dgv_hoso_khamsk, 3)
        _HS_CsLaodong.Create_Frame(dgv_hsnghiphep, 4)
        _HS_CsLaodong.Create_Frame(dgv_cbtainan, 5)

        'Bind dữ liệu ComBobox: Loại nghỉ phép - với id_goc = 34
        arr_LoaiNP.Clear()
        cb_hsnghiphep_loai_ngphep.Items.Clear()
        arr_LoaiNP = _Globals.Bind_ComBoBox(cb_hsnghiphep_loai_ngphep, clsHT_DanhMuc.Sql_LoaiNP, "---Loại nghỉ phép---")
        'Bind dữ liệu ComBobox: Tên vụ tai nạn
        Dim strSQL As String = ""
        strSQL = String.Format("Select IdVuTaiNan, TenVu_TN From HS_VuTaiNan Order by TenVu_TN")
        arr_Tengoi.Clear()
        cb_cbtainan_tengoi.Items.Clear()
        arr_Tengoi = _Globals.Bind_ComBoBox(cb_cbtainan_tengoi, strSQL, "---Tên vụ tai nạn---")
        strNode = ""
        'Reset and set status of controls
        ResetAll_Controls(0)
        lkl_xemct.Enabled = False
        Check_Permits()

        'Nếu được gọi từ Hồ sơ cán bộ
        If (FlagShow = True) Then
            Dim _DonviId As Integer = 0         'Id đơn vị của cán bộ
            Dim _CodeBranch As String = ""      'Mã hiệu của chi nhánh cap tinh hoac tuong duong
            Dim _PhongBanId As Integer = 0      'Id phòng ban của cán bộ
            If (TagNode <> "") Then
                Dim arrElement() As String = TagNode.Split("_")
                If (arrElement.Length > 0) Then
                    If TagNode.Substring(TagNode.ToString.Length - 2, 2) <> "00" Then   'Trường hợp đơn vị là PGD
                        'Tim Id don vi cua can bo dua vao id PGD
                        _DonviId = _ListDocument.GetRootId(CType(arrElement(1).ToString().Trim(), Integer))
                    Else                                    'Trường hợp đơn vị là tỉnh hoặc tương đương
                        _DonviId = CType(arrElement(1).ToString().Trim(), Integer)
                    End If
                End If
            End If
            If (_DonviId > 0) Then
                _CodeBranch = _ListDocument.GetCodeForId(_DonviId)
                _PhongBanId = _PersonnelFile.GetDebtId(HumanId, _DonviId)
            End If
            Dim _NodeFind As String = ""
            Dim _node As TreeNode = Nothing
            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
            'Nếu là Cài đặt tại Hội sở
            If (DONVI = gMaDonViTW) Then
                _NodeFind = "ROOT_1_" & gMaDonViTW
                'Thực hiện tìm Cấp 0
                _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_node Is Nothing) Then
                    tv_main.SelectedNode = _node
                End If
                _NodeFind = "DV_" + _DonviId.ToString() + "_" + _CodeBranch
            Else
                _NodeFind = "ROOT_" + _DonviId.ToString() + "_" + _CodeBranch
            End If

            'Tìm cấp 1 (Gốc root: Ví dụ ROOT_3_10500)
            If (_NodeFind <> "") Then
                _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                If Not (_node Is Nothing) Then
                    tv_main.SelectedNode = _node
                End If
                'Kiem tra xem Phong Ban co cua Huyen Khong?
                If TagNode.Substring(TagNode.ToString.Length - 2, 2) <> "00" Then
                    _node = Globals.TreeViewFindNode(tv_main.Nodes, TagNode)
                    If Not (_node Is Nothing) Then
                        tv_main.SelectedNode = _node
                    End If
                    'Tim den Phong ban cua PGD
                    _NodeFind = "PB_" + TagNode.Substring(TagNode.IndexOf("_") + 1, TagNode.LastIndexOf("_") - TagNode.IndexOf("_") - 1) + "_" + _PhongBanId.ToString()
                Else
                    'Tìm đến cấp 2 (Ví dụ: PB_3_25)
                    _NodeFind = "PB_" + _DonviId.ToString() + "_" + _PhongBanId.ToString()
                End If
                If (_NodeFind <> "") Then
                    _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                    If Not (_node Is Nothing) Then
                        tv_main.SelectedNode = _node
                    End If
                    'Tìm đến cấp 3 (Ví dụ: CB_CNTT00000000004)
                    _NodeFind = "CB_" + HumanId.ToString()
                    If (_NodeFind <> "") Then
                        _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeFind)
                        If Not (_node Is Nothing) Then
                            tv_main.SelectedNode = _node
                        End If
                    End If
                End If
            End If
            tctrl_main.SelectedIndex = TabVal
        End If

        If HumanId <> "" Then
            If checkRight_CreateRecord(HumanId) Then
                btn_accept.Enabled = True
                btn_delete.Enabled = True
            Else
                btn_accept.Enabled = False
                btn_delete.Enabled = False
            End If
        End If

    End Sub

    Private Sub btn_showdetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_showdetail.Click
        obj_hs_vutainan = New frmHS_VuTaiNan()
        obj_hs_vutainan.Index = CType(cb_cbtainan_tengoi.SelectedIndex, Integer)
        obj_hs_vutainan.Progress_Changed = New frmHS_VuTaiNan.ProgressChangedEventHandler(AddressOf ReLoad_Main)
        obj_hs_vutainan.ShowDialog()
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        Try
            strNode = ""
            lkl_xemct.Enabled = False
            If (tv_main.Nodes.Count > 0) Then
                'Thực hiện Load child node khi click vào parent node 
                Dim arrElement() As String
                If tv_main.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not (tv_main.SelectedNode.IsExpanded) Then
                        arrElement = tv_main.SelectedNode.Tag.ToString().Trim().Split("_")
                        If (arrElement.Length > 0) Then
                            If (arrElement(0) = "DV") Then
                                _PersonnelFile.Fill_Node(tv_main.SelectedNode, arrElement(1), arrElement(2))
                            ElseIf arrElement(0) = "PB" Then
                                _PersonnelFile.Fill_NodeCanbo(tv_main.SelectedNode, arrElement(1), arrElement(2), True)
                            End If
                        End If
                    End If
                End If
                'Thực hiện set lại các control
                strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                'Lấy Chỉ số xác định đơn vị Hiện tại của cán bộ (Xét với trường hợp cài đặt đơn vị '" & gMaDonViTW & "')
                If (DONVI = gMaDonViTW) Then
                    If (strNode.IndexOf("DV_") >= 0) Then
                        _IdDonviHT = CType(strNode.Substring(3, 1), Integer)
                    End If
                End If
                ResetAll_Controls(0)
                'Fill data các hồ sơ liên quan
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "CB") Then
                    strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                    _IdCanBo = tv_main.SelectedNode.Tag.ToString().Trim().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1)
                    If (_IdCanBo <> "") Then
                        If (Globals.Roles.IndexOf(";71;") < 0) Then
                            lkl_xemct.Enabled = False
                        Else
                            lkl_xemct.Enabled = True
                        End If
                        Dim dr As DataRow
                        dr = _PersonnelFile.GetHuman_ForCode(_IdCanBo)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                lbl_macb.Text = dr("MaCB").ToString().Trim()
                                lbl_hoten.Text = dr("HoTen").ToString().Trim()
                                lbl_gioitinh.Text = IIf(dr("GioiTinh").ToString() = False, "Nam", "Nữ")
                                If dr("NgaySinh").ToString().Trim() <> "" Then
                                    lbl_ngaysinh.Text = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                End If
                                lbl_so_cmt.Text = dr("CMT_So").ToString().Trim()
                                If (dr("IdNew").ToString().Trim() <> "") Then  'Nếu cán bộ đã chuyển sang loại Ngắn hạn rồi
                                    btn_accept.Visible = False
                                    btn_add.Visible = False
                                Else
                                    btn_accept.Visible = True
                                    btn_add.Visible = True
                                End If
                            End If
                        End If
                        FillAll_Documents(_IdCanBo)

                        If checkRight_CreateRecord(_IdCanBo) Then
                            btn_accept.Enabled = True
                            btn_delete.Enabled = True
                        Else
                            btn_accept.Enabled = False
                            btn_delete.Enabled = False
                        End If
                    End If
                End If
                tctrl_main_SelectedIndexChanged(sender, Nothing)
                tv_main.SelectedNode.Expand()
            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị hồ sơ khi click vào cây dữ liệu: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub lkl_xemct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkl_xemct.LinkClicked
        If (_IdCanBo <> "") Then
            Dim obj_detail As frmHS_ChiTiet = New frmHS_ChiTiet()
            'Lấy danh sách mảng các id hiện có trên lưới dl
            Dim arrRows As ArrayList = New ArrayList()
            arrRows.Add(_IdCanBo)
            obj_detail.RecordCurrent = 0
            obj_detail.Records = 1
            obj_detail.IdCanBo = _IdCanBo
            obj_detail._IdDonviHT = _IdDonviHT
            obj_detail.arr_RecordId = arrRows
            obj_detail.ShowDialog()
        Else
            Return
        End If
    End Sub
#End Region

#Region "---> Bắt các sự kiện ngoại lệ người dùng <---"
    'Băt sự kiện người dùng nhập - Số ngày nghỉ phép của cán bộ

    Private Sub edt_hsnghiphep_songay_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_hsnghiphep_songay.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    'Băt sự kiện người dùng nhập - Số ngày nghỉ do Cán bộ Liên quan đến vụ tai nạn
    Private Sub edt_cbtainan_songay_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_cbtainan_songay.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    'Băt sự kiện người dùng nhập - Cân nặng - Chiều cao của cán bộ trong hồ sơ Khám sức khoẻ
    Private Sub edt_suckhoe_cannang_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_suckhoe_cannang.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If (KeyAscii = 46) Then
            Dim n As Integer = edt_suckhoe_cannang.Text.IndexOf(".")
            If (n >= 0) Then e.Handled = True
        End If
        If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46) Then
            e.Handled = True
        End If
    End Sub

    Private Sub edt_suckhoe_cannang_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_cannang.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_suckhoe_cannang.Text.Trim() = "") Then edt_suckhoe_cannang.Text = "0"
            edt_suckhoe_chieucao.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_chieucao_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_suckhoe_chieucao.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If (KeyAscii = 46) Then
            Dim n As Integer = edt_suckhoe_chieucao.Text.IndexOf(".")
            If (n >= 0) Then e.Handled = True
        End If
        If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46) Then
            e.Handled = True
        End If
    End Sub

    Private Sub edt_suckhoe_chieucao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_chieucao.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_suckhoe_chieucao.Text.Trim() = "") Then edt_suckhoe_chieucao.Text = "0"
            cb_suckhoe_loaisk.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_cannang_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_suckhoe_cannang.Leave
        If (edt_suckhoe_cannang.Text.Trim() = "") Then
            edt_suckhoe_cannang.Text = "0"
        End If
    End Sub

    Private Sub edt_suckhoe_chieucao_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_suckhoe_chieucao.Leave
        If (edt_suckhoe_chieucao.Text.Trim() = "") Then
            edt_suckhoe_chieucao.Text = "0"
        End If
    End Sub

    Private Sub edt_suckhoe_cannang_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles edt_suckhoe_cannang.MouseClick
        edt_suckhoe_cannang.SelectAll()
    End Sub

    Private Sub edt_suckhoe_chieucao_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles edt_suckhoe_chieucao.MouseClick
        edt_suckhoe_chieucao.SelectAll()
    End Sub

    Private Sub edt_hsnghiphep_songay_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles edt_hsnghiphep_songay.MouseClick
        edt_hsnghiphep_songay.SelectAll()
    End Sub

    Private Sub edt_cbtainan_songay_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles edt_cbtainan_songay.MouseClick
        edt_cbtainan_songay.SelectAll()
    End Sub

    Private Sub edt_hsnghiphep_tyle_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_hsnghiphep_tyle.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_hsnghiphep_tyle_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles edt_hsnghiphep_tyle.MouseClick
        edt_hsnghiphep_tyle.SelectAll()
    End Sub

    Private Sub frmHS_CsLaoDong_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub
#End Region

#Region "---> Events: Cell click of datagridview <---"
    Private Sub dgv_bhyt_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_bhyt.CellClick
        Try
            ResetAll_Controls(1)
            If (dgv_bhyt.Rows.Count > 0) Then
                If ((dgv_bhyt.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ bảo hiểm y tế của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_bhyt_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_bhyt.KeyUp
        dgv_bhyt_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_hoso_khamsk_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hoso_khamsk.CellClick
        Try
            ResetAll_Controls(2)
            If (dgv_hoso_khamsk.Rows.Count > 0) Then
                If ((dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ khám sức khoẻ của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_hoso_khamsk_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hoso_khamsk.KeyUp
        dgv_hoso_khamsk_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_hsnghiphep_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_hsnghiphep.CellClick
        Try
            ResetAll_Controls(3)
            If (dgv_hsnghiphep.Rows.Count > 0) Then
                If ((dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ nghỉ phép của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_hsnghiphep_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_hsnghiphep.KeyUp
        dgv_hsnghiphep_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_cbtainan_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_cbtainan.CellClick
        Try
            ResetAll_Controls(4)
            If (dgv_cbtainan.Rows.Count > 0) Then
                If ((dgv_cbtainan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ nghỉ phép của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_cbtainan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_cbtainan.KeyUp
        dgv_cbtainan_CellClick(sender, Nothing)
    End Sub

    Private Sub tctrl_main_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tctrl_main.SelectedIndexChanged
        btn_add.Enabled = True
        btn_delete.Enabled = True
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"      'Hồ sơ Bảo hiểm y tế
                dgv_bhyt_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";119;") < 0) Then
                    dgv_bhyt.Enabled = False
                Else
                    dgv_bhyt.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";120;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";122;") < 0) Then
                    btn_delete.Enabled = False
                End If
            Case "tp_suckhoe"   'Hồ sơ khám sức khoẻ
                dgv_hoso_khamsk_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";123;") < 0) Then
                    dgv_hoso_khamsk.Enabled = False
                Else
                    dgv_hoso_khamsk.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";124;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";126;") < 0) Then
                    btn_delete.Enabled = False
                End If

            Case "tp_nghiphep"  'Hồ sơ nghỉ phép
                dgv_hsnghiphep_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";127;") < 0) Then
                    dgv_hsnghiphep.Enabled = False
                Else
                    dgv_hsnghiphep.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";128;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";130;") < 0) Then
                    btn_delete.Enabled = False
                End If
            Case "tp_cbtainan"  'Hồ sơ cán bộ tai nạn
                dgv_cbtainan_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";131;") < 0) Then
                    dgv_cbtainan.Enabled = False
                Else
                    dgv_cbtainan.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";132;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";134;") < 0) Then
                    btn_delete.Enabled = False
                End If
        End Select
    End Sub
#End Region

#Region "---> Events: Các sự kiện cập nhật dữ liệu<---"
    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        Dim _currRow As String = strCodeId
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"      'Hồ sơ Bảo hiểm y tế
                ResetAll_Controls(1)
                If (dgv_bhyt.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_bhyt.CurrentRow.Selected = False
                dgv_bhyt.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_bhyt, "cln_Code")).Selected = True
                If (dgv_bhyt.Rows.Count > 0) Then
                    If ((dgv_bhyt.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
            Case "tp_suckhoe"   'Hồ sơ khám sức khoẻ
                ResetAll_Controls(2)
                If (dgv_hoso_khamsk.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_hoso_khamsk.CurrentRow.Selected = False
                dgv_hoso_khamsk.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hoso_khamsk, "cln_Code")).Selected = True
                If (dgv_hoso_khamsk.Rows.Count > 0) Then
                    If ((dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
            Case "tp_nghiphep"  'Hồ sơ nghỉ phép
                ResetAll_Controls(3)
                If (dgv_hsnghiphep.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_hsnghiphep.CurrentRow.Selected = False
                dgv_hsnghiphep.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hsnghiphep, "cln_Code")).Selected = True
                If (dgv_hsnghiphep.Rows.Count > 0) Then
                    If ((dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
            Case "tp_cbtainan"  'Hồ sơ cán bộ tai nạn
                ResetAll_Controls(4)
                If (dgv_cbtainan.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_cbtainan.CurrentRow.Selected = False
                dgv_cbtainan.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_cbtainan, "cln_Code")).Selected = True
                If (dgv_cbtainan.Rows.Count > 0) Then
                    If ((dgv_cbtainan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
        End Select
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ cần thêm mới " + GetName() + "!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"
                ResetAll_Controls(1)
                ActiveControl = edt_bhyt_sothe
            Case "tp_suckhoe"
                ResetAll_Controls(2)
                ActiveControl = dtpk_suckhoe_nam
            Case "tp_nghiphep"
                ResetAll_Controls(3)
                cb_hsnghiphep_loai_ngphep.SelectedIndex = IIf(cb_hsnghiphep_loai_ngphep.FindString("Nghỉ phép") > 0, cb_hsnghiphep_loai_ngphep.FindString("Nghỉ phép"), 0)
                edt_hsnghiphep_tyle.Text = "100"
                'edt_hsnghiphep_nguoiky.Text = _PersonnelFile.GetVarNam("GIAMDOC")
                'edt_hsnghiphep_noinghi.Text = _PersonnelFile.GetVarNam("DIABAN")
                ActiveControl = dtpk_hsnghiphep_namnghi
            Case "tp_cbtainan"
                ResetAll_Controls(4)
                ActiveControl = cb_cbtainan_tengoi
        End Select
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If (_IdCanBo = "") Then Return
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"
                If (dgv_bhyt.Rows.Count <= 0) Then
                    Return
                End If
                'Thực hiện xoá dữ liệu - Hồ sơ bảo hiểm Y tế
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_bhyt.Rows.Count > 0) Then
                        If ((dgv_bhyt.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Integer = 0 To dgv_bhyt.Rows.Count - 1
                                If (dgv_bhyt.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_bhyt.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_bhyt.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ bảo hiểm y tế đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CsLaodong.Delete_HS_BHYT(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ bảo hiểm y tế!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetAll_Controls(1)
                            dgv_bhyt.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_bhyt_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_bhyt, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ bảo hiểm y tế: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try

            Case "tp_suckhoe"
                If (dgv_hoso_khamsk.Rows.Count <= 0) Then
                    Return
                End If
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_hoso_khamsk.Rows.Count > 0) Then
                        If ((dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Integer = 0 To dgv_hoso_khamsk.Rows.Count - 1
                                If (dgv_hoso_khamsk.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_hoso_khamsk.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_hoso_khamsk.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ khám sức khoẻ đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CsLaodong.Delete_HS_SucKhoe(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ khám sức khoẻ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetAll_Controls(2)
                            dgv_hoso_khamsk.Rows.Clear()
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_hoso_khamsk_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_hoso_khamsk, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ khám sức khoẻ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try

            Case "tp_nghiphep"
                If (dgv_hsnghiphep.Rows.Count <= 0) Then
                    Return
                End If
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_hsnghiphep.Rows.Count > 0) Then
                        If ((dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Integer = 0 To dgv_hsnghiphep.Rows.Count - 1
                                If (dgv_hsnghiphep.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_hsnghiphep.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_hsnghiphep.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ nghỉ phép đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CsLaodong.Delete_HS_NghiPhep(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ nghỉ phép của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetAll_Controls(3)
                            dgv_hsnghiphep.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_hsnghiphep_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_hsnghiphep, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Return
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ nghỉ phép: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try

            Case "tp_cbtainan"
                If (dgv_cbtainan.Rows.Count <= 0) Then
                    Return
                End If
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_cbtainan.Rows.Count > 0) Then
                        If ((dgv_cbtainan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Integer = 0 To dgv_cbtainan.Rows.Count - 1
                                If (dgv_cbtainan.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_cbtainan.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_cbtainan.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ cán bộ tai nạn đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _HS_CsLaodong.Delete_CB_TaiNan(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ tai nạn của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetAll_Controls(4)
                            dgv_cbtainan.Rows.Clear()

                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_cbtainan_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_cbtainan, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ cán bộ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
        End Select
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        If (_IdCanBo = "") Then Return
        Dim _currRow As String = ""
        Dim _rowIndex As Integer = 0
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_bhyt"      'Hồ sơ Bảo hiểm y tế
                If (strCodeId <> "") Then
                    _currRow = strCodeId
                    If (Globals.Roles.IndexOf(";121;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ bảo hiểm y tế của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(1)
                        dgv_bhyt.CurrentRow.Selected = False
                        dgv_bhyt.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_bhyt, "cln_Code")).Selected = True
                        If (dgv_bhyt.Rows.Count > 0) Then
                            If ((dgv_bhyt.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                End If
                If (IsValid()) Then
                    Dim obj_bhyt As clsHS_CsLaodong.HS_BHYT = New clsHS_CsLaodong.HS_BHYT()
                    obj_bhyt.IdCanBo = _IdCanBo
                    obj_bhyt.SoThe = Globals.Find_Replace(edt_bhyt_sothe.Text.ToString().Trim())
                    obj_bhyt.NgayCap = dtpk_bhyt_ngaycap.Value
                    obj_bhyt.NoiCap = Globals.Find_Replace(edt_bhyt_noicap.Text.ToString().Trim())
                    obj_bhyt.NoiDangKy_KCB = Globals.Find_Replace(edt_bhyt_noi_dkkb.Text.ToString().Trim())
                    obj_bhyt.TuNgay = dtpk_bhyt_tungay.Value
                    obj_bhyt.DenNgay = dtpk_bhyt_denngay.Value
                    obj_bhyt.GhiChu = Globals.Find_Replace(edt_bhyt_ghichu.Text.ToString().Trim())
                    If (strCodeId = "") Then
                        _currRow = _HS_CsLaodong.Insert_HS_BHYT(obj_bhyt)
                    Else
                        obj_bhyt.IdBHYT = strCodeId
                        _HS_CsLaodong.Update_HS_BHYT(obj_bhyt)
                        _currRow = strCodeId
                    End If
                    'Thực hiện load lại thông tin sau khi thêm mới dữ liệu
                    ResetAll_Controls(1)
                    dgv_bhyt.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    'Gọi lại dự kiện Row click
                    dgv_bhyt.CurrentRow.Selected = False
                    dgv_bhyt.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_bhyt, "cln_Code")).Selected = True
                    If (dgv_bhyt.Rows.Count > 0) Then
                        If ((dgv_bhyt.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_bhyt.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                End If
            Case "tp_suckhoe"   'Hồ sơ khám sức khoẻ
                If (strCodeId <> "") Then
                    _currRow = strCodeId
                    If (Globals.Roles.IndexOf(";125;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ khám sức khoẻ của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(2)
                        dgv_hoso_khamsk.CurrentRow.Selected = False
                        dgv_hoso_khamsk.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hoso_khamsk, "cln_Code")).Selected = True
                        If (dgv_hoso_khamsk.Rows.Count > 0) Then
                            If ((dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                End If
                If (IsValid()) Then
                    Dim obj_suckhoe As clsHS_CsLaodong.HS_SucKhoe = New clsHS_CsLaodong.HS_SucKhoe()
                    obj_suckhoe.IdCanBo = _IdCanBo
                    obj_suckhoe.Nam = dtpk_suckhoe_nam.Value.Year
                    obj_suckhoe.Dot = CType(edt_suckhoe_dotkham.Value, Byte)
                    obj_suckhoe.NoiDung = Globals.Find_Replace(edt_suckhoe_noidung.Text.ToString().Trim())
                    obj_suckhoe.NoiKham = Globals.Find_Replace(edt_suckhoe_noikham.Text.ToString().Trim())
                    obj_suckhoe.TuNgay = dtpk_suckhoe_tungay.Value
                    obj_suckhoe.DenNgay = dtpk_suckhoe_denngay.Value
                    obj_suckhoe.CanNang = IIf(edt_suckhoe_cannang.Text.Trim() <> "", CType(edt_suckhoe_cannang.Text.Trim().Replace(" ", ""), Double), "0")
                    obj_suckhoe.ChieuCao = IIf(edt_suckhoe_chieucao.Text.Trim() <> "", CType(edt_suckhoe_chieucao.Text.Trim().Replace(" ", ""), Double), 0)
                    obj_suckhoe.LoaiSucKhoe = CType(cb_suckhoe_loaisk.SelectedIndex, Byte)
                    obj_suckhoe.KetLuan = Globals.Find_Replace(edt_suckhoe_ketluan.Text.ToString().Trim())
                    obj_suckhoe.GhiChu = Globals.Find_Replace(edt_suckhoe_ghichu.Text.ToString().Trim())

                    If (strCodeId = "") Then
                        _currRow = _HS_CsLaodong.Insert_HS_SucKhoe(obj_suckhoe)
                    Else
                        obj_suckhoe.IdKhamChua = strCodeId
                        _HS_CsLaodong.Update_HS_SucKhoe(obj_suckhoe)
                        _currRow = strCodeId
                    End If

                    'Thực hiện load lại thông tin sau khi thêm mới dữ liệu
                    ResetAll_Controls(2)
                    dgv_hoso_khamsk.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    'Gọi lại sự kiện Row click trên lưới dữ liệu
                    dgv_hoso_khamsk.CurrentRow.Selected = False
                    dgv_hoso_khamsk.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hoso_khamsk, "cln_Code")).Selected = True
                    If (dgv_hoso_khamsk.Rows.Count > 0) Then
                        If ((dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hoso_khamsk.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                End If

            Case "tp_nghiphep"  'Hồ sơ nghỉ phép
                If (strCodeId <> "") Then
                    _currRow = strCodeId
                    If (Globals.Roles.IndexOf(";129;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ nghỉ phép của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(3)
                        dgv_hsnghiphep.CurrentRow.Selected = False
                        dgv_hsnghiphep.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hsnghiphep, "cln_Code")).Selected = True
                        If (dgv_hsnghiphep.Rows.Count > 0) Then
                            If ((dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                End If
                If (IsValid()) Then
                    Dim obj_nghiphep As clsHS_CsLaodong.HS_NghiPhep = New clsHS_CsLaodong.HS_NghiPhep()
                    obj_nghiphep.IdCanBo = _IdCanBo
                    obj_nghiphep.Nam = dtpk_hsnghiphep_namnghi.Value.Year
                    obj_nghiphep.TuNgay = dtpk_hsnghiphep_tungay.Value
                    obj_nghiphep.DenNgay = dtpk_hsnghiphep_denngay.Value
                    If String.IsNullOrEmpty(edt_hsnghiphep_songay.Text.Trim()) Then
                        obj_nghiphep.SoNgay = 0
                    Else : obj_nghiphep.SoNgay = CType(edt_hsnghiphep_songay.Text.Trim().Replace(" ", "").Replace(",", "").Replace(".", ""), Int16)
                    End If

                    obj_nghiphep.IdLoaiNghi = CType(IIf(arr_LoaiNP.Count > 0, arr_LoaiNP(cb_hsnghiphep_loai_ngphep.SelectedIndex), "0"), Integer)
                    obj_nghiphep.NoiNghi = Globals.Find_Replace(edt_hsnghiphep_noinghi.Text.ToString().Trim())
                    obj_nghiphep.LyDo = Globals.Find_Replace(edt_hsnghiphep_lydo.Text.ToString().Trim())
                    If String.IsNullOrEmpty(edt_hsnghiphep_tyle.Text.Trim().ToString().Trim()) Then
                        obj_nghiphep.TyleHuong = 0
                    Else
                        obj_nghiphep.TyleHuong = CType(Globals.Find_Replace(edt_hsnghiphep_tyle.Text.Trim().ToString()), Double)
                    End If
                    obj_nghiphep.TienTroCap = IIf(edt_hsnghiphep_tientrcap.Text.Trim() <> "", CType(edt_hsnghiphep_tientrcap.Text.Trim().Replace(",", "").Replace(" ", ""), Double), 0)
                    obj_nghiphep.NguoiKy_Nghi = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_hsnghiphep_nguoiky.Text.ToString().Trim()))
                    obj_nghiphep.GhiChu = Globals.Find_Replace(edt_hsnghiphep_ghichu.Text.ToString().Trim())
                    obj_nghiphep.PosCode = DONVI 'IdDONVI
                    obj_nghiphep.CreatedBy = Globals.UserVal
                    obj_nghiphep.ModifiedBy = Globals.UserVal
                    If (strCodeId = "") Then
                        obj_nghiphep.IdNghiPhep = ""

                    Else
                        obj_nghiphep.IdNghiPhep = strCodeId
                    End If
                    _currRow = _HS_CsLaodong.Insert_Update_HS_NgPh(obj_nghiphep)

                    'Thực hiện load lại thông tin sau khi thêm mới dữ liệu
                    ResetAll_Controls(3)
                    dgv_hsnghiphep.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    'Gọi lại sự kiện row click của lưới dữ liệu
                    dgv_hsnghiphep.CurrentRow.Selected = False
                    dgv_hsnghiphep.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_hsnghiphep, "cln_Code")).Selected = True
                    If (dgv_hsnghiphep.Rows.Count > 0) Then
                        If ((dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_hsnghiphep.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                End If

            Case "tp_cbtainan"  'Hồ sơ cán bộ tai nạn
                If (strCodeId <> "") Then
                    _currRow = strCodeId
                    If (Globals.Roles.IndexOf(";133;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ cán bộ tai nạn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(4)
                        'Gọi lại sự kiện cell click của lưới dữ liệu
                        dgv_cbtainan.CurrentRow.Selected = False
                        dgv_cbtainan.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_cbtainan, "cln_Code")).Selected = True
                        If (dgv_cbtainan.Rows.Count > 0) Then
                            If ((dgv_cbtainan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                End If
                If (IsValid()) Then
                    Dim obj_cbtainan As clsHS_CsLaodong.CB_TaiNan = New clsHS_CsLaodong.CB_TaiNan()
                    obj_cbtainan.IdCanBo = _IdCanBo
                    obj_cbtainan.IdVuTN = IIf(arr_Tengoi.Count > 0, arr_Tengoi(cb_cbtainan_tengoi.SelectedIndex), "")
                    obj_cbtainan.MucDo_TN = CType(cb_cbtainan_mucdo.SelectedIndex, Byte)
                    obj_cbtainan.SoNgayNghi = IIf(edt_cbtainan_songay.Text.Trim() <> "", CType(edt_cbtainan_songay.Text.Trim().Replace(" ", ""), Integer), 0)
                    obj_cbtainan.ChiPhi = IIf(edt_cbtainan_sotien_th.Text.Trim() <> "", CType(MoneyValue(edt_cbtainan_sotien_th.Text.Trim()), Double), 0)
                    obj_cbtainan.TT_ThuongTat = Globals.Find_Replace(edt_cbtainan_tt_thuongtat.Text.ToString().Trim())
                    obj_cbtainan.Noi_PP_DieuTri = Globals.Find_Replace(edt_cbtainan_noi_ppdt.Text.ToString().Trim())
                    obj_cbtainan.GhiChu = Globals.Find_Replace(edt_cbtainan_ghichu.Text.ToString().Trim())

                    If (strCodeId = "") Then
                        _currRow = _HS_CsLaodong.Insert_CB_TaiNan(obj_cbtainan)
                    Else
                        obj_cbtainan.IdCBTaiNan = strCodeId
                        _HS_CsLaodong.Update_CB_TaiNan(obj_cbtainan)
                        _currRow = strCodeId
                    End If
                    'Thực hiện load lại thông tin sau khi thêm mới dữ liệu
                    ResetAll_Controls(4)
                    dgv_cbtainan.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    'Gọi lại sự kiện row click của lưới dữ liệu
                    dgv_cbtainan.CurrentRow.Selected = False
                    dgv_cbtainan.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_cbtainan, "cln_Code")).Selected = True
                    If (dgv_cbtainan.Rows.Count > 0) Then
                        If ((dgv_cbtainan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_cbtainan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "---> Bắt các sự kiện - Di chuyển đến các controls bằng phím <---"
    'Tab - Hồ sơ bảo hiểm y tê
    Private Sub edt_bhyt_sothe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_sothe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_bhyt_ngaycap.Focus()
            edt_bhyt_sothe.Text = edt_bhyt_sothe.Text.ToUpper()
        End If
    End Sub

    Private Sub edt_bhyt_sothe_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_bhyt_sothe.Leave
        edt_bhyt_sothe.Text = edt_bhyt_sothe.Text.ToUpper()
    End Sub

    Private Sub dtpk_bhyt_ngaycap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_bhyt_ngaycap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_bhyt_noicap.Focus()
            If (strCodeId = "") Then
                dtpk_bhyt_tungay.Value = dtpk_bhyt_ngaycap.Value
                dtpk_bhyt_denngay.Value = New DateTime(dtpk_bhyt_ngaycap.Value.Year, 12, 31)
            End If
        End If
    End Sub

    Private Sub dtpk_bhyt_ngaycap_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_bhyt_ngaycap.Leave
        If (strCodeId = "") Then
            dtpk_bhyt_tungay.Value = dtpk_bhyt_ngaycap.Value
            dtpk_bhyt_denngay.Value = New DateTime(dtpk_bhyt_ngaycap.Value.Year, 12, 31)
        End If
    End Sub

    Private Sub edt_bhyt_noicap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_noicap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_bhyt_noi_dkkb.Focus()
        End If
    End Sub

    Private Sub edt_bhyt_noicap_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_noicap.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_bhyt_ngaycap.Focus()
        End If
    End Sub

    Private Sub edt_bhyt_noi_dkkb_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_noi_dkkb.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_bhyt_tungay.Focus()
        End If
    End Sub

    Private Sub edt_bhyt_noi_dkkb_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_noi_dkkb.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_bhyt_noicap.Focus()
        End If
    End Sub

    Private Sub dtpk_bhyt_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_bhyt_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_bhyt_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_bhyt_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_bhyt_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_bhyt_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_bhyt_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_bhyt_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_bhyt_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_bhyt_denngay.Focus()
        End If
    End Sub

    'Tab - Hồ sơ Khám sức khoẻ
    Private Sub dtpk_suckhoe_nam_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_suckhoe_nam.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_suckhoe_dotkham.Focus()
            If (strCodeId = "") Then
                dtpk_suckhoe_tungay.Value = New DateTime(dtpk_suckhoe_nam.Value.Year, 1, 1)
                dtpk_suckhoe_denngay.Value = New DateTime(dtpk_suckhoe_nam.Value.Year, 12, 31)
            End If
        End If
    End Sub

    Private Sub dtpk_suckhoe_nam_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_suckhoe_nam.Leave
        If (strCodeId = "") Then
            dtpk_suckhoe_tungay.Value = New DateTime(dtpk_suckhoe_nam.Value.Year, 1, 1)
            dtpk_suckhoe_denngay.Value = New DateTime(dtpk_suckhoe_nam.Value.Year, 12, 31)
        End If
    End Sub

    Private Sub edt_suckhoe_dotkham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_dotkham.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_suckhoe_noidung.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_noidung_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_noidung.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_suckhoe_noikham.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_noidung_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_noidung.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_suckhoe_dotkham.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_noikham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_noikham.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_suckhoe_tungay.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_noikham_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_noikham.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_suckhoe_noidung.Focus()
        End If
    End Sub

    Private Sub dtpk_suckhoe_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_suckhoe_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_suckhoe_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_suckhoe_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_suckhoe_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_suckhoe_cannang.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_cannang_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_cannang.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_suckhoe_denngay.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_chieucao_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_chieucao.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_suckhoe_cannang.Focus()
        End If
    End Sub

    Private Sub cb_suckhoe_loaisk_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_suckhoe_loaisk.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_suckhoe_ketluan.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_ketluan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_ketluan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_suckhoe_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_ketluan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_ketluan.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_suckhoe_loaisk.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_suckhoe_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_suckhoe_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_suckhoe_ketluan.Focus()
        End If
    End Sub

    'Tab - Hồ sơ nghỉ phép của cán bộ
    Private Sub dtpk_hsnghiphep_tungay_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_hsnghiphep_tungay.ValueChanged
        If (strCodeId = "") Then
            dtpk_hsnghiphep_denngay.Value = dtpk_hsnghiphep_tungay.Value.AddDays(1)
        End If
        edt_hsnghiphep_songay.Text = Globals.GetCountWorkDays(dtpk_hsnghiphep_tungay.Value, dtpk_hsnghiphep_denngay.Value).ToString().Trim()
    End Sub

    Private Sub dtpk_hsnghiphep_denngay_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_hsnghiphep_denngay.ValueChanged
        edt_hsnghiphep_songay.Text = Globals.GetCountWorkDays(dtpk_hsnghiphep_tungay.Value, dtpk_hsnghiphep_denngay.Value).ToString().Trim()
    End Sub

    Private Sub dtpk_hsnghiphep_namnghi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hsnghiphep_namnghi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hsnghiphep_tungay.Focus()
            If (strCodeId = "") Then
                dtpk_hsnghiphep_tungay.Value = New DateTime(dtpk_hsnghiphep_namnghi.Value.Year, DateTime.Now.Month, DateTime.Now.Day)
                dtpk_hsnghiphep_denngay.Value = dtpk_hsnghiphep_tungay.Value.AddDays(1)
            End If
        End If
    End Sub

    Private Sub dtpk_hsnghiphep_namnghi_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_hsnghiphep_namnghi.Leave
        If (strCodeId = "") Then
            dtpk_hsnghiphep_tungay.Value = New DateTime(dtpk_hsnghiphep_namnghi.Value.Year, DateTime.Now.Month, DateTime.Now.Day)
            dtpk_hsnghiphep_denngay.Value = dtpk_hsnghiphep_tungay.Value.AddDays(1)
        End If
    End Sub

    Private Sub dtpk_hsnghiphep_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hsnghiphep_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_hsnghiphep_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_hsnghiphep_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_hsnghiphep_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_hsnghiphep_loai_ngphep.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_songay_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hsnghiphep_songay.Leave
        If (edt_hsnghiphep_songay.Text.Trim() = "") Then
            edt_hsnghiphep_songay.Text = Globals.GetCountWorkDays(dtpk_hsnghiphep_tungay.Value, dtpk_hsnghiphep_denngay.Value).ToString().Trim()
        End If
    End Sub

    Private Sub edt_hsnghiphep_songay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_songay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hsnghiphep_songay.Text.Trim() <> "") Then
                cb_hsnghiphep_loai_ngphep.Focus()
            Else
                edt_hsnghiphep_songay.Text = "0"
                edt_hsnghiphep_songay.Focus()
            End If
        End If
    End Sub

    Private Sub edt_hsnghiphep_songay_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_songay.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_hsnghiphep_denngay.Focus()
        End If
    End Sub

    Private Sub cb_hsnghiphep_loai_ngphep_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_hsnghiphep_loai_ngphep.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_hsnghiphep_noinghi.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_noinghi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_noinghi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_hsnghiphep_lydo.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_noinghi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_noinghi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_hsnghiphep_loai_ngphep.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_hsnghiphep_tyle.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hsnghiphep_noinghi.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_tyle_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_tyle.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hsnghiphep_songay.Text.Trim() <> "") Then
                edt_hsnghiphep_tientrcap.Focus()
            Else
                edt_hsnghiphep_tyle.Focus()
            End If
        End If
    End Sub

    Private Sub edt_hsnghiphep_tyle_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_tyle.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hsnghiphep_lydo.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_tientrcap_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_tientrcap.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hsnghiphep_tientrcap.Text.Trim() = "") Then
                edt_hsnghiphep_tientrcap.Text = "0"
            End If
            edt_hsnghiphep_nguoiky.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_tientrcap_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_tientrcap.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hsnghiphep_tyle.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_tientrcap_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hsnghiphep_tientrcap.Leave
        If (edt_hsnghiphep_tientrcap.Text.Trim() = "") Then
            edt_hsnghiphep_tientrcap.Text = "0"
        End If
    End Sub

    'Thực hiện format số tiền nhập vào của người dùng
    Private Sub edt_hsnghiphep_tientrcap_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_hsnghiphep_tientrcap.TextChanged
        Try
            edt_hsnghiphep_tientrcap = formatMoneyinTextbox(edt_hsnghiphep_tientrcap)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub edt_hsnghiphep_nguoiky_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_nguoiky.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_hsnghiphep_nguoiky.Text.Trim() <> "") Then
                edt_hsnghiphep_ghichu.Focus()
            Else
                edt_hsnghiphep_nguoiky.Focus()
            End If
        End If
    End Sub

    Private Sub edt_hsnghiphep_nguoiky_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_nguoiky.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hsnghiphep_tientrcap.Focus()
        End If
    End Sub

    Private Sub edt_hsnghiphep_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_ghichu.KeyDown
        If (e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Down) Then
            ActiveControl = btn_accept
        End If
    End Sub

    Private Sub edt_hsnghiphep_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_hsnghiphep_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_hsnghiphep_nguoiky.Focus()
        End If
    End Sub

    Private Sub cb_hsnghiphep_loai_ngphep_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_hsnghiphep_loai_ngphep.SelectedIndexChanged
        edt_hsnghiphep_tyle.ReadOnly = False
        edt_hsnghiphep_tyle.Text = "0"
        If (strCodeId = "") Then
            edt_hsnghiphep_tyle.Text = "100"
        End If
        If (cb_hsnghiphep_loai_ngphep.SelectedIndex > 0 And cb_hsnghiphep_loai_ngphep.Items.Count <> 0) Then
            Dim _IdLoaiNP As Integer = CType(arr_LoaiNP(IIf(cb_hsnghiphep_loai_ngphep.SelectedIndex > 0, cb_hsnghiphep_loai_ngphep.SelectedIndex, "0")), Integer)
            If (_IdLoaiNP > 0) Then
                Dim strSQL As String = String.Format("Select * from DanhMuc where id_goc = 34 and Status = 1 And id = {0}", _IdLoaiNP)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            'Xét trường hợp nghỉ Không lương hoặc Nghỉ hưởng chế độ BHXH
                            If db.Rows(0)("ma_so").ToString().Trim() = "3402" Or db.Rows(0)("ma_so").ToString().Trim() = "3403" Then
                                edt_hsnghiphep_tyle.Text = "0"
                                edt_hsnghiphep_tyle.ReadOnly = True
                            End If
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    'Tab - Hồ sơ cán bộ tai nạn
    Private Sub cb_cbtainan_tengoi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cbtainan_tengoi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_cbtainan_mucdo.Focus()
        End If
    End Sub

    Private Sub cb_cbtainan_mucdo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_cbtainan_mucdo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_cbtainan_songay.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_songay_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_cbtainan_songay.Leave
        If (edt_cbtainan_songay.Text.Trim() = "") Then
            edt_cbtainan_songay.Text = "0"
        End If
    End Sub

    Private Sub edt_cbtainan_songay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_songay.KeyDown
        If (e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Down) Then
            If (edt_cbtainan_songay.Text.Trim() = "") Then
                edt_cbtainan_songay.Text = "0"
            End If
            ActiveControl = edt_cbtainan_sotien_th
        End If
    End Sub

    Private Sub edt_cbtainan_songay_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_songay.KeyUp
        If (e.KeyCode = Keys.Up) Then
            cb_cbtainan_mucdo.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_sotien_th_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_sotien_th.KeyDown
        If (e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Down) Then
            If edt_cbtainan_sotien_th.Text.Trim() = "" Then
                edt_cbtainan_sotien_th.Text = "0"
            End If
            ActiveControl = edt_cbtainan_tt_thuongtat
        End If
    End Sub

    Private Sub edt_cbtainan_sotien_th_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_sotien_th.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cbtainan_songay.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_sotien_th_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_cbtainan_sotien_th.Leave
        If edt_cbtainan_sotien_th.Text.Trim() = "" Then
            edt_cbtainan_sotien_th.Text = "0"
        End If
    End Sub

    Private Sub edt_cbtainan_sotien_th_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_cbtainan_sotien_th.TextChanged
        Try
            edt_cbtainan_sotien_th = formatMoneyinTextbox(edt_cbtainan_sotien_th)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub edt_cbtainan_tt_thuongtat_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_tt_thuongtat.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_cbtainan_noi_ppdt.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_tt_thuongtat_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_tt_thuongtat.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cbtainan_sotien_th.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_noi_ppdt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_noi_ppdt.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_cbtainan_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_noi_ppdt_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_noi_ppdt.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cbtainan_tt_thuongtat.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_cbtainan_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_cbtainan_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_cbtainan_noi_ppdt.Focus()
        End If
    End Sub
#End Region

End Class