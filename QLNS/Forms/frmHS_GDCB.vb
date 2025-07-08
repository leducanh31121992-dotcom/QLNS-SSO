Public Class frmHS_GDCB

#Region "---> Defined parametter and properties <---"
    Private _PersonnelFile As clsHS_CanBo = New clsHS_CanBo()
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _ListDocument As clsHT_DanhMuc = New clsHT_DanhMuc()
    'Khai báo mảng lưu danh sách các bản ghi - của dữ liệu fill ra combobox
    Private arrQuanHe As ArrayList = New ArrayList 'Tên gọi vụ tai nạn
    Private arrQueQuan As ArrayList = New ArrayList
    Private arrXaPhuong As ArrayList = New ArrayList
    Private arrQuocTich As ArrayList = New ArrayList
    Private arrUtBanThan As ArrayList = New ArrayList
    Private arrThVien As ArrayList = New ArrayList
    Private arrQuocTichGTGC As ArrayList = New ArrayList

    'Khai báo các biến lưu id của các hồ sơ
    Private _RecordId As String = ""
    'Biến lưu id cán bộ khi select trên cây dữ liệu
    Private _IdCanBo As String = ""
    'Biến lưu tag của node hiện thời đang select trên cây dữ liệu
    Private _NodeCurrent As String = ""

    Private strSQL As String = ""
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
    Public FlagShow As Boolean = False 'Giá trị True là được gọi từ Hồ sơ cán bộ. False: Giá trị mặc định
    Private _IdDonviHT As Integer = IdDONVI
    Public TagNode As String = ""       'Giá trị lưu lại tag đơn vị đang select bên hồ sơ cán bộ
#End Region

#Region "---> Functions main: Các hàm chính <---"


    ''' <summary>
    ''' Hàm thực hiện reset các controls - 
    ''' </summary>
    ''' <param name="_Status">Chỉ số xác định reset:
    '''                                   0 - Reset toàn bộ các controls
    '''                                   1 - Reset Hồ sơ Gia đình cán bộ
    '''                                   2 - Reset Hồ sơ khám sức khoẻ gia đình cán bộ
    '''                                   3 - Reset Hồ sơ Giảm trừ gia cảnh
    '''                                   4 - Reset Hồ sơ Thu nhap gia dinh
    ''' </param>
    ''' <remarks></remarks>
    Private Sub ResetControls(ByVal _Status As Byte)
        Select Case _Status
            Case 0      'Reset toàn bộ các controls
                lbl_macb.Text = ""
                lbl_hoten.Text = ""
                lbl_ngaysinh.Text = ""
                lbl_gioitinh.Text = ""
                lbl_so_cmt.Text = ""
                _RecordId = ""
                _IdCanBo = ""
                chkTWQuanLy.Checked = False

                'Hồ sơ gia đình cán bộ
                dgv_gdcb.Rows.Clear()
                edt_gdcb_hoten.Text = ""
                cb_gdcb_gioitinh.SelectedIndex = 0
                dtpk_gdcb_namsinh.Text = DateTime.Now.ToShortDateString()
                cb_gdcb_quanhe.SelectedIndex = 0
                cb_gdcb_quoctich.SelectedIndex = 1
                cb_gdcb_quequan.SelectedIndex = 0
                cb_gdcb_quequan_SelectedIndexChanged(Nothing, Nothing)
                edt_gdcb_nghenghiep.Text = ""
                edt_gdcb_dienthoai.Text = ""
                edt_gdcb_diachi.Text = ""
                cb_gdcb_utbanthan.SelectedIndex = 0
                cb_gdcb_tinhtrang.SelectedIndex = 1
                cb_gdcb_tinhtrang_SelectedIndexChanged(Nothing, Nothing)

                ''Hồ sơ Khám sức khoẻ gia đình cán bộ
                'dgv_suckhoe.Rows.Clear()
                'If (cb_sk_thanhvien.Items.Count <> 0) Then
                '    cb_sk_thanhvien.SelectedIndex = 0
                'End If

                'dtpk_sk_namkham.Text = DateTime.Now.ToShortDateString()
                'edt_sk_dotkham.Value = 0
                'edt_sk_ndkham.Text = ""
                'edt_sk_noikham.Text = ""
                'dtpk_sk_tungay.Text = DateTime.Now.ToShortDateString()
                'dtpk_sk_denngay.Text = DateTime.Now.ToShortDateString()
                'edt_sk_cannang.Text = "0.00"
                'edt_sk_chieucao.Text = "0.00"
                'cb_sk_loaisk.SelectedIndex = 0
                'edt_sk_ketluan.Text = ""
                'edt_sk_ghichu.Text = ""

                'Hồ sơ Giảm trừ gia cảnh
                dgv_gtgc.Rows.Clear()
                edt_gtgc_hoten.Text = ""
                cb_gtgc_quanhe.SelectedIndex = 0
                dtpk_gtgc_tungay.Text = DateTime.Now.ToShortDateString()
                Dim _dateVal As DateTime = New DateTime(DateTime.Now.Year, 12, 31)
                dtpk_gtgc_denngay.Value = _dateVal
                edt_gtgc_ghichu.Text = ""

            Case 1      'Reset Hồ sơ gia đình cán bộ
                _RecordId = ""
                edt_gdcb_hoten.Text = ""
                cb_gdcb_gioitinh.SelectedIndex = 0
                dtpk_gdcb_namsinh.Text = DateTime.Now.ToShortDateString()
                cb_gdcb_quanhe.SelectedIndex = 0
                cb_gdcb_quoctich.SelectedIndex = 0
                If Not (cb_gdcb_quoctich Is Nothing) Then
                    If (cb_gdcb_quoctich.Items.Count > 0) Then
                        cb_gdcb_quoctich.SelectedIndex = 1
                    End If
                End If
                cb_gdcb_quequan.SelectedIndex = 0
                edt_gdcb_nghenghiep.Text = ""
                edt_gdcb_dienthoai.Text = ""
                edt_gdcb_diachi.Text = ""
                cb_gdcb_utbanthan.SelectedIndex = 0
                cb_gdcb_tinhtrang.SelectedIndex = 0
                cb_gdcb_tinhtrang_SelectedIndexChanged(Nothing, Nothing)

                'Case 2      'Reset Hồ sơ khám sức khoẻ gia đình cán bộ
                '    _RecordId = ""
                '    If (cb_sk_thanhvien.Items.Count <> 0) Then
                '        cb_sk_thanhvien.SelectedIndex = 0
                '    End If
                '    dtpk_sk_namkham.Text = DateTime.Now.ToShortDateString()
                '    edt_sk_dotkham.Value = 0
                '    edt_sk_ndkham.Text = ""
                '    edt_sk_noikham.Text = ""
                '    dtpk_sk_tungay.Text = DateTime.Now.ToShortDateString()
                '    dtpk_sk_denngay.Text = DateTime.Now.ToShortDateString()
                '    edt_sk_cannang.Text = "0.00"
                '    edt_sk_chieucao.Text = "0.00"
                '    cb_sk_loaisk.SelectedIndex = 0
                '    edt_sk_ketluan.Text = ""
                '    edt_sk_ghichu.Text = ""

            Case 3           'Hồ sơ Giảm trừ gia cảnh
                _RecordId = ""
                edt_gtgc_hoten.Text = ""
                cb_gtgc_quanhe.SelectedIndex = 0
                dtpk_gtgc_tungay.Text = DateTime.Now.ToShortDateString()
                Dim _dateVal As DateTime = New DateTime(DateTime.Now.Year, 12, 31)
                dtpk_gtgc_denngay.Value = _dateVal
                dtpk_gtgc_denngay.Checked = False
                edt_gtgc_ghichu.Text = ""

                dtpk_gtgc_ngaysinh.Text = DateTime.Now.ToShortDateString()
                edt_gtgc_masothue.Text = ""
                cb_gtgc_quoctich.SelectedIndex = 1
                edt_gtgc_socmt.Text = ""
                dtpk_gtgc_ngaycap.Text = DateTime.Now.ToShortDateString()
                edt_gtgc_noicap.Text = ""
                edt_gtgc_diachi.Text = ""
                edt_gtgc_dienthoai.Text = ""
                rb_mo.Checked = True

            Case 4
                _RecordId = ""
                dpkThoiGianKK.Text = DateTime.Now.ToShortDateString()
                txtLuong.Text = ""
                txtNguonKhac.Text = ""
                txtNhaCap.Text = ""
                txtDT_Nha_Cap.Text = ""
                txtNhaMua.Text = ""
                txtDT_Nha_Mua.Text = ""
                txtDT_Dat_Cap.Text = ""
                txtDT_Dat_Mua.Text = ""
                txtDat_XSKD.Text = ""
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu - Các hồ sơ theo Id cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Mã hiệu xác định cán bộ</param>
    ''' <remarks></remarks>
    Private Sub FillAll_Documents(ByVal _IdCanBo As String)
        'Fill data - Hồ sơ Gia đình cán bộ
        Using db As DataTable = _PersonnelFile.GetAll(_IdCanBo, 0)
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
                        'GetAddress(CType(db.Rows(i)("IdQueQuan").ToString().Trim(), Integer), db.Rows(i)("DiaChi").ToString().Trim())
                        dgv_gdcb.Rows(i).Cells("cln_NgheNghiep").Value = db.Rows(i)("NgheNghiep").ToString().Trim()
                    Next
                End If
            End If
        End Using

        ''Fill data - Hồ sơ Khám sức khoẻ Gia đình cán bộ
        'Using db As DataTable = _PersonnelFile.GetAll(_IdCanBo, 1)
        '    If Not (db Is Nothing) Then
        '        If (db.Rows.Count > 0) Then
        '            For i As Integer = 0 To db.Rows.Count - 1
        '                dgv_suckhoe.Rows.Add()
        '                dgv_suckhoe.Rows(i).Cells("cln_Id").Value = db.Rows(i)("IdKhamchuaGDCB").ToString().Trim()
        '                dgv_suckhoe.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
        '                dgv_suckhoe.Rows(i).Cells("cln_ThVien").Value = db.Rows(i)("ThanhVien").ToString().Trim()
        '                dgv_suckhoe.Rows(i).Cells("cln_Namkham").Value = db.Rows(i)("Nam").ToString().Trim()
        '                dgv_suckhoe.Rows(i).Cells("cln_Dotkham").Value = db.Rows(i)("Dot").ToString().Trim()
        '                dgv_suckhoe.Rows(i).Cells("cln_Ndkham").Value = db.Rows(i)("NoiDungKC").ToString().Trim()
        '                dgv_suckhoe.Rows(i).Cells("cln_Noikham").Value = db.Rows(i)("NoiKham").ToString().Trim()
        '                If (db.Rows(i)("TuNgay").ToString().Trim() <> "") Then
        '                    dgv_suckhoe.Rows(i).Cells("cln_Tungay").Value = Format(CType(db.Rows(i)("TuNgay"), DateTime), "dd-MM-yyyy")
        '                Else
        '                    dgv_suckhoe.Rows(i).Cells("cln_Tungay").Value = ""
        '                End If
        '                If (db.Rows(i)("DenNgay").ToString().Trim() <> "") Then
        '                    dgv_suckhoe.Rows(i).Cells("cln_Denngay").Value = Format(CType(db.Rows(i)("DenNgay"), DateTime), "dd-MM-yyyy")
        '                Else
        '                    dgv_suckhoe.Rows(i).Cells("cln_Denngay").Value = ""
        '                End If
        '                dgv_suckhoe.Rows(i).Cells("cln_Cannang").Value = IIf(db.Rows(i)("CanNang").ToString() <> "", CType(db.Rows(i)("CanNang"), Double).ToString("N2"), "0")
        '                dgv_suckhoe.Rows(i).Cells("cln_Chieucao").Value = IIf(db.Rows(i)("ChieuCao").ToString() <> "", CType(db.Rows(i)("ChieuCao"), Double).ToString("N2"), "0")

        '                If (db.Rows(i)("LoaiSuckhoe").ToString().Trim() = "1") Then
        '                    dgv_suckhoe.Rows(i).Cells("cln_Loaisk").Value = "Loại A"
        '                ElseIf (db.Rows(i)("LoaiSuckhoe").ToString().Trim() = "2") Then
        '                    dgv_suckhoe.Rows(i).Cells("cln_Loaisk").Value = "Loại B"
        '                ElseIf (db.Rows(i)("LoaiSuckhoe").ToString().Trim() = "3") Then
        '                    dgv_suckhoe.Rows(i).Cells("cln_Loaisk").Value = "Loại C"
        '                End If
        '                dgv_suckhoe.Rows(i).Cells("cln_Ketluan").Value = db.Rows(i)("KetLuan").ToString().Trim()
        '                dgv_suckhoe.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
        '            Next
        '        End If
        '    End If
        'End Using

        'Fill data - Hồ sơ Giảm trừ gia cảnh
        Using db As DataTable = _PersonnelFile.GetAll(_IdCanBo, 2)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_gtgc.Rows.Add()
                        dgv_gtgc.Rows(i).Cells("cln_Id").Value = db.Rows(i)("IdGTGC").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_IdQuocTich").Value = db.Rows(i)("IdQuocTich").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_IdQuanHe").Value = db.Rows(i)("IdQuanHe").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_NgayCap").Value = db.Rows(i)("NgayCap").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_NoiCap").Value = db.Rows(i)("NoiCap").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_TrangThai").Value = db.Rows(i)("TrangThai").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_IdCanBo").Value = db.Rows(i)("IdCanBo").ToString().Trim()

                        dgv_gtgc.Rows(i).Cells("cln_STT").Value = CType(i + 1, String)
                        dgv_gtgc.Rows(i).Cells("cln_HoTen").Value = db.Rows(i)("HoTenNguoiPT").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_QuanHe_HT").Value = db.Rows(i)("QuanHe_HT").ToString().Trim()
                        If (Not String.IsNullOrEmpty(db.Rows(i)("NgaySinh").ToString()) And db.Rows(i)("NgaySinh").ToString().Trim() <> "") Then
                            dgv_gtgc.Rows(i).Cells("cln_NgaySinh").Value = Format(CType(db.Rows(i)("NgaySinh"), DateTime), "dd/MM/yyyy")
                        Else
                            dgv_gtgc.Rows(i).Cells("cln_NgaySinh").Value = ""
                        End If
                        dgv_gtgc.Rows(i).Cells("cln_MaSoThueNPT").Value = db.Rows(i)("MaSoThueNPT").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_SoCMT").Value = db.Rows(i)("SoCMT").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_DiaChi").Value = db.Rows(i)("DiaChi").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_QuocTich_HT").Value = db.Rows(i)("QuocTich_HT").ToString().Trim()
                        dgv_gtgc.Rows(i).Cells("cln_GhiChu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                        Dim _Time As String = ""
                        If (db.Rows(i)("TuNgay").ToString().Trim() <> "") Then
                            _Time = Format(CType(db.Rows(i)("TuNgay"), DateTime), "dd/MM/yyyy")
                        Else
                            _Time = ""
                        End If
                        If (db.Rows(i)("DenNgay").ToString().Trim() <> "") Then
                            If CType(db.Rows(i)("DenNgay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900" Then
                                _Time = _Time + " - " + DateTimeUtil.DateTimeToString(Globals.GetDateTime_ForServerDB, "dd/MM/yyyy")

                            Else
                                _Time = _Time + " - " + DateTimeUtil.DateTimeToString(CType(db.Rows(i)("DenNgay"), DateTime), "dd/MM/yyyy")
                            End If
                        Else
                            _Time = "Từ " + _Time
                        End If
                        dgv_gtgc.Rows(i).Cells("cln_ThoiGian").Value = _Time.Trim()
                        dgv_gtgc.Rows(i).Cells("cln_TrangThai_HT").Value = db.Rows(i)("TrangThai_HT").ToString().Trim()


                    Next
                End If
            End If
        End Using

        'Fill data - Hồ sơ thong tin thu nhap gia dinh
        grid_TNGD.Rows.Clear()
        Using db As DataTable = _PersonnelFile.GetAll(_IdCanBo, 3)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        grid_TNGD.Rows.Add()
                        grid_TNGD.Rows(i).Cells("cln_Id").Value = db.Rows(i)("IdThuNhapGD").ToString().Trim()
                        grid_TNGD.Rows(i).Cells("cln_Luong").Value = db.Rows(i)("Luong").ToString().Trim()
                        Dim _Time As String = ""
                        If (db.Rows(i)("NgayKekhai").ToString().Trim() <> "") Then
                            _Time = Format(CType(db.Rows(i)("NgayKekhai"), DateTime), "dd/MM/yyyy")
                        Else
                            _Time = ""
                        End If
                        grid_TNGD.Rows(i).Cells("cln_ThoiGian").Value = _Time.Trim()
                    Next
                End If
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Hàm trả về tên của tab - Đang thao tác
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetName() As String
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"
                Return "hồ sơ Gia đình cán bộ"
            Case "tp_suckhoe"
                Return "hồ sơ sức khoẻ gia đình cán bộ"
            Case "tp_gtgc"
                Return "hồ sơ giảm trừ gia cảnh"
            Case Else
                Return ""
        End Select
    End Function

    ''' <summary>
    ''' Hàm thực hiện Kiểm tra xem đã có dữ liệu Thành viên nào của gia đình cán bộ chưa
    ''' </summary>
    ''' <returns>True: Đã có. False: Chưa có dl</returns>
    ''' <remarks></remarks>
    Private Function CheckAdd() As Boolean
        Using db As DataTable = _PersonnelFile.GetAll(_IdCanBo, 0)
            If (db Is Nothing) Then
                MessageBox.Show("Hiện chưa có dữ liệu về thành viên của gia đình cán bộ, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Return False
            End If
        End Using
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện kiểm tra sự hợp lệ về mặt dữ liệu - Khi cập nhật các hồ sơ
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"          'Hồ sơ gia đình cán bộ
                If (edt_gdcb_hoten.Text.Trim() = "") Then
                    MessageBox.Show("Họ và tên thành viên trong gia đình cán bộ không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_gdcb_hoten
                    Return False
                End If
                If (cb_gdcb_gioitinh.SelectedIndex <= 0) Then
                    MessageBox.Show("Bạn chưa chọn giới tính của thành viên trong gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gdcb_gioitinh
                    Return False
                End If
                If (cb_gdcb_quanhe.SelectedIndex <= 0 And cb_gdcb_quanhe.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn quan hệ của thành viên với cán bộ trong gia đình!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gdcb_quanhe
                    Return False
                End If
                If (cb_gdcb_quoctich.SelectedIndex <= 0 And cb_gdcb_quoctich.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn quốc tịch của thành viên trong gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gdcb_quoctich
                    Return False
                End If
                If (cb_gdcb_quequan.SelectedIndex <= 0 And cb_gdcb_quequan.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn quê quán (tỉnh - thành phố) của thành viên trong gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gdcb_quequan
                    Return False
                End If
                If (cb_gdcb_xaphuong.SelectedIndex <= 0 And cb_gdcb_xaphuong.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn quê quán (quận - huyện) của thành viên trong gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gdcb_xaphuong
                    Return False
                End If

                If (cb_gdcb_tinhtrang.SelectedIndex <= 0) Then
                    MessageBox.Show("Bạn chưa chọn tình trạng của thành viên trong gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gdcb_tinhtrang
                    Return False
                End If
                'Case "tp_suckhoe"       'Hồ sơ Sức khoẻ gia đình cán bộ
                '    If (cb_sk_thanhvien.SelectedIndex <= 0 And cb_sk_thanhvien.Items.Count <> 0) Then
                '        MessageBox.Show("Bạn chưa chọn thành viên trong gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = cb_sk_thanhvien
                '        Return False
                '    End If
                '    If dtpk_sk_namkham.Value > DateTime.Now Then
                '        MessageBox.Show("Năm khám sức khoẻ không thể lớn hơn năm hiện thời!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = dtpk_sk_namkham
                '        Return False
                '    End If
                '    If (edt_sk_dotkham.Value <= 0) Then
                '        MessageBox.Show("Bạn chưa chọn đợt khám sức khoẻ trong năm!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = edt_sk_dotkham
                '        Return False
                '    End If
                '    If dtpk_sk_tungay.Value > dtpk_sk_denngay.Value Then
                '        MessageBox.Show("Ngày bắt đầu không thể lớn hơn ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = dtpk_sk_tungay
                '        Return False
                '    End If
                '    If (dtpk_sk_namkham.Value.Year <> dtpk_sk_tungay.Value.Year Or dtpk_sk_namkham.Value.Year <> dtpk_sk_denngay.Value.Year) Then
                '        MessageBox.Show("Năm khám không phù hợp với khoảng thời gian trong hồ sơ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = dtpk_sk_namkham
                '        Return False
                '    End If

                '    If (cb_sk_loaisk.SelectedIndex <= 0) Then
                '        MessageBox.Show("Bạn chưa chọn loại sức khoẻ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = cb_sk_loaisk
                '        Return False
                '    End If
                '    If (edt_sk_ketluan.Text.Trim() = "") Then
                '        MessageBox.Show("Kết luận đợt khám sức khoẻ không thể để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '        ActiveControl = edt_sk_ketluan
                '        Return False
                '    End If
                '    If (edt_sk_cannang.Text.Trim() <> "") Then
                '        If (CType(edt_sk_cannang.Text.Trim().Replace(" ", ""), Double) > 300) Then
                '            MessageBox.Show(String.Format("Giá trị cân nặng trong hồ sơ khám sức khoẻ không hợp lệ!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '            ActiveControl = edt_sk_cannang
                '            Return False
                '        End If
                '    End If
                '    If (edt_sk_chieucao.Text.Trim() <> "") Then
                '        If (CType(edt_sk_chieucao.Text.Trim().Replace(" ", ""), Double) > 250) Then
                '            MessageBox.Show(String.Format("Giá trị chiều cao trong hồ sơ khám sức khoẻ không hợp lệ!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '            ActiveControl = edt_sk_chieucao
                '            Return False
                '        End If
                '    End If
                '    'Bắt trùng đợi khám trong cùng một năm - của cán bộ
                '    If (edt_sk_dotkham.Value > 0) Then
                '        If (_RecordId <> "") Then
                '            strSQL = String.Format("Select * from HS_SKhGDCB Where Dot = {0} And Nam = {1} And IdTVien = '{2}' And IdKhamchuaGDCB <> '{3}'", CType(edt_sk_dotkham.Value, Int16), dtpk_sk_namkham.Value.Year, arrThVien(cb_sk_thanhvien.SelectedIndex), _RecordId)
                '        Else
                '            strSQL = String.Format("Select * from HS_SKhGDCB Where Dot = {0} And Nam = {1} And IdTVien = '{2}'", CType(edt_sk_dotkham.Value, Int16), dtpk_sk_namkham.Value.Year, arrThVien(cb_sk_thanhvien.SelectedIndex))
                '        End If
                '        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '            If Not (db Is Nothing) Then
                '                If (db.Rows.Count > 0) Then
                '                    MessageBox.Show(String.Format("Đợt khám sức khoẻ trong năm đã tồn tại. Bạn vui lòng kiểm tra lại!"), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '                    ActiveControl = edt_sk_dotkham
                '                    Return False
                '                End If
                '            End If
                '        End Using
                '    End If
                '    'Bắt điều kiện nhập dữ liệu khoảng thời gian không cho trùng hoặc đan xen lẫn nhau
                '    If (_RecordId = "") Then
                '        strSQL = String.Format("Select * From HS_SKhGDCB Where IdTVien = '{0}' Order by TuNgay Asc", arrThVien(cb_sk_thanhvien.SelectedIndex))
                '    Else
                '        strSQL = String.Format("Select * From HS_SKhGDCB Where IdTVien = '{0}' And IdKhamchuaGDCB <> '{1}' Order by TuNgay Asc", arrThVien(cb_sk_thanhvien.SelectedIndex), _RecordId)
                '    End If
                '    Dim Exist As Boolean = False
                '    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '        If Not (db Is Nothing) Then
                '            If (db.Rows.Count > 0) Then
                '                For i As Integer = 0 To db.Rows.Count - 1
                '                    Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                '                    Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                '                    If ((_TuNgay <= dtpk_sk_tungay.Value And dtpk_sk_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_sk_denngay.Value And dtpk_sk_denngay.Value <= _DenNgay) Or (dtpk_sk_tungay.Value <= _TuNgay And dtpk_sk_denngay.Value >= _DenNgay)) Then
                '                        Exist = True
                '                        Exit For
                '                    End If
                '                Next
                '                If (Exist = True) Then
                '                    MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ khám sức khoẻ không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '                    ActiveControl = dtpk_sk_tungay
                '                    Return False
                '                End If
                '            End If
                '        End If
                '    End Using

            Case "tp_gtgc"
                If (edt_gtgc_hoten.Text.Trim() = "") Then
                    MessageBox.Show("Họ tên của người phụ thuộc không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_gtgc_hoten
                    Return False
                End If
                'If (edt_gtgc_masothue.Text.Trim() = "") Then
                '    MessageBox.Show("Mã số thuế của người phụ thuộc không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_gtgc_masothue
                '    Return False
                'End If
                If (cb_gtgc_quoctich.SelectedIndex <= 0 And cb_gtgc_quoctich.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa Quốc tích của người phụ thuộc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gtgc_quanhe
                    Return False
                End If
                'If (edt_gtgc_socmt.Text.Trim() = "") Then
                '    MessageBox.Show("Số CMT/Thẻ căn cước/Số giấy khai sinh của người phụ thuộc không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_gtgc_socmt
                '    Return False
                'End If
                If (cb_gtgc_quanhe.SelectedIndex <= 0 And cb_gtgc_quanhe.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn mối quan hệ của người phụ thuộc với cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_gtgc_quanhe
                    Return False
                End If
                If (dtpk_gtgc_denngay.Checked = True) Then
                    If dtpk_gtgc_tungay.Value >= dtpk_gtgc_denngay.Value Then
                        MessageBox.Show("Ngày bắt đầu không thể lớn hơn ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_gtgc_tungay
                        Return False
                    End If
                End If
                'Kiểm tra không cho phép khoảng thời gian trùng hoặc sen kẽ nhau với Giảm trừ gia cảnh cùng 1 người
                If (_RecordId = "") Then
                    strSQL = String.Format("SELECT * FROM HS_GTGC WHERE IdCanBo = '{0}' And HoTenNguoiPT LIKE N'{1}' And IdQuanHe = {2} Order By TuNgay Asc", _IdCanBo, Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_gtgc_hoten.Text.ToString().Trim())), CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer))
                Else
                    strSQL = String.Format("SELECT * FROM HS_GTGC WHERE IdCanBo = '{0}' And HoTenNguoiPT LIKE N'{1}' And IdQuanHe = {2} And IdGTGC <> '{3}' Order By TuNgay Asc", _IdCanBo, Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_gtgc_hoten.Text.ToString().Trim())), CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer), _RecordId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = DateTime.Now
                                If (db.Rows(i)("DenNgay").ToString() <> "") Then
                                    If CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900" Then
                                        _DenNgay = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                    End If
                                End If
                                If (dtpk_gtgc_denngay.Checked = True) Then
                                    If ((_TuNgay <= dtpk_gtgc_tungay.Value And dtpk_gtgc_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_gtgc_denngay.Value And dtpk_gtgc_denngay.Value <= _DenNgay) Or (dtpk_gtgc_tungay.Value <= _TuNgay And dtpk_gtgc_denngay.Value >= _DenNgay)) Then
                                        Exist = True
                                        Exit For
                                    End If
                                Else
                                    If ((_TuNgay <= dtpk_gtgc_tungay.Value And dtpk_gtgc_tungay.Value <= _DenNgay) Or (dtpk_gtgc_tungay.Value <= _TuNgay)) Then
                                        Exist = True
                                        Exit For
                                    End If
                                End If
                                If (db.Rows(i)("DenNgay").ToString() = "") Then
                                    If (dtpk_gtgc_tungay.Value > DateTime.Now) Then
                                        Exist = True
                                    End If
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc trong hồ sơ giảm trừ gia cảnh không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau trong cùng một hồ sơ giảm trừ gia cảnh cho cùng một người.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_gtgc_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using
            Case "tp_TNGD"
                If (txtLuong.Text.Trim() = "") Then
                    MessageBox.Show("Thông tin thu nhập theo lương không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = txtLuong
                    Return False
                End If
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện kiểm tra hợp lệ về mặt dữ liệu giữa Hồ sơ giảm trừ gia cảnh so với Hồ sơ gia đình cán bộ
    ''' Xét 4 trường hợp sau phải có tính thống nhất về mặt dữ liệu. Ví dụ họ tên phải tương đương. 
    ''' Nếu đã mất rùi thì phải thông báo
    '''                               2301: Ông nội
    '''                               2302: Bà nội
    '''                               2303: Bố đẻ
    '''                               2304: Mẹ đẻ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckRelations() As Boolean
        Dim _Name As String = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_gtgc_hoten.Text.Trim().ToString()))
        Dim _Code As String = ""
        strSQL = String.Format("Select * from DanhMuc Where id_goc = 23 and Status = 1 and id = '{0}'", CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer))
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _Code = db.Rows(0)("ma_so").Trim().ToString()
                End If
            End If
        End Using

        'Xét nếu quan hệ rơi vào 4 trường hợp kể trên thì thực hiện kiểm tra
        If (_Code = "2301" Or _Code = "2302" Or _Code = "2303" Or _Code = "2304") Then
            'Xét trường hợp trong hồ sơ của cán bộ kê khai giảm trừ gia cảnh - đã có trong cùng một khoảng thời gian 4 trường hợp trên rồi
            If (_RecordId <> "") Then
                strSQL = String.Format("Select * from HS_GTGC Where IdCanBo = '{0}' and IdQuanHe = {1} and TuNgay = '{2}' and DenNgay = '{3}' and IdGTGC <> '{4}'", _IdCanBo, CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer), Format(dtpk_gtgc_tungay.Value, "MM/dd/yyyy"), Format(dtpk_gtgc_denngay.Value, "MM/dd/yyyy"), _RecordId)
            Else
                strSQL = String.Format("Select * from HS_GTGC Where IdCanBo = '{0}' and IdQuanHe = {1} and TuNgay = '{2}' and DenNgay = '{3}'", _IdCanBo, CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer), Format(dtpk_gtgc_tungay.Value, "MM/dd/yyyy"), Format(dtpk_gtgc_denngay.Value, "MM/dd/yyyy"))
            End If

            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        MessageBox.Show(cb_gtgc_quanhe.SelectedItem.ToString() & " của cán bộ mà bạn kê khai giảm trừ gia cảnh trong khoảng thời gian này đã tồn tại!" + vbCrLf + "Lưu ý: với những trường hợp (Ông nội, Bà nội, Bố đẻ, Mẹ đẻ) thì dữ liệu trong cùng một khoảng thời gian là duy nhất với mỗi cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = cb_gtgc_quanhe
                        Return False
                    End If
                End If
            End Using

            'Tìm DL trong bảng HS_GDCB để so sánh
            strSQL = String.Format("Select * from HS_GDCB Where IdCanBo = '{0}' and IdQuanHe = {1}", _IdCanBo, CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer))
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        If (db.Rows(0)("ConMat").ToString().Trim() = True) Then
                            If (dtpk_gtgc_denngay.Value.Year > CType(db.Rows(0)("NamMat").ToString().Trim(), Integer)) Then
                                MessageBox.Show("Người thân bạn kê khai giảm trừ gia cảnh của cán bộ trong khoảng thời gian này đã mất." + vbCrLf + "Lưu ý: với những trường hợp (Ông nội, Bà nội, Bố đẻ, Mẹ đẻ) thì dữ liệu phải tương thích với hồ sơ gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = edt_gtgc_hoten
                                Return False
                            End If
                        End If

                        Dim _NameTag As String = db.Rows(0)("HoTen").Trim().ToString()
                        If (_NameTag <> _Name) Then
                            MessageBox.Show("Dữ liệu giảm trừ gia cảnh bạn khai báo không phù hợp với hồ sơ gia đình cán bộ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                            ActiveControl = edt_gtgc_hoten
                            Return False
                        End If
                    End If
                End If
            End Using
        End If
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện kiểm tra quyền thành viên được phép thao tác với các chức năng
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";167;168;169;170;", _roles)) Then       ' Hồ sơ - Thành viên gia đình cán bộ
            If tctrl_main.TabPages.Contains(tp_gdcb) Then tctrl_main.TabPages.Remove(tp_gdcb)
        End If
        'If Not (Globals.IsIntersect(";171;172;173;174;", _roles)) Then       ' Hồ sơ - Sức khoẻ của Thành viên gia đình cán bộ
        '    If tctrl_main.TabPages.Contains(tp_suckhoe) Then tctrl_main.TabPages.Remove(tp_suckhoe)
        'End If
        If Not (Globals.IsIntersect(";175;176;177;178;", _roles)) Then       ' Hồ sơ - Sức khoẻ của Thành viên gia đình cán bộ
            If tctrl_main.TabPages.Contains(tp_gtgc) Then tctrl_main.TabPages.Remove(tp_gtgc)
            If tctrl_main.TabPages.Contains(tp_TNGD) Then tctrl_main.TabPages.Remove(tp_TNGD)
        End If
        'Nếu không có quyền Xem thông tin với cả 3 hồ sơ
        If (Globals.Roles.IndexOf(";167;") < 0 And Globals.Roles.IndexOf(";171;") < 0 And Globals.Roles.IndexOf(";175;") < 0) Then
            tv_main.Enabled = False
        End If
        'Nếu không có quyền thêm mới với cả 3 hồ sơ
        If (Globals.Roles.IndexOf(";168;") < 0 And Globals.Roles.IndexOf(";172;") < 0 And Globals.Roles.IndexOf(";176;") < 0) Then
            btn_themdl.Enabled = False
        End If
        'Nếu không có quyền xoá bỏ với cả 3 hồ sơ
        If (Globals.Roles.IndexOf("170;") < 0 And Globals.Roles.IndexOf("174;") < 0 And Globals.Roles.IndexOf(";178;") < 0) Then
            btn_xoadl.Enabled = False
        End If
        'Nếu không có quyền thêm mới và sửa đổi với cả 3 hồ sơ
        If (Globals.Roles.IndexOf(";168;") < 0 And Globals.Roles.IndexOf(";169;") < 0 And Globals.Roles.IndexOf(";172;") < 0 And Globals.Roles.IndexOf(";173;") < 0 And Globals.Roles.IndexOf(";176;") < 0 And Globals.Roles.IndexOf(";177;") < 0) Then
            btn_ghidl.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện select row trên lưới dữ liệu --> Fill dữ liệu vào các controls
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi (Id record) đang select</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String)
        Dim dr As DataRow
        _RecordId = _Idnew
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"          'Hồ sơ Thông tin Gia đình cán bộ
                dr = _PersonnelFile.GetRecord(_RecordId, 0)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        Dim currentDate As DateTime = New DateTime()
                        edt_gdcb_hoten.Text = dr("HoTen").ToString().Trim()
                        cb_gdcb_gioitinh.SelectedIndex = IIf(dr("GioiTinh").ToString().Trim() = False, 1, 2)

                        currentDate = New DateTime(CType(dr("NamSinh").ToString().Trim(), Integer), 1, 1)
                        dtpk_gdcb_namsinh.Value = currentDate
                        cb_gdcb_quanhe.SelectedIndex = IIf(dr("IdQuanHe").ToString() <> "", CType(arrQuanHe.IndexOf(dr("IdQuanHe").ToString()), Integer), 0)
                        cb_gdcb_quoctich.SelectedIndex = IIf(dr("IdQuocGia").ToString() <> "", CType(arrQuocTich.IndexOf(dr("IdQuocGia").ToString()), Integer), 0)

                        Dim _IdXaPhuong As Integer = IIf(dr("IdQueQuan").ToString() <> "", CType(dr("IdQueQuan").ToString(), Integer), 0)
                        Dim _IdTTP As Integer = 0
                        If (_IdXaPhuong <= 0) Then
                            cb_gdcb_quequan.SelectedIndex = 0
                            cb_gdcb_quequan_SelectedIndexChanged(Nothing, Nothing)
                        Else
                            strSQL = String.Format("Select A.* From Dm_DiaPhuong A Where A.Ma_Xa='00' And A.Ma_Thon='00' And A.Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By X.TrangThai) Order By A.TrangThai", _IdXaPhuong)
                            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        _IdTTP = CType(db.Rows(0)("Id").ToString(), Integer)
                                    End If
                                End If
                            End Using
                        End If
                        cb_gdcb_quequan.SelectedIndex = IIf(_IdTTP > 0, CType(arrQueQuan.IndexOf(_IdTTP.ToString()), Integer), 0)
                        cb_gdcb_quequan_SelectedIndexChanged(Nothing, Nothing)
                        If (dr("IdQueQuan").ToString() <> "") Then
                            cb_gdcb_xaphuong.SelectedIndex = CType(arrXaPhuong.IndexOf(dr("IdQueQuan").ToString()), Integer)
                        End If
                        edt_gdcb_nghenghiep.Text = dr("NgheNghiep").ToString().Trim()
                        edt_gdcb_dienthoai.Text = dr("DienThoai").ToString().Trim()
                        edt_gdcb_diachi.Text = dr("DiaChi").ToString().Trim()
                        cb_gdcb_utbanthan.SelectedIndex = IIf(dr("IdUT_BThan").ToString() <> "", CType(arrUtBanThan.IndexOf(dr("IdUT_BThan").ToString()), Integer), 0)
                        If (dr("ConMat").ToString().Trim() = "") Then
                            cb_gdcb_tinhtrang.SelectedIndex = 0
                        Else
                            cb_gdcb_tinhtrang.SelectedIndex = IIf(dr("ConMat").ToString().Trim() = False, 1, 2)
                        End If
                        cb_gdcb_tinhtrang_SelectedIndexChanged(Nothing, Nothing)

                        If (dr("ConMat").ToString().Trim() = True) Then
                            currentDate = New DateTime(CType(dr("NamMat").ToString().Trim(), Int32), 1, 1)
                            dtpk_gdcb_nammat.Value = currentDate
                            edt_gdcb_lydo.Text = dr("LyDo_Mat").ToString().Trim()
                        End If
                    End If
                End If
            Case "tp_gtgc"          'Hồ sơ Giảm trừ gia cảnh của cán bộ
                dr = _PersonnelFile.GetRecord(_RecordId, 2)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        edt_gtgc_hoten.Text = dr("HoTenNguoiPT").ToString().Trim()
                        cb_gtgc_quanhe.SelectedIndex = IIf(dr("IdQuanHe").ToString() <> "", CType(arrQuanHe.IndexOf(dr("IdQuanHe").ToString()), Integer), 0)
                        cb_gtgc_quoctich.SelectedIndex = IIf(dr("IdQuocTich").ToString() <> "", CType(arrQuocTichGTGC.IndexOf(dr("IdQuocTich").ToString()), Integer), 0)

                        If dr("TuNgay").ToString().Trim() <> "" Then
                            dtpk_gtgc_tungay.Value = CType(dr("TuNgay").ToString(), DateTime)
                        End If
                        If (dr("DenNgay").ToString().Trim() <> "") Then
                            If (CType(dr("DenNgay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_gtgc_denngay.Checked = False
                            Else
                                dtpk_gtgc_denngay.Checked = True
                                dtpk_gtgc_denngay.Value = CType(dr("DenNgay").ToString().Trim(), DateTime)
                            End If
                        Else
                            dtpk_gtgc_denngay.Checked = False
                        End If
                        edt_gtgc_ghichu.Text = dr("GhiChu").ToString().Trim()
                        If (dr("TrangThai").ToString() <> "") Then
                            If CType(dr("TrangThai").ToString(), Byte) = 1 Then
                                rb_mo.Checked = True
                            Else
                                rb_dong.Checked = True
                            End If
                        End If
                        edt_gtgc_masothue.Text = dr("MaSoThueNPT").ToString().Trim()
                        edt_gtgc_diachi.Text = dr("DiaChi").ToString().Trim()
                        edt_gtgc_dienthoai.Text = dr("DienThoai").ToString().Trim()
                        edt_gtgc_socmt.Text = dr("SoCMT").ToString().Trim()
                        If dr("NgayCap").ToString().Trim() <> "" Then
                            dtpk_gtgc_ngaycap.Value = CType(dr("NgayCap").ToString(), DateTime)
                        End If
                        If dr("NgaySinh").ToString().Trim() <> "" Then
                            dtpk_gtgc_ngaysinh.Value = CType(dr("NgaySinh").ToString(), DateTime)
                        End If
                        edt_gtgc_noicap.Text = dr("NoiCap").ToString().Trim()
                    End If
                End If
            Case "tp_TNGD"          'Hồ sơ thu nhap gia dinh can bo
                dr = _PersonnelFile.GetRecord(_RecordId, 3)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        If dr("NgayKekhai").ToString().Trim() <> "" Then
                            dpkThoiGianKK.Value = CType(dr("NgayKekhai").ToString(), DateTime)
                        End If
                        txtLuong.Text = dr("Luong").ToString().Trim()
                        txtNguonKhac.Text = dr("NguonKhac").ToString().Trim()
                        txtNhaCap.Text = dr("NhaO_DuocCap").ToString().Trim()
                        txtDT_Nha_Cap.Text = IIf(dr("NhaO_DuocCap_DT").ToString().Trim() <> "", CType(dr("NhaO_DuocCap_DT"), Double).ToString("N2"), "0")
                        txtNhaMua.Text = dr("NhaO_TuMua").ToString().Trim()
                        txtDT_Nha_Mua.Text = IIf(dr("NhaO_TuMua_DT").ToString().Trim() <> "", CType(dr("NhaO_TuMua_DT"), Double).ToString("N2"), "0")
                        txtDT_Dat_Cap.Text = IIf(dr("DatO_DuocCap_DT").ToString().Trim() <> "", CType(dr("DatO_DuocCap_DT"), Double).ToString("N2"), "0")
                        txtDT_Dat_Mua.Text = IIf(dr("DatO_TuMua_DT").ToString().Trim() <> "", CType(dr("DatO_TuMua_DT"), Double).ToString("N2"), "0")
                        txtDat_XSKD.Text = dr("DatSXKT").ToString().Trim()
                    End If
                End If
                'Case "tp_suckhoe"       'Hồ sơ Thông tin Khám sức khoẻ của Gia đình cán bộ
                '    dr = _PersonnelFile.GetRecord(_RecordId, 1)
                '    If Not (dr Is Nothing) Then
                '        If (dr.Table.Rows.Count > 0) Then
                '            cb_sk_thanhvien.SelectedIndex = IIf(dr("IdTVien").ToString() <> "", CType(arrThVien.IndexOf(dr("IdTVien").ToString()), Integer), 0)
                '            Dim currentDate As DateTime = New DateTime()
                '            currentDate = New DateTime(CType(dr("Nam").ToString().Trim(), Int32), 1, 1)
                '            dtpk_sk_namkham.Value = currentDate
                '            edt_sk_dotkham.Value = CType(dr("Dot").ToString().Trim(), Byte)

                '            edt_sk_ndkham.Text = dr("NoiDungKC").ToString().Trim()
                '            edt_sk_noikham.Text = dr("NoiKham").ToString().Trim()

                '            If dr("TuNgay").ToString().Trim() <> "" Then
                '                dtpk_sk_tungay.Value = CType(dr("TuNgay").ToString(), DateTime)
                '            End If
                '            If dr("DenNgay").ToString().Trim() <> "" Then
                '                dtpk_sk_denngay.Value = CType(dr("DenNgay").ToString(), DateTime)
                '            End If
                '            edt_sk_cannang.Text = IIf(dr("CanNang").ToString().Trim() <> "", CType(dr("CanNang"), Double).ToString("N2"), "0")
                '            edt_sk_chieucao.Text = IIf(dr("ChieuCao").ToString().Trim() <> "", CType(dr("ChieuCao"), Double).ToString("N2"), "0")
                '            cb_sk_loaisk.SelectedIndex = CType(dr("LoaiSuckhoe").ToString(), Byte)
                '            edt_sk_ketluan.Text = dr("KetLuan").ToString().Trim()
                '            edt_sk_ghichu.Text = dr("GhiChu").ToString().Trim()
                '        End If
                '    End If
        End Select
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHS_GDCB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Fill data ra cây dữ liệu Treeview - Cơ cấu tổ chức
        _PersonnelFile.Fill_Tree(tv_main)
        'Gọi - Hàm định nghĩa lưới dữ liệu
        _PersonnelFile.Create_Frame(dgv_gdcb, 14)
        '_PersonnelFile.Create_Frame(dgv_suckhoe, 15)
        _PersonnelFile.Create_Frame(dgv_gtgc, 16)
        _PersonnelFile.Create_Frame(grid_TNGD, 17)

        'Thực hiện fill dữ liệu vào ComboBox trên giao diện
        arrQuanHe.Clear()
        cb_gdcb_quanhe.Items.Clear()
        arrQuanHe = _Globals.Bind_ComBoBox(cb_gdcb_quanhe, clsHT_DanhMuc.Sql_QuanHe, "---Quan hệ với cán bộ---")

        cb_gtgc_quanhe.Items.Clear()
        _Globals.Bind_ComBoBox(cb_gtgc_quanhe, clsHT_DanhMuc.Sql_QuanHe, "---Mối quan hệ---")

        arrQueQuan.Clear()
        cb_gdcb_quequan.Items.Clear()
        arrQueQuan = _Globals.Bind_ComBoBox(cb_gdcb_quequan, clsHT_DanhMuc.Sql_TinhTP, "---Tỉnh - Thành phố---")

        arrQuocTich.Clear()
        cb_gdcb_quoctich.Items.Clear()
        arrQuocTich = _Globals.Bind_ComBoBox(cb_gdcb_quoctich, clsHT_DanhMuc.Sql_Quocgia, "---Quốc tịch---")


        arrQuocTichGTGC.Clear()
        cb_gtgc_quoctich.Items.Clear()
        arrQuocTichGTGC = _Globals.Bind_ComBoBox(cb_gtgc_quoctich, clsHT_DanhMuc.Sql_Quocgia, "---Quốc tịch---")

        arrUtBanThan.Clear()
        cb_gdcb_utbanthan.Items.Clear()
        arrUtBanThan = _Globals.Bind_ComBoBox(cb_gdcb_utbanthan, clsHT_DanhMuc.Sql_Ut_banthan, "---Ưu tiên bản thân---")

        arrThVien.Clear()
        ResetControls(0)
        lkl_xemct.Enabled = False
        'Hàm kiểm tra quyền thành viên Thao tác Hồ sơ Gia đình cán bộ
        Check_Permits()

        'Nếu được gọi từ Hồ sơ cán bộ
        If (FlagShow = True) Then
            Dim _DonviId As Integer = 0         'Id đơn vị của cán bộ
            Dim _CodeBranch As String = ""      'Mã hiệu của chi nhánh cap tinh hoac tuong duong
            Dim _PhongBanId As Integer = 0      'Id phòng ban của cán bộ
            If (TagNode <> "") Then
                Dim arrElement() As String = TagNode.Split("_")
                If (arrElement.Length > 0) Then
                    'If TagNode.LastIndexOf("00") < 0 Then   'Trường hợp đơn vị là PGD
                    Dim iId_Goc As Integer = 0
                    iId_Goc = SoftSqlHelper.GetNumber("Select Id_Goc From ChiNhanh Where Ma_So='" + arrElement(2) + "'", 0)
                    'If TagNode.Substring(TagNode.ToString.Length - 2, 2) <> "00" Then   'Trường hợp đơn vị là PGD
                    If iId_Goc <> 0 And iId_Goc <> 1 Then   'Trường hợp đơn vị là PGD
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
        End If

        If HumanId <> "" Then
            If checkRight_CreateRecord(HumanId) OrElse (IdDONVI = 1 AndAlso chkTWQuanLy.Checked = True) Then
                btn_ghidl.Enabled = True
                btn_xoadl.Enabled = True
                chkTWQuanLy.Checked = True
            Else
                btn_ghidl.Enabled = False
                btn_xoadl.Enabled = False
                chkTWQuanLy.Checked = False
            End If
        End If

    End Sub

    Private Sub frmHS_GDCB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub cb_gdcb_tinhtrang_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_gdcb_tinhtrang.SelectedIndexChanged
        If (cb_gdcb_tinhtrang.SelectedIndex >= 0 And cb_gdcb_tinhtrang.Items.Count <> 0) Then
            If (cb_gdcb_tinhtrang.SelectedIndex = 2) Then   'Còn sống
                lbl_lydo_mat.Visible = True
                lbl_nammat.Visible = True
                dtpk_gdcb_nammat.Visible = True
                edt_gdcb_lydo.Visible = True
                cb_gdcb_tinhtrang.Width = 111
            Else
                lbl_lydo_mat.Visible = False
                lbl_nammat.Visible = False
                dtpk_gdcb_nammat.Visible = False
                edt_gdcb_lydo.Visible = False
                cb_gdcb_tinhtrang.Width = 182
                dtpk_gdcb_nammat.Text = DateTime.Now.ToShortDateString()
                edt_gdcb_lydo.Text = ""
            End If
        End If
    End Sub

    Private Sub cb_gdcb_quequan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cb_gdcb_quequan.SelectedIndexChanged
        cb_gdcb_xaphuong.Items.Clear()
        arrXaPhuong.Clear()
        If (cb_gdcb_quequan.SelectedIndex > 0 And cb_gdcb_quequan.Items.Count <> 0) Then
            Dim _Id_TTP As Integer = CType(arrQueQuan(IIf(cb_gdcb_quequan.SelectedIndex > 0, cb_gdcb_quequan.SelectedIndex, "0")), Integer)
            If _Id_TTP > 0 Then
                Dim strSQL As String = String.Format("Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa <> '00' And Ma_Thon = '00' And TrangThai = 'A' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Id={0}) Order by Ma_Tinh,Ma_Xa,Ma_Thon Asc", _Id_TTP)
                arrXaPhuong = _Globals.Bind_ComBoBox(cb_gdcb_xaphuong, strSQL, "--- Xã/Phường ---")
            End If
        End If
    End Sub

    Private Sub lkl_xemct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lkl_xemct.LinkClicked
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
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        Try
            'arrThVien.Clear()
            'cb_sk_thanhvien.Items.Clear()

            _NodeCurrent = ""
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
                _NodeCurrent = tv_main.SelectedNode.Tag.ToString().Trim()
                ResetControls(0)
                'Lấy Chỉ số xác định đơn vị Hiện tại của cán bộ (Xét với trường hợp cài đặt đơn vị '000100')
                If (DONVI = gMaDonViTW) Then
                    If (_NodeCurrent.IndexOf("DV_") >= 0) Then
                        _IdDonviHT = CType(_NodeCurrent.Substring(3, 1), Integer)
                    End If
                End If
                'Fill data các hồ sơ liên quan
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "CB") Then
                    _IdCanBo = tv_main.SelectedNode.Tag.ToString().Trim().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1)
                    If (_IdCanBo <> "") Then
                        'Fill lại combobox - Danh sách thành viên trong gia đình cán bộ
                        'arrThVien.Clear()
                        'cb_sk_thanhvien.Items.Clear()
                        'arrThVien = _Globals.Bind_ComBoBox(cb_sk_thanhvien, String.Format("Select IdTVien,HoTen From HS_GDCB Where IdCanBo = '{0}'", _IdCanBo), "---Thành viên trong gia đình---")
                        'cb_sk_thanhvien.SelectedIndex = 0
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
                                    btn_ghidl.Visible = False
                                    btn_themdl.Visible = False
                                Else
                                    btn_ghidl.Visible = True
                                    btn_themdl.Visible = True
                                End If
                                chkTWQuanLy.Checked = My_CBool(dr("CapQuanLy"), False)
                            End If
                        End If
                        FillAll_Documents(_IdCanBo)

                    End If
                End If
                tctrl_main_SelectedIndexChanged(sender, Nothing)
                tv_main.SelectedNode.Expand()
                If _IdCanBo <> "" Then
                    If checkRight_CreateRecord(_IdCanBo) OrElse (IdDONVI = 1 AndAlso chkTWQuanLy.Checked = True) Then
                        btn_ghidl.Enabled = True
                        btn_xoadl.Enabled = True
                        chkTWQuanLy.Checked = True
                    Else
                        btn_ghidl.Enabled = False
                        btn_xoadl.Enabled = False
                        chkTWQuanLy.Checked = False
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Hiển thị hồ sơ khi click vào cây dữ liệu: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub tctrl_main_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tctrl_main.SelectedIndexChanged
        btn_themdl.Enabled = True
        btn_xoadl.Enabled = True
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"          'Hồ sơ Thông tin Gia đình cán bộ
                dgv_gdcb_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";167;") < 0) Then
                    dgv_gdcb.Enabled = False
                Else
                    dgv_gdcb.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";168;") < 0) Then
                    btn_themdl.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";170;") < 0) Then
                    btn_xoadl.Enabled = False
                End If
                'Case "tp_suckhoe"       'Hồ sơ Thông tin Khám sức khoẻ của Gia đình cán bộ
                '    dgv_suckhoe_CellClick(sender, Nothing)
                '    If (Globals.Roles.IndexOf(";171;") < 0) Then
                '        dgv_suckhoe.Enabled = False
                '    Else
                '        dgv_suckhoe.Enabled = True
                '    End If
                '    If (Globals.Roles.IndexOf(";172;") < 0) Then
                '        btn_themdl.Enabled = False
                '    End If
                '    If (Globals.Roles.IndexOf(";174;") < 0) Then
                '        btn_xoadl.Enabled = False
                '    End If
            Case "tp_gtgc"
                dgv_gtgc_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";175;") < 0) Then
                    dgv_gtgc.Enabled = False
                Else
                    dgv_gtgc.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";176;") < 0) Then
                    btn_themdl.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";178;") < 0) Then
                    btn_xoadl.Enabled = False
                End If
            Case "tp_TNGD"
                grid_TNGD_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";175;") < 0) Then
                    grid_TNGD.Enabled = False
                Else
                    grid_TNGD.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";176;") < 0) Then
                    btn_themdl.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";178;") < 0) Then
                    btn_xoadl.Enabled = False
                End If
        End Select
    End Sub

    Private Sub dgv_gdcb_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_gdcb.CellClick
        Try
            ResetControls(1)
            If (dgv_gdcb.Rows.Count > 0) Then
                If ((dgv_gdcb.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    SelectRow(dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ các thành viên trong gia đình cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Private Sub dgv_suckhoe_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    '    Try
    '        ResetControls(2)
    '        If (dgv_suckhoe.Rows.Count > 0) Then
    '            If ((dgv_suckhoe.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
    '                SelectRow(dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString())
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show("Chi tiết hồ sơ khám sức khoẻ của gia đình cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub

    Private Sub dgv_gtgc_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_gtgc.CellClick
        Try
            ResetControls(3)
            If (dgv_gtgc.Rows.Count > 0) Then
                If ((dgv_gtgc.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    SelectRow(dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Chi tiết hồ sơ giảm trừ gia cảnh: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub grid_TNGD_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grid_TNGD.CellClick
        Try
            ResetControls(4)
            If (grid_TNGD.Rows.Count > 0) Then
                If ((grid_TNGD.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    SelectRow(grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi trong quá trình load hồ sơ thu nhập gia đình: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_gdcb_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_gdcb.KeyUp
        dgv_gdcb_CellClick(sender, Nothing)
    End Sub

    'Private Sub dgv_suckhoe_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    dgv_suckhoe_CellClick(sender, Nothing)
    'End Sub

    Private Sub dgv_gtgc_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_gtgc.KeyUp
        dgv_gtgc_CellClick(sender, Nothing)
    End Sub

    Private Sub grid_TNGD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grid_TNGD.KeyUp
        grid_TNGD_CellClick(sender, Nothing)
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub btn_huybo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_huybo.Click
        Dim _currRow As String = _RecordId
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"          'Hồ sơ Thông tin các thành viên trong gia đình cán bộ
                ResetControls(1)
                If (dgv_gdcb.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString()
                dgv_gdcb.CurrentRow.Selected = False
                dgv_gdcb.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_gdcb, "cln_Id")).Selected = True
                If (dgv_gdcb.Rows.Count > 0) Then
                    If ((dgv_gdcb.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
                'Case "tp_suckhoe"       'Hồ sơ khám sức khoẻ của các thành viên trong gia đình cán bộ
                '    ResetControls(2)
                '    If (dgv_suckhoe.Rows.Count <= 0) Then Return
                '    If (_currRow = "") Then _currRow = dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString()
                '    dgv_suckhoe.CurrentRow.Selected = False
                '    dgv_suckhoe.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_suckhoe, "cln_Id")).Selected = True
                '    If (dgv_suckhoe.Rows.Count > 0) Then
                '        If ((dgv_suckhoe.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                '            SelectRow(_currRow)
                '        End If
                '    End If
            Case "tp_gtgc"
                ResetControls(3)
                If (dgv_gtgc.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString()
                dgv_gtgc.CurrentRow.Selected = False
                dgv_gtgc.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_gtgc, "cln_Id")).Selected = True
                If (dgv_gtgc.Rows.Count > 0) Then
                    If ((dgv_gtgc.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
            Case "tp_TNGD"
                ResetControls(4)
                If (grid_TNGD.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString()
                grid_TNGD.CurrentRow.Selected = False
                grid_TNGD.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, grid_TNGD, "cln_Id")).Selected = True
                If (grid_TNGD.Rows.Count > 0) Then
                    If ((grid_TNGD.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
        End Select
    End Sub

    Private Sub btn_themdl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_themdl.Click
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ cần thêm mới " + GetName() + "!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"
                ResetControls(1)
                'Thực hiện load một số controls cho người dùng đỡ phải nhập
                Dim dr As DataRow
                If (dgv_gdcb.Rows.Count > 0) Then
                    If ((dgv_gdcb.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                        dr = _PersonnelFile.GetRecord(dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString().Trim(), 0)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                cb_gdcb_quoctich.SelectedIndex = IIf(dr("IdQuocGia").ToString() <> "", CType(arrQuocTich.IndexOf(dr("IdQuocGia").ToString()), Integer), 0)
                                'Lấy lại thông tin địa chỉ của một thành viên trước đó
                                Dim _IdXaPhuong As Integer = IIf(dr("IdQueQuan").ToString() <> "", CType(dr("IdQueQuan").ToString(), Integer), 0)
                                Dim _IdTTP As Integer = 0
                                If (_IdXaPhuong <= 0) Then
                                    cb_gdcb_quequan.SelectedIndex = 0
                                    cb_gdcb_quequan_SelectedIndexChanged(sender, Nothing)
                                Else
                                    strSQL = String.Format("Select A.* From Dm_DiaPhuong A Where A.Ma_Xa='00' And A.Ma_Thon='00' And A.Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By X.TrangThai) Order By A.TrangThai", _IdXaPhuong)
                                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                        If Not (db Is Nothing) Then
                                            If (db.Rows.Count > 0) Then
                                                _IdTTP = CType(db.Rows(0)("Id").ToString(), Integer)
                                            End If
                                        End If
                                    End Using
                                End If
                                cb_gdcb_quequan.SelectedIndex = IIf(_IdTTP > 0, CType(arrQueQuan.IndexOf(_IdTTP.ToString()), Integer), 0)
                                cb_gdcb_quequan_SelectedIndexChanged(sender, Nothing)
                                If (dr("IdQueQuan").ToString() <> "") Then
                                    cb_gdcb_xaphuong.SelectedIndex = CType(arrXaPhuong.IndexOf(dr("IdQueQuan").ToString()), Integer)
                                End If
                                edt_gdcb_diachi.Text = dr("DiaChi").ToString().Trim()
                            End If
                        End If
                    End If
                Else    'Trường hợp chưa có bản ghi nào trong hồ sơ gia đình cán bộ
                    dr = _PersonnelFile.GetHuman_ForCode(_IdCanBo)
                    If Not (dr Is Nothing) Then
                        If (dr.Table.Rows.Count > 0) Then
                            cb_gdcb_quoctich.SelectedIndex = IIf(dr("IdQuocTich").ToString() <> "", CType(arrQuocTich.IndexOf(dr("IdQuocTich").ToString()), Integer), 0)
                            cb_gdcb_quequan.SelectedIndex = IIf(dr("IdNS_Tinh").ToString() <> "", CType(arrQueQuan.IndexOf(dr("IdNS_Tinh").ToString()), Integer), 0)
                            cb_gdcb_quequan_SelectedIndexChanged(sender, Nothing)
                            If (dr("IdNS_Xa").ToString() <> "") Then
                                cb_gdcb_xaphuong.SelectedIndex = CType(arrXaPhuong.IndexOf(dr("IdNS_Xa").ToString()), Integer)
                            End If
                            edt_gdcb_diachi.Text = dr("NS_DChi").ToString().Trim()
                        End If
                    End If
                End If
                cb_gdcb_tinhtrang.SelectedIndex = 1
                cb_gdcb_tinhtrang_SelectedIndexChanged(sender, Nothing)
                ActiveControl = edt_gdcb_hoten
                'Case "tp_suckhoe"
                '    If (CheckAdd()) Then
                '        ResetControls(2)
                '        ActiveControl = cb_sk_thanhvien
                '    End If
            Case "tp_gtgc"
                ResetControls(3)
                ActiveControl = edt_gtgc_hoten
            Case "tp_TNGD"
                ResetControls(4)
                ActiveControl = dpkThoiGianKK
        End Select
    End Sub

#Region "---> <---"
#End Region

    Private Sub btn_ghidl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ghidl.Click
        If (_IdCanBo = "") Then Return
        Dim _currRow As String = ""
        If (_RecordId <> "") Then
            _currRow = _RecordId
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_gdcb"                  'Hồ sơ Gia đình cán bộ
                    
                    If (Globals.Roles.IndexOf(";169;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ thành viên trong gia đình cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetControls(1)
                        dgv_gdcb.CurrentRow.Selected = False
                        dgv_gdcb.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_gdcb, "cln_Id")).Selected = True
                        If (dgv_gdcb.Rows.Count > 0) Then
                            If ((dgv_gdcb.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                    'Case "tp_suckhoe"               'Hồ sơ Khám sức khoẻ của Gia đình cán bộ
                    '    If (Globals.Roles.IndexOf(";173;") < 0) Then
                    '        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ sức khoẻ của thành viên trong gia đình cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                    '        ResetControls(2)
                    '        dgv_suckhoe.CurrentRow.Selected = False
                    '        dgv_suckhoe.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_suckhoe, "cln_Id")).Selected = True
                    '        If (dgv_suckhoe.Rows.Count > 0) Then
                    '            If ((dgv_suckhoe.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    '                SelectRow(_currRow)
                    '            End If
                    '        End If
                    '        Return
                    '    End If
                Case "tp_gtgc"               'Hồ sơ Giảm trừ gia cảnh của cán bộ
                    If (Globals.Roles.IndexOf(";177;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ giảm trừ gia cảnh của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetControls(3)
                        dgv_gtgc.CurrentRow.Selected = False
                        dgv_gtgc.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_gtgc, "cln_Id")).Selected = True
                        If (dgv_gtgc.Rows.Count > 0) Then
                            If ((dgv_gtgc.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
                Case "tp_TNGD" ' Ho so thu nhap gia dinh
                    If (Globals.Roles.IndexOf(";177;") < 0) Then
                        MessageBox.Show("Bạn không có quyền sửa đổi hồ sơ thu nhập gia đình của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetControls(4)
                        grid_TNGD.CurrentRow.Selected = False
                        grid_TNGD.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, grid_TNGD, "cln_Id")).Selected = True
                        If (grid_TNGD.Rows.Count > 0) Then
                            If ((grid_TNGD.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        Return
                    End If
            End Select
        End If

        If (IsValid()) Then
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_gdcb"                  'Hồ sơ Gia đình cán bộ
                    Dim obj_gdcb As clsHS_CanBo.HS_GDCB = New clsHS_CanBo.HS_GDCB()
                    obj_gdcb.IdCanBo = _IdCanBo
                    obj_gdcb.Ma_TVien = ""
                    obj_gdcb.HoTen = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_gdcb_hoten.Text.ToString().Trim()))
                    obj_gdcb.GioiTinh = CType(cb_gdcb_gioitinh.SelectedIndex - 1, Byte)
                    obj_gdcb.NamSinh = dtpk_gdcb_namsinh.Value.Year
                    obj_gdcb.IdQuanHe = CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gdcb_quanhe.SelectedIndex), "0"), Integer)
                    obj_gdcb.ConMat = CType(cb_gdcb_tinhtrang.SelectedIndex - 1, Byte)
                    If (cb_gdcb_tinhtrang.SelectedIndex = 2) Then
                        obj_gdcb.NamMat = dtpk_gdcb_nammat.Value.Year
                        obj_gdcb.LyDo_Mat = Globals.Find_Replace(edt_gdcb_lydo.Text.ToString().Trim())
                    End If
                    obj_gdcb.IdUT_BThan = CType(IIf(arrUtBanThan.Count > 0, arrUtBanThan(cb_gdcb_utbanthan.SelectedIndex), "0"), Integer)
                    obj_gdcb.IdQueQuan = CType(IIf(arrXaPhuong.Count > 0, arrXaPhuong(cb_gdcb_xaphuong.SelectedIndex), "0"), Integer)
                    obj_gdcb.DienThoai = Globals.Find_Replace(edt_gdcb_dienthoai.Text.ToString().Trim())
                    obj_gdcb.DiaChi = Globals.Find_Replace(edt_gdcb_diachi.Text.ToString().Trim())
                    obj_gdcb.IdQuocGia = CType(IIf(arrQuocTich.Count > 0, arrQuocTich(cb_gdcb_quoctich.SelectedIndex), "0"), Integer)
                    obj_gdcb.NgheNghiep = Globals.Find_Replace(edt_gdcb_nghenghiep.Text.ToString().Trim())

                    If (_RecordId = "") Then        'Trường hợp thêm mới dữ liệu
                        _currRow = _PersonnelFile.Insert_HS_GDCB(obj_gdcb)
                    Else                            'Trường hợp sửa đổi dữ liệu
                        obj_gdcb.IdTVien = _RecordId
                        _PersonnelFile.Update_HS_GDCB(obj_gdcb)
                        _currRow = _RecordId
                    End If
                    'Thực hiện load lại thông tin sau khi cập nhật song dl
                    ResetControls(1)
                    dgv_gdcb.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    Dim _node As TreeNode = Nothing
                    If (_NodeCurrent <> "") Then
                        _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                        If Not (_node Is Nothing) Then
                            tv_main.SelectedNode = _node
                        End If
                    End If
                    dgv_gdcb.CurrentRow.Selected = False
                    dgv_gdcb.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_gdcb, "cln_Id")).Selected = True
                    If (dgv_gdcb.Rows.Count > 0) Then
                        If ((dgv_gdcb.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If

                    'Case "tp_suckhoe"               'Hồ sơ Khám sức khoẻ của Gia đình cán bộ
                    '    Dim obj_sk_gdcb As clsHS_CanBo.HS_SKhGDCB = New clsHS_CanBo.HS_SKhGDCB()
                    '    obj_sk_gdcb.IdTVien = IIf(arrThVien.Count > 0, arrThVien(cb_sk_thanhvien.SelectedIndex), "")
                    '    obj_sk_gdcb.Nam = dtpk_sk_namkham.Value.Year
                    '    obj_sk_gdcb.Dot = CType(edt_sk_dotkham.Value, Byte)
                    '    obj_sk_gdcb.IdCanBo = _IdCanBo
                    '    obj_sk_gdcb.NoiDungKC = Globals.Find_Replace(edt_sk_ndkham.Text.ToString().Trim())
                    '    obj_sk_gdcb.NoiKham = Globals.Find_Replace(edt_sk_noikham.Text.ToString().Trim())
                    '    obj_sk_gdcb.TuNgay = dtpk_sk_tungay.Value
                    '    obj_sk_gdcb.DenNgay = dtpk_sk_denngay.Value
                    '    obj_sk_gdcb.CanNang = CType(edt_sk_cannang.Text.Trim().Replace(" ", ""), Double)
                    '    obj_sk_gdcb.ChieuCao = CType(edt_sk_chieucao.Text.Trim().Replace(" ", ""), Double)
                    '    obj_sk_gdcb.LoaiSuckhoe = CType(cb_sk_loaisk.SelectedIndex, Byte)
                    '    obj_sk_gdcb.KetLuan = Globals.Find_Replace(edt_sk_ketluan.Text.ToString().Trim())
                    '    obj_sk_gdcb.GhiChu = Globals.Find_Replace(edt_sk_ghichu.Text.ToString().Trim())

                    '    If (_RecordId = "") Then        'Trường hợp thêm mới dữ liệu
                    '        _currRow = _PersonnelFile.Insert_HS_SKhGDCB(obj_sk_gdcb)
                    '    Else                            'Trường hợp sửa đổi dữ liệu
                    '        obj_sk_gdcb.IdKhamchuaGDCB = _RecordId
                    '        _PersonnelFile.Update_HS_SKhGDCB(obj_sk_gdcb)
                    '        _currRow = _RecordId
                    '    End If

                    '    'Thực hiện load lại thông tin sau khi cập nhật song dl
                    '    ResetControls(2)
                    '    dgv_suckhoe.Rows.Clear()
                    '    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    '    Dim _node As TreeNode = Nothing
                    '    If (_NodeCurrent <> "") Then
                    '        _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                    '        If Not (_node Is Nothing) Then
                    '            tv_main.SelectedNode = _node
                    '        End If
                    '    End If
                    '    dgv_suckhoe.CurrentRow.Selected = False
                    '    dgv_suckhoe.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_suckhoe, "cln_Id")).Selected = True
                    '    If (dgv_suckhoe.Rows.Count > 0) Then
                    '        If ((dgv_suckhoe.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    '            SelectRow(_currRow)
                    '        End If
                    '    End If

                Case "tp_gtgc"               'Hồ sơ Giảm trừ gia cảnh của cán bộ
                    If (CheckRelations()) Then
                        Dim objHS_GTGC As HsCanBo.HS_GTGC = New HsCanBo.HS_GTGC()
                        objHS_GTGC.PosCode = DONVI 'IdDONVI
                        objHS_GTGC.IdCanBo = _IdCanBo
                        If (_RecordId = "") Then        'Trường hợp thêm mới dữ liệu
                            objHS_GTGC.IdGTGC = ""
                        Else                            'Trường hợp sửa đổi dữ liệu
                            objHS_GTGC.IdGTGC = _RecordId
                        End If
                        objHS_GTGC.IdQuanHe = CType(IIf(arrQuanHe.Count > 0, arrQuanHe(cb_gtgc_quanhe.SelectedIndex), "0"), Integer)
                        objHS_GTGC.HoTenNguoiPT = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_gtgc_hoten.Text.ToString().Trim()))
                        objHS_GTGC.IdQuocTich = CType(IIf(arrQuocTichGTGC.Count > 0, arrQuocTichGTGC(cb_gtgc_quoctich.SelectedIndex), "0"), Integer)
                        objHS_GTGC.NgaySinh = dtpk_gtgc_ngaysinh.Value
                        objHS_GTGC.TuNgay = dtpk_gtgc_tungay.Value
                        objHS_GTGC.DenNgay = IIf(dtpk_gtgc_denngay.Checked, dtpk_gtgc_denngay.Value, DateTime.Parse("01/01/1900"))
                        objHS_GTGC.MaSoThueNPT = Globals.Find_Replace(edt_gtgc_masothue.Text.Trim.ToString().Replace(" ", "").Replace(".", "").Replace(",", "").Replace("-", "").Replace("_", ""))
                        objHS_GTGC.DiaChi = Globals.Find_Replace(edt_gtgc_diachi.Text.Trim.ToString())
                        objHS_GTGC.DienThoai = Globals.Find_Replace(edt_gtgc_dienthoai.Text.Trim.ToString())
                        objHS_GTGC.SoCMT = Globals.Find_Replace(edt_gtgc_socmt.Text.Trim.ToString())
                        objHS_GTGC.NgayCap = dtpk_gtgc_ngaycap.Value
                        objHS_GTGC.NoiCap = Globals.Find_Replace(edt_gtgc_noicap.Text.Trim.ToString())
                        objHS_GTGC.GhiChu = Globals.Find_Replace(edt_gtgc_ghichu.Text.Trim.ToString())
                        objHS_GTGC.TrangThai = CType(IIf(rb_mo.Checked = True, 1, 0), Byte)
                        objHS_GTGC.CreatedBy = Globals.UserVal
                        objHS_GTGC.ModifiedBy = Globals.UserVal
                        _currRow = _PersonnelFile.Insert_Update_HS_GTGC(objHS_GTGC)

                        'Thực hiện load lại thông tin sau khi cập nhật song dl
                        ResetControls(3)
                        dgv_gtgc.Rows.Clear()
                        tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                        Dim _node As TreeNode = Nothing
                        If (_NodeCurrent <> "") Then
                            _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                            If Not (_node Is Nothing) Then
                                tv_main.SelectedNode = _node
                            End If
                        End If
                        dgv_gtgc.CurrentRow.Selected = False
                        dgv_gtgc.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_gtgc, "cln_Id")).Selected = True
                        If (dgv_gtgc.Rows.Count > 0) Then
                            If ((dgv_gtgc.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                    End If
                Case "tp_TNGD"
                    Dim m_thunhapGD As ThuNhapGiaDinh = New ThuNhapGiaDinh
                    If txtDT_Nha_Cap.Text.Trim = "" Then txtDT_Nha_Cap.Text = "0"
                    If txtDT_Nha_Mua.Text.Trim = "" Then txtDT_Nha_Mua.Text = "0"
                    If txtDT_Dat_Cap.Text.Trim = "" Then txtDT_Dat_Cap.Text = "0"
                    If txtDT_Dat_Mua.Text.Trim = "" Then txtDT_Dat_Mua.Text = "0"
                    m_thunhapGD.IdThuNhapGD = _RecordId
                    m_thunhapGD.IdCanBo = _IdCanBo
                    m_thunhapGD.NgayKekhai = DateTimeUtil.getDate(dpkThoiGianKK.Text)
                    m_thunhapGD.Luong = txtLuong.Text.Trim
                    m_thunhapGD.NguonKhac = txtNguonKhac.Text.Trim
                    m_thunhapGD.NhaO_DuocCap = txtNhaCap.Text.Trim
                    m_thunhapGD.NhaO_DuocCap_DT = CDec(txtDT_Nha_Cap.Text.Trim)
                    m_thunhapGD.NhaO_TuMua = txtNhaMua.Text.Trim
                    m_thunhapGD.NhaO_TuMua_DT = CDec(txtDT_Nha_Mua.Text.Trim)
                    m_thunhapGD.DatO_DuocCap_DT = CDec(txtDT_Dat_Cap.Text.Trim)
                    m_thunhapGD.DatO_TuMua_DT = CDec(txtDT_Dat_Mua.Text.Trim)
                    m_thunhapGD.DatSXKT = txtDat_XSKD.Text.Trim
                    If (_RecordId = "") Then        'Trường hợp thêm mới dữ liệu
                        _currRow = m_thunhapGD.Add
                    Else                            'Trường hợp sửa đổi dữ liệu
                        m_thunhapGD.Update()
                        _currRow = _RecordId
                    End If

                    'Thực hiện load lại thông tin sau khi cập nhật song dl
                    ResetControls(4)
                    grid_TNGD.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    Dim _node As TreeNode = Nothing
                    If (_NodeCurrent <> "") Then
                        _node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                        If Not (_node Is Nothing) Then
                            tv_main.SelectedNode = _node
                        End If
                    End If
                    grid_TNGD.CurrentRow.Selected = False
                    grid_TNGD.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, grid_TNGD, "cln_Id")).Selected = True
                    If (grid_TNGD.Rows.Count > 0) Then
                        If ((grid_TNGD.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            SelectRow(_currRow)
                        End If
                    End If
            End Select
        End If
    End Sub

    Private Sub btn_xoadl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_xoadl.Click
        If (_IdCanBo = "") Then Return
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_gdcb"
                If (dgv_gdcb.Rows.Count <= 0) Then Return
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Integer = 0
                    If (dgv_gdcb.Rows.Count > 0) Then
                        If ((dgv_gdcb.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gdcb.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            For i As Integer = 0 To dgv_gdcb.Rows.Count - 1
                                If (dgv_gdcb.Rows(i).Cells("cln_Id").Value IsNot Nothing) Then
                                    If (CType(dgv_gdcb.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_gdcb.Rows(i).Cells("cln_Id").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ thành viên gia đình cán bộ đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Integer = 0 To arr_Del.Count - 1
                                _PersonnelFile.Delete_HS_GDCB(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ các thành viên gia đình cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetControls(1)
                            dgv_gdcb.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (_NodeCurrent <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_gdcb_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_gdcb, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Cập nhật xoá bỏ hồ sơ thành viên trong gia đình cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try

                'Case "tp_suckhoe"
                '    If (dgv_suckhoe.Rows.Count <= 0) Then Return
                '    Try
                '        Dim arr_Del As ArrayList = New ArrayList()
                '        Dim _count As Int16 = 0
                '        If (dgv_suckhoe.Rows.Count > 0) Then
                '            If ((dgv_suckhoe.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_suckhoe.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                '                For i As Integer = 0 To dgv_suckhoe.Rows.Count - 1
                '                    If (dgv_suckhoe.Rows(i).Cells("cln_Id").Value IsNot Nothing) Then
                '                        If (CType(dgv_suckhoe.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                '                            _count += 1
                '                            arr_Del.Add(dgv_suckhoe.Rows(i).Cells("cln_Id").Value.ToString())
                '                        End If
                '                    End If
                '                Next
                '            End If
                '        End If
                '        If (_count > 0) Then
                '            Dim _mess As String = IIf(_count = 1, "", "các ")
                '            If (MessageBox.Show("Bạn có thực sự muốn xoá " & _mess & "bản ghi hồ sơ đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                '                'Thực hiện xoá dữ liệu khi đã chấp thuận
                '                For i As Int16 = 0 To arr_Del.Count - 1
                '                    _PersonnelFile.Delete_HS_SKhGDCB(arr_Del(i).ToString())
                '                Next
                '                'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                '                MessageBox.Show("Bạn đã xoá thành công hồ sơ khám sức khoẻ của thành viên trong gia đình cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                '                ResetControls(2)
                '                dgv_suckhoe.Rows.Clear()
                '                'Select item của cây dữ liệu
                '                tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                '                Dim node As TreeNode = Nothing
                '                If (_NodeCurrent <> "") Then
                '                    node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                '                    If Not (node Is Nothing) Then
                '                        tv_main.SelectedNode = node
                '                    End If
                '                End If
                '                dgv_suckhoe_CellClick(sender, Nothing)
                '            Else
                '                arr_Del.Clear()
                '                Globals.Check_All_Items(dgv_suckhoe, False)
                '            End If
                '        Else
                '            MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                '        End If
                '    Catch ex As Exception
                '        MessageBox.Show("Xoá bỏ hồ sơ khám sức khoẻ thành viên trong gia đình cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                '    End Try

            Case "tp_gtgc"
                If (dgv_gtgc.Rows.Count <= 0) Then Return
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (dgv_gtgc.Rows.Count > 0) Then
                        If ((dgv_gtgc.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_gtgc.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            For i As Integer = 0 To dgv_gtgc.Rows.Count - 1
                                If (dgv_gtgc.Rows(i).Cells("cln_Id").Value IsNot Nothing) Then
                                    If (CType(dgv_gtgc.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_gtgc.Rows(i).Cells("cln_Id").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " & _mess & "bản ghi hồ sơ giảm trừ gia cảnh đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                _PersonnelFile.Delete_HS_GTGC(arr_Del(i).ToString(), Globals.UserVal, Convert.ToByte("2"))
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ giảm trừ gia cảnh của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetControls(3)
                            dgv_gtgc.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (_NodeCurrent <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_gtgc_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_gtgc, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Xoá bỏ hồ sơ giảm trừ gia cảnh của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
            Case "tp_TNGD"
                If (grid_TNGD.Rows.Count <= 0) Then Return
                Try
                    Dim arr_Del As ArrayList = New ArrayList()
                    Dim _count As Int16 = 0
                    If (grid_TNGD.Rows.Count > 0) Then
                        If ((grid_TNGD.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (grid_TNGD.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                            For i As Integer = 0 To grid_TNGD.Rows.Count - 1
                                If (grid_TNGD.Rows(i).Cells("cln_Id").Value IsNot Nothing) Then
                                    If (CType(grid_TNGD.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(grid_TNGD.Rows(i).Cells("cln_Id").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " & _mess & "bản ghi hồ sơ kê khai thông tin thu nhập gia đình đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                Dim m_ThuNhapGD As ThuNhapGiaDinh = New ThuNhapGiaDinh
                                m_ThuNhapGD.IdThuNhapGD = arr_Del(i).ToString()
                                m_ThuNhapGD.Delete()
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            MessageBox.Show("Bạn đã xoá thành công hồ sơ kê khai thông tin thu nhập gia đình của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            ResetControls(4)
                            grid_TNGD.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (_NodeCurrent <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, _NodeCurrent)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            grid_TNGD_CellClick(sender, Nothing)
                        Else
                            arr_Del.Clear()
                            Globals.Check_All_Items(grid_TNGD, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Xoá bỏ hồ sơ kê khai thông tin thu nhập gia đình của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End Try
        End Select
    End Sub
#End Region

#Region "---> Events: Các sự kiện bắt ngoại lệ nhập liệu <---"
    'Bắt điều kiện nhập dl - Ví dụ như nhập số
    'Private Sub edt_sk_cannang_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '    Dim KeyAscii As Integer
    '    KeyAscii = Asc(e.KeyChar)
    '    If (KeyAscii = 46) Then
    '        Dim n As Integer = edt_sk_cannang.Text.IndexOf(".")
    '        If (n >= 0) Then e.Handled = True
    '    End If
    '    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46) Then
    '        e.Handled = True
    '    End If
    'End Sub

    'Private Sub edt_sk_chieucao_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '    Dim KeyAscii As Integer
    '    KeyAscii = Asc(e.KeyChar)
    '    If (KeyAscii = 46) Then
    '        Dim n As Integer = edt_sk_chieucao.Text.IndexOf(".")
    '        If (n >= 0) Then e.Handled = True
    '    End If
    '    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46) Then
    '        e.Handled = True
    '    End If
    'End Sub

    'Private Sub edt_sk_cannang_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    edt_sk_cannang.Text = Globals.Convert_ToDouble(edt_sk_cannang, vNFInfo).ToString("N", vNFInfo)
    'End Sub

    'Private Sub edt_sk_chieucao_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    edt_sk_chieucao.Text = Globals.Convert_ToDouble(edt_sk_chieucao, vNFInfo).ToString("N", vNFInfo)
    'End Sub

    'Private Sub edt_sk_cannang_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    '    edt_sk_cannang.SelectAll()
    'End Sub

    'Private Sub edt_sk_chieucao_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    '    edt_sk_chieucao.SelectAll()
    'End Sub

    Private Sub edt_gdcb_dienthoai_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles edt_gdcb_dienthoai.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    'Các sự kiện di chuyển cho người nhập liệu - Tab Hồ sơ gia đình cán bộ

    Private Sub edt_gdcb_hoten_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_hoten.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = cb_gdcb_gioitinh
        End If
    End Sub

    Private Sub cb_gdcb_gioitinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_gdcb_gioitinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = dtpk_gdcb_namsinh
        End If
    End Sub

    Private Sub dtpk_gdcb_namsinh_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_gdcb_namsinh.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_gdcb_quanhe
        End If
    End Sub

    Private Sub cb_gdcb_quanhe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_gdcb_quanhe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_gdcb_quoctich
        End If
    End Sub

    Private Sub cb_gdcb_quoctich_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_gdcb_quoctich.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_gdcb_quequan
        End If
    End Sub

    Private Sub cb_gdcb_quequan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_gdcb_quequan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_gdcb_nghenghiep
        End If
    End Sub

    Private Sub edt_gdcb_nghenghiep_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_nghenghiep.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_gdcb_dienthoai
        End If
    End Sub

    Private Sub edt_gdcb_nghenghiep_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_nghenghiep.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = cb_gdcb_quequan
        End If
    End Sub

    Private Sub edt_gdcb_dienthoai_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_dienthoai.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = edt_gdcb_diachi
        End If
    End Sub

    Private Sub edt_gdcb_dienthoai_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_dienthoai.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_gdcb_nghenghiep
        End If
    End Sub

    Private Sub edt_gdcb_diachi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_diachi.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = cb_gdcb_utbanthan
        End If
    End Sub

    Private Sub edt_gdcb_diachi_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_diachi.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_gdcb_dienthoai
        End If
    End Sub

    Private Sub cb_gdcb_utbanthan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_gdcb_utbanthan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = cb_gdcb_tinhtrang
        End If
    End Sub

    Private Sub cb_gdcb_tinhtrang_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_gdcb_tinhtrang.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (cb_gdcb_tinhtrang.SelectedIndex = 2) Then
                ActiveControl = dtpk_gdcb_nammat
            Else
                ActiveControl = btn_ghidl
            End If
        End If
    End Sub

    Private Sub dtpk_gdcb_nammat_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_gdcb_nammat.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_gdcb_lydo
        End If
    End Sub

    Private Sub edt_gdcb_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = btn_ghidl
        End If
    End Sub

    Private Sub edt_gdcb_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_gdcb_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = dtpk_gdcb_nammat
        End If
    End Sub

    'Các sự kiện di chuyển cho người nhập liệu - Tab Hồ sơ Khám sức khoẻ gia đình cán bộ

    'Private Sub cb_sk_thanhvien_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = dtpk_sk_namkham
    '    End If
    'End Sub

    'Private Sub dtpk_sk_namkham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = edt_sk_dotkham
    '        If (_RecordId = "") Then
    '            dtpk_sk_tungay.Value = New DateTime(dtpk_sk_namkham.Value.Year, 1, 1)
    '            dtpk_sk_denngay.Value = New DateTime(dtpk_sk_namkham.Value.Year, 12, 31)
    '        End If
    '    End If
    'End Sub

    'Private Sub dtpk_sk_namkham_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If (_RecordId = "") Then
    '        dtpk_sk_tungay.Value = New DateTime(dtpk_sk_namkham.Value.Year, 1, 1)
    '        dtpk_sk_denngay.Value = New DateTime(dtpk_sk_namkham.Value.Year, 12, 31)
    '    End If
    'End Sub

    'Private Sub edt_sk_dotkham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = edt_sk_ndkham
    '    End If
    'End Sub

    'Private Sub edt_sk_ndkham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = edt_sk_noikham
    '    End If
    'End Sub

    'Private Sub edt_sk_ndkham_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = edt_sk_dotkham
    '    End If
    'End Sub

    'Private Sub edt_sk_noikham_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = dtpk_sk_tungay
    '    End If
    'End Sub

    'Private Sub edt_sk_noikham_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = edt_sk_ndkham
    '    End If
    'End Sub

    'Private Sub dtpk_sk_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = dtpk_sk_denngay
    '    End If
    'End Sub

    'Private Sub dtpk_sk_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = edt_sk_cannang
    '    End If
    'End Sub

    'Private Sub edt_sk_cannang_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = edt_sk_chieucao
    '    End If
    'End Sub

    'Private Sub edt_sk_cannang_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = dtpk_sk_denngay
    '    End If
    'End Sub

    'Private Sub edt_sk_chieucao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = cb_sk_loaisk
    '    End If
    'End Sub

    'Private Sub edt_sk_chieucao_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = edt_sk_cannang
    '    End If
    'End Sub

    'Private Sub cb_sk_loaisk_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = edt_sk_ketluan
    '    End If
    'End Sub

    'Private Sub edt_sk_ketluan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = edt_sk_ghichu
    '    End If
    'End Sub

    'Private Sub edt_sk_ketluan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = cb_sk_loaisk
    '    End If
    'End Sub

    Private Sub edt_sk_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            ActiveControl = btn_ghidl
        End If
    End Sub

    'Private Sub edt_sk_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = edt_sk_ketluan
    '    End If
    'End Sub

    'Các sự kiện di chuyển cho người nhập liệu - Tab Hồ sơ Giảm trừ gia cảnh

    'Private Sub edt_gtgc_hoten_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = cb_gtgc_quanhe
    '    End If
    'End Sub

    'Private Sub cb_gtgc_quanhe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = dtpk_gtgc_tungay
    '    End If
    'End Sub

    'Private Sub dtpk_gtgc_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = dtpk_gtgc_denngay
    '    End If
    'End Sub

    'Private Sub dtpk_gtgc_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
    '        ActiveControl = edt_gtgc_ghichu
    '    End If
    'End Sub

    'Private Sub edt_gtgc_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
    '        ActiveControl = btn_ghidl
    '    End If
    'End Sub

    'Private Sub edt_gtgc_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If (e.KeyCode = Keys.Up) Then
    '        ActiveControl = dtpk_gtgc_denngay
    '    End If
    'End Sub
#End Region

    Private Sub cb_gdcb_quanhe_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_gdcb_quanhe.SelectedIndexChanged
        If (_RecordId = "" And cb_gdcb_quanhe.SelectedIndex <> 0) Then
            Dim sDiaChiTmp As String = ""
            Dim IDTV As String = ""
            Dim iId_Dc_Tinh, iId_Dc_XaPhuong, Id_QuocTich As Integer
            Dim db As DBAccess = New DBAccess
            Dim dt As DataTable
            Select Case arrQuanHe(cb_gdcb_quanhe.SelectedIndex)  'CType(arrQuanHe.IndexOf(dr("IdQuanHe").ToString())
                Case 463, 458, 460, 459  ' lay theo ca nhan
                    dt = db.SelectDBRows("SELECT IdThT_Tinh, IdThT_Xa,ThT_DiaChi, IdQuocTich FROM HS_CanBo WHERE IDCanBo='" & _IdCanBo & "'")
                    If dt.Rows.Count > 0 Then
                        sDiaChiTmp = dt.Rows(0).Item("ThT_DiaChi")
                        iId_Dc_Tinh = dt.Rows(0).Item("IdThT_Tinh")
                        iId_Dc_XaPhuong = dt.Rows(0).Item("IdThT_Xa")
                        Id_QuocTich = dt.Rows(0).Item("IdQuocTich")
                    End If
                Case 450, 466, 451, 462, 461  'lay theo dia chi cua bo
                    IDTV = db.getString(" SELECT IdTVien FROM HS_GDCB WHERE IdCanBo='" & _IdCanBo & "' And IdQuanhe = 449")
                    If IDTV <> "" Then
                        dt = db.SelectDBRows("SELECT IdQueQuan, DiaChi, IdQuocGia FROM HS_GDCB WHERE IdTVien='" & IDTV & "'")
                        If dt.Rows.Count > 0 Then
                            sDiaChiTmp = dt.Rows(0).Item("DiaChi")
                            iId_Dc_XaPhuong = dt.Rows(0).Item("IdQueQuan")
                            Id_QuocTich = dt.Rows(0).Item("IdQuocGia")
                            iId_Dc_Tinh = db.getNumber(String.Format("Select Id From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By TrangThai)", iId_Dc_XaPhuong))
                        End If
                    End If
                Case 448, 455, 454, 453, 467 ' lay theo dia chi ong noi
                    IDTV = db.getString(" SELECT IdTVien FROM HS_GDCB WHERE IdCanBo='" & _IdCanBo & "' and IdQuanhe=447")
                    If IDTV <> "" Then
                        dt = db.SelectDBRows("SELECT IdQueQuan, DiaChi, IdQuocGia FROM HS_GDCB WHERE IdTVien='" & IDTV & "'")
                        If dt.Rows.Count > 0 Then
                            sDiaChiTmp = dt.Rows(0).Item("DiaChi")
                            iId_Dc_XaPhuong = dt.Rows(0).Item("IdQueQuan")
                            Id_QuocTich = dt.Rows(0).Item("IdQuocGia")
                            iId_Dc_Tinh = db.getNumber(String.Format("Select Id From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By TrangThai)", iId_Dc_XaPhuong))
                        End If
                    End If
                Case 695, 457, 452, 456  ' lay theo dia chi ong ngoai
                    IDTV = db.getString(" SELECT IdTVien FROM HS_GDCB WHERE IdCanBo='" & _IdCanBo & "' and IdQuanhe=767")
                    If IDTV <> "" Then
                        dt = db.SelectDBRows("SELECT IdQueQuan, DiaChi, IdQuocGia FROM HS_GDCB WHERE IdTVien='" & IDTV & "'")
                        If dt.Rows.Count > 0 Then
                            sDiaChiTmp = dt.Rows(0).Item("DiaChi")
                            iId_Dc_XaPhuong = dt.Rows(0).Item("IdQueQuan")
                            Id_QuocTich = dt.Rows(0).Item("IdQuocGia")
                            iId_Dc_Tinh = db.getNumber(String.Format("Select Id From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By TrangThai)", iId_Dc_XaPhuong))
                        End If
                    End If
                Case 471, 475, 476, 477  ' lay theo dia chi bo chong
                    IDTV = db.getString(" SELECT IdTVien FROM HS_GDCB WHERE IdCanBo='" & _IdCanBo & "' and IdQuanhe=470")
                    If IDTV <> "" Then
                        dt = db.SelectDBRows("SELECT IdQueQuan, DiaChi, IdQuocGia FROM HS_GDCB WHERE IdTVien='" & IDTV & "'")
                        If dt.Rows.Count > 0 Then
                            sDiaChiTmp = dt.Rows(0).Item("DiaChi")
                            iId_Dc_XaPhuong = dt.Rows(0).Item("IdQueQuan")
                            Id_QuocTich = dt.Rows(0).Item("IdQuocGia")
                            iId_Dc_Tinh = db.getNumber(String.Format("Select Id From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By TrangThai)", iId_Dc_XaPhuong))
                        End If
                    End If
                Case 469, 472, 473, 474  ' lay theo dia chi bo vo
                    IDTV = db.getString(" SELECT IdTVien FROM HS_GDCB WHERE IdCanBo='" & _IdCanBo & "' and IdQuanhe=468")
                    If IDTV <> "" Then
                        dt = db.SelectDBRows("SELECT IdQueQuan, DiaChi, IdQuocGia FROM HS_GDCB WHERE IdTVien='" & IDTV & "'")
                        If dt.Rows.Count > 0 Then
                            sDiaChiTmp = dt.Rows(0).Item("DiaChi")
                            iId_Dc_XaPhuong = dt.Rows(0).Item("IdQueQuan")
                            Id_QuocTich = dt.Rows(0).Item("IdQuocGia")
                            iId_Dc_Tinh = db.getNumber(String.Format("Select Id From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Ma_Tinh In (Select Top 1 X.Ma_Tinh From Dm_DiaPhuong X Where X.Ma_Xa<>'00' And X.Ma_Thon='00' And X.Id={0} Order By TrangThai)", iId_Dc_XaPhuong))
                        End If
                    End If
                Case Else
                    Id_QuocTich = 0
            End Select
            If Id_QuocTich <> 0 Then
                edt_gdcb_diachi.Text = sDiaChiTmp
                cb_gdcb_quoctich.SelectedIndex = CType(arrQuocTich.IndexOf(Id_QuocTich.ToString()), Integer)
                cb_gdcb_quequan.SelectedIndex = CType(arrQueQuan.IndexOf(iId_Dc_Tinh.ToString()), Integer)
                cb_gdcb_xaphuong.SelectedIndex = CType(arrXaPhuong.IndexOf(iId_Dc_XaPhuong.ToString()), Integer)
            End If
        End If
    End Sub

    Private Sub txtDT_Nha_Cap_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDT_Nha_Cap.TextChanged
        txtDT_Nha_Cap.Text = formatDouble(txtDT_Nha_Cap.Text.Trim)
    End Sub

    Private Sub txtDT_Nha_Mua_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDT_Nha_Mua.TextChanged
        txtDT_Nha_Mua.Text = formatDouble(txtDT_Nha_Mua.Text.Trim)
    End Sub

    Private Sub txtDT_Dat_Cap_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDT_Dat_Cap.TextChanged
        txtDT_Dat_Cap.Text = formatDouble(txtDT_Dat_Cap.Text.Trim)
    End Sub

    Private Sub txtDT_Dat_Mua_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDT_Dat_Mua.TextChanged
        txtDT_Dat_Mua.Text = formatDouble(txtDT_Dat_Mua.Text.Trim)
    End Sub

    Private Sub edt_gtgc_masothue_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_gtgc_masothue.KeyPress
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

    Private Sub edt_gtgc_dienthoai_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_gtgc_dienthoai.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_gtgc_hoten_Leave(sender As Object, e As EventArgs) Handles edt_gtgc_hoten.Leave
        If _RecordId = "" And _IdCanBo <> "" Then
            strSQL = String.Format("Select * From HS_GDCB Where IdCanBo = '{0}' And HoTen Like N'%{1}%' ", _IdCanBo, edt_gtgc_hoten.Text)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        If db.Rows(0)("NamSinh").ToString().Trim() <> "" Then
                            dtpk_gtgc_ngaysinh.Value = DateTimeUtil.StringToDateTime("01/01/" + db.Rows(0)("NamSinh").ToString(), "dd/MM/yyyy")
                            If CType(db.Rows(0)("NamSinh").ToString(), Integer) >= 2014 Then
                                dtpk_gtgc_tungay.Value = dtpk_gtgc_ngaysinh.Value.AddMonths(1)
                            End If
                        End If
                        cb_gtgc_quoctich.SelectedIndex = IIf(db.Rows(0)("IdQuocGia").ToString() <> "", CType(arrQuocTichGTGC.IndexOf(db.Rows(0)("IdQuocGia").ToString()), Integer), 0)
                        edt_gtgc_diachi.Text = db.Rows(0)("DiaChi").ToString().Trim()
                        edt_gtgc_dienthoai.Text = db.Rows(0)("DienThoai").ToString().Trim()
                        cb_gtgc_quanhe.SelectedIndex = IIf(db.Rows(0)("IdQuanHe").ToString() <> "", CType(arrQuanHe.IndexOf(db.Rows(0)("IdQuanHe").ToString()), Integer), 0)
                    End If
                End If
            End Using
        End If
    End Sub
End Class