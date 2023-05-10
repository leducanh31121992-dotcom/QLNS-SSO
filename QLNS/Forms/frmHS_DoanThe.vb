Public Class frmHS_DoanThe

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Globals As Globals = New Globals
    Private _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()
    Private _ListDocument As clsHT_DanhMuc = New clsHT_DanhMuc()
    Private _HS_DangDT As clsHS_DangDT = New clsHS_DangDT()
    Private _IdCanBo As String = ""
    Private _RowId As String = ""  'Biến lưu Id hồ sơ - để lưu id khi sửa đổi
    Private strNode As String = ""
    Private _SqlHelper As DBAccess = New DBAccess()

    'Khai báo biến sử dụng khi được gọi từ Hồ sơ cán bộ
    Public HumanId As String = ""       'Giá trị id của cán bộ
    Public FlagShow As Boolean = False 'Giá trị True là được gọi từ Hồ sơ cán bộ. False: Giá trị mặc định
    Public TagNode As String = ""       'Giá trị lưu lại tag đơn vị đang select bên hồ sơ cán bộ
    Private _IdDonviHT As Integer = IdDONVI
#End Region

#Region "---> Functions: Các hàm chính <---"
    ''' <summary>
    ''' Hàm thực hiện reset controls về trạng thái mặc định
    ''' </summary>
    ''' <param name="_State">Chỉ số xác định trạng thái reset. Với quy ước:
    '''         0: Reset all controls: Reset toàn bộ controls
    '''         1: Reset hồ sơ đảng viên
    '''         2: Reset hồ sơ đoàn viên
    '''         3: Reset hồ sơ Công đoàn
    ''' </param>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls(ByVal _State As Byte)
        _RowId = ""
        Select Case _State
            Case 0
                _IdCanBo = ""
                dgv_dangvien.Rows.Clear()
                dgv_doanvien.Rows.Clear()
                dgv_congdoan.Rows.Clear()
                'Thông tin Chung về hồ sơ nhân sự
                lbl_macb.Text = ""
                lbl_hoten.Text = ""
                lbl_ngaysinh.Text = ""
                lbl_gioitinh.Text = ""
                lbl_so_cmt.Text = ""
                chkTWQuanLy.Checked = False
                'Thông tin Hồ sơ đảng viên - của cán bộ
                edt_dang_sothe.Text = ""
                dtpk_dang_ngaykn.Text = DateTime.Now.ToShortDateString()
                dtpk_dang_ngayvao.Text = DateTime.Now.ToShortDateString()
                edt_dang_noi_kn.Text = ""
                edt_dang_nguoi_gt.Text = ""
                edt_dang_noicapthe.Text = ""
                dtpk_dang_ngayra.Text = DateTime.Now.ToShortDateString()
                dtpk_dang_ngayra.Checked = False
                edt_dang_lydo.ReadOnly = True
                edt_dang_lydo.Text = ""
                'Thông tin Hồ sơ đoàn viên - của cán bộ
                edt_doan_sothe.Text = ""
                dtpk_doan_ngayvao.Text = DateTime.Now.ToShortDateString()
                edt_doan_noi_kn.Text = ""
                edt_doan_noicapthe.Text = ""
                dtpk_doan_ngayra.Text = DateTime.Now.ToShortDateString()
                dtpk_doan_ngayra.Checked = False
                edt_doan_lydo.ReadOnly = True
                edt_doan_lydo.Text = ""
                'Thông tin Hồ sơ Công đoàn - của cán bộ
                edt_congdoan_sothe.Text = ""
                dtpk_congdoan_ngayvao.Text = DateTime.Now.ToShortDateString()
                edt_congdoan_coquan.Text = ""
                edt_congdoan_noicapthe.Text = ""
                dtpk_congdoan_ngayra.Text = DateTime.Now.ToShortDateString()
                dtpk_congdoan_ngayra.Checked = False
                edt_congdoan_lydo.ReadOnly = True
                edt_congdoan_lydo.Text = ""
            Case 1      'Hồ sơ Đảng viên
                edt_dang_sothe.Text = ""
                dtpk_dang_ngaykn.Text = DateTime.Now.ToShortDateString()
                dtpk_dang_ngayvao.Text = DateTime.Now.ToShortDateString()
                edt_dang_noi_kn.Text = ""
                edt_dang_nguoi_gt.Text = ""
                edt_dang_noicapthe.Text = ""
                dtpk_dang_ngayra.Text = DateTime.Now.ToShortDateString()
                dtpk_dang_ngayra.Checked = False
                edt_dang_lydo.ReadOnly = True
                edt_dang_lydo.Text = ""
            Case 2      'Hồ sơ Đoàn viên
                edt_doan_sothe.Text = ""
                dtpk_doan_ngayvao.Text = DateTime.Now.ToShortDateString()
                edt_doan_noi_kn.Text = ""
                edt_doan_noicapthe.Text = ""
                dtpk_doan_ngayra.Text = DateTime.Now.ToShortDateString()
                dtpk_doan_ngayra.Checked = False
                edt_doan_lydo.ReadOnly = True
                edt_doan_lydo.Text = ""
            Case 3      'Hồ sơ Công đoàn
                edt_congdoan_sothe.Text = ""
                dtpk_congdoan_ngayvao.Text = DateTime.Now.ToShortDateString()
                edt_congdoan_coquan.Text = ""
                edt_congdoan_noicapthe.Text = ""
                dtpk_congdoan_ngayra.Text = DateTime.Now.ToShortDateString()
                dtpk_congdoan_ngayra.Checked = False
                edt_congdoan_lydo.ReadOnly = True
                edt_congdoan_lydo.Text = ""
        End Select
    End Sub

    ''' <summary>
    ''' Fucntion check permits of current user manipulation programs
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        'Kiểm tra có quyền nào về Hồ sơ đảng không (Nếu không thực hiện remove tab - Hồ sơ đảng)?
        If Not (Globals.IsIntersect(";163;164;165;166;", _roles)) Then       ' Hồ sơ - Đảng viên
            If tctrl_main.TabPages.Contains(tp_hs_dang) Then tctrl_main.TabPages.Remove(tp_hs_dang)
        End If
        'Kiểm tra có quyền nào về Hồ sơ Đoàn không (Nếu không thực hiện remove tab - Hồ sơ Đoàn)?
        If Not (Globals.IsIntersect(";220;221;222;223;", _roles)) Then       ' Hồ sơ - Đoàn viên
            If tctrl_main.TabPages.Contains(tp_hs_doan) Then tctrl_main.TabPages.Remove(tp_hs_doan)
        End If
        'Kiểm tra có quyền nào về Hồ sơ Công đoàn không (Nếu không thực hiện remove tab - Hồ sơ Công đoàn)?
        If Not (Globals.IsIntersect(";225;226;227;228;", _roles)) Then       ' Hồ sơ - Công đoàn
            If tctrl_main.TabPages.Contains(tp_hs_congdoan) Then tctrl_main.TabPages.Remove(tp_hs_congdoan)
        End If

        'Nếu không có quyền Xem thông tin với cả 3 hồ sơ
        If (Globals.Roles.IndexOf(";163;") < 0 And Globals.Roles.IndexOf(";220;") < 0 And Globals.Roles.IndexOf(";225;") < 0) Then
            tv_main.Enabled = False
        End If
        'Nếu không có quyền thêm mới với cả 3 hồ sơ
        If (Globals.Roles.IndexOf(";164;") < 0 And Globals.Roles.IndexOf(";221;") < 0 And Globals.Roles.IndexOf(";226;") < 0) Then
            btn_add.Enabled = False
        End If
        ''Nếu không có quyền xoá bỏ với cả 3 hồ sơ
        If (Globals.Roles.IndexOf("166;") < 0 And Globals.Roles.IndexOf("223;") < 0 And Globals.Roles.IndexOf(";228;") < 0) Then
            btn_delete.Enabled = False
        End If
        ''Nếu không có quyền thêm mới và sửa đổi với cả 3 hồ sơ
        If (Globals.Roles.IndexOf(";164;") < 0 And Globals.Roles.IndexOf(";165;") < 0 And Globals.Roles.IndexOf(";221;") < 0 And Globals.Roles.IndexOf(";222;") < 0 And Globals.Roles.IndexOf(";226;") < 0 And Globals.Roles.IndexOf(";227;") < 0) Then
            btn_accept.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu - Các hồ sơ theo Id cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Mã hiệu xác định cán bộ</param>
    ''' <remarks></remarks>
    Private Sub FillAll_Documents(ByVal _IdCanBo As String)
        'Fill data - Hồ sơ Đảng của cán bộ
        Using db As DataTable = _HS_DangDT.GetAll(_IdCanBo, 1)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_dangvien.Rows.Add()
                        dgv_dangvien.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("IdDangVien").ToString() <> "", db.Rows(i)("IdDangVien").ToString(), "")
                        dgv_dangvien.Rows(i).Cells("cln_SoThe").Value = IIf(db.Rows(i)("SoThe").ToString() <> "", db.Rows(i)("SoThe").ToString(), "")
                        If db.Rows(i)("NgayKN").ToString().Trim() <> "" Then
                            dgv_dangvien.Rows(i).Cells("cln_NgayKN").Value = IIf(CType(db.Rows(i)("NgayKN").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayKN").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_dangvien.Rows(i).Cells("cln_NgayKN").Value = ""
                        End If
                        If db.Rows(i)("NgayVao").ToString().Trim() <> "" Then
                            dgv_dangvien.Rows(i).Cells("cln_Ngayvao").Value = IIf(CType(db.Rows(i)("NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_dangvien.Rows(i).Cells("cln_Ngayvao").Value = ""
                        End If
                        dgv_dangvien.Rows(i).Cells("cln_Noi_kn").Value = IIf(db.Rows(i)("Noi_KetNap").ToString() <> "", db.Rows(i)("Noi_KetNap").ToString(), "")
                        dgv_dangvien.Rows(i).Cells("cln_NoiCapThe").Value = IIf(db.Rows(i)("NoiCapThe").ToString() <> "", db.Rows(i)("NoiCapThe").ToString(), "")

                        dgv_dangvien.Rows(i).Cells("cln_Nguoi_gt").Value = IIf(db.Rows(i)("Nguoi_GT").ToString() <> "", db.Rows(i)("Nguoi_GT").ToString(), "")
                        If (db.Rows(i)("NgayRa").ToString().Trim() <> "") Then
                            dgv_dangvien.Rows(i).Cells("cln_Ngayra").Value = IIf(CType(db.Rows(i)("NgayRa").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayRa").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_dangvien.Rows(i).Cells("cln_Ngayra").Value = ""
                        End If
                        dgv_dangvien.Rows(i).Cells("cln_Lydo").Value = IIf(db.Rows(i)("LyDo").ToString() <> "", db.Rows(i)("LyDo").ToString(), "")
                    Next
                End If
            End If
        End Using

        'Fill data - Hồ sơ Đoàn viên của cán bộ
        Using db As DataTable = _HS_DangDT.GetAll(_IdCanBo, 2)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_doanvien.Rows.Add()
                        dgv_doanvien.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("IdDoanVien").ToString() <> "", db.Rows(i)("IdDoanVien").ToString(), "")
                        dgv_doanvien.Rows(i).Cells("cln_SoThe").Value = IIf(db.Rows(i)("SoThe").ToString() <> "", db.Rows(i)("SoThe").ToString(), "")
                        If db.Rows(i)("NgayVao").ToString().Trim() <> "" Then
                            dgv_doanvien.Rows(i).Cells("cln_Ngayvao").Value = IIf(CType(db.Rows(i)("NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_doanvien.Rows(i).Cells("cln_Ngayvao").Value = ""
                        End If
                        dgv_doanvien.Rows(i).Cells("cln_Noi_kn").Value = IIf(db.Rows(i)("Noi_KetNap").ToString() <> "", db.Rows(i)("Noi_KetNap").ToString(), "")
                        dgv_doanvien.Rows(i).Cells("cln_NoiCapThe").Value = IIf(db.Rows(i)("NoiCapThe").ToString() <> "", db.Rows(i)("NoiCapThe").ToString(), "")
                        If (db.Rows(i)("NgayRa").ToString().Trim() <> "") Then
                            dgv_doanvien.Rows(i).Cells("cln_Ngayra").Value = IIf(CType(db.Rows(i)("NgayRa").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayRa").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_doanvien.Rows(i).Cells("cln_Ngayra").Value = ""
                        End If
                        dgv_doanvien.Rows(i).Cells("cln_Lydo").Value = IIf(db.Rows(i)("LyDo").ToString() <> "", db.Rows(i)("LyDo").ToString(), "")
                    Next
                End If
            End If
        End Using

        'Fill data - Hồ sơ Công đoàn của cán bộ
        Using db As DataTable = _HS_DangDT.GetAll(_IdCanBo, 3)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_congdoan.Rows.Add()
                        dgv_congdoan.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("IdCongDoan").ToString() <> "", db.Rows(i)("IdCongDoan").ToString(), "")
                        dgv_congdoan.Rows(i).Cells("cln_Sothe").Value = IIf(db.Rows(i)("SoThe").ToString() <> "", db.Rows(i)("SoThe").ToString(), "")
                        If db.Rows(i)("NgayVao").ToString().Trim() <> "" Then
                            dgv_congdoan.Rows(i).Cells("cln_Ngayvao").Value = IIf(CType(db.Rows(i)("NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayVao").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_congdoan.Rows(i).Cells("cln_Ngayvao").Value = ""
                        End If
                        dgv_congdoan.Rows(i).Cells("cln_Noi_kn").Value = IIf(db.Rows(i)("Noi_Vao").ToString() <> "", db.Rows(i)("Noi_Vao").ToString(), "")
                        dgv_congdoan.Rows(i).Cells("cln_NoiCapThe").Value = IIf(db.Rows(i)("NoiCapThe").ToString() <> "", db.Rows(i)("NoiCapThe").ToString(), "")
                        If (db.Rows(i)("NgayRa").ToString().Trim() <> "") Then
                            dgv_congdoan.Rows(i).Cells("cln_Ngayra").Value = IIf(CType(db.Rows(i)("NgayRa").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("NgayRa").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                        Else
                            dgv_congdoan.Rows(i).Cells("cln_Ngayra").Value = ""
                        End If
                        dgv_congdoan.Rows(i).Cells("cln_Lydo").Value = IIf(db.Rows(i)("LyDo").ToString() <> "", db.Rows(i)("LyDo").ToString(), "")
                    Next
                End If
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Hàm thực hiện select row trên lưới dữ liệu --> Fill dữ liệu vào các controls
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi (Id record) đang select</param>
    ''' <param name="_State">Chỉ số xác định hồ sơ. Với quy ước: 1-Đảng, 2-Đoàn, 3-Công đoàn</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String, ByVal _State As Byte)
        Dim dr As DataRow
        _RowId = _Idnew
        dr = _HS_DangDT.GetRecord(_RowId, _State)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                Select Case _State
                    Case 1  'Hồ sơ đảng
                        edt_dang_sothe.Text = dr("SoThe").ToString().Trim()
                        If dr("NgayKN").ToString().Trim() <> "" Then
                            dtpk_dang_ngaykn.Value = CType(dr("NgayKN").ToString(), DateTime)
                        End If
                        If dr("NgayVao").ToString().Trim() <> "" Then
                            If (CType(dr("NgayVao"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_dang_ngayvao.Checked = False
                            Else
                                dtpk_dang_ngayvao.Checked = True
                                dtpk_dang_ngayvao.Value = CType(dr("NgayVao").ToString(), DateTime)
                            End If
                        Else
                            dtpk_dang_ngayvao.Checked = False
                        End If
                        edt_dang_noi_kn.Text = dr("Noi_KetNap").ToString().Trim()
                        edt_dang_nguoi_gt.Text = dr("Nguoi_GT").ToString().Trim()
                        edt_dang_noicapthe.Text = dr("NoiCapThe").ToString().Trim()
                        If dr("NgayRa").ToString().Trim() <> "" Then
                            If (CType(dr("NgayRa"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_dang_ngayra.Checked = False
                            Else
                                dtpk_dang_ngayra.Checked = True
                                dtpk_dang_ngayra.Value = CType(dr("NgayRa").ToString(), DateTime)
                            End If
                        Else
                            dtpk_dang_ngayra.Checked = False
                        End If
                        edt_dang_lydo.Text = dr("LyDo").ToString().Trim()
                    Case 2  'Hồ sơ đoàn
                        edt_doan_sothe.Text = dr("SoThe").ToString().Trim()
                        If dr("NgayVao").ToString().Trim() <> "" Then
                            dtpk_doan_ngayvao.Value = CType(dr("NgayVao").ToString(), DateTime)
                        End If
                        edt_doan_noi_kn.Text = dr("Noi_KetNap").ToString().Trim()
                        edt_doan_noicapthe.Text = dr("NoiCapThe").ToString().Trim()
                        If dr("NgayRa").ToString().Trim() <> "" Then
                            If (CType(dr("NgayRa"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_doan_ngayra.Checked = False
                            Else
                                dtpk_doan_ngayra.Checked = True
                                dtpk_doan_ngayra.Value = CType(dr("NgayRa").ToString(), DateTime)
                            End If
                        Else
                            dtpk_doan_ngayra.Checked = False
                        End If
                        edt_doan_lydo.Text = dr("LyDo").ToString().Trim()
                    Case 3  'Hồ sơ Công đoàn
                        edt_congdoan_sothe.Text = dr("SoThe").ToString().Trim()
                        If dr("NgayVao").ToString().Trim() <> "" Then
                            dtpk_congdoan_ngayvao.Value = CType(dr("NgayVao").ToString(), DateTime)
                        End If
                        edt_congdoan_coquan.Text = dr("Noi_Vao").ToString().Trim()
                        edt_congdoan_noicapthe.Text = dr("NoiCapThe").ToString().Trim()
                        If dr("NgayRa").ToString().Trim() <> "" Then
                            If (CType(dr("NgayRa"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                dtpk_congdoan_ngayra.Checked = False
                            Else
                                dtpk_congdoan_ngayra.Checked = True
                                dtpk_congdoan_ngayra.Value = CType(dr("NgayRa").ToString(), DateTime)
                            End If
                        Else
                            dtpk_congdoan_ngayra.Checked = False
                        End If
                        edt_congdoan_lydo.Text = dr("LyDo").ToString().Trim()
                End Select
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra xem Hồ sơ đoàn thể còn hiệu lực không. Nếu còn hiệu lực thì không cho thêm tiếp
    ''' </summary>
    ''' <returns>True-Hợp lệ. False-Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsExist() As Boolean
        Dim strSQL As String = ""
        Dim vTitle As String = ""
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_dang"
                strSQL = String.Format("Select * From HS_DangVien Where IdCanBo = '{0}' Order by NgayVao Desc", _IdCanBo)
                vTitle = "Đảng"
            Case "tp_hs_doan"
                strSQL = String.Format("Select * From HS_DoanVien Where IdCanBo = '{0}' Order by NgayVao Desc", _IdCanBo)
                vTitle = "Đoàn"
            Case "tp_hs_congdoan"
                strSQL = String.Format("Select * From HS_CongDoan Where IdCanBo = '{0}' Order by NgayVao Desc", _IdCanBo)
                vTitle = "Công đoàn"
        End Select
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    If (db.Rows(0)("NgayRa").ToString() = "") Then
                        MessageBox.Show("Hồ sơ " & vTitle & " của cán bộ " + lbl_hoten.Text.Trim() + " đã tồn tại. Thao tác bị huỷ bỏ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        Return False
                    ElseIf (CType(db.Rows(0)("NgayRa"), DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
                        MessageBox.Show("Hồ sơ " & vTitle & " của cán bộ " + lbl_hoten.Text.Trim() + " đã tồn tại. Thao tác bị huỷ bỏ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        Return False
                    End If
                End If
            End If
        End Using
        Return True
    End Function

    ''' <summary>
    ''' 'Hàm thực hiện kiểm tra tính hợp lệ của việc cập nhật dữ liệu
    ''' </summary>
    ''' <returns>True: Is success. Reverse -> False</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_dang"
                'If (edt_dang_sothe.Text.Trim() = "") Then
                '    MessageBox.Show("Số thẻ đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_dang_sothe
                '    Return False
                'End If
                'Bắt thông tin 16 tuổi trở lên mới được kết nạp đảng
                Dim dt_ngaysinh As DateTime = New DateTime(lbl_ngaysinh.Text.Trim().Substring(6), lbl_ngaysinh.Text.Trim().Substring(3, 2), lbl_ngaysinh.Text.Trim().Substring(0, 2))
                Dim dt_ngay_dtc As DateTime = dt_ngaysinh.AddYears(16)
                'If (dtpk_dang_ngaykn.Value < dt_ngay_dtc) Then
                '    MessageBox.Show("Ngày kết nạp đảng không hợp lệ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = dtpk_dang_ngaykn
                '    Return False
                'End If
                If (dtpk_dang_ngaykn.Checked = True) Then
                    If (dtpk_dang_ngaykn.Value >= dtpk_dang_ngayvao.Value) Then
                        MessageBox.Show("Ngày vào đảng chính thức không thể nhỏ hơn ngày kết nạp. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_dang_ngaykn
                        Return False
                    End If
                End If
                If (dtpk_dang_ngayra.Checked = True) Then
                    If dtpk_dang_ngaykn.Value >= dtpk_dang_ngayra.Value Then
                        MessageBox.Show("Ngày vào đảng chính thức không thể lớn hơn hoặc bằng ngày ra khỏi đảng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_dang_ngaykn
                        Return False
                    End If
                    If dtpk_dang_ngayra.Value > DateTime.Now Then
                        MessageBox.Show("Ngày ra khỏi đảng không thể lớn hơn ngày hiện thời!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_dang_ngayra
                        Return False
                    End If
                    If (edt_dang_lydo.Text.Trim() = "") Then
                        MessageBox.Show("Lý do ra khỏi đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_dang_lydo
                        Return False
                    End If
                End If
                'If (edt_dang_noi_kn.Text.Trim() = "") Then
                '    MessageBox.Show("Nơi kết nạp đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_dang_noi_kn
                '    Return False
                'End If
                'If (edt_dang_nguoi_gt.Text.Trim() = "") Then
                '    MessageBox.Show("Người giới thiệu vào đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_dang_nguoi_gt
                '    Return False
                'End If
                'Kiêm tra xem đã có bản ghi nào chưa ---> Có rùi thì kiểm tra khoảng thời gian không có trùng 2 lần kết nạp đảng
                Dim strSQL As String = ""
                If (_RowId = "") Then
                    strSQL = String.Format("Select * from HS_DangVien Where IdCanBo = '{0}' Order by NgayKN Desc", _IdCanBo)
                ElseIf (_RowId <> "") Then
                    strSQL = String.Format("Select * from HS_DangVien Where IdCanBo = '{0}' And IdDangVien <> '{1}' Order by NgayKN Desc", _IdCanBo, _RowId)
                End If
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                If (dtpk_dang_ngaykn.Value <= CType(db.Rows(0)("NgayRa").ToString(), DateTime)) Then
                                    MessageBox.Show("        Thời gian vào đảng của cán bộ không hợp lệ." & vbCrLf & "Thời gian này bị trùng vào khoảng thời gian của hồ sơ trước!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = dtpk_dang_ngaykn
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End Using
                'Kiểm tra khi sửa một bản ghi hồ sơ đảng mà đã có ngày ra. Người sử dụng Sửa lại cho dl ngày ra = NoThing
                If (_RowId <> "") Then
                    Dim _Flag As Boolean = False
                    strSQL = String.Format("Select * from HS_DangVien Where IdCanBo = '{0}' and IdDangVien = '{1}' Order by NgayKN Desc", _IdCanBo, _RowId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                    _Flag = True
                                End If
                            End If
                        End If
                    End Using
                    If (_Flag = True) Then
                        If (dgv_dangvien.Rows.Count > 1 And dtpk_dang_ngayra.Checked = False) Then
                            MessageBox.Show("Hồ sơ đảng này đã có ngày ra khỏi đảng, bạn vui lòng chọn lại ngày ra khỏi đảng", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                            ActiveControl = dtpk_dang_ngayra
                            Return False
                        End If
                    End If
                End If
            Case "tp_hs_doan"   'Hồ sơ đoàn viên
                'If (edt_doan_sothe.Text.Trim() = "") Then
                '    MessageBox.Show("Số thẻ đoàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_doan_sothe
                '    Return False
                'End If
                'Bắt thông tin 16 tuổi trở lên mới được kết nạp đảng
                Dim dt_ngaysinh As DateTime = New DateTime(lbl_ngaysinh.Text.Trim().Substring(6), lbl_ngaysinh.Text.Trim().Substring(3, 2), lbl_ngaysinh.Text.Trim().Substring(0, 2))
                Dim dt_ngay_dtc As DateTime = dt_ngaysinh.AddYears(14)
                If (dtpk_doan_ngayvao.Value < dt_ngay_dtc) Then
                    MessageBox.Show("Ngày vào đoàn không hợp lệ. Vui lòng kiểm tra lại!" + vbCrLf + "Lưu ý: Theo quy định đoàn viên kết nạp tối thiểu phải từ 14 tuổi trở lên.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_doan_ngayvao
                    Return False
                End If

                If (dtpk_doan_ngayra.Checked = True) Then
                    If dtpk_doan_ngayvao.Value >= dtpk_doan_ngayra.Value Then
                        MessageBox.Show("Ngày vào đoàn không thể lớn hơn hoặc bằng ngày ra khỏi đoàn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_doan_ngayra
                        Return False
                    End If
                    If dtpk_doan_ngayra.Value > DateTime.Now Then
                        MessageBox.Show("Ngày ra khỏi đoàn không thể lớn hơn ngày hiện thời!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_doan_ngayra
                        Return False
                    End If
                    If (edt_doan_lydo.Text.Trim() = "") Then
                        MessageBox.Show("Lý do ra khỏi đoàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_doan_lydo
                        Return False
                    End If
                End If
                'If (edt_doan_noi_kn.Text.Trim() = "") Then
                '    MessageBox.Show("Nơi kết nạp đoàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_doan_noi_kn
                '    Return False
                'End If
                'Kiêm tra xem đã có bản ghi nào chưa ---> Có rùi thì kiểm tra khoảng thời gian không có trùng 2 lần kết nạp đảng
                Dim strSQL As String = ""
                If (_RowId = "") Then
                    strSQL = String.Format("Select * from HS_DoanVien Where IdCanBo = '{0}' Order by NgayVao Desc", _IdCanBo)
                ElseIf (_RowId <> "") Then
                    strSQL = String.Format("Select * from HS_DoanVien Where IdCanBo = '{0}' And IdDoanVien <> '{1}' Order by NgayVao Desc", _IdCanBo, _RowId)
                End If
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                If (dtpk_doan_ngayvao.Value <= CType(db.Rows(0)("NgayRa").ToString(), DateTime)) Then
                                    MessageBox.Show("        Thời gian vào đoàn của cán bộ không hợp lệ." & vbCrLf & "Lưu ý: Thời gian này bị trùng vào khoảng thời gian của hồ sơ trước!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = dtpk_doan_ngayvao
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End Using
                'Kiểm tra khi sửa một bản ghi hồ sơ Đoàn mà đã có ngày ra. Người sử dụng Sửa lại cho dl ngày ra = NoThing
                If (_RowId <> "") Then
                    Dim _Flag As Boolean = False
                    strSQL = String.Format("Select * from HS_DoanVien Where IdCanBo = '{0}' and IdDoanVien = '{1}' Order by NgayVao Desc", _IdCanBo, _RowId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                    _Flag = True
                                End If
                            End If
                        End If
                    End Using
                    If (_Flag = True) Then
                        If (dgv_doanvien.Rows.Count > 1 And dtpk_doan_ngayra.Checked = False) Then
                            MessageBox.Show("Hồ sơ đoàn này đã có ngày ra khỏi đoàn trước khi bạn sửa, bạn vui lòng chọn lại ngày ra khỏi đoàn.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                            ActiveControl = dtpk_doan_ngayra
                            Return False
                        End If
                    End If
                End If
            Case "tp_hs_congdoan"
                'If (edt_congdoan_sothe.Text.Trim() = "") Then
                '    MessageBox.Show("Số thẻ công đoàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_congdoan_sothe
                '    Return False
                'End If
                'Bắt thông tin 16 tuổi trở lên mới được kết nạp đảng
                Dim dt_ngaysinh As DateTime = New DateTime(lbl_ngaysinh.Text.Trim().Substring(6), lbl_ngaysinh.Text.Trim().Substring(3, 2), lbl_ngaysinh.Text.Trim().Substring(0, 2))
                Dim dt_ngay_dtc As DateTime = dt_ngaysinh.AddYears(18)
                If (dtpk_congdoan_ngayvao.Value < dt_ngay_dtc) Then
                    MessageBox.Show("Ngày vào công đoàn không hợp lệ. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_congdoan_ngayvao
                    Return False
                End If
                If (dtpk_congdoan_ngayra.Checked = True) Then
                    If dtpk_congdoan_ngayvao.Value >= dtpk_congdoan_ngayra.Value Then
                        MessageBox.Show("Ngày vào không thể lớn hơn hoặc bằng ngày ra!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_congdoan_ngayra
                        Return False
                    End If
                    If dtpk_congdoan_ngayra.Value > DateTime.Now Then
                        MessageBox.Show("Ngày ra không thể lớn hơn ngày hiện thời!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = dtpk_congdoan_ngayra
                        Return False
                    End If
                    If (edt_congdoan_lydo.Text.Trim() = "") Then
                        MessageBox.Show("Lý do ra khỏi công đoàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                        ActiveControl = edt_congdoan_lydo
                        Return False
                    End If
                End If
                'If (edt_congdoan_coquan.Text.Trim() = "") Then
                '    MessageBox.Show("Đơn vị bạn tham gia vào công đoàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                '    ActiveControl = edt_congdoan_coquan
                '    Return False
                'End If

                'Kiêm tra xem đã có bản ghi nào chưa ---> Có rùi thì kiểm tra khoảng thời gian không có trùng, đan xen thời gian không?
                Dim strSQL As String = ""
                If (_RowId = "") Then
                    strSQL = String.Format("Select * from HS_CongDoan Where IdCanBo = '{0}' Order by NgayVao Desc", _IdCanBo)
                ElseIf (_RowId <> "") Then
                    strSQL = String.Format("Select * from HS_CongDoan Where IdCanBo = '{0}' And IdCongDoan <> '{1}' Order by NgayVao Desc", _IdCanBo, _RowId)
                End If
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                If (dtpk_congdoan_ngayvao.Value <= CType(db.Rows(0)("NgayRa").ToString(), DateTime)) Then
                                    MessageBox.Show("        Thời gian vào đảng của cán bộ không hợp lệ." & vbCrLf & "Thời gian này bị trùng vào khoảng thời gian của hồ sơ trước!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = dtpk_congdoan_ngayvao
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End Using
                'Kiểm tra khi sửa một bản ghi hồ sơ Công Đoàn mà đã có ngày ra. Người sử dụng Sửa lại cho dl ngày ra = NoThing
                If (_RowId <> "") Then
                    Dim _Flag As Boolean = False
                    strSQL = String.Format("Select * from HS_CongDoan Where IdCanBo = '{0}' and IdCongDoan = '{1}' Order by NgayVao Desc", _IdCanBo, _RowId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                    _Flag = True
                                End If
                            End If
                        End If
                    End Using
                    If (_Flag = True) Then
                        If (dgv_congdoan.Rows.Count > 1 And dtpk_congdoan_ngayra.Checked = False) Then
                            MessageBox.Show("Hồ sơ công đoàn này đã có ngày ra khỏi đoàn trước khi bạn sửa, bạn vui lòng chọn lại ngày ra khỏi công đoàn.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                            ActiveControl = dtpk_congdoan_ngayra
                            Return False
                        End If
                    End If
                End If
        End Select
        Return True
    End Function
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHS_DoanThe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '--> Định nghĩa lưới dữ liệu hiển thị Hồ sơ
        _HS_DangDT.Create_Frame(dgv_dangvien, 1)
        _HS_DangDT.Create_Frame(dgv_doanvien, 3)
        _HS_DangDT.Create_Frame(dgv_congdoan, 5)
        '--> Fill dữ liệu ra cây dữ liệu cơ cấu tổ chức
        _HS_CanBo.Fill_Tree(tv_main)
        '--> Thực hiện reset all controls
        ResetAll_Controls(0)
        lkl_xemct.Enabled = False
        btn_quatrinhsh.Visible = False
        tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
        strNode = ""
        '--> Nếu được gọi từ Menu trên lưới của Hồ sơ cán bộ
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
                _PhongBanId = _HS_CanBo.GetDebtId(HumanId, _DonviId)
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
                btn_accept.Enabled = True
                btn_delete.Enabled = True
                btn_quatrinhsh.Enabled = True
            Else
                btn_accept.Enabled = False
                btn_delete.Enabled = False
                btn_quatrinhsh.Enabled = False
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

    Private Sub tctrl_main_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tctrl_main.SelectedIndexChanged
        btn_accept.Enabled = True
        btn_delete.Enabled = True
        btn_quatrinhsh.Visible = False
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_dang"       ' Quyền Hồ sơ Đảng:         163;164;165;166;
                dgv_dangvien_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";163;") < 0) Then
                    dgv_dangvien.Enabled = False
                Else
                    dgv_dangvien.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";164;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";166;") < 0) Then
                    btn_delete.Enabled = False
                End If
                If (dgv_dangvien.Rows.Count > 0) Then
                    btn_quatrinhsh.Visible = True
                End If
                btn_quatrinhsh.Text = "Quá trình sinh &hoạt đảng"
                btn_quatrinhsh.Width = 161
            Case "tp_hs_doan"       ' Quyền hồ sơ Đoàn viên:    220;221;222;223;
                dgv_doanvien_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";220;") < 0) Then
                    dgv_doanvien.Enabled = False
                Else
                    dgv_doanvien.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";221;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";223;") < 0) Then
                    btn_delete.Enabled = False
                End If
                If (dgv_doanvien.Rows.Count > 0) Then
                    btn_quatrinhsh.Visible = True
                End If
                btn_quatrinhsh.Text = "Quá trình sinh &hoạt đoàn"
                btn_quatrinhsh.Width = 161
            Case "tp_hs_congdoan"       ' Quyền hồ sơ Công đoàn:    225;226;227;228;
                dgv_congdoan_CellClick(sender, Nothing)
                If (Globals.Roles.IndexOf(";225;") < 0) Then
                    dgv_congdoan.Enabled = False
                Else
                    dgv_congdoan.Enabled = True
                End If
                If (Globals.Roles.IndexOf(";226;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (Globals.Roles.IndexOf(";228;") < 0) Then
                    btn_delete.Enabled = False
                End If
                If (dgv_congdoan.Rows.Count > 0) Then
                    btn_quatrinhsh.Visible = True
                End If
                btn_quatrinhsh.Text = "Quá trình sinh &hoạt công đoàn"
                btn_quatrinhsh.Width = 191
        End Select
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        Try
            lkl_xemct.Enabled = False
            ResetAll_Controls(0)
            If (tv_main.Nodes.Count > 0) Then
                'Thực hiện Load child node khi click vào parent node 
                Dim arrElement() As String
                If tv_main.SelectedNode.GetNodeCount(True) = 0 Then
                    If Not (tv_main.SelectedNode.IsExpanded) Then
                        arrElement = tv_main.SelectedNode.Tag.ToString().Trim().Split("_")
                        If (arrElement.Length > 0) Then
                            If (arrElement(0) = "DV") Then
                                _HS_CanBo.Fill_Node(tv_main.SelectedNode, arrElement(1), arrElement(2))
                            ElseIf arrElement(0) = "PB" Then
                                _HS_CanBo.Fill_NodeCanbo(tv_main.SelectedNode, arrElement(1), arrElement(2), True)
                            End If
                        End If
                    End If
                End If
                'Thực hiện load dữ liệu cán bộ khi click vào Từng cán bộ
                strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                'Lấy Chỉ số xác định đơn vị Hiện tại của cán bộ (Xét với trường hợp cài đặt đơn vị '" & gMaDonViTW & "')
                If (DONVI = gMaDonViTW) Then
                    If (strNode.IndexOf("DV_") >= 0) Then
                        _IdDonviHT = CType(strNode.Substring(3, 1), Integer)
                    End If
                End If
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "CB") Then
                    strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                    _IdCanBo = tv_main.SelectedNode.Tag.ToString().Trim().Substring(tv_main.SelectedNode.Tag.ToString().LastIndexOf("_") + 1)
                    If (_IdCanBo <> "") Then
                        'Fill data - thông tin chung về nhân sự
                        If (Globals.Roles.IndexOf(";71;") < 0) Then
                            lkl_xemct.Enabled = False
                        Else
                            lkl_xemct.Enabled = True
                        End If
                        Dim dr As DataRow
                        dr = _HS_CanBo.GetHuman_ForCode(_IdCanBo)
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
                                chkTWQuanLy.Checked = My_CBool(dr("CapQuanLy"), False)
                            End If
                        End If

                        'Thực hiện fill data HS đảng viên của cán bộ ra lưới dữ liệu
                        FillAll_Documents(_IdCanBo)
                        '--> Thực thi sự kiện select tab
                        tctrl_main_SelectedIndexChanged(sender, Nothing)

                        If _IdCanBo <> "" Then
                            If checkRight_CreateRecord(_IdCanBo) Then
                                btn_accept.Enabled = True
                                btn_delete.Enabled = True
                                btn_quatrinhsh.Enabled = True
                            Else
                                btn_accept.Enabled = False
                                btn_delete.Enabled = False
                                btn_quatrinhsh.Enabled = False
                            End If
                        End If
                        
                    End If
                End If
                tv_main.SelectedNode.Expand()
            End If
        Catch ex As Exception
            MessageBox.Show("Load dữ liệu khi click vào cây dữ liệu: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub dgv_dangvien_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_dangvien.CellClick
        ResetAll_Controls(1)
        If (dgv_dangvien.Rows.Count > 0) Then
            If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                SelectRow(dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString(), 1)
            End If
        End If
        If (dtpk_dang_ngayra.Checked = False) Then
            edt_dang_lydo.ReadOnly = True
        Else
            edt_dang_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dgv_doanvien_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_doanvien.CellClick
        ResetAll_Controls(2)
        If (dgv_doanvien.Rows.Count > 0) Then
            If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                SelectRow(dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString(), 2)
            End If
        End If
        If (dtpk_doan_ngayra.Checked = False) Then
            edt_doan_lydo.ReadOnly = True
        Else
            edt_doan_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dgv_congdoan_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_congdoan.CellClick
        ResetAll_Controls(3)
        If (dgv_congdoan.Rows.Count > 0) Then
            If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                SelectRow(dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString(), 3)
            End If
        End If
        If (dtpk_congdoan_ngayra.Checked = False) Then
            edt_congdoan_lydo.ReadOnly = True
        Else
            edt_congdoan_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dgv_dangvien_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_dangvien.KeyUp
        dgv_dangvien_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_doanvien_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_doanvien.KeyUp
        dgv_doanvien_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_congdoan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_congdoan.KeyUp
        dgv_congdoan_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_dangvien_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_dangvien.CellDoubleClick
        If (dgv_dangvien.Rows.Count > 0) Then
            If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Dim obj_qtsinhhoat As New frmHS_DoanTheQT
                obj_qtsinhhoat.DocumentId = dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString()
                obj_qtsinhhoat.FlagDocument = 1
                obj_qtsinhhoat.ShowDialog()
            End If
        End If
    End Sub

    Private Sub dgv_doanvien_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_doanvien.CellDoubleClick
        If (dgv_doanvien.Rows.Count > 0) Then
            If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Dim obj_qtsinhhoat As New frmHS_DoanTheQT
                obj_qtsinhhoat.DocumentId = dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString()
                obj_qtsinhhoat.FlagDocument = 2
                obj_qtsinhhoat.ShowDialog()
            End If
        End If
    End Sub

    Private Sub dgv_congdoan_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_congdoan.CellDoubleClick
        If (dgv_congdoan.Rows.Count > 0) Then
            If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Dim obj_qtsinhhoat As New frmHS_DoanTheQT
                obj_qtsinhhoat.DocumentId = dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString()
                obj_qtsinhhoat.FlagDocument = 3
                obj_qtsinhhoat.ShowDialog()
            End If
        End If
    End Sub

    Private Sub btn_quatrinhsh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quatrinhsh.Click
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_dang"
                If (dgv_dangvien.Rows.Count > 0) Then
                    If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        Dim obj_qtsinhhoat As New frmHS_DoanTheQT
                        obj_qtsinhhoat.DocumentId = dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString()
                        obj_qtsinhhoat.FlagDocument = 1
                        obj_qtsinhhoat.ShowDialog()
                    End If
                End If
            Case "tp_hs_doan"
                If (dgv_doanvien.Rows.Count > 0) Then
                    If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        Dim obj_qtsinhhoat As New frmHS_DoanTheQT
                        obj_qtsinhhoat.DocumentId = dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString()
                        obj_qtsinhhoat.FlagDocument = 2
                        obj_qtsinhhoat.ShowDialog()
                    End If
                End If
            Case "tp_hs_congdoan"
                If (dgv_congdoan.Rows.Count > 0) Then
                    If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        Dim obj_qtsinhhoat As New frmHS_DoanTheQT
                        obj_qtsinhhoat.DocumentId = dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString()
                        obj_qtsinhhoat.FlagDocument = 3
                        obj_qtsinhhoat.ShowDialog()
                    End If
                End If
        End Select
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        If (_IdCanBo = "") Then
            MessageBox.Show("Bạn chưa chọn cán bộ cần thêm mới hồ sơ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        If (IsExist()) Then
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_hs_dang"       'Hồ sơ Đảng viên
                    ResetAll_Controls(1)
                    'Nếu cán bộ này chuyển từ Hồ sơ tập sự sang hồ sơ chính. Thực hiện load dữ liệu hồ sơ đảng (nếu có) có sẵn sang
                    If (dgv_dangvien.Rows.Count <= 0) Then
                        'Kiểm tra xem có phải là cán bộ chuyển từ Tập sự sang không
                        Dim strSQL As String = String.Format("Select * From HSCB_TS Where idNew = '{0}' And TrangThai = 1", _IdCanBo)
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    edt_dang_sothe.Text = db.Rows(0)("Dang_SoThe").ToString()
                                    If db.Rows(0)("Dang_NgayVao").ToString().Trim() <> "" Then
                                        dtpk_dang_ngaykn.Value = CType(db.Rows(0)("Dang_NgayVao").ToString(), DateTime)
                                    End If
                                    If db.Rows(0)("Dang_NgayChTh").ToString().Trim() <> "" Then
                                        dtpk_dang_ngayvao.Value = CType(db.Rows(0)("Dang_NgayChTh").ToString(), DateTime)
                                    End If
                                    edt_dang_noi_kn.Text = db.Rows(0)("Dang_NoiKN").ToString()
                                    edt_dang_nguoi_gt.Text = db.Rows(0)("Dang_NGT").ToString()
                                    If db.Rows(0)("Dang_NgayRa").ToString().Trim() <> "" Then
                                        If (CType(db.Rows(0)("Dang_NgayRa"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                            dtpk_dang_ngayra.Checked = False
                                        Else
                                            dtpk_dang_ngayra.Checked = True
                                            dtpk_dang_ngayra.Value = CType(db.Rows(0)("Dang_NgayRa").ToString(), DateTime)
                                        End If
                                    Else
                                        dtpk_dang_ngayra.Checked = False
                                    End If
                                    edt_dang_lydo.Text = db.Rows(0)("Dang_LyDoRa").ToString()
                                    If (dtpk_dang_ngayra.Checked = False) Then
                                        edt_dang_lydo.ReadOnly = True
                                    Else
                                        edt_dang_lydo.ReadOnly = False
                                    End If
                                    MessageBox.Show("Thông tin hồ sơ đảng của cán bộ '" + lbl_hoten.Text.Trim() + "' được lấy từ hồ sơ cán bộ tập sự mà người sử dụng đã khai báo trước đó!" + vbCrLf + "Lưu ý: Các thông tin hồ sơ đảng được hiển thị kế thừa giúp người sử dụng không phải nhập liệu lại, người sử dụng có thể nhập lại sau đó cập nhật.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                                    ActiveControl = edt_dang_sothe
                                    Return
                                End If
                            End If
                        End Using
                    End If
                    ActiveControl = edt_dang_sothe
                Case "tp_hs_doan"       'Hồ sơ Đoàn viên
                    ResetAll_Controls(2)
                    ActiveControl = edt_doan_sothe
                Case "tp_hs_congdoan"   'Hồ sơ Công đoàn
                    ResetAll_Controls(3)
                    ActiveControl = edt_congdoan_sothe
            End Select
        End If
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        If (_IdCanBo = "") Then Return
        Dim _currRow As String = ""
        '--> Thực hiện kiểm tra quyền được sửa đổi dữ liệu - Hồ sơ
        If (_RowId <> "") Then
            _currRow = _RowId
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_hs_dang"
                    If (Globals.Roles.IndexOf(";165;") < 0) Then
                        MessageBox.Show("Bạn không có quyền thực hiện sửa đổi hồ sơ đảng viên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(1)
                        'Gọi lại sự kiện cell click của lưới dữ liệu
                        dgv_dangvien.CurrentRow.Selected = False
                        dgv_dangvien.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_dangvien, "cln_Code")).Selected = True
                        If (dgv_dangvien.Rows.Count > 0) Then
                            If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow, 1)
                            End If
                        End If
                        If (dtpk_dang_ngayra.Checked = False) Then
                            edt_dang_lydo.ReadOnly = True
                        Else : edt_dang_lydo.ReadOnly = False
                        End If
                        Return
                    End If
                Case "tp_hs_doan"
                    If (Globals.Roles.IndexOf(";222;") < 0) Then
                        MessageBox.Show("Bạn không có quyền thực hiện sửa đổi hồ sơ đoàn viên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(2)
                        'Gọi lại sự kiện cell click của lưới dữ liệu
                        dgv_doanvien.CurrentRow.Selected = False
                        dgv_doanvien.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_doanvien, "cln_Code")).Selected = True
                        If (dgv_doanvien.Rows.Count > 0) Then
                            If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow, 2)
                            End If
                        End If
                        If (dtpk_doan_ngayra.Checked = False) Then
                            edt_doan_lydo.ReadOnly = True
                        Else : edt_doan_lydo.ReadOnly = False
                        End If
                        Return
                    End If
                Case "tp_hs_congdoan"
                    If (Globals.Roles.IndexOf(";227;") < 0) Then
                        MessageBox.Show("Bạn không có quyền thực hiện sửa đổi hồ sơ công đoàn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls(3)
                        'Gọi lại sự kiện cell click của lưới dữ liệu
                        dgv_congdoan.CurrentRow.Selected = False
                        dgv_congdoan.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_congdoan, "cln_Code")).Selected = True
                        If (dgv_congdoan.Rows.Count > 0) Then
                            If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow, 2)
                            End If
                        End If
                        If (dtpk_congdoan_ngayra.Checked = False) Then
                            edt_congdoan_lydo.ReadOnly = True
                        Else : edt_congdoan_lydo.ReadOnly = False
                        End If
                        Return
                    End If
            End Select
        End If

        '--> Băt đầu thực hiện cập nhật dữ liệu - Hồ sơ đoàn thể
        If (IsValid()) Then
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_hs_dang"       'HỒ SƠ ĐẢNG VIÊN
                    ' Biến lấy giá trị khoá hiện thời khi thêm hoặc sửa để select về đúng row đó trên lưới dữ liệu
                    Dim obj_hs_dang As clsHS_DangDT.HS_DangVien = New clsHS_DangDT.HS_DangVien()
                    obj_hs_dang.IdCanBo = _IdCanBo
                    obj_hs_dang.SoThe = Globals.Find_Replace(edt_dang_sothe.Text.ToString().Trim())
                    obj_hs_dang.NgayKN = dtpk_dang_ngaykn.Value
                    obj_hs_dang.NgayVao = IIf(dtpk_dang_ngayvao.Checked = True, dtpk_dang_ngayvao.Value, DateTime.Parse("01/01/1900"))
                    obj_hs_dang.Noi_KetNap = Globals.Find_Replace(edt_dang_noi_kn.Text.ToString().Trim())
                    obj_hs_dang.NoiCapThe = Globals.Find_Replace(edt_dang_noicapthe.Text.ToString().Trim())
                    obj_hs_dang.Nguoi_GT = Globals.MakeFirstWordCharUpper(Globals.Find_Replace(edt_dang_nguoi_gt.Text.ToString().Trim()))
                    obj_hs_dang.NgayRa = IIf(dtpk_dang_ngayra.Checked = True, dtpk_dang_ngayra.Value, DateTime.Parse("01/01/1900"))
                    obj_hs_dang.LyDo = Globals.Find_Replace(edt_dang_lydo.Text.ToString().Trim())
                    If (_RowId = "") Then
                        _currRow = _HS_DangDT.Insert_PartyDocument(obj_hs_dang)
                    Else
                        obj_hs_dang.IdDangVien = _RowId
                        _HS_DangDT.Update_PartyDocument(obj_hs_dang)
                        _currRow = _RowId
                    End If
                    btn_quatrinhsh.Visible = True
                    ResetAll_Controls(1)
                    _IdCanBo = ""
                    dgv_dangvien.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    'Gọi lại sự kiện cell click của lưới dữ liệu
                    dgv_dangvien.CurrentRow.Selected = False
                    dgv_dangvien.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_dangvien, "cln_Code")).Selected = True
                    If (dgv_dangvien.Rows.Count > 0) Then
                        If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow, 1)
                        End If
                    End If
                    If (dtpk_dang_ngayra.Checked = False) Then
                        edt_dang_lydo.ReadOnly = True
                    Else
                        edt_dang_lydo.ReadOnly = False
                    End If
                Case "tp_hs_doan"       'HỒ SƠ ĐOÀN VIÊN
                    Dim obj_hsdoan As clsHS_DangDT.HS_DoanVien = New clsHS_DangDT.HS_DoanVien()
                    obj_hsdoan.IdCanBo = _IdCanBo
                    obj_hsdoan.SoThe = Globals.Find_Replace(edt_doan_sothe.Text.ToString().Trim())
                    obj_hsdoan.NgayVao = dtpk_doan_ngayvao.Value
                    obj_hsdoan.Noi_KetNap = Globals.Find_Replace(edt_doan_noi_kn.Text.ToString().Trim())
                    obj_hsdoan.NoiCapThe = Globals.Find_Replace(edt_doan_noicapthe.Text.ToString().Trim())
                    obj_hsdoan.NgayRa = IIf(dtpk_doan_ngayra.Checked = True, dtpk_doan_ngayra.Value, DateTime.Parse("01/01/1900"))
                    obj_hsdoan.LyDo = Globals.Find_Replace(edt_doan_lydo.Text.ToString().Trim())
                    If (_RowId = "") Then
                        _currRow = _HS_DangDT.Insert_UnionMember(obj_hsdoan)
                    Else
                        obj_hsdoan.IdDoanVien = _RowId
                        _HS_DangDT.Update_UnionMember(obj_hsdoan)
                        _currRow = _RowId
                    End If
                    btn_quatrinhsh.Visible = True
                    ResetAll_Controls(2)
                    _IdCanBo = ""
                    dgv_doanvien.Rows.Clear()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If
                    'Gọi lại sự kiện cell click của lưới dữ liệu
                    dgv_doanvien.CurrentRow.Selected = False
                    dgv_doanvien.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_doanvien, "cln_Code")).Selected = True
                    If (dgv_doanvien.Rows.Count > 0) Then
                        If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow, 2)
                        End If
                    End If
                    If (dtpk_doan_ngayra.Checked = False) Then
                        edt_doan_lydo.ReadOnly = True
                    Else
                        edt_doan_lydo.ReadOnly = False
                    End If
                Case "tp_hs_congdoan"       'HỒ SƠ CÔNG ĐOÀN    
                    Dim obj_TradeUnion As clsHS_DangDT.HS_CongDoan = New clsHS_DangDT.HS_CongDoan()
                    obj_TradeUnion.IdCanBo = _IdCanBo
                    obj_TradeUnion.SoThe = Globals.Find_Replace(edt_congdoan_sothe.Text.ToString().Trim())
                    obj_TradeUnion.NgayVao = dtpk_congdoan_ngayvao.Value
                    obj_TradeUnion.Noi_Vao = Globals.Find_Replace(edt_congdoan_coquan.Text.ToString().Trim())
                    obj_TradeUnion.NoiCapThe = Globals.Find_Replace(edt_congdoan_noicapthe.Text.ToString().Trim())
                    obj_TradeUnion.NgayRa = IIf(dtpk_congdoan_ngayra.Checked = True, dtpk_congdoan_ngayra.Value, DateTime.Parse("01/01/1900"))
                    obj_TradeUnion.LyDo = Globals.Find_Replace(edt_congdoan_lydo.Text.ToString().Trim())
                    If (_RowId = "") Then
                        _currRow = _HS_DangDT.Insert_TradeUnion(obj_TradeUnion)
                    Else
                        obj_TradeUnion.IdCongDoan = _RowId
                        _HS_DangDT.Update_TradeUnion(obj_TradeUnion)
                        _currRow = _RowId
                    End If
                    ResetAll_Controls(3)
                    _IdCanBo = ""
                    dgv_congdoan.Rows.Clear()
                    btn_quatrinhsh.Visible = True
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                    'Select item của cây dữ liệu
                    Dim node As TreeNode = Nothing
                    If (strNode <> "") Then
                        node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                        If Not (node Is Nothing) Then
                            tv_main.SelectedNode = node
                        End If
                    End If

                    'Gọi lại sự kiện cell click của lưới dữ liệu
                    dgv_congdoan.CurrentRow.Selected = False
                    dgv_congdoan.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_congdoan, "cln_Code")).Selected = True
                    If (dgv_congdoan.Rows.Count > 0) Then
                        If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            SelectRow(_currRow, 3)
                        End If
                    End If
                    If (dtpk_congdoan_ngayra.Checked = False) Then
                        edt_congdoan_lydo.ReadOnly = True
                    Else
                        edt_congdoan_lydo.ReadOnly = False
                    End If
            End Select
        End If
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If (_IdCanBo = "") Then Return
        Try
            Dim _count As Int16 = 0
            Dim arr_Del As ArrayList = New ArrayList()
            Select Case tctrl_main.SelectedTab.Name
                Case "tp_hs_dang"       'THỰC HIỆN XOÁ HỒ SƠ ĐẢNG CÙNG QUÁ TRÌNH SINH HOẠT ĐẢNG
                    If (dgv_dangvien.Rows.Count <= 0) Then Return
                    If (dgv_dangvien.Rows.Count > 0) Then
                        If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Int32 = 0 To dgv_dangvien.Rows.Count - 1
                                If (dgv_dangvien.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_dangvien.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_dangvien.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ đảng đã chọn không?" + vbCrLf + "Lưu ý: Khi thực hiện xoá hồ sơ đảng thì thông tin về quá trình sinh hoạt đảng của hồ sơ này cũng sẽ được xoá theo.", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                'Thực hiện xoá thông tin hồ sơ đảng và hồ sơ sinh hoạt đảng của Cán bộ
                                _HS_DangDT.Delete_PartyDocument(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            ResetAll_Controls(1)
                            dgv_dangvien.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_dangvien_CellClick(sender, Nothing)
                        Else
                            _count = 0
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_dangvien, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                Case "tp_hs_doan"       'THỰC HIỆN XOÁ HỒ SƠ ĐOÀN CÙNG QUÁ TRÌNH SINH HOẠT ĐOÀN
                    If (dgv_doanvien.Rows.Count <= 0) Then Return
                    If (dgv_doanvien.Rows.Count > 0) Then
                        If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Int32 = 0 To dgv_doanvien.Rows.Count - 1
                                If (dgv_doanvien.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_doanvien.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_doanvien.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ đoàn đã chọn không?" + vbCrLf + "Lưu ý: Khi thực hiện xoá hồ sơ đoàn thì thông tin về quá trình sinh hoạt đoàn của hồ sơ này cũng sẽ được xoá theo.", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                'Thực hiện xoá thông tin hồ sơ đảng và hồ sơ sinh hoạt đảng của Cán bộ
                                _HS_DangDT.Delete_UnionMember(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            ResetAll_Controls(2)
                            dgv_doanvien.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_doanvien_CellClick(sender, Nothing)
                        Else
                            _count = 0
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_doanvien, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
                Case "tp_hs_congdoan"   'THỰC HIỆN XOÁ HỒ SƠ CÔNG ĐOÀN CÙNG QUÁ TRÌNH SINH HOẠT CÔNG ĐOÀN
                    If (dgv_congdoan.Rows.Count <= 0) Then Return
                    If (dgv_congdoan.Rows.Count > 0) Then
                        If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                            For i As Int32 = 0 To dgv_congdoan.Rows.Count - 1
                                If (dgv_congdoan.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                                    If (CType(dgv_congdoan.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                        _count += 1
                                        arr_Del.Add(dgv_congdoan.Rows(i).Cells("cln_Code").Value.ToString())
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If (_count > 0) Then
                        Dim _mess As String = IIf(_count = 1, "", "các ")
                        If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi hồ sơ công đoàn đã chọn không?" + vbCrLf + "Lưu ý: Khi thực hiện xoá hồ sơ công đoàn thì thông tin về quá trình sinh hoạt đoàn của hồ sơ này cũng sẽ được xoá theo.", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                            'Thực hiện xoá dữ liệu khi đã chấp thuận
                            For i As Int16 = 0 To arr_Del.Count - 1
                                'Thực hiện xoá thông tin hồ sơ đảng và hồ sơ sinh hoạt đảng của Cán bộ
                                _HS_DangDT.Delete_TradeUnion(arr_Del(i).ToString())
                            Next
                            'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                            ResetAll_Controls(3)
                            dgv_congdoan.Rows.Clear()
                            'Select item của cây dữ liệu
                            tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                            Dim node As TreeNode = Nothing
                            If (strNode <> "") Then
                                node = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                                If Not (node Is Nothing) Then
                                    tv_main.SelectedNode = node
                                End If
                            End If
                            dgv_congdoan_CellClick(sender, Nothing)
                        Else
                            _count = 0
                            arr_Del.Clear()
                            Globals.Check_All_Items(dgv_congdoan, False)
                        End If
                    Else
                        MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    End If
            End Select
        Catch ex As Exception
            MessageBox.Show("Xoá thông tin hồ sơ " + IIf(tctrl_main.SelectedTab.Name = "tp_hs_dang", "Đảng", IIf(tctrl_main.SelectedTab.Name = "tp_hs_doan", "Đoàn", "Công đoàn")) + " của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        Dim _currRow As String = _RowId
        Select Case tctrl_main.SelectedTab.Name
            Case "tp_hs_dang"
                ResetAll_Controls(1)
                If (dgv_dangvien.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_dangvien.CurrentRow.Selected = False
                dgv_dangvien.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_dangvien, "cln_Code")).Selected = True
                If (dgv_dangvien.Rows.Count > 0) Then
                    If ((dgv_dangvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_dangvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow, 1)
                    End If
                End If
                If (dtpk_dang_ngayra.Checked = False) Then
                    edt_dang_lydo.ReadOnly = True
                Else
                    edt_dang_lydo.ReadOnly = False
                End If
            Case "tp_hs_doan"
                ResetAll_Controls(2)
                If (dgv_doanvien.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_doanvien.CurrentRow.Selected = False
                dgv_doanvien.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_doanvien, "cln_Code")).Selected = True
                If (dgv_doanvien.Rows.Count > 0) Then
                    If ((dgv_doanvien.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_doanvien.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow, 2)
                    End If
                End If
                If (dtpk_doan_ngayra.Checked = False) Then
                    edt_doan_lydo.ReadOnly = True
                Else
                    edt_doan_lydo.ReadOnly = False
                End If
            Case "tp_hs_congdoan"
                ResetAll_Controls(3)
                If (dgv_congdoan.Rows.Count <= 0) Then Return
                If (_currRow = "") Then _currRow = dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString()
                dgv_congdoan.CurrentRow.Selected = False
                dgv_congdoan.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_congdoan, "cln_Code")).Selected = True
                If (dgv_congdoan.Rows.Count > 0) Then
                    If ((dgv_congdoan.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_congdoan.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow, 3)
                    End If
                End If
                If (dtpk_congdoan_ngayra.Checked = False) Then
                    edt_congdoan_lydo.ReadOnly = True
                Else
                    edt_congdoan_lydo.ReadOnly = False
                End If
        End Select
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Close()
    End Sub
#End Region

#Region "---> Events: Các sự kiện ngoại lệ của người dùng <---"
    'Ngoại lệ: Hồ sơ Đảng viên của Cán bộ
    Private Sub dtpk_dang_ngaykn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_dang_ngaykn.ValueChanged
        Dim _FlagTS As Boolean = False
        Dim strSQL As String = String.Format("Select * From HSCB_TS Where idNew = '{0}' And TrangThai = 1", _IdCanBo)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _FlagTS = True
                End If
            End If
        End Using
        If (_RowId = "" And _FlagTS = False) Then
            dtpk_dang_ngayvao.Value = dtpk_dang_ngaykn.Value.AddYears(1)
        End If
    End Sub

    Private Sub dtpk_dang_ngayra_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_dang_ngayra.Leave
        If (dtpk_dang_ngayra.Checked = False) Then
            edt_dang_lydo.Text = ""
            edt_dang_lydo.ReadOnly = True
        Else
            edt_dang_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dtpk_dang_ngayra_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtpk_dang_ngayra.MouseUp
        If (dtpk_dang_ngayra.Checked = False) Then
            edt_dang_lydo.Text = ""
            edt_dang_lydo.ReadOnly = True
        Else
            edt_dang_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dtpk_doan_ngayra_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_doan_ngayra.Leave
        If (dtpk_doan_ngayra.Checked = False) Then
            edt_doan_lydo.Text = ""
            edt_doan_lydo.ReadOnly = True
        Else
            edt_doan_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dtpk_doan_ngayra_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtpk_doan_ngayra.MouseUp
        If (dtpk_doan_ngayra.Checked = False) Then
            edt_doan_lydo.Text = ""
            edt_doan_lydo.ReadOnly = True
        Else
            edt_doan_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dtpk_congdoan_ngayra_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpk_congdoan_ngayra.Leave
        If (dtpk_congdoan_ngayra.Checked = False) Then
            edt_congdoan_lydo.Text = ""
            edt_congdoan_lydo.ReadOnly = True
        Else
            edt_congdoan_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub dtpk_congdoan_ngayra_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtpk_congdoan_ngayra.MouseUp
        If (dtpk_congdoan_ngayra.Checked = False) Then
            edt_congdoan_lydo.Text = ""
            edt_congdoan_lydo.ReadOnly = True
        Else
            edt_congdoan_lydo.ReadOnly = False
        End If
    End Sub

    Private Sub edt_dang_sothe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_sothe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_dang_sothe.Text.Trim() <> "") Then
                dtpk_dang_ngaykn.Focus()
            Else
                edt_dang_sothe.Focus()
            End If
        End If
    End Sub

    Private Sub dtpk_dang_ngaykn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_dang_ngaykn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_dang_ngayvao.Focus()
        End If
    End Sub

    Private Sub dtpk_dang_ngayvao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_dang_ngayvao.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_dang_nguoi_gt.Focus()
        End If
    End Sub

    Private Sub edt_dang_nguoi_gt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_nguoi_gt.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_dang_nguoi_gt.Text.Trim() <> "") Then
                edt_dang_noi_kn.Focus()
            Else
                edt_dang_nguoi_gt.Focus()
            End If
        End If
    End Sub

    Private Sub edt_dang_nguoi_gt_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_nguoi_gt.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_dang_ngayvao.Focus()
        End If
    End Sub

    Private Sub edt_dang_noi_kn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_noi_kn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_dang_noi_kn.Text.Trim() <> "") Then
                edt_dang_noicapthe.Focus()
            Else
                edt_dang_noi_kn.Focus()
            End If
        End If
    End Sub

    Private Sub edt_dang_noi_kn_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_noi_kn.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dang_nguoi_gt.Focus()
        End If
    End Sub

    Private Sub edt_dang_noi_kn_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_dang_noi_kn.TextChanged
        If (_RowId = "") Then edt_dang_noicapthe.Text = edt_dang_noi_kn.Text
    End Sub

    Private Sub edt_dang_noicapthe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_noicapthe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_dang_ngayra.Focus()
        End If
    End Sub

    Private Sub edt_dang_noicapthe_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_noicapthe.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_dang_noi_kn.Focus()
        End If
    End Sub

    Private Sub dtpk_dang_ngayra_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_dang_ngayra.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (dtpk_dang_ngayra.Checked = True) Then
                edt_dang_lydo.Focus()
            Else
                btn_accept.Focus()
            End If
        End If
    End Sub

    Private Sub edt_dang_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_dang_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_dang_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then dtpk_dang_ngayra.Focus()
    End Sub

    'Ngoại lệ: Hồ sơ đoàn viên của Cán bộ
    Private Sub edt_doan_sothe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_sothe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_doan_sothe.Text.Trim() <> "") Then
                dtpk_doan_ngayvao.Focus()
            Else
                edt_doan_sothe.Focus()
            End If
        End If
    End Sub

    Private Sub dtpk_doan_ngayvao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_doan_ngayvao.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_doan_noi_kn.Focus()
        End If
    End Sub

    Private Sub edt_doan_noi_kn_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_noi_kn.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_doan_noicapthe.Focus()
        End If
    End Sub

    Private Sub edt_doan_noi_kn_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_noi_kn.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_doan_ngayvao.Focus()
        End If
    End Sub

    Private Sub edt_doan_noicapthe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_noicapthe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_doan_ngayra.Focus()
        End If
    End Sub

    Private Sub edt_doan_noicapthe_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_noicapthe.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_doan_noi_kn.Focus()
        End If
    End Sub

    Private Sub dtpk_doan_ngayra_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_doan_ngayra.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (dtpk_doan_ngayra.Checked = True) Then
                edt_doan_lydo.Focus()
            Else
                btn_accept.Focus()
            End If
        End If
    End Sub

    Private Sub edt_doan_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_doan_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_doan_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_doan_ngayra.Focus()
        End If
    End Sub

    'Ngoại lệ: Hồ sơ Công đoàn của Cán bộ
    Private Sub edt_congdoan_sothe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_sothe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_congdoan_sothe.Text.Trim() <> "") Then
                dtpk_congdoan_ngayvao.Focus()
            Else
                edt_congdoan_sothe.Focus()
            End If
        End If
    End Sub

    Private Sub dtpk_congdoan_ngayvao_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_congdoan_ngayvao.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_congdoan_coquan.Focus()
        End If
    End Sub

    Private Sub edt_congdoan_coquan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_coquan.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_congdoan_noicapthe.Focus()
        End If
    End Sub

    Private Sub edt_congdoan_coquan_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_coquan.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_congdoan_ngayvao.Focus()
        End If
    End Sub

    Private Sub edt_congdoan_noicapthe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_noicapthe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            dtpk_congdoan_ngayra.Focus()
        End If
    End Sub

    Private Sub edt_congdoan_noicapthe_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_noicapthe.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_congdoan_coquan.Focus()
        End If
    End Sub

    Private Sub dtpk_congdoan_ngayra_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_congdoan_ngayra.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            If (dtpk_congdoan_ngayra.Checked = True) Then
                edt_congdoan_lydo.Focus()
            Else
                btn_accept.Focus()
            End If
        End If
    End Sub

    Private Sub edt_congdoan_lydo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_lydo.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_congdoan_lydo_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_congdoan_lydo.KeyUp
        If (e.KeyCode = Keys.Up) Then
            dtpk_congdoan_ngayra.Focus()
        End If
    End Sub
#End Region

End Class