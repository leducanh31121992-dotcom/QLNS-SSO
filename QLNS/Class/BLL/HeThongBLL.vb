Imports System.Data
Imports System.Data.SqlClient
Public Class HeThongBLL

    Private Sub FixGrid_Columns(ByVal dgv_name As DataGridView)
        For Each _column As DataGridViewColumn In dgv_name.Columns
            _column.ReadOnly = True
            _column.SortMode = DataGridViewColumnSortMode.NotSortable
            Select Case _column.HeaderText.Trim()
                Case "STT"
                    _column.HeaderText = "STT"
                    _column.Width = 40
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                    _column.Frozen = True

                Case "HoTen"
                    _column.HeaderText = "Đơn vị/Phòng ban/Họ tên"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                    _column.Frozen = True

                Case "MaCB"
                    _column.HeaderText = "Mã CB"
                    _column.Width = 60
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NgaySinh"
                    _column.HeaderText = "Ngày sinh"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "GioiTinh"
                    _column.HeaderText = "Giới tính"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_So"
                    _column.HeaderText = "Số CMND/Thẻ căn cước"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_NgayCap"
                    _column.HeaderText = "Ngày cấp"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "CMT_NoiCap"
                    _column.HeaderText = "Nơi cấp"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NS_DiaChi"
                    _column.HeaderText = "Nơi sinh"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NQ_DiaChi"
                    _column.HeaderText = "Nguyên quán"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "TTr_DiaChi"
                    _column.HeaderText = "Thường trú/Tạm trú"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "ThT_DiaChi"
                    _column.HeaderText = "Địa chỉ theo hộ khẩu"
                    _column.Width = 190
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DanToc"
                    _column.HeaderText = "Dân tộc"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "TonGiao"
                    _column.HeaderText = "Tôn giáo"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "TrinhDoCT"
                    _column.HeaderText = "Trình độ chính trị"
                    _column.Width = 70
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DienThoai_DD"
                    _column.HeaderText = "Số điện thoại"
                    _column.Width = 90
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "Email"
                    _column.HeaderText = "Địa chỉ e-mail"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "QDNS_CV_PB_ChMon"
                    _column.HeaderText = "Chức danh/Chuyên môn đảm nhận"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "CVTruocTuyenDung"
                    _column.HeaderText = "Công việc trước khi Tuyển dụng/Tiếp nhận"
                    _column.Width = 120
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "NgayTiepNhan"
                    _column.HeaderText = "Ngày tiếp nhận"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "NgayTuyenDung"
                    _column.HeaderText = "Ngày tuyển dụng"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DVTuyenDung"
                    _column.HeaderText = "Đơn vị tuyển dụng"
                    _column.Width = 120
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "QDNS_Ngay_BN_BNL"
                    _column.HeaderText = "Ngày Bổ nhiệm/Bổ nhiệm lại"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DV_NgayCThuc"
                    _column.HeaderText = "Ngày vào đảng chính thức"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "DV_SoThe"
                    _column.HeaderText = "Số thẻ đảng"
                    _column.Width = 70
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DV_NoiCapThe"
                    _column.HeaderText = "Nơi cấp thẻ đảng"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_TrinhDo"
                    _column.HeaderText = "Cấp đào tạo"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_CoSoDT"
                    _column.HeaderText = "Trường đào tạo"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_HeDT"
                    _column.HeaderText = "Hệ đào tạo"
                    _column.Width = 80
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_ChuyenNganh"
                    _column.HeaderText = "Chuyên ngành đào tạo"
                    _column.Width = 130
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "DT_NamTN"
                    _column.HeaderText = "Năm tốt nghiệp"
                    _column.Width = 45
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "LCB_NgachLuong"
                    _column.HeaderText = "Ngạch lương"
                    _column.Width = 150
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "LCB_BacLuong"
                    _column.HeaderText = "Bậc lương"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "LCB_HeSoLuong"
                    _column.HeaderText = "Hệ số lương/Tiền lương"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "LCB_NgayHuong"
                    _column.HeaderText = "Ngày hưởng lương"
                    _column.Width = 77
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "LCB_NgayLenDK"
                    _column.HeaderText = "Ngày nâng lương dự kiến"
                    _column.Width = 77
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "PCCV"
                    _column.HeaderText = "Phụ cấp chức vụ"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCKV"
                    _column.HeaderText = "Phụ cấp khu vực"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCTN"
                    _column.HeaderText = "Phụ cấp trách nhiệm"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCDH"
                    _column.HeaderText = "Phụ cấp độc hại"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCTH"
                    _column.HeaderText = "Phụ cấp thu hút"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "PCDang"
                    _column.HeaderText = "Phụ cấp đảng"
                    _column.Width = 50
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                Case "HDLD_LoaiHD"
                    _column.HeaderText = "Loại HĐLĐ"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "HDLD_So"
                    _column.HeaderText = "Số HĐLĐ"
                    _column.Width = 110
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case "HDLD_NgayHL"
                    _column.HeaderText = "Ngày hiệu lực HĐLĐ"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HDLD_NgayKy"
                    _column.HeaderText = "Ngày ký HĐLĐ"
                    _column.Width = 75
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                Case "HDLD_DVKyHD"
                    _column.HeaderText = "Đơn vị ký HĐLĐ"
                    _column.Width = 180
                    _column.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
                Case Else
                    _column.Visible = False
            End Select
        Next
    End Sub

    ''' <summary>
    ''' Hàm thực hiện tìm kiếm danh sách Bảng dữ liệu tham số hệ thống
    ''' </summary>
    ''' <param name="pSysVarId"></param>
    ''' <param name="pDonViId"></param>
    ''' <param name="pDonVi_Cd"></param>
    ''' <param name="pDonVi_HT"></param>
    ''' <param name="pNgayHL_BD">dd/MM/yyyy</param>
    ''' <param name="pNgayHL_KT">dd/MM/yyyy</param>
    ''' <param name="pTrangThai"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListSysVarSearch(ByVal pSysVarId As Integer, ByVal pDonViId As Integer, ByVal pDonVi_Cd As String, ByVal pDonVi_HT As String, ByVal pNgayHL_BD As String, ByVal pNgayHL_KT As String, ByVal pTrangThai As Byte) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("SysVar_GetSearch", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pSysVarId", pSysVarId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViId", pDonViId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonVi_Cd", pDonVi_Cd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonVi_HT", pDonVi_HT, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNgayHL_BD", pNgayHL_BD, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNgayHL_KT", pNgayHL_KT, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pTrangThai", pTrangThai, ParameterDirection.Input))
            
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "SysVar")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("SysVar")
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

    Public Function Insert_Update_SysVar(ByVal _SysVar As HeThong.SysVar) As Integer
        Dim _Ret As Integer = 0
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "SysVar_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.Int).Direction = ParameterDirection.Output

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Id", _SysVar.Id, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Id_DonVi", _SysVar.Id_DonVi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TrucThuoc", _SysVar.TrucThuoc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DiaBan", _SysVar.DiaBan, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ALL", _SysVar.ALL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Max_Ky", _SysVar.Max_Ky, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tinh_Thue_TNCN", _SysVar.Tinh_Thue_TNCN, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo_GiamDoc", _SysVar.IdCanBo_GiamDoc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GiamDoc", _SysVar.GiamDoc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo_PhoGD", _SysVar.IdCanBo_PhoGD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhoGiamDoc", _SysVar.PhoGiamDoc, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo_HCTC", _SysVar.IdCanBo_HCTC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TruongHCTC", _SysVar.TruongHCTC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo_KeToan", _SysVar.IdCanBo_KeToan, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_KeToanTruong", _SysVar.KeToanTruong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Ten_VT", _SysVar.Ten_VT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DonVi_Cd", _SysVar.DonVi_Cd, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_VungLuongTT", _SysVar.VungLuongTT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_HeSoTamUng_V2", _SysVar.HeSoTamUng_V2, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_MucTamUng_V2", _SysVar.MucTamUng_V2, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgayHL", _SysVar.NgayHL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgayHetHL", _SysVar.NgayHetHL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TrangThai", _SysVar.TrangThai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CreatedBy", _SysVar.CreatedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _SysVar.ModifiedBy, ParameterDirection.Input))
                    
                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _SysVar.IdOutPut = CType(_command.Parameters("@_IdOutPut").Value.ToString(), Integer)
                        _Ret = _SysVar.IdOutPut
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Cập nhật tham số hệ thông (SYSVAR)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = 0
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    ''' <summary>
    ''' Hàm xóa/đánh dấu xóa bản ghi tham số hệ thống
    ''' </summary>
    ''' <param name="_Id">Chỉ số bản ghi cần xóa</param>
    ''' <param name="_ModifiedBy">Người dùng thực hiện</param>
    ''' <param name="_FlagDelete">Cờ xác định thực hiện: 1 - Xóa hẳn dữ liệu; 2 - Đánh dấu xóa bản ghi</param>
    ''' <returns>True - Thành công; False - Thất bại</returns>
    ''' <remarks></remarks>
    Public Function Delete_SysVar(ByVal _Id As Integer, ByVal _ModifiedBy As String, _FlagDelete As Byte) As Boolean
        Dim _Ret As Boolean = False

        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "SysVar_Delete"

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Id", _Id, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _ModifiedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_FlagDelete", _FlagDelete, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Xóa Tham số hệ thống (SYSVAR)", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = False
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    Public Shared Function GetQuery_ListCanBos(ByVal _BranchId As Integer, ByVal _NgayDL As String, ByVal _ChucVu_Cds As String, ByVal _PhongBan_Cds As String)
        Dim sRetQuery As String = ""
        sRetQuery = "Select T1.IdCanBo Code,(Case T1.GioiTinh When 0 Then N'Ông' Else N'Bà' End) + N' ' + T1.HoTen + N', '+dbo.GetTenDMuc(T.IdChucVu_Moi,1) + N' (Ngày sinh: '+ Convert(Varchar(10),T1.NgaySinh,103) + N', Mã CB: ' + T1.MaCB + N')' Name, "
        sRetQuery = sRetQuery + " dbo.GetTenDMuc(T.IdPhong_Moi,6) PhongBan_Cd,dbo.GetTenDMuc(T.IdPhong_Moi,3) PhongBan_HT,"
        sRetQuery = sRetQuery + " dbo.GetTenDMuc(T.IdChucVu_Moi,4) ChucVu_Cd,dbo.GetTenDMuc(T.IdChucVu_Moi,1) ChucVu_HT"
        sRetQuery = sRetQuery + " From QDNhanSu T, HS_Canbo T1,"
        sRetQuery = sRetQuery + " (Select IdCanBo, Max(NgayHL) As NgayHL From QDNhanSu Where IsKiemNhiem = 0 And IsQD_NHCS = 1 And Cast(NgayHL As Date) <= Convert(Date,'" + _NgayDL + "',103)"
        sRetQuery = sRetQuery + "         Group By IdCanBo Having Max(Cast(NgayHL As Date)) <= Convert(Date,'" + _NgayDL + "',103)"
        sRetQuery = sRetQuery + " ) T2,(Select * From ChiNhanh Where Status=1) CN"
        sRetQuery = sRetQuery + " Where T.IdCanBo = T1.IdCanBo And T.IdCanBo = T2.IdCanBo And T.NgayHL = T2.NgayHL"
        sRetQuery = sRetQuery + " And CN.Id = T.IdDonVi_Moi "
        If _BranchId <> 0 Then
            sRetQuery = sRetQuery + " And T.IdDonvi_Moi = " + _BranchId.ToString()
        End If
        If _ChucVu_Cds <> "" Then
            sRetQuery = sRetQuery + " And T.IdPhong_Moi In (Select X.Id From PhongBan X Where X.Status=1 And X.Ma_So In (" + _PhongBan_Cds + "))"
        End If
        If _PhongBan_Cds <> "" Then
            sRetQuery = sRetQuery + " And T.IdChucVu_Moi In (Select X.Id From DanhMuc X Where X.Status=1 And X.Ma_So In (" + _ChucVu_Cds + "))"
        End If
        sRetQuery = sRetQuery + " And T1.IdCanBo Not In "
        sRetQuery = sRetQuery + " ("
        sRetQuery = sRetQuery + " Select T4.IdCanBo From "
        sRetQuery = sRetQuery + "    ("
        sRetQuery = sRetQuery + "     Select IdCanBo, Max(Ngay_HL) As Ngay_HL From HS_CBThoiviec Where IsQD_NHCS = 1 And Cast(Ngay_HL As Date) <= Convert(DateTime,'" + _NgayDL + "',103)"
        sRetQuery = sRetQuery + "            Group By IdCanBo Having Max(Cast(Ngay_HL As Date)) <= Convert(DateTime,'" + _NgayDL + "',103)"
        sRetQuery = sRetQuery + "    ) T4 Where T4.IdCanBo= T.IdCanBo And DateAdd(Second, 10, T4.Ngay_HL) > Cast(T.NgayHL As Date)"
        sRetQuery = sRetQuery + " )"
        sRetQuery = sRetQuery + " Order By dbo.GetTenDMuc(T.IdPhong_Moi,6),(Case When dbo.GetTenDMuc(T.IdChucVu_Moi,4) In ('1444','1426') Then 1 Else 0 End),dbo.GetTenDMuc(T.IdChucVu_Moi,4)"


        Return sRetQuery
    End Function

    ''' <summary>
    ''' Hàm lấy danh sách Fill dữ liệu vào Treeview toàn bộ các Node
    ''' </summary>
    ''' <param name="pFlagCall"> Chỉ số xác định kiểu Danh mục cho Treeview. Giá trị quy ước:
    '''                                1 -  Danh mục các Chi nhánh => Xong đến Phòng ban trực thuộc (Chỉ lấy Ban CMNV) => Tháng có phát sinh chi lương
    ''' </param>
    ''' <param name="pCapBC">Cấp BC: 1: Trung ương; 2: Chi nhánh; 3 - Phòng giao dịch</param>
    ''' <param name="pNamBC">Năm báo cáo</param>
    ''' <param name="pDonviCd">Mã Chi nhánh/PGD</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDanhMuc_TreeViews(ByVal pFlagCall As Byte, ByVal pCapBC As Byte, ByVal pNamBC As Integer, ByVal pThangBC As Integer, ByVal pDonviCd As String) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("GetDanhMuc_TreeViews", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pFlagCall", pFlagCall, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pCapBC", pCapBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNamBC", pNamBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThangBC", pThangBC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonviCd", pDonviCd, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "DanhMuc_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("DanhMuc_TMP")
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
End Class
