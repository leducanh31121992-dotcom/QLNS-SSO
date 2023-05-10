
Imports System.Data
Imports System.Data.SqlClient
''' <summary>
''' Các hàm liên quan đến báo cáo - Write by Duong Van Chu - 09/12/2008
''' </summary>
''' <remarks></remarks>
Public Class clsHS_BaoCao
    Private _SqlHelper As DBAccess

    Public Sub New()
        _SqlHelper = New DBAccess
    End Sub

    ''' <summary>
    ''' Hàm trả về ngày khi biết Quý và năm truyền vào
    ''' </summary>
    ''' <param name="_Quarter">Quý truyền vào</param>
    ''' <param name="_Year">Năm truyền vào</param>
    ''' <param name="_flag">True: Lấy ngày đầu. False: Lấy ngày cuối</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDate(ByVal _Quarter As Byte, ByVal _Year As Integer, ByVal _flag As Boolean) As DateTime
        Dim _dateResult As DateTime = New DateTime()
        If (_flag = True) Then  'Ngày đầu
            If _Quarter <> 5 And _Quarter <> 6 Then
                _dateResult = New DateTime(_Year, _Quarter * 3 - 2, 1)
            Else
                _dateResult = New DateTime(_Year, 1, 1)
            End If
        Else        'Ngày cuối
            '_dateResult = New DateTime(_Year, _Quarter * 3, 30)
            If _Quarter <> 5 And _Quarter <> 6 Then
                _dateResult = New DateTime(_Year, _Quarter * 3, 1)
                _dateResult = _dateResult.AddMonths(1)
                _dateResult = _dateResult.AddDays(-1)
            Else
                If _Quarter = 5 Then
                    _dateResult = New DateTime(_Year, 6, 30)
                Else
                    _dateResult = New DateTime(_Year, 12, 31)
                End If
            End If
        End If
        Return _dateResult
    End Function

    ''' <summary>
    ''' Hàm lấy dữ liệu cho báo cáo - Danh sách tuyển dụng - Tiếp nhận cán bộ
    ''' </summary>
    ''' <param name="_ChiNhanhId">Id chi nhánh</param>
    ''' <param name="_Year">Năm</param>
    ''' <param name="_KyBC">Kỳ hạn báo cáo</param>
    ''' <returns>Table dữ liệu cần dùng</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Report04(ByVal _ChiNhanhId As Integer, ByVal _Year As Integer, ByVal _KyBC As Byte, ByVal _FlagALL As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As System.Data.SqlClient.SqlCommand = Nothing
            If _FlagALL = 0 Then    'Lấy dữ liệu toàn quốc
                command = New System.Data.SqlClient.SqlCommand("BC04TQ", connection)
                command.CommandTimeout = 150
            Else                    'Lấy dữ liệu đơn vị
                command = New System.Data.SqlClient.SqlCommand("BC04DV", connection)
            End If
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BranchId", SqlDbType.Int))
            command.Parameters("@_BranchId").Value = _ChiNhanhId

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_NamBC", SqlDbType.Int))
            command.Parameters("@_NamBC").Value = _Year

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_KyBC", SqlDbType.Int))
            command.Parameters("@_KyBC").Value = _KyBC

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC_04")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC_04")
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
    ''' 02/TCCB: Hàm lấy dữ liệu cho báo cáo - Tình hình thực hiện công tác bổ nhiệm Lãnh đạo, Quản lý
    ''' </summary>
    ''' <param name="_ChiNhanhId">Chỉ số xác định chi nhánh. Nếu là 0 thì lấy tất cả toàn quốc</param>
    ''' <param name="_NgayBD">Ngày bắt đầu lấy số liệu. Nếu lấy SL đến thời điểm hiện tại thì mặc định giá trị '1900-01-01'</param>
    ''' <param name="_NgayKT">Ngày kết thúc lấy số liệu</param>
    ''' <param name="_State">Chỉ số xác định lấy DL Báo cáo: 1,2,3,4 - Đến thời điểm hiện tại (Ngày cuối quý); 5 - Lấy từ ngày đến ngày</param>
    ''' <param name="_FlagALL">1 - Tổng hợp đến Chi nhánh; 0 - Chi tiết cả PGD</param>
    ''' <param name="_FlagReportExcel">1 - In ra báo cáo; 2 - Xuất danh sách ra excel;</param>
    ''' <returns>DataTable dữ liệu báo cáo</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Report02(ByVal _ChiNhanhId As Integer, ByVal _NgayBD As DateTime, ByVal _NgayKT As DateTime, ByVal _State As Byte, ByVal _FlagALL As Byte, ByVal _FlagReportExcel As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As System.Data.SqlClient.SqlCommand = Nothing
            If _ChiNhanhId = 0 Then
                command = New System.Data.SqlClient.SqlCommand("BC02_ToanQuoc", connection)
            ElseIf _ChiNhanhId = 1 Then
                command = New System.Data.SqlClient.SqlCommand("BC02_HSC", connection)
            Else
                command = New System.Data.SqlClient.SqlCommand("BC02_ChiNhanh", connection)
            End If
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BranchId", SqlDbType.Int))
            command.Parameters("@_BranchId").Value = _ChiNhanhId
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_TuNgay", SqlDbType.DateTime))
            command.Parameters("@_TuNgay").Value = _NgayBD
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_DenNgay", SqlDbType.DateTime))
            command.Parameters("@_DenNgay").Value = _NgayKT
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagALL", SqlDbType.TinyInt))
            command.Parameters("@_FlagALL").Value = _FlagALL
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagReportExcel", SqlDbType.TinyInt))
            command.Parameters("@_FlagReportExcel").Value = _FlagReportExcel
            command.CommandTimeout = 90
            command.CommandType = CommandType.StoredProcedure

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC02")
                If (ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC02")
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

    Public Function GetAll_Report02_XoaDi(ByVal _ChiNhanhId As Integer, ByVal begin_date As DateTime, ByVal end_date As DateTime, ByVal _State As Byte, ByVal _FlagALL As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As System.Data.SqlClient.SqlCommand = Nothing
            If _State = 9 Then
                If _ChiNhanhId = 0 Then
                    command = New System.Data.SqlClient.SqlCommand("BC02_TH_CT_TQ", connection)
                ElseIf _ChiNhanhId = 1 Then
                    command = New System.Data.SqlClient.SqlCommand("BC02C_HSC_2019", connection)
                End If
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BranchId", SqlDbType.Int))
                command.Parameters("@_BranchId").Value = _ChiNhanhId
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_TuNgay", SqlDbType.DateTime))
                command.Parameters("@_TuNgay").Value = begin_date
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_DenNgay", SqlDbType.DateTime))
                command.Parameters("@_DenNgay").Value = end_date
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagALL", SqlDbType.TinyInt))
                command.Parameters("@_FlagALL").Value = _FlagALL
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagReportExcel", SqlDbType.TinyInt))
                command.Parameters("@_FlagReportExcel").Value = 1
                command.CommandTimeout = 90
                command.CommandType = CommandType.StoredProcedure
            Else

                If _State = 0 Then
                    If _ChiNhanhId = 1 Then
                        command = New System.Data.SqlClient.SqlCommand("BC02C_HSC_2019", connection)
                        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagReportExcel", SqlDbType.TinyInt))
                        command.Parameters("@_FlagReportExcel").Value = 1
                    Else
                        command = New System.Data.SqlClient.SqlCommand("BC02C_2019", connection)
                    End If
                Else
                    command = New System.Data.SqlClient.SqlCommand("BC02D_2019", connection)
                End If
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_TuNgay", SqlDbType.DateTime))
                command.Parameters("@_TuNgay").Value = begin_date
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_DenNgay", SqlDbType.DateTime))
                command.Parameters("@_DenNgay").Value = end_date

                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BranchId", SqlDbType.Int))
                command.Parameters("@_BranchId").Value = _ChiNhanhId
                If _State = 0 Then
                    command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagALL", SqlDbType.TinyInt))
                    command.Parameters("@_FlagALL").Value = _FlagALL
                End If
            End If

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC02")
                If (ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC02")
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
    ''' Hàm lấy dữ liệu cho báo cáo - Tình hình thực hiện lao động - Mạng lưới hoạt động
    ''' </summary>
    ''' <param name="_ChiNhanhId">Chi nhánh cần báo cáo</param>
    ''' <param name="_Year">Nam báo cáo</param>
    ''' <param name="_ThangT1">Tháng thứ nhất - trong quý</param>
    ''' <param name="_ThangT2">Tháng thứ hai - trong quý</param>
    ''' <param name="_ThangT3">Tháng thứ ba - trong quý</param>
    ''' <param name="_FlagALL">Chỉ số xác định lấy báo cáo đơn vị hay toàn quốc</param>
    ''' <returns>Table dư liệu báo cáo 01</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Report01(ByVal _ChiNhanhId As Integer, ByVal _Year As Integer, ByVal _ThangT1 As Integer, ByVal _ThangT2 As Integer, ByVal _ThangT3 As Integer, ByVal _FlagALL As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand("BC01", connection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BranchId", SqlDbType.Int))
            command.Parameters("@_BranchId").Value = _ChiNhanhId
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Nam", SqlDbType.Int))
            command.Parameters("@_Nam").Value = _Year
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_ThangT1", SqlDbType.Int))
            command.Parameters("@_ThangT1").Value = _ThangT1
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_ThangT2", SqlDbType.Int))
            command.Parameters("@_ThangT2").Value = _ThangT2
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_ThangT3", SqlDbType.Int))
            command.Parameters("@_ThangT3").Value = _ThangT3
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagALL", SqlDbType.TinyInt))
            command.Parameters("@_FlagALL").Value = _FlagALL
            command.CommandTimeout = 90
            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC_01")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC_01")
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
    ''' Hàm lấy dữ liệu báo cáo Thống kê Số lượng, chất lượng cán bộ 03/BC-TCCB
    ''' </summary>
    ''' <param name="pDonViId">Chỉ số xác định đơn vị cần lấy Số liệu. 0 - Lất tất cả toàn quốc</param>
    ''' <param name="pFlagCall">1 - Báo cáo 03/TCCB Report; 2 - Báo cáo 03/TCCB xuất excel; 9 - Danh sách chi tiết cán bộ ra excel;</param>
    ''' <param name="pNgayBC">Ngày báo cáo. Kiểu DateTime</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll_Report03(ByVal pDonViId As Integer, ByVal pFlagCall As Byte, ByVal pNgayBC As DateTime) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As System.Data.SqlClient.SqlCommand = Nothing
            command = New System.Data.SqlClient.SqlCommand("BC03_2022", connection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pDonViId", SqlDbType.Int))
            command.Parameters("@pDonViId").Value = pDonViId
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pThoiDiem", SqlDbType.VarChar))
            command.Parameters("@pThoiDiem").Value = pNgayBC.ToString("yyyy-MM-dd")

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pFlagCall", SqlDbType.TinyInt))
            command.Parameters("@pFlagCall").Value = pFlagCall
            'If pDonViId = 0 Then command.CommandTimeout = 480

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC02")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC02")
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

    Public Function GetAll_Report03_KhongDung_Xoa(ByVal _ChiNhanhId As Integer, ByVal _Code As String, ByVal _DateEnd As DateTime) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As System.Data.SqlClient.SqlCommand = Nothing
            command = New System.Data.SqlClient.SqlCommand("BC03_TCCB", connection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pUnitId", SqlDbType.Int))
            command.Parameters("@pUnitId").Value = _ChiNhanhId
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pThoiDiem", SqlDbType.VarChar))
            command.Parameters("@pThoiDiem").Value = _DateEnd.ToString("dd/MM/yyyy")

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pFlagALL", SqlDbType.TinyInt))
            command.Parameters("@pFlagALL").Value = 1
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pFlagReportExcel", SqlDbType.TinyInt))
            command.Parameters("@pFlagReportExcel").Value = 1
            If _ChiNhanhId = 0 Then command.CommandTimeout = 480

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC02")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC02")
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
    ''' Bao cao ngoai nganh
    ''' </summary>
    ''' <param name="_ChiNhanhId"></param>
    ''' <param name="_Year"></param>
    ''' <param name="_ThangT1"></param>
    ''' <param name="_ThangT2"></param>
    ''' <param name="_ThangT3"></param>
    ''' <param name="_FlagALL"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll_Report01_NN(ByVal _ChiNhanhId As Integer, ByVal _Year As Integer, ByVal _ThangT1 As Integer, ByVal _ThangT2 As Integer, ByVal _ThangT3 As Integer, ByVal _FlagALL As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand("BC01_NN", connection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BranchId", SqlDbType.Int))
            command.Parameters("@_BranchId").Value = _ChiNhanhId
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Nam", SqlDbType.Int))
            command.Parameters("@_Nam").Value = _Year
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_ThangT1", SqlDbType.Int))
            command.Parameters("@_ThangT1").Value = _ThangT1
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_ThangT2", SqlDbType.Int))
            command.Parameters("@_ThangT2").Value = _ThangT2
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_ThangT3", SqlDbType.Int))
            command.Parameters("@_ThangT3").Value = _ThangT3
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_FlagALL", SqlDbType.TinyInt))
            command.Parameters("@_FlagALL").Value = _FlagALL
            command.CommandTimeout = 90
            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC_01NN")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC_01NN")
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

    '@_IdDonVi Int, @_NamBC Int, @_QuyBC TinyInt, @_IsFlagAll TinyInt, @_FlagReportExcel TinyInt


    Public Function GetDanhSach_NangLuong(ByVal _LoaiDSach As Byte, ByVal _IdDonVi As Integer, ByVal _NamBC As Integer, ByVal _QuyBC As Byte, ByVal _IsFlagAll As Byte, ByVal _FlagReportExcel As Byte) As DataTable
        Try
            Dim ds_ret As DataSet = New DataSet()
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                If (_LoaiDSach = 1 Or _LoaiDSach = 2) Then        'Danh sách nâng bậc lương
                    Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand("GetDanhSach_NangLuong", conn_obj)
                        _command.CommandType = CommandType.StoredProcedure

                        _command.Parameters.Add(New SqlParameter("@_IdDonVi", SqlDbType.Int))
                        _command.Parameters("@_IdDonVi").Value = _IdDonVi

                        _command.Parameters.Add(New SqlParameter("@_NamBC", SqlDbType.Int))
                        _command.Parameters("@_NamBC").Value = _NamBC

                        _command.Parameters.Add(New SqlParameter("@_QuyBC", SqlDbType.TinyInt))
                        _command.Parameters("@_QuyBC").Value = _QuyBC

                        _command.Parameters.Add(New SqlParameter("@_IsFlagAll", SqlDbType.TinyInt))
                        _command.Parameters("@_IsFlagAll").Value = _IsFlagAll

                        _command.Parameters.Add(New SqlParameter("@_FlagReportExcel", SqlDbType.TinyInt))
                        _command.Parameters("@_FlagReportExcel").Value = _FlagReportExcel
                        Using _sqldap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(_command)
                            _sqldap.Fill(ds_ret, "DanhSach_NangLuong")
                            If (ds_ret Is Nothing Or ds_ret.Tables.Count = 0 Or ds_ret.Tables(0).Rows.Count = 0) Then
                                Dim dt As DataTable
                                Return dt
                            End If
                        End Using
                    End Using

                End If
            End Using
            Return ds_ret.Tables(0)
        Catch ex As Exception
            MessageBox.Show("Lỗi truy vấn Danh sách cán bộ nâng bậc lương: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm lấy dữ liệu báo cáo 09/TKCD - Thống kê trình độ chuyên môn - Chuyên ngành theo chức danh
    ''' </summary>
    ''' <param name="pDonViId">Chỉ số xác định đơn vị cần lấy Số liệu. 0 - Lất tất cả toàn quốc</param>
    ''' <param name="pFlagCall">1 - Báo cáo 03/TKCD Report; 2 - Báo cáo 03/TKCD xuất excel; 9 - Danh sách chi tiết cán bộ ra excel;</param>
    ''' <param name="pNgayBC">Ngày báo cáo. Kiểu DateTime</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll_Report09_TKCD(ByVal pDonViId As Integer, ByVal pFlagCall As Byte, ByVal pNgayBC As DateTime) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As System.Data.SqlClient.SqlCommand = Nothing
            command = New System.Data.SqlClient.SqlCommand("BC09_V1", connection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pDonViId", SqlDbType.Int))
            command.Parameters("@pDonViId").Value = pDonViId
            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pThoiDiem", SqlDbType.VarChar))
            command.Parameters("@pThoiDiem").Value = pNgayBC.ToString("yyyy-MM-dd")

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@pFlagCall", SqlDbType.TinyInt))
            command.Parameters("@pFlagCall").Value = pFlagCall

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "BC09")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("BC09")
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
