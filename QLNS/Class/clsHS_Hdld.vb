Imports System.Data
Imports System.Data.SqlClient

Public Class clsHS_Hdld
    'aaaaaaaaaaaa
    Private _SqlHelper As DBAccess

    Public Sub New()
        _SqlHelper = New DBAccess
    End Sub

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Public Class HS_HDLD
        Private _IdCB_HDLD As String
        Private _IdCanBo As String
        Private _SoHD As String
        Private _IdLoaiHD As Int32
        Private _Ngay_HL As DateTime
        Private _NgayKy_HD As DateTime
        Private _NguoKy_QD As String
        Private _IdCV_Nguoiky_QD As Int32
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _IdHT_TraLuong As Int32
        Private _IdBacLuong As Int32
        Private _HeSo As Double
        Private _TyleHuong As Double
        Private _DVKyHDLD As String
        Private _NoiLamViec As String
        Private _GhiChu As String

        Public Sub New()
            _IdCB_HDLD = ""
            _IdCanBo = ""
            _SoHD = ""
            _IdLoaiHD = 0
            _Ngay_HL = DateTime.Now
            _NgayKy_HD = DateTime.Now
            _NguoKy_QD = ""
            _IdCV_Nguoiky_QD = 0
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _IdHT_TraLuong = 0
            _IdBacLuong = 0
            _HeSo = 0
            _TyleHuong = 0
            _DVKyHDLD = ""
            _NoiLamViec = ""
            _GhiChu = ""
        End Sub

        Public Property IdCB_HDLD() As String
            Get
                Return _IdCB_HDLD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hợp đồng lao động có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCB_HDLD = value
            End Set
        End Property

        Public Property IdCanBo() As String
            Get
                Return _IdCanBo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id cán bộ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCanBo = value
            End Set
        End Property

        Public Property SoHD() As String
            Get
                Return _SoHD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số hợp đồng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoHD = value
            End Set
        End Property

        Public Property IdLoaiHD() As Int32
            Get
                Return _IdLoaiHD
            End Get
            Set(ByVal value As Int32)
                _IdLoaiHD = value
            End Set
        End Property

        Public Property Ngay_HL() As DateTime
            Get
                Return _Ngay_HL
            End Get
            Set(ByVal value As DateTime)
                _Ngay_HL = value
            End Set
        End Property

        Public Property NgayKy_HD() As DateTime
            Get
                Return _NgayKy_HD
            End Get
            Set(ByVal value As DateTime)
                _NgayKy_HD = value
            End Set
        End Property

        Public Property NguoKy_QD() As String
            Get
                Return _NguoKy_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NguoKy_QD = value
            End Set
        End Property

        Public Property IdCV_Nguoiky_QD() As Int32
            Get
                Return _IdCV_Nguoiky_QD
            End Get
            Set(ByVal value As Int32)
                _IdCV_Nguoiky_QD = value
            End Set
        End Property

        Public Property TuNgay() As DateTime
            Get
                Return _TuNgay
            End Get
            Set(ByVal value As DateTime)
                _TuNgay = value
            End Set
        End Property

        Public Property DenNgay() As DateTime
            Get
                Return _DenNgay
            End Get
            Set(ByVal value As DateTime)
                _DenNgay = value
            End Set
        End Property

        Public Property IdHT_TraLuong() As Int32
            Get
                Return _IdHT_TraLuong
            End Get
            Set(ByVal value As Int32)
                _IdHT_TraLuong = value
            End Set
        End Property

        Public Property IdBacLuong() As Int32
            Get
                Return _IdBacLuong
            End Get
            Set(ByVal value As Int32)
                _IdBacLuong = value
            End Set
        End Property

        Public Property HeSo() As Double
            Get
                Return _HeSo
            End Get
            Set(ByVal value As Double)
                _HeSo = value
            End Set
        End Property

        Public Property TyleHuong() As Double
            Get
                Return _TyleHuong
            End Get
            Set(ByVal value As Double)
                _TyleHuong = value
            End Set
        End Property

        Public Property DVKyHDLD() As String
            Get
                Return _DVKyHDLD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị đơn vị ký HĐLĐ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DVKyHDLD = value
            End Set
        End Property

        Public Property NoiLamViec() As String
            Get
                Return _NoiLamViec
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi làm việc có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiLamViec = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property
    End Class

    Public Class HS_CBThoiViec
        Private _IdCBThoiViec As String
        Private _IdCanBo As String
        Private _IdLoaiQD As Int32
        Private _LyDo As String
        Private _NgayKy_QD As DateTime
        Private _Ngay_HL As DateTime
        Private _NguoiKy_QD As String
        Private _IdCV_NguoKy_QD As Int32
        Private _TroCap_ThoiViec As Double
        Private _TroCap_Khac As Double
        Private _GhiChu As String

        Public Sub New()
            _IdCBThoiViec = ""
            _IdCanBo = ""
            _IdLoaiQD = 0
            _LyDo = ""
            _NgayKy_QD = DateTime.Now
            _Ngay_HL = DateTime.Now
            _NguoiKy_QD = ""
            _IdCV_NguoKy_QD = 0
            _TroCap_ThoiViec = 0
            _TroCap_Khac = 0
            _GhiChu = ""
        End Sub

        Public Property IdCBThoiViec() As String
            Get
                Return _IdCBThoiViec
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ thôi việc có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCBThoiViec = value
            End Set
        End Property

        Public Property IdCanBo() As String
            Get
                Return _IdCanBo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ cán bộ thôi việc có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCanBo = value
            End Set
        End Property

        Public Property IdLoaiQD() As Int32
            Get
                Return _IdLoaiQD
            End Get
            Set(ByVal value As Int32)
                _IdLoaiQD = value
            End Set
        End Property

        Public Property LyDo() As String
            Get
                Return _LyDo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do thôi việc có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LyDo = value
            End Set
        End Property

        Public Property NgayKy_QD() As DateTime
            Get
                Return _NgayKy_QD
            End Get
            Set(ByVal value As DateTime)
                _NgayKy_QD = value
            End Set
        End Property

        Public Property Ngay_HL() As DateTime
            Get
                Return _Ngay_HL
            End Get
            Set(ByVal value As DateTime)
                _Ngay_HL = value
            End Set
        End Property

        Public Property NguoiKy_QD() As String
            Get
                Return _NguoiKy_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NguoiKy_QD = value
            End Set
        End Property

        Public Property IdCV_NguoKy_QD() As Int32
            Get
                Return _IdCV_NguoKy_QD
            End Get
            Set(ByVal value As Int32)
                _IdCV_NguoKy_QD = value
            End Set
        End Property

        Public Property TroCap_ThoiViec() As Double
            Get
                Return _TroCap_ThoiViec
            End Get
            Set(ByVal value As Double)
                _TroCap_ThoiViec = value
            End Set
        End Property

        Public Property TroCap_Khac() As Double
            Get
                Return _TroCap_Khac
            End Get
            Set(ByVal value As Double)
                _TroCap_Khac = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property
    End Class
#End Region

#Region "---> Hàm Định nghĩa lưới dữ liệu <---"
    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu - Hiên thị thông tin liên quan đến hồ sơ lao động
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="status">True: Hồ sơ hợp đồng lao động. False: Hồ sơ cán bộ nghỉ việc</param>
    ''' <remarks></remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal status As Boolean)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        Select Case status
            Case True       'Định nghĩa lưới dữ liệu - Hồ sơ Hợp đồng lao động
                'Add Cột chọn xoá - Checkeditems
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_SoHd", "Số hợp đồng")
                dgv_name.Columns.Add("cln_LoaiHd", "Loại hợp đồng")
                dgv_name.Columns.Add("cln_NgayHL", "Ngày hiệu lực")
                dgv_name.Columns.Add("cln_Ngayky", "Ngày ký")
                dgv_name.Columns.Add("cln_Nguoiky", "Người ký")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ người ký")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_HTTL", "Hình thức trả lương")
                dgv_name.Columns.Add("cln_Bacluong", "Bậc lương")
                dgv_name.Columns.Add("cln_Heso", "Hệ số")
                dgv_name.Columns.Add("cln_Tyle", "Tỷ lệ được hưởng (%)")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_SoHd").Width = 90
                dgv_name.Columns("cln_LoaiHd").Width = 150
                dgv_name.Columns("cln_NgayHL").Width = 88
                dgv_name.Columns("cln_Ngayky").Width = 85
                dgv_name.Columns("cln_Nguoiky").Width = 140
                dgv_name.Columns("cln_Chucvu").Width = 140
                dgv_name.Columns("cln_Tungay").Width = 85
                dgv_name.Columns("cln_Denngay").Width = 85
                dgv_name.Columns("cln_HTTL").Width = 120
                dgv_name.Columns("cln_Bacluong").Width = 80
                dgv_name.Columns("cln_Heso").Width = 75
                dgv_name.Columns("cln_Tyle").Width = 138
                dgv_name.Columns("cln_Ghichu").Width = 200
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 13
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_NgayHL").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Ngayky").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Heso").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Tyle").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Choice").Frozen = True
            Case False       'Định nghĩa lưới dữ liệu - Hồ sơ Cán bộ thôi việc
                'Add Cột chọn xoá - Checkeditems
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_LoaiQd", "Loại quyết định")
                dgv_name.Columns.Add("cln_Lydo", "Lý do nghỉ việc")
                dgv_name.Columns.Add("cln_NgayKy", "Ngày ký quyết định")
                dgv_name.Columns.Add("cln_NgayHl", "Ngày hiệu lực")
                dgv_name.Columns.Add("cln_NguoiKy", "Người ký")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ người ký")
                dgv_name.Columns.Add("cln_Trocap_Tviec", "Trợ cấp thôi việc")
                dgv_name.Columns.Add("cln_Trocap_Khac", "Trợ cấp khác")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(1).Width = 120
                dgv_name.Columns(2).Width = 170
                dgv_name.Columns(3).Width = 200
                dgv_name.Columns(4).Width = 125
                dgv_name.Columns(5).Width = 90
                dgv_name.Columns(6).Width = 170
                dgv_name.Columns(7).Width = 180
                dgv_name.Columns(8).Width = 130
                dgv_name.Columns(9).Width = 130
                dgv_name.Columns(10).Width = 200

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 9
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_NgayKy").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NgayHl").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Trocap_Tviec").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Trocap_Khac").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        End Select
        dgv_name.Columns("cln_Code").Visible = False
    End Sub
#End Region

#Region "---> Hàm lấy dữ liệu - liên quan đến Lao động <---"
    ''' <summary>
    ''' Hàm trả về danh sách bản ghi - Hợp đồng lao động theo Id cán bộ truyền vào
    ''' </summary>
    ''' <param name="_IdCanBo">Id cán bộ</param>
    ''' <returns>Danh sách hợp đồng lao động</returns>
    ''' <remarks></remarks>
    Public Function GetAllLaborContract(ByVal _IdCanBo As String) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = New SqlCommand("HS_HDLD_GetForCanBoId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
            command.Parameters("@_IdCanBo").Value = _IdCanBo

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_HDLD")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_HDLD")
            End Using
        Catch ex As Exception
            MessageBox.Show("Lấy dữ liệu hồ sơ hợp đồng lao động: " + ex.Message.ToString(), "Lỗi xẩy ra", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try

    End Function

    ''' <summary>
    ''' Hàm trả về row hồ sơ hợp đồng lao động - theo mã hiệu truyền vào
    ''' </summary>
    ''' <param name="_Code">Mã hiệu hồ sơ hợp đồng lao động</param>
    ''' <returns>Row cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetLaborContract(ByVal _Code As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HS_HDLD_GetForId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdCB_HDLD", SqlDbType.VarChar))
            command.Parameters("@_IdCB_HDLD").Value = _Code

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_HDLD")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_HDLD").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về danh sách bản ghi - Hồ sơ cán bộ thôi việc -> Theo Id cán bộ truyền vào
    ''' </summary>
    ''' <param name="_IdCanBo">Id cán bộ</param>
    ''' <returns>Danh sách cán bộ thôi việc</returns>
    ''' <remarks></remarks>
    Public Function GetAll_StopWork(ByVal _IdCanBo As String) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = New SqlCommand("HS_CBThoiViec_GetForCanBoId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
            command.Parameters("@_IdCanBo").Value = _IdCanBo

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_CBThoiViec")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_CBThoiViec")
            End Using
        Catch ex As Exception
            MessageBox.Show("Lấy dữ liệu hồ sơ cán bộ thôi việc: " + ex.Message.ToString(), "Lỗi xẩy ra", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try

    End Function

    ''' <summary>
    ''' Hàm trả về row hồ sơ cán bộ thôi việc - theo mã hiệu truyền vào
    ''' </summary>
    ''' <param name="_Code">Mã hiệu hồ sơ cán bộ thôi việc</param>
    ''' <returns>Row cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetStopWork(ByVal _Code As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = New SqlCommand("HS_CBThoiViec_GetForId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdCBThoiViec", SqlDbType.VarChar))
            command.Parameters("@_IdCBThoiViec").Value = _Code

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_CBThoiViec")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_CBThoiViec").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về Id Ngạch lương khi biết - Id của bậc lương
    ''' </summary>
    ''' <param name="_IdBacLuong">Id bậc lương</param>
    ''' <returns>Id ngạch lương cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetIdNgachLuong(ByVal _IdBacLuong As Int32) As Int32
        Try
            Dim _ResultId As Int32 = 0
            Dim strSQL As String = String.Format("Select * From BacLuong Where Status = 1 and IdBacLuong = {0}", _IdBacLuong)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        _ResultId = CType(db.Rows(0)("IdNgachLuong").ToString(), Int32)
                    End If
                End If
            End Using
            Return _ResultId
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về Id Bảng lương khi biết - Id của Ngạch lương
    ''' </summary>
    ''' <param name="_IdNgachLuong">Id ngạch lương</param>
    ''' <returns>Id Bảng lương cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetIdBangLuong(ByVal _IdNgachLuong As Int32) As Int32
        Try
            Dim _ResultId As Int32 = 0
            Dim strSQL As String = String.Format("Select * From NgachLuong Where IdNgachLuong = {0} and Status = 1", _IdNgachLuong)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        _ResultId = CType(db.Rows(0)("IdBangLuong").ToString(), Int32)
                    End If
                End If
            End Using
            Return _ResultId
        Catch
            Return 0
        End Try
    End Function

    Public Function GetIdNghiDinhLuong(ByVal _IdBangLuong As Int32) As Int32
        Try
            Dim _ResultId As Int32 = 0
            Dim strSQL As String = String.Format("Select * From BangLuong Where IdBangLuong = {0} and Status = 1", _IdBangLuong)
            Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        _ResultId = CType(db.Rows(0)("IdND_Luong").ToString(), Int32)
                    End If
                End If
            End Using
            Return _ResultId
        Catch
            Return 0
        End Try
    End Function
#End Region

#Region "---> Các hàm cập nhật dữ liệu - Hợp đồng lao động <---"
    ''' <summary>
    ''' Hàm cập nhật - Thêm mới hợp đồng lao động của cán bộ
    ''' </summary>
    ''' <param name="obj_contract">Hợp đồng lao động</param>
    ''' <returns>Chỉ số tự sinh xác định tính duy nhất của bản ghi HĐLĐ</returns>
    ''' <remarks></remarks>
    Public Function Insert_LaborContract(ByVal obj_contract As HS_HDLD) As String
        Dim command As SqlCommand = New SqlCommand("HS_HDLD_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdCB_HDLD", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_contract.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoHD", obj_contract.SoHD))
        command.Parameters.Add(New SqlParameter("@_IdLoaiHD", obj_contract.IdLoaiHD))
        command.Parameters.Add(New SqlParameter("@_Ngay_HL", obj_contract.Ngay_HL))
        command.Parameters.Add(New SqlParameter("@_NgayKy_HD", obj_contract.NgayKy_HD))
        command.Parameters.Add(New SqlParameter("@_NguoKy_QD", obj_contract.NguoKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_Nguoiky_QD", obj_contract.IdCV_Nguoiky_QD))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_contract.TuNgay))
        If (CType(obj_contract.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_contract.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_IdHT_TraLuong", obj_contract.IdHT_TraLuong))
        command.Parameters.Add(New SqlParameter("@_IdBacLuong", obj_contract.IdBacLuong))
        command.Parameters.Add(New SqlParameter("@_HeSo", obj_contract.HeSo))
        command.Parameters.Add(New SqlParameter("@_TyleHuong", obj_contract.TyleHuong))
        command.Parameters.Add(New SqlParameter("@_DVKyHDLD", obj_contract.DVKyHDLD))
        command.Parameters.Add(New SqlParameter("@_NoiLamViec", obj_contract.NoiLamViec))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_contract.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_contract.IdCB_HDLD = command.Parameters("@_IdCB_HDLD").Value.ToString
            Return obj_contract.IdCB_HDLD
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm cập nhật - Sửa đổi hợp đồng lao động của cán bộ
    ''' </summary>
    ''' <param name="obj_contract">Hồ sơ hợp đồng lao động</param>
    ''' <remarks></remarks>
    Public Sub Update_LaborContract(ByVal obj_contract As HS_HDLD)

        Dim command As SqlCommand = New SqlCommand("HS_HDLD_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCB_HDLD", obj_contract.IdCB_HDLD))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_contract.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoHD", obj_contract.SoHD))
        command.Parameters.Add(New SqlParameter("@_IdLoaiHD", obj_contract.IdLoaiHD))
        command.Parameters.Add(New SqlParameter("@_Ngay_HL", obj_contract.Ngay_HL))
        command.Parameters.Add(New SqlParameter("@_NgayKy_HD", obj_contract.NgayKy_HD))
        command.Parameters.Add(New SqlParameter("@_NguoKy_QD", obj_contract.NguoKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_Nguoiky_QD", obj_contract.IdCV_Nguoiky_QD))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_contract.TuNgay))
        If (CType(obj_contract.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_contract.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_IdHT_TraLuong", obj_contract.IdHT_TraLuong))
        command.Parameters.Add(New SqlParameter("@_IdBacLuong", obj_contract.IdBacLuong))
        command.Parameters.Add(New SqlParameter("@_HeSo", obj_contract.HeSo))
        command.Parameters.Add(New SqlParameter("@_TyleHuong", obj_contract.TyleHuong))
        command.Parameters.Add(New SqlParameter("@_DVKyHDLD", obj_contract.DVKyHDLD))
        command.Parameters.Add(New SqlParameter("@_NoiLamViec", obj_contract.NoiLamViec))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_contract.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhât sửa đổi hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện  - Xoá hồ sơ hợp đồng lao động của cán bộ
    ''' </summary>
    ''' <param name="_ContractId">Id hợp đồng lao động</param>
    ''' <remarks></remarks>
    Public Sub Delete_LaborContract(ByVal _ContractId As String)
        Dim command As SqlCommand = New SqlCommand("HS_HDLD_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCB_HDLD", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdCB_HDLD").Value = _ContractId
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ hợp đồng lao động: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật thông tin  - Thêm mới hồ sơ thôi việc của cán bộ
    ''' </summary>
    ''' <param name="obj_stopwork">Đối tượng - Cán bộ thôi việc</param>
    ''' <remarks></remarks>
    Public Function Insert_StopWork(ByVal obj_stopwork As HS_CBThoiViec) As String
        Dim command As SqlCommand = New SqlCommand("HS_CBThoiViec_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdCBThoiViec", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_stopwork.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_IdLoaiQD", obj_stopwork.IdLoaiQD))
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_stopwork.LyDo))
        command.Parameters.Add(New SqlParameter("@_NgayKy_QD", obj_stopwork.NgayKy_QD))
        command.Parameters.Add(New SqlParameter("@_Ngay_HL", obj_stopwork.Ngay_HL))
        command.Parameters.Add(New SqlParameter("@_NguoiKy_QD", obj_stopwork.NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_NguoKy_QD", obj_stopwork.IdCV_NguoKy_QD))
        command.Parameters.Add(New SqlParameter("@_TroCap_ThoiViec", obj_stopwork.TroCap_ThoiViec))
        command.Parameters.Add(New SqlParameter("@_TroCap_Khac", obj_stopwork.TroCap_Khac))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_stopwork.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_stopwork.IdCBThoiViec = command.Parameters("@_IdCBThoiViec").Value.ToString
            Return obj_stopwork.IdCBThoiViec
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ cán bộ thôi việc: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Hàm cập nhật thông tin  - Sửa đổi hồ sơ thôi việc của cán bộ
    ''' </summary>
    ''' <param name="obj_stopwork">Đối tượng - Cán bộ thôi việc</param>
    ''' <remarks></remarks>
    Public Sub Update_StopWork(ByVal obj_stopwork As HS_CBThoiViec)
        Dim command As SqlCommand = New SqlCommand("HS_CBThoiViec_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCBThoiViec", obj_stopwork.IdCBThoiViec))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_stopwork.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_IdLoaiQD", obj_stopwork.IdLoaiQD))
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_stopwork.LyDo))
        command.Parameters.Add(New SqlParameter("@_NgayKy_QD", obj_stopwork.NgayKy_QD))
        command.Parameters.Add(New SqlParameter("@_Ngay_HL", obj_stopwork.Ngay_HL))
        command.Parameters.Add(New SqlParameter("@_NguoiKy_QD", obj_stopwork.NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_NguoKy_QD", obj_stopwork.IdCV_NguoKy_QD))
        command.Parameters.Add(New SqlParameter("@_TroCap_ThoiViec", obj_stopwork.TroCap_ThoiViec))
        command.Parameters.Add(New SqlParameter("@_TroCap_Khac", obj_stopwork.TroCap_Khac))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_stopwork.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ cán bộ thôi việc: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện  - Xoá hồ sơ cán bộ thôi việc
    ''' </summary>
    ''' <param name="_StopWorkId">Id hồ sơ cán bộ thôi việc</param>
    ''' <remarks></remarks>
    Public Sub Delete_StopWork(ByVal _StopWorkId As String)
        Dim command As SqlCommand = New SqlCommand("HS_CBThoiViec_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCBThoiViec", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdCBThoiViec").Value = _StopWorkId
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hố sơ cán bộ thôi việc: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

End Class
