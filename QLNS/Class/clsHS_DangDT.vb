Imports System.Data
Imports System.Data.SqlClient

Public Class clsHS_DangDT
    Private _SqlHelper As DBAccess

    Public Sub New()
        _SqlHelper = New DBAccess
    End Sub

#Region "---> Hàm định nghĩa lưới dữ liệu: Hồ sơ và Quá trình sinh hoạt: Đảng - Đoàn - Công Đoàn<---"
    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu - Hiên thị thông tin liên quan Hồ sơ Đảng - Đoàn - Công Đoàn
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="status">Chỉ số xác định Tạo lưới cho hồ sơ. Với quy ước
    ''' 1: Hồ sơ Đảng viên
    ''' 2: Quá trình sinh hoạt Đảng
    ''' 3: Hồ sơ Đoàn viên
    ''' 4: Quá trình sinh hoạt Đoàn
    ''' 5: Hồ sơ Công đoàn
    ''' 6: Quá trình sinh hoạt Công đoàn
    ''' </param>
    ''' <remarks></remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal status As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        'True: Hồ sơ đảng viên. False: Quá trình sinh hoạt đảng
        Select Case status
            Case 1       'Định nghĩa lưới dữ liệu - Hồ sơ đảng viên
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 43
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_SoThe", "Số thẻ đảng")
                dgv_name.Columns.Add("cln_NgayKN", "Ngày kết nạp")
                dgv_name.Columns.Add("cln_Ngayvao", "Ngày vào chính thức")
                dgv_name.Columns.Add("cln_Noi_kn", "Nơi kết nạp")
                dgv_name.Columns.Add("cln_Nguoi_gt", "Người giới thiệu")
                dgv_name.Columns.Add("cln_NoiCapThe", "Nơi cấp thẻ đảng")
                dgv_name.Columns.Add("cln_Ngayra", "Ngày ra khỏi đảng")
                dgv_name.Columns.Add("cln_Lydo", "Lý do ra khỏi đảng")
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_SoThe").Width = 85
                dgv_name.Columns("cln_NgayKN").Width = 85
                dgv_name.Columns("cln_Ngayvao").Width = 111
                dgv_name.Columns("cln_Noi_kn").Width = 200
                dgv_name.Columns("cln_Nguoi_gt").Width = 140
                dgv_name.Columns("cln_NoiCapThe").Width = 152
                dgv_name.Columns("cln_Ngayra").Width = 120
                dgv_name.Columns("cln_Lydo").Width = 210

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 9
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_NgayKN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Ngayvao").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Ngayra").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

            Case 2      'Định nghĩa lưới dữ liệu - Quá trình sinh hoạt đảng
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 43
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ đảng")
                dgv_name.Columns.Add("cln_Chibo", "Chi bộ đảng")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Tungay").Width = 85
                dgv_name.Columns("cln_Denngay").Width = 85
                dgv_name.Columns("cln_Chucvu").Width = 200
                dgv_name.Columns("cln_Chibo").Width = 180
                dgv_name.Columns("cln_Ghichu").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 6
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Choice").Frozen = True
            Case 3       'Định nghĩa lưới dữ liệu - Hồ sơ Đoàn viên
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 43
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_SoThe", "Số thẻ đoàn")
                dgv_name.Columns.Add("cln_Ngayvao", "Ngày vào đoàn")
                dgv_name.Columns.Add("cln_Noi_kn", "Nơi kết nạp")
                dgv_name.Columns.Add("cln_NoiCapThe", "Nơi cấp thẻ đoàn")
                dgv_name.Columns.Add("cln_Ngayra", "Ngày ra khỏi đoàn")
                dgv_name.Columns.Add("cln_Lydo", "Lý do ra khỏi đoàn")
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_SoThe").Width = 85
                dgv_name.Columns("cln_Ngayvao").Width = 95
                dgv_name.Columns("cln_Noi_kn").Width = 140
                dgv_name.Columns("cln_NoiCapThe").Width = 140
                dgv_name.Columns("cln_Ngayra").Width = 120
                dgv_name.Columns("cln_Lydo").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_Ngayvao").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Ngayra").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

            Case 4      'Định nghĩa lưới dữ liệu - Quá trình sinh hoạt Đoàn
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 43
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_TuNgay", "Từ ngày")
                dgv_name.Columns.Add("cln_DenNgay", "Đến ngày")
                dgv_name.Columns.Add("cln_ChucVu", "Chức vụ đoàn")
                dgv_name.Columns.Add("cln_ChiDoan", "Chi đoàn")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_TuNgay").Width = 85
                dgv_name.Columns("cln_DenNgay").Width = 85
                dgv_name.Columns("cln_ChucVu").Width = 200
                dgv_name.Columns("cln_ChiDoan").Width = 180
                dgv_name.Columns("cln_Ghichu").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 6
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_TuNgay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DenNgay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Choice").Frozen = True
            Case 5       'Định nghĩa lưới dữ liệu - Hồ sơ Công đoàn
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_SoThe", "Số thẻ")
                dgv_name.Columns.Add("cln_Ngayvao", "Ngày vào")
                dgv_name.Columns.Add("cln_Noi_kn", "Nơi vào")
                dgv_name.Columns.Add("cln_NoiCapThe", "Nơi cấp thẻ đoàn")
                dgv_name.Columns.Add("cln_Ngayra", "Ngày ra")
                dgv_name.Columns.Add("cln_Lydo", "Lý do ra khỏi công đoàn")
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_SoThe").Width = 85
                dgv_name.Columns("cln_Ngayvao").Width = 95
                dgv_name.Columns("cln_Noi_kn").Width = 140
                dgv_name.Columns("cln_NoiCapThe").Width = 140
                dgv_name.Columns("cln_Ngayra").Width = 120
                dgv_name.Columns("cln_Lydo").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_Ngayvao").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Ngayra").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 6      'Định nghĩa lưới dữ liệu - Quá trình sinh hoạt Công Đoàn
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 43
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ công đoàn")
                dgv_name.Columns.Add("cln_CoQuan", "Cơ quan (CĐ cơ sở - CĐ trực thuộc)")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Tungay").Width = 85
                dgv_name.Columns("cln_Denngay").Width = 85
                dgv_name.Columns("cln_Chucvu").Width = 200
                dgv_name.Columns("cln_CoQuan").Width = 180
                dgv_name.Columns("cln_Ghichu").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 6
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Choice").Frozen = True
        End Select
        dgv_name.Columns("cln_Code").Visible = False
    End Sub
#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Public Class HS_DangVien
        Private _IdDangVien As String
        Private _IdCanBo As String
        Private _SoThe As String
        Private _NgayKN As DateTime
        Private _NgayVao As DateTime
        Private _Noi_KetNap As String
        Private _Nguoi_GT As String
        Private _NoiCapThe As String
        Private _NgayRa As DateTime
        Private _LyDo As String
        Private _Date_Create As DateTime
        Private _Date_Update As DateTime

        Public Sub New()
            _IdDangVien = ""
            _IdCanBo = ""
            _SoThe = ""
            _NgayKN = DateTime.Now
            _NgayVao = DateTime.Now
            _Noi_KetNap = ""
            _Nguoi_GT = ""
            _NoiCapThe = ""
            _NgayRa = DateTime.Now
            _LyDo = ""
            _Date_Create = DateTime.Now
            _Date_Update = DateTime.Now
        End Sub

        Public Property IdDangVien() As String
            Get
                Return _IdDangVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id đảng viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdDangVien = value
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

        Public Property SoThe() As String
            Get
                Return _SoThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số thẻ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoThe = value
            End Set
        End Property

        Public Property NgayKN() As DateTime
            Get
                Return _NgayKN
            End Get
            Set(ByVal value As DateTime)
                _NgayKN = value
            End Set
        End Property

        Public Property NgayVao() As DateTime
            Get
                Return _NgayVao
            End Get
            Set(ByVal value As DateTime)
                _NgayVao = value
            End Set
        End Property

        Public Property Noi_KetNap() As String
            Get
                Return _Noi_KetNap
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi kết nạp có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Noi_KetNap = value
            End Set
        End Property

        Public Property Nguoi_GT() As String
            Get
                Return _Nguoi_GT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người giới thiệu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Nguoi_GT = value
            End Set
        End Property

        Public Property NoiCapThe() As String
            Get
                Return _NoiCapThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp thẻ đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiCapThe = value
            End Set
        End Property

        Public Property NgayRa() As DateTime
            Get
                Return _NgayRa
            End Get
            Set(ByVal value As DateTime)
                _NgayRa = value
            End Set
        End Property

        Public Property LyDo() As String
            Get
                Return _LyDo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LyDo = value
            End Set
        End Property

        Public Property Date_Create() As DateTime
            Get
                Return _Date_Create
            End Get
            Set(ByVal value As DateTime)
                _Date_Create = value
            End Set
        End Property

        Public Property Date_Update() As DateTime
            Get
                Return _Date_Update
            End Get
            Set(ByVal value As DateTime)
                _Date_Update = value
            End Set
        End Property
    End Class

    Public Class HS_Dang
        Private _IdHSDang As String
        Private _IdDangVien As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _IdCVDang As Int32
        Private _ChiBo As String
        Private _IsKiemNhiem As Byte
        Private _GhiChu As String
        Public Sub New()
            _IdHSDang = ""
            _IdDangVien = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _IdCVDang = 0
            _ChiBo = ""
            _IsKiemNhiem = 0
            _GhiChu = ""
        End Sub

        Public Property IdHSDang() As String
            Get
                Return _IdHSDang
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHSDang = value
            End Set
        End Property

        Public Property IdDangVien() As String
            Get
                Return _IdDangVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id đảng viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdDangVien = value
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

        Public Property IdCVDang() As Int32
            Get
                Return _IdCVDang
            End Get
            Set(ByVal value As Int32)
                _IdCVDang = value
            End Set
        End Property

        Public Property ChiBo() As String
            Get
                Return _ChiBo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chi bộ đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ChiBo = value
            End Set
        End Property

        Public Property IsKiemNhiem() As Byte
            Get
                Return _IsKiemNhiem
            End Get
            Set(ByVal value As Byte)
                _IsKiemNhiem = value
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

        Public Sub GetData(ByVal _row As DataRow)
            IdHSDang = _row("IdHSDang").ToString()
            IdDangVien = _row("IdDangVien").ToString()
            TuNgay = _row("TuNgay").ToString()
            DenNgay = _row("DenNgay").ToString()
            IdCVDang = _row("IdCVDang").ToString()
            ChiBo = _row("ChiBo").ToString()
            GhiChu = _row("GhiChu").ToString()
        End Sub

        Public Sub GetData(ByVal reader As SqlDataReader)
            IdHSDang = reader("IdHSDang").ToString()
            IdDangVien = reader("IdDangVien").ToString()
            TuNgay = reader("TuNgay").ToString()
            DenNgay = reader("DenNgay").ToString()
            IdCVDang = reader("IdCVDang").ToString()
            ChiBo = reader("ChiBo").ToString()
            GhiChu = reader("GhiChu").ToString()
        End Sub
    End Class

    Public Class HS_DoanVien
        Private _IdDoanVien As String
        Private _IdCanBo As String
        Private _SoThe As String
        Private _NgayVao As DateTime
        Private _Noi_KetNap As String
        Private _NoiCapThe As String
        Private _NgayRa As DateTime
        Private _LyDo As String

        Public Sub New()
            _IdDoanVien = ""
            _IdCanBo = ""
            _SoThe = ""
            _NgayVao = DateTime.Now
            _Noi_KetNap = ""
            _NoiCapThe = ""
            _NgayRa = DateTime.Now
            _LyDo = ""
        End Sub

        Public Property IdDoanVien() As String
            Get
                Return _IdDoanVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id đoàn viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdDoanVien = value
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

        Public Property SoThe() As String
            Get
                Return _SoThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số thẻ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoThe = value
            End Set
        End Property

        Public Property NgayVao() As DateTime
            Get
                Return _NgayVao
            End Get
            Set(ByVal value As DateTime)
                _NgayVao = value
            End Set
        End Property

        Public Property Noi_KetNap() As String
            Get
                Return _Noi_KetNap
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi kết nạp có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Noi_KetNap = value
            End Set
        End Property

        Public Property NoiCapThe() As String
            Get
                Return _NoiCapThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp thẻ đoàn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiCapThe = value
            End Set
        End Property

        Public Property NgayRa() As DateTime
            Get
                Return _NgayRa
            End Get
            Set(ByVal value As DateTime)
                _NgayRa = value
            End Set
        End Property

        Public Property LyDo() As String
            Get
                Return _LyDo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LyDo = value
            End Set
        End Property
    End Class

    Public Class HS_Doan
        Private _IdHSDoan As String
        Private _IdDoanVien As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _IdCVDoan As Int32
        Private _ChiDoan As String
        Private _GhiChu As String

        Public Sub New()
            _IdHSDoan = ""
            _IdDoanVien = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _IdCVDoan = 0
            _ChiDoan = ""
            _GhiChu = ""
        End Sub

        Public Property IdHSDoan() As String
            Get
                Return _IdHSDoan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id quá trình sinh hoạt đoàn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHSDoan = value
            End Set
        End Property

        Public Property IdDoanVien() As String
            Get
                Return _IdDoanVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ đoàn viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdDoanVien = value
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

        Public Property IdCVDoan() As Int32
            Get
                Return _IdCVDoan
            End Get
            Set(ByVal value As Int32)
                _IdCVDoan = value
            End Set
        End Property

        Public Property ChiDoan() As String
            Get
                Return _ChiDoan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chi đoàn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ChiDoan = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

    End Class

    Public Class HS_CongDoan
        Private _IdCongDoan As String
        Private _IdCanBo As String
        Private _SoThe As String
        Private _NgayVao As DateTime
        Private _Noi_Vao As String
        Private _NoiCapThe As String
        Private _NgayRa As DateTime
        Private _LyDo As String
        Public Sub New()
            _IdCongDoan = ""
            _IdCanBo = ""
            _SoThe = ""
            _NgayVao = DateTime.Now
            _Noi_Vao = ""
            _NoiCapThe = ""
            _NgayRa = DateTime.Now
            _LyDo = ""
        End Sub

        Public Property IdCongDoan() As String
            Get
                Return _IdCongDoan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ công đoàn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCongDoan = value
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

        Public Property SoThe() As String
            Get
                Return _SoThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số thẻ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoThe = value
            End Set
        End Property

        Public Property NgayVao() As DateTime
            Get
                Return _NgayVao
            End Get
            Set(ByVal value As DateTime)
                _NgayVao = value
            End Set
        End Property

        Public Property Noi_Vao() As String
            Get
                Return _Noi_Vao
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi vào có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Noi_Vao = value
            End Set
        End Property

        Public Property NoiCapThe() As String
            Get
                Return _NoiCapThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp thẻ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiCapThe = value
            End Set
        End Property

        Public Property NgayRa() As DateTime
            Get
                Return _NgayRa
            End Get
            Set(ByVal value As DateTime)
                _NgayRa = value
            End Set
        End Property

        Public Property LyDo() As String
            Get
                Return _LyDo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LyDo = value
            End Set
        End Property
    End Class

    Public Class HS_CongDoanQT
        Private _IdHSCDoan As String
        Private _IdCongDoan As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _IdChucVu As Int32
        Private _CoQuan As String
        Private _GhiChu As String

        Public Sub New()
            _IdHSCDoan = ""
            _IdCongDoan = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _IdChucVu = 0
            _CoQuan = ""
            _GhiChu = ""
        End Sub

        Public Property IdHSCDoan() As String
            Get
                Return _IdHSCDoan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id quá trình sinh hoạt công đoàn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHSCDoan = value
            End Set
        End Property

        Public Property IdCongDoan() As String
            Get
                Return _IdCongDoan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ công đoàn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCongDoan = value
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

        Public Property IdChucVu() As Int32
            Get
                Return _IdChucVu
            End Get
            Set(ByVal value As Int32)
                _IdChucVu = value
            End Set
        End Property

        Public Property CoQuan() As String
            Get
                Return _CoQuan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị cơ quan có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CoQuan = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property
    End Class
#End Region

#Region "---> Các hàm liên quan đến lấy dữ liệu Đảng - Đoàn - Công Đoàn <---"
    ''' <summary>
    ''' Hàm trả về danh sách các bản ghi Hồ sơ đảng viên theo Id cán bộ truyền vào
    ''' </summary>
    ''' <param name="_IdCanBo">Id cán bộ</param>
    ''' <param name="_State">Chỉ số xác định Hồ sơ cần lấy. Với quy ước: 
    ''' 1-Hồ sơ Đảng viên
    ''' 2-Hồ sơ Đoàn viên
    ''' 3-Hồ sơ Công Đoàn
    ''' </param>
    ''' <returns>Mảng các bản ghi hồ sơ đảng</returns>
    ''' <remarks></remarks>
    Public Function GetAll(ByVal _IdCanBo As String, ByVal _State As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = Nothing
            Select Case _State
                Case 1
                    command = New SqlCommand("HS_DangVien_GetForIdCanBo", connection)
                Case 2
                    command = New SqlCommand("HS_DoanVien_GetForIdCanBo", connection)
                Case 3
                    command = New SqlCommand("HS_CongDoan_GetForIdCanBo", connection)
            End Select
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
            command.Parameters("@_IdCanBo").Value = _IdCanBo

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_DoanThe")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_DoanThe")
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Lấy dữ liệu hồ sơ đoàn thể", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về row hồ sơ Đoàn thể theo mã hiệu truyền vào
    ''' </summary>
    ''' <param name="_Code">Chuỗi chỉ số duy nhất xác định hồ sơ</param>
    ''' <param name="_State">Chỉ số xác định hồ sơ cần lấy. Với quy uơcs
    '''                   1-Hồ sơ Đảng viên
    '''                   2-Hồ sơ Đoàn viên
    '''                   3-Hồ sơ Công Đoàn
    ''' </param>
    ''' <returns>Row cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetRecord(ByVal _Code As String, ByVal _State As Byte) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = Nothing
            Select Case _State
                Case 1
                    command = New SqlCommand("HS_DangVien_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdDangVien", SqlDbType.VarChar))
                    command.Parameters("@_IdDangVien").Value = _Code
                Case 2
                    command = New SqlCommand("HS_DoanVien_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdDoanVien", SqlDbType.VarChar))
                    command.Parameters("@_IdDoanVien").Value = _Code
                Case 3
                    command = New SqlCommand("HS_CongDoan_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCongDoan", SqlDbType.VarChar))
                    command.Parameters("@_IdCongDoan").Value = _Code
            End Select
            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_DoanThe")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_DoanThe").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện lấy danh sách quá trình sinh hoạt theo id hồ sơ đoàn thể truyền vào
    ''' </summary>
    ''' <param name="_DocumentId">Id hồ sơ đoàn thể truyền vào</param>
    ''' <param name="_State">Chỉ số xác định Hồ sơ cần lấy. Với quy ước:
    '''                   1-Hồ sơ Đảng viên
    '''                   2-Hồ sơ Đoàn viên
    '''                   3-Hồ sơ Công Đoàn
    ''' </param>
    ''' <returns>Danh sách quá trình sinh hoạt</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Process(ByVal _DocumentId As String, ByVal _State As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = Nothing
            Select Case _State
                Case 1
                    command = New SqlCommand("HS_Dang_GetAll", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdDangVien", SqlDbType.VarChar))
                    command.Parameters("@_IdDangVien").Value = _DocumentId
                Case 2
                    command = New SqlCommand("HS_Doan_GetAll", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdDoanVien", SqlDbType.VarChar))
                    command.Parameters("@_IdDoanVien").Value = _DocumentId
                Case 3
                    command = New SqlCommand("HS_CongDoanQT_GetAll", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCongDoan", SqlDbType.VarChar))
                    command.Parameters("@_IdCongDoan").Value = _DocumentId
            End Select
            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_QTSinhHoat")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_QTSinhHoat")
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Lấy dữ liệu quá trình sinh hoạt", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về bản ghi quá trình sinh hoạt Đảng - Đoàn - Công đoàn
    ''' </summary>
    ''' <param name="_Code">Chuỗi chỉ số xác định tính duy nhất của bản ghi hồ sơ quá trình sinh hoạt</param>
    ''' <param name="_State">Chỉ số xác định Hồ sơ cần lấy. Với quy ước:
    '''                   1-Hồ sơ Đảng viên
    '''                   2-Hồ sơ Đoàn viên
    '''                   3-Hồ sơ Công Đoàn
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRecordProcess(ByVal _Code As String, ByVal _State As Byte) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = Nothing
            Select Case _State
                Case 1
                    command = New SqlCommand("HS_Dang_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdHSDang", SqlDbType.VarChar))
                    command.Parameters("@_IdHSDang").Value = _Code
                Case 2
                    command = New SqlCommand("HS_Doan_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdHSDoan", SqlDbType.VarChar))
                    command.Parameters("@_IdHSDoan").Value = _Code
                Case 3
                    command = New SqlCommand("HS_CongDoanQT_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdHSCDoan", SqlDbType.VarChar))
                    command.Parameters("@_IdHSCDoan").Value = _Code
            End Select
            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "RecordProcess")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("RecordProcess").Rows(0)
            End Using
        Catch ex As Exception
            MessageBox.Show("Lấy bản ghi quá trình sinh hoạt: " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function
#End Region

#Region "---> Các hàm liên quan đến cập nhật dữ liệu <---"
    ' Hàm cập nhật dữ liệu Hồ sơ Đảng viên
    ''' <summary>
    ''' Hàm cập nhật thêm mới Hồ sơ đảng viên của cán bộ. Trả về id của hồ sơ đảng viên
    ''' </summary>
    ''' <param name="obj_hs_dang">Đối tượng hs đảng truyền vào</param>
    ''' <returns>IdDangVien</returns>
    ''' <remarks></remarks>
    Public Function Insert_PartyDocument(ByVal obj_hs_dang As HS_DangVien) As String
        Dim command As SqlCommand = New SqlCommand("HS_DangVien_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdDangVien", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_dang.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_hs_dang.SoThe))
        command.Parameters.Add(New SqlParameter("@_NgayKN", obj_hs_dang.NgayKN))
        command.Parameters.Add(New SqlParameter("@_NgayVao", obj_hs_dang.NgayVao))
        command.Parameters.Add(New SqlParameter("@_Noi_KetNap", obj_hs_dang.Noi_KetNap))
        command.Parameters.Add(New SqlParameter("@_Nguoi_GT", obj_hs_dang.Nguoi_GT))
        command.Parameters.Add(New SqlParameter("@_NoiCapThe", obj_hs_dang.NoiCapThe))
        If (CType(obj_hs_dang.NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayRa", obj_hs_dang.NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_hs_dang.LyDo))
        Try
            _SqlHelper.executeSQL(command)
            obj_hs_dang.IdDangVien = command.Parameters("@_IdDangVien").Value.ToString
            Return obj_hs_dang.IdDangVien
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ đảng", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu Hồ sơ đảng viên của cán bộ
    ''' </summary>
    ''' <param name="obj_hs_dang">Đối tượng Hồ sơ đảng viên</param>
    ''' <remarks></remarks>
    Public Sub Update_PartyDocument(ByVal obj_hs_dang As HS_DangVien)
        Dim command As SqlCommand = New SqlCommand("HS_DangVien_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdDangVien", obj_hs_dang.IdDangVien))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_dang.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_NgayKN", obj_hs_dang.NgayKN))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_hs_dang.SoThe))
        command.Parameters.Add(New SqlParameter("@_NgayVao", obj_hs_dang.NgayVao))
        command.Parameters.Add(New SqlParameter("@_Noi_KetNap", obj_hs_dang.Noi_KetNap))
        command.Parameters.Add(New SqlParameter("@_Nguoi_GT", obj_hs_dang.Nguoi_GT))
        command.Parameters.Add(New SqlParameter("@_NoiCapThe", obj_hs_dang.NoiCapThe))
        If (CType(obj_hs_dang.NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayRa", obj_hs_dang.NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_hs_dang.LyDo))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ đảng", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ đảng viên của cán bộ
    ''' </summary>
    ''' <param name="_Id">Id hồ sơ đảng viên</param>
    ''' <remarks></remarks>
    Public Sub Delete_PartyDocument(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HS_DangVien_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdDangVien", SqlDbType.VarChar))
        command.Parameters("@_IdDangVien").Value = _Id
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ đảng viên: " + ex.Message.ToString(), "Xoá bỏ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu - Quá trình sinh hoạt đảng
    ''' </summary>
    ''' <param name="obj_qt_dang"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert_PartyProcess(ByVal obj_qt_dang As HS_Dang) As String
        Dim command As SqlCommand = New SqlCommand("HS_Dang_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdHSDang", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdDangVien", obj_qt_dang.IdDangVien))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_qt_dang.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_qt_dang.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdCVDang", obj_qt_dang.IdCVDang))
        command.Parameters.Add(New SqlParameter("@_ChiBo", obj_qt_dang.ChiBo))
        command.Parameters.Add(New SqlParameter("@_IsKiemNhiem", obj_qt_dang.IsKiemNhiem))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_qt_dang.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_qt_dang.IdHSDang = command.Parameters("@_IdHSDang").Value.ToString
            Return obj_qt_dang.IdHSDang
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới quá trình sinh hoạt đảng", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện Cập nhật dữ liệu - Quá trình sinh hoạt đảng
    ''' </summary>
    ''' <param name="obj_qt_dang"></param>
    ''' <remarks></remarks>
    Public Sub Update_PartyProcess(ByVal obj_qt_dang As HS_Dang)
        Dim command As SqlCommand = New SqlCommand("HS_Dang_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSDang", obj_qt_dang.IdHSDang))
        command.Parameters.Add(New SqlParameter("@_IdDangVien", obj_qt_dang.IdDangVien))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_qt_dang.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_qt_dang.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdCVDang", obj_qt_dang.IdCVDang))
        command.Parameters.Add(New SqlParameter("@_ChiBo", obj_qt_dang.ChiBo))
        command.Parameters.Add(New SqlParameter("@_IsKiemNhiem", obj_qt_dang.IsKiemNhiem))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_qt_dang.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật quá trình sinh hoạt đảng", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Quá trình sinh hoạt đảng
    ''' </summary>
    ''' <param name="_Id">Id quá trình sinh hoạt đảng</param>
    ''' <remarks></remarks>
    Public Sub Delete_PartyProcess(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HS_Dang_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSDang", SqlDbType.VarChar))
        command.Parameters("@_IdHSDang").Value = _Id
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ quá trình sinh hoạt đảng: " + ex.Message.ToString(), "Xoá bỏ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ' Hàm cập nhật dữ liệu Hồ sơ Đoàn viên
    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu Hồ sơ đoàn viên
    ''' </summary>
    ''' <param name="obj_hs_doanvien">Hồ sơ đoàn viên</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert_UnionMember(ByVal obj_hs_doanvien As HS_DoanVien) As String
        Dim command As SqlCommand = New SqlCommand("HS_DoanVien_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdDoanVien", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_doanvien.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_hs_doanvien.SoThe))
        command.Parameters.Add(New SqlParameter("@_NgayVao", obj_hs_doanvien.NgayVao))
        command.Parameters.Add(New SqlParameter("@_Noi_KetNap", obj_hs_doanvien.Noi_KetNap))
        command.Parameters.Add(New SqlParameter("@_NoiCapThe", obj_hs_doanvien.NoiCapThe))
        If (CType(obj_hs_doanvien.NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayRa", obj_hs_doanvien.NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_hs_doanvien.LyDo))
        Try
            _SqlHelper.executeSQL(command)
            obj_hs_doanvien.IdDoanVien = command.Parameters("@_IdDoanVien").Value.ToString
            Return obj_hs_doanvien.IdDoanVien
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ đoàn viên", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện Cập nhật dữ liệu hồ sơ đoàn viên
    ''' </summary>
    ''' <param name="obj_hs_doanvien">Hồ sơ đoàn viên</param>
    ''' <remarks></remarks>
    Public Sub Update_UnionMember(ByVal obj_hs_doanvien As HS_DoanVien)
        Dim command As SqlCommand = New SqlCommand("HS_DoanVien_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdDoanVien", obj_hs_doanvien.IdDoanVien))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_doanvien.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_hs_doanvien.SoThe))
        command.Parameters.Add(New SqlParameter("@_NgayVao", obj_hs_doanvien.NgayVao))
        command.Parameters.Add(New SqlParameter("@_Noi_KetNap", obj_hs_doanvien.Noi_KetNap))
        command.Parameters.Add(New SqlParameter("@_NoiCapThe", obj_hs_doanvien.NoiCapThe))
        If (CType(obj_hs_doanvien.NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayRa", obj_hs_doanvien.NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_hs_doanvien.LyDo))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ đoàn viên", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ đoàn viên của cán bộ
    ''' </summary>
    ''' <param name="_Id">Id hồ sơ đoàn viên</param>
    ''' <remarks></remarks>
    Public Sub Delete_UnionMember(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HS_DoanVien_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdDoanVien", SqlDbType.VarChar))
        command.Parameters("@_IdDoanVien").Value = _Id
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ đoàn viên: " + ex.Message.ToString(), "Xoá bỏ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu - Quá trình sinh hoạt đoàn
    ''' </summary>
    ''' <param name="obj_qt_doan">Quá trình sinh hoạt đoàn</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert_UnionMemberProcess(ByVal obj_qt_doan As HS_Doan) As String
        Dim command As SqlCommand = New SqlCommand("HS_Doan_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdHSDoan", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdDoanVien", obj_qt_doan.IdDoanVien))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_qt_doan.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_qt_doan.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdCVDoan", obj_qt_doan.IdCVDoan))
        command.Parameters.Add(New SqlParameter("@_ChiDoan", obj_qt_doan.ChiDoan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_qt_doan.GhiChu))
        Try
            _SqlHelper.executeSQL(command)
            obj_qt_doan.IdHSDoan = command.Parameters("@_IdHSDoan").Value.ToString
            Return obj_qt_doan.IdHSDoan
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới quá trình sinh hoạt đoàn", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện Cập nhật dữ liệu - Quá trình sinh hoạt đoàn
    ''' </summary>
    ''' <param name="obj_qt_doan">Quá trình sinh hoạt đoàn</param>
    ''' <remarks></remarks>
    Public Sub Update_UnionMemberProcess(ByVal obj_qt_doan As HS_Doan)
        Dim command As SqlCommand = New SqlCommand("HS_Doan_Update")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_IdHSDoan", obj_qt_doan.IdHSDoan))
        command.Parameters.Add(New SqlParameter("@_IdDoanVien", obj_qt_doan.IdDoanVien))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_qt_doan.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_qt_doan.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdCVDoan", obj_qt_doan.IdCVDoan))
        command.Parameters.Add(New SqlParameter("@_ChiDoan", obj_qt_doan.ChiDoan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_qt_doan.GhiChu))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật quá trình sinh hoạt đoàn", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Quá trình sinh hoạt đoàn
    ''' </summary>
    ''' <param name="_Id">Id quá trình sinh hoạt đoàn</param>
    ''' <remarks></remarks>
    Public Sub Delete_UnionMemberProcess(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HS_Doan_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSDoan", SqlDbType.VarChar))
        command.Parameters("@_IdHSDoan").Value = _Id
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ quá trình sinh hoạt đoàn: " + ex.Message.ToString(), "Xoá bỏ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ' Hàm cập nhật dữ liệu Hồ sơ Công Đoàn

    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu Hồ sơ Công Đoàn
    ''' </summary>
    ''' <param name="obj_hs_congdoan">Hồ sơ Công Đoàn</param>
    ''' <returns>Chuỗi chỉ số xác định bản ghi vừa tạo thêm</returns>
    ''' <remarks></remarks>
    Public Function Insert_TradeUnion(ByVal obj_hs_congdoan As HS_CongDoan) As String
        Dim command As SqlCommand = New SqlCommand("HS_CongDoan_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdCongDoan", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_congdoan.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_hs_congdoan.SoThe))
        command.Parameters.Add(New SqlParameter("@_NgayVao", obj_hs_congdoan.NgayVao))
        command.Parameters.Add(New SqlParameter("@_Noi_Vao", obj_hs_congdoan.Noi_Vao))
        command.Parameters.Add(New SqlParameter("@_NoiCapThe", obj_hs_congdoan.NoiCapThe))
        If (CType(obj_hs_congdoan.NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayRa", obj_hs_congdoan.NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_hs_congdoan.LyDo))
        Try
            _SqlHelper.executeSQL(command)
            obj_hs_congdoan.IdCongDoan = command.Parameters("@_IdCongDoan").Value.ToString
            Return obj_hs_congdoan.IdCongDoan
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ công đoàn", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện Cập nhật dữ liệu hồ sơ Công đoàn
    ''' </summary>
    ''' <param name="obj_hs_congdoan">Hồ sơ Công đoàn</param>
    ''' <remarks></remarks>
    Public Sub Update_TradeUnion(ByVal obj_hs_congdoan As HS_CongDoan)
        Dim command As SqlCommand = New SqlCommand("HS_CongDoan_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCongDoan", obj_hs_congdoan.IdCongDoan))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_congdoan.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_hs_congdoan.SoThe))
        command.Parameters.Add(New SqlParameter("@_NgayVao", obj_hs_congdoan.NgayVao))
        command.Parameters.Add(New SqlParameter("@_Noi_Vao", obj_hs_congdoan.Noi_Vao))
        command.Parameters.Add(New SqlParameter("@_NoiCapThe", obj_hs_congdoan.NoiCapThe))
        If (CType(obj_hs_congdoan.NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayRa", obj_hs_congdoan.NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_hs_congdoan.LyDo))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ công đoàn", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ công đoàn của cán bộ
    ''' </summary>
    ''' <param name="_Id">Id hồ sơ đoàn viên</param>
    ''' <remarks></remarks>
    Public Sub Delete_TradeUnion(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HS_CongDoan_Delete")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_IdCongDoan", SqlDbType.VarChar))
        command.Parameters("@_IdCongDoan").Value = _Id
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ công đoàn: " + ex.Message.ToString(), "Xoá bỏ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu - Quá trình sinh Công đoàn
    ''' </summary>
    ''' <param name="obj_qt_congdoan">Quá trình sinh hoạt công đoàn</param>
    ''' <returns>Chuỗi chỉ số xác định bản ghi vừa thêm mới</returns>
    ''' <remarks></remarks>
    Public Function Insert_TradeUnionProcess(ByVal obj_qt_congdoan As HS_CongDoanQT) As String
        Dim command As SqlCommand = New SqlCommand("HS_CongDoanQT_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdHSCDoan", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCongDoan", obj_qt_congdoan.IdCongDoan))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_qt_congdoan.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_qt_congdoan.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdChucVu", obj_qt_congdoan.IdChucVu))
        command.Parameters.Add(New SqlParameter("@_CoQuan", obj_qt_congdoan.CoQuan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_qt_congdoan.GhiChu))
        Try
            _SqlHelper.executeSQL(command)
            obj_qt_congdoan.IdHSCDoan = command.Parameters("@_IdHSCDoan").Value.ToString
            Return obj_qt_congdoan.IdHSCDoan
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới quá trình sinh hoạt công đoàn", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện Cập nhật dữ liệu - Quá trình sinh hoạt Công đoàn
    ''' </summary>
    ''' <param name="obj_qt_congdoan">Quá trình sinh hoạt Công đoàn</param>
    ''' <remarks></remarks>
    Public Sub Update_TradeUnionProcess(ByVal obj_qt_congdoan As HS_CongDoanQT)
        Dim command As SqlCommand = New SqlCommand("HS_CongDoanQT_Update")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_IdHSCDoan", obj_qt_congdoan.IdHSCDoan))
        command.Parameters.Add(New SqlParameter("@_IdCongDoan", obj_qt_congdoan.IdCongDoan))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_qt_congdoan.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_qt_congdoan.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdChucVu", obj_qt_congdoan.IdChucVu))
        command.Parameters.Add(New SqlParameter("@_CoQuan", obj_qt_congdoan.CoQuan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_qt_congdoan.GhiChu))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật quá trình sinh hoạt công đoàn", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Quá trình sinh hoạt Công đoàn của cán bộ
    ''' </summary>
    ''' <param name="_Id">Id quá trình sinh hoạt Công đoàn</param>
    ''' <remarks></remarks>
    Public Sub Delete_TradeUnionProcess(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HS_CongDoanQT_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSCDoan", SqlDbType.VarChar))
        command.Parameters("@_IdHSCDoan").Value = _Id
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ quá trình sinh hoạt công đoàn: " + ex.Message.ToString(), "Xoá bỏ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

End Class
