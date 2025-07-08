Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections

Public Class clsHT_DanhMuc
    Private _SqlHelper As DBAccess

    Public Sub New()
        _SqlHelper = New DBAccess
    End Sub

#Region "---> Const string: Khai báo - Các câu lệnh truy vấn danh mục <---"
    Public Const Sql_UnitAll As String = "Select Id,dbo.Replace_BranchName(Ten_Goi) Name,Ma_So Code From ChiNhanh Where Status=1 Order By Substring(Ma_So,1,4), Id_Goc,Ma_So"

    Public Const Sql_Province As String = "Select Id,dbo.Replace_BranchName(Ten_Goi) Name,Ma_So Code From ChiNhanh Where Status=1 And Id_Goc In (0,1) Order By Substring(Ma_So,1,4), Id_Goc,Ma_So"

    'Danh sách các quốc gia trên thế giới --> Tham chiếu tới bảng Quốc gia
    Public Const Sql_Quocgia As String = "Select Id,Ten_Goi From QuocGia Where Status = 1"

    'Danh mục Dân tộc  -->  Tham chiếu tới bảng danh mục với id_goc = 21
    Public Const Sql_Dantoc As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 21 and Status = 1"

    'Danh mục Tôn giáo  -->  Tham chiếu tới bảng danh mục với id_goc = 24
    Public Const Sql_Tongiao As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 24 and Status = 1"

    'Danh mục Thành phần gia đình  -->  Tham chiếu tới bảng danh mục với id_goc = 13
    Public Const Sql_Tpgd As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 13 and Status = 1 Order by ten_goi Asc"

    'Danh mục Ưu tiên bản thân  -->  Tham chiếu tới bảng danh mục với id_goc = 18
    Public Const Sql_Ut_banthan As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 18 and Status = 1"

    'Danh mục Ưu tiên gia đình  -->  Tham chiếu tới bảng danh mục với id_goc = 19
    Public Const Sql_Ut_giadinh As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 19 and Status = 1 Order by ten_goi Asc"

    'Danh mục Trình độ Văn hoá  -->  Tham chiếu tới bảng danh mục với id_goc = 6
    Public Const Sql_Tdvh As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 6 and Status = 1 Order by ten_goi Asc"

    'Danh mục Trình độ chính trị  -->  Tham chiếu tới bảng danh mục với id_goc = 36
    Public Const Sql_Tdct As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 36 and Status = 1 Order by ten_goi Asc"

    'Danh mục Học hàm  -->  Tham chiếu tới bảng danh mục với id_goc = 26
    Public Const Sql_Hocham As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 26 and Status = 1 Order by Ma_so Asc"

    'Danh mục Học vị  -->  Tham chiếu tới bảng danh mục với id_goc = 20
    Public Const Sql_Hocvi As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 20 and Status = 1 Order by Ma_so Asc"

    'Danh sanh địa danh - Tỉnh (Thành phố) trong toàn quốc --> Tham chiếu tới bảng địa danh với điều kiện id_goc = 0
    'Public Const Sql_TinhTP As String = "Select Id,Ten_Goi From DiaDanh Where id_goc = 0 and Status = 1 Order by Ten_Goi Asc"
    Public Const Sql_TinhTP As String = "Select Id,Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa = '00' And Ma_Thon = '00' And TrangThai = 'A' Order by Ma_Tinh Asc"

    'Danh mục chức vụ trong cơ quan --> Tham chiếu tới bảng danh mục với id_goc = 14
    Public Const Sql_Chucvu As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 14 and Status = 1 Order by ten_goi Asc"

    'Phân loại - Lực lượng vũ trang  -->  Tham chiếu tới bảng danh mục với id_goc = 41
    Public Const Sql_PL_LLVT As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 41 and Status = 1 Order by ten_goi Asc"

    'Danh mục các loại Quân hàm  -->  Tham chiếu tới bảng danh mục với id_goc = 2
    Public Const Sql_Quanham As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 2 and Status = 1 Order by ten_goi Asc"

    'Loại quyết định nhân sự  --> Tham chiếu tới bảng danh mục với id_goc = 15
    'Public Const Sql_LoaiQd_Ns As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 15 and Status = 1 Order by ten_goi Asc"
    Public Const Sql_LoaiQd_Ns As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 15 and ma_so not in ('1510', '1511', '1506', '1505', '1512') and Status = 1 Order by ten_goi Asc"

    'Danh mục các Chuyên môn nghiệp vụ --> Tham chiếu tới bảng danh mục với id_goc = 12
    Public Const Sql_Chuyenmon As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 12 and Status = 1 Order by ten_goi Asc"

    'Danh mục Chức vụ trong Đảng --> Tham chiếu tới bảng danh mục với id_goc = 16
    Public Const Sql_Cv_dang As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 16 and Status = 1 Order by ten_goi Asc"

    'Danh mục Chức vụ trong Đoàn --> Tham chiếu tới bảng danh mục với id_goc = 10
    Public Const Sql_Cv_doan As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 10 and Status = 1 Order by ten_goi Asc"

    'Danh mục Chức vụ trong Công Đoàn --> Tham chiếu tới bảng danh mục với id_goc = 17
    Public Const Sql_Cv_Congdoan As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 17 and Status = 1 Order by ten_goi Asc"

    'Loại quyết định Nghỉ việc --> Tham chiếu đến bảng danh mục với id_goc = 37
    Public Const Sql_LoaiQd_Nv As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 37 and Status = 1 Order by ten_goi Asc"

    'Loại Hợp đồng lao động --> Tham chiếu đến bảng danh mục với id_goc = 30
    Public Const Sql_LoaiHd As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 30 and Status = 1 Order by ten_goi desc"

    'Hình thức trả lương --> Tham chiếu đến bảng danh mục với id_goc = 40
    Public Const Sql_HtTrLuong As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 40 and Status = 1 Order by ten_goi Asc"

    'Loại tai nạn --> Tham chiếu đến bảng danh mục với id_goc = 32
    Public Const Sql_LoaiTn As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 32 and Status = 1 Order by ten_goi Asc"

    'Nguyên nhân tai nạn --> Tham chiếu đến bảng danh mục với id_goc = 35
    Public Const Sql_NgNhanTn As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 35 and Status = 1 Order by ten_goi Asc"

    'Loại nghỉ phép --> Tham chiếu đến bảng danh mục với id_goc = 34
    Public Const Sql_LoaiNP As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 34 and Status = 1 Order by ten_goi Asc"

    'Chuyên ngành đào thạo --> Tham chiếu đến bảng danh mục với id_goc = 11
    Public Const Sql_ChNganh As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 11 and Status = 1 Order by ten_goi Asc"

    'Mức phụ cấp --> Tham chiếu đến bảng danh mục với id_goc = 31
    Public Const Sql_MucPhCap As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 31 and Status = 1 Order by ten_goi Asc"

    'Quan hệ của cán bộ trong gia đình --> Tham chiếu đến bảng danh mục với id_goc = 23
    Public Const Sql_QuanHe As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 23 and Status = 1 Order by ten_goi Asc"

    'Trình độ chuyên môn --> Tham chiếu đến bảng danh mục với id_goc = 38
    Public Const Sql_TrdChuyenMon As String = "Select Id,Ten_Goi From DanhMuc Where id_goc = 38 and Status = 1"

    Public Const Sql_TTrangHonNhan As String = "Select Ma_So,Ten_Goi from DanhMuc Where id_goc = 835 and Status = 1 Order By Ma_So"

    Public Const Sql_LoaiChiLuongALL As String = "Select ValueKey Code,ValueDesc Name,ValueName,OrderNo,Status,ListKey From SysListValue Where Status=1 And ListKey='LOAICHILUONG' Order By OrderNo"

    Public Const Sql_LoaiChiLuongBoEOD As String = "Select ValueKey Code,ValueDesc Name,ValueName,OrderNo,Status,ListKey From SysListValue Where Status=1 And ListKey='LOAICHILUONG' And ValueKey <> '04' Order By OrderNo"
#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"

    ''' <summary>
    ''' Class khởi tạo phần danh mục và Địa danh
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DanhMuc
        Private _Id As Integer
        Private _RootId As Integer
        Private _Code As String
        Private _Name As String
        Private _Status As Integer

        Public Sub New()
            _Id = 0
            _RootId = 0
            _Code = ""
            _Name = ""
            _Status = 1
        End Sub

        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property RootId() As Integer
            Get
                Return _RootId
            End Get
            Set(ByVal value As Integer)
                _RootId = value
            End Set
        End Property

        Public Property Code() As String
            Get
                Return _Code
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 8) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Code = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Name = value
            End Set
        End Property

        Public Property Status() As Integer
            Get
                Return _Status
            End Get
            Set(ByVal value As Integer)
                _Status = value
            End Set
        End Property
    End Class

    Public Class QuocGia
        Private _Id As Integer
        Private _Code As String
        Private _Name As String
        Private _Status As String

        Public Sub New()
            _Id = 0
            _Code = ""
            _Name = ""
            _Status = 1
        End Sub

        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property Code() As String
            Get
                Return _Code
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 8) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Code = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Name = value
            End Set
        End Property

        Public Property Status() As Integer
            Get
                Return _Status
            End Get
            Set(ByVal value As Integer)
                _Status = value
            End Set
        End Property
    End Class

    Public Class ChiNhanh
        Private _Id As Integer
        Private _RootId As Integer
        Private _Code As String
        Private _Name As String
        Private _AliasBranch As String
        Private _Address As String
        Private _Tel As String
        Private _Email As String
        Private _Web As String
        Private _Status As String
        Private _Villagers As Integer

        Public Sub New()
            _Id = 0
            _RootId = 0
            _Code = ""
            _Name = ""
            _AliasBranch = ""
            _Address = ""
            _Tel = ""
            _Email = ""
            _Web = ""
            _Status = 1
            _Villagers = 0
        End Sub

        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property RootId() As Integer
            Get
                Return _RootId
            End Get
            Set(ByVal value As Integer)
                _RootId = value
            End Set
        End Property

        Public Property Code() As String
            Get
                Return _Code
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 8) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Code = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Name = value
            End Set
        End Property

        Public Property AliasBranch() As String
            Get
                Return _AliasBranch
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên viết tắt có độ dài không hợp lệ!", value, value.ToString())
                End If
                _AliasBranch = value
            End Set
        End Property

        Public Property Tel() As String
            Get
                Return _Tel
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Tel = value
            End Set
        End Property

        Public Property Email() As String
            Get
                Return _Email
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ e-mail có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Email = value
            End Set
        End Property

        Public Property Address() As String
            Get
                Return _Address
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Address = value
            End Set
        End Property

        Public Property Web() As String
            Get
                Return _Web
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ website có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Web = value
            End Set
        End Property

        Public Property Status() As Integer
            Get
                Return _Status
            End Get
            Set(ByVal value As Integer)
                _Status = value
            End Set
        End Property

        Public Property Villagers() As Integer
            Get
                Return _Villagers
            End Get
            Set(ByVal value As Integer)
                _Villagers = value
            End Set
        End Property
    End Class

    Public Class DiaDanh
        Private _Id As Integer
        Private _RootId As Integer
        Private _Code As String
        Private _Name As String
        Private _Status As String

        Public Sub New()
            _Id = 0
            _RootId = 0
            _Code = ""
            _Name = ""
            _Status = 1
        End Sub

        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property RootId() As Integer
            Get
                Return _RootId
            End Get
            Set(ByVal value As Integer)
                _RootId = value
            End Set
        End Property

        Public Property Code() As String
            Get
                Return _Code
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 7) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Code = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Name = value
            End Set
        End Property

        Public Property Status() As Integer
            Get
                Return _Status
            End Get
            Set(ByVal value As Integer)
                _Status = value
            End Set
        End Property
    End Class

    Public Class PhongBan
        Private _Id As Integer
        Private _Code As String
        Private _Name As String
        Private _TrucThuoc As String
        Private _Status As String

        Public Sub New()
            _Id = 0
            _TrucThuoc = ""
            _Code = ""
            _Name = ""
            _Status = 1
        End Sub

        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property Code() As String
            Get
                Return _Code
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 5) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã hiệu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Code = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Name = value
            End Set
        End Property

        Public Property TrucThuoc() As String
            Get
                Return _TrucThuoc
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 10) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị trực thuộc có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TrucThuoc = value
            End Set
        End Property

        Public Property Status() As Integer
            Get
                Return _Status
            End Get
            Set(ByVal value As Integer)
                _Status = value
            End Set
        End Property
    End Class

    Public Class NghiDinhLuong
        Private _IdNDLuong As Integer
        Private _MaND As String
        Private _SoND As String
        Private _TenND As String
        Private _LoaiND As String
        Private _GhiChu As String
        Private _Status As Byte

        Public Sub New()
            _IdNDLuong = 0
            _MaND = ""
            _SoND = ""
            _TenND = ""
            _LoaiND = ""
            _GhiChu = ""
            _Status = 0
        End Sub

        Public Property IdNDLuong() As Integer
            Get
                Return _IdNDLuong
            End Get
            Set(ByVal value As Integer)
                _IdNDLuong = value
            End Set
        End Property

        Public Property MaND() As String
            Get
                Return _MaND
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã nghị định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MaND = value
            End Set
        End Property

        Public Property SoND() As String
            Get
                Return _SoND
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số nghị định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoND = value
            End Set
        End Property

        Public Property TenND() As String
            Get
                Return _TenND
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên nghị định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TenND = value
            End Set
        End Property

        Public Property LoaiND() As String
            Get
                Return _LoaiND
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị loại nghị định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LoaiND = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 255) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú nghị định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

        Public Property Status() As Byte
            Get
                Return _Status
            End Get
            Set(ByVal value As Byte)
                _Status = value
            End Set
        End Property
    End Class

    Public Class BangLuong
        Private _IdBangLuong As Integer
        Private _IdND_Luong As Integer
        Private _BangLuong As String
        Private _MoTa As String
        Private _Status As Byte

        Public Sub New()
            _IdBangLuong = 0
            _IdND_Luong = 0
            _BangLuong = ""
            _MoTa = ""
            _Status = 0
        End Sub

        Public Property IdBangLuong() As Integer
            Get
                Return _IdBangLuong
            End Get
            Set(ByVal value As Integer)
                _IdBangLuong = value
            End Set
        End Property

        Public Property IdND_Luong() As Integer
            Get
                Return _IdND_Luong
            End Get
            Set(ByVal value As Integer)
                _IdND_Luong = value
            End Set
        End Property

        Public Property BangLuong() As String
            Get
                Return _BangLuong
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 5) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị bảng lương có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BangLuong = value
            End Set
        End Property

        Public Property MoTa() As String
            Get
                Return _MoTa
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mô tả bảng lương có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MoTa = value
            End Set
        End Property

        Public Property Status() As Byte
            Get
                Return _Status
            End Get
            Set(ByVal value As Byte)
                _Status = value
            End Set
        End Property
    End Class

    Public Class NgachLuong
        Private _IdNgachLuong As Integer
        Private _IdBangLuong As Integer
        Private _NgachLuong As String
        Private _Loai As String
        Private _Mota As String
        Private _TimeNangNgach As Integer
        Private _Status As Byte

        Public Sub New()
            _IdNgachLuong = 0
            _IdBangLuong = 0
            _NgachLuong = ""
            _Loai = ""
            _Mota = ""
            _TimeNangNgach = 0
            _Status = 0
        End Sub

        Public Property IdNgachLuong() As Integer
            Get
                Return _IdNgachLuong
            End Get
            Set(ByVal value As Integer)
                _IdNgachLuong = value
            End Set
        End Property

        Public Property IdBangLuong() As Integer
            Get
                Return _IdBangLuong
            End Get
            Set(ByVal value As Integer)
                _IdBangLuong = value
            End Set
        End Property

        Public Property NgachLuong() As String
            Get
                Return _NgachLuong
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 5) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ngạch lương có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NgachLuong = value
            End Set
        End Property

        Public Property Loai() As String
            Get
                Return _Loai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 10) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị phân loại ngạch lương có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Loai = value
            End Set
        End Property

        Public Property Mota() As String
            Get
                Return _Mota
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mô tả chi tiết ngạch lương có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Mota = value
            End Set
        End Property

        Public Property TimeNangNgach() As Integer
            Get
                Return _TimeNangNgach
            End Get
            Set(ByVal value As Integer)
                _TimeNangNgach = value
            End Set
        End Property

        Public Property Status() As Byte
            Get
                Return _Status
            End Get
            Set(ByVal value As Byte)
                _Status = value
            End Set
        End Property
    End Class

    Public Class BacLuong
        Private _IdBacLuong As Integer
        Private _IdNgachLuong As Integer
        Private _BacLuong As Integer
        Private _Heso As Double
        Private _Mota As String
        Private _Status As Byte

        Public Sub New()
            _IdBacLuong = 0
            _IdNgachLuong = 0
            _BacLuong = 0
            _Heso = 0
            _Mota = ""
            _Status = 0
        End Sub

        Public Property IdBacLuong() As Integer
            Get
                Return _IdBacLuong
            End Get
            Set(ByVal value As Integer)
                _IdBacLuong = value
            End Set
        End Property

        Public Property IdNgachLuong() As Integer
            Get
                Return _IdNgachLuong
            End Get
            Set(ByVal value As Integer)
                _IdNgachLuong = value
            End Set
        End Property

        Public Property BacLuong() As Integer
            Get
                Return _BacLuong
            End Get
            Set(ByVal value As Integer)
                _BacLuong = value
            End Set
        End Property

        Public Property Heso() As Double
            Get
                Return _Heso
            End Get
            Set(ByVal value As Double)
                _Heso = value
            End Set
        End Property

        Public Property Mota() As String
            Get
                Return _Mota
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mô tả chi tiết bậc lương có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Mota = value
            End Set
        End Property

        Public Property Status() As Byte
            Get
                Return _Status
            End Get
            Set(ByVal value As Byte)
                _Status = value
            End Set
        End Property
    End Class

    Public Class TienLuong
        Private _IdTienLuong As Integer
        Private _LuongCoBan As Double
        Private _HeSoNganh As Double
        Private _NgayHuong As DateTime
        Private _SoQD As String
        Private _NgayQD As DateTime
        Private _NguoiQD As String
        Private _IdCV_Nguoi_QD As Integer
        Private _Ghichu As String
        Private _Status As Byte

        Public Sub New()
            _IdTienLuong = 0
            _LuongCoBan = 0
            _HeSoNganh = 0
            _NgayHuong = DateTime.Now
            _SoQD = ""
            _NgayQD = DateTime.Now
            _NguoiQD = ""
            _IdCV_Nguoi_QD = 0
            _Ghichu = ""
            _Status = 0
        End Sub

        Public Property IdTienLuong() As Integer
            Get
                Return _IdTienLuong
            End Get
            Set(ByVal value As Integer)
                _IdTienLuong = value
            End Set
        End Property

        Public Property LuongCoBan() As Double
            Get
                Return _LuongCoBan
            End Get
            Set(ByVal value As Double)
                _LuongCoBan = value
            End Set
        End Property

        Public Property HeSoNganh() As Double
            Get
                Return _HeSoNganh
            End Get
            Set(ByVal value As Double)
                _HeSoNganh = value
            End Set
        End Property

        Public Property NgayHuong() As DateTime
            Get
                Return _NgayHuong
            End Get
            Set(ByVal value As DateTime)
                _NgayHuong = value
            End Set
        End Property

        Public Property SoQD() As String
            Get
                Return _SoQD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoQD = value
            End Set
        End Property

        Public Property NgayQD() As DateTime
            Get
                Return _NgayQD
            End Get
            Set(ByVal value As DateTime)
                _NgayQD = value
            End Set
        End Property

        Public Property NguoiQD() As String
            Get
                Return _NguoiQD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người quyết định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NguoiQD = value
            End Set
        End Property

        Public Property IdCV_Nguoi_QD() As Integer
            Get
                Return _IdCV_Nguoi_QD
            End Get
            Set(ByVal value As Integer)
                _IdCV_Nguoi_QD = value
            End Set
        End Property

        Public Property Ghichu() As String
            Get
                Return _Ghichu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Ghichu = value
            End Set
        End Property

        Public Property Status() As Byte
            Get
                Return _Status
            End Get
            Set(ByVal value As Byte)
                _Status = value
            End Set
        End Property
    End Class

    Public Class MucPhuCap
        Private _IdMuc_PhC As Integer
        Private _IdLoai_PhC As Integer
        Private _Muc_PhC As Double
        Private _Status As Byte

        Public Sub New()
            _IdMuc_PhC = 0
            _IdLoai_PhC = 0
            _Muc_PhC = 0
            _Status = 0
        End Sub

        Public Property IdMuc_PhC() As Integer
            Get
                Return _IdMuc_PhC
            End Get
            Set(ByVal value As Integer)
                _IdMuc_PhC = value
            End Set
        End Property

        Public Property IdLoai_PhC() As Integer
            Get
                Return _IdLoai_PhC
            End Get
            Set(ByVal value As Integer)
                _IdLoai_PhC = value
            End Set
        End Property

        Public Property Muc_PhC() As Double
            Get
                Return _Muc_PhC
            End Get
            Set(ByVal value As Double)
                _Muc_PhC = value
            End Set
        End Property

        Public Property Status() As Byte
            Get
                Return _Status
            End Get
            Set(ByVal value As Byte)
                _Status = value
            End Set
        End Property
    End Class

#End Region

#Region "---> Hàm định nghĩa lưới dữ liệu <---"
    ''' <summary>
    ''' Hàm định nghĩa cấu trúc lưới dữ liệu hiển thị - Phân hệ danh mục
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu truyền vào</param>
    ''' <param name="_State">
    ''' Biến lưu chỉ số xác định Lưới dữ liệu. Với 
    '''                    _State =  1 Or _State = 3: Danh mục chung và Địa danh
    '''                    _State =  2: Chi nhánh
    '''                    _State =  4: Nghị định lương
    '''                    _State =  5: Bảng lương
    '''                    _State =  6: Ngạch lương
    '''                    _State =  7: Bậc lương
    '''                    _State =  8: Lương cơ bản
    '''                    _State =  9: Mức phụ cấp
    '''                    _State = 10: Tham số Hệ thống
    '''                    _State = 11: Khai báo tham số hệ thống Lương
    '''                    _State = 12: Nhập thông tin Tham số của Chi nhánh trên lưới
    ''' </param>
    ''' <remarks>Eg: Create_Frame(dgv_name, True)</remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal _State As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        Select Case _State
            Case 1      'Danh mục chung
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Name", "Tên gọi")
                dgv_name.Columns.Add("cln_Code", "Mã số")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Name").Width = 245
                dgv_name.Columns("cln_Code").Width = 70
                dgv_name.Columns("cln_Status").Width = 113

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns("cln_Id").ReadOnly = True
                dgv_name.Columns("cln_STT").ReadOnly = True
                dgv_name.Columns("cln_Name").ReadOnly = True
                dgv_name.Columns("cln_Code").ReadOnly = True
                dgv_name.Columns("cln_Status").ReadOnly = True
                'Ẩn cột đầu tiên lưu chỉ số id xác định tính duy nhất của ban ghi
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                'Xét quyền xoá bản ghi của người dùng
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";58;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 2      'Chi nhánh
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Name", "Tên gọi")
                dgv_name.Columns.Add("cln_Code", "Mã số")
                dgv_name.Columns.Add("cln_Alias", "Tên viết tắt")
                dgv_name.Columns.Add("cln_Address", "Địa chỉ")
                dgv_name.Columns.Add("cln_Tel", "Điện thoại")
                dgv_name.Columns.Add("cln_Email", "Địa chỉ e-mail")
                dgv_name.Columns.Add("cln_Web", "Địa chỉ website")
                dgv_name.Columns.Add("cln_Villagers", "Số xã phường")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Name").Width = 201
                dgv_name.Columns("cln_Code").Width = 70
                dgv_name.Columns("cln_Alias").Width = 78
                dgv_name.Columns("cln_Address").Width = 210
                dgv_name.Columns("cln_Tel").Width = 80
                dgv_name.Columns("cln_Email").Width = 140
                dgv_name.Columns("cln_Web").Width = 140
                dgv_name.Columns("cln_Villagers").Width = 90
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 11
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_Tel").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Villagers").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns(0).Visible = False 'Ẩn cột dữ liệu Id
                'Xét quyền xoá bản ghi của người dùng
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";62;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 3      'Địa danh
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Name", "Tên gọi")
                dgv_name.Columns.Add("cln_Code", "Mã số")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Name").Width = 245
                dgv_name.Columns("cln_Code").Width = 70
                dgv_name.Columns("cln_Status").Width = 113

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns("cln_Id").ReadOnly = True
                dgv_name.Columns("cln_STT").ReadOnly = True
                dgv_name.Columns("cln_Name").ReadOnly = True
                dgv_name.Columns("cln_Code").ReadOnly = True
                dgv_name.Columns("cln_Status").ReadOnly = True
                'Ẩn cột đầu tiên lưu chỉ số id xác định tính duy nhất của ban ghi
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";66;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 4      'Định nghĩa lưới dữ liệu - Nghị định lương
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_Code", "Mã nghị định")
                dgv_name.Columns.Add("cln_SoNgd", "Số nghị định")
                dgv_name.Columns.Add("cln_TenNd", "Tên nghị định")
                dgv_name.Columns.Add("cln_LoaiNd", "Loại nghị định")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")

                dgv_name.Columns("cln_Code").Width = 88
                dgv_name.Columns("cln_SoNgd").Width = 95
                dgv_name.Columns("cln_TenNd").Width = 125
                dgv_name.Columns("cln_LoaiNd").Width = 91
                dgv_name.Columns("cln_Ghichu").Width = 280
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Ẩn cột đầu tiên lưu chỉ số id xác định tính duy nhất của ban ghi
                dgv_name.Columns(0).Visible = False   'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 5      'Định nghĩa lưới dữ liệu - Bảng lương
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_NgdLuong", "Nghị định lương")
                dgv_name.Columns.Add("cln_BangLuong", "Bảng lương")
                dgv_name.Columns.Add("cln_Mota", "Mô tả chi tiết")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_NgdLuong").Width = 125
                dgv_name.Columns("cln_BangLuong").Width = 83
                dgv_name.Columns("cln_Mota").Width = 270
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 6
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 6      'Định nghĩa lưới dữ liệu - Ngạch lương
                dgv_name.Columns.Add("cln_Id", "Id")

                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_BangLuong", "Bảng lương")
                dgv_name.Columns.Add("cln_NgLuong", "Ngạch lương")
                dgv_name.Columns.Add("cln_Phanloai", "Phân loại")
                dgv_name.Columns.Add("cln_Mota", "Mô tả chi tiết")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_BangLuong").Width = 80
                dgv_name.Columns("cln_NgLuong").Width = 85
                dgv_name.Columns("cln_Phanloai").Width = 75
                dgv_name.Columns("cln_Mota").Width = 230
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 7      'Định nghĩa lưới dữ liệu - Bậc lương
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_NgLuong", "Ngạch lương")
                dgv_name.Columns.Add("cln_BacLuong", "Bậc lương")
                dgv_name.Columns.Add("cln_HeSo", "Hệ số")
                dgv_name.Columns.Add("cln_Mota", "Mô tả chi tiết")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")


                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_NgLuong").Width = 85
                dgv_name.Columns("cln_BacLuong").Width = 80
                dgv_name.Columns("cln_HeSo").Width = 77
                dgv_name.Columns("cln_Mota").Width = 120
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_HeSo").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 8      'Định nghĩa lưới dữ liệu - Lương cơ bản
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_LuongCb", "Lương tối thiểu")
                dgv_name.Columns.Add("cln_HsNganh", "Hệ số ngành")
                dgv_name.Columns.Add("cln_NgApdung", "Ngày áp dụng")
                dgv_name.Columns.Add("cln_SoQd", "Số quyết định")
                dgv_name.Columns.Add("cln_NgayQd", "Ngày quyết định")
                dgv_name.Columns.Add("cln_NguoiQd", "Người ký quyết định")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ người ký")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")

                'Hiệu chỉnh độ rộng của các columns
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_LuongCb").Width = 100
                dgv_name.Columns("cln_HsNganh").Width = 83
                dgv_name.Columns("cln_NgApdung").Width = 88
                dgv_name.Columns("cln_SoQd").Width = 105
                dgv_name.Columns("cln_NgayQd").Width = 103
                dgv_name.Columns("cln_NguoiQd").Width = 127
                dgv_name.Columns("cln_Chucvu").Width = 135
                dgv_name.Columns("cln_Ghichu").Width = 145
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 11
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_LuongCb").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_HsNganh").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_NgApdung").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NgayQd").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 9       'Định nghĩa lưới dữ liệu - Mức phụ cấp
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_LoaiPc", "Loại phụ cấp")
                dgv_name.Columns.Add("cln_MucPc", "Mức phụ cấp")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_LoaiPc").Width = 150
                dgv_name.Columns("cln_MucPc").Width = 83
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns("cln_Id").ReadOnly = True
                dgv_name.Columns("cln_STT").ReadOnly = True
                dgv_name.Columns("cln_LoaiPc").ReadOnly = True
                dgv_name.Columns("cln_MucPc").ReadOnly = True
                dgv_name.Columns("cln_Status").ReadOnly = True
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_MucPc").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns(0).Visible = False 'Ẩn cột dữ liệu Id 
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If
            Case 10      'Định nghĩa lưới dữ liệu - Khai báo tham số hệ thống
                dgv_name.Columns.Add("cln_Id", "Id")
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Name", "Tên tham số")
                dgv_name.Columns.Add("cln_Value", "Giá trị tham số")
                dgv_name.Columns.Add("cln_Discript", "Mô tả chi tiết")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Name").Width = 120
                dgv_name.Columns("cln_Value").Width = 120
                dgv_name.Columns("cln_Discript").Width = 390

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                dgv_name.Columns(1).ReadOnly = True
                dgv_name.Columns(2).ReadOnly = True
                dgv_name.Columns(3).ReadOnly = True
                dgv_name.Columns(4).ReadOnly = True
                dgv_name.Columns(0).Visible = False

                'Không cho sắp xếp theo cột
                dgv_name.Columns("cln_Id").SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns("cln_STT").SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns("cln_Name").SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns("cln_Value").SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns("cln_Discript").SortMode = DataGridViewColumnSortMode.NotSortable
            Case 11      'Định nghĩa lưới dữ liệu - Khai báo tham số hệ thống Lương
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Name", "Tên tham số")
                dgv_name.Columns.Add("cln_Value", "Giá trị tham số")
                dgv_name.Columns.Add("cln_Rate", "Tỷ lệ tương ứng với giá trị")
                dgv_name.Columns.Add("cln_DateApply", "Ngày áp dụng")
                dgv_name.Columns.Add("cln_Discript", "Mô tả chi tiết")
                dgv_name.Columns.Add("cln_Status", "Trạng thái bản ghi")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Name").Width = 140
                dgv_name.Columns("cln_Value").Width = 230
                dgv_name.Columns("cln_Rate").Width = 230
                dgv_name.Columns("cln_DateApply").Width = 88
                dgv_name.Columns("cln_Discript").Width = 280
                dgv_name.Columns("cln_Status").Width = 113
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Integer = 2 To 8
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_DateApply").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                If (DONVI <> gMaDonViTW) Then
                    dgv_name.Columns("cln_Choice").Visible = False
                Else
                    Dim _roles As String = Globals.Roles
                    If (_roles.IndexOf(";70;") < 0) Then
                        dgv_name.Columns("cln_Choice").Visible = False
                    End If
                End If

            Case 12      'Định nghĩa lưới dữ liệu - Tham số hệ thống
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_Id_DonVi", "DonViId")
                dgv_name.Columns.Add("cln_TrucThuoc", "Trực thuộc")
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_DonVi_Cd", "Mã đơn vị")
                dgv_name.Columns.Add("cln_DonVi_HT", "Tên đơn vị")
                dgv_name.Columns.Add("cln_Ten_VT", "Tên VT")
                dgv_name.Columns.Add("cln_DiaBan", "Tên địa bàn")
                dgv_name.Columns.Add("cln_Max_Ky", "Số kỳ trả lương (tháng)")
                dgv_name.Columns.Add("cln_Tinh_Thue_TNCN", "Tính thuế TNCN")
                dgv_name.Columns.Add("cln_VungLuongTT", "Vùng lương tối thiểu")
                dgv_name.Columns.Add("cln_MucTamUng_V2", "Mức tạm ứng (%) tiền lương V2")
                dgv_name.Columns.Add("cln_HeSoTamUng_V2", "Hệ số tạm ứng lương V2 hàng tháng ")
                dgv_name.Columns.Add("cln_GiamDoc", IIf(Cap_Nd = 1, "Tổng Giám đốc", "Giám đốc"))
                dgv_name.Columns.Add("cln_PhoGiamDoc", IIf(Cap_Nd = 1, "Phó Tổng Giám đốc", "Phó Giám đốc"))
                dgv_name.Columns.Add("cln_TruongHCTC", IIf(Cap_Nd = 1, "Giám đốc Ban TCCB", "TP.HCTC"))
                dgv_name.Columns.Add("cln_TruongKT", IIf(Cap_Nd = 1, "Giám đốc Ban KTTC", "TP. Kế toán/Trưởng KT"))
                dgv_name.Columns.Add("cln_NgayHL", "Ngày áp dụng")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_DonVi_Cd").Width = 50
                dgv_name.Columns("cln_DonVi_HT").Width = 230
                dgv_name.Columns("cln_Ten_VT").Width = 50
                dgv_name.Columns("cln_DiaBan").Width = 80
                dgv_name.Columns("cln_Max_Ky").Width = 60
                dgv_name.Columns("cln_Tinh_Thue_TNCN").Width = 60
                dgv_name.Columns("cln_VungLuongTT").Width = 80
                dgv_name.Columns("cln_GiamDoc").Width = 140
                dgv_name.Columns("cln_PhoGiamDoc").Width = 140
                dgv_name.Columns("cln_TruongHCTC").Width = 140
                dgv_name.Columns("cln_TruongKT").Width = 140
                dgv_name.Columns("cln_NgayHL").Width = 80
                dgv_name.Columns("cln_MucTamUng_V2").Width = 80
                dgv_name.Columns("cln_HeSoTamUng_V2").Width = 80
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Integer = 2 To 18
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DonVi_Cd").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Max_Ky").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Tinh_Thue_TNCN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_VungLuongTT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NgayHL").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_MucTamUng_V2").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_HeSoTamUng_V2").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False  'Ẩn cột dữ liệu Id
                dgv_name.Columns(2).Visible = False  'Ẩn cột dữ liệu DonViId
                dgv_name.Columns(3).Visible = False  'Ẩn cột dữ liệu TrucThuoc
                dgv_name.ColumnHeadersHeight = 35
                dgv_name.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

        End Select
    End Sub
#End Region

#Region "---> Các Hàm lấy dữ liệu <---"
    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào treeview - Phân hệ danh mục chung
    ''' </summary>
    ''' <param name="tv_name">Tên cây dữ liệu</param>
    ''' <param name="_State">
    ''' Chỉ số xác định cây dữ liệu
    '''                   1: Danh mục chung
    '''                   2: Chi nhánh
    '''                   3: Địa danh
    ''' </param>
    ''' <remarks></remarks>
    Public Sub FillData_TreeView(ByVal tv_name As System.Windows.Forms.TreeView, ByVal _State As Byte)
        tv_name.Nodes.Clear()
        Dim tn_parent As TreeNode
        Dim tn_child As TreeNode
        Dim strSQL As String = ""
        Select Case _State
            Case 1          'FILL: dữ liệu vào cây Danh mục chung
                strSQL = "Select id, ten_goi from DanhMuc Where id_goc = 0 and status=1 order by ten_goi"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                tn_parent = tv_name.Nodes.Add(db.Rows(i)("ten_goi").ToString())
                                tn_parent.Tag = "00DM" + db.Rows(i)("id").ToString().Trim()
                                strSQL = String.Format("Select id, ten_goi from DanhMuc Where id_goc = {0} and id_goc != 0 Order by ten_goi Asc", CType(db.Rows(i)("id"), Integer))
                                Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                    If Not (db_child Is Nothing) Then
                                        If (db_child.Rows.Count > 0) Then
                                            For j As Integer = 0 To (db_child.Rows.Count - 1)
                                                tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("ten_goi").ToString())
                                                tn_child.Tag = "01DM" + db_child.Rows(j)("id").ToString().Trim()
                                            Next
                                        End If
                                    End If
                                End Using
                            Next
                        End If
                    End If
                End Using

                'Add thêm danh mục phòng ban vào Treeview hệ thống danh mục chung
                tn_parent = tv_name.Nodes.Add("Phòng ban")
                tn_parent.Tag = "00PB"
                strSQL = String.Format("Select * from PhongBan order by ten_phong asc ")
                Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db_child Is Nothing) Then
                        If (db_child.Rows.Count > 0) Then
                            For j As Integer = 0 To (db_child.Rows.Count - 1)
                                tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("ten_phong").ToString())
                                tn_child.Tag = "01PB" + db_child.Rows(j)("id").ToString().Trim()
                            Next
                        End If
                    End If
                End Using

                'Add thêm danh mục - Quốc gia vào treeview Hệ thống danh mục chung
                tn_parent = tv_name.Nodes.Add("Quốc gia")
                tn_parent.Tag = "00QG"
                strSQL = String.Format("Select Id,Ten_Goi From QuocGia Order by Ma_so asc")
                Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db_child Is Nothing) Then
                        If (db_child.Rows.Count > 0) Then
                            For j As Integer = 0 To (db_child.Rows.Count - 1)
                                tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("ten_goi").ToString())
                                tn_child.Tag = "01QG" + db_child.Rows(j)("id").ToString().Trim()
                            Next
                        End If
                    End If
                End Using

            Case 2          'FILL: dữ liệu vào cây Chi nhánh
                'Thực hiện add node gốc "Ngân hàng chính sách xã hội" trước
                Dim tn_node As TreeNode
                tn_parent = tv_name.Nodes.Add("Ngân hàng chính sách xã hội")
                tn_parent.Tag = "NHCSXH"

                'Add node child - Hội sở chính trước (Mục đích việc add riêng node HSC trước để add các phòng ban trực thuộc vào node con của HSC luôn)
                strSQL = String.Format("Select id, ten_goi from ChiNhanh Where id_goc = 0")
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            'Add node child - các phòng ban trực thuộc hội sở chính
                            For i As Integer = 0 To (db.Rows.Count - 1)
                                tn_child = tn_parent.Nodes.Add(db.Rows(i)("ten_goi").ToString())
                                tn_child.Tag = "00CN1" + db.Rows(i)("id").ToString().Trim()
                                strSQL = String.Format("Select * From PhongBan Where Charindex('1',truc_thuoc) > 0 Order by Ma_so")
                                Using db_pb As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                    If Not (db_pb Is Nothing) Then
                                        If (db_pb.Rows.Count > 0) Then
                                            For k As Integer = 0 To (db_pb.Rows.Count - 1)
                                                tn_node = tn_child.Nodes.Add(db_pb.Rows(k)("ten_phong").ToString())
                                                tn_node.Tag = "01CN1" + db_pb.Rows(k)("id").ToString().Trim()
                                            Next
                                        End If
                                    End If
                                End Using
                            Next

                            'Thực hiện add các node child - Chi nhánh các tỉnh, các trung tâm, sở giao dịch, văn phòng miền trực thuộc
                            For i As Integer = 0 To (db.Rows.Count - 1)
                                strSQL = String.Format("Select id, ten_goi from ChiNhanh Where id_goc = {0} and id_goc != 0", CType(db.Rows(i)("id"), Integer))
                                Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                    If Not (db_child Is Nothing) Then
                                        If (db_child.Rows.Count > 0) Then
                                            For j As Integer = 0 To (db_child.Rows.Count - 1)
                                                tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("ten_goi").ToString())
                                                tn_child.Tag = "00CN2" + db_child.Rows(j)("id").ToString().Trim()
                                                strSQL = String.Format("Select Id,Ten_Goi From ChiNhanh where id_goc = {0} and id_goc != 0", CType(db_child.Rows(j)("id"), Integer))
                                                Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                                    If Not (_db Is Nothing) Then
                                                        If (_db.Rows.Count > 0) Then
                                                            For k As Integer = 0 To (_db.Rows.Count - 1)
                                                                tn_node = tn_child.Nodes.Add(_db.Rows(k)("ten_goi").ToString())
                                                                tn_node.Tag = "01CN2" + _db.Rows(k)("id").ToString().Trim()
                                                            Next
                                                        End If
                                                    End If
                                                End Using
                                            Next
                                        End If
                                    End If
                                End Using
                            Next
                        End If
                    End If
                End Using

            Case 3          'FILL: dữ liệu vào cây Địa danh (Cây dữ liệu 3 cấp: Tỉnh -> Huyện -> Xã)
                Dim tn_node As TreeNode
                strSQL = "Select id, ten_goi From DiaDanh Where id_goc = 0 Order by Ma_so Asc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim _rows As Int32 = db.Rows.Count - 1
                            For i As Integer = 0 To _rows
                                tn_parent = tv_name.Nodes.Add(db.Rows(i)("ten_goi").ToString())
                                tn_parent.Tag = "00DD" + db.Rows(i)("id").ToString().Trim()
                                strSQL = String.Format("Select Id,Ten_Goi From DiaDanh Where id_goc = {0} and id_goc != 0 Order by Ma_so asc", Convert.ToInt32(db.Rows(i)("id")))
                                Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                    If Not (db_child Is Nothing) Then
                                        If (db_child.Rows.Count > 0) Then
                                            For j As Integer = 0 To (db_child.Rows.Count - 1)
                                                tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("ten_goi").ToString())
                                                tn_child.Tag = "01DD" + db_child.Rows(j)("id").ToString().Trim()
                                                'Fill dữ liệu Xã Phường thuộc Quận huyện đang duyệt
                                                strSQL = String.Format("Select Id,Ten_Goi From DiaDanh Where id_goc = {0} and id_goc != 0 Order by Ma_so asc", Convert.ToInt32(db_child.Rows(j)("id")))
                                                Using db_childXP As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                                    If Not (db_childXP Is Nothing) Then
                                                        If (db_childXP.Rows.Count > 0) Then
                                                            For k As Integer = 0 To (db_childXP.Rows.Count - 1)
                                                                tn_node = tn_child.Nodes.Add(db_childXP.Rows(k)("ten_goi").ToString())
                                                                tn_node.Tag = "02DD" + db_childXP.Rows(k)("id").ToString().Trim()
                                                            Next
                                                        End If
                                                    End If
                                                End Using
                                            Next
                                        End If
                                    End If
                                End Using
                            Next
                        End If
                    End If
                End Using
            Case Else
        End Select
    End Sub

    ''' <summary>
    ''' Hmà trả về danh sách các bản ghi theo RoodId (chỉ số xác định danh mục cha) truyền vào
    ''' </summary>
    ''' <param name="_RootId">Chỉ số xác định danh mục gốc</param>
    ''' <param name="_FlagState">
    '''     Chỉ số xác định danh mục gốc: 
    '''                     1: Danh mục chung
    '''                     2: Chi nhánh
    '''                     3: Địa Danh
    ''' </param>
    ''' <param name="_FlagList">
    '''      Chỉ số xác định danh mục con thuộc hệ thống danh mục chung
    '''                 0: Danh mục chung
    '''                 1: Phòng ban
    '''                 2: Quốc gia        
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll(ByVal _RootId As Int32, ByVal _FlagState As Byte, ByVal _FlagList As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim strSQL As String = ""
            Select Case _FlagState
                Case 1  'Danh mục chung
                    If (_FlagList = 0) Then
                        strSQL = "Select * From DanhMuc Where id_goc != 0 and id_goc = @_RootId Order by ten_goi"
                    ElseIf (_FlagList = 1) Then
                        strSQL = "Select id, ma_so,ten_phong as ten_goi, truc_thuoc, Status From PhongBan Order by Ma_so Asc"
                    ElseIf (_FlagList = 2) Then
                        strSQL = "Select * From QuocGia Order by Ma_so Asc"
                    End If
                Case 2  'Chi nhánh
                    strSQL = "Select * From ChiNhanh Where id_goc != 0 and id_goc = @_RootId Order by id_goc,Ma_so Asc"
                Case 3  'Địa danh
                    strSQL = "Select * From DiaDanh Where id_goc != 0 and id_goc = @_RootId Order by Ma_so Asc"
            End Select

            Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL, connection)
            command.CommandType = CommandType.Text
            If Not (_FlagList = 1 Or _FlagList = 2) Then
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_RootId", SqlDbType.Int))
                command.Parameters("@_RootId").Value = _RootId
            End If

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "ListRecords")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("ListRecords")
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
    ''' Hàm trả về row của hệ thống danh mục thoả mãn điều kiện truyền vào
    ''' </summary>
    ''' <param name="_Id">Chỉ số xác định bản ghi</param>
    ''' <param name="_FlagState">Chỉ số xác định Phân hệ. Với 1: Danh mục chung. 2: Chi nhánh. 3 Địa danh</param>
    ''' <param name="_FlagList">
    '''      Chỉ số xác định danh mục con thuộc hệ thống danh mục chung
    '''                 0: Danh mục chung
    '''                 1: Phòng ban
    '''                 2: Quốc gia        
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataForId(ByVal _Id As Integer, ByVal _FlagState As Byte, ByVal _FlagList As Byte) As DataRow
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim strSQL As String = ""
            Select Case _FlagState
                Case 1  ' Hệ thống danh mục chung
                    If (_FlagList = 0) Then
                        strSQL = "Select * From DanhMuc Where id = @_Id"
                    ElseIf (_FlagList = 1) Then
                        strSQL = "Select id, ma_so,ten_phong as ten_goi, truc_thuoc, Status From PhongBan Where id = @_Id"
                    ElseIf (_FlagList = 2) Then
                        strSQL = "Select * From QuocGia Where id = @_Id Order by Ma_so Asc"
                    End If
                Case 2  ' Chi nhánh
                    strSQL = "Select * From ChiNhanh Where id_goc != 0 And id = @_Id"
                Case 3  ' Địa danh
                    strSQL = "Select * From DiaDanh Where id_goc != 0 And id = @_Id"
            End Select

            Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL, connection)
            command.CommandType = CommandType.Text

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", SqlDbType.Int))
            command.Parameters("@_Id").Value = _Id

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "ListRecord")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("ListRecord").Rows(0)
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
    ''' Hàm trả về tên gọi của danh mục khi biết chỉ số xác định bản ghi
    ''' </summary>
    ''' <param name="_Id"></param>
    ''' <param name="_FlagState">Chỉ số xác định kho truy xuất dl. 1-Danh mục. 2-Chi nhánh. 3: Địa danh</param>
    ''' <param name="_FlagChild">Chỉ cần chú ý khi _FlagState = 1. Quy định: 0-Danh mục chung. 1-Phòng ban. 2-Quốc gia</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNameList(ByVal _Id As Integer, ByVal _FlagState As Byte, ByVal _FlagChild As Byte)
        Dim _result As String = ""
        Dim strSQL As String = ""
        Select Case _FlagState
            Case 1
                If (_FlagChild = 0) Then
                    strSQL = String.Format("Select id, ten_goi, ma_so from DanhMuc Where id = '{0}'", _Id)
                ElseIf (_FlagChild = 1) Then
                    strSQL = String.Format("Select id, ten_phong as ten_goi, ma_so from DanhMuc Where id = '{0}'", _Id)
                Else
                    strSQL = String.Format("Select id, ten_goi, ma_so from QuocGia Where id = '{0}'", _Id)
                End If
            Case 2
        End Select
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _result = db.Rows(0)("ten_goi").ToString().Trim()
                End If
            End If
        End Using
        Return _result
    End Function

    ''' <summary>
    ''' Hàm thực hiện lấy mã hiệu tự động của danh mục khi thêm mới
    ''' </summary>
    ''' <param name="_RootId">Chỉ số xác định danh mục gốc</param>
    ''' <param name="_FlagState">1: Hệ thống danh mục chung. 2: Chi nhánh. 3: Địa danh</param>
    ''' <param name="_FlagChild">Xét với _FlagState = 1 thì quy định. 0-DM chung. 1-Phòng ban 2-Quốc gia</param>
    ''' <returns>Mã hiệu gen tự động</returns>
    ''' <remarks></remarks>
    Public Function GetCode(ByVal _RootId As Integer, ByVal _FlagState As Byte, ByVal _FlagChild As Byte, ByVal _Node As String) As String
        Dim _result As String = ""
        Dim _valTemp As String = ""
        Dim strSQL As String = ""
        Select Case _FlagState
            Case 1  'Hệ thống danh mục chung (danh mục chung, phòng ban, quốc gia)
                If (_FlagChild = 0) Then
                    strSQL = String.Format("Select * From DanhMuc Where id_goc = {0} Order by Ma_so Desc", _RootId)
                ElseIf (_FlagChild = 1) Then
                    strSQL = String.Format("Select * From PhongBan Order by Ma_so Desc")
                Else
                    strSQL = String.Format("Select * From QuocGia Order by Ma_so Desc")
                End If
            Case 2  'Chi nhánh
                strSQL = String.Format("Select * From ChiNhanh Where id_goc = {0} Order by id_goc,Ma_so Desc", _RootId)
            Case 3  'Địa danh
                strSQL = String.Format("Select * From DiaDanh Where id_goc = {0} And ma_so <> '99900' And Substring(ma_so,LEN(ma_so)-1,2) <> '99' Order by Ma_so Desc", _RootId)
        End Select

        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _result = db.Rows(0)("ma_so").ToString().Trim()
                    If (_result <> "") Then
                        _valTemp = _result
                        Dim _valCode As Integer = 0
                        If (Globals.IsNumeric(_result)) Then
                            If (_FlagState = 2) Then
                                If (_RootId = 1) Then
                                    _valCode = CType(_valTemp.Substring(0, 3), Integer)
                                    _valCode += 2
                                    _result = _valCode.ToString() + "00"
                                Else
                                    _valCode = CType(_valTemp.Substring(3), Integer)
                                    _valCode += 2
                                    _result = _valTemp.Substring(0, 3) + _valCode.ToString()
                                End If
                            ElseIf (_FlagState = 3) Then
                                If (_RootId = 0) Then
                                    _valCode = CType(_valTemp.Substring(0, 3), Integer)
                                    _valCode += 2
                                    _result = _valCode.ToString() + "00"
                                Else
                                    If _valTemp.Length = 7 Then         'Truong hop sinh ma xa - phuong
                                        _valCode = CType(_valTemp.Substring(5), Integer)
                                        _valCode += 2
                                        If (_valCode.ToString().Length < 2) Then
                                            _result = _valTemp.Substring(0, 5) + "0" + _valCode.ToString()
                                        Else
                                            _result = _valTemp.Substring(0, 5) + _valCode.ToString()
                                        End If
                                    Else
                                        _valCode = CType(_valTemp.Substring(3), Integer)
                                        _valCode += 2
                                        If (_valCode.ToString().Length < 2) Then
                                            _result = _valTemp.Substring(0, 3) + "0" + _valCode.ToString()
                                        Else
                                            _result = _valTemp.Substring(0, 3) + _valCode.ToString()
                                        End If
                                    End If
                                End If
                                If (_result = db.Rows(0)("ma_so").ToString().Trim()) Then
                                    _result = "Tràn số mã hiệu địa danh. Hãy kiểm tra lại sinh mã tiếp theo!"
                                End If
                            ElseIf (_FlagState = 1) Then
                                If (_FlagChild = 0) Then
                                    _valCode = CType(_valTemp.Substring(2), Integer)
                                    _valCode += 1
                                    If (_valCode.ToString().Length < 2) Then
                                        _result = _valTemp.Substring(0, 2) + "0" + _valCode.ToString()
                                    Else
                                        _result = _valTemp.Substring(0, 2) + _valCode.ToString()
                                    End If
                                Else
                                    _valCode = CType(_valTemp, Integer)
                                    _valCode += 1
                                    _result = _valCode.ToString()
                                    While (_result.Length < _valTemp.Length)
                                        _result = "0" + _result
                                    End While
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End Using

        If (_result = "") Then
            Select Case _FlagState
                Case 2      'Chi nhánh
                    If (_RootId > 1) Then
                        strSQL = String.Format("Select * From ChiNhanh Where id = {0}", _RootId)
                        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db_child Is Nothing) Then
                                If (db_child.Rows.Count > 0) Then
                                    _result = db_child.Rows(0)("ma_so").ToString().Substring(0, 3) + "02"
                                End If
                            End If
                        End Using
                    End If
                Case 3
                    If (_RootId > 0) Then
                        strSQL = String.Format("Select * From DiaDanh Where id = {0}", _RootId)
                        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db_child Is Nothing) Then
                                If (db_child.Rows.Count > 0) Then
                                    If _Node.Substring(0, 2).ToString() = "00" Then
                                        _result = db_child.Rows(0)("ma_so").ToString().Substring(0, 3) + "01"
                                    Else
                                        _result = db_child.Rows(0)("ma_so").ToString().Substring(0, 5) + "01"
                                    End If
                                End If
                            End If
                        End Using
                    End If
            End Select
        End If
        Return _result
    End Function

    ''' <summary>
    ''' Hàm thực hiện lấy số lượng xã phường của một tỉnh
    ''' </summary>
    ''' <param name="_BranchId">Chỉ số xác định chi nhánh tỉnh</param>
    ''' <returns>Số lượng xã phường của một chi nhánh tỉnh</returns>
    ''' <remarks></remarks>
    Public Function GetVillagers(ByVal _BranchId As Integer)
        Dim _Count As Integer = 0
        Dim strSQL As String = ""
        strSQL = String.Format("SELECT SUM(solg_xaphuong) As Totals FROM ChiNhanh WHERE id_goc NOT IN (0,1) And id_goc = {0}", _BranchId)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    If (db.Rows(0)("Totals").ToString().Trim() <> "") Then
                        _Count = CType(db.Rows(0)("Totals").ToString().Trim(), Integer)
                    End If
                End If
            End If
        End Using
        Return _Count
    End Function

    ''' <summary>
    ''' Hàm trả về chỉ số tìm được của dòng đang select trước đó trên GridView
    ''' </summary>
    ''' <param name="_RowId">Id bản ghi cần tìm trên lưới</param>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="_ColumnName">Tên cột lưu chỉ số id bản ghi trong lưới dữ liệu</param>
    ''' <returns>Chỉ số tìm được của dòng đang select trước đó trên GridView</returns>
    ''' <remarks></remarks>
    Public Shared Function GetRowIndex(ByVal _RowId As String, ByVal dgv_name As DataGridView, ByVal _ColumnName As String) As Int16
        Dim _rowIndex As Int16 = 0
        If (dgv_name.Rows.Count > 0) Then
            If ((dgv_name.CurrentRow.Cells(_ColumnName).Value IsNot Nothing) And (dgv_name.CurrentRow.Cells(_ColumnName).Value.ToString() <> "")) Then
                For i As Int16 = 0 To dgv_name.Rows.Count - 1
                    If dgv_name.Rows(i).Cells(_ColumnName).Value.ToString().Trim() = _RowId Then
                        _rowIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
        Return _rowIndex
    End Function
#End Region

#Region "---> Hàm cập nhật dữ liệu liên quan đến Hệ thống danh mục <---"
    ''' <summary>
    ''' Hàm cập nhật dữ liệu danh mục
    ''' </summary>
    ''' <param name="obj_list">Object name</param>
    ''' <remarks></remarks>
    Public Sub Save_List(ByVal obj_list As DanhMuc)
        Dim strSQL As String = ""
        If (obj_list.Id > 0) Then   'Update data
            strSQL = "Update DanhMuc Set id_goc = @_RootId, ma_so = @_Code, ten_goi = @_Name, Status = @_Status Where id = @_Id"
        Else                        'Insert data
            strSQL = "Insert Into DanhMuc(id_goc,ma_so,ten_goi,Status) Values(@_RootId, @_Code, @_Name, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", obj_list.Id))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_RootId", obj_list.RootId))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Code", obj_list.Code))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Name", obj_list.Name))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu danh mục: " & ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật danh mục - Chi nhánh
    ''' </summary>
    ''' <param name="obj_branch">Object chi nhánh</param>
    ''' <remarks></remarks>
    Public Sub Save_Branch(ByVal obj_branch As ChiNhanh)
        Dim strSQL As String = ""
        If (obj_branch.Id > 0) Then   'Update data
            strSQL = "Update ChiNhanh Set id_goc = @_RootId, ma_so = @_Code, ten_goi = @_Name, ten_vt = @_Alias, dia_chi = @_Address, dien_thoai = @_Tel, email = @_Email, website = @_Web, solg_xaphuong = @_Villagers, Status = @_Status Where id = @_Id"
        Else                        'Insert data
            strSQL = "Insert Into ChiNhanh(id_goc, ma_so, ten_goi, ten_vt, dia_chi, dien_thoai, email, website, solg_xaphuong, Status) Values(@_RootId, @_Code, @_Name, @_Alias, @_Address, @_Tel, @_Email, @_Web, @_Villagers, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", obj_branch.Id))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_RootId", obj_branch.RootId))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Code", obj_branch.Code))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Name", obj_branch.Name))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Alias", obj_branch.AliasBranch))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Address", obj_branch.Address))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Tel", obj_branch.Tel))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Email", obj_branch.Email))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Web", obj_branch.Web))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_branch.Status))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Villagers", obj_branch.Villagers))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu chi nhánh: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật thông tin danh mục - Địa danh
    ''' </summary>
    ''' <param name="obj_diadanh"></param>
    ''' <remarks></remarks>
    Public Sub Save_DiaDanh(ByVal obj_diadanh As DiaDanh)
        Dim strSQL As String = ""
        If (obj_diadanh.Id > 0) Then   'Update data
            strSQL = "Update DiaDanh Set id_goc = @_RootId, ma_so = @_Code, ten_goi = @_Name, Status = @_Status Where id = @_Id"
        Else                        'Insert data
            strSQL = "Insert Into DiaDanh(id_goc,ma_so,ten_goi,Status) Values(@_RootId, @_Code, @_Name, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", obj_diadanh.Id))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_RootId", obj_diadanh.RootId))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Code", obj_diadanh.Code))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Name", obj_diadanh.Name))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_diadanh.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu địa danh: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật dữ liệu - Quốc gia
    ''' </summary>
    ''' <param name="obj_quocgia"></param>
    ''' <remarks></remarks>
    Public Sub Save_QuocGia(ByVal obj_quocgia As QuocGia)
        Dim strSQL As String = ""
        If (obj_quocgia.Id > 0) Then   'Update data
            strSQL = "Update QuocGia Set ma_so = @_Code, ten_goi = @_Name, Status = @_Status Where id = @_Id"
        Else                        'Insert data
            strSQL = "Insert Into QuocGia(ma_so,ten_goi,Status) Values(@_Code, @_Name, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", obj_quocgia.Id))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Code", obj_quocgia.Code))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Name", obj_quocgia.Name))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_quocgia.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu quốc gia: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật dữ liệu Phòng ban
    ''' </summary>
    ''' <param name="obj_dept"></param>
    ''' <remarks></remarks>
    Public Sub Save_PhongBan(ByVal obj_dept As PhongBan)
        Dim strSQL As String = ""
        If (obj_dept.Id > 0) Then   'Update data
            strSQL = "Update PhongBan Set ma_so = @_Code, ten_phong = @_Name, truc_thuoc = @_TrucThuoc, Status = @_Status Where id = @_Id"
        Else                        'Insert data
            strSQL = "Insert Into PhongBan(ma_so, ten_phong, truc_thuoc, Status) Values(@_Code, @_Name, @_TrucThuoc, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", obj_dept.Id))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Code", obj_dept.Code))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Name", obj_dept.Name))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_TrucThuoc", obj_dept.TrucThuoc))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_dept.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu phòng ban: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ' ''' <summary>
    ' ''' Hàm thực hiện cập nhật dữ liệu - Tham số Hệ thống chương trình
    ' ''' </summary>
    ' ''' <param name="_Id">Chỉ số xác định bản ghi</param>
    ' ''' <param name="_VarName">Tên gọi tham số</param>
    ' ''' <param name="_VarValue">Giá trị tham số</param>
    ' ''' <param name="_Descript">Mô tả chi tiết</param>
    ' ''' <remarks></remarks>
    'Public Sub Save_SysVar(ByVal _Id As Integer, ByVal _VarName As String, ByVal _VarValue As String, ByVal _Descript As String)
    '    Dim strSQL As String = ""
    '    If (_Id > 0) Then   'Update data
    '        strSQL = String.Format("Update SysVar Set VarName = N'{0}', VarValue = N'{1}', Descript = N'{2}' Where id = {3}", _VarName, _VarValue, _Descript, _Id)
    '    Else                        'Insert data
    '        strSQL = String.Format("Insert Into SysVar(VarName,VarValue,Descript) Values(N'{0}', N'{1}', N'{2}')", _VarName, _VarValue, _Descript)
    '    End If
    '    Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
    '    command.CommandType = CommandType.Text
    '    Try
    '        _SqlHelper.executeSQL(command)
    '    Catch ex As Exception
    '        MessageBox.Show("Cập nhật dữ liệu tham số hệ thống: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub

    ' ''' <summary>
    ' ''' Hàm cập nhật dữ liệu - Tham số hệ thống
    ' ''' </summary>
    ' ''' <param name="strSQL">Câu lệnh update dữ liệu</param>
    ' ''' <remarks></remarks>
    'Public Sub Save_SysVar(ByVal strSQL As String)
    '    Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
    '    command.CommandType = CommandType.Text
    '    Try
    '        _SqlHelper.executeSQL(command)
    '    Catch ex As Exception
    '        MessageBox.Show("Cập nhật dữ liệu tham số hệ thống: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ tham số hệ thống theo id truyền vào
    ''' </summary>
    ''' <param name="_Id"></param>
    ''' <remarks></remarks>
    Public Sub Delete_SystemVar(ByVal _Id As Integer)
        Dim strSQL As String = ""
        strSQL = String.Format("Delete SysVar Where id = {0}", _Id)
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xóa bỏ tham số hệ thống: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Tham số Hệ thống lương của chương trình
    ''' </summary>
    ''' <param name="_IdSalaryVar">Chỉ số xác định tham số hệ thống</param>
    ''' <param name="_NameVar">Tên gọi của tham số</param>
    ''' <param name="_Value">Giá trị của tham số</param>
    ''' <param name="_Rate">Tỷ lệ - tương đương với giá trị</param>
    ''' <param name="_DateApply">Ngày áp dụng</param>
    ''' <param name="_Descript">Mô tả chi tiết</param>
    ''' <param name="_Status">Trạng thái bản ghi</param>
    ''' <remarks></remarks>
    Public Sub Save_SalaryVar(ByVal _IdSalaryVar As Integer, ByVal _NameVar As String, ByVal _Value As String, ByVal _Rate As String, ByVal _DateApply As DateTime, ByVal _Descript As String, ByVal _Status As Byte)
        Dim strSQL As String = ""
        If (_IdSalaryVar > 0) Then   'Update data
            strSQL = String.Format("Update SalaryVar Set NameVar = '{0}', [Value] = '{1}', Rate = '{2}', DateApply = '{3}', Descript = N'{4}', Status = {5} WHERE IdSalaryVar = {6}", _NameVar, _Value, _Rate, _DateApply, _Descript, _Status, _IdSalaryVar)
        Else                        'Insert data
            strSQL = String.Format("Insert Into SalaryVar(NameVar,[Value],Rate,DateApply,Descript,Status) Values(N'{0}', N'{1}', N'{2}', N'{3}', N'{4}', {5})", _NameVar, _Value, _Rate, _DateApply, _Descript, _Status)
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu tham số hệ thống: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

#End Region

#Region "---> Các hàm xử lý liên quan đến Danh mục - Lương <---"
    ''' <summary>
    ''' Hàm fill data vào treeview - Phần danh mục liên quan đến lương
    ''' </summary>
    ''' <param name="tv_name">Tên cây dữ liệu</param>
    ''' <remarks></remarks>
    Public Sub Fill_Data(ByVal tv_name As System.Windows.Forms.TreeView)
        tv_name.Nodes.Clear()
        Dim tn_parent As TreeNode
        Dim tn_child As TreeNode
        Dim strSQL As String = ""

        '1 - Node Danh mục - Nghị định lương
        tn_parent = tv_name.Nodes.Add("Nghị định lương")
        tn_parent.Tag = "00ND"
        strSQL = String.Format("Select * from NghiDinhLuong order by TenND asc")
        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db_child Is Nothing) Then
                If (db_child.Rows.Count > 0) Then
                    For j As Integer = 0 To (db_child.Rows.Count - 1)
                        tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("TenND").ToString())
                        tn_child.Tag = "01ND" + db_child.Rows(j)("IdNDLuong").ToString().Trim()
                    Next
                End If
            End If
        End Using

        '2 - Node Danh mục - Bảng lương
        tn_parent = tv_name.Nodes.Add("Bảng lương")
        tn_parent.Tag = "00BL"
        strSQL = String.Format("Select * from BangLuong")
        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db_child Is Nothing) Then
                If (db_child.Rows.Count > 0) Then
                    For j As Integer = 0 To (db_child.Rows.Count - 1)
                        tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("MoTa").ToString())
                        tn_child.Tag = "01BL" + db_child.Rows(j)("IdBangLuong").ToString().Trim()
                    Next
                End If
            End If
        End Using

        '3 - Node Danh mục - Ngạch lương
        tn_parent = tv_name.Nodes.Add("Ngạch lương")
        tn_parent.Tag = "00NL"
        strSQL = String.Format("Select * from NgachLuong")
        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db_child Is Nothing) Then
                If (db_child.Rows.Count > 0) Then
                    For j As Integer = 0 To (db_child.Rows.Count - 1)
                        tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("Mota").ToString())
                        tn_child.Tag = "01NL" + db_child.Rows(j)("IdNgachLuong").ToString().Trim()
                    Next
                End If
            End If
        End Using

        '4 - Node Danh mục - Bậc lương
        tn_parent = tv_name.Nodes.Add("Bậc lương")
        tn_parent.Tag = "00BA"
        strSQL = String.Format("Select * from BacLuong")
        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db_child Is Nothing) Then
                If (db_child.Rows.Count > 0) Then
                    For j As Integer = 0 To (db_child.Rows.Count - 1)
                        tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("Mota").ToString())
                        tn_child.Tag = "01BA" + db_child.Rows(j)("IdBacLuong").ToString().Trim()
                    Next
                End If
            End If
        End Using

        '5 - Node Danh mục - Lương cơ bảng
        tn_parent = tv_name.Nodes.Add("Lương cơ bản")
        tn_parent.Tag = "00CB"

        '6 - Node Danh mục - Mức phụ cấp
        tn_parent = tv_name.Nodes.Add("Mức phụ cấp")
        tn_parent.Tag = "00PC"


        '7 - Node Danh mục tham số hệ thống lương
        tn_parent = tv_name.Nodes.Add("Tham số lương hệ thống")
        tn_parent.Tag = "00TS"
    End Sub

    ''' <summary>
    ''' Hàm trả về danh sách các bản ghi - Danh mục liên quan đến lương
    ''' </summary>
    ''' <param name="_Status">
    '''  Chỉ số xác định danh mục lương. Với quy định sau:
    '''                              _State = 4: Nghị định lương
    '''                              _State = 5: Bảng lương
    '''                              _State = 6: Ngạch lương
    '''                              _State = 7: Bậc lương
    '''                              _State = 8: Lương cơ bản
    '''                              _State = 9: Mức phụ cấp
    '''                              _State = 10: Tham số lương hệ thống
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll(ByVal _Status As Byte) As DataTable
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim strSQL As String = ""
            Select Case _Status
                Case 4  'Nghị định lương
                    strSQL = "Select * From NghiDinhLuong  Order by TenND asc"
                Case 5  'Bảng lương
                    strSQL = "Select a.*, b.TenND as Nd_Luong From BangLuong a, NghiDinhLuong b Where b.IdNDLuong = a.IdND_Luong"
                Case 6  'Ngạch lương
                    strSQL = "Select a.*,b.BangLuong From NgachLuong a, BangLuong b Where a.IdBangLuong = b.IdBangLuong"
                Case 7  'Bậc lương
                    strSQL = "Select a.*,b.NgachLuong From BacLuong a, NgachLuong b Where a.IdNgachLuong = b.IdNgachLuong"
                Case 8  'Lương cơ bản
                    strSQL = "SELECT a.*,b.ten_goi as Chucvu From TienLuong a, DanhMuc b Where a.IdCV_Nguoi_QD = b.id and b.id_goc = 14 Order by NgayHuong desc"
                Case 9  'Mức phụ cấp
                    strSQL = "SELECT a.*,b.ten_goi as LoaiPhCap From MucPhuCap a, DanhMuc b Where a.IdLoai_PhC = b.id and b.id_goc = 31 Order by IdLoai_PhC, Muc_PhC"
                Case 10  'Tham số Lương hệ thống
                    strSQL = "SELECT * from SalaryVar order by status desc, nameVar"
            End Select

            Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL, connection)
            command.CommandType = CommandType.Text

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "SalaryRecords")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("SalaryRecords")
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
    ''' Hàm trả về bản ghi thoả mãn điều kiện truyền vào theo chỉ số id
    ''' </summary>
    ''' <param name="_Id">Chỉ số xác định bản ghi</param>
    ''' <param name="_Status">
    '''  Chỉ số xác định danh mục lương. Với quy định sau:
    '''                              _State = 4: Nghị định lương
    '''                              _State = 5: Bảng lương
    '''                              _State = 6: Ngạch lương
    '''                              _State = 7: Bậc lương
    '''                              _State = 8: Lương cơ bản
    '''                              _State = 9: Mức phụ cấp
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataForId(ByVal _Id As Integer, ByVal _Status As Byte) As DataRow
        Dim connection As System.Data.SqlClient.SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim strSQL As String = ""
            Select Case _Status
                Case 4  'Nghị định lương
                    strSQL = "Select * From NghiDinhLuong Where IdNDLuong = @_Id Order by TenND asc"
                Case 5  'Bảng lương
                    strSQL = "Select a.*, b.TenND as Nd_Luong From BangLuong a, NghiDinhLuong b Where b.IdNDLuong = a.IdND_Luong and a.IdBangLuong = @_Id"
                Case 6  'Ngạch lương
                    strSQL = "Select a.*,b.BangLuong From NgachLuong a, BangLuong b Where a.IdBangLuong = b.IdBangLuong and a.IdNgachLuong = @_Id"
                Case 7  'Bậc lương
                    strSQL = "Select a.*,b.NgachLuong From BacLuong a, NgachLuong b Where a.IdNgachLuong = b.IdNgachLuong and a.IdBacLuong = @_Id"
                Case 8  'Lương cơ bản
                    strSQL = "SELECT a.*,b.ten_goi as Chucvu From TienLuong a, DanhMuc b Where (a.IdCV_Nguoi_QD = b.id and b.id_goc = 14) and (a.IdTienLuong = @_Id) Order by NgayHuong desc"
                Case 9  'Mức phụ cấp
                    strSQL = "SELECT a.*,b.ten_goi as LoaiPhCap From MucPhuCap a, DanhMuc b Where (a.IdLoai_PhC = b.id and b.id_goc = 31) and (a.IdMuc_PhC = @_Id) Order by IdLoai_PhC, Muc_PhC"
                Case 10  'Tham số Lương hệ thống
                    strSQL = "SELECT * From SalaryVar Where IdSalaryVar = @_Id"
            End Select

            Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL, connection)
            command.CommandType = CommandType.Text

            command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Id", SqlDbType.Int))
            command.Parameters("@_Id").Value = _Id

            Using mydap As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "SalaryRecord")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("SalaryRecord").Rows(0)
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
    ''' Hàm thực hiện cập nhật dữ liệu - Danh mục Nghị định lương
    ''' </summary>
    ''' <param name="obj_list">Nghị định lương</param>
    ''' <remarks></remarks>
    Public Sub Save_NghiDinhLuong(ByVal obj_list As NghiDinhLuong)
        Dim strSQL As String = ""
        If (obj_list.IdNDLuong > 0) Then   'Update data
            strSQL = "Update NghiDinhLuong Set MaND = @_MaND, SoND = @_SoND, TenND = @_TenND, LoaiND = @_LoaiND, GhiChu = @_GhiChu, Status = @_Status Where IdNDLuong = @_IdNDLuong"
        Else                        'Insert data
            strSQL = "Insert Into NghiDinhLuong(MaND, SoND, TenND, LoaiND, GhiChu, Status) Values(@_MaND, @_SoND, @_TenND, @_LoaiND, @_GhiChu, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdNDLuong", obj_list.IdNDLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_MaND", obj_list.MaND))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_SoND", obj_list.SoND))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_TenND", obj_list.TenND))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_LoaiND", obj_list.LoaiND))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_GhiChu", obj_list.GhiChu))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu nghị định lương: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Danh mục Bảng lương
    ''' </summary>
    ''' <param name="obj_list">Bảng lương</param>
    ''' <remarks></remarks>
    Public Sub Save_BangLuong(ByVal obj_list As BangLuong)
        Dim strSQL As String = ""
        If (obj_list.IdBangLuong > 0) Then   'Update data
            strSQL = "Update BangLuong Set IdND_Luong = @_IdND_Luong, BangLuong = @_BangLuong, MoTa = @_MoTa, Status = @_Status Where IdBangLuong = @_IdBangLuong"
        Else                        'Insert data
            strSQL = "Insert Into BangLuong(IdND_Luong, BangLuong, MoTa, Status) Values(@_IdND_Luong, @_BangLuong, @_MoTa, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdBangLuong", obj_list.IdBangLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdND_Luong", obj_list.IdND_Luong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BangLuong", obj_list.BangLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_MoTa", obj_list.MoTa))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu bảng lương: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Danh mục Ngạch lương
    ''' </summary>
    ''' <param name="obj_list">Ngạch lương</param>
    ''' <remarks></remarks>
    Public Sub Save_NgachLuong(ByVal obj_list As NgachLuong)
        Dim strSQL As String = ""
        If (obj_list.IdNgachLuong > 0) Then   'Update data
            strSQL = "Update NgachLuong Set IdBangLuong = @_IdBangLuong, NgachLuong = @_NgachLuong, Loai = @_Loai, Mota = @_Mota, TimeNangNgach = @_TimeNangNgach, Status = @_Status Where IdNgachLuong = @_IdNgachLuong"
        Else                        'Insert data
            strSQL = "Insert Into NgachLuong(IdBangLuong, NgachLuong, Loai, Mota, TimeNangNgach, Status) Values(@_IdBangLuong, @_NgachLuong, @_Loai, @_Mota, @_TimeNangNgach, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdNgachLuong", obj_list.IdNgachLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdBangLuong", obj_list.IdBangLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_NgachLuong", obj_list.NgachLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Loai", obj_list.Loai))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_MoTa", obj_list.Mota))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_TimeNangNgach", obj_list.TimeNangNgach))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu ngạch lương: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Danh mục Bậc lương
    ''' </summary>
    ''' <param name="obj_list">Bậc lương</param>
    ''' <remarks></remarks>
    Public Sub Save_BacLuong(ByVal obj_list As BacLuong)
        Dim strSQL As String = ""
        If (obj_list.IdBacLuong > 0) Then   'Update data
            strSQL = "Update BacLuong Set IdNgachLuong = @_IdNgachLuong, BacLuong = @_BacLuong, Heso = @_Heso, Mota = @_Mota, Status = @_Status Where IdBacLuong = @_IdBacLuong"
        Else                        'Insert data
            strSQL = "Insert Into BacLuong(IdNgachLuong, BacLuong, Heso, Mota, Status) Values(@_IdNgachLuong, @_BacLuong, @_Heso, @_Mota, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdBacLuong", obj_list.IdBacLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdNgachLuong", obj_list.IdNgachLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_BacLuong", obj_list.BacLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Heso", obj_list.Heso))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_MoTa", obj_list.Mota))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu bạc lương: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Danh mục Lương cơ bản
    ''' </summary>
    ''' <param name="obj_list">Lương cơ bản</param>
    ''' <remarks></remarks>
    Public Sub Save_TienLuong(ByVal obj_list As TienLuong)
        Dim strSQL As String = ""
        If (obj_list.IdTienLuong > 0) Then   'Update data
            strSQL = "Update TienLuong Set LuongCoBan = @_LuongCoBan, HeSoNganh = @_HeSoNganh, NgayHuong = @_NgayHuong, SoQD = @_SoQD, NgayQD = @_NgayQD, NguoiQD = @_NguoiQD, IdCV_Nguoi_QD = @_IdCV_Nguoi_QD, Ghichu = @_Ghichu, Status = @_Status Where IdTienLuong = @_IdTienLuong"
        Else                        'Insert data
            strSQL = "Insert Into TienLuong(LuongCoBan, HeSoNganh, NgayHuong, SoQD, NgayQD, NguoiQD, IdCV_Nguoi_QD, Ghichu, Status) Values(@_LuongCoBan, @_HeSoNganh, @_NgayHuong, @_SoQD, @_NgayQD, @_NguoiQD, @_IdCV_Nguoi_QD, @_Ghichu, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdTienLuong", obj_list.IdTienLuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_LuongCoBan", obj_list.LuongCoBan))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_HeSoNganh", obj_list.HeSoNganh))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_NgayHuong", obj_list.NgayHuong))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_SoQD", obj_list.SoQD))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_NgayQD", obj_list.NgayQD))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_NguoiQD", obj_list.NguoiQD))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdCV_Nguoi_QD", obj_list.IdCV_Nguoi_QD))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Ghichu", obj_list.Ghichu))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu lương cơ bản: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Danh mục Mức phụ cấp
    ''' </summary>
    ''' <param name="obj_list">Mức phụ cấp</param>
    ''' <remarks></remarks>
    Public Sub Save_MucPhuCap(ByVal obj_list As MucPhuCap)
        Dim strSQL As String = ""
        If (obj_list.IdMuc_PhC > 0) Then   'Update data
            strSQL = "Update MucPhuCap Set IdLoai_PhC = @_IdLoai_PhC, Muc_PhC = @_Muc_PhC, Status = @_Status Where IdMuc_PhC = @_IdMuc_PhC"
        Else                        'Insert data
            strSQL = "Insert Into MucPhuCap(IdLoai_PhC, Muc_PhC, Status) Values(@_IdLoai_PhC, @_Muc_PhC, @_Status)"
        End If
        Dim command As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdMuc_PhC", obj_list.IdMuc_PhC))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_IdLoai_PhC", obj_list.IdLoai_PhC))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Muc_PhC", obj_list.Muc_PhC))
        command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@_Status", obj_list.Status))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật dữ liệu mức phụ cấp: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

#Region "---> Một số hàm khác về lấy và kiểm tra dữ liệu <---"
    ''' <summary>
    ''' Hàm thực hiện trả về id chi nhánh tỉnh khi biết id của phòng giao dịch
    ''' </summary>
    ''' <param name="_IdChild">Id của pgd</param>
    ''' <returns>Chỉ số xác định chi nhánh tỉnh</returns>
    ''' <remarks></remarks>
    Public Function GetRootId(ByVal _IdChild As Integer) As String
        Dim _IdRoot As Integer = 0
        Dim strSQL As String = String.Format("Select * from ChiNhanh Where id = {0} and Status = 1 and id_goc > 1", _IdChild)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _IdRoot = db.Rows(0)("id_goc").ToString().Trim()
                End If
            End If
        End Using
        Return _IdRoot
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về mã hiệu chi nhánh từ id chi nhánh
    ''' </summary>
    ''' <param name="_IdIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCodeForId(ByVal _IdIndex As Integer) As String
        Dim _result As String = ""
        Dim strSQL As String = String.Format("Select * from ChiNhanh Where id = {0} And Status = 1", _IdIndex)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _result = db.Rows(0)("ma_so").ToString().Trim()
                End If
            End If
        End Using
        Return _result
    End Function

    ''' <summary>
    ''' Hàm thực hiện Kiểm tra dữ liệu của 1 cột trong một bảng có trùng nhau không
    ''' Ví dụ: Kiêm tra xem có Code nào trùng nhau trong bảng Danh mục không?
    ''' </summary>
    ''' <param name="_TableName">Tên bảng</param>
    ''' <param name="_FieldName">Tên cột cần kiểm tra dữ liệu</param>
    ''' <param name="_ValExist">Trả về giá trị khi phát hiện đó là giá trị trùng nhau
    ''' <returns>True: Trùng nhau. False: Không Trùng nhau</returns>
    ''' <remarks></remarks>
    Public Function IsExistCode(ByVal _TableName As String, ByVal _FieldName As String, ByRef _ValExist As String) As Boolean
        Dim strSQL As String = "SELECT * FROM " + _TableName
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        If i < db.Rows.Count - 1 Then
                            If (db.Rows(i)(_FieldName).ToString().Trim() = db.Rows(i + 1)(_FieldName).ToString().Trim()) Then
                                _ValExist = db.Rows(i)(_FieldName).ToString().Trim()
                                Return True
                            End If
                        End If
                    Next
                End If
            End If
        End Using
        Return False
    End Function
#End Region


     
End Class