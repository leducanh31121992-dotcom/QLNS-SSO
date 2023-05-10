Imports System.Data
Imports System.Data.SqlClient

Public Class clsHS_CsLaodong

    Private _SqlHelper As DBAccess

    Public Sub New()
        _SqlHelper = New DBAccess
    End Sub

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"

    Public Class HS_BHXH

        Private _IdBHXH As String
        Private _IdCanBo As String
        Private _So As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _Muc_DongBHXH As Double
        Private _TenCQ_DongBHXH As String
        Private _GhiChu As String

        Public Sub New()
            _IdBHXH = ""
            _IdCanBo = ""
            _So = ""
            _TuNgay = DateTime.Now.ToShortDateString()
            _DenNgay = DateTime.Now.ToShortDateString()
            _Muc_DongBHXH = 0
            _TenCQ_DongBHXH = ""
            _GhiChu = ""
        End Sub

        Public Property IdBHXH() As String
            Get
                Return _IdBHXH
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ BHXH có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdBHXH = value
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

        Public Property So() As String
            Get
                Return _So
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số sổ BHXH có độ dài không hợp lệ!", value, value.ToString())
                End If
                _So = value
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

        Public Property Muc_DongBHXH() As Double
            Get
                Return _Muc_DongBHXH
            End Get
            Set(ByVal value As Double)
                _Muc_DongBHXH = value
            End Set
        End Property

        Public Property TenCQ_DongBHXH() As String
            Get
                Return _TenCQ_DongBHXH
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên cơ quan đóng BHXH có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TenCQ_DongBHXH = value
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

    Public Class HS_BHYT
        Private _IdBHYT As String
        Private _IdCanBo As String
        Private _SoThe As String
        Private _NgayCap As DateTime
        Private _NoiCap As String
        Private _NoiDangKy_KCB As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _GhiChu As String

        Public Sub New()
            _IdBHYT = ""
            _IdCanBo = ""
            _SoThe = ""
            _NgayCap = DateTime.Now.ToShortDateString()
            _NoiCap = ""
            _NoiDangKy_KCB = ""
            _TuNgay = DateTime.Now.ToShortDateString()
            _DenNgay = DateTime.Now.ToShortDateString()
            _GhiChu = ""
        End Sub

        Public Property IdBHYT() As String
            Get
                Return _IdBHYT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ BHYT có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdBHYT = value
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
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số thẻ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoThe = value
            End Set
        End Property

        Public Property NgayCap() As DateTime
            Get
                Return _NgayCap
            End Get
            Set(ByVal value As DateTime)
                _NgayCap = value
            End Set
        End Property

        Public Property NoiCap() As String
            Get
                Return _NoiCap
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiCap = value
            End Set
        End Property

        Public Property NoiDangKy_KCB() As String
            Get
                Return _NoiDangKy_KCB
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi đăng ký khám chữa bệnh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiDangKy_KCB = value
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

    Public Class HS_SucKhoe
        Private _IdKhamChua As String
        Private _IdCanBo As String
        Private _Nam As Int32
        Private _Dot As Byte
        Private _NoiDung As String
        Private _NoiKham As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _CanNang As Double
        Private _ChieuCao As Double
        Private _LoaiSucKhoe As Byte
        Private _KetLuan As String
        Private _GhiChu As String

        Public Sub New()
            _IdKhamChua = ""
            _IdCanBo = ""
            _Nam = 0
            _Dot = 0
            _NoiDung = ""
            _NoiKham = ""
            _TuNgay = DateTime.Now.ToShortDateString()
            _DenNgay = DateTime.Now.ToShortDateString()
            _CanNang = 0
            _ChieuCao = 0
            _LoaiSucKhoe = 0
            _KetLuan = ""
            _GhiChu = ""
        End Sub

        Public Property IdKhamChua() As String
            Get
                Return _IdKhamChua
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id khám chữa bệnh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdKhamChua = value
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

        Public Property Nam() As Int32
            Get
                Return _Nam
            End Get
            Set(ByVal value As Int32)
                _Nam = value
            End Set
        End Property

        Public Property Dot() As Byte
            Get
                Return _Dot
            End Get
            Set(ByVal value As Byte)
                _Dot = value
            End Set
        End Property

        Public Property NoiDung() As String
            Get
                Return _NoiDung
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nội dung có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiDung = value
            End Set
        End Property

        Public Property NoiKham() As String
            Get
                Return _NoiKham
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi khám có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiKham = value
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

        Public Property CanNang() As Double
            Get
                Return _CanNang
            End Get
            Set(ByVal value As Double)
                _CanNang = value
            End Set
        End Property

        Public Property ChieuCao() As Double
            Get
                Return _ChieuCao
            End Get
            Set(ByVal value As Double)
                _ChieuCao = value
            End Set
        End Property

        Public Property LoaiSucKhoe() As Byte
            Get
                Return _LoaiSucKhoe
            End Get
            Set(ByVal value As Byte)
                _LoaiSucKhoe = value
            End Set
        End Property

        Public Property KetLuan() As String
            Get
                Return _KetLuan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết luận có độ dài không hợp lệ!", value, value.ToString())
                End If
                _KetLuan = value
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

    Public Class HS_NghiPhep
        Private _IdNghiPhep As String
        Private _IdCanBo As String
        Private _Nam As Integer
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _SoNgay As Int16
        Private _IdLoaiNghi As Integer
        Private _NoiNghi As String
        Private _LyDo As String
        Private _TyleHuong As Double
        Private _TienTroCap As Double
        Private _NguoiKy_Nghi As String
        Private _GhiChu As String

        Private _ChiNhanh_Cd As String
        Private _ChiNhanh_HT As String
        Private _DonVi_Cd As String
        Private _DonVi_HT As String
        Private _PhongBan_Cd As String
        Private _PhongBan_HT As String
        Private _ChucVu_Cd As String
        Private _ChucVu_HT As String
        Private _MaCB As String
        Private _HoTen As String
        Private _NgaySinh As String
        Private _GioiTinh As String
        Private _SoCMT As String
        Private _IdOutPut As String
        Private _PosCode As String
        Private _LoaiNghi_Cd As String
        Private _CreatedBy As String
        Private _CreatedDate As DateTime
        Private _ModifiedBy As String
        Private _ModifiedDate As DateTime

        Public Sub New()
            _IdNghiPhep = ""
            _IdCanBo = ""
            _Nam = 0
            _TuNgay = Globals.GetDateTime_ForServerDB()
            _DenNgay = Globals.GetDateTime_ForServerDB()
            _SoNgay = 0
            _IdLoaiNghi = 0
            _NoiNghi = ""
            _LyDo = ""
            _TyleHuong = 0
            _TienTroCap = 0
            _NguoiKy_Nghi = ""
            _GhiChu = ""
            _MaCB = ""
            _HoTen = ""
            _NgaySinh = ""
            _GioiTinh = ""
            _SoCMT = ""
            _IdOutPut = ""
            _PosCode = ""
            _LoaiNghi_Cd = ""
            _CreatedBy = ""
            _CreatedDate = Globals.GetDateTime_ForServerDB()
            _ModifiedBy = ""
            _ModifiedDate = Globals.GetDateTime_ForServerDB()

            _ChiNhanh_Cd = ""
            _ChiNhanh_HT = ""
            _DonVi_Cd = ""
            _DonVi_HT = ""
            _PhongBan_Cd = ""
            _PhongBan_HT = ""
            _ChucVu_Cd = ""
            _ChucVu_HT = ""
        End Sub

        Public Property IdNghiPhep() As String
            Get
                Return _IdNghiPhep
            End Get
            Set(ByVal value As String)
                _IdNghiPhep = value
            End Set
        End Property
        Public Property IdCanBo() As String
            Get
                Return _IdCanBo
            End Get
            Set(ByVal value As String)
                _IdCanBo = value
            End Set
        End Property
        Public Property Nam() As Integer
            Get
                Return _Nam
            End Get
            Set(ByVal value As Integer)
                _Nam = value
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
        Public Property SoNgay() As Int16
            Get
                Return _SoNgay
            End Get
            Set(ByVal value As Int16)
                _SoNgay = value
            End Set
        End Property
        Public Property IdLoaiNghi() As Int32
            Get
                Return _IdLoaiNghi
            End Get
            Set(ByVal value As Int32)
                _IdLoaiNghi = value
            End Set
        End Property
        Public Property NoiNghi() As String
            Get
                Return _NoiNghi
            End Get
            Set(ByVal value As String)
                _NoiNghi = value
            End Set
        End Property
        Public Property LyDo() As String
            Get
                Return _LyDo
            End Get
            Set(ByVal value As String)
                _LyDo = value
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
        Public Property TienTroCap() As Double
            Get
                Return _TienTroCap
            End Get
            Set(ByVal value As Double)
                _TienTroCap = value
            End Set
        End Property
        Public Property NguoiKy_Nghi() As String
            Get
                Return _NguoiKy_Nghi
            End Get
            Set(ByVal value As String)
                _NguoiKy_Nghi = value
            End Set
        End Property
        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                _GhiChu = value
            End Set
        End Property
        Public Property MaCB() As String
            Get
                Return _MaCB
            End Get
            Set(ByVal value As String)
                _MaCB = value
            End Set
        End Property
        Public Property HoTen() As String
            Get
                Return _HoTen
            End Get
            Set(ByVal value As String)
                _HoTen = value
            End Set
        End Property
        Public Property NgaySinh() As String
            Get
                Return _NgaySinh
            End Get
            Set(ByVal value As String)
                _NgaySinh = value
            End Set
        End Property
        Public Property GioiTinh() As String
            Get
                Return _GioiTinh
            End Get
            Set(ByVal value As String)
                _GioiTinh = value
            End Set
        End Property
        Public Property SoCMT() As String
            Get
                Return _SoCMT
            End Get
            Set(ByVal value As String)
                _SoCMT = value
            End Set
        End Property
        Public Property IdOutPut() As String
            Get
                Return _IdOutPut
            End Get
            Set(ByVal value As String)
                _IdOutPut = value
            End Set
        End Property
        Public Property PosCode() As String
            Get
                Return _PosCode
            End Get
            Set(ByVal value As String)
                _PosCode = value
            End Set
        End Property
        Public Property LoaiNghi_Cd() As String
            Get
                Return _LoaiNghi_Cd
            End Get
            Set(ByVal value As String)
                _LoaiNghi_Cd = value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return _CreatedBy
            End Get
            Set(ByVal value As String)
                _CreatedBy = value
            End Set
        End Property
        Public Property CreatedDate() As DateTime
            Get
                Return _CreatedDate
            End Get
            Set(ByVal value As DateTime)
                _CreatedDate = value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return _ModifiedBy
            End Get
            Set(ByVal value As String)
                _ModifiedBy = value
            End Set
        End Property
        Public Property ModifiedDate() As DateTime
            Get
                Return _ModifiedDate
            End Get
            Set(ByVal value As DateTime)
                _ModifiedDate = value
            End Set
        End Property


        Public Property ChiNhanh_Cd() As String
            Get
                Return _ChiNhanh_Cd
            End Get
            Set(ByVal value As String)
                _ChiNhanh_Cd = value
            End Set
        End Property
        Public Property ChiNhanh_HT() As String
            Get
                Return _ChiNhanh_HT
            End Get
            Set(ByVal value As String)
                _ChiNhanh_HT = value
            End Set
        End Property
        Public Property DonVi_Cd() As String
            Get
                Return _DonVi_Cd
            End Get
            Set(ByVal value As String)
                _DonVi_Cd = value
            End Set
        End Property
        Public Property DonVi_HT() As String
            Get
                Return _DonVi_HT
            End Get
            Set(ByVal value As String)
                _DonVi_HT = value
            End Set
        End Property
        Public Property PhongBan_Cd() As String
            Get
                Return _PhongBan_Cd
            End Get
            Set(ByVal value As String)
                _PhongBan_Cd = value
            End Set
        End Property
        Public Property PhongBan_HT() As String
            Get
                Return _PhongBan_HT
            End Get
            Set(ByVal value As String)
                _PhongBan_HT = value
            End Set
        End Property
        Public Property ChucVu_Cd() As String
            Get
                Return _ChucVu_Cd
            End Get
            Set(ByVal value As String)
                _ChucVu_Cd = value
            End Set
        End Property
        Public Property ChucVu_HT() As String
            Get
                Return _ChucVu_HT
            End Get
            Set(ByVal value As String)
                _ChucVu_HT = value
            End Set
        End Property
    End Class

    Public Class CB_TaiNan
        Private _IdCBTaiNan As String
        Private _IdVuTN As String
        Private _IdCanBo As String
        Private _MucDo_TN As Byte
        Private _SoNgayNghi As Int32
        Private _ChiPhi As Double
        Private _TT_ThuongTat As String
        Private _Noi_PP_DieuTri As String
        Private _GhiChu As String
        Public Sub New()
            _IdCBTaiNan = ""
            _IdVuTN = ""
            _IdCanBo = ""
            _MucDo_TN = 0
            _SoNgayNghi = 0
            _ChiPhi = 0
            _TT_ThuongTat = ""
            _Noi_PP_DieuTri = ""
            _GhiChu = ""
        End Sub
        Public Property IdCBTaiNan() As String
            Get
                Return _IdCBTaiNan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id cán bộ tai nạn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCBTaiNan = value
            End Set
        End Property

        Public Property IdVuTN() As String
            Get
                Return _IdVuTN
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id vụ tai nạn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdVuTN = value
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

        Public Property MucDo_TN() As Byte
            Get
                Return _MucDo_TN
            End Get
            Set(ByVal value As Byte)
                _MucDo_TN = value
            End Set
        End Property

        Public Property SoNgayNghi() As Int32
            Get
                Return _SoNgayNghi
            End Get
            Set(ByVal value As Int32)
                _SoNgayNghi = value
            End Set
        End Property

        Public Property ChiPhi() As Double
            Get
                Return _ChiPhi
            End Get
            Set(ByVal value As Double)
                _ChiPhi = value
            End Set
        End Property

        Public Property TT_ThuongTat() As String
            Get
                Return _TT_ThuongTat
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị thông tin thương tật có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TT_ThuongTat = value
            End Set
        End Property

        Public Property Noi_PP_DieuTri() As String
            Get
                Return _Noi_PP_DieuTri
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi phối hợp điều trị có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Noi_PP_DieuTri = value
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

    Public Class HS_VuTaiNan
        Private _IdVuTaiNan As String
        Private _IdLoaiTN As Int32
        Private _TenVu_TN As String
        Private _Ngay_TN As DateTime
        Private _NgoaiDonVi As Byte
        Private _Noi_TN As String
        Private _IdNguyenNhanTN As Int32
        Private _SoTien_ThietHai As Double
        Private _Ngay_BaoCao As DateTime
        Private _Ngay_DieuTra As DateTime
        Private _Nguoi_DieuTra As String
        Private _MoTa_TN As String
        Private _DeNghi As String
        Private _GhiChu As String

        Public Sub New()
            _IdVuTaiNan = ""
            _IdLoaiTN = 0
            _TenVu_TN = ""
            _Ngay_TN = DateTime.Now.ToShortDateString()
            _NgoaiDonVi = 0
            _Noi_TN = ""
            _IdNguyenNhanTN = 0
            _SoTien_ThietHai = 0
            _Ngay_BaoCao = DateTime.Now.ToShortDateString()
            _Ngay_DieuTra = DateTime.Now.ToShortDateString()
            _Nguoi_DieuTra = ""
            _MoTa_TN = ""
            _DeNghi = ""
            _GhiChu = ""
        End Sub

        Public Property IdVuTaiNan() As String
            Get
                Return _IdVuTaiNan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id vụ tai nạn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdVuTaiNan = value
            End Set
        End Property

        Public Property IdLoaiTN() As Int32
            Get
                Return _IdLoaiTN
            End Get
            Set(ByVal value As Int32)
                _IdLoaiTN = value
            End Set
        End Property

        Public Property TenVu_TN() As String
            Get
                Return _TenVu_TN
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 30) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên vụ tai nạn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TenVu_TN = value
            End Set
        End Property

        Public Property Ngay_TN() As DateTime
            Get
                Return _Ngay_TN
            End Get
            Set(ByVal value As DateTime)
                _Ngay_TN = value
            End Set
        End Property

        Public Property NgoaiDonVi() As Byte
            Get
                Return _NgoaiDonVi
            End Get
            Set(ByVal value As Byte)
                _NgoaiDonVi = value
            End Set
        End Property

        Public Property Noi_TN() As String
            Get
                Return _Noi_TN
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi tai nạn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Noi_TN = value
            End Set
        End Property

        Public Property IdNguyenNhanTN() As Int32
            Get
                Return _IdNguyenNhanTN
            End Get
            Set(ByVal value As Int32)
                _IdNguyenNhanTN = value
            End Set
        End Property

        Public Property SoTien_ThietHai() As Double
            Get
                Return _SoTien_ThietHai
            End Get
            Set(ByVal value As Double)
                _SoTien_ThietHai = value
            End Set
        End Property

        Public Property Ngay_BaoCao() As DateTime
            Get
                Return _Ngay_BaoCao
            End Get
            Set(ByVal value As DateTime)
                _Ngay_BaoCao = value
            End Set
        End Property

        Public Property Ngay_DieuTra() As DateTime
            Get
                Return _Ngay_DieuTra
            End Get
            Set(ByVal value As DateTime)
                _Ngay_DieuTra = value
            End Set
        End Property

        Public Property Nguoi_DieuTra() As String
            Get
                Return _Nguoi_DieuTra
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 30) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người điều tra có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Nguoi_DieuTra = value
            End Set
        End Property

        Public Property MoTa_TN() As String
            Get
                Return _MoTa_TN
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mô tả vụ tai nạn có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MoTa_TN = value
            End Set
        End Property

        Public Property DeNghi() As String
            Get
                Return _DeNghi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị đề nghị có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DeNghi = value
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

#Region "---> Hồ sơ bảo hiểm xã hội - Hàm cập nhật <---"
    ''' <summary>
    ''' Hàm cập nhật thêm mới - Hồ sơ bảo hiểm xã hội
    ''' </summary>
    ''' <param name="obj_bhxh">Đối tượng Bảo hiểm XH</param>
    ''' <returns>Chỉ số xác định tính duy nhất của bản ghi (Tự sinh)</returns>
    ''' <remarks></remarks>
    Public Function Insert_HS_BHXH(ByVal obj_bhxh As HS_BHXH) As String
        Dim command As SqlCommand = New SqlCommand("HS_BHXH_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdBHXH", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_bhxh.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_So", obj_bhxh.So))
        If (CType(obj_bhxh.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_TuNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_TuNgay", obj_bhxh.TuNgay))
        End If
        If (CType(obj_bhxh.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_bhxh.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_Muc_DongBHXH", obj_bhxh.Muc_DongBHXH))
        command.Parameters.Add(New SqlParameter("@_TenCQ_DongBHXH", obj_bhxh.TenCQ_DongBHXH))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_bhxh.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_bhxh.IdBHXH = command.Parameters("@_IdBHXH").Value.ToString
            Return obj_bhxh.IdBHXH
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ bảo hiểm xã hội: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Function

    ''' <summary>
    ''' Hàm cập nhật sửa đổi - Hồ sơ bảo hiểm xã hội
    ''' </summary>
    ''' <param name="obj_bhxh">Đối tượng bảo hiểm xã hội</param>
    ''' <remarks></remarks>
    Public Sub Update_HS_BHXH(ByVal obj_bhxh As HS_BHXH)
        Dim command As SqlCommand = New SqlCommand("HS_BHXH_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdBHXH", obj_bhxh.IdBHXH))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_bhxh.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_So", obj_bhxh.So))
        If (CType(obj_bhxh.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_TuNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_TuNgay", obj_bhxh.TuNgay))
        End If
        If (CType(obj_bhxh.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_bhxh.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_Muc_DongBHXH", obj_bhxh.Muc_DongBHXH))
        command.Parameters.Add(New SqlParameter("@_TenCQ_DongBHXH", obj_bhxh.TenCQ_DongBHXH))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_bhxh.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ bảo hiểm xã hội: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật Xoá bỏ - Hồ sơ bảo hiểm xã hội
    ''' </summary>
    ''' <param name="_IdBHXH">Id hồ sơ bảo hiểm xã hội</param>
    ''' <remarks></remarks>
    Public Sub Delete_HS_BHXH(ByVal _IdBHXH As String)
        Dim command As SqlCommand = New SqlCommand("HS_BHXH_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdBHXH", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdBHXH").Value = _IdBHXH
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ bảo hiểm xã hội: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

#Region "---> Hồ sơ bảo hiểm Y tế - Hàm cập nhật <---"
    ''' <summary>
    ''' Hàm thực hiện cập nhật - Thêm mới hồ sơ bảo hiểm y tế
    ''' </summary>
    ''' <param name="obj_bhyt">Đối tượng Bảo hiểm y tế</param>
    ''' <returns>Chỉ số xác định tính duy nhất bản ghi vừa sinh - Bảo hiểm y tế</returns>
    ''' <remarks></remarks>
    Public Function Insert_HS_BHYT(ByVal obj_bhyt As HS_BHYT) As String
        Dim command As SqlCommand = New SqlCommand("HS_BHYT_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdBHYT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_bhyt.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_bhyt.SoThe))

        If (CType(obj_bhyt.NgayCap, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayCap", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayCap", obj_bhyt.NgayCap))
        End If

        command.Parameters.Add(New SqlParameter("@_NoiCap", obj_bhyt.NoiCap))
        command.Parameters.Add(New SqlParameter("@_NoiDangKy_KCB", obj_bhyt.NoiDangKy_KCB))

        If (CType(obj_bhyt.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_TuNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_TuNgay", obj_bhyt.TuNgay))
        End If

        If (CType(obj_bhyt.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_bhyt.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_bhyt.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_bhyt.IdBHYT = command.Parameters("@_IdBHYT").Value.ToString
            Return obj_bhyt.IdBHYT
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ bảo hiểm y tế: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật - Thêm mới hồ sơ bảo hiểm y tế
    ''' </summary>
    ''' <param name="obj_bhyt"></param>
    ''' <remarks></remarks>
    Public Sub Update_HS_BHYT(ByVal obj_bhyt As HS_BHYT)
        Dim command As SqlCommand = New SqlCommand("HS_BHYT_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdBHYT", obj_bhyt.IdBHYT))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_bhyt.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_SoThe", obj_bhyt.SoThe))

        If (CType(obj_bhyt.NgayCap, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayCap", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayCap", obj_bhyt.NgayCap))
        End If

        command.Parameters.Add(New SqlParameter("@_NoiCap", obj_bhyt.NoiCap))
        command.Parameters.Add(New SqlParameter("@_NoiDangKy_KCB", obj_bhyt.NoiDangKy_KCB))

        If (CType(obj_bhyt.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_TuNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_TuNgay", obj_bhyt.TuNgay))
        End If

        If (CType(obj_bhyt.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_bhyt.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_bhyt.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ bảo hiểm y tế: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thiực hiện cập nhật Xoá bỏ - Hồ sơ bảo hiểm y tế
    ''' </summary>
    ''' <param name="_IdBHYT"></param>
    ''' <remarks></remarks>
    Public Sub Delete_HS_BHYT(ByVal _IdBHYT As String)
        Dim command As SqlCommand = New SqlCommand("HS_BHYT_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdBHYT", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdBHYT").Value = _IdBHYT
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ bảo hiểm y tế: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

#Region "---> Hồ sơ Sức khoẻ - Hàm cập nhật <---"
    ''' <summary>
    ''' Hàm thực hiện cập nhật - Thêm mới hồ sơ khám sức khoẻ của Cán bộ
    ''' </summary>
    ''' <param name="obj_hs_suckhoe">Đối tượng hồ sơ khám sức khoẻ</param>
    ''' <returns>Chỉ số bản ghi vừa sinh (Chỉ số xác định tính duy nhất)</returns>
    ''' <remarks></remarks>
    Public Function Insert_HS_SucKhoe(ByVal obj_hs_suckhoe As HS_SucKhoe) As String
        Dim command As SqlCommand = New SqlCommand("HS_SucKhoe_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdKhamChua", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_Nam", obj_hs_suckhoe.Nam))
        command.Parameters.Add(New SqlParameter("@_Dot", obj_hs_suckhoe.Dot))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_suckhoe.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_NoiDung", obj_hs_suckhoe.NoiDung))
        command.Parameters.Add(New SqlParameter("@_NoiKham", obj_hs_suckhoe.NoiKham))
        If (CType(obj_hs_suckhoe.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_TuNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_TuNgay", obj_hs_suckhoe.TuNgay))
        End If

        If (CType(obj_hs_suckhoe.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_hs_suckhoe.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_CanNang", obj_hs_suckhoe.CanNang))
        command.Parameters.Add(New SqlParameter("@_ChieuCao", obj_hs_suckhoe.ChieuCao))
        command.Parameters.Add(New SqlParameter("@_LoaiSucKhoe", obj_hs_suckhoe.LoaiSucKhoe))
        command.Parameters.Add(New SqlParameter("@_KetLuan", obj_hs_suckhoe.KetLuan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hs_suckhoe.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_hs_suckhoe.IdKhamChua = command.Parameters("@_IdKhamChua").Value.ToString
            Return obj_hs_suckhoe.IdKhamChua
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ khám chữa bệnh: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật - Sửa đổi hồ sơ khám sức khoẻ của Cán bộ
    ''' </summary>
    ''' <param name="obj_hs_suckhoe">Hồ sơ sức khoẻ của cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Update_HS_SucKhoe(ByVal obj_hs_suckhoe As HS_SucKhoe)
        Dim command As SqlCommand = New SqlCommand("HS_SucKhoe_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdKhamChua", obj_hs_suckhoe.IdKhamChua))
        command.Parameters.Add(New SqlParameter("@_Nam", obj_hs_suckhoe.Nam))
        command.Parameters.Add(New SqlParameter("@_Dot", obj_hs_suckhoe.Dot))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hs_suckhoe.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_NoiDung", obj_hs_suckhoe.NoiDung))
        command.Parameters.Add(New SqlParameter("@_NoiKham", obj_hs_suckhoe.NoiKham))
        If (CType(obj_hs_suckhoe.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_TuNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_TuNgay", obj_hs_suckhoe.TuNgay))
        End If

        If (CType(obj_hs_suckhoe.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_hs_suckhoe.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_CanNang", obj_hs_suckhoe.CanNang))
        command.Parameters.Add(New SqlParameter("@_ChieuCao", obj_hs_suckhoe.ChieuCao))
        command.Parameters.Add(New SqlParameter("@_LoaiSucKhoe", obj_hs_suckhoe.LoaiSucKhoe))
        command.Parameters.Add(New SqlParameter("@_KetLuan", obj_hs_suckhoe.KetLuan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hs_suckhoe.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ khám chữa bệnh: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật - Xoá bỏ hồ sơ sức khoẻ của Cán bộ
    ''' </summary>
    ''' <param name="_IdKhamChua">Id hồ sơ sức khoẻ của cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Delete_HS_SucKhoe(ByVal _IdKhamChua As String)
        Dim command As SqlCommand = New SqlCommand("HS_SucKhoe_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdKhamChua", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdKhamChua").Value = _IdKhamChua
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ sức khoẻ của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub
#End Region

#Region "---> Hồ sơ nghỉ phép - Hàm cập nhật <---"
    ''' <summary>
    ''' Hàm Thêm mới/Sửa đổi thông tin Nghỉ phép/Nghỉ khác của cán bộ
    ''' </summary>
    ''' <param name="_HS_NghiPhep">Đối tượng Hồ sơ nghỉ phép của cán bộ</param>
    ''' <returns>Chuổi Id xác định bản ghi</returns>
    ''' <remarks></remarks>
    Public Function Insert_Update_HS_NgPh(ByVal _HS_NghiPhep As HS_NghiPhep) As String
        Dim _Ret As String = ""
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "HS_NgPh_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.VarChar, 16).Direction = ParameterDirection.Output
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdNghiPhep", _HS_NghiPhep.IdNghiPhep, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PosCode", _HS_NghiPhep.PosCode, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo", _HS_NghiPhep.IdCanBo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Nam", _HS_NghiPhep.Nam, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TuNgay", _HS_NghiPhep.TuNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DenNgay", _HS_NghiPhep.DenNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoNgay", _HS_NghiPhep.SoNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdLoaiNghi", _HS_NghiPhep.IdLoaiNghi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NoiNghi", _HS_NghiPhep.NoiNghi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_LyDo", _HS_NghiPhep.LyDo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TyleHuong", _HS_NghiPhep.TyleHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TienTroCap", _HS_NghiPhep.TienTroCap, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NguoiKy_Nghi", _HS_NghiPhep.NguoiKy_Nghi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GhiChu", _HS_NghiPhep.GhiChu, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CreatedBy", _HS_NghiPhep.CreatedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _HS_NghiPhep.ModifiedBy, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _HS_NghiPhep.IdOutPut = _command.Parameters("@_IdOutPut").Value.ToString()
                        _Ret = _HS_NghiPhep.IdOutPut
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi Cập nhật thêm mới hồ sơ nghỉ phép: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Globals.Logger.Error("Lỗi Cập nhật Nghỉ phép cán bộ Insert_Update_HS_NgPh(): " + ex.Message)
                    _Ret = ""
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật - Xoá bỏ hồ sơ nghỉ phép của cán bộ
    ''' </summary>
    ''' <param name="_IdNghiPhep">Id hồ sơ nghỉ phép</param>
    ''' <returns>True - Thành Công; False - Thất bại</returns>
    ''' <remarks></remarks>
    Public Function Delete_HS_NghiPhep(ByVal _IdNghiPhep As String) As Boolean
        Dim _Ret As Boolean = False
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "HS_NgPh_Delete"
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdNghiPhep", _IdNghiPhep, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdBranch", IdDONVI, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi Xoá bỏ hồ sơ nghỉ phép của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Globals.Logger.Error("Lỗi Xoá bỏ hồ sơ nghỉ phép của cán bộ Delete_HS_NghiPhep() lỗi: " + ex.Message)
                    _Ret = False
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

#End Region

#Region "---> Hồ sơ Cán bộ tai nạn - Hàm cập nhật <---"

    ''' <summary>
    ''' Hàm thực hiện cập nhật  - Thêm mới Hồ sơ cán bộ Tai nạn
    ''' </summary>
    ''' <param name="obj_cbtainan">Đối tượng hồ sơ cán bộ tai nạn</param>
    ''' <returns>Chỉ số xác định tính duy nhất của bản ghi hồ sơ. Tự sinh</returns>
    ''' <remarks></remarks>
    Public Function Insert_CB_TaiNan(ByVal obj_cbtainan As CB_TaiNan) As String
        Dim command As SqlCommand = New SqlCommand("CB_TaiNan_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdCBTaiNan", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdVuTN", obj_cbtainan.IdVuTN))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_cbtainan.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_MucDo_TN", obj_cbtainan.MucDo_TN))
        command.Parameters.Add(New SqlParameter("@_SoNgayNghi", obj_cbtainan.SoNgayNghi))
        command.Parameters.Add(New SqlParameter("@_ChiPhi", obj_cbtainan.ChiPhi))
        command.Parameters.Add(New SqlParameter("@_TT_ThuongTat", obj_cbtainan.TT_ThuongTat))
        command.Parameters.Add(New SqlParameter("@_Noi_PP_DieuTri", obj_cbtainan.Noi_PP_DieuTri))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_cbtainan.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_cbtainan.IdCBTaiNan = command.Parameters("@_IdCBTaiNan").Value.ToString
            Return obj_cbtainan.IdCBTaiNan
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ cán bộ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật - Sửa đổi hồ sơ cán bộ Tai nạn
    ''' </summary>
    ''' <param name="obj_cbtainan">Hồ sơ cán bộ tai nạn</param>
    ''' <remarks></remarks>
    Public Sub Update_CB_TaiNan(ByVal obj_cbtainan As CB_TaiNan)
        Dim command As SqlCommand = New SqlCommand("CB_TaiNan_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCBTaiNan", obj_cbtainan.IdCBTaiNan))
        command.Parameters.Add(New SqlParameter("@_IdVuTN", obj_cbtainan.IdVuTN))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_cbtainan.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_MucDo_TN", obj_cbtainan.MucDo_TN))
        command.Parameters.Add(New SqlParameter("@_SoNgayNghi", obj_cbtainan.SoNgayNghi))
        command.Parameters.Add(New SqlParameter("@_ChiPhi", obj_cbtainan.ChiPhi))
        command.Parameters.Add(New SqlParameter("@_TT_ThuongTat", obj_cbtainan.TT_ThuongTat))
        command.Parameters.Add(New SqlParameter("@_Noi_PP_DieuTri", obj_cbtainan.Noi_PP_DieuTri))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_cbtainan.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ cán bộ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật - Xoá bỏ hồ sơ cán bộ Tai nạn
    ''' </summary>
    ''' <param name="_IdCBTaiNan"></param>
    ''' <remarks></remarks>
    Public Sub Delete_CB_TaiNan(ByVal _IdCBTaiNan As String)
        Dim command As SqlCommand = New SqlCommand("CB_TaiNan_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCBTaiNan", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdCBTaiNan").Value = _IdCBTaiNan
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ tai nạn của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Các hàm liên quan đến - Chi tiết Hồ sơ vụ tai nạn
    ''' <summary>
    ''' Thàm thực hiện cập nhật - Thêm mới hồ sơ chi tiết vụ tai nạn
    ''' </summary>
    ''' <param name="obj_hs_vutainan"></param>
    ''' <returns>Chỉ số xác định bản ghi hồ sơ vụ tai nạn</returns>
    ''' <remarks></remarks>
    Public Function Insert_HS_VuTaiNan(ByVal obj_hs_vutainan As HS_VuTaiNan) As String
        Dim command As SqlCommand = New SqlCommand("HS_VuTaiNan_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdVuTaiNan", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdLoaiTN", obj_hs_vutainan.IdLoaiTN))
        command.Parameters.Add(New SqlParameter("@_TenVu_TN", obj_hs_vutainan.TenVu_TN))

        If (CType(obj_hs_vutainan.Ngay_TN, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_TN", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_TN", obj_hs_vutainan.Ngay_TN))
        End If

        command.Parameters.Add(New SqlParameter("@_NgoaiDonVi", obj_hs_vutainan.NgoaiDonVi))
        command.Parameters.Add(New SqlParameter("@_Noi_TN", obj_hs_vutainan.Noi_TN))
        command.Parameters.Add(New SqlParameter("@_IdNguyenNhanTN", obj_hs_vutainan.IdNguyenNhanTN))
        command.Parameters.Add(New SqlParameter("@_SoTien_ThietHai", obj_hs_vutainan.SoTien_ThietHai))

        If (CType(obj_hs_vutainan.Ngay_BaoCao, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_BaoCao", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_BaoCao", obj_hs_vutainan.Ngay_BaoCao))
        End If

        If (CType(obj_hs_vutainan.Ngay_DieuTra, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_DieuTra", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_DieuTra", obj_hs_vutainan.Ngay_DieuTra))
        End If

        command.Parameters.Add(New SqlParameter("@_Nguoi_DieuTra", obj_hs_vutainan.Nguoi_DieuTra))
        command.Parameters.Add(New SqlParameter("@_MoTa_TN", obj_hs_vutainan.MoTa_TN))

        command.Parameters.Add(New SqlParameter("@_DeNghi", obj_hs_vutainan.DeNghi))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hs_vutainan.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_hs_vutainan.IdVuTaiNan = command.Parameters("@_IdVuTaiNan").Value.ToString
            Return obj_hs_vutainan.IdVuTaiNan
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ chi tiết vụ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Thàm thực hiện cập nhật - Sửa đổi hồ sơ chi tiết vụ tai nạn
    ''' </summary>
    ''' <param name="obj_hs_vutainan">Hồ sơ vụ tai nạn</param>
    ''' <remarks></remarks>
    Public Sub Update_HS_VuTaiNan(ByVal obj_hs_vutainan As HS_VuTaiNan)
        Dim command As SqlCommand = New SqlCommand("HS_VuTaiNan_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdVuTaiNan", obj_hs_vutainan.IdVuTaiNan))
        command.Parameters.Add(New SqlParameter("@_IdLoaiTN", obj_hs_vutainan.IdLoaiTN))
        command.Parameters.Add(New SqlParameter("@_TenVu_TN", obj_hs_vutainan.TenVu_TN))

        If (CType(obj_hs_vutainan.Ngay_TN, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_TN", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_TN", obj_hs_vutainan.Ngay_TN))
        End If

        command.Parameters.Add(New SqlParameter("@_NgoaiDonVi", obj_hs_vutainan.NgoaiDonVi))
        command.Parameters.Add(New SqlParameter("@_Noi_TN", obj_hs_vutainan.Noi_TN))
        command.Parameters.Add(New SqlParameter("@_IdNguyenNhanTN", obj_hs_vutainan.IdNguyenNhanTN))
        command.Parameters.Add(New SqlParameter("@_SoTien_ThietHai", obj_hs_vutainan.SoTien_ThietHai))

        If (CType(obj_hs_vutainan.Ngay_BaoCao, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_BaoCao", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_BaoCao", obj_hs_vutainan.Ngay_BaoCao))
        End If

        If (CType(obj_hs_vutainan.Ngay_DieuTra, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_DieuTra", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_DieuTra", obj_hs_vutainan.Ngay_DieuTra))
        End If

        command.Parameters.Add(New SqlParameter("@_Nguoi_DieuTra", obj_hs_vutainan.Nguoi_DieuTra))
        command.Parameters.Add(New SqlParameter("@_MoTa_TN", obj_hs_vutainan.MoTa_TN))

        command.Parameters.Add(New SqlParameter("@_DeNghi", obj_hs_vutainan.DeNghi))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hs_vutainan.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ chi tiết vụ tai nạn: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Thàm thực hiện cập nhật - Xoá bỏ hồ sơ chi tiết vụ tai nạn
    ''' </summary>
    ''' <param name="_IdVuTaiNan">Id Hồ sơ vụ tai nạn</param>
    ''' <remarks></remarks>
    Public Sub Delete_HS_VuTaiNan(ByVal _IdVuTaiNan As String)
        Dim command As SqlCommand = New SqlCommand("HS_VuTaiNan_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdVuTaiNan", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdVuTaiNan").Value = _IdVuTaiNan
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ vụ tai nạn của cán bộ: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

#End Region

#Region "---> Các hàm dùng chung và lấy dữ liệu <---"
    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào Treeview - Hồ sơ chi tiết các vụ tai nạn
    ''' </summary>
    ''' <param name="tv_name"></param>
    ''' <remarks></remarks>
    Public Sub BindData_TreeView(ByVal tv_name As TreeView)
        tv_name.Nodes.Clear()
        Dim tn_parent As TreeNode
        Dim tn_child As TreeNode
        Dim strSQL As String = ""
        Using db As DataTable = _SqlHelper.SelectDBRows(clsHT_DanhMuc.Sql_LoaiTn)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        tn_parent = tv_name.Nodes.Add(db.Rows(i)("ten_goi").ToString())
                        tn_parent.Tag = "00DM_" + db.Rows(i)("id").ToString().Trim()
                        strSQL = String.Format("Select * from HS_VuTaiNan Where IdLoaiTN = {0}", CType(db.Rows(i)("id").ToString().Trim(), Int32))
                        Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db_child Is Nothing) Then
                                If (db_child.Rows.Count > 0) Then
                                    For j As Integer = 0 To (db_child.Rows.Count - 1)
                                        tn_child = tn_parent.Nodes.Add(db_child.Rows(j)("TenVu_TN").ToString())
                                        tn_child.Tag = "01DM_" + db_child.Rows(j)("IdVuTaiNan").ToString().Trim()
                                    Next
                                End If
                            End If
                        End Using
                    Next
                End If
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu - Liên quan đến quản lý Tai nạn
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <param name="_state">
    '''               0: Lưới dữ liệu hồ sơ các vụ tai nạn.
    '''               1: Hồ sơ bảo hiểm Xã hội của Cán bộ
    '''               2: Hồ sơ bảo hiểm Y tế của Cán bộ
    '''               3: Hồ sơ Khám sức khoẻ của Cán bộ
    '''               4: Hồ sơ nghỉ phép của Cán bộ
    '''               5: Hồ sơ Cán bộ tai nạn
    ''' </param>
    ''' <remarks></remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal _state As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        Select Case _state
            Case 0       'Định nghĩa lưới dữ liệu - Lưới dữ liệu hồ sơ các vụ tai nạn
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tengoi", "Tên vụ tai nạn")
                dgv_name.Columns.Add("cln_LoaiTN", "Loại tai nạn")
                dgv_name.Columns.Add("cln_NgayTN", "Ngày tai nạn")
                dgv_name.Columns.Add("cln_ViTri", "Vị trí xẩy ra")
                dgv_name.Columns.Add("cln_NoiTN", "Nơi tai nạn")
                dgv_name.Columns.Add("cln_NgNhan", "Nguyên nhân tai nạn")
                dgv_name.Columns.Add("cln_SoTien", "Số tiền thiệt hại")
                dgv_name.Columns.Add("cln_Ngay_DT", "Ngày điều tra")
                dgv_name.Columns.Add("cln_Ngay_BC", "Ngày báo cáo")
                dgv_name.Columns.Add("cln_NguoiDT", "Người điều tra")
                dgv_name.Columns.Add("cln_MoTa", "Mô tả vụ tai nạn")
                dgv_name.Columns.Add("cln_DeNghi", "Đề nghị")
                dgv_name.Columns.Add("cln_GhiChu", "Ghi chú")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Tengoi").Width = 150
                dgv_name.Columns("cln_LoaiTN").Width = 180
                dgv_name.Columns("cln_NgayTN").Width = 85      'Columns:   Ngày xẩy ra
                dgv_name.Columns("cln_ViTri").Width = 95      'Columns:   Vị trí xẩy ra
                dgv_name.Columns("cln_NoiTN").Width = 180
                dgv_name.Columns("cln_NgNhan").Width = 210     'Columns:   Nguyên nhân tai nạn
                dgv_name.Columns("cln_SoTien").Width = 130     'Columns:   Số tiền thiệt hại
                dgv_name.Columns("cln_Ngay_DT").Width = 95     'Columns:   Ngày điều tra
                dgv_name.Columns("cln_Ngay_BC").Width = 95      'Columns:   Ngày báo cáo
                dgv_name.Columns("cln_NguoiDT").Width = 140
                dgv_name.Columns("cln_MoTa").Width = 230    'Columns:   Mô tả
                dgv_name.Columns("cln_DeNghi").Width = 180
                dgv_name.Columns("cln_GhiChu").Width = 180
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 15
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_NgayTN").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoTien").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Ngay_DT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Ngay_BC").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

            Case 1      'Định nghĩa lưới dữ liệu - Hồ sơ bảo hiểm xã hội của Cán bộ
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_SoHieu", "Số sổ")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Mucdong", "Mức đóng (%)")
                dgv_name.Columns.Add("cln_TenCq", "Tên cơ quan đóng BHXH")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_SoHieu").Width = 110
                dgv_name.Columns("cln_Tungay").Width = 88
                dgv_name.Columns("cln_Denngay").Width = 88
                dgv_name.Columns("cln_Mucdong").Width = 100
                dgv_name.Columns("cln_TenCq").Width = 210
                dgv_name.Columns("cln_Ghichu").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh dữ liệu hiển thị trên lưới dữ liệu
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Mucdong").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight

            Case 2      'Định nghĩa lưới dữ liệu - Hồ sơ bảo hiểm Y tế của Cán bộ
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_SoHieu", "Số thẻ")
                dgv_name.Columns.Add("cln_Ngaycap", "Ngày cấp")
                dgv_name.Columns.Add("cln_Noicap", "Nơi cấp")
                dgv_name.Columns.Add("cln_Noi_Dk", "Nơi đăng ký khám chữa bệnh")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_SoHieu").Width = 78
                dgv_name.Columns("cln_Ngaycap").Width = 88
                dgv_name.Columns("cln_Noicap").Width = 190
                dgv_name.Columns("cln_Noi_Dk").Width = 210
                dgv_name.Columns("cln_Tungay").Width = 88
                dgv_name.Columns("cln_Denngay").Width = 88
                dgv_name.Columns("cln_Ghichu").Width = 210

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 8
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh dữ liệu hiển thị trên lưới dữ liệu
                dgv_name.Columns("cln_Ngaycap").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

            Case 3      'Định nghĩa lưới dữ liệu - Hồ sơ Khám sức khoẻ của cán bộ
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Namkham", "Năm khám")
                dgv_name.Columns.Add("cln_DotKham", "Đợt khám")
                dgv_name.Columns.Add("cln_Noidung", "Nội dung khám sức khoẻ")
                dgv_name.Columns.Add("cln_Noikham", "Nơi khám sức khoẻ")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Cannang", "Cân nặng (Kg)")
                dgv_name.Columns.Add("cln_Chieucao", "Chiều cao (Kg)")
                dgv_name.Columns.Add("cln_LoaiSk", "Loại sức khoẻ")
                dgv_name.Columns.Add("cln_Ketluan", "Kết luận")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Namkham").Width = 73
                dgv_name.Columns("cln_DotKham").Width = 73
                dgv_name.Columns("cln_Noidung").Width = 280
                dgv_name.Columns("cln_Noikham").Width = 280
                dgv_name.Columns("cln_Tungay").Width = 88 'Columns: Từ ngày
                dgv_name.Columns("cln_Denngay").Width = 88
                dgv_name.Columns("cln_Cannang").Width = 95
                dgv_name.Columns("cln_Chieucao").Width = 95
                dgv_name.Columns("cln_LoaiSk").Width = 95
                dgv_name.Columns("cln_Ketluan").Width = 300
                dgv_name.Columns("cln_Ghichu").Width = 210
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 12
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh dữ liệu hiển thị trên lưới dữ liệu
                dgv_name.Columns("cln_Namkham").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DotKham").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Cannang").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Chieucao").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight

            Case 4      'Định nghĩa lưới dữ liệu - Hồ sơ Nghỉ phép của Cán bộ
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 55
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_NamNghi", "Năm")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_SoNgay", "Số ngày")
                dgv_name.Columns.Add("cln_LoaiNgPhep", "Loại nghỉ phép")
                dgv_name.Columns.Add("cln_Noi_NgPhep", "Nơi nghỉ phép")
                dgv_name.Columns.Add("cln_Lydonghi", "Lý do nghỉ phép")
                dgv_name.Columns.Add("cln_Nguoiky", "Người ký")
                dgv_name.Columns.Add("cln_TyLe", "Tỷ lệ hưởng lượng (%)")
                dgv_name.Columns.Add("cln_TienTC", "Tiền trợ cấp (VNĐ)")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_NamNghi").Width = 50
                dgv_name.Columns("cln_Tungay").Width = 83
                dgv_name.Columns("cln_Denngay").Width = 83
                dgv_name.Columns("cln_SoNgay").Width = 50  'Columns: Số ngày nghỉ
                dgv_name.Columns("cln_LoaiNgPhep").Width = 105
                dgv_name.Columns("cln_Noi_NgPhep").Width = 160
                dgv_name.Columns("cln_Lydonghi").Width = 230
                dgv_name.Columns("cln_Nguoiky").Width = 130
                dgv_name.Columns("cln_TyLe").Width = 80
                dgv_name.Columns("cln_TienTC").Width = 90
                dgv_name.Columns("cln_Ghichu").Width = 240

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 12
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh dữ liệu hiển thị trên lưới dữ liệu
                dgv_name.Columns("cln_NamNghi").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoNgay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_TyLe").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_TienTC").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.ColumnHeadersHeight = 38
            Case 5      'Định nghĩa lưới dữ liệu - Hồ sơ Cán bộ bị tai nạn
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tengoi", "Tên vụ tai nạn")
                dgv_name.Columns.Add("cln_Mucdo", "Mức độ")
                dgv_name.Columns.Add("cln_Songay", "Số ngày nghỉ")
                dgv_name.Columns.Add("cln_Chiphi", "Chi phí (VNĐ)")
                dgv_name.Columns.Add("cln_TTThuongtat", "Tình trạng thương tật")
                dgv_name.Columns.Add("cln_Noi_PPdieutri", "Nơi và phương pháp điều trị")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Tengoi").Width = 180
                dgv_name.Columns("cln_Mucdo").Width = 80
                dgv_name.Columns("cln_Songay").Width = 88
                dgv_name.Columns("cln_Chiphi").Width = 120  'Columns: Chi phí
                dgv_name.Columns("cln_TTThuongtat").Width = 190
                dgv_name.Columns("cln_Noi_PPdieutri").Width = 250
                dgv_name.Columns("cln_Ghichu").Width = 290

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 1 To 8
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh dữ liệu hiển thị trên lưới dữ liệu
                dgv_name.Columns("cln_Songay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Chiphi").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        End Select
        dgv_name.Columns("cln_Code").Visible = False
    End Sub

    ''' <summary>
    ''' Hàm trả về danh sách các bản ghi Hồ sơ vụ tai nạn theo Id loại tai nạn truyền vào
    ''' </summary>
    ''' <param name="_IdLoaiTN">Id loại tai nạn</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll_HS_VuTaiNan(ByVal _IdLoaiTN As Int32) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try

            Dim command As SqlCommand = New SqlCommand("HS_VuTaiNan_GetForIdLoaiTN", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdLoaiTN", SqlDbType.Int))
            command.Parameters("@_IdLoaiTN").Value = _IdLoaiTN

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_VuTaiNan")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_VuTaiNan")
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
    ''' Hàm trả về bản ghi - Hồ sơ vụ tai nạn dựa vào Id hồ sơ vụ tai nạn truyền vào
    ''' </summary>
    ''' <param name="_Code">Id hồ sơ vụ tai nạn</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRecord_VuTaiNan(ByVal _Code As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try

            Dim command As SqlCommand = New SqlCommand("HS_VuTaiNan_GetForId", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdVuTaiNan", SqlDbType.VarChar))
            command.Parameters("@_IdVuTaiNan").Value = _Code

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_VuTaiNan")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_VuTaiNan").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về danh sách các hồ sơ liên quan đến chính sách lao động theo id cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Chuỗi chỉ số xác định cán bộ</param>
    ''' <param name="_State">
    ''' Chỉ số phân loại hồ sơ. Với quy ước như sau:
    '''        1 - Hồ sơ BHXH
    '''        2 - Hồ sơ BHYT  
    '''        3 - Hồ sơ khám sức khoẻ
    '''        4 - HS nghỉ phép
    '''        5 - Hồ sơ cán bộ tai nạn
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll(ByVal _IdCanBo As String, ByVal _State As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand
            Select Case _State
                Case 1
                    command = New SqlCommand("HS_BHXH_GetForIdCanBo", connection)
                Case 2
                    command = New SqlCommand("HS_BHYT_GetForIdCanBo", connection)
                Case 3
                    command = New SqlCommand("HS_SucKhoe_GetForIdCanBo", connection)
                    'Case 4
                    '    command = New SqlCommand("HS_NgPh_GetForIdCanBo", connection)
                Case 5
                    command = New SqlCommand("CB_TaiNan_GetForIdCanBo", connection)
                Case Else
            End Select
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
            command.Parameters("@_IdCanBo").Value = _IdCanBo

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "TableRows")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("TableRows")
            End Using
        Catch ex As Exception
            MessageBox.Show("Lấy dữ liệu hồ sơ về chính sách lao động: " + ex.Message.ToString(), "Lỗi xẩy ra", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try

    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về row record thoả mãn điều kiện chuỗi id truyền vào
    ''' </summary>
    ''' <param name="_CodeId">Chuỗi Id hồ sơ</param>
    ''' <param name="_State">
    ''' Chỉ số phân loại hồ sơ. Với quy ước như sau:
    '''        1 - Hồ sơ BHXH
    '''        2 - Hồ sơ BHYT
    '''        3 - Hồ sơ khám sức khoẻ
    '''        4 - HS nghỉ phép
    '''        5 - Hồ sơ cán bộ tai nạn
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRecord(ByVal _CodeId As String, ByVal _State As Byte) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand = Nothing
            Select Case _State
                Case 1
                    command = New SqlCommand("HS_BHXH_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@_IdBHXH", SqlDbType.VarChar))
                    command.Parameters("@_IdBHXH").Value = _CodeId
                Case 2
                    command = New SqlCommand("HS_BHYT_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@_IdBHYT", SqlDbType.VarChar))
                    command.Parameters("@_IdBHYT").Value = _CodeId
                Case 3
                    command = New SqlCommand("HS_SucKhoe_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@_IdKhamChua", SqlDbType.VarChar))
                    command.Parameters("@_IdKhamChua").Value = _CodeId
                    'Case 4
                    '    command = New SqlCommand("HS_NgPh_GetForId", connection)
                    '    command.CommandType = CommandType.StoredProcedure

                    '    command.Parameters.Add(New SqlParameter("@_IdNghiPhep", SqlDbType.VarChar))
                    '    command.Parameters("@_IdNghiPhep").Value = _CodeId
                Case 5
                    command = New SqlCommand("CB_TaiNan_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@_IdCBTaiNan", SqlDbType.VarChar))
                    command.Parameters("@_IdCBTaiNan").Value = _CodeId
                Case Else
            End Select

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "TableRow")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("TableRow").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function
 
    ''' <summary>
    ''' Hàm lấy danh sách bản ghi Nghỉ phép/Nghỉ khác theo điều kiện truyền vào của cán bộ
    ''' </summary>
    ''' <param name="pDonViCd"></param>
    ''' <param name="pPhongBanCd"></param>
    ''' <param name="pIdCanBo"></param>
    ''' <param name="pIdNghiPhep"></param>
    ''' <param name="pNamNghi"></param>
    ''' <param name="pIdLoaiNghi"></param>
    ''' <param name="pLoaiNghiCd"></param>
    ''' <param name="pThoiDiemDL">Thời điểm dữ liệu cán bộ (Nếu trống lấy Ngày hiện thời). Giá trị định dạng: dd/MM/yyyy</param>
    ''' <param name="pIsAll">0 - Lấy theo mã Đơn vị; 1 - Lấy cả các đơn vị trực thuộc</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HS_NgPh_GetSearch(ByVal pDonViCd As String, ByVal pPhongBanCd As String, ByVal pIdCanBo As String, ByVal pIdNghiPhep As String, ByVal pNamNghi As Integer, ByVal pIdLoaiNghi As Integer, ByVal pLoaiNghiCd As String, ByVal pThoiDiemDL As String, ByVal pIsAll As Byte) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("HS_NgPh_GetSearch", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViCd", pDonViCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhongBanCd", pPhongBanCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdCanBo", pIdCanBo, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdNghiPhep", pIdNghiPhep, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pNamNghi", pNamNghi, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdLoaiNghi", pIdLoaiNghi, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pLoaiNghiCd", pLoaiNghiCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiemDL", pThoiDiemDL, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIsAll", pIsAll, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_NgPh_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_NgPh_TMP")
            End Using
        Catch ex As Exception
            Throw ex
            Globals.Logger.Error("Lỗi Lấy danh sách nghỉ phép HS_NgPh_GetSearch(): " + ex.Message)
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

#End Region

End Class

Public Class ThuNhapGiaDinh
    Dim _IdThuNhapGD As String
    Dim _IdCanBo As String
    Dim _NgayKekhai As Date
    Dim _Luong As String '] [nvarchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Dim _NguonKhac As String '] [nvarchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Dim _NhaO_DuocCap As String '] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Dim _NhaO_DuocCap_DT As Double
    Dim _NhaO_TuMua As String '] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
    Dim _NhaO_TuMua_DT As Double
    Dim _DatO_DuocCap_DT As Double
    Dim _DatO_TuMua_DT As Double
    Dim _DatSXKT As String '] [nvarchar] (300)
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdThuNhapGD() As String
        Get
            Return _IdThuNhapGD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID thu nhập cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdThuNhapGD = Value
        End Set
    End Property

    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdCanBo = Value
        End Set
    End Property

    Public Property NgayKekhai() As Date
        Get
            Return _NgayKekhai
        End Get
        Set(ByVal Value As Date)
            _NgayKekhai = Value
        End Set
    End Property
    
    Public Property Luong() As String
        Get
            Return _Luong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nguồn thu nhập từ lương có độ dài quá 200 kí tự!", Value, Value.ToString())
            End If
            _Luong = Value
        End Set
    End Property

    Public Property NguonKhac() As String
        Get
            Return _NguonKhac
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 300) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị thu nhập từ nguồn khác có độ dài quá 300 kí tự!", Value, Value.ToString())
            End If
            _NguonKhac = Value
        End Set
    End Property

    Public Property NhaO_DuocCap() As String
        Get
            Return _NhaO_DuocCap
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nhà ở được cấp/thuê có độ dài quá 50 kí tự!", Value, Value.ToString())
            End If
            _NhaO_DuocCap = Value
        End Set
    End Property

    Public Property NhaO_DuocCap_DT() As Double
        Get
            Return _NhaO_DuocCap_DT
        End Get
        Set(ByVal Value As Double)
            _NhaO_DuocCap_DT = Value
        End Set
    End Property

    Public Property NhaO_TuMua() As String
        Get
            Return _NhaO_TuMua
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nhà ở tự mua/xây có độ dài quá 50 kí tự!", Value, Value.ToString())
            End If
            _NhaO_TuMua = Value
        End Set
    End Property

    Public Property NhaO_TuMua_DT() As Double
        Get
            Return _NhaO_TuMua_DT
        End Get
        Set(ByVal Value As Double)
            _NhaO_TuMua_DT = Value
        End Set
    End Property

    Public Property DatO_DuocCap_DT() As Double
        Get
            Return _DatO_DuocCap_DT
        End Get
        Set(ByVal Value As Double)
            _DatO_DuocCap_DT = Value
        End Set
    End Property

    Public Property DatO_TuMua_DT() As Double
        Get
            Return _DatO_TuMua_DT
        End Get
        Set(ByVal Value As Double)
            _DatO_TuMua_DT = Value
        End Set
    End Property

    Public Property DatSXKT() As String
        Get
            Return _DatSXKT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 300) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị đất sản xuất kinh doanh có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _DatSXKT = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ThuNhapGD_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@NgayKekhai", NgayKekhai))
            cmd.Parameters.Add(New SqlParameter("@Luong", Luong))
            cmd.Parameters.Add(New SqlParameter("@NguonKhac", NguonKhac))
            cmd.Parameters.Add(New SqlParameter("@NhaO_DuocCap", NhaO_DuocCap))
            cmd.Parameters.Add(New SqlParameter("@NhaO_DuocCap_DT", NhaO_DuocCap_DT))
            cmd.Parameters.Add(New SqlParameter("@NhaO_TuMua", NhaO_TuMua))
            cmd.Parameters.Add(New SqlParameter("@NhaO_TuMua_DT", NhaO_TuMua_DT))
            cmd.Parameters.Add(New SqlParameter("@DatO_DuocCap_DT", DatO_DuocCap_DT))
            cmd.Parameters.Add(New SqlParameter("@DatO_TuMua_DT", DatO_TuMua_DT))
            cmd.Parameters.Add(New SqlParameter("@DatSXKT", DatSXKT))
            cmd.Parameters.Add("@IdThuNhapGD", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdThuNhapGD = cmd.Parameters("@IdThuNhapGD").Value.ToString
        Catch ex As Exception
            IdThuNhapGD = ""
        End Try
        Return IdThuNhapGD
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ThuNhapGD_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdThuNhapGD", IdThuNhapGD))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NgayKekhai", NgayKekhai))
            cmd.Parameters.Add(New SqlParameter("@Luong", Luong))
            cmd.Parameters.Add(New SqlParameter("@NguonKhac", NguonKhac))
            cmd.Parameters.Add(New SqlParameter("@NhaO_DuocCap", NhaO_DuocCap))
            cmd.Parameters.Add(New SqlParameter("@NhaO_DuocCap_DT", NhaO_DuocCap_DT))
            cmd.Parameters.Add(New SqlParameter("@NhaO_TuMua", NhaO_TuMua))
            cmd.Parameters.Add(New SqlParameter("@NhaO_TuMua_DT", NhaO_TuMua_DT))
            cmd.Parameters.Add(New SqlParameter("@DatO_DuocCap_DT", DatO_DuocCap_DT))
            cmd.Parameters.Add(New SqlParameter("@DatO_TuMua_DT", DatO_TuMua_DT))
            cmd.Parameters.Add(New SqlParameter("@DatSXKT", DatSXKT))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ThuNhapGD_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdThuNhapGD", IdThuNhapGD))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            db.executeSQL(cmd)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function init(ByVal cmd As SqlCommand) As IList
        Dim conn As SqlConnection = db.getConnection
        cmd.Connection = conn
        Dim list As ArrayList = New ArrayList
        Try
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim smartReader As SmartDataReader = New SmartDataReader(reader)
            While smartReader.Read
                Dim m_ThuNhapGD As ThuNhapGiaDinh = New ThuNhapGiaDinh
                m_ThuNhapGD.IdThuNhapGD = smartReader.getString("IdThuNhapGD")
                m_ThuNhapGD.IdCanBo = smartReader.getString("IdCanBo")
                m_ThuNhapGD.NgayKekhai = smartReader.getDatetime("NgayKekhai")
                m_ThuNhapGD.Luong = smartReader.getString("Luong")
                m_ThuNhapGD.NguonKhac = smartReader.getString("NguonKhac")
                m_ThuNhapGD.NhaO_DuocCap = smartReader.getString("NhaO_DuocCap")
                m_ThuNhapGD.NhaO_DuocCap_DT = smartReader.getFloat("NhaO_DuocCap_DT")
                m_ThuNhapGD.NhaO_TuMua = smartReader.getString("NhaO_TuMua")
                m_ThuNhapGD.NhaO_TuMua_DT = smartReader.getFloat("NhaO_TuMua_DT")
                m_ThuNhapGD.DatO_DuocCap_DT = smartReader.getFloat("DatO_DuocCap_DT")
                m_ThuNhapGD.DatO_TuMua_DT = smartReader.getFloat("DatO_TuMua_DT")
                m_ThuNhapGD.DatSXKT = smartReader.getString("DatSXKT")
                list.Add(m_ThuNhapGD)
            End While
            smartReader.disposeReader(reader)
            Return list
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            db.closeConnection(conn)
        End Try

    End Function

    Public Function getAll() As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ThuNhapGD Order by NgayKekhai desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByCanbo(ByVal vIdCanbo As String) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ThuNhapGD WHERE idCanbo='" & vIdCanbo.Trim & "' order by NgayKekhai desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdThuNhapGD As String) As ThuNhapGiaDinh
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ThuNhapGD WHERE IdThuNhapGD='" & vIdThuNhapGD.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), ThuNhapGiaDinh)
            Else
                Return New ThuNhapGiaDinh
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As ThuNhapGiaDinh
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_ThuNhapGD WHERE idCanbo='" & vIdCanbo.Trim & "' order by NgayKekhai desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), ThuNhapGiaDinh)
            Else
                Return New ThuNhapGiaDinh
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
#End Region

End Class

