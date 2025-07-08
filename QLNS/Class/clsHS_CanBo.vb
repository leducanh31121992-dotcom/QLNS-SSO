Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class clsHS_CanBo
    Private _SqlHelper As DBAccess
    Private _HS_Dang As clsHS_DangDT = New clsHS_DangDT()

    Public Sub New()
        _SqlHelper = New DBAccess
    End Sub
    'Biến lưu đường dẫn ảnh lock icon

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ cán bộ<---"

    ''' <summary>
    ''' Khai báo thuộc tính và get - set giá trị đối tượng
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HS_CanBo
        'Khai báo thuộc tính
        Private _IdCanBo As String = ""
        Private _MaCB As String = ""
        Private _HoTen As String = ""
        Private _IdDonVi As Integer = 0
        Private _TenThuongGoi As String = ""
        Private _BiDanh As String = ""
        Private _GioiTinh As Byte
        Private _NgaySinh As DateTime = DateTime.Now
        Private _IdNS_Tinh As Integer = 0
        Private _IdNS_Huyen As Integer = 0
        Private _IdNS_Xa As Integer = 0
        Private _IdNS_Thon As Integer = 0
        Private _NS_DChi As String = ""

        Private _IdNQ_Tinh As Integer = 0
        Private _IdNQ_Huyen As Integer = 0
        Private _IdNQ_Xa As Integer = 0
        Private _IdNQ_Thon As Integer = 0
        Private _NQ_DChi As String = ""

        Private _IdThT_Tinh As Integer = 0
        Private _IdThT_Huyen As Integer = 0
        Private _IdThT_Xa As Integer = 0
        Private _IdThT_Thon As Integer = 0
        Private _ThT_Diachi As String = ""
        Private _ThT_Dienthoai As String = ""

        Private _IdTTr_Tinh As Integer = 0
        Private _IdTTr_Huyen As Integer = 0
        Private _IdTTr_Xa As Integer = 0
        Private _IdTTr_Thon As Integer = 0
        Private _TTr_Diachi As String = ""
        Private _TTr_Dienthoai As String = ""

        Private _IdQuocTich As Integer = 0
        Private _IdDanToc As Integer = 0
        Private _IdTonGiao As Integer = 0
        Private _IdThanhPhanGD As Integer = 0

        Private _NhomMau As String = ""
        Private _CMT_So As String = ""
        Private _CMT_NgayCap As DateTime = DateTime.Now
        Private _CMT_NoiCap As String = ""

        Private _IdUT_GDinh As Integer = 0
        Private _IdUT_BThan As String = ""

        Private _Ngay_ThamNien As DateTime = DateTime.Now
        Private _Ngay_NH As DateTime = DateTime.Now
        Private _Ngay_VBSP As DateTime = DateTime.Now
        Private _Ngay_BienChe As DateTime = DateTime.Now
        Private _Ngay_CQ As DateTime = DateTime.Now

        Private _Email As String
        Private _DienThoai_NR As String = ""
        Private _DienThoai_DD As String = ""
        Private _DienThoai_CQ As String = ""
        Private _SoFax As String = ""

        Private _IdTrinhDoVH As Integer = 0
        Private _IdTrinhDoCT As Integer = 0
        Private _HonNhan_Cd As String = ""

        Private _SoTruong_CT As String = ""
        Private _CV_Lau As String = ""
        Private _CM_Ngay As DateTime = DateTime.Now
        Private _CM_ToChuc As String = ""
        Private _DacDiem_BT As String = ""
        Private _QuanHe_Nguoi_NN As String = ""

        Private _NH_MaKH As String = ""
        Private _NH_SoTK As String = ""
        Private _NH_TenNH As String = ""

        Private _MaSoThue As String = ""

        Private _BHXH_SoSo As String = ""
        Private _BHXH_NgaySo As DateTime = DateTime.Now
        Private _BHXH_NgayBatDau As DateTime = DateTime.Now
        Private _BHXH_NoiLam As String = ""
        Private _AnhThe As Byte()
        Private _GhiChu As String = ""
        Private _IdNew As String = ""

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

        Public Property MaCB() As String
            Get
                Return _MaCB
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã cán bộ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MaCB = value
            End Set
        End Property

        Public Property HoTen() As String
            Get
                Return _HoTen
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị họ tên cán bộ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _HoTen = value
            End Set
        End Property

        Public Property IdDonVi() As Integer
            Get
                Return _IdDonVi
            End Get
            Set(ByVal value As Integer)
                _IdDonVi = value
            End Set
        End Property

        Public Property TenThuongGoi() As String
            Get
                Return _TenThuongGoi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên thường gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TenThuongGoi = value
            End Set
        End Property

        Public Property BiDanh() As String
            Get
                Return _BiDanh
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị bí danh gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BiDanh = value
            End Set
        End Property

        Public Property GioiTinh() As Byte
            Get
                Return _GioiTinh
            End Get
            Set(ByVal value As Byte)
                _GioiTinh = value
            End Set
        End Property

        Public Property NgaySinh() As DateTime
            Get
                Return _NgaySinh
            End Get
            Set(ByVal value As DateTime)
                _NgaySinh = value
            End Set
        End Property

        Public Property IdNS_Tinh() As Integer
            Get
                Return _IdNS_Tinh
            End Get
            Set(ByVal value As Integer)
                _IdNS_Tinh = value
            End Set
        End Property

        Public Property IdNS_Huyen() As Integer
            Get
                Return _IdNS_Huyen
            End Get
            Set(ByVal value As Integer)
                _IdNS_Huyen = value
            End Set
        End Property

        Public Property IdNS_Xa() As Integer
            Get
                Return _IdNS_Xa
            End Get
            Set(ByVal value As Integer)
                _IdNS_Xa = value
            End Set
        End Property
        Public Property IdNS_Thon() As Integer
            Get
                Return _IdNS_Thon
            End Get
            Set(ByVal value As Integer)
                _IdNS_Thon = value
            End Set
        End Property

        Public Property NS_DChi() As String
            Get
                Return _NS_DChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ nơi sinh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NS_DChi = value
            End Set
        End Property

        Public Property IdNQ_Tinh() As Integer
            Get
                Return _IdNQ_Tinh
            End Get
            Set(ByVal value As Integer)
                _IdNQ_Tinh = value
            End Set
        End Property

        Public Property IdNQ_Huyen() As Integer
            Get
                Return _IdNQ_Huyen
            End Get
            Set(ByVal value As Integer)
                _IdNQ_Huyen = value
            End Set
        End Property

        Public Property IdNQ_Xa() As Integer
            Get
                Return _IdNQ_Xa
            End Get
            Set(ByVal value As Integer)
                _IdNQ_Xa = value
            End Set
        End Property

        Public Property IdNQ_Thon() As Integer
            Get
                Return _IdNQ_Thon
            End Get
            Set(ByVal value As Integer)
                _IdNQ_Thon = value
            End Set
        End Property

        Public Property NQ_DChi() As String
            Get
                Return _NQ_DChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ nguyên quán có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NQ_DChi = value
            End Set
        End Property

        Public Property IdThT_Tinh() As Integer
            Get
                Return _IdThT_Tinh
            End Get
            Set(ByVal value As Integer)
                _IdThT_Tinh = value
            End Set
        End Property

        Public Property IdThT_Huyen() As Integer
            Get
                Return _IdThT_Huyen
            End Get
            Set(ByVal value As Integer)
                _IdThT_Huyen = value
            End Set
        End Property

        Public Property IdThT_Xa() As Integer
            Get
                Return _IdThT_Xa
            End Get
            Set(ByVal value As Integer)
                _IdThT_Xa = value
            End Set
        End Property

        Public Property IdThT_Thon() As Integer
            Get
                Return _IdThT_Thon
            End Get
            Set(ByVal value As Integer)
                _IdThT_Thon = value
            End Set
        End Property

        Public Property ThT_Diachi() As String
            Get
                Return _ThT_Diachi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ thường trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ThT_Diachi = value
            End Set
        End Property

        Public Property ThT_Dienthoai() As String
            Get
                Return _ThT_Dienthoai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại thường trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ThT_Dienthoai = value
            End Set
        End Property

        Public Property IdTTr_Tinh() As Integer
            Get
                Return _IdTTr_Tinh
            End Get
            Set(ByVal value As Integer)
                _IdTTr_Tinh = value
            End Set
        End Property

        Public Property IdTTr_Huyen() As Integer
            Get
                Return _IdTTr_Huyen
            End Get
            Set(ByVal value As Integer)
                _IdTTr_Huyen = value
            End Set
        End Property

        Public Property IdTTr_Xa() As Integer
            Get
                Return _IdTTr_Xa
            End Get
            Set(ByVal value As Integer)
                _IdTTr_Xa = value
            End Set
        End Property

        Public Property IdTTr_Thon() As Integer
            Get
                Return _IdTTr_Thon
            End Get
            Set(ByVal value As Integer)
                _IdTTr_Thon = value
            End Set
        End Property

        Public Property TTr_Diachi() As String
            Get
                Return _TTr_Diachi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ tạm trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TTr_Diachi = value
            End Set
        End Property

        Public Property TTr_Dienthoai() As String
            Get
                Return _TTr_Dienthoai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại tạm trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TTr_Dienthoai = value
            End Set
        End Property

        Public Property IdQuocTich() As Integer
            Get
                Return _IdQuocTich
            End Get
            Set(ByVal value As Integer)
                _IdQuocTich = value
            End Set
        End Property

        Public Property IdDanToc() As Integer
            Get
                Return _IdDanToc
            End Get
            Set(ByVal value As Integer)
                _IdDanToc = value
            End Set
        End Property

        Public Property IdTonGiao() As Integer
            Get
                Return _IdTonGiao
            End Get
            Set(ByVal value As Integer)
                _IdTonGiao = value
            End Set
        End Property

        Public Property IdThanhPhanGD() As Integer
            Get
                Return _IdThanhPhanGD
            End Get
            Set(ByVal value As Integer)
                _IdThanhPhanGD = value
            End Set
        End Property

        Public Property NhomMau() As String
            Get
                Return _NhomMau
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 2) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nhóm máu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NhomMau = value
            End Set
        End Property

        Public Property CMT_So() As String
            Get
                Return _CMT_So
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số chứng minh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CMT_So = value
            End Set
        End Property

        Public Property CMT_NgayCap() As DateTime
            Get
                Return _CMT_NgayCap
            End Get
            Set(ByVal value As DateTime)
                _CMT_NgayCap = value
            End Set
        End Property

        Public Property CMT_NoiCap() As String
            Get
                Return _CMT_NoiCap
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp chứng minh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CMT_NoiCap = value
            End Set
        End Property

        Public Property IdUT_GDinh() As Integer
            Get
                Return _IdUT_GDinh
            End Get
            Set(ByVal value As Integer)
                _IdUT_GDinh = value
            End Set
        End Property

        Public Property IdUT_BThan() As String
            Get
                Return _IdUT_BThan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ưu tiên bản thân có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdUT_BThan = value
            End Set
        End Property

        Public Property Ngay_ThamNien() As DateTime
            Get
                Return _Ngay_ThamNien
            End Get
            Set(ByVal value As DateTime)
                _Ngay_ThamNien = value
            End Set
        End Property

        Public Property Ngay_NH() As DateTime
            Get
                Return _Ngay_NH
            End Get
            Set(ByVal value As DateTime)
                _Ngay_NH = value
            End Set
        End Property

        Public Property Ngay_VBSP() As DateTime
            Get
                Return _Ngay_VBSP
            End Get
            Set(ByVal value As DateTime)
                _Ngay_VBSP = value
            End Set
        End Property

        Public Property Ngay_BienChe() As DateTime
            Get
                Return _Ngay_BienChe
            End Get
            Set(ByVal value As DateTime)
                _Ngay_BienChe = value
            End Set
        End Property

        Public Property Ngay_CQ() As DateTime
            Get
                Return _Ngay_CQ
            End Get
            Set(ByVal value As DateTime)
                _Ngay_CQ = value
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

        Public Property DienThoai_NR() As String
            Get
                Return _DienThoai_NR
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại nhà riêng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai_NR = value
            End Set
        End Property

        Public Property DienThoai_DD() As String
            Get
                Return _DienThoai_DD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại di động có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai_DD = value
            End Set
        End Property

        Public Property DienThoai_CQ() As String
            Get
                Return _DienThoai_CQ
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại cơ quan có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai_CQ = value
            End Set
        End Property

        Public Property SoFax() As String
            Get
                Return _SoFax
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số máy fax cơ quan có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoFax = value
            End Set
        End Property

        Public Property IdTrinhDoVH() As Integer
            Get
                Return _IdTrinhDoVH
            End Get
            Set(ByVal value As Integer)
                _IdTrinhDoVH = value
            End Set
        End Property

        Public Property IdTrinhDoCT() As Integer
            Get
                Return _IdTrinhDoCT
            End Get
            Set(ByVal value As Integer)
                _IdTrinhDoCT = value
            End Set
        End Property
        Public Property HonNhan_Cd() As String
            Get
                Return _HonNhan_Cd
            End Get
            Set(ByVal value As String)
                _HonNhan_Cd = value
            End Set
        End Property
        Public Property SoTruong_CT() As String
            Get
                Return _SoTruong_CT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị sở trường công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoTruong_CT = value
            End Set
        End Property

        Public Property CV_Lau() As String
            Get
                Return _CV_Lau
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị công việc làm lâu nhất có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CV_Lau = value
            End Set
        End Property

        Public Property CM_Ngay() As DateTime
            Get
                Return _CM_Ngay
            End Get
            Set(ByVal value As DateTime)
                _CM_Ngay = value
            End Set
        End Property

        Public Property CM_ToChuc() As String
            Get
                Return _CM_ToChuc
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tổ chức cách mạng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CM_ToChuc = value
            End Set
        End Property

        Public Property DacDiem_BT() As String
            Get
                Return _DacDiem_BT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị đặc điểm bản thân có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DacDiem_BT = value
            End Set
        End Property

        Public Property QuanHe_Nguoi_NN() As String
            Get
                Return _QuanHe_Nguoi_NN
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị quan hệ người nước ngoài có độ dài không hợp lệ!", value, value.ToString())
                End If
                _QuanHe_Nguoi_NN = value
            End Set
        End Property

        Public Property NH_MaKH() As String
            Get
                Return _NH_MaKH
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã khách hàng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NH_MaKH = value
            End Set
        End Property

        Public Property NH_SoTK() As String
            Get
                Return _NH_SoTK
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số tài khoản có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NH_SoTK = value
            End Set
        End Property

        Public Property NH_TenNH() As String
            Get
                Return _NH_TenNH
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên ngân hàng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NH_TenNH = value
            End Set
        End Property

        Public Property MaSoThue() As String
            Get
                Return _MaSoThue
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã số thuế có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MaSoThue = value
            End Set
        End Property

        Public Property BHXH_SoSo() As String
            Get
                Return _BHXH_SoSo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số sổ BHXH có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BHXH_SoSo = value
            End Set
        End Property

        Public Property BHXH_NgaySo() As DateTime
            Get
                Return _BHXH_NgaySo
            End Get
            Set(ByVal value As DateTime)
                _BHXH_NgaySo = value
            End Set
        End Property

        Public Property BHXH_NgayBatDau() As DateTime
            Get
                Return _BHXH_NgayBatDau
            End Get
            Set(ByVal value As DateTime)
                _BHXH_NgayBatDau = value
            End Set
        End Property

        Public Property BHXH_NoiLam() As String
            Get
                Return _BHXH_NoiLam
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi làm BHXH có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BHXH_NoiLam = value
            End Set
        End Property

        Public Property AnhThe() As Byte()
            Get
                Return _AnhThe
            End Get
            Set(ByVal value As Byte())
                _AnhThe = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú thêm có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

        Public Property IdNew() As String
            Get
                Return _IdNew
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id mới có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdNew = value
            End Set
        End Property
    End Class

#End Region

#Region "---> Hồ sơ cán bộ tmp - Bang tam luu thong thong tin import tu file excel <---"

#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ cán bộ Tập sự<---"
    Public Class HSCB_TS
        Private _Id As String = ""
        Private _MaCB As String = ""
        Private _HoTen As String = ""
        Private _IdChiNhanh As Integer = 0
        Private _IdPhongBan As String = ""
        Private _TenThuongGoi As String = ""
        Private _BiDanh As String = ""
        Private _GioiTinh As Byte
        Private _NgaySinh As DateTime = DateTime.Now
        Private _IdNS_Tinh As Integer = 0
        Private _IdNS_Huyen As Integer = 0
        Private _IdNS_Xa As Integer = 0
        Private _IdNS_Thon As Integer = 0
        Private _NS_DiaChi As String = ""

        Private _IdNQ_Tinh As Integer = 0
        Private _IdNQ_Huyen As Integer = 0
        Private _IdNQ_Xa As Integer = 0
        Private _IdNQ_Thon As Integer = 0
        Private _NQ_DiaChi As String = ""

        Private _IdThT_Tinh As Integer = 0
        Private _IdThT_Huyen As Integer = 0
        Private _IdThT_Xa As Integer = 0
        Private _IdThT_Thon As Integer = 0
        Private _ThT_DiaChi As String = ""
        Private _ThT_DienThoai As String = ""

        Private _IdTTr_Tinh As Integer = 0
        Private _IdTTr_Huyen As Integer = 0
        Private _IdTTr_Xa As Integer = 0
        Private _IdTTr_Thon As Integer = 0
        Private _TTr_DiaChi As String = ""
        Private _TTr_DienThoai As String = ""

        Private _IdQuocTich As Integer = 0
        Private _IdDanToc As Integer = 0
        Private _IdTonGiao As Integer = 0
        Private _IdThanhPhan_GD As Integer = 0
        Private _NhomMau As String = ""
        Private _CMT_So As String = ""
        Private _CMT_NgayCap As DateTime = DateTime.Now
        Private _CMT_NoiCap As String = ""
        Private _IdUT_GD As Integer = 0
        Private _IdUT_BT As String = ""
        Private _Email As String = ""
        Private _DienThoai_NR As String = ""
        Private _DienThoai_DD As String = ""
        Private _DienThoai_CQ As String = ""
        Private _SoFax As String = ""
        Private _IdTrinhDoVH As Integer = 0
        Private _IdTrinhDoCT As Integer = 0
        Private _SoTruong_CT As String = ""
        Private _CV_Lau As String = ""
        Private _CM_Ngay As DateTime = DateTime.Now
        Private _CM_ToChuc As String = ""
        Private _Dang_NgayVao As DateTime = DateTime.Now
        Private _Dang_NgayChTh As DateTime = DateTime.Now
        Private _Dang_NgayRa As DateTime = DateTime.Now
        Private _Dang_NoiKN As String = ""
        Private _Dang_SoThe As String = ""
        Private _Dang_NGT As String = ""
        Private _Dang_LyDoRa As String = ""
        Private _NH_MSKH As String = ""
        Private _NH_SHTK As String = ""
        Private _NH_Ten_NH As String = ""
        Private _MaSoThue As String = ""
        Private _BHXH_SoSo As String = ""
        Private _BHXH_NgayLam As DateTime = DateTime.Now
        Private _BHXH_NgayDong As DateTime = DateTime.Now
        Private _BHXH_NoiLam As String = ""
        Private _IdHocHam As Integer = 0
        Private _IdHocVi As Integer = 0
        Private _IdTrdChMon As Integer = 0
        Private _IdTrdNgoaiNgu As Integer = 0
        Private _IdTrdTinHoc As Integer = 0
        Private _IdChuyenNganhDT As Integer = 0
        Private _AnhThe As Byte()
        Private _GhiChu As String = ""
        Private _idNew As String = ""
        Private _TrangThai As Byte
        Private _HonNhan_Cd As String

        Public Property Id() As String
            Get
                Return _Id
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Id = value
            End Set
        End Property

        Public Property MaCB() As String
            Get
                Return _MaCB
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã cán bộ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MaCB = value
            End Set
        End Property

        Public Property HoTen() As String
            Get
                Return _HoTen
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị họ tên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _HoTen = value
            End Set
        End Property

        Public Property IdChiNhanh() As Int32
            Get
                Return _IdChiNhanh
            End Get
            Set(ByVal value As Int32)
                _IdChiNhanh = value
            End Set
        End Property

        Public Property IdPhongBan() As String
            Get
                Return _IdPhongBan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 8) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị phòng ban có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdPhongBan = value
            End Set
        End Property

        Public Property TenThuongGoi() As String
            Get
                Return _TenThuongGoi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên thường gọi có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TenThuongGoi = value
            End Set
        End Property

        Public Property BiDanh() As String
            Get
                Return _BiDanh
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị bí danh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BiDanh = value
            End Set
        End Property

        Public Property GioiTinh() As Byte
            Get
                Return _GioiTinh
            End Get
            Set(ByVal value As Byte)
                _GioiTinh = value
            End Set
        End Property

        Public Property NgaySinh() As DateTime
            Get
                Return _NgaySinh
            End Get
            Set(ByVal value As DateTime)
                _NgaySinh = value
            End Set
        End Property

        Public Property IdNS_Tinh() As Int32
            Get
                Return _IdNS_Tinh
            End Get
            Set(ByVal value As Int32)
                _IdNS_Tinh = value
            End Set
        End Property

        Public Property IdNS_Huyen() As Int32
            Get
                Return _IdNS_Huyen
            End Get
            Set(ByVal value As Int32)
                _IdNS_Huyen = value
            End Set
        End Property

        Public Property IdNS_Xa() As Integer
            Get
                Return _IdNS_Xa
            End Get
            Set(ByVal value As Integer)
                _IdNS_Xa = value
            End Set
        End Property
        Public Property IdNS_Thon() As Integer
            Get
                Return _IdNS_Thon
            End Get
            Set(ByVal value As Integer)
                _IdNS_Thon = value
            End Set
        End Property

        Public Property NS_DiaChi() As String
            Get
                Return _NS_DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ nơi sinh có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NS_DiaChi = value
            End Set
        End Property

        Public Property IdNQ_Tinh() As Int32
            Get
                Return _IdNQ_Tinh
            End Get
            Set(ByVal value As Int32)
                _IdNQ_Tinh = value
            End Set
        End Property

        Public Property IdNQ_Huyen() As Int32
            Get
                Return _IdNQ_Huyen
            End Get
            Set(ByVal value As Int32)
                _IdNQ_Huyen = value
            End Set
        End Property

        Public Property IdNQ_Xa() As Integer
            Get
                Return _IdNQ_Xa
            End Get
            Set(ByVal value As Integer)
                _IdNQ_Xa = value
            End Set
        End Property

        Public Property IdNQ_Thon() As Integer
            Get
                Return _IdNQ_Thon
            End Get
            Set(ByVal value As Integer)
                _IdNQ_Thon = value
            End Set
        End Property

        Public Property NQ_DiaChi() As String
            Get
                Return _NQ_DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ nguyên quán có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NQ_DiaChi = value
            End Set
        End Property

        Public Property IdThT_Tinh() As Int32
            Get
                Return _IdThT_Tinh
            End Get
            Set(ByVal value As Int32)
                _IdThT_Tinh = value
            End Set
        End Property

        Public Property IdThT_Huyen() As Int32
            Get
                Return _IdThT_Huyen
            End Get
            Set(ByVal value As Int32)
                _IdThT_Huyen = value
            End Set
        End Property

        Public Property IdThT_Xa() As Integer
            Get
                Return _IdThT_Xa
            End Get
            Set(ByVal value As Integer)
                _IdThT_Xa = value
            End Set
        End Property

        Public Property IdThT_Thon() As Integer
            Get
                Return _IdThT_Thon
            End Get
            Set(ByVal value As Integer)
                _IdThT_Thon = value
            End Set
        End Property

        Public Property ThT_DiaChi() As String
            Get
                Return _ThT_DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ thường trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ThT_DiaChi = value
            End Set
        End Property

        Public Property ThT_DienThoai() As String
            Get
                Return _ThT_DienThoai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại thường trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ThT_DienThoai = value
            End Set
        End Property

        Public Property IdTTr_Tinh() As Int32
            Get
                Return _IdTTr_Tinh
            End Get
            Set(ByVal value As Int32)
                _IdTTr_Tinh = value
            End Set
        End Property

        Public Property IdTTr_Huyen() As Int32
            Get
                Return _IdTTr_Huyen
            End Get
            Set(ByVal value As Int32)
                _IdTTr_Huyen = value
            End Set
        End Property

        Public Property IdTTr_Xa() As Integer
            Get
                Return _IdTTr_Xa
            End Get
            Set(ByVal value As Integer)
                _IdTTr_Xa = value
            End Set
        End Property

        Public Property IdTTr_Thon() As Integer
            Get
                Return _IdTTr_Thon
            End Get
            Set(ByVal value As Integer)
                _IdTTr_Thon = value
            End Set
        End Property

        Public Property TTr_DiaChi() As String
            Get
                Return _TTr_DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ tạm trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TTr_DiaChi = value
            End Set
        End Property

        Public Property TTr_DienThoai() As String
            Get
                Return _TTr_DienThoai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại tạm trú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TTr_DienThoai = value
            End Set
        End Property

        Public Property IdQuocTich() As Int32
            Get
                Return _IdQuocTich
            End Get
            Set(ByVal value As Int32)
                _IdQuocTich = value
            End Set
        End Property

        Public Property IdDanToc() As Int32
            Get
                Return _IdDanToc
            End Get
            Set(ByVal value As Int32)
                _IdDanToc = value
            End Set
        End Property

        Public Property IdTonGiao() As Int32
            Get
                Return _IdTonGiao
            End Get
            Set(ByVal value As Int32)
                _IdTonGiao = value
            End Set
        End Property

        Public Property IdThanhPhan_GD() As Int32
            Get
                Return _IdThanhPhan_GD
            End Get
            Set(ByVal value As Int32)
                _IdThanhPhan_GD = value
            End Set
        End Property

        Public Property NhomMau() As String
            Get
                Return _NhomMau
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 2) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nhóm máu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NhomMau = value
            End Set
        End Property

        Public Property CMT_So() As String
            Get
                Return _CMT_So
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số chứng minh thư có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CMT_So = value
            End Set
        End Property

        Public Property CMT_NgayCap() As DateTime
            Get
                Return _CMT_NgayCap
            End Get
            Set(ByVal value As DateTime)
                _CMT_NgayCap = value
            End Set
        End Property

        Public Property CMT_NoiCap() As String
            Get
                Return _CMT_NoiCap
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp chứng minh thư có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CMT_NoiCap = value
            End Set
        End Property

        Public Property IdUT_GD() As Int32
            Get
                Return _IdUT_GD
            End Get
            Set(ByVal value As Int32)
                _IdUT_GD = value
            End Set
        End Property

        Public Property IdUT_BT() As String
            Get
                Return _IdUT_BT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ưu tiên bản thân có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdUT_BT = value
            End Set
        End Property

        Public Property Email() As String
            Get
                Return _Email
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị e-mail có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Email = value
            End Set
        End Property

        Public Property DienThoai_NR() As String
            Get
                Return _DienThoai_NR
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại nhà riêng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai_NR = value
            End Set
        End Property

        Public Property DienThoai_DD() As String
            Get
                Return _DienThoai_DD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại di động có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai_DD = value
            End Set
        End Property

        Public Property DienThoai_CQ() As String
            Get
                Return _DienThoai_CQ
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại cơ quan có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai_CQ = value
            End Set
        End Property

        Public Property SoFax() As String
            Get
                Return _SoFax
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số fax cơ quan có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoFax = value
            End Set
        End Property

        Public Property IdTrinhDoVH() As Int32
            Get
                Return _IdTrinhDoVH
            End Get
            Set(ByVal value As Int32)
                _IdTrinhDoVH = value
            End Set
        End Property

        Public Property IdTrinhDoCT() As Int32
            Get
                Return _IdTrinhDoCT
            End Get
            Set(ByVal value As Int32)
                _IdTrinhDoCT = value
            End Set
        End Property

        Public Property SoTruong_CT() As String
            Get
                Return _SoTruong_CT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị sở trường công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoTruong_CT = value
            End Set
        End Property

        Public Property CV_Lau() As String
            Get
                Return _CV_Lau
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị công việc làm lâu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CV_Lau = value
            End Set
        End Property

        Public Property CM_Ngay() As DateTime
            Get
                Return _CM_Ngay
            End Get
            Set(ByVal value As DateTime)
                _CM_Ngay = value
            End Set
        End Property

        Public Property CM_ToChuc() As String
            Get
                Return _CM_ToChuc
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tổ chức cách mạng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CM_ToChuc = value
            End Set
        End Property

        Public Property Dang_NgayVao() As DateTime
            Get
                Return _Dang_NgayVao
            End Get
            Set(ByVal value As DateTime)
                _Dang_NgayVao = value
            End Set
        End Property

        ''' <summary>
        ''' Ngày vào đảng chính thức
        ''' </summary>
        Public Property Dang_NgayChTh() As DateTime
            Get
                Return _Dang_NgayChTh
            End Get
            Set(ByVal value As DateTime)
                _Dang_NgayChTh = value
            End Set
        End Property

        Public Property Dang_NgayRa() As DateTime
            Get
                Return _Dang_NgayRa
            End Get
            Set(ByVal value As DateTime)
                _Dang_NgayRa = value
            End Set
        End Property

        Public Property Dang_NoiKN() As String
            Get
                Return _Dang_NoiKN
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi kết nạp đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Dang_NoiKN = value
            End Set
        End Property

        Public Property Dang_SoThe() As String
            Get
                Return _Dang_SoThe
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số thẻ đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Dang_SoThe = value
            End Set
        End Property

        Public Property Dang_NGT() As String
            Get
                Return _Dang_NGT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người giới thiệu vào đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Dang_NGT = value
            End Set
        End Property

        Public Property Dang_LyDoRa() As String
            Get
                Return _Dang_LyDoRa
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do ra khỏi đảng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Dang_LyDoRa = value
            End Set
        End Property

        Public Property NH_MSKH() As String
            Get
                Return _NH_MSKH
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã khách hàng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NH_MSKH = value
            End Set
        End Property

        Public Property NH_SHTK() As String
            Get
                Return _NH_SHTK
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số tài khoản có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NH_SHTK = value
            End Set
        End Property

        Public Property NH_Ten_NH() As String
            Get
                Return _NH_Ten_NH
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị tên ngân hàng có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NH_Ten_NH = value
            End Set
        End Property

        Public Property MaSoThue() As String
            Get
                Return _MaSoThue
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã số thuế có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MaSoThue = value
            End Set
        End Property

        Public Property BHXH_SoSo() As String
            Get
                Return _BHXH_SoSo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số sổ bảo hiểm xã hội có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BHXH_SoSo = value
            End Set
        End Property

        Public Property BHXH_NgayLam() As DateTime
            Get
                Return _BHXH_NgayLam
            End Get
            Set(ByVal value As DateTime)
                _BHXH_NgayLam = value
            End Set
        End Property

        Public Property BHXH_NgayDong() As DateTime
            Get
                Return _BHXH_NgayDong
            End Get
            Set(ByVal value As DateTime)
                _BHXH_NgayDong = value
            End Set
        End Property

        Public Property BHXH_NoiLam() As String
            Get
                Return _BHXH_NoiLam
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi làm sổ hiểm xã hội có độ dài không hợp lệ!", value, value.ToString())
                End If
                _BHXH_NoiLam = value
            End Set
        End Property

        Public Property IdHocHam() As Int32
            Get
                Return _IdHocHam
            End Get
            Set(ByVal value As Int32)
                _IdHocHam = value
            End Set
        End Property

        Public Property IdHocVi() As Int32
            Get
                Return _IdHocVi
            End Get
            Set(ByVal value As Int32)
                _IdHocVi = value
            End Set
        End Property

        Public Property IdTrdChMon() As Integer
            Get
                Return _IdTrdChMon
            End Get
            Set(ByVal value As Integer)
                _IdTrdChMon = value
            End Set
        End Property

        Public Property IdTrdNgoaiNgu() As Integer
            Get
                Return _IdTrdNgoaiNgu
            End Get
            Set(ByVal value As Integer)
                _IdTrdNgoaiNgu = value
            End Set
        End Property

        Public Property IdTrdTinHoc() As Integer
            Get
                Return _IdTrdTinHoc
            End Get
            Set(ByVal value As Integer)
                _IdTrdTinHoc = value
            End Set
        End Property

        Public Property IdChuyenNganhDT() As Integer
            Get
                Return _IdChuyenNganhDT
            End Get
            Set(ByVal value As Integer)
                _IdChuyenNganhDT = value
            End Set
        End Property

        Public Property AnhThe() As Byte()
            Get
                Return _AnhThe
            End Get
            Set(ByVal value As Byte())
                _AnhThe = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

        Public Property idNew() As String
            Get
                Return _idNew
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id mới có độ dài không hợp lệ!", value, value.ToString())
                End If
                _idNew = value
            End Set
        End Property

        Public Property TrangThai() As Byte
            Get
                Return _TrangThai
            End Get
            Set(ByVal value As Byte)
                _TrangThai = value
            End Set
        End Property
        Public Property HonNhan_Cd() As String
            Get
                Return _HonNhan_Cd
            End Get
            Set(ByVal value As String)
                _HonNhan_Cd = value
            End Set
        End Property
    End Class

    Public Class HSCB_TS_HDLD
        Private _IdCBTS_HDLD As String
        Private _IdCanBo As String
        Private _SoQD As String
        Private _NgayQD As DateTime
        Private _NguoiQD As String
        Private _IdCV_Nguoi_QD As Integer
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _TuGio As String
        Private _DenGio As String
        Private _IdHT_TraLuong As Integer
        Private _HeSo As Double
        Private _IdBacLuong As Integer
        Private _TyleHuong As Double
        Private _Tien_Luong As Double
        Private _IdChuyenMon As Integer
        Private _CongViec As String
        Private _BHXH As Byte
        Private _BHYT As Byte
        Private _Loai As Byte
        Private _Status As Byte
        Private _TV_Ngay_HL As DateTime
        Private _TV_So_QD As String
        Private _TV_NgayKy_QD As DateTime
        Private _TV_IdLyDo As Integer
        Private _TV_TroCap_ThoiViec As Double
        Private _TV_TroCap_Khac As Double
        Private _TV_SoTien_BoiThuong As Double
        Private _TV_SoTien_ThuHoi As Double
        Private _GhiChu As String
        Private _ChiNhanhId As Integer
        Private _PhongBanDonViId As String

        Public Sub New()
            _IdCBTS_HDLD = ""
            _IdCanbo = ""
            _SoQD = ""
            _NgayQD = DateTime.Now
            _NguoiQD = ""
            _IdCV_Nguoi_QD = 0
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _TuGio = ""
            _DenGio = ""
            _IdHT_TraLuong = 0
            _HeSo = 0
            _IdBacLuong = 0
            _TyleHuong = 0
            _Tien_Luong = 0
            _IdChuyenMon = 0
            _CongViec = ""
            _BHXH = 0
            _BHYT = 0
            _Loai = 0
            _Status = 0
            _TV_Ngay_HL = DateTime.Now
            _TV_So_QD = ""
            _TV_NgayKy_QD = DateTime.Now
            _TV_IdLyDo = 0
            _TV_TroCap_ThoiViec = 0
            _TV_TroCap_Khac = 0
            _TV_SoTien_BoiThuong = 0
            _TV_SoTien_ThuHoi = 0
            _GhiChu = ""
            _ChiNhanhId = 0
            _PhongBanDonViId = ""
        End Sub

        Public Property IdCBTS_HDLD() As String
            Get
                Return _IdCBTS_HDLD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCBTS_HDLD = value
            End Set
        End Property

        Public Property IdCanBo() As String
            Get
                Return _IdCanbo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id cán bộ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCanbo = value
            End Set
        End Property

        Public Property SoQD() As String
            Get
                Return _SoQD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số hợp đồng có độ dài không hợp lệ!", value, value.ToString())
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
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký hợp đồng có độ dài không hợp lệ!", value, value.ToString())
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

        Public Property TuGio() As String
            Get
                Return _TuGio
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 10) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị từ giờ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TuGio = value
            End Set
        End Property

        Public Property DenGio() As String
            Get
                Return _DenGio
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 10) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị đến giờ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DenGio = value
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

        Public Property HeSo() As Double
            Get
                Return _HeSo
            End Get
            Set(ByVal value As Double)
                _HeSo = value
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

        Public Property TyleHuong() As Double
            Get
                Return _TyleHuong
            End Get
            Set(ByVal value As Double)
                _TyleHuong = value
            End Set
        End Property

        Public Property Tien_Luong() As Double
            Get
                Return _Tien_Luong
            End Get
            Set(ByVal value As Double)
                _Tien_Luong = value
            End Set
        End Property

        Public Property IdChuyenMon() As Integer
            Get
                Return _IdChuyenMon
            End Get
            Set(ByVal value As Integer)
                _IdChuyenMon = value
            End Set
        End Property

        Public Property CongViec() As String
            Get
                Return _CongViec
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị công việc có độ dài không hợp lệ!", value, value.ToString())
                End If
                _CongViec = value
            End Set
        End Property

        Public Property BHXH() As Byte
            Get
                Return _BHXH
            End Get
            Set(ByVal value As Byte)
                _BHXH = value
            End Set
        End Property

        Public Property BHYT() As Byte
            Get
                Return _BHYT
            End Get
            Set(ByVal value As Byte)
                _BHYT = value
            End Set
        End Property

        Public Property Loai() As Byte
            Get
                Return _Loai
            End Get
            Set(ByVal value As Byte)
                _Loai = value
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

        Public Property TV_Ngay_HL() As DateTime
            Get
                Return _TV_Ngay_HL
            End Get
            Set(ByVal value As DateTime)
                _TV_Ngay_HL = value
            End Set
        End Property

        Public Property TV_So_QD() As String
            Get
                Return _TV_So_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số QĐ chấm dứt HĐ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _TV_So_QD = value
            End Set
        End Property

        Public Property TV_NgayKy_QD() As DateTime
            Get
                Return _TV_NgayKy_QD
            End Get
            Set(ByVal value As DateTime)
                _TV_NgayKy_QD = value
            End Set
        End Property

        Public Property TV_IdLyDo() As Integer
            Get
                Return _TV_IdLyDo
            End Get
            Set(ByVal value As Integer)
                _TV_IdLyDo = value
            End Set
        End Property

        Public Property TV_TroCap_ThoiViec() As Double
            Get
                Return _TV_TroCap_ThoiViec
            End Get
            Set(ByVal value As Double)
                _TV_TroCap_ThoiViec = value
            End Set
        End Property

        Public Property TV_TroCap_Khac() As Double
            Get
                Return _TV_TroCap_Khac
            End Get
            Set(ByVal value As Double)
                _TV_TroCap_Khac = value
            End Set
        End Property

        Public Property TV_SoTien_BoiThuong() As Double
            Get
                Return _TV_SoTien_BoiThuong
            End Get
            Set(ByVal value As Double)
                _TV_SoTien_BoiThuong = value
            End Set
        End Property

        Public Property TV_SoTien_ThuHoi() As Double
            Get
                Return _TV_SoTien_ThuHoi
            End Get
            Set(ByVal value As Double)
                _TV_SoTien_ThuHoi = value
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
        Public Property ChiNhanhId() As Integer
            Get
                Return _ChiNhanhId
            End Get
            Set(ByVal value As Integer)
                _ChiNhanhId = value
            End Set
        End Property

        Public Property PhongBanDonViId() As String
            Get
                Return _PhongBanDonViId
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 16) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị PhongBanDonViId có độ dài không hợp lệ!", value, value.ToString())
                End If
                _PhongBanDonViId = value
            End Set
        End Property
    End Class

#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Quyết định nhân sự <---"
    Public Class QD_NhanSu
        Private _IdQDNhanSu As String
        Private _So_QD As String
        Private _NgayKy_QD As DateTime
        Private _IdCanBo As String
        Private _IdLoaiQD As Int32
        Private _NgayHL As DateTime
        Private _NgayBoNhiem_TT As DateTime
        Private _NgayThoiLuong As DateTime

        Private _NguoiKy_QD As String
        Private _IdCV_Nguoiky_QD As String

        Private _IdDonVi_Cu As Int32
        Private _IdPhong_Cu As Int32
        Private _IdChucVu_Cu As Int32
        Private _IdChuyenMon_Cu As Int32

        Private _IdDonVi_Moi As Int32
        Private _IdPhong_Moi As Int32
        Private _IdChucVu_Moi As Int32
        Private _IdChuyenMon_Moi As Int32
        Private _GhiChu As String
        Private _Active As Byte
        Private _IsKiemNhiem As Byte

        Public Sub New()
            _IdQDNhanSu = ""
            _So_QD = ""
            _NgayKy_QD = DateTime.Now
            _IdCanBo = ""
            _IdLoaiQD = 0
            _NgayHL = DateTime.Now
            _NgayBoNhiem_TT = DateTime.Now
            _NgayThoiLuong = DateTime.Now
            _NguoiKy_QD = ""
            _IdCV_Nguoiky_QD = ""

            _IdDonVi_Cu = 0
            _IdPhong_Cu = 0
            _IdChucVu_Cu = 0
            _IdChuyenMon_Cu = 0

            _IdDonVi_Moi = 0
            _IdPhong_Moi = 0
            _IdChucVu_Moi = 0
            _IdChuyenMon_Moi = 0
            _GhiChu = ""
            _Active = 0
            _IsKiemNhiem = 0
        End Sub

        Public Property IdQDNhanSu() As String
            Get
                Return _IdQDNhanSu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id quyết định nhân sự có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdQDNhanSu = value
            End Set
        End Property

        Public Property So_QD() As String
            Get
                Return _So_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định nhân sự có độ dài không hợp lệ!", value, value.ToString())
                End If
                _So_QD = value
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

        Public Property IdLoaiQD() As Int32
            Get
                Return _IdLoaiQD
            End Get
            Set(ByVal value As Int32)
                _IdLoaiQD = value
            End Set
        End Property

        Public Property NgayHL() As DateTime
            Get
                Return _NgayHL
            End Get
            Set(ByVal value As DateTime)
                _NgayHL = value
            End Set
        End Property

        Public Property NgayBoNhiem_TT() As DateTime
            Get
                Return _NgayBoNhiem_TT
            End Get
            Set(ByVal value As DateTime)
                _NgayBoNhiem_TT = value
            End Set
        End Property

        Public Property NgayThoiLuong() As DateTime
            Get
                Return _NgayThoiLuong
            End Get
            Set(ByVal value As DateTime)
                _NgayThoiLuong = value
            End Set
        End Property

        Public Property NguoiKy_QD() As String
            Get
                Return _NguoiKy_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký quyết định nhân sự có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NguoiKy_QD = value
            End Set
        End Property

        Public Property IdCV_Nguoiky_QD() As String
            Get
                Return _IdCV_Nguoiky_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chức vụ của người ký quyết định nhân sự có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdCV_Nguoiky_QD = value
            End Set
        End Property

        Public Property IdDonVi_Cu() As Int32
            Get
                Return _IdDonVi_Cu
            End Get
            Set(ByVal value As Int32)
                _IdDonVi_Cu = value
            End Set
        End Property

        Public Property IdPhong_Cu() As Int32
            Get
                Return _IdPhong_Cu
            End Get
            Set(ByVal value As Int32)
                _IdPhong_Cu = value
            End Set
        End Property

        Public Property IdChucVu_Cu() As Int32
            Get
                Return _IdChucVu_Cu
            End Get
            Set(ByVal value As Int32)
                _IdChucVu_Cu = value
            End Set
        End Property

        Public Property IdChuyenMon_Cu() As Int32
            Get
                Return _IdChuyenMon_Cu
            End Get
            Set(ByVal value As Int32)
                _IdChuyenMon_Cu = value
            End Set
        End Property

        Public Property IdDonVi_Moi() As Int32
            Get
                Return _IdDonVi_Moi
            End Get
            Set(ByVal value As Int32)
                _IdDonVi_Moi = value
            End Set
        End Property

        Public Property IdPhong_Moi() As Int32
            Get
                Return _IdPhong_Moi
            End Get
            Set(ByVal value As Int32)
                _IdPhong_Moi = value
            End Set
        End Property

        Public Property IdChucVu_Moi() As Int32
            Get
                Return _IdChucVu_Moi
            End Get
            Set(ByVal value As Int32)
                _IdChucVu_Moi = value
            End Set
        End Property

        Public Property IdChuyenMon_Moi() As Int32
            Get
                Return _IdChuyenMon_Moi
            End Get
            Set(ByVal value As Int32)
                _IdChuyenMon_Moi = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú của quyết định nhân sự có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

        Public Property Active() As Byte
            Get
                Return _Active
            End Get
            Set(ByVal value As Byte)
                _Active = value
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
    End Class
#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ công tác của cán bộ <---"
    Public Class HS_CongTac
        Private _IdHSCongTac As String
        Private _IdCanBo As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _ChucVu As String
        Private _DiaChi As String
        Private _LyDo As String
        Private _GhiChu As String
        Public Sub New()
            _IdHSCongTac = ""
            _IdCanBo = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _ChucVu = ""
            _DiaChi = ""
            _LyDo = ""
            _GhiChu = ""
        End Sub

        Public Property IdHSCongTac() As String
            Get
                Return _IdHSCongTac
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ cơ công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHSCongTac = value
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

        Public Property ChucVu() As String
            Get
                Return _ChucVu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chức vụ của hồ sơ công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ChucVu = value
            End Set
        End Property

        Public Property DiaChi() As String
            Get
                Return _DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ của hồ sơ công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DiaChi = value
            End Set
        End Property

        Public Property LyDo() As String
            Get
                Return _LyDo
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do của hồ sơ công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LyDo = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú của hồ sơ công tác có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

    End Class
#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ Xuất ngoại của cán bộ <---"
    Public Class HS_XuatNgoai
        Private _IdXuatNgoai As String
        Private _IdCanBo As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _IdNuocDen As Int32
        Private _MucDich As String
        Private _SoQD As String
        Private _NgayKy_QD As DateTime
        Private _NguoiKy_QD As String
        Private _IdCV_NguoiKy_QD As Int32
        Private _GhiChu As String

        Public Sub New()
            _IdXuatNgoai = ""
            _IdCanBo = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _IdNuocDen = 0
            _MucDich = ""
            _SoQD = ""
            _NgayKy_QD = DateTime.Now
            _NguoiKy_QD = ""
            _IdCV_NguoiKy_QD = 0
            _GhiChu = ""
        End Sub

        Public Property IdXuatNgoai() As String
            Get
                Return _IdXuatNgoai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id xuất ngoại có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdXuatNgoai = value
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

        Public Property IdNuocDen() As Int32
            Get
                Return _IdNuocDen
            End Get
            Set(ByVal value As Int32)
                _IdNuocDen = value
            End Set
        End Property

        Public Property MucDich() As String
            Get
                Return _MucDich
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mục đích có độ dài không hợp lệ!", value, value.ToString())
                End If
                _MucDich = value
            End Set
        End Property

        Public Property SoQD() As String
            Get
                Return _SoQD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 30) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định có độ dài không hợp lệ!", value, value.ToString())
                End If
                _SoQD = value
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

        Public Property NguoiKy_QD() As String
            Get
                Return _NguoiKy_QD
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NguoiKy_QD = value
            End Set
        End Property

        Public Property IdCV_NguoiKy_QD() As Int32
            Get
                Return _IdCV_NguoiKy_QD
            End Get
            Set(ByVal value As Int32)
                _IdCV_NguoiKy_QD = value
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

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ Hộ chiếu của cán bộ <---"
    Public Class HS_HoChieu
        Private _IdHoChieu As String
        Private _IdCanBo As String
        Private _So_HoChieu As String
        Private _Loai_HC As Byte
        Private _NgayCap As DateTime
        Private _NoiCap As String
        Private _NgayHH As DateTime
        Private _TinhTrang As Byte
        Private _GhiChu As String

        Public Sub New()
            _IdHoChieu = ""
            _IdCanBo = ""
            _So_HoChieu = ""
            _Loai_HC = 0
            _NgayCap = DateTime.Now
            _NoiCap = ""
            _NgayHH = DateTime.Now
            _TinhTrang = 0
            _GhiChu = ""
        End Sub

        Public Property IdHoChieu() As String
            Get
                Return _IdHoChieu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hộ chiếu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHoChieu = value
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

        Public Property So_HoChieu() As String
            Get
                Return _So_HoChieu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số hộ chiếu có độ dài không hợp lệ!", value, value.ToString())
                End If
                _So_HoChieu = value
            End Set
        End Property

        Public Property Loai_HC() As Byte
            Get
                Return _Loai_HC
            End Get
            Set(ByVal value As Byte)
                _Loai_HC = value
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
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi cấp có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiCap = value
            End Set
        End Property

        Public Property NgayHH() As DateTime
            Get
                Return _NgayHH
            End Get
            Set(ByVal value As DateTime)
                _NgayHH = value
            End Set
        End Property

        Public Property TinhTrang() As Byte
            Get
                Return _TinhTrang
            End Get
            Set(ByVal value As Byte)
                _TinhTrang = value
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

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ Lực lượng vũ trang của cán bộ <---"
    Public Class HS_LLVT
        Private _IdHSLLVT As String
        Private _IdCanBo As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _IdLoaiLLVT As Int32
        Private _IdQuanHam As Int32
        Private _ChucVu As String
        Private _DonVi As String
        Private _GhiChu As String

        Public Sub New()
            _IdHSLLVT = ""
            _IdCanBo = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _IdLoaiLLVT = 0
            _IdQuanHam = 0
            _ChucVu = ""
            _DonVi = ""
            _GhiChu = ""
        End Sub

        Public Property IdHSLLVT() As String
            Get
                Return _IdHSLLVT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id lực lượng vũ trang có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHSLLVT = value
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

        Public Property IdLoaiLLVT() As Int32
            Get
                Return _IdLoaiLLVT
            End Get
            Set(ByVal value As Int32)
                _IdLoaiLLVT = value
            End Set
        End Property

        Public Property IdQuanHam() As Int32
            Get
                Return _IdQuanHam
            End Get
            Set(ByVal value As Int32)
                _IdQuanHam = value
            End Set
        End Property

        Public Property ChucVu() As String
            Get
                Return _ChucVu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chức vụ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _ChucVu = value
            End Set
        End Property

        Public Property DonVi() As String
            Get
                Return _DonVi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Đơn vị có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DonVi = value
            End Set
        End Property

        Public Property GhiChu() As String
            Get
                Return _GhiChu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property
    End Class
#End Region

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ Cũ của cán bộ <---"
    Public Class HS_Cu
        Private _IdHSCu As String
        Private _IdCanBo As String
        Private _TuThang As DateTime
        Private _DenThang As DateTime
        Private _NgheNghiep As String
        Private _DiaChi As String
        Private _GhiChu As String

        Public Sub New()
            _IdHSCu = ""
            _IdCanBo = ""
            _TuThang = DateTime.Now
            _DenThang = DateTime.Now
            _NgheNghiep = ""
            _DiaChi = ""
            _GhiChu = ""
        End Sub

        Public Property IdHSCu() As String
            Get
                Return _IdHSCu
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ cũ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdHSCu = value
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

        Public Property TuThang() As DateTime
            Get
                Return _TuThang
            End Get
            Set(ByVal value As DateTime)
                _TuThang = value
            End Set
        End Property

        Public Property DenThang() As DateTime
            Get
                Return _DenThang
            End Get
            Set(ByVal value As DateTime)
                _DenThang = value
            End Set
        End Property

        Public Property NgheNghiep() As String
            Get
                Return _NgheNghiep
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nghề nghiệp có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NgheNghiep = value
            End Set
        End Property

        Public Property DiaChi() As String
            Get
                Return _DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DiaChi = value
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

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng - Hồ sơ Gia đình cán bộ <---"
    ''' <summary>
    ''' Class get - set - Hồ sơ Gia đình cán bộ
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HS_GDCB

        Private _IdTVien As String
        Private _IdCanBo As String
        Private _Ma_TVien As String
        Private _HoTen As String
        Private _GioiTinh As Byte
        Private _NamSinh As Integer
        Private _IdQuanHe As Integer
        Private _ConMat As Byte
        Private _NamMat As Integer
        Private _LyDo_Mat As String
        Private _IdUT_BThan As Integer
        Private _IdQueQuan As Integer
        Private _DienThoai As String
        Private _DiaChi As String
        Private _IdQuocGia As Integer
        Private _NgheNghiep As String

        Public Sub New()
            _IdTVien = ""
            _IdCanBo = ""
            _Ma_TVien = ""
            _HoTen = ""
            _GioiTinh = 0
            _NamSinh = 0
            _IdQuanHe = 0
            _ConMat = 0
            _NamMat = 0
            _LyDo_Mat = ""
            _IdUT_BThan = 0
            _IdQueQuan = 0
            _DienThoai = ""
            _DiaChi = ""
            _IdQuocGia = 0
            _NgheNghiep = ""
        End Sub

        Public Property IdTVien() As String
            Get
                Return _IdTVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdTVien = value
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

        Public Property Ma_TVien() As String
            Get
                Return _Ma_TVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị mã thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _Ma_TVien = value
            End Set
        End Property

        Public Property HoTen() As String
            Get
                Return _HoTen
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị họ tên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _HoTen = value
            End Set
        End Property

        Public Property GioiTinh() As Byte
            Get
                Return _GioiTinh
            End Get
            Set(ByVal value As Byte)
                _GioiTinh = value
            End Set
        End Property

        Public Property NamSinh() As Integer
            Get
                Return _NamSinh
            End Get
            Set(ByVal value As Integer)
                _NamSinh = value
            End Set
        End Property

        Public Property IdQuanHe() As Integer
            Get
                Return _IdQuanHe
            End Get
            Set(ByVal value As Integer)
                _IdQuanHe = value
            End Set
        End Property

        Public Property ConMat() As Byte
            Get
                Return _ConMat
            End Get
            Set(ByVal value As Byte)
                _ConMat = value
            End Set
        End Property

        Public Property NamMat() As Integer
            Get
                Return _NamMat
            End Get
            Set(ByVal value As Integer)
                _NamMat = value
            End Set
        End Property

        Public Property LyDo_Mat() As String
            Get
                Return _LyDo_Mat
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do mất có độ dài không hợp lệ!", value, value.ToString())
                End If
                _LyDo_Mat = value
            End Set
        End Property

        Public Property IdUT_BThan() As Integer
            Get
                Return _IdUT_BThan
            End Get
            Set(ByVal value As Integer)
                _IdUT_BThan = value
            End Set
        End Property

        Public Property IdQueQuan() As Integer
            Get
                Return _IdQueQuan
            End Get
            Set(ByVal value As Integer)
                _IdQueQuan = value
            End Set
        End Property

        Public Property DienThoai() As String
            Get
                Return _DienThoai
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 32) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị điện thoại có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DienThoai = value
            End Set
        End Property

        Public Property DiaChi() As String
            Get
                Return _DiaChi
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị địa chỉ có độ dài không hợp lệ!", value, value.ToString())
                End If
                _DiaChi = value
            End Set
        End Property

        Public Property IdQuocGia() As Integer
            Get
                Return _IdQuocGia
            End Get
            Set(ByVal value As Integer)
                _IdQuocGia = value
            End Set
        End Property

        Public Property NgheNghiep() As String
            Get
                Return _NgheNghiep
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nghề nghiệp có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NgheNghiep = value
            End Set
        End Property

    End Class

    ''' <summary>
    ''' Class get - set - Hồ sơ khám sức khoẻ Gia đình cán bộ
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HS_SKhGDCB

        Private _IdKhamchuaGDCB As String
        Private _Nam As Integer
        Private _Dot As Byte
        Private _IdCanBo As String
        Private _IdTVien As String
        Private _NoiDungKC As String
        Private _NoiKham As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _CanNang As Double
        Private _ChieuCao As Double
        Private _LoaiSuckhoe As Byte
        Private _KetLuan As String
        Private _GhiChu As String

        Public Sub New()
            _IdKhamchuaGDCB = ""
            _Nam = 0
            _Dot = 0
            _IdCanBo = ""
            _IdTVien = ""
            _NoiDungKC = ""
            _NoiKham = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _CanNang = 0
            _ChieuCao = 0
            _LoaiSuckhoe = 0
            _KetLuan = ""
            _GhiChu = ""
        End Sub

        Public Property IdKhamchuaGDCB() As String
            Get
                Return _IdKhamchuaGDCB
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdKhamchuaGDCB = value
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

        Public Property Dot() As Byte
            Get
                Return _Dot
            End Get
            Set(ByVal value As Byte)
                _Dot = value
            End Set
        End Property

        Public Property IdTVien() As String
            Get
                Return _IdTVien
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdTVien = value
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

        Public Property NoiDungKC() As String
            Get
                Return _NoiDungKC
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nội dung khám chữa có độ dài không hợp lệ!", value, value.ToString())
                End If
                _NoiDungKC = value
            End Set
        End Property

        Public Property NoiKham() As String
            Get
                Return _NoiKham
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nơi khám chữa có độ dài không hợp lệ!", value, value.ToString())
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

        Public Property LoaiSuckhoe() As Byte
            Get
                Return _LoaiSuckhoe
            End Get
            Set(ByVal value As Byte)
                _LoaiSuckhoe = value
            End Set
        End Property

        Public Property KetLuan() As String
            Get
                Return _KetLuan
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết luận khám chữa bệnh có độ dài không hợp lệ!", value, value.ToString())
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
                    If (value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property

    End Class

    ''' <summary>
    ''' Class get - set - Giảm trừ gia cảnh
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HS_GTGC_Xoa
        Private _IdGTGC As String
        Private _IdCanBo As String
        Private _IdQuanHe As Integer
        Private _HoTenNguoiPT As String
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _GhiChu As String

        Public Sub New()
            _IdGTGC = ""
            _IdCanBo = ""
            _IdQuanHe = 0
            _HoTenNguoiPT = ""
            _TuNgay = DateTime.Now
            _DenNgay = DateTime.Now
            _GhiChu = ""
        End Sub

        Public Property IdGTGC() As String
            Get
                Return _IdGTGC
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị xác định thành viên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _IdGTGC = value
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

        Public Property IdQuanHe() As Integer
            Get
                Return _IdQuanHe
            End Get
            Set(ByVal value As Integer)
                _IdQuanHe = value
            End Set
        End Property

        Public Property HoTenNguoiPT() As String
            Get
                Return _HoTenNguoiPT
            End Get
            Set(ByVal value As String)
                If Not (value Is Nothing) Then
                    If (value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị họ tên có độ dài không hợp lệ!", value, value.ToString())
                End If
                _HoTenNguoiPT = value
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
                    If (value.Length > 254) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", value, value.ToString())
                End If
                _GhiChu = value
            End Set
        End Property
    End Class

#End Region

#Region "---> Hàm định nghĩa lưới dữ <---"
    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu - Liên quan đến hồ sơ cán bộ
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu truyền vào</param>
    ''' <param name="state">Chỉ số xác định lưới dữ liệu cần định nghĩa
    '''                           0: Hồ sơ nhân sự
    '''                           1: Hồ sơ cán bộ tập sự
    '''                           2: Hồ sơ Công tác
    '''                           3: Hồ sơ Xuất ngoại
    '''                           4: Hồ sơ Hộ chiếu
    '''                           5: Hồ sơ Tham gia lực lượng vũ trang
    '''                           6: Hồ sơ Cũ (Lưu các thông tin về cán bộ khi chưa thoát ly)
    ''' 
    '''                           7: Chi tiết Hồ sơ Công tác
    '''                           8: Chi tiết Hồ sơ Xuất ngoại
    '''                           9: Chi tiết Hồ sơ Hộ chiếu
    '''                           10: Chi tiết Hồ sơ Tham gia lực lượng vũ trang
    '''                           11: Chi tiết Hồ sơ Cũ (Lưu các thông tin về cán bộ khi chưa thoát ly)
    '''                           12: Chi tiết Hồ sơ Học hàm
    '''                           13: Chi tiết Hồ sơ Học vị
    '''                           14: Hồ sơ Gia đình cán bộ
    '''                           15: Hồ sơ Khám sức khoẻ Gia đình cán bộ
    '''                           16: Hồ sơ Giảm trừ gia cảnh của Cán bộ
    '''                           17: Hồ sơ thu nhap gia dinh cua can bo
    '''                           18: Hồ sơ can bo ve huu - chuyen cong tac
    '''                           19: Danh sach cac quyets dinh Nhan su + Thoi viec - trong Hồ sơ cán bộ nghỉ hưu-chuyển công tác 
    '''                           20: Hồ sơ Dang - trong Hồ sơ cán bộ nghỉ hưu-chuyển công tác
    '''                           21: Ho so Dao tao Van bang chung chi - trong Hồ sơ cán bộ nghỉ hưu-chuyển công tác
    '''                           22: Ho sơ lương - trong Hồ sơ cán bộ nghỉ hưu-chuyển công tác
    ''' </param>
    ''' <remarks></remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView, ByVal state As Byte)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        Select Case state
            Case 0 ' Định nghĩa lưới dữ liệu - Hồ sơ Nhân sự
                dgv_name.Columns.Add("cln_Id", "Id") '0

                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 35
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice) '1

                dgv_name.Columns.Add("cln_STT", "STT") '2
                dgv_name.Columns.Add("cln_MaCB", "Mã CB") '3
                dgv_name.Columns.Add("cln_HoTen", "Họ tên") '4
                dgv_name.Columns.Add("cln_Username", "Tên đăng nhập") '4
                dgv_name.Columns.Add("cln_Quyen", "Quyền") '4
                dgv_name.Columns.Add("cln_TWQuanLy", "TW quản lý") '4
                dgv_name.Columns.Add("cln_GioiTinh", "Giới tính") '5
                dgv_name.Columns.Add("cln_NgaySinh", "Ngày sinh") '6
                dgv_name.Columns.Add("cln_DonVi", "Đơn vị") '7
                dgv_name.Columns.Add("cln_Phong", "Phòng/Ban") '8
                dgv_name.Columns.Add("cln_SoCMT", "Số CMT") '9
                dgv_name.Columns.Add("cln_DiaChi", "Địa chỉ - thường trú") '10
                dgv_name.Columns.Add("cln_DienThoai", "Điện thoại") '11
                dgv_name.Columns.Add("cln_Email", "Địa chỉ e-mail") '12
                dgv_name.Columns.Add("cln_BienChe", "Ngày biên chế NHCSXH") '13
                dgv_name.Columns.Add("cln_HonNhan", "Hôn nhân") '14
                dgv_name.Columns("cln_STT").Width = 30
                dgv_name.Columns("cln_MaCB").Width = 60
                dgv_name.Columns("cln_HoTen").Width = 120
                dgv_name.Columns("cln_Username").Width = 90
                dgv_name.Columns("cln_TWQuanLy").Width = 80
                dgv_name.Columns("cln_Quyen").Width = 60
                dgv_name.Columns("cln_GioiTinh").Width = 60
                dgv_name.Columns("cln_NgaySinh").Width = 77
                dgv_name.Columns("cln_DonVi").Width = 210
                'dgv_name.Columns("cln_Phong").Width = 210
                dgv_name.Columns("cln_SoCMT").Width = 90
                dgv_name.Columns("cln_DiaChi").Width = 220
                dgv_name.Columns("cln_DienThoai").Width = 90
                dgv_name.Columns("cln_Email").Width = 125
                dgv_name.Columns("cln_BienChe").Width = 125
                dgv_name.Columns("cln_HonNhan").Width = 80

                dgv_name.Columns("cln_Username").Visible = False
                dgv_name.Columns("cln_TWQuanLy").Visible = False
                dgv_name.Columns("cln_Quyen").Visible = False
                dgv_name.Columns("cln_Phong").Visible = False

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Integer = 0 To dgv_name.Columns.Count - 1
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Để column check có thể edit
                dgv_name.Columns("cln_Choice").ReadOnly = True

                'Căn chỉnh tiêu đề
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TWQuanLy").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GioiTinh").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NgaySinh").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_SoCMT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DienThoai").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_BienChe").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Id").Visible = False
                dgv_name.Columns("cln_Choice").Frozen = True

            Case 1 ' Định nghĩa lưới dữ liệu - Hồ sơ Cán bộ tập sự
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = ""
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_MaCB", "Mã cán bộ")
                dgv_name.Columns.Add("cln_HoTen", "Họ tên")
                dgv_name.Columns.Add("cln_GioiTinh", "Giới tính")
                dgv_name.Columns.Add("cln_NgaySinh", "Ngày sinh")
                dgv_name.Columns.Add("cln_PhongBan", "Phòng ban (Phòng giao dịch)")
                dgv_name.Columns.Add("cln_SoCMT", "Số CMT")
                dgv_name.Columns.Add("cln_NgayCap", "Ngày cấp")
                dgv_name.Columns.Add("cln_NoiCap", "Nơi cấp")
                dgv_name.Columns.Add("cln_DiaChi", "Địa chỉ - thường trú")
                dgv_name.Columns.Add("cln_DienThoai", "Điện thoại")
                dgv_name.Columns.Add("cln_Email", "Địa chỉ e-mail")
                dgv_name.Columns.Add("cln_LoaiHinh", "Loại hình công việc")
                dgv_name.Columns.Add("cln_HT_TraLuong", "Hình thức trả lương")
                dgv_name.Columns.Add("cln_NgayBD", "Ngày bắt đầu")
                dgv_name.Columns.Add("cln_NgayKT", "Ngày kết thúc")
                dgv_name.Columns.Add("cln_GioLV", "Giờ làm việc") '18
                dgv_name.Columns.Add("cln_HT_HuongLuong", "Hình thức hưởng lương")
                dgv_name.Columns.Add("cln_BacLuong", "Bậc lương")
                dgv_name.Columns.Add("cln_Heso", "Hệ số")
                dgv_name.Columns.Add("cln_TyLe", "Tỷ lệ hưởng (%)")
                dgv_name.Columns.Add("cln_TienLuong", "Tiền lương")
                dgv_name.Columns.Add("cln_Cv_DamNhan", "Công việc đảm nhận")
                dgv_name.Columns.Add("cln_HieuLuc", "Thông tin hiệu lực HĐLD")
                dgv_name.Columns.Add("cln_GhiChu_HieuLuc", "Thông tin hiệu lực HĐLD")
                dgv_name.Columns.Add("cln_HonNhan", "Hôn nhân")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_MaCB").Width = 75
                dgv_name.Columns("cln_HoTen").Width = 120
                dgv_name.Columns("cln_GioiTinh").Width = 60
                dgv_name.Columns("cln_NgaySinh").Width = 77
                dgv_name.Columns("cln_PhongBan").Width = 210
                dgv_name.Columns("cln_SoCMT").Width = 90
                dgv_name.Columns("cln_NgayCap").Width = 77
                dgv_name.Columns("cln_NoiCap").Width = 130
                dgv_name.Columns("cln_DiaChi").Width = 220
                dgv_name.Columns("cln_DienThoai").Width = 90
                dgv_name.Columns("cln_Email").Width = 125
                dgv_name.Columns("cln_LoaiHinh").Width = 120    'Loại hình công việc
                dgv_name.Columns("cln_HT_TraLuong").Width = 125    'Column: Hình thức trả lương

                dgv_name.Columns("cln_NgayBD").Width = 90
                dgv_name.Columns("cln_NgayKT").Width = 90
                dgv_name.Columns("cln_GioLV").Width = 85
                dgv_name.Columns("cln_HT_HuongLuong").Width = 150
                dgv_name.Columns("cln_BacLuong").Width = 75 'Column: Bậc lương
                dgv_name.Columns("cln_Heso").Width = 75 'Column: Hệ số lương

                dgv_name.Columns("cln_TyLe").Width = 105
                dgv_name.Columns("cln_TienLuong").Width = 110
                dgv_name.Columns("cln_Cv_DamNhan").Width = 200
                dgv_name.Columns("cln_HieuLuc").Width = 90
                dgv_name.Columns("cln_GhiChu_HieuLuc").Width = 150
                dgv_name.Columns("cln_HonNhan").Width = 80
                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Integer = 2 To 26
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(6).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(9).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(12).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(16).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(17).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(18).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(20).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(21).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns(22).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns(23).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_HieuLuc").Visible = False
                dgv_name.Columns("cln_Id").Visible = False
                dgv_name.Columns("cln_Choice").Frozen = True
            Case 2 ' Định nghĩa lưới dữ liệu - Hồ sơ Công tác
                'dgv_name.Columns.Add("cln_Id", "Id")
                'Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                'ckb_choice.Name = "cln_Choice"
                'ckb_choice.HeaderText = "Chọn"
                'ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'ckb_choice.Width = 45
                'Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                'dgv_name.Columns.Add("cln_STT", "STT")
                'dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                'dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                'dgv_name.Columns.Add("cln_Chucvu", "Chức vụ")
                'dgv_name.Columns.Add("cln_Diachi", "Địa chỉ nơi công tác")
                'dgv_name.Columns.Add("cln_Lydo", "Lý do nghỉ việc")
                'dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                'dgv_name.Columns(2).Width = 35
                'dgv_name.Columns(3).Width = 80
                'dgv_name.Columns(4).Width = 80
                'dgv_name.Columns(5).Width = 130
                'dgv_name.Columns(6).Width = 200
                'dgv_name.Columns(7).Width = 180
                'dgv_name.Columns(8).Width = 150

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày - Đến ngày")
                dgv_name.Columns.Add("cln_QuaTrinh", "Chức danh, đơn vị công tác")
                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 200
                dgv_name.Columns(2).Width = 800

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                'dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 0 To 2
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                'dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'dgv_name.Columns(0).Visible = False

            Case 3 ' Định nghĩa lưới dữ liệu - Hồ sơ Xuất ngoại
                dgv_name.Columns.Add("cln_Id", "Id")
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Quocgia", "Quốc gia đến")
                dgv_name.Columns.Add("cln_Mucdich", "Mục đích")
                dgv_name.Columns.Add("cln_Soqd", "Số quyết định")
                dgv_name.Columns.Add("cln_Ngayky", "Ngày ký")
                dgv_name.Columns.Add("cln_Nguoiky", "Người ký")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ - người ký")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(2).Width = 35
                dgv_name.Columns(3).Width = 75
                dgv_name.Columns(4).Width = 75
                dgv_name.Columns(5).Width = 125
                dgv_name.Columns(6).Width = 150
                dgv_name.Columns(7).Width = 100
                dgv_name.Columns(8).Width = 75
                dgv_name.Columns(9).Width = 130
                dgv_name.Columns(10).Width = 130
                dgv_name.Columns(11).Width = 130
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 11
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(8).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False

            Case 4 ' Định nghĩa lưới dữ liệu - Hồ sơ Hộ chiếu
                dgv_name.Columns.Add("cln_Id", "Id")

                'Add column Chọn xoá - bản ghi vào lưới dữ liệu
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Sohc", "Số hộ chiếu")
                dgv_name.Columns.Add("cln_Loaihc", "Loại hộ chiếu")
                dgv_name.Columns.Add("cln_Ngaycap", "Ngày cấp")
                dgv_name.Columns.Add("cln_Noicap", "Nơi cấp")
                dgv_name.Columns.Add("cln_Ngayhh", "Ngày hết hạn")
                dgv_name.Columns.Add("cln_Tinhtrang", "Tình trạng")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(2).Width = 35
                dgv_name.Columns(3).Width = 85
                dgv_name.Columns(4).Width = 115
                dgv_name.Columns(5).Width = 80
                dgv_name.Columns(6).Width = 150
                dgv_name.Columns(7).Width = 90
                dgv_name.Columns(8).Width = 105
                dgv_name.Columns(9).Width = 130
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 9
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(7).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False

            Case 5 ' Định nghĩa lưới dữ liệu - Hồ sơ Tham gia lực lượng vũ trang
                dgv_name.Columns.Add("cln_Id", "Id")
                'Add column Chọn xoá - bản ghi vào lưới dữ liệu
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Phanloai", "Phân loại")
                dgv_name.Columns.Add("cln_Quanham", "Quân hàm")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ")
                dgv_name.Columns.Add("cln_Donvi", "Đơn vị")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(2).Width = 35
                dgv_name.Columns(3).Width = 75
                dgv_name.Columns(4).Width = 75
                dgv_name.Columns(5).Width = 125
                dgv_name.Columns(6).Width = 140
                dgv_name.Columns(7).Width = 140
                dgv_name.Columns(8).Width = 160
                dgv_name.Columns(9).Width = 130
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 9
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False

            Case 6 ' Định nghĩa lưới dữ liệu - Hồ sơ Cũ (Lưu các thông tin về cán bộ khi chưa thoát ly)
                dgv_name.Columns.Add("cln_Id", "Id")
                'Add column Chọn xoá - bản ghi vào lưới dữ liệu
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tuthang", "Từ tháng")
                dgv_name.Columns.Add("cln_Denthang", "Đến tháng")
                dgv_name.Columns.Add("cln_Diachi", "Địa chỉ")
                dgv_name.Columns.Add("cln_Nghenghiep", "Nghề nghiệp")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(2).Width = 35
                dgv_name.Columns(3).Width = 75
                dgv_name.Columns(4).Width = 75
                dgv_name.Columns(5).Width = 160
                dgv_name.Columns(6).Width = 140
                dgv_name.Columns(7).Width = 140

                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 7
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(0).Visible = False

            Case 7 ' Định nghĩa lưới dữ liệu - Chi tiết Hồ sơ Công tác
                'dgv_name.Columns.Add("cln_STT", "STT")
                'dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                'dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                'dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                'dgv_name.Columns.Add("cln_Chucvu", "Chức vụ")
                'dgv_name.Columns.Add("cln_Diachi", "Địa chỉ nơi công tác")
                'dgv_name.Columns.Add("cln_Lydo", "Lý do nghỉ việc")
                'dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                'dgv_name.Columns("cln_STT").Width = 50
                'dgv_name.Columns("cln_Code").Width = 120
                'dgv_name.Columns("cln_Tungay").Width = 80
                'dgv_name.Columns("cln_Denngay").Width = 80
                'dgv_name.Columns("cln_Chucvu").Width = 150
                'dgv_name.Columns("cln_Diachi").Width = 260
                'dgv_name.Columns("cln_Lydo").Width = 220
                'dgv_name.Columns("cln_Ghichu").Width = 220
                ''ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                'For i As Byte = 0 To 7
                '    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                'Next
                ''Căn chỉnh tiêu đề
                'dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'dgv_name.Columns("cln_Code").Visible = False

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày - Đến ngày")
                dgv_name.Columns.Add("cln_QuaTrinh", "Chức danh, đơn vị công tác")
                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 200
                dgv_name.Columns(2).Width = 600
                For i As Byte = 0 To 2
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns(0).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

            Case 8 ' Định nghĩa lưới dữ liệu - Chi tiết Hồ sơ Xuất ngoại
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Quocgia", "Quốc gia đến")
                dgv_name.Columns.Add("cln_Mucdich", "Mục đích")
                dgv_name.Columns.Add("cln_Soqd", "Số quyết định")
                dgv_name.Columns.Add("cln_Ngayky", "Ngày ký")
                dgv_name.Columns.Add("cln_Nguoiky", "Người ký")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ - người ký")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 120
                dgv_name.Columns(2).Width = 75
                dgv_name.Columns(3).Width = 75
                dgv_name.Columns(4).Width = 125
                dgv_name.Columns(5).Width = 190
                dgv_name.Columns(6).Width = 100
                dgv_name.Columns(7).Width = 75
                dgv_name.Columns(8).Width = 130
                dgv_name.Columns(9).Width = 130
                dgv_name.Columns(10).Width = 170
                For i As Byte = 0 To 10
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                dgv_name.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(7).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Code").Visible = False
            Case 9 ' Định nghĩa lưới dữ liệu - Chi tiết Hồ sơ Hộ chiếu
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Sohc", "Số hộ chiếu")
                dgv_name.Columns.Add("cln_Loaihc", "Loại hộ chiếu")
                dgv_name.Columns.Add("cln_Ngaycap", "Ngày cấp")
                dgv_name.Columns.Add("cln_Noicap", "Nơi cấp")
                dgv_name.Columns.Add("cln_Ngayhh", "Ngày hết hạn")
                dgv_name.Columns.Add("cln_Tinhtrang", "Tình trạng")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 120
                dgv_name.Columns(2).Width = 85
                dgv_name.Columns(3).Width = 125
                dgv_name.Columns(4).Width = 80
                dgv_name.Columns(5).Width = 150
                dgv_name.Columns(6).Width = 90
                dgv_name.Columns(7).Width = 115
                dgv_name.Columns(8).Width = 130
                For i As Byte = 0 To 8
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(6).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Code").Visible = False
            Case 10 ' Định nghĩa lưới dữ liệu - Chi tiết Hồ sơ Tham gia lực lượng vũ trang
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Phanloai", "Phân loại")
                dgv_name.Columns.Add("cln_Quanham", "Quân hàm")
                dgv_name.Columns.Add("cln_Chucvu", "Chức vụ")
                dgv_name.Columns.Add("cln_Donvi", "Đơn vị")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 120
                dgv_name.Columns(2).Width = 75
                dgv_name.Columns(3).Width = 75
                dgv_name.Columns(4).Width = 125
                dgv_name.Columns(5).Width = 140
                dgv_name.Columns(6).Width = 170
                dgv_name.Columns(7).Width = 190
                dgv_name.Columns(8).Width = 180
                For i As Byte = 0 To 8
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                dgv_name.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Code").Visible = False
            Case 11 ' Định nghĩa lưới dữ liệu - Chi tiết Hồ sơ Cũ (Lưu các thông tin về cán bộ khi chưa thoát ly)
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Tuthang", "Từ tháng")
                dgv_name.Columns.Add("cln_Denthang", "Đến tháng")
                dgv_name.Columns.Add("cln_Diachi", "Địa chỉ")
                dgv_name.Columns.Add("cln_Nghenghiep", "Nghề nghiệp")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Tuthang").Width = 75
                dgv_name.Columns("cln_Denthang").Width = 75
                dgv_name.Columns("cln_Diachi").Width = 160
                dgv_name.Columns("cln_Nghenghiep").Width = 140
                dgv_name.Columns("cln_Ghichu").Width = 140
                For i As Byte = 0 To 6
                    dgv_name.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
                dgv_name.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Code").Visible = False
            Case 12 ' Thông tin Học Hàm của cán bộ
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Hocham", "Học hàm")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Hocham").Width = 220
                dgv_name.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns("cln_Code").Visible = False
            Case 13 ' Thông tin Học vị của cán bộ
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Code", "Mã hiệu")
                dgv_name.Columns.Add("cln_Hocvi", "Học vị")
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_Code").Width = 120
                dgv_name.Columns("cln_Hocvi").Width = 220
                dgv_name.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                dgv_name.Columns("cln_Code").Visible = False
            Case 14 ' Thông tin Hồ sơ Gia đình cán bộ
                dgv_name.Columns.Add("cln_Id", "Id")

                'Add column Chọn xoá - bản ghi vào lưới dữ liệu
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_HoTen", "Họ và tên")
                dgv_name.Columns.Add("cln_GioiTinh", "Giới tính")
                dgv_name.Columns.Add("cln_NamSinh", "Năm sinh")
                dgv_name.Columns.Add("cln_QuanHe", "Quan hệ")
                dgv_name.Columns.Add("cln_TinhTrang", "Tình trạng")
                dgv_name.Columns.Add("cln_DienThoai", "Điện thoại")
                dgv_name.Columns.Add("cln_DiaChi", "Địa chỉ")
                dgv_name.Columns.Add("cln_NgheNghiep", "Nghề nghiệp")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_HoTen").Width = 140
                dgv_name.Columns("cln_GioiTinh").Width = 63
                dgv_name.Columns("cln_NamSinh").Width = 65
                dgv_name.Columns("cln_QuanHe").Width = 71
                dgv_name.Columns("cln_TinhTrang").Width = 77
                dgv_name.Columns("cln_DienThoai").Width = 95
                dgv_name.Columns("cln_DiaChi").Width = 210
                dgv_name.Columns("cln_NgheNghiep").Width = 210
                dgv_name.Columns(0).ReadOnly = True
                For i As Byte = 2 To 9
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_STT").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_GioiTinh").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_NamSinh").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_TinhTrang").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_DienThoai").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Id").Visible = False

            Case 15 ' Thông tin Hồ sơ Khám sức khoẻ Gia đình cán bộ
                'Add column Chọn xoá - bản ghi vào lưới dữ liệu
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Id", "Id")
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_ThVien", "Thành viên")
                dgv_name.Columns.Add("cln_Namkham", "Năm khám")
                dgv_name.Columns.Add("cln_Dotkham", "Đợt khám")
                dgv_name.Columns.Add("cln_Ndkham", "Nội dung khám")
                dgv_name.Columns.Add("cln_Noikham", "Nơi khám")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày")
                dgv_name.Columns.Add("cln_Denngay", "Đến ngày")
                dgv_name.Columns.Add("cln_Cannang", "Cân nặng")
                dgv_name.Columns.Add("cln_Chieucao", "Chiều cao")
                dgv_name.Columns.Add("cln_Loaisk", "Loại sức khoẻ")
                dgv_name.Columns.Add("cln_Ketluan", "Kết luận")
                dgv_name.Columns.Add("cln_Ghichu", "Ghi chú")

                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_ThVien").Width = 140
                dgv_name.Columns("cln_Namkham").Width = 75
                dgv_name.Columns("cln_Dotkham").Width = 75
                dgv_name.Columns("cln_Ndkham").Width = 190
                dgv_name.Columns("cln_Noikham").Width = 180
                dgv_name.Columns("cln_Tungay").Width = 85
                dgv_name.Columns("cln_Denngay").Width = 85
                dgv_name.Columns("cln_Cannang").Width = 77
                dgv_name.Columns("cln_Chieucao").Width = 77
                dgv_name.Columns("cln_Loaisk").Width = 95
                dgv_name.Columns("cln_Ketluan").Width = 210
                dgv_name.Columns("cln_Ghichu").Width = 210

                For i As Byte = 1 To 14
                    dgv_name.Columns(i).ReadOnly = True
                Next

                dgv_name.Columns("cln_Tungay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Denngay").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Namkham").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Dotkham").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Cannang").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Chieucao").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
                dgv_name.Columns("cln_Id").Visible = False

            Case 16
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Id", "Id")
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_HoTen", "Họ tên người phụ thuộc")
                dgv_name.Columns.Add("cln_NgaySinh", "Ngày sinh")
                dgv_name.Columns.Add("cln_MaSoThueNPT", "Mã số thuế NPT")
                dgv_name.Columns.Add("cln_SoCMT", "Số CMT/Thẻ căn cước")
                dgv_name.Columns.Add("cln_QuanHe_HT", "Quan hệ")
                dgv_name.Columns.Add("cln_ThoiGian", "Thời gian kê khai")
                dgv_name.Columns.Add("cln_DiaChi", "Địa chỉ")
                dgv_name.Columns.Add("cln_DienThoai", "Điện thoại")
                dgv_name.Columns.Add("cln_QuocTich_HT", "Quốc tịch")
                dgv_name.Columns.Add("cln_GhiChu", "Ghi chú thêm")
                dgv_name.Columns.Add("cln_TrangThai_HT", "Trạng thái")

                dgv_name.Columns.Add("cln_IdQuocTich", "IdQuocTich")
                dgv_name.Columns.Add("cln_IdQuanHe", "IdQuanHe")
                dgv_name.Columns.Add("cln_NgayCap", "NgayCap")
                dgv_name.Columns.Add("cln_NoiCap", "NoiCap")
                dgv_name.Columns.Add("cln_TrangThai", "TrangThai")
                dgv_name.Columns.Add("cln_IdCanBo", "IdCanBo")
                
                dgv_name.Columns("cln_STT").Width = 35
                dgv_name.Columns("cln_HoTen").Width = 151
                dgv_name.Columns("cln_NgaySinh").Width = 85
                dgv_name.Columns("cln_MaSoThueNPT").Width = 85
                dgv_name.Columns("cln_SoCMT").Width = 85
                dgv_name.Columns("cln_QuanHe_HT").Width = 80
                dgv_name.Columns("cln_ThoiGian").Width = 160
                dgv_name.Columns("cln_DiaChi").Width = 220
                dgv_name.Columns("cln_DienThoai").Width = 85
                dgv_name.Columns("cln_QuocTich_HT").Width = 90
                dgv_name.Columns("cln_GhiChu").Width = 220
                dgv_name.Columns("cln_TrangThai_HT").Width = 60
                For i As Byte = 1 To 19
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_ThoiGian").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Id").Visible = False
                dgv_name.Columns("cln_IdQuocTich").Visible = False
                dgv_name.Columns("cln_IdQuanHe").Visible = False
                dgv_name.Columns("cln_NgayCap").Visible = False
                dgv_name.Columns("cln_NoiCap").Visible = False
                dgv_name.Columns("cln_TrangThai").Visible = False
                dgv_name.Columns("cln_IdCanBo").Visible = False

            Case 17
                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chọn"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 45
                Dim valCol As Integer = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_Id", "Id")
                dgv_name.Columns.Add("cln_ThoiGian", "Thời gian kê khai")
                dgv_name.Columns.Add("cln_Luong", "Thu nhập từ lương")
                dgv_name.Columns("cln_ThoiGian").Width = 100
                dgv_name.Columns("cln_Luong").Width = 400 'DataGridViewAutoSizeColumnMode.Fill
                For i As Byte = 1 To 2
                    dgv_name.Columns(i).ReadOnly = True
                Next
                dgv_name.Columns("cln_ThoiGian").CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Id").Visible = False
            Case 18 ' Định nghĩa lưới dữ liệu - Can bo ve huu, chuyen cong tac
                dgv_name.Columns.Add("cln_Id", "Id")

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_MaCB", "Mã cán bộ")
                dgv_name.Columns.Add("cln_HoTen", "Họ tên")
                dgv_name.Columns.Add("cln_GioiTinh", "Giới tính")
                dgv_name.Columns.Add("cln_NgaySinh", "Ngày sinh")
                dgv_name.Columns.Add("cln_GhiChu", "Nghỉ hưu/Chuyển công tác")
                dgv_name.Columns.Add("cln_DonVi", "Đơn vị")
                dgv_name.Columns.Add("cln_Phong", "Phòng/Ban")

                Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_choice.Name = "cln_Choice"
                ckb_choice.HeaderText = "Chuyển DL tạm"
                ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_choice.Width = 60
                Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                Dim ckb_report As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                ckb_report.Name = "cln_Report"
                ckb_report.HeaderText = "Chuyển báo cáo"
                ckb_report.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_report.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                ckb_report.Width = 60
                Dim valColBC As Int32 = dgv_name.Columns.Add(ckb_report)

                dgv_name.Columns(1).Width = 50
                dgv_name.Columns(2).Width = 75
                dgv_name.Columns(3).Width = 120
                dgv_name.Columns(4).Width = 65
                dgv_name.Columns(5).Width = 80
                dgv_name.Columns(6).Width = 250
                dgv_name.Columns(7).Width = 210
                dgv_name.Columns(8).Width = 250

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Integer = 1 To 8
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(5).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Id").Visible = False
                dgv_name.Columns("cln_Choice").Frozen = True
                dgv_name.Columns("cln_Report").Frozen = True

            Case 19  ' Định nghĩa lưới dữ liệu - Quyết định Nhân sự - trong Ho so can bo nghi huu-chuyen cong tac
                dgv_name.Columns.Add("cln_Id", "Id")

                'Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
                'ckb_choice.Name = "cln_Choice"
                'ckb_choice.HeaderText = ""
                'ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                'ckb_choice.Width = 35
                'Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)

                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_NgayHL", "Ngày hiệu lực")
                dgv_name.Columns.Add("cln_SoQD", "Số quyết định")
                dgv_name.Columns.Add("cln_DVraQD", "Đơn vị ra quyết định")
                dgv_name.Columns.Add("cln_LoaiQD", "Loại quyết định")
                dgv_name.Columns.Add("cln_NoiDung", "Nội dung")

                dgv_name.Columns(1).Width = 35
                dgv_name.Columns(2).Width = 90
                dgv_name.Columns(3).Width = 125
                dgv_name.Columns(4).Width = 240
                dgv_name.Columns(5).Width = 180
                dgv_name.Columns(6).Width = 595

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Integer = 1 To 6
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns("cln_Id").Visible = False
                'dgv_name.Columns("cln_Choice").Frozen = True
            Case 20  ' Định nghĩa lưới dữ liệu - Qua trinh sinh hoat Dang - trong Ho so can bo nghi huu-chuyen cong tac
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tungay", "Từ ngày - Đến ngày")
                dgv_name.Columns.Add("cln_KiemNhiem", "Kiêm nhiệm")
                dgv_name.Columns.Add("cln_CVDang", "Chức vụ Đảng")
                dgv_name.Columns.Add("cln_ChiBo", "Chi bộ")

                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 200
                dgv_name.Columns(2).Width = 100
                dgv_name.Columns(3).Width = 200
                dgv_name.Columns(4).Width = 400

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 0 To 4
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(0).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 21 ' Định nghĩa lưới dữ liệu - Ho so Dao tao Van bang chung chi - trong Ho so can bo nghi huu-chuyen cong tac
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_Tungay", "Thời gian đào tạo/Năm TN")
                dgv_name.Columns.Add("cln_HeDT", "Hệ đào tạo")
                dgv_name.Columns.Add("cln_TrinhDo", "Trình độ")
                dgv_name.Columns.Add("cln_ChuyenNganh", "Chuyên ngành")
                dgv_name.Columns.Add("cln_NganhHoc", "Ngành học")
                dgv_name.Columns.Add("cln_CoSoDT", "Cơ sở đào tạo")

                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 200
                dgv_name.Columns(2).Width = 100
                dgv_name.Columns(3).Width = 100
                dgv_name.Columns(4).Width = 200
                dgv_name.Columns(5).Width = 250
                dgv_name.Columns(6).Width = 380

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                For i As Byte = 0 To 6
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(0).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 22 ' Định nghĩa lưới dữ liệu - Ho so lương - trong Ho so can bo nghi huu-chuyen cong tac
                dgv_name.Columns.Add("cln_STT", "STT")
                dgv_name.Columns.Add("cln_NgayHL", "Ngày hiệu lực")
                dgv_name.Columns.Add("cln_SoQD", "Số quyết định")
                dgv_name.Columns.Add("cln_DVraQD", "Đơn vị ra quyết định")
                dgv_name.Columns.Add("cln_LoaiQD", "Loại quyết định")
                dgv_name.Columns.Add("cln_NoiDung", "Nội dung")

                dgv_name.Columns(0).Width = 35
                dgv_name.Columns(1).Width = 90
                dgv_name.Columns(2).Width = 125
                dgv_name.Columns(3).Width = 240
                dgv_name.Columns(4).Width = 180
                dgv_name.Columns(5).Width = 595

                'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
                dgv_name.Columns(0).ReadOnly = True
                For i As Integer = 1 To 5
                    dgv_name.Columns(i).ReadOnly = True
                Next
                'Căn chỉnh tiêu đề
                dgv_name.Columns(0).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
                dgv_name.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        End Select
    End Sub

    ''' <summary>
    ''' Hàm định nghĩa lưới dữ liệu Hợp đồng lao động của cán bộ tập sự
    ''' </summary>
    ''' <param name="dgv_name">Tên lưới dữ liệu</param>
    ''' <remarks></remarks>
    Public Sub Create_Frame(ByVal dgv_name As DataGridView)
        dgv_name.AutoGenerateColumns = True
        dgv_name.Columns.Clear()
        dgv_name.Columns.Add("cln_Id", "Id")
        Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
        ckb_choice.Name = "cln_Choice"
        ckb_choice.HeaderText = ""
        ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        ckb_choice.Width = 35
        Dim valCol As Int32 = dgv_name.Columns.Add(ckb_choice)
        dgv_name.Columns.Add("cln_PhBanDonVi_HT", "Phòng ban/Đơn vị")
        dgv_name.Columns.Add("cln_SoHD", "Số HĐ - QĐ")
        dgv_name.Columns.Add("cln_Loaihinh", "Loại hình công việc")
        dgv_name.Columns.Add("cln_Tungay", "Ngày bắt đầu")
        dgv_name.Columns.Add("cln_Denngay", "Ngày kết thúc")
        dgv_name.Columns.Add("cln_Cv_Damnhan", "Công việc đảm nhận")

        'Cột Hình thức Hưởng lương
        dgv_name.Columns.Add("cln_HT_TraLuong", "Hình thức trả lương")
        dgv_name.Columns.Add("cln_HT_HuongLuong", "Hình thức hưởng lương")
        dgv_name.Columns.Add("cln_Bacluong", "Bậc lương")
        dgv_name.Columns.Add("cln_Hs_Luong", "Hệ số")
        dgv_name.Columns.Add("cln_Tl_Huong", "Tỷ lệ hưởng (%)")
        dgv_name.Columns.Add("cln_Tienluong", "Tiền lương")

        'dgv_name.Columns.Add("cln_NgayKy", "Ngày ký")
        'dgv_name.Columns.Add("cln_NguoiKy", "Người ký")
        'dgv_name.Columns.Add("cln_Chucvu", "Chức vụ người ký")

        dgv_name.Columns.Add("cln_TV_Ngay_HL", "Ngày chấm dứt HĐ")
        dgv_name.Columns(2).Width = 180
        dgv_name.Columns(3).Width = 85
        dgv_name.Columns(4).Width = 120
        dgv_name.Columns(5).Width = 85  'Ngày bắt đầu
        dgv_name.Columns(6).Width = 85
        dgv_name.Columns(7).Width = 180    'Công việc đảm nhận
        dgv_name.Columns(8).Width = 120
        dgv_name.Columns(9).Width = 150
        dgv_name.Columns(10).Width = 75    'Bậc lương
        dgv_name.Columns(11).Width = 60    'Hệ số
        dgv_name.Columns(12).Width = 95    'Tỷ lệ hưởng
        dgv_name.Columns(13).Width = 110    'Tiền lương
        'dgv_name.Columns(13).Width = 78
        'dgv_name.Columns(14).Width = 110 'Người ký
        'dgv_name.Columns(15).Width = 130  'Chức vụ
        dgv_name.Columns(14).Width = 120

        'ReadOnly columns - Không cho phép edit các Cell trên lưới dữ liệu
        dgv_name.Columns(0).ReadOnly = True
        For i As Integer = 2 To 13
            dgv_name.Columns(i).ReadOnly = True
        Next
        'Căn chỉnh tiêu đề
        dgv_name.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        dgv_name.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        dgv_name.Columns(6).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        dgv_name.Columns(10).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        dgv_name.Columns(11).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        dgv_name.Columns(12).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        dgv_name.Columns(14).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

        dgv_name.Columns(0).Visible = False
    End Sub

#End Region

#Region "---> Hàm thực hiện fill dữ liệu vào treeview <---"
    ''' <summary>
    ''' Hàm trả lại giá trị - Trực thuộc
    ''' </summary>
    ''' <param name="pPosCode">Mã số của chi nhánh</param>
    ''' <remarks>Bỏ việc lấy thông tin từ AppSetting vì trong StoredProcedure đã lấy fix giá trị</remarks>
    Public Function GetTrucThuoc(ByVal pPosCode As String) As String
        Dim strReturn As String = ""
        Select Case pPosCode
            Case gMaDonViTW, "000199"
                strReturn = 1 'Globals.GetAppSetting("HSC").Trim()
            Case "000196"
                strReturn = 5 'Globals.GetAppSetting("TTCNTT").Trim()
            Case "000197"
                strReturn = 7 'Globals.GetAppSetting("TTDT").Trim()
            Case "000101"
                strReturn = 9 'Globals.GetAppSetting("SGD").Trim()
                'Chữ bổ sung Cơ sở đâò tạo
            Case "002821", "001114", "002734", "003799", "004532", "005399"
                strReturn = 6 'Cơ sở đào tạo
            Case Else
                Dim iId_Goc As Integer = SoftSqlHelper.GetNumber(String.Format("Select Id_Goc From ChiNhanh Where Ma_So = '{0}'", pPosCode), 0)
                'Dim iCap As Integer = _SqlHelper.getNumber("Select Id_Goc From ChiNhanh Where Ma_So = '" & pPosCode & "'")
                strReturn = IIf(iId_Goc = 1, 2, 4)
        End Select
        Return strReturn
    End Function

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào Cây dữ liệu - Fill root node
    ''' </summary>
    ''' <param name="tv_name"></param>
    ''' <param name="notAll"></param>
    ''' <remarks></remarks>
    Public Sub Fill_Tree(ByRef tv_name As TreeView, Optional ByVal notAll As Boolean = False)
        Dim db As DataTable
        Dim i As Integer
        Dim childNode As TreeNode
        db = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where Id_Goc In (0,1) And Ma_So='{0}' And Status = 1", DONVI.Trim()))
        If Not (db Is Nothing) Then
            If db.Rows.Count > 0 Then
                For i = 0 To db.Rows.Count - 1
                    childNode = tv_name.Nodes.Add(db.Rows(i)("Id").ToString(), db.Rows(i)("Ten_Goi").ToString())
                    childNode.Tag = "ROOT_" & db.Rows(i)("Id").ToString() & "_" & db.Rows(i)("Ma_So").ToString()
                    Fill_Node(childNode, db.Rows(i)("Id").ToString(), db.Rows(i)("Ma_So").ToString(), notAll)
                Next
            End If
        End If
        If Not notAll AndAlso tv_name.Nodes.Count > 0 Then tv_name.Nodes(0).ExpandAll()
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào Node cha là  - Hội sở chính: Gồm Chi nhánh và các phòng ban tương đương
    ''' </summary>
    ''' <param name="parent_Node"></param>
    ''' <param name="IdParentNode"></param>
    ''' <param name="CodeParentNode"></param>
    ''' <param name="notAll"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Fill_Node(ByVal parent_Node As TreeNode, ByVal IdParentNode As Integer, ByVal CodeParentNode As String, Optional ByVal notAll As Boolean = False) As TreeNode
        Dim i As Integer
        Dim db_cn, db_pb As DataTable
        Dim childNode As TreeNode
        Dim currTructhuoc As String
        'CHARINDEX(string1,string2[,start]): Hàm trả về vị trí đầu tiên tính từ vị trí start tại đó chuỗi string1 xuất hiện trong chuỗi string2.
        currTructhuoc = GetTrucThuoc(CodeParentNode)
        db_pb = _SqlHelper.SelectDBRows("Select * from PhongBan Where CharIndex('" + currTructhuoc.ToString + "',truc_thuoc) > 0 and Status = 1 Order by Ma_so")
        If Not (db_pb Is Nothing) Then
            If db_pb.Rows.Count > 0 Then
                For i = 0 To db_pb.Rows.Count - 1
                    childNode = parent_Node.Nodes.Add(db_pb.Rows(i)("id").ToString(), db_pb.Rows(i)("ten_phong").ToString())
                    childNode.Tag = "PB_" & IdParentNode & "_" & db_pb.Rows(i)("id").ToString()
                Next
            End If
        End If

        'Nếu cấp cha là HSC thì kiểm tra quyền quản lý Chi nhánh
        If Globals.Group <> "S03" And Globals.Group <> "S04" Then
            If IdParentNode = 1 Then
                db_cn = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where id_goc = {0} and Status = 1 and id in (" & Globals.QuyenQuanLyCN & ")", IdParentNode))
            Else
                db_cn = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where id_goc = {0} and Status = 1", IdParentNode))
            End If
            If Not (db_cn Is Nothing) Then
                If db_cn.Rows.Count > 0 Then
                    For i = 0 To db_cn.Rows.Count - 1
                        childNode = parent_Node.Nodes.Add(db_cn.Rows(i)("id").ToString(), db_cn.Rows(i)("Ten_Goi").ToString())
                        childNode.Tag = "DV_" & db_cn.Rows(i)("id").ToString() & "_" & db_cn.Rows(i)("ma_so").ToString()
                    Next
                End If
            End If
        End If

        Return parent_Node
    End Function

    Public Sub Fill_Tree_NghiHuu_ChuyenCT(ByVal tv_name As TreeView, Optional ByVal notAll As Boolean = False)
        Dim db As DataTable
        Dim i As Integer
        Dim childNode As TreeNode
        db = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where ma_so='{0}' And Status = 1", DONVI.Trim()))
        If Not (db Is Nothing) Then
            If db.Rows.Count > 0 Then
                For i = 0 To db.Rows.Count - 1
                    childNode = tv_name.Nodes.Add(db.Rows(i)("id").ToString(), db.Rows(i)("Ten_Goi").ToString())
                    childNode.Tag = "ROOT_" & db.Rows(i)("id").ToString() & "_" & db.Rows(i)("ma_so").ToString()
                    Fill_Node_NghiHuu_ChuyenCT(childNode, db.Rows(i)("id").ToString(), db.Rows(i)("ma_so").ToString(), notAll)
                Next
            End If
        End If
    End Sub

    Public Function Fill_Node_NghiHuu_ChuyenCT(ByVal parent_Node As TreeNode, ByVal IdParentNode As Integer, ByVal CodeParentNode As String, Optional ByVal notAll As Boolean = False) As TreeNode
        Dim i As Integer
        Dim db_cn As DataTable
        Dim childNode As TreeNode
        'Dim db_pb As DataTable
        'Dim currTructhuoc As String
        ''CHARINDEX(string1,string2[,start]): Hàm trả về vị trí đầu tiên tính từ vị trí start tại đó chuỗi string1 xuất hiện trong chuỗi string2.
        'currTructhuoc = GetTrucThuoc(CodeParentNode)
        'db_pb = _SqlHelper.SelectDBRows("Select * from PhongBan Where CharIndex('" + currTructhuoc.ToString + "',truc_thuoc) > 0 and Status = 1 Order by Ma_so")
        'If Not (db_pb Is Nothing) Then
        '    If db_pb.Rows.Count > 0 Then
        '        For i = 0 To db_pb.Rows.Count - 1
        '            childNode = parent_Node.Nodes.Add(db_pb.Rows(i)("id").ToString(), db_pb.Rows(i)("ten_phong").ToString())
        '            childNode.Tag = "PB_" & IdParentNode & "_" & db_pb.Rows(i)("id").ToString()
        '        Next
        '    End If
        'End If

        db_cn = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where id_goc = {0} and Status = 1", IdParentNode))
        If Not (db_cn Is Nothing) Then
            If db_cn.Rows.Count > 0 Then
                For i = 0 To db_cn.Rows.Count - 1
                    childNode = parent_Node.Nodes.Add(db_cn.Rows(i)("id").ToString(), db_cn.Rows(i)("Ten_Goi").ToString())
                    childNode.Tag = "DV_" & db_cn.Rows(i)("id").ToString() & "_" & db_cn.Rows(i)("ma_so").ToString()
                Next
            End If
        End If
        Return parent_Node
    End Function

    ''' <summary>
    ''' Hàm fill danh sách cán bộ - Thuộc phòng ban ra theo Parent node -> Đơn vị
    ''' </summary>
    ''' <param name="parentNode">Node cha cần fill các node con</param>
    ''' <param name="Id_Dv">Id đơn vị</param>
    ''' <param name="Id_Pb">Id phòng ban thuộc đơn vị</param>
    ''' <param name="FlagNode">True: Fill tất cả danh sách cán bộ trừ các cán bộ đã thôi việc. False: Fill tất cả danh sách cán bộ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Fill_NodeCanbo(ByVal parentNode As TreeNode, ByVal Id_Dv As Integer, ByVal Id_Pb As Integer, ByVal FlagNode As Boolean, Optional ByVal isNghiHuu_ChuyenCongTac As Boolean = False) As TreeNode
        Dim i As Integer
        Dim childNode As TreeNode
        Dim db As DataTable
        Dim strSQL As String = ""

        If isNghiHuu_ChuyenCongTac Then
            strSQL = getSQL_DSCB(Id_Dv, Id_Pb, False, False, False, False, False, False, False, True)
        Else
            If (FlagNode = True) Then   ' Fill tất cả trừ các cán bộ đã thôi việc
                'strSQL = "SELECT t2.IdCanbo, MaCB, HoTen, t2.IdNew FROM QDNhansu t1, HS_Canbo t2 WHERE " & _
                '                       " t1.idCanbo = t2.idCanbo" & _
                '                       " And t2.idCanbo not in (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS=1) " & _
                '                       " And t1.idDonvi_Moi=" & Id_Dv & " AND t1.idPhong_Moi=" & Id_Pb & _
                '                       " And t1.idcanbo not in (SELECT idcanbo FROM QDNhanSu t3 WHERE t3.idCanbo=t1.idCanbo and datediff(second,t1.ngayHL,t3.ngayHL)>0) " & _
                '                       " And t1.IsKiemNhiem = 0 AND t1.IsQD_NHCS=1 order by IdChucvu_moi "
                strSQL = getSQL_DSCB(Id_Dv, Id_Pb)
            Else                        ' Fill tất cả các cán bộ
                strSQL = "Select * From QDNhanSu a, HS_Canbo b Where a.IdCanBo = b.IdCanBo" &
                                   " And a.IdDonvi_Moi=" & Id_Dv & " And a.IdPhong_Moi=" & Id_Pb &
                                   " And a.Idcanbo Not In (Select IdCanBo From QDNhanSu c Where c.IdCanBo = a.IdCanBo And DatedIff(day, a.NgayHL, c.NgayHL) > 0) " &
                                   " And a.IsKiemNhiem = 0 AND a.IsQD_NHCS=1 Order by IdChucvu_moi "
            End If
        End If

        db = _SqlHelper.SelectDBRows(strSQL)
        If Not (db Is Nothing) Then
            If db.Rows.Count > 0 Then
                For i = 0 To db.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(db.Rows(i)("IdCanbo").ToString(), db.Rows(i)("hoten").ToString())
                    childNode.Tag = "CB_" & db.Rows(i)("IdCanbo").ToString()
                    If (db.Rows(i)("IdNew").ToString().Trim() <> "") Then
                        childNode.BackColor = Color.Red
                        childNode.ToolTipText = "Lưu ý: Cán bộ " + childNode.Text.Trim() + " đã được chuyển sang hồ sơ lao động ngắn hạn!"
                    End If
                Next
            End If
        End If
        Return parentNode
    End Function

    ''' <summary>
    ''' Hàm thực hiện fill data vào cây dữ liệu - Hồ sơ cán bộ tập sự 
    ''' </summary>
    ''' <param name="tv_name"></param>
    ''' <param name="notAll"></param>
    ''' <remarks></remarks>
    Public Sub Fill_Tree_HSCBTS(ByVal tv_name As TreeView, Optional ByVal notAll As Boolean = False)
        Dim db As DataTable
        Dim dbCN As DataTable
        Dim i, j As Integer
        Dim childNode, Node_ As TreeNode

        If DONVI = gMaDonViTW Then
            'db = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where Status = 1 and id_goc IN (0,1)"))
            db = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where Status = 1 and id_goc <=1 and id in (" & Globals.QuyenQuanLyCN & ")"))
            If Not (db Is Nothing) Then
                If db.Rows.Count > 0 Then
                    For i = 0 To db.Rows.Count - 1
                        childNode = tv_name.Nodes.Add(db.Rows(i)("id").ToString(), db.Rows(i)("Ten_Goi").ToString())
                        childNode.Tag = "00CN" & db.Rows(i)("id").ToString()

                        'Không đọc lại cấp con của Hội sở chính
                        If db.Rows(i)("id") > 1 Then
                            dbCN = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where Status = 1 and id_goc =" & db.Rows(i)("id")))
                            If Not (dbCN Is Nothing) Then
                                If dbCN.Rows.Count > 0 Then
                                    For j = 0 To dbCN.Rows.Count - 1
                                        Node_ = childNode.Nodes.Add(dbCN.Rows(j)("id").ToString(), dbCN.Rows(j)("Ten_Goi").ToString())
                                        Node_.Tag = "00CN" & dbCN.Rows(j)("id").ToString()
                                    Next
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Else
            db = _SqlHelper.SelectDBRows(String.Format("SELECT * FROM CHINHANH WHERE ma_so='" & DONVI & "' AND status=1 and id_goc = 1"))
            If Not (db Is Nothing) Then
                If db.Rows.Count > 0 Then
                    For i = 0 To db.Rows.Count - 1
                        childNode = tv_name.Nodes.Add(db.Rows(i)("id").ToString(), db.Rows(i)("Ten_Goi").ToString())
                        childNode.Tag = "00CN" & db.Rows(i)("id").ToString()
                        dbCN = _SqlHelper.SelectDBRows(String.Format("Select * from ChiNhanh Where Status = 1 and id_goc =" & db.Rows(i)("id")))
                        If Not (dbCN Is Nothing) Then
                            If dbCN.Rows.Count > 0 Then
                                For j = 0 To dbCN.Rows.Count - 1
                                    Node_ = childNode.Nodes.Add(dbCN.Rows(j)("id").ToString(), dbCN.Rows(j)("Ten_Goi").ToString())
                                    Node_.Tag = "00CN" & dbCN.Rows(j)("id").ToString()
                                Next
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill các node con vào Node cha của treeview hồ sơ cán bộ tập sự
    ''' </summary>
    ''' <param name="tn_parent">Node cha cần add thêm các node con</param>
    ''' <param name="IdChiNhanh">Id chi nhánh</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Fill_Node(ByVal tn_parent As TreeNode, ByVal IdChiNhanh As Int32) As TreeNode
        Dim i As Integer
        Dim tn_child As TreeNode
        Dim strSQL As String = ""
        'strSQL = String.Format("Select * From HSCB_TS Where IdChiNhanh ={0}", IdChiNhanh)
        strSQL = "Select Id,MaCB,HoTen,IdChiNhanh,TenThuongGoi,BiDanh,GioiTinh,NgaySinh,IdPhongBan,IdNew From HSCB_TS Where (IdChiNhanh =" & IdChiNhanh & " AND Substring(IdPhongBan, 1, 2)='PB') Or (Substring(IdPhongBan, CharIndex('_',IdPhongBan)+1, len(IdPhongBan)-CharIndex('_',IdPhongBan))=" & IdChiNhanh & " AND Substring(IdPhongBan, 1, 2)='DV')"
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If db.Rows.Count > 0 Then
                    For i = 0 To db.Rows.Count - 1
                        tn_child = tn_parent.Nodes.Add(db.Rows(i)("HoTen").ToString().Trim())
                        tn_child.Tag = "01CN" + db.Rows(i)("Id").ToString().Trim()
                        If (db.Rows(i)("IdNew").ToString().Trim() <> "") Then
                            tn_child.BackColor = Color.Red
                            tn_child.ToolTipText = "Lưu ý: Cán bộ tập sự " + tn_child.Text.Trim() + " đã được chuyển sang hồ sơ cán bộ chính thức!"
                        End If
                    Next
                End If
            End If
        End Using
        Return tn_parent
    End Function

#End Region

#Region "---> Hàm thực hiện lấy danh sách bản ghi nhân sự <---"
    ''' <summary>
    ''' Hàm thực hiện tìm kiếm danh sách cán bộ theo điều kiện cần tìm
    ''' </summary>
    ''' <param name="pThoiDiem">Thời điểm lấy danh sách, nếu trống lấy theo ngày hiện tại. Chuoi dinh dạng dd/MM/yyyy</param>
    ''' <param name="pIdCanBo">IdCanBo cần tìm</param>
    ''' <param name="pMaCB">Mã cán bộ</param>
    ''' <param name="pHoTen">Họ tên cán bộ</param>
    ''' <param name="pDonVi_Id">Chỉ số xác định đơn vị công tác</param>
    ''' <param name="pDonVi_Cd">Mã hiệu xác định đơn vị công tác</param>
    ''' <param name="pDonVi_HT">Tên gọi xác định đơn vị công tác</param>
    ''' <param name="pPhongBan_Id">Chỉ số xác định phòng/ban</param>
    ''' <param name="pPhongBan_Cd">Chuỗi các Mã hiệu xác định phòng/ban. Ex: '6015,6016'</param>
    ''' <param name="pPhongBan_HT">Tên phòng ban</param> 
    ''' <param name="pChucVu_Id">Chỉ số xác định Chức vụ</param>
    ''' <param name="pChucVu_Cd">Chuỗi các Mã hiệu xác định Chức vụ</param>
    ''' <param name="pChucVu_HT">Tên chức vụ</param>
    ''' <param name="pNgaySinh_BD">Ngày sinh bắt đầu. Định dạng dd/MM/yyyy</param>
    ''' <param name="pNgaySinh_KT">Ngày sinh kết thúc. Định dạng dd/MM/yyyy</param>
    ''' <param name="pGioiTinh_Cd">Mã giới tính. 'M' - Nam; 'F' - Nữ</param>
    ''' <param name="pGioiTinh_HT">Chuỗi giới tính cần tìm</param>
    ''' <param name="pHonNhan_Cd">Chuỗi giá trị xác định tình trạng hôn nhân</param>
    ''' <param name="pSoCMT">Số CMTND/Thẻ căn cước</param>
    ''' <param name="pDanToc_Id">List Id danh mục dân tộc</param>
    ''' <param name="pTonGiao_Id">List Id danh mục tôn giáo</param>
    ''' <param name="pTWQuanLy">Gía trị: 1 - TW Quản lý; 0 - Đơn vị quản lý</param>
    ''' <param name="pNghiHuu">Gía trị: 1 - Lấy nghỉ hưu; 0 - Không xét</param>
    ''' <param name="pNgayTinhNghiHuu">Ngày tính tuổi nghỉ hưu (Chỉ lấy năm)</param>
    ''' <param name="pNgayTinhNghiHuuDenNam">Ngày tính tuổi nghỉ hưu (Chỉ lấy năm)</param>
    ''' <param name="pSoBHXH">0: Tất cả; 1: Chưa làm sổ BHXH</param>
    ''' <param name="pNgayVaoNHCS_BD">Ngày vào NHCSXH bắt đầu. Truyen vao chuoi dinh dạng dd/MM/yyyy</param>
    ''' <param name="pNgayVaoNHCS_KT">Ngày vào NHCSXH kết thúc. Truyen vao chuoi dinh dạng dd/MM/yyyy</param>
    ''' <param name="pNgayBoNhiemLai_BD">Ngày bổ nhiệm lại bắt đầu. Truyen vao chuoi dinh dạng dd/MM/yyyy</param>
    ''' <param name="pNgayBoNhiemLai_KT">Ngày bổ nhiệm lại kết thúc. Truyen vao chuoi dinh dạng dd/MM/yyyy</param>
    ''' <param name="pIsALL">pIsALL = 1 => Lấy tất cả đơn vị; 0 - Lấy duy nhất 1 đơn vị</param>
    ''' <param name="pFlagDL">pFlagDL - Nhận giá trị: 0 - Lấy một số Cột DL chính; 1 - Lấy đầy đủ DL</param>
    ''' <returns>DataTable cán bộ</returns>
    ''' <remarks></remarks>
    Public Function GetHS_CanBo_Search(ByVal pThoiDiem As String, ByVal pIdCanBo As String, pMaCB As String, pHoTen As String, pDonVi_Id As Int32, pDonVi_Cd As String, pDonVi_HT As String, pPhongBan_Id As Int32, pPhongBan_Cd As String, pPhongBan_HT As String, pChucVu_Id As Int32, pChucVu_Cd As String, pChucVu_HT As String, pNgaySinh_BD As String, pNgaySinh_KT As String, pGioiTinh_Cd As String, pGioiTinh_HT As String, pHonNhan_Cd As String, pSoCMT As String, pDanToc_Id As String, pTonGiao_Id As String, pTWQuanLy As Byte, pNghiHuu As Byte, pNgayTinhNghiHuu As String, pNgayTinhNghiHuuDenNam As String, pSoBHXH As Int32, pNgayVaoNHCS_BD As String, pNgayVaoNHCS_KT As String, pNgayBoNhiemLai_BD As String, pNgayBoNhiemLai_KT As String, pIsALL As Byte, pFlagDL As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = New SqlCommand("GetHS_CanBo_Search", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@pThoiDiem", SqlDbType.VarChar, 10))
            command.Parameters("@pThoiDiem").Value = pThoiDiem

            command.Parameters.Add(New SqlParameter("@pIdCanBo", SqlDbType.NVarChar, 64))
            command.Parameters("@pIdCanBo").Value = pIdCanBo

            command.Parameters.Add(New SqlParameter("@pMaCB", SqlDbType.NVarChar, 64))
            command.Parameters("@pMaCB").Value = pMaCB
            command.Parameters.Add(New SqlParameter("@pHoTen", SqlDbType.NVarChar, 256))
            command.Parameters("@pHoTen").Value = pHoTen

            command.Parameters.Add(New SqlParameter("@pDonVi_Id", SqlDbType.Int))
            command.Parameters("@pDonVi_Id").Value = pDonVi_Id
            command.Parameters.Add(New SqlParameter("@pDonVi_Cd", SqlDbType.VarChar, 8))
            command.Parameters("@pDonVi_Cd").Value = pDonVi_Cd
            command.Parameters.Add(New SqlParameter("@pDonVi_HT", SqlDbType.NVarChar, 128))
            command.Parameters("@pDonVi_HT").Value = pDonVi_HT

            command.Parameters.Add(New SqlParameter("@pPhongBan_Id", SqlDbType.Int))
            command.Parameters("@pPhongBan_Id").Value = pPhongBan_Id
            command.Parameters.Add(New SqlParameter("@pPhongBan_Cd", SqlDbType.VarChar, 512))
            command.Parameters("@pPhongBan_Cd").Value = pPhongBan_Cd
            command.Parameters.Add(New SqlParameter("@pPhongBan_HT", SqlDbType.NVarChar, 128))
            command.Parameters("@pPhongBan_HT").Value = pPhongBan_HT

            command.Parameters.Add(New SqlParameter("@pChucVu_Id", SqlDbType.Int))
            command.Parameters("@pChucVu_Id").Value = pChucVu_Id
            command.Parameters.Add(New SqlParameter("@pChucVu_Cd", SqlDbType.VarChar, 512))
            command.Parameters("@pChucVu_Cd").Value = pChucVu_Cd
            command.Parameters.Add(New SqlParameter("@pChucVu_HT", SqlDbType.NVarChar, 128))
            command.Parameters("@pChucVu_HT").Value = pChucVu_HT

            command.Parameters.Add(New SqlParameter("@pNgaySinh_BD", SqlDbType.VarChar, 10))
            command.Parameters("@pNgaySinh_BD").Value = pNgaySinh_BD
            command.Parameters.Add(New SqlParameter("@pNgaySinh_KT", SqlDbType.VarChar, 10))
            command.Parameters("@pNgaySinh_KT").Value = pNgaySinh_KT

            command.Parameters.Add(New SqlParameter("@pGioiTinh_Cd", SqlDbType.VarChar, 1))
            command.Parameters("@pGioiTinh_Cd").Value = pGioiTinh_Cd
            command.Parameters.Add(New SqlParameter("@pGioiTinh_HT", SqlDbType.NVarChar, 16))
            command.Parameters("@pGioiTinh_HT").Value = pGioiTinh_HT

            command.Parameters.Add(New SqlParameter("@pHonNhan_Cd", SqlDbType.NVarChar, 512))
            command.Parameters("@pHonNhan_Cd").Value = pHonNhan_Cd
            command.Parameters.Add(New SqlParameter("@pSoCMT", SqlDbType.NVarChar, 16))
            command.Parameters("@pSoCMT").Value = pSoCMT

            command.Parameters.Add(New SqlParameter("@pDanToc_Id", SqlDbType.NVarChar, 512))
            command.Parameters("@pDanToc_Id").Value = pDanToc_Id
            command.Parameters.Add(New SqlParameter("@pTonGiao_Id", SqlDbType.NVarChar, 512))
            command.Parameters("@pTonGiao_Id").Value = pTonGiao_Id

            command.Parameters.Add(New SqlParameter("@pTWQuanLy", SqlDbType.TinyInt))
            command.Parameters("@pTWQuanLy").Value = pTWQuanLy

            command.Parameters.Add(New SqlParameter("@pNghiHuu", SqlDbType.TinyInt))
            command.Parameters("@pNghiHuu").Value = pNghiHuu

            command.Parameters.Add(New SqlParameter("@pNgayTinhNghiHuu", SqlDbType.VarChar, 10))
            command.Parameters("@pNgayTinhNghiHuu").Value = pNgayTinhNghiHuu

            command.Parameters.Add(New SqlParameter("@pNgayTinhNghiHuuDenNam", SqlDbType.VarChar, 10))
            command.Parameters("@pNgayTinhNghiHuuDenNam").Value = pNgayTinhNghiHuuDenNam

            command.Parameters.Add(New SqlParameter("@pSoBHXH", SqlDbType.Int))
            command.Parameters("@pSoBHXH").Value = pSoBHXH

            command.Parameters.Add(New SqlParameter("@pNgayVaoNHCS_BD", SqlDbType.VarChar, 10))
            command.Parameters("@pNgayVaoNHCS_BD").Value = pNgayVaoNHCS_BD
            command.Parameters.Add(New SqlParameter("@pNgayVaoNHCS_KT", SqlDbType.VarChar, 10))
            command.Parameters("@pNgayVaoNHCS_KT").Value = pNgayVaoNHCS_KT

            command.Parameters.Add(New SqlParameter("@pNgayBoNhiemLai_BD", SqlDbType.VarChar, 10))
            command.Parameters("@pNgayBoNhiemLai_BD").Value = pNgayBoNhiemLai_BD
            command.Parameters.Add(New SqlParameter("@pNgayBoNhiemLai_KT", SqlDbType.VarChar, 10))
            command.Parameters("@pNgayBoNhiemLai_KT").Value = pNgayBoNhiemLai_KT

            command.Parameters.Add(New SqlParameter("@pIsALL", SqlDbType.TinyInt))
            command.Parameters("@pIsALL").Value = pIsALL
            command.Parameters.Add(New SqlParameter("@pFlagDL", SqlDbType.TinyInt))
            command.Parameters("@pFlagDL").Value = pFlagDL
            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_CanBo")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_CanBo")
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
    ''' Hàm trả về bản ghi Hồ sơ nhân sự thoả mãn mã số nhân sự truyền vào
    ''' </summary>
    ''' <param name="_Code">Id nhân sự</param>
    ''' <returns>Row human document</returns>
    ''' <remarks></remarks>
    Public Function GetHuman_ForCode(ByVal _Code As String) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try

            Dim command As SqlCommand = New SqlCommand("HS_CanBo_GetForIdCanBo", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
            command.Parameters("@_IdCanBo").Value = _Code

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_CanBo")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_CanBo").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về danh sách Hồ sơ nhân sự theo phòng ban và đơn vị truyền vào
    ''' </summary>
    ''' <param name="_deptId">Id phòng ban truyền vào</param>
    ''' <param name="_unitId">Id đơn vị truyền vào</param>
    ''' <returns>Danh sách nhân sự thoả mãn</returns>
    ''' <remarks></remarks>
    Public Function GetAllHuman(ByVal _deptId As Int32, ByVal _unitId As Int32) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = New SqlCommand("HS_CanBo_GetAll", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdDonVi", SqlDbType.Int))
            command.Parameters("@_IdDonVi").Value = _unitId

            command.Parameters.Add(New SqlParameter("@_IdPhongBan", SqlDbType.Int))
            command.Parameters("@_IdPhongBan").Value = _deptId

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_CanBo")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_CanBo")
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
    ''' Hàm trả về tên gọi khi biết Id - truy xuất theo câu lệnh SQL truyền vào
    ''' </summary>
    ''' <param name="strSQL">Câu lệnh truy xuất dữ liệu</param>
    ''' <returns>Tên gọi cần lấy</returns>
    ''' <remarks></remarks>
    Public Function GetNameByCode(ByVal strSQL As String) As String
        Dim strName = ""
        Dim db As DataTable = Nothing
        db = _SqlHelper.SelectDBRows(strSQL)
        If Not (db Is Nothing) Then
            If db.Rows.Count > 0 Then
                strName = db.Rows(0)(1).ToString().Trim()
            End If
        End If
        Return strName
    End Function

    ''' <summary>
    ''' Hàm trả về Địa chỉ chi tiết khi biết: Chỉ số tỉnh - Huyện và địa chỉ
    ''' </summary>
    ''' <param name="pTinhId">Chi số xác định tỉnh</param>
    ''' <param name="pXaId">Chỉ số xác định xã</param>
    ''' <param name="pThonId">Chỉ số xác định thôn</param>
    ''' <param name="pAddress">Địa chỉ chi tiết</param>
    ''' <returns>Địa chỉ chi tiết</returns>
    ''' <remarks></remarks>
    Public Function GetAddress(ByVal pTinhId As Int32, ByVal pXaId As Int32, ByVal pThonId As Int32, ByVal pAddress As String) As String
        Dim strResult As String = ""
        If (pAddress <> "") Then
            strResult = pAddress
            If (pThonId > 0) Then
                strResult += " - " + GetNameByCode(String.Format("Select Id, Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa<>'00' And Ma_Thon<>'00' And Id = {0} Order By TrangThai ", pThonId)).Trim()
            End If

            If (pXaId > 0) Then
                strResult += " - " + GetNameByCode(String.Format("Select Id, Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa<>'00' And Ma_Thon='00' And Id = {0} Order By TrangThai ", pXaId)).Trim()
            End If
            If (pTinhId > 0) Then
                strResult += " - " + GetNameByCode(String.Format("Select Id, Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Id = {0} Order By TrangThai ", pTinhId)).Trim()
            End If

            'If (HuyenId > 0) Then
            '    strResult += " - " + GetNameByCode(String.Format("Select id, Ten_Goi from DiaDanh Where Id = {0} And Id_Goc != 0 And Status = 1", HuyenId)).Trim().Replace("Huyện", "H.").Replace("Quận", "Q.")
            '    If (TinhId > 0) Then
            '        strResult += " - " + GetNameByCode(String.Format("Select id, Ten_Goi from DiaDanh Where id = {0} and id_goc = 0 and Status = 1", TinhId)).Trim()
            '    End If
            'End If
        End If
        'Thực hiện chuẩn hoá lại địa chỉ chi tiết nếu Address truyền vào rỗng
        While strResult.EndsWith("-") 'Bỏ dấu - ở cuối chuỗi nếu có
            strResult = strResult.Substring(0, strResult.Length - 1)
        End While
        'Chuẩn hoá ký tự đầu tiên của chuỗi
        If (strResult <> "") Then
            If (strResult.Substring(0, 1) = "-") Then ' Nếu đầu chuỗi chứa ký tự - thì remove khỏi chuỗi
                strResult = strResult.Substring(1)
            End If
        End If
        strResult = strResult.Replace("Thành phố", "TP.")
        Return strResult
    End Function

    ''' <summary>
    ''' Hàm trả về danh sách các bản ghi - Hồ sơ cán bộ tập sự theo đơn vị
    ''' </summary>
    ''' <param name="_BranchId">Id đơn vị - Chi nhánh</param>
    ''' <returns>Danh sách bản ghi thoả mãn</returns>
    ''' <remarks></remarks>
    Public Function GetAll_Trainee_New(ByVal _BranchId As Int32, ByVal _Id As String, ByVal _HoTen As String, ByVal _DonVi_Cd As String, ByVal _IsAll As Byte, ByVal _FlagAnhThe As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = New SqlCommand("HSCB_TS_GetForIdChiNhanh", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdChiNhanh", SqlDbType.Int))
            command.Parameters("@_IdChiNhanh").Value = _BranchId
            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.NVarChar))
            command.Parameters("@_IdCanBo").Value = _Id
            command.Parameters.Add(New SqlParameter("@_HoTen", SqlDbType.NVarChar))
            command.Parameters("@_HoTen").Value = _HoTen
            command.Parameters.Add(New SqlParameter("@_DonVi_Cd", SqlDbType.NVarChar))
            command.Parameters("@_DonVi_Cd").Value = _DonVi_Cd
            command.Parameters.Add(New SqlParameter("@_IsAll", SqlDbType.TinyInt))
            command.Parameters("@_IsAll").Value = _IsAll
            command.Parameters.Add(New SqlParameter("@_FlagAnhThe", SqlDbType.TinyInt))
            command.Parameters("@_FlagAnhThe").Value = _FlagAnhThe
            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_CanBoTS")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_CanBoTS")
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
    ''' Hàm trả về bản ghi - Hồ sơ cán bộ tập sự theo Id cán bộ tập sự truyền vào
    ''' </summary>
    ''' <param name="_Code">Id truyền vào</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTrainee(ByVal _BranchId As Int32, ByVal _Code As String, ByVal _HoTen As String, ByVal _DonVi_Cd As String, ByVal _IsAll As Byte, ByVal _FlagAnhThe As Byte) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try

            Dim command As SqlCommand = New SqlCommand("HSCB_TS_GetForIdChiNhanh", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdChiNhanh", SqlDbType.Int))
            command.Parameters("@_IdChiNhanh").Value = _BranchId
            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.NVarChar))
            command.Parameters("@_IdCanBo").Value = _Code
            command.Parameters.Add(New SqlParameter("@_HoTen", SqlDbType.NVarChar))
            command.Parameters("@_HoTen").Value = _HoTen
            command.Parameters.Add(New SqlParameter("@_DonVi_Cd", SqlDbType.NVarChar))
            command.Parameters("@_DonVi_Cd").Value = _DonVi_Cd
            command.Parameters.Add(New SqlParameter("@_IsAll", SqlDbType.TinyInt))
            command.Parameters("@_IsAll").Value = _IsAll
            command.Parameters.Add(New SqlParameter("@_FlagAnhThe", SqlDbType.TinyInt))
            command.Parameters("@_FlagAnhThe").Value = _FlagAnhThe

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HSCB_TS")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HSCB_TS").Rows(0)
            End Using
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try
    End Function

    ''' <summary>
    ''' Hàm trả về tên loại hình công việc - từ id loại hình
    ''' </summary>
    ''' <param name="LoaiId">Id loại hình công việc</param>
    ''' <returns>Chuỗi cần lấy</returns>
    ''' <remarks></remarks>
    Public Shared Function GetLoaiHinh(ByVal LoaiId As Byte) As String
        Dim strResult As String = ""
        Select Case LoaiId
            Case 1
                strResult = "Tập sự"
            Case 2
                strResult = "Bán thời gian"
            Case 3
                strResult = "Hợp đồng ngắn hạn"
            Case 4
                strResult = "Tư vấn"
            Case Else
                strResult = ""
        End Select
        Return strResult
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về danh sách record theo id cán bộ truyền vào
    ''' </summary>
    ''' <param name="_IdCanBo">Id cán bộ truyền vào</param>
    ''' <param name="state">Chỉ số xác định hồ sơ cần trả dữ liệu
    '''                             1: Quyết định nhân sự
    '''                             2: Hồ sơ công tác
    '''                             3: Hồ sơ xuất ngoại
    '''                             4: Hồ sơ hộ chiếu
    '''                             5: Hồ sơ Tham gia lực lượng vũ trang
    '''                             6: Hồ sơ cũ (Thông tin của cán bộ khi chưa thoát ly)
    '''                             7: Hồ sơ Cán bộ học hàm 
    '''                             8: Hồ sơ Cán bộ học vị 
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll_Document(ByVal _IdCanBo As String, ByVal state As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = Nothing

            Select Case state
                Case 1
                    command = New SqlCommand("QDNhanSu_GetForIdCanBo", connection)
                Case 2
                    command = New SqlCommand("HS_CongTac_GetForIdCanBo", connection)
                Case 3
                    command = New SqlCommand("HS_XuatNgoai_GetForIdCanBo", connection)
                Case 4
                    command = New SqlCommand("HS_HoChieu_GetForIdCanBo", connection)
                Case 5
                    command = New SqlCommand("HS_LLVT_GetForIdCanBo", connection)
                Case 6
                    command = New SqlCommand("HS_Cu_GetForIdCanBo", connection)
                Case 7
                    command = New SqlCommand("CB_HocHam_GetForIdCanBo", connection)
                Case 8
                    command = New SqlCommand("CB_HocVi_GetForIdCanBo", connection)
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
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm trả về Quyết định nhân sự mới đây nhất của cán bộ Theo Id cán bộ truyền vào và đơn vị hiện thời của Cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Chỉ số xác định cán bộ</param>
    ''' <param name="_IdDonVi">Chỉ số xác định đơn vị của cán bộ (Với Id_Goc = 0 hoặc Id_Goc = 1)</param>
    ''' <param name="_IsNew">
    ''' Chỉ số xác định loại dữ liệu cần lấy. Với quy ước:
    '''                         0: Bản ghi mới nhất hiện tại khi vào đơn vị trong QĐ nhân sự
    '''                         1: Bản ghi mới nhất trong quyết định nhân sự của cán bộ
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAll_Decision(ByVal _IdCanBo As String, ByVal _IdDonVi As Integer, ByVal _IsNew As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand
            command = New SqlCommand("QDNhanSu_GetForIdCanBo_New", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@_IdDonVi", _IdDonVi))
            command.Parameters.Add(New SqlParameter("@_IsNew", _IsNew))

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
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm thực hiện thay thế một số từ trong chuỗi thành từ viết tắt.
    ''' Mục đích khi xuất dữ liệu ra Word không bị vỡ hình.
    ''' </summary>
    ''' <param name="_ValRep">Chuỗi cần thay thế</param>
    ''' <returns>Chuỗi đã được rút gọn</returns>
    ''' <remarks>Replace_Branch("Công nghệ thông tin") --> "CNTT"</remarks>
    Public Function Replace_Branch(ByVal _ValRep As String) As String
        Dim _result As String = ""
        _ValRep = _ValRep.Replace("Chi nhánh Tỉnh", "NHCSXH").Trim()
        _ValRep = _ValRep.Replace("Trung tâm", "TT").Trim()
        _ValRep = _ValRep.Replace("Công nghệ thông tin", "CNTT").Trim()
        _ValRep = _ValRep.Replace("Phòng GD", "PGD").Trim()
        _ValRep = _ValRep.Replace("Thành phố ", "TP.").Trim()
        _ValRep = _ValRep.Replace("Kĩ thuật Phần mềm & Quản trị ứng dụng", "KTPM&QTUD").Trim()
        _ValRep = _ValRep.Replace("Tín dụng", "TD").Trim()
        _ValRep = _ValRep.Replace("Xây dựng cơ bản", "XDCB").Trim()
        _ValRep = _ValRep.Replace("Văn phòng", "VP").Trim()
        _ValRep = _ValRep.Replace("(CN cấp I & tương đương)", "").Trim()
        _ValRep = _ValRep.Replace("(HSC)", "").Trim()
        _result = _ValRep.Trim()
        Return _ValRep
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về Chuỗi dữ liệu cần lấy về cán bộ mà bạn đang xuất CV
    ''' </summary>
    ''' <param name="_sIdCanBo">Chỉ số xác định cán bộ</param>
    ''' <param name="_IsFlag">Chỉ số xác định dữ liệu cần lấy: Với quy ước 
    '''               1: Chuỗi Ngày vào và Chức vụ trong Đoàn (2 chuỗi này cách nhau dấy chấm phẩy (;))
    '''               2: Chuỗi phụ cấp chức vụ và phụ cấp trách nhiệm của cán bộ
    '''               3: Chuỗi tình trạng sức khoẻ: Chiều cao, cân nặng được lấy cách nhau dấu chấm phẩy (;)
    '''               4: Tập thông tin Học vị của Cán bộ
    '''               5: Chuyên ngành đào tạo, Trường đào tạo, Năm tốt nghiệp
    '''               6: Tập thông tin Học hàm của Cán bộ
    '''               7: Trình độ cao nhất về Ngoại ngữ
    '''               8: Trình độ cao nhất về Tin học
    '''               9: Don vi dau tien tuyển dụng
    '''              10: Ngay vao cq hien dang cong tac   
    '''              11: Hoc ham - hoc vi cao nhat
    '''              12: Danh hieu đươc phong cao nhat
    '''              13: khen thuong cao nhat
    '''              14: ki luat
    ''' </param>
    ''' <returns>Chức vụ cần lấy của cán bộ</returns>
    ''' <remarks></remarks>
    Public Function GetValue(ByVal _sIdCanBo As String, ByVal _IsFlag As Byte) As String
        Dim _result As String = ""
        Dim strSQL As String = ""
        Select Case _IsFlag
            Case 1      'Lấy Chức vụ đoàn và Ngày vào đoàn TNCSHCM hiện nay của cán bộ
                'Hàm thực hiện trả về Mảng Chuỗi chức vụ và Ngày vào trong Đoàn của cán bộ (trong đó đã loại trừ chức vụ là: Đoàn viên bình thường)
                strSQL = "SELECT TOP 1 IdDoanVien,SoThe, Convert(Varchar, NgayVao, 103) as NgayVao, Noi_KetNap,"
                strSQL += "NoiCapThe From HS_DoanVien Where IdCanBo = '" + _sIdCanBo + "'  And (NgayRa Is Null Or NgayRa = '')"
                strSQL += " Order by NgayVao Desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            _result = db.Rows(0)("NgayVao").ToString().Trim()
                            strSQL = "SELECT Top 1 a.*,b.ma_so,b.Ten_Goi FROM HS_Doan a, DanhMuc b WHERE (a.IdDoanVien = '" + db.Rows(0)("IdDoanVien").ToString() + "'"
                            strSQL += " And a.Tungay <= GetDate() And GetDate() <= a.DenNgay) And "
                            strSQL += " (a.IdCVDoan = b.id And b.id_goc = 10 And b.Status = 1) Order by DenNgay Desc"
                            Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (_db Is Nothing) Then
                                    If (_db.Rows.Count > 0) Then
                                        If (_db.Rows(0)("ma_so").ToString().Trim() <> "1013") Then
                                            _result += ";" + _db.Rows(0)("Ten_Goi").ToString().Trim()
                                        End If
                                    End If
                                End If
                            End Using
                        End If
                    End If
                End Using
            Case 2      'Lấy phụ cấp chức vụ của cán bộ
                Dim _pcChucvu As String = ""
                Dim _pcTrachnhiem As String = ""
                strSQL = "SELECT TOP 1 a.IdCB_PhuCap, a.IdCanBo, a.TuNgay, a.DenNgay, b.IdLoai_PhC, b.Muc_PhC FROM HS_phucapCB a, MucPhuCap b"
                strSQL += " WHERE a.IdMucPC=b.IdMuc_PhC and Idcanbo = '" + _sIdCanBo + "'and IdLoai_phC in (Select [id] From DanhMuc Where ma_so = '3101' and Status =1)  and IsQD_NHCS=1"
                strSQL += " Order by TuNgay desc"
                '--> Thực hiện lấy phụ cấp chức vụ của cán bộ
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If db.Rows(0)("DenNgay").ToString().Trim() = "" Or db.Rows(0)("DenNgay") Is DBNull.Value Then
                                _pcChucvu = CType(db.Rows(0)("Muc_PhC").ToString().Trim(), Double).ToString("N2")
                            Else
                                Dim vDenNgay As Date
                                vDenNgay = CType(db.Rows(0)("DenNgay").ToString().Trim(), Date)
                                If Now() <= vDenNgay Or db.Rows(0)("DenNgay") Is DBNull.Value Then
                                    _pcChucvu = CType(db.Rows(0)("Muc_PhC").ToString().Trim(), Double).ToString("N2")
                                End If
                            End If
                        End If
                    End If
                End Using

                '--> Thực hiện lấy phụ cấp Trách nhiệm của cán bộ
                strSQL = "Select a.IdCB_PhuCap,a.IdCanBo,a.TuNgay,a.DenNgay,b.IdLoai_PhC,b.Muc_PhC From HS_PhuCapCB a, MucPhuCap b"
                strSQL += " Where (a.Idcanbo = '" + _sIdCanBo + "') And ((a.Tungay <= GetDate() And GetDate() <= a.DenNgay) Or (a.DenNgay is Null)) And (a.IdMucPC = b.IdMuc_PhC and b.Status = 1) And b.IdLoai_PhC IN "
                strSQL += " (Select [id] From DanhMuc Where ma_so = '3105' and Status =1) Order by TuNgay desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            _pcTrachnhiem = CType(db.Rows(0)("Muc_PhC").ToString().Trim(), Double).ToString("N2")
                        End If
                    End If
                End Using
                _result = _pcChucvu + ";" + _pcTrachnhiem
            Case 3      'Tình trạng sức khoẻ: Chiều cao và cân nặng
                Dim _ChieuCao As String = ""
                Dim _CanNang As String = ""
                Dim _Ketluan As String = ""
                strSQL = String.Format("Select * From  HS_SucKhoe Where IdCanBo = '{0}' Order by Nam Desc", _sIdCanBo)
                Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (_db Is Nothing) Then
                        If _db.Rows.Count > 0 Then
                            If (CType(_db.Rows(0)("ChieuCao"), Double) >= 10) Then
                                _ChieuCao = String.Format("{0:0.00}", CType(_db.Rows(0)("ChieuCao"), Double) / 100)
                            Else
                                _ChieuCao = String.Format("{0:0.00}", CType(_db.Rows(0)("ChieuCao"), Double))
                            End If
                            If (_ChieuCao <> "") Then
                                _ChieuCao = _ChieuCao.Replace(".", "m")
                            End If
                            _CanNang = String.Format("{0:0.00}", CType(_db.Rows(0)("CanNang"), Double))
                            _Ketluan = _db.Rows(0)("KetLuan").ToString().Trim()
                        End If
                    End If
                End Using
                _result = _ChieuCao + ";" + _CanNang + ";" + _Ketluan
            Case 4      'Lấy thông tin Học vị
                strSQL = "Select a.*,b.Ten_Goi,b.ma_so from CB_HocVi a, DanhMuc b Where IdCanBo = '" + _sIdCanBo + "'"
                strSQL += " And (b.id = a.IdHocVi And b.id_goc = 20 And b.ma_so <> '2007' And b.Status = 1) Order by b.ma_so Asc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                _result += db.Rows(i)("Ten_Goi").ToString().Trim() + ", "
                            Next
                        End If
                    End If
                End Using
                _result = _result.Trim()
                If (_result <> "") Then
                    If (_result.Substring(0, 1) = ",") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                        _result = _result.Substring(1).Trim()
                    End If
                    While _result.EndsWith(",")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
                        _result = _result.Substring(0, _result.Length - 1)
                    End While
                End If
            Case 5
                strSQL = "SELECT a.*,b.Ten_Goi as ChuyenNganhDT FROM HS_DTVBCC a, DanhMuc b Where a.IdCanBo = '" + _sIdCanBo + "' And"
                strSQL += " (a.IdChuyenNganhDT = b.id and b.id_goc = 11 and b.Status = 1) And a.VBCC=1 And a.HoanThanh = 1"
                strSQL += "  And (a.NgayHH >= GetDate() Or a.NgayHH Is Null) Order by a.NamTN Desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            _result = "Chuyên ngành đào tạo: " + IIf(db.Rows(0)("ChuyenNganhDT").ToString().Trim() <> "", db.Rows(0)("ChuyenNganhDT").ToString().Trim() + ", " & vbTab, vbTab & vbTab & vbTab)
                            _result += "Trường Đào tạo: " + IIf(db.Rows(0)("CoSo_DT").ToString().Trim() <> "", db.Rows(0)("CoSo_DT").ToString().Trim() + ", " & vbTab, vbTab & vbTab & vbTab)
                            _result += "Năm tốt nghiệp: " + IIf(db.Rows(0)("NamTN").ToString().Trim() <> "", db.Rows(0)("NamTN").ToString().Trim() & vbTab, "")
                        End If
                    End If
                End Using
            Case 6      'Lấy thông tin học hàm của cán bộ
                strSQL = "Select a.*,b.Ten_Goi,b.ma_so from CB_HocHam a, DanhMuc b Where IdCanBo = '" + _sIdCanBo + "'"
                strSQL += " And (b.id = a.IdHocHam And b.id_goc = 26 And b.ma_so <> '2603' And b.Status = 1) Order by b.ma_so Asc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                _result += db.Rows(i)("Ten_Goi").ToString().Trim() + ", "
                            Next
                        End If
                    End If
                End Using
                _result = _result.Trim()
                If (_result <> "") Then
                    If (_result.Substring(0, 1) = ",") Then   '   Bỏ dấu phẩy ở đầu chuỗi nếu có
                        _result = _result.Substring(1).Trim()
                    End If
                    While _result.EndsWith(",")       '   Bỏ dấu phẩy (,) ở cuối chuỗi nếu có
                        _result = _result.Substring(0, _result.Length - 1)
                    End While
                End If
            Case 7      'Lấy thông tin Trình độ cao nhất về Ngoại ngữ
                strSQL = "SELECT (Select ma_so from DanhMuc Where id = a.IdTrinhDo And id_goc = 38 And Status = 1) As ma_so,"
                strSQL += " (Select Ten_Goi from DanhMuc Where id = a.IdTrinhDo And id_goc = 38 And Status = 1) As TrinhDo, a.*"
                strSQL += " FROM HS_DTVBCC a left join DanhMuc b on a.IdLoaiVBCC = b.id Where a.IdCanBo = '" + _sIdCanBo + "' "
                strSQL += " And b.id_goc = 33 and b.Status = 1 And b.ma_so = '3302' And a.HoanThanh = 1"
                strSQL += " Order by Ma_so, a.VBCC, a.NamTN Desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            _result = db.Rows(0)("TrinhDo").ToString().Trim()
                        End If
                    End If
                End Using
                ' Trinh do dai hoc --> Ngoai ngu trinh do C
                'If _result.ToString.Trim = "" Or _result.ToString.Trim = "A" Or _result.ToString.Trim = "B" Then
                '    Dim db As DBAccess = New DBAccess
                '    If db.getString("select IdDTVBCC from HS_DTVBCC WHERE IDcanbo='" + _sIdCanBo + "' and IdtrinhDo in (SELECT id FROM danhmuc Where id_goc=38 and ma_so in ('3801','3802','3803'))") <> "" Then
                '        _result = " C "
                '    End If
                'End If
            Case 8      'Lấy thông tin Trình độ cao nhất về tin học
                Dim IDMAX As Integer = getChinhDo_Max(_sIdCanBo, "1113")
                If IDMAX > 0 Then
                    Dim db As DBAccess = New DBAccess
                    _result = db.getString("Select Ten_Goi from DanhMuc Where id = " & IDMAX & " And id_goc = 38 And Status = 1")
                Else
                    _result = ""
                End If
            Case 9      'Don vi dau tien tuyển dụng
                ''Trước hết kiểm tra trong HS_CongTac có bản ghi nào trước Ngày vào NHCSXH không?
                ''Nếu không có thực hiện lấy bản ghi ngày gần đây nhất của Hồ sơ cũ
                'strSQL = "SELECT a.* FROM HS_CongTac a, HS_CanBo b WHERE a.IdCanBo = b.IdCanBo And a.IdCanBo = '" + _sIdCanBo + "'"
                'strSQL += " And (a.TuNgay < b.Ngay_VBSP  And Convert(Varchar, a.TuNgay, 103) <> Convert(Varchar, b.Ngay_VBSP, 103))"
                'strSQL += " And a.DenNgay Is Not Null Order By a.DenNgay Desc"
                'Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '    If Not (db Is Nothing) Then
                '        If db.Rows.Count > 0 Then
                '            If (db.Rows(0)("ChucVu").ToString().Trim() <> "") Then
                '                _result = db.Rows(0)("ChucVu").ToString().Trim()
                '                If (db.Rows(0)("DiaChi").ToString().Trim() <> "") Then
                '                    _result += " (" + db.Rows(0)("DiaChi").ToString().Trim() + ")"
                '                End If
                '            End If
                '        End If
                '    End If
                'End Using
                '_result = _result.Trim()
                'If (_result = "") Then
                '    strSQL = "SELECT * FROM HS_Cu WHERE IdCanBo = '" + _sIdCanBo + "' Order By DenThang Desc"
                '    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '        If Not (db Is Nothing) Then
                '            If db.Rows.Count > 0 Then
                '                If (db.Rows(0)("NgheNghiep").ToString().Trim() <> "") Then
                '                    _result = db.Rows(0)("NgheNghiep").ToString().Trim()
                '                    If (db.Rows(0)("DiaChi").ToString().Trim() <> "") Then
                '                        _result += " (" + db.Rows(0)("DiaChi").ToString().Trim() + ")"
                '                    End If
                '                End If
                '            End If
                '        End If
                '    End Using
                'End If

                'Bỏ từ 1/1/2022 vì đầy là quyết định không phải của NHCSXH
                'strSQL += " SELECT NoiDung FROM QDNhanSu WHERE Idcanbo='" + _sIdCanBo + "' and isQD_NHCS=0 and CHARINDEX('NHCS',upper(So_QD))=0 and IdLoaiQD not in (SELECT id FROM DanhMuc WHERE id_goc=15 and ma_so in ('1505','1510','1519')) order by NgayHL asc"

                'Chuyển sang lấy QĐ của NHCSXH
                ' --1505 - Miễn nhiệm; 1510 - Thôi giữ chức (Cách chức); 1519 - Thôi Kiêm nhiệm
                strSQL = " Select (ZZ.ChucVu_HT + N', ' + PhongBan_HT + N', '+ DonVi_TMP) NoiDung From "
                strSQL += " ( "
                strSQL += " Select Top 1 dbo.GetTenDMuc(A.IdChucVu_Moi,1) ChucVu_HT,dbo.Replace_Ten_PhBanCN(dbo.GetTenDMuc(A.IdPhong_Moi,3)) PhongBan_HT,"
                strSQL += "        dbo.GetTenDMuc(A.IdDonVi_Moi,2) DonVi_HT,dbo.GetTenDMuc(A.IdDonVi_Moi,5) DonVi_Cd,"
                strSQL += "        dbo.GetTenDMuc_MaSo(dbo.GetTenDMuc(A.IdDonVi_Moi,5),12) MaCN,dbo.GetTenDMuc_MaSo(dbo.GetTenDMuc(A.IdDonVi_Moi,5),13) TenChiNhanh,"
                strSQL += "        dbo.Replace_Ten_PhBanCN("
                strSQL += "         Case When dbo.GetTenDMuc(A.IdDonVi_Moi,5) = dbo.GetTenDMuc_MaSo(dbo.GetTenDMuc(A.IdDonVi_Moi,5),12) Then dbo.GetTenDMuc(A.IdDonVi_Moi,2)"
                strSQL += "              Else  dbo.GetTenDMuc(A.IdDonVi_Moi,2) + N' - ' + dbo.GetTenDMuc_MaSo(dbo.GetTenDMuc(A.IdDonVi_Moi,5),13)"
                strSQL += "         End"
                strSQL += "        ) DonVi_TMP,"
                strSQL += "  A.* FROM QDNhanSu A Where IdCanBo='" + _sIdCanBo + "' And IsQD_NHCS = 1 And IsKiemNhiem = 0 "
                strSQL += " And IdLoaiQD Not In (Select Id From DanhMuc Where Id_Goc = 15 And Ma_So in ('1505','1510','1519')) "
                strSQL += " Order By NgayHL Asc "
                strSQL += " ) ZZ "
                _result = SoftSqlHelper.GetString(strSQL, "")

                'Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '    If Not (db Is Nothing) Then
                '        If db.Rows.Count > 0 Then
                '            If (db.Rows(0)("NoiDung").ToString().Trim() <> "") Then
                '                _result = db.Rows(0)("NoiDung").ToString().Trim()
                '            Else
                '                _result = ""
                '            End If
                '        End If
                '    End If
                'End Using
            Case 10
                Dim dbconn As DBAccess = New DBAccess
                Dim idDV As Integer = dbconn.getNumber(" SELECT TOP 1 IdDonVi_Moi FROM QDNhanSu WHERE Idcanbo='" + _sIdCanBo + "' and IsQD_NHCS=1 order by NgayHL desc")
                strSQL += "  SELECT TOP 1 NgayHL, Convert(Varchar, NgayHL, 103) as NgayHL_VN FROM QDNhanSu WHERE Idcanbo='" + _sIdCanBo + "' and (IdDonVi_Moi=" & idDV & " or IdDonVi_Moi in (SELECT ID FROM Chinhanh WHERE id_goc=" & idDV & ")) order by NgayHL asc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            If (db.Rows(0)("NgayHL_VN").ToString().Trim() <> "") Then
                                _result = db.Rows(0)("NgayHL_VN").ToString().Trim()
                            Else
                                _result = ""
                            End If
                        End If
                    End If
                End Using
            Case 11
                'Lấy thông tin học hàm của cán bộ
                strSQL = "Select a.*,b.Ten_Goi,b.ma_so from CB_HocHam a, DanhMuc b Where IdCanBo = '" + _sIdCanBo + "'"
                strSQL += " And (b.id = a.IdHocHam And b.id_goc = 26 And b.ma_so <> '2603' And b.Status = 1) Order by b.ma_so Asc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                _result += db.Rows(i)("Ten_Goi").ToString().Trim() + ", "
                            Next
                        End If
                    End If
                End Using
                _result = _result.Trim()
                If (_result <> "") Then
                    If (_result.Substring(0, 1) = ",") Then   '   Bỏ dấu phẩy ở đầu chuỗi nếu có
                        _result = _result.Substring(1).Trim()
                    End If
                    While _result.EndsWith(",")       '   Bỏ dấu phẩy (,) ở cuối chuỗi nếu có
                        _result = _result.Substring(0, _result.Length - 1)
                    End While
                End If
                'Hoc vi:(SELECT ma_so FROM DanhMuc WHERE id=IdTrinhDo) as MaTrinhDo
                strSQL = "SELECT (SELECT X.Ten_Goi FROM DanhMuc X WHERE X.Id=IdTrinhDo) as TrinhDo, (SELECT X.Ten_Goi FROM DanhMuc X WHERE X.Id=A.IdChuyenNganhDT) as ChuyenNganhDT , * "
                strSQL += " FROM HS_DTVBCC A Where A.IdCanBo = '" + _sIdCanBo + "' And A.VBCC=1 And A.HoanThanh = 1 Order by A.MaTrinhDo asc, NamTN"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If _result <> "" Then _result = " - "
                        If db.Rows.Count > 0 Then
                            _result += db.Rows(0)("TrinhDo").ToString().Trim()
                            _result += IIf(db.Rows(0)("NamTN").ToString().Trim() <> "", " Năm " & db.Rows(0)("NamTN").ToString().Trim(), " ")
                            _result += " Chuyên ngành đào tạo: " + IIf(db.Rows(0)("ChuyenNganhDT").ToString().Trim() <> "", db.Rows(0)("ChuyenNganhDT").ToString().Trim(), "")
                            '_result += "Trường Đào tạo: " + IIf(db.Rows(0)("CoSo_DT").ToString().Trim() <> "", db.Rows(0)("CoSo_DT").ToString().Trim() + ", " & vbTab, vbTab & vbTab & vbTab)
                            '_result += "Năm tốt nghiệp: " + IIf(db.Rows(0)("NamTN").ToString().Trim() <> "", db.Rows(0)("NamTN").ToString().Trim() & vbTab, "")
                        End If
                    End If
                End Using
            Case 12      'Lấy danh hieu được phong cao nhat
                strSQL = "SELECT a.IdKhenThuong, a.NamKT, a.SoQD, a.NgayQD, a.IdCapKT, N'', a.IdChucVuKyQD, N'', a.NguoiKyQD, a.KhenThuong, N'',"
                strSQL += " b.IdKhenThuong_CT, b.IdDonvi, N'', b.IdPhong, N'', b.IdDanhHieuHinhThuc, N'', b.DaDuyet, b.GhiChu,"
                strSQL += " (SELECT ma_so FROM DanhMuc WHERE id=a.IdCapKT) as MaCapKT,"
                strSQL += " (SELECT DanhHieu_HinhThuc FROM thiduakhenthuong WHERE idTDKT=b.IdDanhHieuHinhThuc) as DanhHieu"
                strSQL += " FROM     HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong"
                strSQL += " WHERE b.IdCN_TT='" + _sIdCanBo + "'"
                strSQL += " 	AND IdDanhHieuHinhThuc in (SELECT IdTDKT FROM ThiDuaKhenThuong WHERE maso in('16061'))"
                strSQL += " Order by MaCapKT asc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            _result = db.Rows(0)("DanhHieu").ToString().Trim() & " Năm " & db.Rows(0)("NamKT").ToString().Trim()
                        End If
                    End If
                End Using
            Case 13      'Lấy huan huy chuong
                strSQL = "SELECT a.IdKhenThuong, a.NamKT, a.SoQD, a.NgayQD, a.IdCapKT, N'', a.IdChucVuKyQD, N'', a.NguoiKyQD, a.KhenThuong, N'',"
                strSQL += " b.IdKhenThuong_CT, b.IdDonvi, N'', b.IdPhong, N'', b.IdDanhHieuHinhThuc, N'', b.DaDuyet, b.GhiChu,"
                strSQL += " (SELECT maso FROM thiduakhenthuong WHERE idTDKT=b.IdDanhHieuHinhThuc) as MaKT,"
                strSQL += " (SELECT DanhHieu_HinhThuc FROM thiduakhenthuong WHERE idTDKT=b.IdDanhHieuHinhThuc) as DanhHieu"
                strSQL += " FROM     HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong"
                strSQL += " WHERE b.IdCN_TT='" + _sIdCanBo + "'"
                strSQL += " 	AND IdDanhHieuHinhThuc in (SELECT IdTDKT FROM ThiDuaKhenThuong WHERE maso in('15051','15052','15053','15081','15091','15101','15102','15103'))"
                strSQL += " Order by MaKT desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            _result = db.Rows(0)("DanhHieu").ToString().Trim() & " Năm " & db.Rows(0)("NamKT").ToString().Trim()
                        End If
                    End If
                End Using
            Case 14      'Ki luat
                strSQL = "SELECT (SELECT Ten_Goi FROM DanhMuc WHERE id=IdHinhThucKL) as HinhThuc, Lydo, year(TuNgay) as Nam"
                strSQL += "  FROM HS_KiLuat WHERE idcanbo='" + _sIdCanBo + "'"
                strSQL += " Order by Nam desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            _result = db.Rows(0)("HinhThuc").ToString().Trim() & ", " & db.Rows(0)("Lydo").ToString().Trim() & ", năm " & db.Rows(0)("Nam").ToString().Trim()
                        End If
                    End If
                End Using
        End Select
        _result = _result.Trim()
        Return _result
    End Function

    ''' <summary>
    ''' Hàm thực hiện lấy các thông tin về Đảng của cán bộ để xuất CV
    ''' </summary>
    ''' <param name="_sIdCanBo">Chỉ số xác định cán bộ</param>
    ''' <returns>
    ''' Một chuỗi các thông tin về đảng cần lấy của cán bộ (Các thông tin này cách nhau bởi dấu chấm phẩy(;)). Quy ước thứ tự như sau
    '''                 -> Chức vụ đảng
    '''                 -> Cấp uỷ hiện tại
    '''                 -> Ngày kết nạp
    '''                 -> Ngày vào chính thức
    '''                 -> Chức vụ kiêm nhiệm
    '''                 -> Cấp uỷ kiêm nhiệm
    ''' </returns>
    ''' <remarks></remarks>
    Public Function GetPartys(ByVal _sIdCanBo As String) As String()
        Dim _arrPartys(5) As String
        Dim strSQL As String = ""
        strSQL = "SELECT TOP 1 IdDangVien,SoThe, Convert(Varchar, NgayKN, 103) as NgayKN, Convert(Varchar, NgayVao, 103) as NgayVao, Noi_KetNap,"
        strSQL += "Nguoi_GT, NoiCapThe From HS_DangVien Where IdCanBo = '" + _sIdCanBo + "'  And (NgayRa Is Null OR NgayRa = '')"
        strSQL += " Order by NgayVao Desc"
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _arrPartys(2) = db.Rows(0)("NgayKN").ToString().Trim()
                    _arrPartys(3) = db.Rows(0)("NgayVao").ToString().Trim()
                    'Câu lệnh lấy thông tin về đảng chính (IsKiemNhiem = 0)
                    strSQL = "SELECT Top 1 a.*,b.ma_so,b.Ten_Goi FROM HS_Dang a, DanhMuc b WHERE (a.IdDangVien = '" + db.Rows(0)("IdDangVien").ToString() + "'"
                    strSQL += " And a.IsKiemNhiem = 0 And a.Tungay <= GetDate() And (GetDate() <= a.DenNgay or a.DenNgay ='01/01/1900')) And "
                    strSQL += " (a.IdCVDang = b.id And b.id_goc = 16 And b.Status = 1) Order by DenNgay Desc"
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If (_db.Rows.Count > 0) Then
                                _arrPartys(1) = _db.Rows(0)("ChiBo").ToString().Trim()
                                'Lấy chức vụ chính (Loại bỏ chức vụ đảng viên bình thường và chức vụ khác)
                                If (_db.Rows(0)("ma_so").ToString().Trim() <> "1621" And _db.Rows(0)("ma_so").ToString().Trim() <> "1622" And _db.Rows(0)("ma_so").ToString().Trim() <> "1623") Then
                                    _arrPartys(0) = _db.Rows(0)("Ten_Goi").ToString().Trim()
                                End If
                            End If
                        End If
                    End Using
                    'Lấy chức vụ kiêm và cấp uỷ kiêm
                    strSQL = "SELECT Top 1 a.*,b.ma_so,b.Ten_Goi FROM HS_Dang a, DanhMuc b WHERE (a.IdDangVien = '" + db.Rows(0)("IdDangVien").ToString() + "'"
                    strSQL += " And a.IsKiemNhiem = 1 And a.Tungay <= GetDate() And (GetDate() <= a.DenNgay or a.DenNgay ='01/01/1900')) And "
                    strSQL += " (a.IdCVDang = b.id And b.id_goc = 16 And b.Status = 1) Order by DenNgay Desc"
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If (_db.Rows.Count > 0) Then
                                _arrPartys(5) = _db.Rows(0)("ChiBo").ToString().Trim()
                                'Lấy chức vụ chính (Loại bỏ chức vụ đảng viên bình thường và chức vụ khác)
                                If (_db.Rows(0)("ma_so").ToString().Trim() <> "1621" And _db.Rows(0)("ma_so").ToString().Trim() <> "1622" And _db.Rows(0)("ma_so").ToString().Trim() <> "1623") Then
                                    _arrPartys(4) = _db.Rows(0)("Ten_Goi").ToString().Trim()
                                End If
                            End If
                        End If
                    End Using
                End If
            End If
        End Using
        Return _arrPartys
    End Function

    ''' <summary>
    ''' Hàm trả về quê quán của thành viên Gia đình cán bộ
    ''' </summary>
    ''' <param name="pIdQueQuan">Id quê quán (Tưởng đương Id Xã/Phường)</param>
    ''' <returns>Địa chỉ quê quán (Huyện - Tỉnh)</returns>
    ''' <remarks></remarks>
    Public Function GetAddress(ByVal pIdQueQuan As Integer) As String
        Dim _result As String = ""
        Dim _XaName As String = ""
        Dim _TinhName As String = ""
        Dim strSQL As String = ""
        strSQL = String.Format("Select * From Dm_DiaPhuong Where Ma_Xa<>'00' And Ma_Thon='00' And Id = {0} Order By TrangThai", pIdQueQuan)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _XaName = db.Rows(0)("Ten_Thon").ToString().Trim()
                    strSQL = String.Format("Select Id, Ten_Thon As Ten_Goi From Dm_DiaPhuong Where Ma_Xa='00' And Ma_Thon='00' And Ma_Tinh='{0}' Order By TrangThai", db.Rows(0)("Ma_Tinh").ToString())
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If (_db.Rows.Count > 0) Then
                                _TinhName = _db.Rows(0)("Ten_Thon").ToString().Trim()
                            End If
                        End If
                    End Using
                End If
            End If
        End Using
        If (_XaName <> "") Then
            _result = _XaName
        End If
        If (_TinhName <> "") Then
            _result += " - " & _TinhName
        End If
        If _result.Trim <> "" Then
            If (_result.Substring(0, 3) = " - ") Then
                _result = _result.Substring(3)
            End If
            If (_result.EndsWith(" - ")) Then
                _result = _result.Substring(0, _result.Length - 3)
            End If
            Return _result.Replace("Huyện", "H.").Replace("Quận", "Q.").Replace("Thành phố", "TP.").Trim()
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Xuất lý lịch theo mẫu 2C/TCTW-98
    ''' </summary>
    ''' <param name="_RowId"></param>
    ''' <remarks></remarks>
    Public Sub ExportWord_2C(ByVal _RowId As String)
        Dim _sDiaban As String = DIABAN
        Dim _sDonviTT As String = ""        'Biến lưu chuỗi Đơn vị Trực thuộc
        Dim _sDonviCS As String = ""        'Biến lưu chuỗi Đơn vị Cơ sở
        Dim _IdDvParent As Integer = 0          'Biến lưu Chỉ số đơn vị cấp nhỏ nhất của Cán bộ
        Dim _IdDvChild As Integer = 0           'Biến lưu Chỉ số đơn vị cấp nhỏ nhất của Cán bộ
        Dim _IdPb As Integer = 0                'Biến lưu chỉ số Phòng ban của cán bộ
        Dim _IdCv As Integer = 0                'Biến lưu chỉ số xác định chức vụ Cán bộ
        Dim _IdChmon As Integer = 0             'Biến lưu chỉ số xác Chuyên môn mới nhất của cán bộ (Công việc chính đang làm)
        Dim _sVal As String = ""
        Dim strSQL As String = ""
        Dim sNonDate As String = "...../...../.........."
        Dim _HS_CanBo As clsHS_CanBo = New clsHS_CanBo()

        'Dim s() As String
        's(0) = "1"
        's(1) = "3"
        's(2) = "5"

        Dim dr As DataRow
        dr = _HS_CanBo.GetHuman_ForCode(_RowId)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                'Lấy các thông tin chính để xuất dữ liệu phần đầu
                _IdDvParent = dr("IdDonVi").ToString().Trim()
                'Lấy dữ liệu Đơn vị cấp dưới của cán bộ đang công tác, Phòng ban, chức vụ
                Dim HS_QDNhansu As QDNhanSu = New QDNhanSu
                Dim dbconn As DBAccess = New DBAccess
                HS_QDNhansu = HS_QDNhansu.getFinalRecord(_RowId, Nothing, True)
                If Not (HS_QDNhansu Is Nothing) Then
                    If (HS_QDNhansu.IdQDNhanSu <> "") Then
                        _IdCv = IIf(HS_QDNhansu.IdChucvu_Moi <> 0, HS_QDNhansu.IdChucvu_Moi, 0)
                        _IdPb = IIf(HS_QDNhansu.IdPhong_Moi <> 0, HS_QDNhansu.IdPhong_Moi, 0)
                        _IdChmon = IIf(HS_QDNhansu.IdChuyenMon_Moi <> 0, HS_QDNhansu.IdChuyenMon_Moi, 0)
                        _IdDvChild = IIf(HS_QDNhansu.IdDonvi_Moi <> 0, HS_QDNhansu.IdDonvi_Moi, 0)
                    End If
                End If

                If dbconn.getNumber("SELECT count(*) FROM Phongban WHERE ma_so in ('01','02','03') and [id]=" & _IdPb) > 0 Then
                    _sDonviTT = "Đơn vị trực thuộc: Ngân hàng Chính sách xã hội Việt Nam"
                Else
                    strSQL = String.Format("Select id, Ten_Goi from ChiNhanh Where Status = 1 and id = {0} And id_goc IN (0,1)", _IdDvParent)
                    _sDonviTT = "Đơn vị trực thuộc: " + Replace_Branch(_HS_CanBo.GetNameByCode(strSQL).ToString())
                End If

                If (_IdDvParent <> _IdDvChild) Then 'Phòng giao dịch
                    strSQL = String.Format("Select id, Ten_Goi from ChiNhanh Where Status = 1 and id = {0} And id_goc > 1", _IdDvChild)
                    _sDonviCS = "Đơn vị cơ sở: " + Replace_Branch(_HS_CanBo.GetNameByCode(strSQL).ToString())
                ElseIf (_IdDvParent = _IdDvChild) Then 'Phòng ban của tỉnh hoặc Hội sở
                    strSQL = String.Format("Select id, ten_phong From PhongBan Where Status = 1 And id = {0}", _IdPb)
                    _sDonviCS = "Đơn vị cơ sở: " + Replace_Branch(_HS_CanBo.GetNameByCode(strSQL).ToString())
                End If

                'Bắt đầu khai báo và xuất dữ liệu ra Word
                Dim objWordApp As New Word.Application      'Tạo Đối tượng Word Application
                Dim objDocument As New Word.Document        'Tạo đối tượng Word Document
                Dim objTable As Word.Table                  'Tạo đối tượng Word Table
                Dim oMissing As Object = System.Reflection.Missing.Value

                'Start Word and open the document template.
                'objWordApp.Visible = True                          'And show word screen (False: Hide msword cho đến khi xuất hết thông tin)
                'objWordApp.Activate()
                objDocument = objWordApp.Documents.Add                  'Add một Document vào trong Application (Create a new word document)
                objDocument.Paragraphs.LineSpacing = 14 'objDocument.LinesToPoints(1.15)
                objDocument.Paragraphs.SpaceAfter = 3
                objDocument.Paragraphs.SpaceBefore = 0

                'Biến lưu vị trí select hiện hành
                Dim objselection As Word.Selection
                'Gán vị trí hiện hành trong Document vào biến selection
                objselection = objDocument.Application.Selection()

                'Định dạng Paragraph
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Size = 13
                objselection.Font.Name = "Times New Roman"
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                'objWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader
                ''THE LOGO IS ASSIGNED TO A SHAPE OBJECT SO THAT WE CAN USE ALL THE 
                ''SHAPE FORMATTING OPTIONS PRESENT FOR THE SHAPE OBJECT 
                'Dim logoCustom As Word.Shape = Nothing
                'Dim logoPath As String = gBackUpFileDir & "\_icon\file_icon.ico"
                'logoCustom = objWordApp.Selection.HeaderFooter.Shapes.AddPicture(logoPath, False, True, oMissing, oMissing, oMissing, oMissing, oMissing)
                'logoCustom.[Select](oMissing)
                'logoCustom.Name = "VBSPLogo"
                'logoCustom.Left = CSng(Word.WdShapePosition.wdShapeLeft)
                'objWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument

                'Tạo một bảng gồm 3 dòng và 3 cột
                'objTable = objselection.Tables.Add(objDocument.Bookmarks.Item("\endofdoc").Range, 3, 3)
                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 3, 3)
                objTable.Columns(1).Width = 180.0F
                objTable.Columns(2).Width = 162.0F
                objTable.Columns(3).Width = 120.0F
                objTable.Cell(1, 2).Merge(objTable.Cell(3, 2))
                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.LeftPadding = 0 '-10
                objTable.RightPadding = 0 '-10
                objTable.Cell(1, 2).Range.InsertAfter("SƠ YẾU LÝ LỊCH")
                objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                objTable.Cell(1, 2).Range.Font.Name = "Times New Roman"
                objTable.Cell(1, 2).Range.Font.Size = 20
                objTable.Cell(1, 2).Range.Font.Bold = 1

                objTable.Cell(1, 1).Range.InsertAfter("Tỉnh (TP): " + _sDiaban)
                objTable.Cell(2, 1).Range.InsertAfter(_sDonviTT)
                objTable.Cell(3, 1).Range.InsertAfter(_sDonviCS)
                objTable.Cell(1, 3).Range.InsertAfter("Mẫu 2C/TCTW-98")
                objTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                objTable.Cell(3, 3).Range.Font.Bold = 1
                objTable.Cell(3, 3).Range.InsertAfter(dr("MaCB").ToString().Trim().ToUpper())
                objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.InsertParagraph()

                'Xóa định dạng Paragraph trước
                objselection.ClearFormatting()
                'Định dạng lại Paragraph
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Size = 13
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                'Add Text vào Paragraph
                objselection.Paragraphs.SpaceAfter = 8
                objselection.Font.Name = "Times New Roman"
                objselection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                objselection.TypeText("Số hiệu cán bộ, công chức")

                'BẮT ĐẦU XUẤT DỮ LIỆU CHÍNH CỦA HỒ SƠ CÁN BỘ

                ''Tạo một bảng có 8 dòng và 2 cột
                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 6, 2)
                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter

                'objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                'objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                'objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.Range.Paragraphs.SpaceBefore = 1
                objTable.Range.Paragraphs.SpaceAfter = 1

                If dr("AnhThe").ToString().Trim() = "" Then
                    _sVal = "Ảnh 4*6"
                    objselection.TypeText(_sVal)
                    objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Else
                    'Dim picbx As System.Windows.Forms.PictureBox = New System.Windows.Forms.PictureBox
                    'picbx.Width = 15
                    'picbx.Height = 20
                    '_HS_CanBo.Show_Picture(dr("AnhThe").ToString(), picbx)
                    ''objselection.InsertFile(String.Format("{0}{1}", Application.StartupPath, dr("AnhThe").ToString().Trim()))
                    'objselection.InlineShapes.AddPicture(String.Format("{0}{1}", Application.StartupPath, dr("AnhThe").ToString().Trim()))
                    ''objselection.InlineShapes.AddPicture(picbx.ImageLocation)
                    Try
                        Dim oPic As Word.InlineShape
                        oPic = objselection.InlineShapes.AddPicture(FileName:=String.Format("{0}{1}", Application.StartupPath, dr("AnhThe").ToString().Trim()),
                        LinkToFile:=False, SaveWithDocument:=True)
                        oPic.Width = 90
                        oPic.Height = 140
                        objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    Catch ex As Exception
                        _sVal = "Ảnh 4*6"
                        objselection.TypeText(_sVal)
                        objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    End Try
                End If
                objTable.Cell(1, 1).Merge(objTable.Cell(6, 1))

                _sVal = dr("HoTen").ToString().Trim().ToUpper() & vbTab & vbTab & " Nam, Nữ: " & IIf(dr("GioiTinh").ToString() = False, " Nam", " Nữ")
                objTable.Cell(1, 2).Range.InsertAfter(" 1) Họ và tên khai sinh: " & _sVal)
                objTable.Cell(1, 2).Range.Paragraphs.LeftIndent = 3.0F
                objTable.Cell(1, 2).Range.Font.Size = 13
                objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                objTable.Cell(2, 2).Range.InsertAfter(" 2) Các tên gọi khác: " & dr("TenThuongGoi").ToString().Trim())
                objTable.Cell(2, 2).Range.Paragraphs.LeftIndent = 3.0F
                objTable.Cell(2, 2).Range.Font.Size = 13
                objTable.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                'Write 4: Chức vụ (Đảng, đoàn thể, chính quyền, kể cả chức vụ kiêm nhiệm)
                '--->Lấy chức vụ đảng trong thời gian hiện tại (Tức chức vụ đó thì ngày hiện tại phải thuộc khoảng từ ngày 1 đến ngày 2)
                Dim arrPartys() As String       'Lấy mảng các thông tin về đảng của cán bộ
                arrPartys = GetPartys(_RowId)
                'Lấy Chức vụ đảng của cán bộ (Chức vụ chính và Chức vụ Kiêm)
                If ((arrPartys(0) <> "" Or Not (arrPartys(0) Is Nothing)) And (arrPartys(4) <> "" Or Not (arrPartys(4) Is Nothing))) Then
                    _sVal = arrPartys(0) + "; " + arrPartys(4)
                Else
                    _sVal = IIf(arrPartys(0) <> "", arrPartys(0), IIf(arrPartys(4) <> "", arrPartys(4), ""))
                End If
                Dim _sDoanvien As String = GetValue(_RowId, 1) 'Lấy chuỗi thông tin đoàn
                If (_sDoanvien <> "" And _sDoanvien.IndexOf(";") >= 0) Then
                    If (_sVal <> "") Then
                        _sVal += "; " + _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
                    Else
                        _sVal = _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
                    End If
                End If

                'If (_sVal <> "") Then
                '    _sVal += "; " + HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                'Else
                '    _sVal = HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                'End If

                'If arrPartys(4) <> "" Or Not (arrPartys(4) Is Nothing) Then

                If (_sVal <> "") Then
                    If _IdPb = 3 Then
                        _sVal += "; " & _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " Ngân hàng Chính sách xã hội"
                    Else
                        If _IdPb = 15 Then
                            _sVal += "; " & _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " " & _HS_CanBo.GetNameByCode(String.Format("Select id, Ten_Goi from Chinhanh Where id = {0} ", _IdDvChild))
                        Else
                            _sVal += "; " & _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                        End If
                    End If
                Else
                    If _IdPb = 3 Then
                        _sVal = _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " Ngân hàng Chính sách xã hội"
                    Else
                        If _IdPb = 15 Then
                            _sVal = _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " " & _HS_CanBo.GetNameByCode(String.Format("Select id, Ten_Goi from Chinhanh Where id = {0} ", _IdDvChild))
                        Else
                            _sVal = _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                        End If
                    End If
                End If

                _sVal = _sVal.Replace(";;", ";").Replace("; ;", ";").Trim()
                While _sVal.EndsWith("-") 'Bỏ dấu gạch ngang (-) ở cuối chuỗi nếu có
                    _sVal = _sVal.Substring(0, _sVal.Length - 1).Trim()
                End While
                If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                    _sVal = _sVal.Substring(1).Trim()
                End If
                While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
                    _sVal = _sVal.Substring(0, _sVal.Length - 1)
                End While
                If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
                    _sVal += vbTab & vbTab & vbTab
                Else
                    _sVal += vbTab
                End If

                '3. Write 3: Cấp uỷ hiện tai, cấp uỷ kiêm
                'objTable.Cell(3, 2).Range.InsertAfter(" 3) Cấp uỷ hiện tại: " + IIf(arrPartys(1) <> "", arrPartys(1), ".......................") + ", Cấp uỷ kiêm: " + IIf(arrPartys(4) <> "", arrPartys(4) & " " & arrPartys(5), "....................."))
                objTable.Cell(3, 2).Range.InsertAfter(" 3) Cấp uỷ hiện tại: " + IIf(arrPartys(0) <> "", arrPartys(0) & " " & arrPartys(1), ".......................") + ", Cấp uỷ kiêm: " + IIf(arrPartys(4) <> "", arrPartys(4) & " " & arrPartys(5), "....................."))
                objTable.Cell(3, 2).Range.Paragraphs.LeftIndent = 3.0F
                objTable.Cell(3, 2).Range.Font.Size = 13
                objTable.Cell(3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                'Chuc vu
                objTable.Cell(4, 2).Range.InsertAfter("     Chức vụ: " & _sVal)
                objTable.Cell(4, 2).Range.Paragraphs.LeftIndent = 3.0F
                objTable.Cell(4, 2).Range.Font.Size = 13
                objTable.Cell(4, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                'Phu cap
                _sVal = GetValue(_RowId, 2)
                _sVal = "     Phụ cấp chức vụ: " + _sVal.Substring(0, _sVal.IndexOf(";"))
                objTable.Cell(5, 2).Range.InsertAfter(_sVal)
                objTable.Cell(5, 2).Range.Paragraphs.LeftIndent = 3.0F
                objTable.Cell(5, 2).Range.Font.Size = 13
                objTable.Cell(5, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                '4. Write 4: Ngày tháng năm sinh, nơi sinh của cán bộ
                _sVal = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                _sVal = " 4) Ngày sinh " & _sVal.Substring(0, 2) & " tháng " & _sVal.Substring(3, 2) & " năm " & _sVal.Substring(6)
                objTable.Cell(6, 2).Range.InsertAfter(_sVal)
                objTable.Cell(6, 2).Range.Paragraphs.LeftIndent = 3.0F
                objTable.Cell(6, 2).Range.Font.Size = 13
                objTable.Cell(6, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                objTable.Columns(1).PreferredWidth = 21
                objTable.Columns(2).PreferredWidth = 79
                objTable.LeftPadding = 0 '-3
                objTable.RightPadding = 0 '-4

                For i As Integer = 1 To 6
                    objTable.Cell(i, 2).Range.Font.Name = "Times New Roman"
                Next
                ''Tạo một bảng có 1 dòng và 2 cột
                'objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 1, 2)
                'objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                'objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                'objTable.Range.Paragraphs.SpaceBefore = 2
                'objTable.Range.Paragraphs.SpaceAfter = 2

                'If dr("AnhThe").ToString().Trim() = "" Then
                '    _sVal = "Ảnh 4*6"
                '    objselection.TypeText(_sVal)
                'Else
                '    Dim oPic As Word.InlineShape
                '    oPic = objselection.InlineShapes.AddPicture(FileName:=String.Format("{0}{1}", Application.StartupPath, dr("AnhThe").ToString().Trim()), _
                '    LinkToFile:=False, SaveWithDocument:=True)
                '    oPic.Width = 80
                '    oPic.Height = 160
                '    objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                'End If
                'objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                'objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

                ''Tạo một Paragraph mới
                ''objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                'objselection.InsertParagraph()
                'objselection.TypeParagraph()
                'objselection.Font.Size = 12
                'objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                'objselection.Paragraphs.SpaceAfter = 4
                'objselection.TypeText(" 1) Họ và tên khai sinh: ")
                'objselection.Font.Italic = 1
                '_sVal = dr("HoTen").ToString().Trim().ToUpper() & vbTab & vbTab & vbTab & vbTab
                'objselection.TypeText(_sVal & vbTab & vbTab & vbTab)
                'objselection.Font.Italic = 0

                'objselection.TypeText(" Nam, Nữ: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(IIf(dr("GioiTinh").ToString() = False, " Nam", " Nữ") + vbCrLf)
                'objselection.Font.Italic = 0

                'objselection.TypeText(" 3) Các tên gọi khác: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(dr("TenThuongGoi").ToString().Trim() + vbCrLf)
                'objselection.Font.Italic = 0

                ''Write 4: Chức vụ (Đảng, đoàn thể, chính quyền, kể cả chức vụ kiêm nhiệm)
                ''--->Lấy chức vụ đảng trong thời gian hiện tại (Tức chức vụ đó thì ngày hiện tại phải thuộc khoảng từ ngày 1 đến ngày 2)
                'Dim arrPartys() As String       'Lấy mảng các thông tin về đảng của cán bộ
                'arrPartys = GetPartys(_RowId)
                ''Lấy Chức vụ đảng của cán bộ (Chức vụ chính và Chức vụ Kiêm)
                'If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
                '    _sVal = arrPartys(0) + "; " + arrPartys(4)
                'Else
                '    _sVal = IIf(arrPartys(0) <> "", arrPartys(0), IIf(arrPartys(4) <> "", arrPartys(4), ""))
                'End If
                'Dim _sDoanvien As String = GetValue(_RowId, 1) 'Lấy chuỗi thông tin đoàn
                'If (_sDoanvien <> "" And _sDoanvien.IndexOf(";") >= 0) Then
                '    If (_sVal <> "") Then
                '        _sVal += "; " + _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
                '    Else
                '        _sVal = _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
                '    End If
                'End If
                'If (_sVal <> "") Then
                '    _sVal += "; " + _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                'Else
                '    _sVal = _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                'End If
                '_sVal = _sVal.Replace(";;", ";").Replace("; ;", ";").Trim()
                'While _sVal.EndsWith("-") 'Bỏ dấu gạch ngang (-) ở cuối chuỗi nếu có
                '    _sVal = _sVal.Substring(0, _sVal.Length - 1).Trim()
                'End While
                'If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                '    _sVal = _sVal.Substring(1).Trim()
                'End If
                'While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
                '    _sVal = _sVal.Substring(0, _sVal.Length - 1)
                'End While
                'If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
                '    _sVal += vbTab & vbTab & vbTab
                'Else
                '    _sVal += vbTab
                'End If

                'objselection.TypeText(" 3) Cấp uỷ hiện tại: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(arrPartys(1) & vbTab & vbTab & vbTab)
                'objselection.Font.Italic = 0
                'objselection.TypeText(", Cấp uỷ kiêm: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(arrPartys(5) & vbCrLf)
                'objselection.Font.Italic = 0

                'objselection.TypeText("  Chức vụ: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(_sVal & vbCrLf)
                'objselection.Font.Italic = 0

                'objselection.TypeText("  Phụ cấp chức vụ: ")
                'objselection.Font.Italic = 1
                '_sVal = GetValue(_RowId, 2)
                '_sVal = _sVal.Substring(0, _sVal.IndexOf(";"))
                'objselection.TypeText(_sVal & vbCrLf)
                'objselection.Font.Italic = 0

                ''4. Write 4: Ngày tháng năm sinh, nơi sinh của cán bộ
                '_sVal = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                'objselection.TypeText(" 4) Sinh ngày: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(_sVal.Substring(0, 2) & " tháng " & _sVal.Substring(3, 2) & " năm " & _sVal.Substring(6) & vbTab & vbTab)
                'objselection.Font.Italic = 0
                '_sVal &= _HS_CanBo.GetAddress(CType(IIf(dr("IdNS_Tinh").ToString() <> "", dr("IdNS_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Huyen").ToString() <> "", dr("IdNS_Huyen").ToString(), "0"), Int32), dr("NS_DChi").ToString().Trim())
                'objselection.TypeText(" 5) Nơi sinh: ")
                'objselection.Font.Italic = 1
                'objselection.TypeText(_sVal & vbTab)
                'objselection.Font.Italic = 0

                'objTable.Columns(1).PreferredWidth = 18
                'objTable.Columns(2).PreferredWidth = 82
                'objTable.LeftPadding = 0' -3
                'objTable.RightPadding = 0' -4

                'Tạo một Paragraph mới
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                'objselection.TypeParagraph()
                'objselection.Font.Size = 13
                'objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                'objselection.Paragraphs.SpaceAfter = 4

                '_sVal = " 5) Nơi sinh: " & _HS_CanBo.GetAddress(CType(IIf(dr("IdNS_Tinh").ToString() <> "", dr("IdNS_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Huyen").ToString() <> "", dr("IdNS_Huyen").ToString(), "0"), Int32), dr("NS_DChi").ToString().Trim())
                'objTable.Cell(7, 2).Range.InsertAfter(_sVal)
                'objTable.Cell(7, 2).Range.Paragraphs.LeftIndent = 3.0F
                'objTable.Cell(7, 2).Range.Font.Size = 13
                'objTable.Cell(7, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                '_sVal = " 6) Quê quán: " & _HS_CanBo.GetAddress(CType(IIf(dr("IdNQ_Tinh").ToString() <> "", dr("IdNQ_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNQ_Huyen").ToString() <> "", dr("IdNQ_Huyen").ToString(), "0"), Int32), dr("NQ_DChi").ToString().Trim())
                'objTable.Cell(8, 2).Range.InsertAfter(_sVal)
                'objTable.Cell(8, 2).Range.Paragraphs.LeftIndent = 3.0F
                'objTable.Cell(8, 2).Range.Font.Size = 13
                'objTable.Cell(8, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                objselection.TypeText("5) Nơi sinh: ")
                objselection.Font.Italic = 1
                _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdNS_Tinh").ToString() <> "", dr("IdNS_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Xa").ToString() <> "", dr("IdNS_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Thon").ToString() <> "", dr("IdNS_Thon").ToString(), "0"), Int32), dr("NS_DChi").ToString().Trim())
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...................") & vbCrLf)
                objselection.Font.Italic = 0

                objselection.TypeText("6) Quê quán: ")
                objselection.Font.Italic = 1
                _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdNQ_Tinh").ToString() <> "", dr("IdNQ_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNQ_Xa").ToString() <> "", dr("IdNQ_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdNQ_Thon").ToString() <> "", dr("IdNQ_Thon").ToString(), "0"), Int32), dr("NQ_DChi").ToString().Trim())
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...................") & vbCrLf)
                objselection.Font.Italic = 0

                objselection.TypeText("7) Nơi ở hiện nay: ")
                objselection.Font.Italic = 1
                _sVal = _HS_CanBo.GetAddress(CType(IIf(dr("IdTTr_Tinh").ToString() <> "", dr("IdTTr_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdTTr_Xa").ToString() <> "", dr("IdTTr_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdTTr_Thon").ToString() <> "", dr("IdTTr_Thon").ToString(), "0"), Int32), dr("TTr_Diachi").ToString().Trim())
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...................") & vbCrLf)
                objselection.Font.Italic = 0

                objselection.TypeText(vbTab & " Điện thoại: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(dr("DienThoai_DD").ToString().Trim() <> "", dr("DienThoai_DD").ToString().Trim(), "...................") & vbCrLf)
                'objselection.Font.Size = 12
                objselection.Font.Italic = 0

                objselection.TypeText("8) Dân tộc: ")
                objselection.Font.Italic = 1
                _sVal = IIf(dr("IdDanToc").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 21 and Status = 1", CType(dr("IdDanToc").ToString(), Int32))), "")
                objselection.TypeText(_sVal & vbTab & vbTab & vbTab & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("9) Tôn giáo: ")
                objselection.Font.Italic = 1
                _sVal = IIf(dr("IdTonGiao").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 24 and Status = 1", CType(dr("IdTonGiao").ToString(), Int32))), "")
                objselection.TypeText(_sVal & vbCrLf)
                objselection.Font.Italic = 0

                objselection.TypeText("10) Thành phần gia đình xuất thân: ")
                objselection.Font.Italic = 1
                _sVal = IIf(dr("IdThanhPhanGD").ToString() <> "", _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 13 and Status = 1", CType(dr("IdThanhPhanGD").ToString(), Integer))), "")
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...................") & vbCrLf)
                objselection.Font.Italic = 0

                '11. Nghề nghiệp bản thân trước khi được tuyển dụng (ghi nghề được đào tạo hoặc công nhân (thợ gì), làm ruộng, buôn bán, học sinh,...)
                objselection.TypeText("11) Nghề nghiệp bản thân trước khi được tuyển dụng: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(dr("GhiChu").ToString().Trim <> "", dr("GhiChu").ToString(), "...................") + vbCrLf)
                objselection.Font.Italic = 0

                '12. Nghề nghiệp bản thân trước khi được tuyển dụng (ghi nghề được đào tạo hoặc công nhân (thợ gì), làm ruộng, buôn bán, học sinh,...)
                objselection.TypeText("12) Ngày được tuyển dụng: ")
                objselection.Font.Italic = 1
                If dr("Ngay_ThamNien").ToString().Trim <> "" Then
                    _sVal = IIf(CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy"), sNonDate)
                Else
                    _sVal = sNonDate
                End If
                objselection.TypeText(_sVal + vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("Vào cơ quan nào, ở đâu: ")
                _sVal = GetValue(_RowId, 9)
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "................") & vbCrLf)
                objselection.Font.Italic = 0

                '13. Ngay vao co quan hien dang cong tac
                objselection.TypeText("13) Ngày vào cơ quan hiện đang công tác: ")
                objselection.Font.Italic = 1
                _sVal = GetValue(_RowId, 10)
                objselection.TypeText(IIf(_sVal <> "", _sVal, sNonDate) & vbTab & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("Ngày tham gia cách mạng: ")
                objselection.Font.Italic = 1
                If dr("CM_Ngay").ToString().Trim <> "" Then
                    _sVal = IIf(CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd-MM-yyyy"), sNonDate)
                Else
                    _sVal = sNonDate
                End If
                objselection.TypeText(_sVal + vbCrLf)
                objselection.Font.Italic = 0

                '14. Ngay vao Dang CSVN
                objselection.TypeText("14) Ngày Đảng Cộng sản Việt Nam: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(arrPartys(2) <> "", arrPartys(2), sNonDate) & vbTab & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("Ngày chính thức: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(arrPartys(3) <> "", arrPartys(3), sNonDate) + vbCrLf)
                objselection.Font.Italic = 0

                '15. Ngày tham gia các tổ chức chính trị, xã hội (Ngày vào Đoàn TNCSHCM, Công đoàn, Hội,...)
                objselection.TypeText("15) Ngày tham gia các tổ chức chính trị, xã hội ")
                objselection.Font.Size = 10
                objselection.TypeText("(Ngày vào Đoàn TNCSHCM, Công đoàn, Hội,...)")
                objselection.Font.Size = 13
                objselection.TypeText(": ")
                _sVal = ""
                If (_sDoanvien <> "") Then
                    If (_sDoanvien.IndexOf(";") >= 0) Then
                        _sVal = _sDoanvien.Substring(0, _sDoanvien.IndexOf(";"))
                    End If
                End If
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, sNonDate) & vbCrLf)
                objselection.Font.Italic = 0
                _sVal = ""

                '16. Ngày nhập ngũ, Ngày xuất ngũ, Quân hàm, chức vụ cao nhất (Năm)
                strSQL = " SELECT a.IdHSLLVT,a.TuNgay,a.DenNgay,a.IdQuanHam, b.Ten_Goi As QuanHam,a.ChucVu,a.DonVi FROM HS_LLVT a, DanhMuc b "
                strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdQuanHam = b.id and b.Status = 1 and b.id_goc = 2) Order by a.TuNgay ASC"
                Dim ngay_nhap As String = ""
                Dim ngay_xuat As String = ""
                Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (_db Is Nothing) Then
                        If (_db.Rows.Count > 0) Then
                            Dim _IdMin As Integer = CType(_db.Rows(0)("IdQuanHam").ToString().Trim(), Integer)
                            Dim _ChuvuQN As String = _db.Rows(0)("ChucVu").ToString().Trim()
                            ngay_nhap = CType(_db.Rows(0)("TuNgay").ToString(), DateTime).ToString("dd/MM/yyyy")
                            For i As Integer = 0 To _db.Rows.Count - 1
                                If (CType(_db.Rows(i)("IdQuanHam").ToString().Trim(), Integer) < _IdMin) Then
                                    _IdMin = CType(_db.Rows(i)("IdQuanHam").ToString().Trim(), Integer)
                                    _ChuvuQN = _db.Rows(i)("ChucVu").ToString().Trim()
                                End If

                                If (i = _db.Rows.Count - 1) Then
                                    ngay_xuat = CType(_db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy")
                                End If
                            Next
                            _sVal = _HS_CanBo.GetNameByCode(String.Format("Select id, Ten_Goi from DanhMuc Where id = {0} And id_goc = 2 and Status = 1", _IdMin)) + ", " + _ChuvuQN
                        End If
                    End If
                End Using
                objselection.TypeText("16) Ngày nhập ngũ: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(ngay_nhap <> "", ngay_nhap, "..../..../......."))
                objselection.Font.Italic = 0

                objselection.TypeText(",  Ngày xuất ngũ: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(ngay_xuat <> "", ngay_xuat, "..../..../......."))
                objselection.Font.Italic = 0

                objselection.TypeText(", Quân hàm, chức vụ cao nhất: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "....................") + vbCrLf)
                objselection.Font.Italic = 0

                '17. Trình độ học vấn: Giáo dục phổ thông (Lớp mấy):
                _sVal = IIf(dr("IdTrinhDoVH").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 6 and Status = 1", CType(dr("IdTrinhDoVH").ToString(), Int32))), "")
                objselection.TypeText("17) Trình độ học vấn: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "..........."))
                objselection.Font.Italic = 0

                objselection.TypeText("  Học hàm, học vị cao nhất: ")
                objselection.Font.Italic = 1
                _sVal = GetValue(_RowId, 11)
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0

                objselection.TypeText(vbTab & "- Lý luận chính trị: ")
                objselection.Font.Italic = 1
                _sVal = IIf(dr("IdTrinhDoCT").ToString() <> "0", _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 36 and Status = 1", CType(dr("IdTrinhDoCT").ToString(), Int32))), "")
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("- Ngoại ngữ: ")
                objselection.Font.Italic = 1
                _sVal = GetValue(_RowId, 7)
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("- Tin học: ")
                objselection.Font.Italic = 1
                _sVal = GetValue(_RowId, 8)
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0

                '18. Công việc chính đang làm:
                _sVal = IIf(_IdChmon > 0, _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 12 and Status = 1", _IdChmon)), "")
                objselection.TypeText("18) Công tác chính đang làm: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0

                '19. Ngach luong:
                Dim vNgachLuong As String = ""
                _sVal = IIf(_IdChmon > 0, _HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 12 and Status = 1", _IdChmon)), "")
                objselection.TypeText("19) Ngạch công chức: ")
                objselection.Font.Italic = 1
                strSQL = "SELECT (Select Mota from NgachLuong Where IdNgachLuong = b.IdNgachLuong And Status = 1) As NgachLuong,"
                strSQL += "a.IdLuongCB,b.BacLuong,a.HeSoLuong,Convert(Varchar, Ngay_Huong, 103) as NgayApDung  From HS_LuongCB a, BacLuong b"
                strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdBacLuong = b.IdBacLuong And b.Status = 1) Order by a.Ngay_Huong Desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            vNgachLuong = vbTab + "Bậc lương: " + db.Rows(0)("BacLuong").ToString().Trim()
                            vNgachLuong &= vbTab + "Hệ số: " + CType(db.Rows(0)("HeSoLuong"), Double).ToString("N2")
                            objselection.TypeText(db.Rows(0)("NgachLuong").ToString().Trim() & vbCrLf)
                            objselection.TypeText(vbTab + "Bậc lương: " + db.Rows(0)("BacLuong").ToString().Trim())
                            objselection.TypeText(vbTab + "Hệ số: " + CType(db.Rows(0)("HeSoLuong"), Double).ToString("N2"))
                            If (db.Rows(0)("NgayApDung").ToString().Trim() <> "") Then
                                Dim arrDate() As String = db.Rows(0)("NgayApDung").ToString().Trim().Split("/")
                                vNgachLuong &= vbTab & "Hưởng từ " + arrDate(0) + " tháng " + arrDate(1) + " năm " + arrDate(2)
                                objselection.TypeText(vbTab & "Hưởng từ " + arrDate(0) + " tháng " + arrDate(1) + " năm " + arrDate(2) + vbCrLf)
                            Else
                                objselection.TypeText(vbTab & "Hưởng từ    tháng    năm    " + vbCrLf)
                            End If
                        Else
                            objselection.TypeText("19) Ngạch công chức, viên chức:              Bậc lương:      Hệ số:      Hưởng từ    tháng    năm" + vbCrLf)
                        End If
                    Else
                        objselection.TypeText("19) Ngạch công chức, viên chức:              Bậc lương:      Hệ số:      Hưởng từ    tháng    năm" + vbCrLf)
                    End If
                End Using
                objselection.Font.Italic = 0

                '20. Danh hieu duoc phong:
                _sVal = GetValue(_RowId, 12)
                objselection.TypeText("20) Danh hiệu được phong: ")
                objselection.Font.Size = 10
                objselection.TypeText("(Anh hùng, nhà giáo,...)")
                objselection.Font.Size = 13
                objselection.TypeText(": ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0

                '21. Danh hieu duoc phong:
                objselection.TypeText("21) Sở trường công tác: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(dr("SoTruong_CT").ToString().Trim() <> "", dr("SoTruong_CT").ToString().Trim(), "...........") & vbTab & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("Công việc đã làm lâu nhất: ")
                objselection.Font.Italic = 1
                _sVal = GetValue(_RowId, 8)
                objselection.TypeText(IIf(dr("CV_Lau").ToString().Trim() <> "", dr("CV_Lau").ToString().Trim(), "...........") & vbCrLf)
                objselection.Font.Italic = 0

                '22. Huan huy chuong duoc phong:
                _sVal = GetValue(_RowId, 13)
                objselection.TypeText("22) Khen thưởng ")
                objselection.Font.Size = 10
                objselection.TypeText("(Huân, huy chương,...)")
                objselection.Font.Size = 13
                objselection.TypeText(": ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0

                '23. Ki luat
                _sVal = GetValue(_RowId, 14)
                objselection.TypeText("23) Kỉ luật ")
                objselection.Font.Size = 10
                objselection.TypeText("(Đảng, Chính quyền, Đoàn thể,...)")
                objselection.Font.Size = 13
                objselection.TypeText(": ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0

                '24. Lấy dữ liệu tình trạng Sức khoẻ của cán bộ trong hồ sơ Sức khoẻ của Cán bộ
                Dim arrTemp() As String = GetValue(_RowId, 3).Split(";")
                objselection.TypeText("24) Tình trạng sức khoẻ: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(arrTemp(2).Trim() <> "", arrTemp(2).Trim(), ".................") & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText(" Cao: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(arrTemp(0).Trim() <> "", arrTemp(0).Trim(), ".......") & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText(" Cân Nặng: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(arrTemp(1).Trim() <> "", arrTemp(1).Trim(), ".......") & " (kg)" & vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText(" Nhóm máu: ")
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(dr("NhomMau").ToString().Trim() <> "", dr("NhomMau").ToString().Trim(), ".......") & vbCrLf)
                objselection.Font.Italic = 0

                '25. CMT
                _sVal = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                objselection.TypeText("25) Số chứng minh nhân dân: ")
                objselection.Font.Italic = 1
                objselection.TypeText(dr("CMT_So").ToString().Trim() & vbTab)
                objselection.Font.Italic = 0
                objselection.TypeText(vbTab & " Ngày cấp: ")
                objselection.Font.Italic = 1
                objselection.TypeText(_sVal & vbTab)
                objselection.Font.Italic = 0
                objselection.TypeText(" Nơi cấp: ")
                objselection.Font.Italic = 1
                objselection.TypeText(dr("CMT_NoiCap").ToString().Trim() & vbCrLf)
                objselection.Font.Italic = 0

                'objselection.TypeText(" Thương binh loại: ")
                If (dr("IdUT_BThan").ToString().Trim() <> "" And dr("IdUT_BThan").ToString().Trim() <> "0") Then
                    _sVal = dr("IdUT_BThan").ToString().Trim()
                    If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                        _sVal = _sVal.Substring(1).Trim()
                    End If
                    While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
                        _sVal = _sVal.Substring(0, _sVal.Length - 1)
                    End While
                    _sVal = _sVal.Trim()
                    Dim arrUTBT() As String = _sVal.Split(";")
                    Dim lengths As Integer = arrUTBT.Length - 1
                    If lengths > 0 Then
                        For i As Integer = 0 To lengths
                            _sVal = IIf(arrUTBT(i).ToString() <> "", _HS_CanBo.GetNameByCode(String.Format("Select id, Ten_Goi from DanhMuc Where id = {0} And id_goc = 18 and ma_so in('1801', '1802', '1803', '1804') and Status = 1", CType(arrUTBT(i).ToString(), Integer))), "")
                            If _sVal <> "" Then
                                objselection.TypeText(vbTab & _sVal & vbTab)
                            End If
                        Next
                    Else
                        objselection.TypeText(vbTab & "Thương binh loại:........" & vbTab)
                    End If
                Else
                    objselection.TypeText(vbTab & "Thương binh loại:........" & vbTab)
                End If

                objselection.TypeText(" Gia đình liệt sĩ: ")
                _sVal = IIf(dr("IdUT_GDinh").ToString() <> "", _HS_CanBo.GetNameByCode(String.Format("Select id, ma_so from DanhMuc Where id = {0} And id_goc = 19 and Status = 1", CType(dr("IdUT_GDinh").ToString(), Integer))), "")
                If (String.Compare(_sVal, "1902", True) = 0) Then
                    objselection.InsertSymbol(254, "Wingdings")
                Else
                    objselection.InsertSymbol(168, "Wingdings")
                End If
                objselection.TypeText(vbCrLf)

                objselection.InsertBreak() '------------- trang mới - LHN

                '26. Xuất dữ liệu: Quá trình đào tạo, Bồi dưỡng về chuyên môn, nghiệp vụ, lý luận chính trị, ngoại ngữ, tin học
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.TypeText("26) ĐÀO TẠO, BỔI DƯỠNG VỀ CHUYÊN MÔN, NGHIỆP VỤ, LÝ LUẬN CHÍNH TRỊ, NGOẠI NGỮ" + vbCrLf)
                objselection.Font.Size = 11
                Dim _rowsTable As Integer = 2
                strSQL = "SELECT CoSo_DT,(Select Ten_Goi From DanhMuc Where IdChuyenNganhDT = Id and id_goc = 11 And Status =1) As ChuyenNganh,"
                strSQL += "NganhHoc,(Convert(Varchar, TuNgay, 103) + ' - ' + Convert(Varchar, DenNgay, 103)) As ThoiGian,"
                strSQL += "(Select Ten_Goi From DanhMuc Where id = IdHinhThucDT And id_goc = 8 And Status = 1) As HinhThuc,"
                strSQL += "(CASE VBCC WHEN 1 THEN 'VB'ELSE 'CC' END) As VBCC,"
                strSQL += "(Select Ten_Goi From DanhMuc Where id = IdTrinhDo And id_goc = 38 And Status = 1) As TrinhDo"
                strSQL += " FROM HS_DTVBCC WHERE IdCanBo = '" + _RowId + "' And HoanThanh = 1 Order By DenNgay Asc"
                Using db_vb As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_vb Is Nothing Or db_vb.Rows.Count = 0) Then
                        _rowsTable = 2
                    Else
                        _rowsTable = db_vb.Rows.Count + 1
                    End If
                    'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 5 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Tên trường")
                    objTable.Cell(1, 2).Range.InsertAfter("Chuyên ngành - Khoa")
                    objTable.Cell(1, 3).Range.InsertAfter("Thời gian học")
                    objTable.Cell(1, 4).Range.InsertAfter("Hình thức học")
                    objTable.Cell(1, 5).Range.InsertAfter("Văn bằng, chứng chỉ, trình độ gì")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    If Not (db_vb Is Nothing) Then
                        If (db_vb.Rows.Count > 0) Then
                            For i As Integer = 0 To db_vb.Rows.Count - 1
                                objTable.Cell(i + 2, 1).Range.InsertAfter(db_vb.Rows(i)("CoSo_DT").ToString().Trim())
                                objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 2).Range.InsertAfter(db_vb.Rows(i)("ChuyenNganh").ToString().Trim() + " - " + db_vb.Rows(i)("NganhHoc").ToString().Trim())
                                objTable.Cell(i + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 3).Range.InsertAfter(db_vb.Rows(i)("ThoiGian").ToString().Trim())
                                objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 4).Range.InsertAfter(db_vb.Rows(i)("HinhThuc").ToString().Trim())
                                objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 5).Range.InsertAfter(db_vb.Rows(i)("VBCC").ToString().Trim() + "-" + db_vb.Rows(i)("TrinhDo").ToString().Trim())
                                objTable.Cell(i + 2, 5).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 27
                    objTable.Columns(2).PreferredWidth = 27
                    objTable.Columns(3).PreferredWidth = 23
                    objTable.Columns(4).PreferredWidth = 13
                    objTable.Columns(5).PreferredWidth = 17
                    objTable.LeftPadding = 0 '-3
                    objTable.RightPadding = 0 '-4
                End Using

                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Font.Size = 10
                objselection.Range.Paragraphs.SpaceAfter = 4
                _sVal = "Ghi chú: Hình thức học: Chính quy, tại chức, chuyên tu, bồi dưỡng,.../ Văn bằng: Tiến sỹ, Phó TS, Thạc sỹ, Cử nhân, kỹ sư,..." + vbCrLf
                objselection.TypeText(_sVal)

                '27. Tóm tắt quá trình công tác:
                objselection.Paragraphs.SpaceBefore = 10
                'objselection.Font.Bold = 1
                objselection.Font.Size = 12
                objselection.TypeText("27) TÓM TẮT QUÁ TRÌNH CÔNG TÁC" + vbCrLf)
                'objselection.Font.Bold = 0
                objselection.Font.Size = 11
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                Using db_ct As DataTable = listQDNhansu(_RowId, 2)
                    If (db_ct Is Nothing Or db_ct.Rows.Count = 0) Then
                        _rowsTable = 2
                    Else
                        _rowsTable = db_ct.Rows.Count + 1
                    End If
                    'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 2 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 2)
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.PreferredWidth = 100
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Từ ngày tháng năm -" + vbCrLf + " Đến ngày tháng năm")
                    objTable.Cell(1, 2).Range.InsertAfter("Chức danh, chức vụ, đơn vị công tác (Đảng, Chính quyền, Đoàn thể)")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    If Not (db_ct Is Nothing) Then
                        If (db_ct.Rows.Count > 0) Then
                            Dim row As Integer = 0
                            For i As Integer = 0 To db_ct.Rows.Count - 1
                                Dim ThoiGian As String = ""
                                If CInt(db_ct.Rows(i)("IsQD_NHCS")) = 0 Then
                                    If i = 0 Then
                                        row = i
                                        ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        Dim j As Integer = i + 1
                                        While j < db_ct.Rows.Count
                                            Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    Exit While
                                                Case Else
                                                    j = j + 1
                                            End Select
                                        End While
                                        objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                        objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                        objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                        objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                    Else
                                        If Not (db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
                                                Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
                                                Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
                                            ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                            If Not (db_ct.Rows(i)("DenNgay") Is DBNull.Value) Then
                                                If Not (CType(db_ct.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(db_ct.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                                    row = row + 1
                                                    objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                                    objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                                    objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                                    objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                                Else

                                                    Dim j As Integer = i + 1
                                                    While j < db_ct.Rows.Count
                                                        Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                                Exit While
                                                            Case Else
                                                                j = j + 1
                                                        End Select
                                                    End While
                                                    row = row + 1
                                                    objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                                    objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                                    objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                                    objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F

                                                End If
                                            Else

                                                Dim j As Integer = i + 1
                                                While j < db_ct.Rows.Count
                                                    Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            Exit While
                                                        Case Else
                                                            j = j + 1
                                                    End Select
                                                End While
                                                row = row + 1
                                                objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                                objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                                objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                                objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                            End If
                                        End If
                                    End If
                                Else
                                    'Đọc các QĐ khác QĐ thôi kiêm nhiệm
                                    If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() <> "1519" Then
                                        ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        Dim j As Integer = i + 1
                                        While j < db_ct.Rows.Count
                                            Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                Case "1534", "1535"
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    Exit While
                                                Case "1518"
                                                    'Bỏ qua nếu QĐ tiếp theo là QĐ kiêm nhiệm
                                                    Exit While
                                                Case "1519"
                                                    'Lấy mốc hết thời gian kiêm nhiệm
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    j = j + 1
                                                    Exit While
                                                Case Else
                                                    'TH không đọc tiếp nếu: Nếu đơn vị, phòng ban, chức vụ cũ; hoặc QĐ tiếp theo là QĐ kiêm nhiệm
                                                    If (CInt(db_ct.Rows(i)("IDDonVi_Moi")) = CInt(db_ct.Rows(j)("IDDonVi_Moi")) And CInt(db_ct.Rows(i)("IdPhong_Moi")) = CInt(db_ct.Rows(j)("IdPhong_Moi")) And CInt(db_ct.Rows(i)("IDChucVu_Moi")) = CInt(db_ct.Rows(j)("IDChucVu_Moi"))) Then
                                                        If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1515" Then 'Thử việc
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            Exit While
                                                        Else
                                                            j = j + 1
                                                        End If
                                                    Else
                                                        'Nếu QĐ trước đó la QD Kiêm nhiệm mà QĐ tiếp theo ko phải là QĐ Thôi kiêm nhiệm thì bỏ qua
                                                        If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                            Exit While
                                                        Else
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            i = j - 1
                                                            Exit While
                                                        End If

                                                    End If
                                            End Select
                                        End While
                                        If i <> 0 Then row = row + 1
                                        objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                        objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                        objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1534" Then
                                            objTable.Cell(row + 2, 2).Range.InsertAfter(db_ct.Rows(i)("LoaiQD").ToString())
                                        Else
                                            If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                objTable.Cell(row + 2, 2).Range.InsertAfter(db_ct.Rows(i)("LoaiQD").ToString() & " " & IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("(CN cấp I & tương đương)", ""), ""))
                                            Else
                                                ' objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", ""), ""))

                                                If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1515" Then 'Thử việc
                                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", "") & " (" & db_ct.Rows(i)("LoaiQD").ToString() & ")", ""))
                                                    'dgv_congtac.Rows(row).Cells("cln_QuaTrinh").Value = IIf(db.Rows(i)("NoiDung").ToString() <> "", db.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", "") & " (" & db.Rows(i)("LoaiQD").ToString() & ")", "")
                                                Else
                                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", ""), ""))
                                                End If

                                            End If

                                        End If
                                        objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                        If j = db_ct.Rows.Count Then Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 24
                    objTable.Columns(2).PreferredWidth = 76
                    objTable.LeftPadding = 0 '-3
                    objTable.RightPadding = 0 '-4
                End Using

                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Font.Name = "Times New Roman"
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Paragraphs.SpaceAfter = 4
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.TypeText("28) ĐẶC ĐIỂM LỊCH SỬ BẢN THÂN: " + vbCrLf)
                objselection.Font.Size = 11
                objselection.Font.Italic = 1
                If (dr("DacDiem_BT").ToString().Trim() <> "") Then
                    objselection.TypeText(vbTab + "- " + dr("DacDiem_BT").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
                Else
                    _sVal = "      a - Khai rõ: bị bắt, bị tù (từ ngày tháng năm nào đến ngày tháng năm nào, ở đâu), đã khai báo cho ai, những vấn đề gì:" + vbCrLf
                    _sVal += "     ....................................................................................................................................................................." + vbCrLf
                    _sVal += "     ....................................................................................................................................................................." + vbCrLf
                    _sVal += "     b - Bản thân có làm việc trong chế độ cũ (Cơ quan, đơn vị nào, địa điểm, chức danh, chức vụ, thời gian làm việc):" + vbCrLf
                    _sVal += "     ....................................................................................................................................................................." + vbCrLf
                    _sVal += "     ....................................................................................................................................................................." + vbCrLf
                    objselection.TypeText(_sVal)
                End If
                objselection.Font.Italic = 0

                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.TypeText("29) QUAN HỆ VỚI NƯỚC NGOÀI: " + vbCrLf)
                objselection.Font.Size = 11
                objselection.Font.Italic = 1
                If (dr("QuanHe_Nguoi_NN").ToString().Trim() <> "") Then
                    objselection.TypeText(vbTab + "- " + dr("QuanHe_Nguoi_NN").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
                Else
                    _sVal = "      - Tham gia hoặc có quan hệ với các tổ chức chính trị, kinh tế, xã hội nào ở nước ngoài (làm gì, tổ chức nào, đặt trụ sở ở đâu?):" + vbCrLf
                    _sVal += "     ......................................................................................................................................................................" + vbCrLf
                    _sVal += "     ......................................................................................................................................................................" + vbCrLf
                    _sVal += "     ......................................................................................................................................................................" + vbCrLf
                    _sVal += "     - Có thân nhân (Bố mẹ, vợ, chồng, con, anh chị em ruột) ở nước ngoài (làm gì, địa chỉ?): ..............................." + vbCrLf
                    _sVal += "     ......................................................................................................................................................................" + vbCrLf
                    _sVal += "     ......................................................................................................................................................................" + vbCrLf
                    _sVal += "     ......................................................................................................................................................................" + vbCrLf
                    objselection.TypeText(_sVal)
                End If
                objselection.Font.Italic = 0

                objselection.InsertBreak() '------------- trang mới - LHN

                '30. Quan hệ gia đình
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Name = "Times New Roman"
                objselection.Font.Size = 12
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.TypeText("30. QUAN HỆ GIA ĐÌNH:" + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Italic = 1
                objselection.TypeText("a) Về bản thân: Bố, Mẹ, Vợ (Chồng), các con, anh chị em ruột" + vbCrLf)
                objselection.Font.Size = 11
                objselection.Font.Italic = 0
                'Tạo một bảng có _rowsTable dòng và 4 cột
                strSQL = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2303','2304','2314','2313','2312','2317','2305','2320','2315','2316'))"
                Using db_rows As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_rows Is Nothing Or CType(db_rows.Rows(0)(0), Byte) = 0) Then
                        _rowsTable = 6
                    Else
                        Dim db As DBAccess = New DBAccess
                        Dim ssql As String = ""
                        'Dem So Bo khác bố đẻ
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so in ('2318', '2332'))"
                        If db.getNumber(ssql) = 0 Then
                            '_rowsTable = 4
                            _rowsTable = 3
                        Else
                            _rowsTable = 3 + db.getNumber(ssql)
                        End If
                        'Dem So Me khác mẹ đẻ
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so in ('2319', '2333'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        'Dem so Vo/Chong
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2314','2313'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        'Dem So con
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2312','2317'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        'If db.getNumber(ssql) = 0 Then
                        '    _rowsTable = 5
                        'Else
                        '    _rowsTable = 4 + db.getNumber(ssql)
                        'End If
                        'Dem So anh chi em
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        '_rowsTable = CType(db_rows.Rows(0)(0), Byte) + 1
                        If _rowsTable <= 6 Then _rowsTable = 6
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
                    objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
                    objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
                    objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
                    objTable.Cell(1, 4).Range.Font.Bold = 0
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(2, 1).Range.InsertAfter("Bố")
                    objTable.Cell(2, 1).Range.Font.Bold = 1
                    objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
                    'objTable.Cell(3, 1).Range.Font.Bold = 1
                    'objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'If (dr("GioiTinh").ToString() = False) Then
                    '    objTable.Cell(4, 1).Range.InsertAfter("Vợ" + vbCrLf)
                    'Else
                    '    objTable.Cell(4, 1).Range.InsertAfter("Chồng" + vbCrLf)
                    'End If
                    'objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'objTable.Cell(4, 1).Range.Font.Bold = 1

                    For i As Byte = 2 To _rowsTable
                        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                    Next

                    If Not (db_rows Is Nothing) Then
                        If (db_rows.Rows.Count > 0) Then
                            'Lấy thông tin Bố đẻ
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2303')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(2, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            Dim _rowGD As Integer = 2

                            'Lấy thông tin Dượng
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2318')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Bố dượng")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Bố nuôi
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2332')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Bố nuôi")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Mẹ đẻ
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2304')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Mẹ")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Mẹ kế
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2319')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Mẹ kế")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Mẹ nuôi
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2333')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Mẹ kế")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Vợ hoặc chồng
                            Dim rowCVC As Integer = 4  'Danh dau dong Chong/Vo/Con
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2314','2313'))"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        If (dr("GioiTinh").ToString() = False) Then
                                            objTable.Cell(_rowGD, 1).Range.InsertAfter("Vợ" + vbCrLf)
                                        Else
                                            objTable.Cell(_rowGD, 1).Range.InsertAfter("Chồng" + vbCrLf)
                                        End If
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    Else
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(vbCrLf)
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter(vbCrLf)
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Các con đẻ
                            Dim _RowCon As Byte = 0
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2312','2317')) order by NamSinh"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _RowCon = db_gd.Rows.Count
                                        _rowGD = _rowGD + 1
                                        For i As Integer = 0 To db_gd.Rows.Count - 1
                                            If (db_gd.Rows.Count = 1) Then
                                                objTable.Cell(i + _rowGD, 1).Range.InsertAfter("Con")
                                            Else
                                                If (i = 0) Then objTable.Cell(i + _rowGD, 1).Range.InsertAfter("Các" + vbCrLf + "con")
                                            End If
                                            objTable.Cell(i + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            objTable.Cell(i + _rowGD, 1).Range.Font.Bold = 1
                                            objTable.Cell(i + _rowGD, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _rowGD, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                                            objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                            If (db_gd.Rows(i)("ConMat") = False) Then
                                                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                                                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                                            Else
                                                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Đã mất")
                                            End If

                                            If (i > 0) Then objTable.Rows(i + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                        Next
                                    Else
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(vbCrLf)
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter(vbCrLf)
                                    End If
                                End If
                            End Using

                            ''Lấy thông tin Các con đẻ
                            'Dim _RowCon As Byte = 0
                            'strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2312','2317')) order by NamSinh"
                            'Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            '    If db_gd.Rows.Count = 0 Then
                            '        _rowGD = _rowGD + 1
                            '        objTable.Cell(_rowGD, 1).Range.InsertAfter("Con")
                            '        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                            '        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            '    End If
                            '    If Not (db_gd Is Nothing) Then
                            '        If (db_gd.Rows.Count > 0) Then
                            '            _RowCon = db_gd.Rows.Count
                            '            _rowGD = _rowGD + 1
                            '            For i As Integer = 0 To db_gd.Rows.Count - 1
                            '                If (i = 0) Then
                            '                    objTable.Cell(i + _rowGD, 1).Range.InsertAfter("Các" + vbCrLf + "con")
                            '                End If
                            '                objTable.Cell(i + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            '                objTable.Cell(i + _rowGD, 1).Range.Font.Bold = 1
                            '                objTable.Cell(i + _rowGD, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                            '                objTable.Cell(i + _rowGD, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                            '                objTable.Cell(i + _rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            '                _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                            '                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                            '                If (db_gd.Rows(0)("ConMat") = False) Then
                            '                    objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                            '                    objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                            '                Else
                            '                    objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Đã mất")
                            '                End If

                            '                If (i > 0) Then objTable.Rows(i + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                            '            Next
                            '        End If
                            '    End If
                            '    '                                If _RowCon = db_gd.Rows.Count Then objTable.Rows(_RowCon - 1 + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                            'End Using

                            'Lấy thông tin Anh chị em ruột thịt
                            Dim _rowAnhChi As Integer = 0
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316')) order by NamSinh"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If db_gd.Rows.Count = 0 Then
                                    If _RowCon = 0 Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(6, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                                        objTable.Cell(6, 1).Range.Font.Bold = 1
                                        objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                    Else
                                        objTable.Cell(_RowCon + _rowGD, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                                        objTable.Cell(_RowCon + _rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_RowCon + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                    End If
                                End If

                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowAnhChi = db_gd.Rows.Count
                                        If _RowCon = 0 Then _RowCon = 1
                                        For i As Integer = 0 To db_gd.Rows.Count - 1
                                            If (i = 0) Then
                                                objTable.Cell(i + _RowCon + _rowGD, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                                                objTable.Rows(i + _RowCon + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                            End If
                                            objTable.Cell(i + _RowCon + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            objTable.Cell(i + _RowCon + _rowGD, 1).Range.Font.Bold = 1
                                            objTable.Cell(i + _RowCon + _rowGD, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _RowCon + _rowGD, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _RowCon + _rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                                            objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                            If (db_gd.Rows(i)("ConMat") = False) Then
                                                objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                                                objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                                            Else
                                                objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Đã mất")
                                            End If

                                            'If (i > 0) Then objTable.Rows(i + _RowCon + 3).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                            If (i > 0) Then objTable.Rows(i + _RowCon + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                            'objTable.Rows(i + _RowCon + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                        Next
                                        objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                        objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                        objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                        objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    End If
                                End If
                            End Using
                            'Merge ô cuối trước
                            If (_rowAnhChi > 1) Then
                                'objTable.Cell(_RowCon + 4, 1).Merge(objTable.Cell(_rowsTable, 1))
                                objTable.Cell(_RowCon + _rowGD, 1).Merge(objTable.Cell(_rowsTable, 1))
                            End If
                            'If _RowCon > 1 Then objTable.Cell(4, 1).Merge(objTable.Cell(3 + _RowCon, 1))
                            If _RowCon > 1 Then objTable.Cell(_rowGD, 1).Merge(objTable.Cell(_rowGD + _RowCon - 1, 1))

                        End If
                    End If
                    'Thực hiện Merge Cell Anh chị em ruột
                    objTable.Columns(1).PreferredWidth = 9
                    objTable.Columns(2).PreferredWidth = 18
                    objTable.Columns(3).PreferredWidth = 9
                    objTable.Columns(4).PreferredWidth = 59
                    objTable.LeftPadding = 0 '-3
                    objTable.RightPadding = 0 '-4
                End Using

                '30 b (Bố mẹ, anh chị em ruột bên vợ hoặc chồng)
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.Font.Italic = 1
                objselection.TypeText("b) Bố, Mẹ, anh chị em ruột (bên vợ hoặc chồng): " + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 11
                objselection.Font.Italic = 0
                strSQL = "SELECT count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2326','2327','2328','2329','2330','2331')) "
                Using db_gdvochong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_gdvochong Is Nothing Or CType(db_gdvochong.Rows(0)(0), Byte) = 0) Then
                        _rowsTable = 4
                    Else
                        _rowsTable = CType(db_gdvochong.Rows(0)(0), Byte) + 3
                        If _rowsTable <= 4 Then _rowsTable = 4
                    End If
                    'Tạo một bảng có 4 dòng và 4 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
                    objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
                    objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
                    objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
                    objTable.Cell(1, 4).Range.Font.Bold = 0
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(2, 1).Range.InsertAfter("Bố")
                    objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
                    objTable.Cell(4, 1).Range.InsertAfter("Anh" + vbCrLf + "chị em" + vbCrLf + "ruột")
                    objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(2, 1).Range.Font.Bold = 1
                    objTable.Cell(3, 1).Range.Font.Bold = 1
                    objTable.Cell(4, 1).Range.Font.Bold = 1
                    For i As Integer = 2 To _rowsTable
                        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                    Next
                    'Lấy thông tin Bố vợ (Bố chồng)
                    strSQL = "SELECT Top 1 * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2322','2324'))"
                    Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db_gd Is Nothing) Then
                            If (db_gd.Rows.Count > 0) Then
                                objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                If (db_gd.Rows(0)("ConMat") = False) Then
                                    objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                    objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                Else
                                    objTable.Cell(2, 4).Range.InsertAfter("Đã mất")
                                End If
                            End If
                        End If
                    End Using
                    'Lấy thông tin Mẹ vợ (Mẹ chồng)
                    strSQL = "SELECT Top 1 * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2323','2325'))"
                    Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db_gd Is Nothing) Then
                            If (db_gd.Rows.Count > 0) Then
                                objTable.Cell(3, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(3, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                objTable.Cell(3, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                If (db_gd.Rows(0)("ConMat") = False) Then
                                    objTable.Cell(3, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                    objTable.Cell(3, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                Else
                                    objTable.Cell(3, 4).Range.InsertAfter("Đã mất")
                                End If
                            End If
                        End If
                    End Using
                    If (CType(db_gdvochong.Rows(0)(0), Byte) > 0) Then
                        'Lấy thông tin Anh chị em ruột bên vợ (hoặc bên chồng)
                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2326','2327','2328','2329','2330','2331'))  order by NamSinh"
                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db_gd Is Nothing) Then
                                If (db_gd.Rows.Count > 0) Then
                                    For i As Integer = 0 To db_gd.Rows.Count - 1
                                        objTable.Cell(i + 4, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(i + 4, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(i + 4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        If (db_gd.Rows(i)("ConMat") = False) Then
                                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Đã mất")
                                        End If
                                        If (i > 0) Then
                                            objTable.Rows(i + 4).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                        End If
                                    Next
                                    'objTable.Cell(_rowsTable, 1).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    'Thực hiện Merge Cell Anh chị em ruột
                                    If (_rowsTable > 4) Then objTable.Cell(4, 1).Merge(objTable.Cell(_rowsTable, 1))
                                End If
                            End If
                        End Using
                    End If
                    objTable.Columns(1).PreferredWidth = 9
                    objTable.Columns(2).PreferredWidth = 18
                    objTable.Columns(3).PreferredWidth = 9
                    objTable.Columns(4).PreferredWidth = 59
                    objTable.LeftPadding = 0 '-3
                    objTable.RightPadding = 0 '-4
                End Using

                '31. Hoan canh kinh te gia dinh
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Name = "Times New Roman"
                objselection.Font.Size = 12
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.TypeText("31. HOÀN CẢNH KINH TẾ GIA ĐÌNH:" + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Bold = 1
                objselection.TypeText("- Quá trình lương của bản thân:" + vbCrLf)
                objselection.Font.Bold = 0
                strSQL = "SELECT (Select Mota from NgachLuong Where IdNgachLuong = b.IdNgachLuong And Status = 1) As NgachLuong,"
                strSQL += " a.IdLuongCB,b.BacLuong As BacLuong,a.HeSoLuong As HeSo,Convert(Varchar, Ngay_Huong, 103) as NgayHuong  From HS_LuongCB a, BacLuong b"
                strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdBacLuong = b.IdBacLuong And b.Status = 1) Order by a.Ngay_Huong "

                'strSQL = "SELECT (Select BacLuong From BacLuong Where IdBacLuong = HS_LuongCB.IdBacLuong) As BacLuong,"
                'strSQL += "(Select Mota From NgachLuong Where IdNgachLuong in (Select IdNgachLuong From BacLuong Where IdBacLuong = HS_LuongCB.IdBacLuong)) As NgachLuong,"
                'strSQL += "CONVERT(VarChar,HeSoLuong) As HeSo,Convert(Varchar,Ngay_Huong,103) As NgayHuong,SoQD,Convert(Varchar,NgayQD,103) As NgayQD,NguoiQD"
                'strSQL += " FROM HS_LuongCB WHERE IdCanBo = '" + _RowId + "' Order By HeSoLuong Asc"
                Using db_qtluong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_qtluong Is Nothing Or db_qtluong.Rows.Count = 0) Then
                        _rowsTable = 10
                    Else
                        _rowsTable = 6
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 4, _rowsTable)
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    For i As Integer = 1 To 4
                        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                    Next
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Tháng/Năm")
                    objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    objTable.Cell(2, 1).Range.InsertAfter("Ngạch")
                    objTable.Cell(3, 1).Range.InsertAfter("Bậc lương")
                    objTable.Cell(4, 1).Range.InsertAfter("Hệ số lương")
                    If Not (db_qtluong Is Nothing) Then
                        If (db_qtluong.Rows.Count > 0) Then
                            Dim _columns As Byte = 6  'Dùng biến này làm biến đếm ngược
                            For i As Integer = db_qtluong.Rows.Count - 1 To 0 Step -1
                                objTable.Cell(1, _columns).Range.InsertAfter(db_qtluong.Rows(i)("NgayHuong").ToString().Trim().Substring(3))
                                objTable.Cell(2, _columns).Range.InsertAfter(db_qtluong.Rows(i)("NgachLuong").ToString().Trim())
                                objTable.Cell(3, _columns).Range.InsertAfter(db_qtluong.Rows(i)("BacLuong").ToString().Trim())
                                objTable.Cell(4, _columns).Range.InsertAfter(db_qtluong.Rows(i)("HeSo").ToString().Trim())
                                _columns -= 1
                                If _columns = 1 Then
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    objTable.LeftPadding = 0 '-3
                    objTable.RightPadding = 0 '-4
                End Using

                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 12
                objselection.Font.Bold = 1
                objselection.TypeText("- Nguồn thu nhập chính của gia đình (hàng năm):" + vbCrLf)
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                Dim vLuong As String = "......................................................................................................................."
                Dim vCacNguonKhac As String = "............................................................................................................"
                Dim vNhaO_Cap As String = "..................."
                Dim vNhaO_Cap_DT As String = ".............."
                Dim vNhaO_Mua As String = "..................."
                Dim vNhaO_Mua_DT As String = "................"
                Dim vDatO_Cap As String = "....................."
                Dim vDatO_Mua As String = "....................."
                Dim vDat_KD As String = ".........................................." + vbCrLf
                vDat_KD += ".................................................................................................................................................................."
                'If vNgachLuong <> "" Then vLuong = vNgachLuong
                'strSQL = " SELECT TOP 1 * FROM HS_ThuNhapGD WHERE IDCanBo='" + _RowId + "' and year(NgayKeKhai)= year(getdate()) order by NgayKeKhai desc"
                strSQL = " SELECT TOP 1 * FROM HS_ThuNhapGD WHERE IDCanBo='" + _RowId + "' order by NgayKeKhai desc"
                Using dtTNGD As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (dtTNGD.Rows.Count > 0) Then
                        vLuong = dtTNGD.Rows(0)("Luong").ToString().Trim()
                        vCacNguonKhac = dtTNGD.Rows(0)("NguonKhac").ToString().Trim()
                        vNhaO_Cap = dtTNGD.Rows(0)("NhaO_DuocCap").ToString().Trim()
                        vNhaO_Cap_DT = dtTNGD.Rows(0)("NhaO_DuocCap_DT").ToString().Trim()
                        vNhaO_Mua = dtTNGD.Rows(0)("NhaO_TuMua").ToString().Trim()
                        vNhaO_Mua_DT = dtTNGD.Rows(0)("NhaO_TuMua_DT").ToString().Trim()
                        vDatO_Cap = dtTNGD.Rows(0)("DatO_DuocCap_DT").ToString().Trim()
                        vDatO_Mua = dtTNGD.Rows(0)("DatO_TuMua_DT").ToString().Trim()
                        vDat_KD = dtTNGD.Rows(0)("DatSXKT").ToString().Trim()
                    End If
                End Using
                objselection.TypeText(vbTab + vbTab + "+ Lương: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vLuong & vbCrLf)
                objselection.Font.Italic = 0
                objselection.TypeText(vbTab + vbTab + "+ Các nguồn khác: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vCacNguonKhac & vbCrLf)
                objselection.Font.Italic = 0

                objselection.Font.Bold = 1
                objselection.TypeText("- Nhà ở:")
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText(vbTab + "+ Được cấp, được thuê, loại nhà: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Cap)
                objselection.Font.Italic = 0
                objselection.TypeText(", tổng diện tích sử dụng: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Cap_DT + "m2" + vbCrLf)
                objselection.Font.Italic = 0
                objselection.TypeText(vbTab + vbTab + "+ Nhà tự mua, tự xây, loại nhà: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Mua)
                objselection.Font.Italic = 0
                objselection.TypeText(", tổng diện tích sử dụng: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Mua_DT + "m2" + vbCrLf)
                objselection.Font.Italic = 0

                objselection.Font.Bold = 1
                objselection.TypeText("- Đất ở:")
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText(vbTab + "+ Đất được cấp: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vDatO_Cap)
                objselection.Font.Italic = 0
                objselection.TypeText("m2," + vbTab + vbTab + vbTab + "+ Đất tự mua: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vDatO_Mua + "m2" + vbCrLf)

                objselection.Font.Bold = 1
                objselection.TypeText("- Đất sản xuất, kinh doanh ")
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText("(Tổng diện tích được cấp, tự mua, tự khai phá,...): ")
                objselection.Font.Italic = 1
                objselection.TypeText(vDat_KD)
                objselection.Font.Italic = 0

                objselection.TypeText(vbCrLf + vbCrLf + vbCrLf + vbCrLf)

                objselection.Font.Size = 13
                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 5, 2)
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.PreferredWidth = 100
                objTable.Range.Paragraphs.SpaceBefore = 0
                objTable.Range.Paragraphs.SpaceAfter = 0

                'objTable.Range.Font.Bold = 0
                objTable.Cell(1, 1).Range.InsertAfter("Người khai")
                objTable.Cell(1, 1).Range.Font.Bold = 1
                objTable.Cell(2, 1).Range.InsertAfter("Tôi xin cam đoan những")
                objTable.Cell(3, 1).Range.InsertAfter("lời khai trên đây là đúng sự thật")
                objTable.Cell(4, 1).Range.InsertAfter("(Ký tên)")
                objTable.Cell(5, 1).Range.InsertAfter(vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + dr("HoTen").ToString().Trim().ToUpper())
                objTable.Cell(5, 1).Range.Font.Bold = 1
                _sVal = _sDiaban + ", Ngày " + DateTime.Now.ToString("dd") + " tháng " + DateTime.Now.ToString("MM") + " năm " + DateTime.Now.ToString("yyyy")
                objTable.Cell(2, 2).Range.InsertAfter(_sVal)
                objTable.Cell(2, 2).Range.Font.Italic = 1
                objTable.Cell(3, 2).Range.InsertAfter("Xác nhận của cơ quan quản lý")
                objTable.Cell(3, 2).Range.Font.Bold = 1
                objTable.Cell(5, 2).Range.InsertAfter(vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf)
                objTable.Cell(1, 1).Merge(objTable.Cell(4, 1))
                objTable.Cell(1, 2).Merge(objTable.Cell(4, 2))
                objTable.LeftPadding = 0 '-3
                objTable.RightPadding = 0 '-4


                'Định dạng thêm về trang word khi xuất ra
                objWordApp.ActiveDocument.PageSetup.HeaderDistance = 3
                objWordApp.ActiveDocument.PageSetup.FooterDistance = 3
                objWordApp.ActiveDocument.PageSetup.LeftMargin = 25         'Quy ước:   72 Points = 1 Inch.
                objWordApp.ActiveDocument.PageSetup.RightMargin = 20
                objWordApp.ActiveDocument.PageSetup.TopMargin = 25
                objWordApp.ActiveDocument.PageSetup.BottomMargin = 25
                objWordApp.ActiveDocument.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4
                'objWord.ActiveDocument.PageSetup.HeaderDistance = 100
                objWordApp.ActiveDocument.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait   'Định dạng Khổ dấy dọc
                objWordApp.ActiveDocument.ShowGrammaticalErrors = False     'Không cho hiển thị các đường viền check chính tả mầu xanh
                objWordApp.ActiveDocument.ShowSpellingErrors = False        'Không cho hiển thị các đường viền check chính tả mầu đỏ

                ''Danh so trang
                'objWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
                'objWordApp.Selection.TypeParagraph()
                'Dim docNumber As String = "1"
                'Dim revisionNumber As String = "0"
                'objWordApp.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                'objWordApp.ActiveWindow.Selection.Font.Name = "Arial"
                'objWordApp.ActiveWindow.Selection.Font.Size = 8
                'objWordApp.ActiveWindow.Selection.TypeText("Document #: " + docNumber + " - Revision #: " + revisionNumber)
                ''INSERTING TAB CHARACTERS 
                'objWordApp.ActiveWindow.Selection.TypeText("" & Chr(9) & "")
                'objWordApp.ActiveWindow.Selection.TypeText("" & Chr(9) & "")
                'objWordApp.ActiveWindow.Selection.TypeText("Page ")
                'Dim CurrentPage As Object = Word.WdFieldType.wdFieldPage
                'objWordApp.ActiveWindow.Selection.Fields.Add(objWordApp.Selection.Range, CurrentPage, oMissing, oMissing)
                'objWordApp.ActiveWindow.Selection.TypeText(" of ")
                'Dim TotalPages As Object = Word.WdFieldType.wdFieldNumPages
                'objWordApp.ActiveWindow.Selection.Fields.Add(objWordApp.Selection.Range, TotalPages, oMissing, oMissing)
                ''SETTING FOCUES BACK TO DOCUMENT 
                'objWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument

                'Dim sFileName As String = SaveToFile("Document File|*.doc|XML Document File|*.docx", "Mau2C_" + dr("MaCB").ToString().Trim().ToUpper() + "_" + getFullName(_RowId.ToString().Trim(), False))
                'objDocument.SaveAs(sFileName)
                'If sFileName <> "" Then
                '    objDocument.SaveAs(sFileName)
                'End If

                SaveToWord(objDocument, "Mau2C_" + dr("MaCB").ToString().Trim().ToUpper() + "_" + getFullName(_RowId.ToString().Trim(), False), True)

                Clipboard.Clear()
                objWordApp.Quit(False)
                objWordApp = Nothing
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xuất dữ liệu Hồ sơ cán bộ thành mẫu Mẫu 01/LLCB (Lý lịch cá nhân) ra file Word
    ''' </summary>
    ''' <param name="_RowId">Chuỗi chỉ số xác định bản ghi của cán bộ</param>
    ''' <remarks></remarks>
    Public Sub ExportWord_CurriculumVitae(ByVal _RowId As String)
        Dim _sDiaban As String = DIABAN
        Dim _sDonviTT As String = ""        'Biến lưu chuỗi Đơn vị Trực thuộc
        Dim _sDonviCS As String = ""        'Biến lưu chuỗi Đơn vị Cơ sở
        Dim _IdDvParent As Integer = 0          'Biến lưu Chỉ số đơn vị cấp nhỏ nhất của Cán bộ
        Dim _IdDvChild As Integer = 0           'Biến lưu Chỉ số đơn vị cấp nhỏ nhất của Cán bộ
        Dim _IdPb As Integer = 0                'Biến lưu chỉ số Phòng ban của cán bộ
        Dim _IdCv As Integer = 0                'Biến lưu chỉ số xác định chức vụ Cán bộ
        Dim _IdChmon As Integer = 0             'Biến lưu chỉ số xác Chuyên môn mới nhất của cán bộ (Công việc chính đang làm)
        Dim _sVal As String = ""
        Dim strSQL As String = ""
        Dim HS_CanBo As clsHS_CanBo = New clsHS_CanBo()

        Dim dr As DataRow
        dr = HS_CanBo.GetHuman_ForCode(_RowId)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                'Lấy các thông tin chính để xuất dữ liệu phần đầu
                _IdDvParent = dr("IdDonVi").ToString().Trim()
                'Lấy dữ liệu Đơn vị cấp dưới của cán bộ đang công tác, Phòng ban, chức vụ
                Dim HS_QDNhansu As QDNhanSu = New QDNhanSu
                Dim dbconn As DBAccess = New DBAccess
                HS_QDNhansu = HS_QDNhansu.getFinalRecord(_RowId)
                If Not (HS_QDNhansu Is Nothing) Then
                    If (HS_QDNhansu.IdQDNhanSu <> "") Then
                        _IdCv = IIf(HS_QDNhansu.IdChucvu_Moi <> 0, HS_QDNhansu.IdChucvu_Moi, 0)
                        _IdPb = IIf(HS_QDNhansu.IdPhong_Moi <> 0, HS_QDNhansu.IdPhong_Moi, 0)
                        _IdChmon = IIf(HS_QDNhansu.IdChuyenMon_Moi <> 0, HS_QDNhansu.IdChuyenMon_Moi, 0)
                        _IdDvChild = IIf(HS_QDNhansu.IdDonvi_Moi <> 0, HS_QDNhansu.IdDonvi_Moi, 0)
                    End If
                End If
                If dbconn.getNumber("SELECT count(*) FROM Phongban WHERE ma_so in ('01','02','03') and [id]=" & _IdPb) > 0 Then
                    _sDonviTT = "Đơn vị trực thuộc: Ngân hàng Chính sách xã hội Việt Nam"
                Else
                    strSQL = String.Format("Select id, Ten_Goi from ChiNhanh Where Status = 1 and id = {0} And id_goc IN (0,1)", _IdDvParent)
                    _sDonviTT = "Đơn vị trực thuộc: " + Replace_Branch(HS_CanBo.GetNameByCode(strSQL).ToString())
                End If

                strSQL = String.Format("Select id, Ten_Goi from ChiNhanh Where Status = 1 and id = {0} And id_goc IN (0,1)", _IdDvParent)
                _sDonviTT = "Đơn vị trực thuộc: " + Replace_Branch(HS_CanBo.GetNameByCode(strSQL).ToString())
                If (_IdDvParent <> _IdDvChild) Then 'Phòng giao dịch
                    strSQL = String.Format("Select id, Ten_Goi from ChiNhanh Where Status = 1 and id = {0} And id_goc > 1", _IdDvChild)
                    _sDonviCS = "Đơn vị cơ sở: " + Replace_Branch(HS_CanBo.GetNameByCode(strSQL).ToString())
                ElseIf (_IdDvParent = _IdDvChild) Then 'Phòng ban của tỉnh hoặc Hội sở
                    strSQL = String.Format("Select id, ten_phong From PhongBan Where Status = 1 And id = {0}", _IdPb)
                    _sDonviCS = "Đơn vị cơ sở: " + Replace_Branch(HS_CanBo.GetNameByCode(strSQL).ToString())
                End If

                'Bắt đầu khai báo và xuất dữ liệu ra Word
                Dim objWordApp As New Word.Application      'Tạo Đối tượng Word Application
                Dim objDocument As New Word.Document        'Tạo đối tượng Word Document
                Dim objTable As Word.Table                  'Tạo đối tượng Word Table
                'Start Word and open the document template.
                'objWordApp.Visible = True                          'And show word screen (False: Hide msword cho đến khi xuất hết thông tin)
                'objWordApp.Activate()
                objDocument = objWordApp.Documents.Add                  'Add một Document vào trong Application (Create a new word document)
                'Biến lưu vị trí select hiện hành
                Dim objselection As Word.Selection
                'Gán vị trí hiện hành trong Document vào biến selection
                objselection = objDocument.Application.Selection()

                'Định dạng Paragraph
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Size = 13
                objselection.Font.Name = "Times New Roman"
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify

                'Tạo một bảng gồm 3 dòng và 3 cột
                'objTable = objselection.Tables.Add(objDocument.Bookmarks.Item("\endofdoc").Range, 3, 3)
                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 3, 3)
                objTable.Columns(1).Width = 180.0F
                objTable.Columns(2).Width = 162.0F
                objTable.Columns(3).Width = 100.0F
                objTable.Cell(1, 2).Merge(objTable.Cell(3, 2))
                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.LeftPadding = 0 ' -10
                objTable.RightPadding = 0 ' -10
                objTable.Cell(1, 2).Range.InsertAfter("LÝ LỊCH CÁ NHÂN")
                objTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                objTable.Cell(1, 2).Range.Font.Name = "Times New Roman"
                objTable.Cell(1, 2).Range.Font.Size = 20
                objTable.Cell(1, 2).Range.Font.Bold = 1

                objTable.Cell(1, 1).Range.InsertAfter("Tỉnh (TP): " + _sDiaban)
                objTable.Cell(2, 1).Range.InsertAfter(_sDonviTT)
                objTable.Cell(3, 1).Range.InsertAfter(_sDonviCS)
                objTable.Cell(1, 3).Range.InsertAfter("Mẫu 01/LLCB")
                objTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                objTable.Cell(3, 3).Range.InsertAfter(dr("MaCB").ToString().Trim().ToUpper())
                objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.InsertParagraph()

                'Xóa định dạng Paragraph trước
                objselection.ClearFormatting()
                'Định dạng lại Paragraph
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Size = 12
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                'Add Text vào Paragraph
                objselection.TypeText("Số hiệu cán bộ, công chức")
                objselection.Paragraphs.SpaceAfter = 14

                'BẮT ĐẦU XUẤT DỮ LIỆU CHÍNH CỦA HỒ SƠ CÁN BỘ
                'Tạo một Paragraph mới
                objselection.TypeParagraph()
                objselection.Font.Size = 12
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.Paragraphs.SpaceAfter = 4
                objselection.TypeText("1. Họ và tên ")
                objselection.Font.Size = 10
                objselection.Font.Italic = 1
                objselection.TypeText("(Viết chữ in hoa): ")
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                _sVal = dr("HoTen").ToString().Trim().ToUpper() & vbTab & vbTab & vbTab & vbTab & "2. Nam, Nữ: " & IIf(dr("GioiTinh").ToString() = False, " Nam", " Nữ")
                objselection.TypeText(_sVal + vbCrLf)
                objselection.TypeText("3. Các tên gọi khác: " + dr("TenThuongGoi").ToString().Trim() + vbCrLf)

                '4.5.6. Write 4: Chức vụ (Đảng, đoàn thể, chính quyền, kể cả chức vụ kiêm nhiệm)
                '--->Lấy chức vụ đảng trong thời gian hiện tại (Tức chức vụ đó thì ngày hiện tại phải thuộc khoảng từ ngày 1 đến ngày 2)
                Dim arrPartys() As String       'Lấy mảng các thông tin về đảng của cán bộ
                arrPartys = GetPartys(_RowId)
                'Lấy Chức vụ đảng của cán bộ (Chức vụ chính và Chức vụ Kiêm)
                If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
                    _sVal = arrPartys(0) + "; " + arrPartys(4)
                Else
                    _sVal = IIf(arrPartys(0) <> "", arrPartys(0), IIf(arrPartys(4) <> "", arrPartys(4), ""))
                End If
                Dim _sDoanvien As String = GetValue(_RowId, 1) 'Lấy chuỗi thông tin đoàn
                If (_sDoanvien <> "" And _sDoanvien.IndexOf(";") >= 0) Then
                    If (_sVal <> "") Then
                        _sVal += "; " + _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
                    Else
                        _sVal = _sDoanvien.Substring(_sDoanvien.IndexOf(";") + 1)
                    End If
                End If

                If (_sVal <> "") Then
                    If _IdPb = 3 Then
                        _sVal += "; " & HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " Ngân hàng Chính sách xã hội"
                    Else
                        If _IdPb = 15 Then
                            _sVal += "; " & HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " " & HS_CanBo.GetNameByCode(String.Format("Select id,Ten_Goi from Chinhanh Where id = {0} ", _IdDvChild))
                        Else
                            _sVal += "; " & HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                        End If
                    End If
                Else
                    If _IdPb = 3 Then
                        _sVal = HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " Ngân hàng Chính sách xã hội"
                    Else
                        If _IdPb = 15 Then
                            _sVal = HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim() & " " & HS_CanBo.GetNameByCode(String.Format("Select id,Ten_Goi from Chinhanh Where id = {0} ", _IdDvChild))
                        Else
                            _sVal = HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 14 and Status = 1", _IdCv)).Trim().Replace("(Ban) Hội sở chính", "").Trim() + " - " + _sDonviCS.Substring(13).Trim()
                        End If
                    End If
                End If

                _sVal = _sVal.Replace(";;", ";").Replace("; ;", ";").Trim()
                While _sVal.EndsWith("-") 'Bỏ dấu gạch ngang (-) ở cuối chuỗi nếu có
                    _sVal = _sVal.Substring(0, _sVal.Length - 1).Trim()
                End While
                If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                    _sVal = _sVal.Substring(1).Trim()
                End If
                While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
                    _sVal = _sVal.Substring(0, _sVal.Length - 1)
                End While
                If (arrPartys(0) <> "" And arrPartys(4) <> "") Then
                    _sVal += vbTab & vbTab & vbTab
                Else
                    _sVal += vbTab
                End If
                objselection.TypeText("4. Chức vụ ")
                'Xóa định dạng Paragraph cũ
                objselection.Font.Size = 10
                objselection.Font.Italic = 1
                objselection.TypeText("(Đảng, đoàn thể, chính quyền, kể cả chức vụ kiêm nghiệm): ")
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.TypeText(_sVal)

                'Lấy phụ cấp chức vụ và Phụ cấp trách nhiệm
                _sVal = GetValue(_RowId, 2)
                _sVal = "5. Phụ cấp chức vụ: " + _sVal.Substring(0, _sVal.IndexOf(";")) + vbTab + vbTab + "6. Phụ cấp trách nhiệm: " + _sVal.Substring(_sVal.IndexOf(";") + 1) + vbCrLf
                objselection.TypeText(_sVal)
                '7. Write 7: Cấp uỷ hiện tai, cấp uỷ kiêm
                _sVal = "7. Cấp uỷ hiện tại: " + arrPartys(1) & vbTab & vbTab & vbTab + ", Cấp uỷ kiêm: " + arrPartys(5) + vbCrLf
                objselection.TypeText(_sVal)

                '8. Write 8: Ngày tháng năm sinh, nơi sinh của cán bộ
                _sVal = IIf(CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                objselection.TypeText("8. Ngày sinh " & _sVal.Substring(0, 2) & " tháng " & _sVal.Substring(3, 2) & " năm " & _sVal.Substring(6) & vbTab & vbTab & "9. Nơi sinh: ")
                _sVal = HS_CanBo.GetAddress(CType(IIf(dr("IdNS_Tinh").ToString() <> "", dr("IdNS_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Xa").ToString() <> "", dr("IdNS_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdNS_Thon").ToString() <> "", dr("IdNS_Thon").ToString(), "0"), Int32), dr("NS_DChi").ToString().Trim())
                objselection.TypeText(_sVal & vbCrLf)
                _sVal = IIf(dr("IdQuocTich").ToString() <> "0", HS_CanBo.GetNameByCode(String.Format("Select id,Ten_Goi from QuocGia Where id = {0} and Status = 1", CType(dr("IdQuocTich").ToString(), Int32))), "")
                objselection.TypeText("10. Quốc tịch: " & _sVal & vbTab & vbTab)

                _sVal = IIf(dr("IdDanToc").ToString() <> "0", HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 21 and Status = 1", CType(dr("IdDanToc").ToString(), Int32))), "")
                objselection.TypeText("11. Dân tộc: " & _sVal & vbTab & vbTab)
                _sVal = IIf(dr("IdTonGiao").ToString() <> "0", HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 24 and Status = 1", CType(dr("IdTonGiao").ToString(), Int32))), "")
                objselection.TypeText("12. Tôn giáo: " & _sVal & vbTab & vbTab & vbCrLf)
                _sVal = HS_CanBo.GetAddress(CType(IIf(dr("IdNQ_Tinh").ToString() <> "", dr("IdNQ_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdNQ_Xa").ToString() <> "", dr("IdNQ_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdNQ_Thon").ToString() <> "", dr("IdNQ_Thon").ToString(), "0"), Int32), dr("NQ_DChi").ToString().Trim())
                objselection.TypeText("13. Quê quán: " & _sVal & vbCrLf)
                _sVal = HS_CanBo.GetAddress(CType(IIf(dr("IdThT_Tinh").ToString() <> "", dr("IdThT_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdThT_Xa").ToString() <> "", dr("IdThT_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdThT_Thon").ToString() <> "", dr("IdThT_Thon").ToString(), "0"), Int32), dr("ThT_Diachi").ToString().Trim())
                objselection.TypeText("14. Nơi đăng ký Hộ khẩu thường trú: " & _sVal & vbCrLf)

                _sVal = HS_CanBo.GetAddress(CType(IIf(dr("IdTTr_Tinh").ToString() <> "", dr("IdTTr_Tinh").ToString(), "0"), Int32), CType(IIf(dr("IdTTr_Xa").ToString() <> "", dr("IdTTr_Xa").ToString(), "0"), Int32), CType(IIf(dr("IdTTr_Thon").ToString() <> "", dr("IdTTr_Thon").ToString(), "0"), Int32), dr("TTr_Diachi").ToString().Trim())
                objselection.TypeText("15. Nơi ở hiện nay: " & _sVal & vbCrLf)
                _sVal = IIf(CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                objselection.TypeText("16. Giấy CMND số: " & dr("CMT_So").ToString().Trim() & vbTab & "Ngày cấp: " & _sVal & " " & vbTab & " Nơi cấp: " & dr("CMT_NoiCap").ToString().Trim() & vbCrLf)
                objselection.TypeText("17. Điện thoại: ")
                'Xóa định dạng Paragraph cũ
                objselection.Font.Italic = 1
                objselection.TypeText(" Di động: " & dr("DienThoai_DD").ToString().Trim() & vbTab & " Cơ quan: " & dr("DienThoai_CQ").ToString().Trim() & vbTab & " Nhà riêng: " & dr("DienThoai_NR").ToString().Trim())
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.Paragraphs.SpaceAfter = 4
                objselection.TypeText(vbCrLf)
                objselection.TypeText("18. Địa chỉ e-mail: " & dr("Emai").ToString().Trim() & vbCrLf)
                '19. Lấy dữ liệu tình trạng Sức khoẻ của cán bộ trong hồ sơ Sức khoẻ của Cán bộ
                Dim arrTemp() As String = GetValue(_RowId, 3).Split(";")
                objselection.TypeText("19. Tình trạng sức khoẻ: " & arrTemp(2).Trim() & "," & vbTab & " Cao: " & arrTemp(0).Trim() & "," & vbTab & " Cân Nặng: " & arrTemp(1).Trim() & " (kg)," & vbTab & " Nhóm máu: " & dr("NhomMau").ToString().Trim() & vbCrLf)
                '20. Trình độ học vấn: Giáo dục phổ thông (Lớp mấy):
                _sVal = IIf(dr("IdTrinhDoVH").ToString() <> "0", HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 6 and Status = 1", CType(dr("IdTrinhDoVH").ToString(), Int32))), "")
                objselection.TypeText("20. Trình độ học vấn: Giáo dục phổ thông (Lớp mấy): " & _sVal & vbCrLf)
                'Thực hiện lấy thông tin Học vị - Chuyên ngành đào tạo - Trường đào tạo - Năm tốt nghiệp
                objselection.TypeText("21. Trình độ chuyên môn")
                objselection.Font.Size = 10
                objselection.TypeText(" (Trung cấp, Cử nhân, Kỹ sư) ")
                objselection.Font.Size = 12
                _sVal = GetValue(_RowId, 4)  '+ ", " + vbTab + GetValue(_RowId, 5)
                objselection.Font.Italic = 1
                objselection.TypeText(_sVal & vbCrLf)
                objselection.Font.Italic = 0
                'Lấy thông tin Học hàm
                '_sVal = GetValue(_RowId, 6) & vbCrLf
                'objselection.TypeText("22. Học hàm cao nhất: " + _sVal)

                objselection.TypeText("22. Học hàm cao nhất: ")
                objselection.Font.Italic = 1
                _sVal = GetValue(_RowId, 11)
                objselection.TypeText(IIf(_sVal <> "", _sVal, "...........") & vbCrLf)
                objselection.Font.Italic = 0


                'Lấy thông tin Lý luận chính trị - Ngoại ngữ - Tin học
                _sVal = IIf(dr("IdTrinhDoCT").ToString() <> "0", HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 36 and Status = 1", CType(dr("IdTrinhDoCT").ToString(), Int32))), "")
                objselection.TypeText("23. Lý luận chính trị: " + _sVal & vbTab & vbTab + "24. Ngoại ngữ: " + GetValue(_RowId, 7) & vbTab & vbTab + "25. Tin học: " + GetValue(_RowId, 8) + vbCrLf)
                objselection.Font.Size = 10
                objselection.TypeText("(Cao cấp, Trung cấp, Sơ cấp)" + vbTab + vbTab + "(T.sĩ, ĐH, C.chỉ Anh, Nga, Pháp A/B/C,...)" + vbTab + "   (T.sĩ, ĐH, C.chỉ Anh, Nga, Pháp A/B/C,...)" + vbCrLf)
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                _sVal = IIf(_IdChmon > 0, HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} and id_goc = 12 and Status = 1", _IdChmon)), "")
                objselection.TypeText("26. Công việc chính đang làm: " + _sVal + vbCrLf)
                strSQL = "SELECT (Select Mota from NgachLuong Where IdNgachLuong = b.IdNgachLuong And Status = 1) As NgachLuong,"
                strSQL += "a.IdLuongCB,b.BacLuong,a.HeSoLuong,Convert(Varchar, Ngay_Huong, 103) as NgayApDung  From HS_LuongCB a, BacLuong b"
                strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdBacLuong = b.IdBacLuong And b.Status = 1) Order by a.Ngay_Huong Desc"
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If db.Rows.Count > 0 Then
                            objselection.TypeText("27. Ngạch công chức, viên chức: " + db.Rows(0)("NgachLuong").ToString().Trim() + vbTab + vbTab)
                            objselection.TypeText("Bậc lương: " + db.Rows(0)("BacLuong").ToString().Trim() + vbTab + vbTab)
                            objselection.TypeText("Hệ số: " + CType(db.Rows(0)("HeSoLuong"), Double).ToString("N2") + vbTab)
                            If (db.Rows(0)("NgayApDung").ToString().Trim() <> "") Then
                                Dim arrDate() As String = db.Rows(0)("NgayApDung").ToString().Trim().Split("/")
                                objselection.TypeText("từ " + arrDate(0) + " tháng " + arrDate(1) + " năm " + arrDate(2) + vbCrLf)
                            Else
                                objselection.TypeText("từ    tháng    năm    " + vbCrLf)
                            End If
                        Else
                            objselection.TypeText("27. Ngạch công chức, viên chức:              Bậc lương:      Hệ số:      từ    tháng    năm" + vbCrLf)
                        End If
                    Else
                        objselection.TypeText("27. Ngạch công chức, viên chức:              Bậc lương:      Hệ số:      từ    tháng    năm" + vbCrLf)
                    End If
                End Using
                '28. Thành phần gia đình xuất thân (công nhân, nông dân, cán bộ, công chức, trí thức, quân nhân, dân nghèo thành thị, tiểu thương, tiểu chư, tư sản,...):
                objselection.TypeText("28. Thành phần gia đình xuất thân ")
                objselection.Font.Size = 10
                objselection.TypeText("(công nhân, nông dân, cán bộ, công chức, trí thức, quân nhân, dân nghèo thành thị, tiểu thương, tiểu chư, tư sản,...)")
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                _sVal = IIf(dr("IdThanhPhanGD").ToString() <> "", HS_CanBo.GetNameByCode(String.Format("Select Id,Ten_Goi From DanhMuc Where Id = {0} And id_goc = 13 and Status = 1", CType(dr("IdThanhPhanGD").ToString(), Integer))), "")
                objselection.TypeText(": " + _sVal & vbCrLf)

                '29. Ưu tiên gia đình
                _sVal = IIf(dr("IdUT_GDinh").ToString() <> "", HS_CanBo.GetNameByCode(String.Format("Select id, ma_so from DanhMuc Where id = {0} And id_goc = 19 and Status = 1", CType(dr("IdUT_GDinh").ToString(), Integer))), "")
                objselection.TypeText("29. Ưu tiên gia đình: " & vbTab & vbTab)
                If (String.Compare(_sVal, "1901", True) = 0) Then
                    objselection.InsertSymbol(254, "Wingdings")
                Else
                    objselection.InsertSymbol(112, "Wingdings")
                End If
                objselection.Font.Italic = 1
                objselection.TypeText("- Gia đình cách mạng" & vbTab & vbTab)
                objselection.Font.Italic = 0

                If (String.Compare(_sVal, "1902", True) = 0) Then
                    objselection.InsertSymbol(254, "Wingdings")
                Else
                    objselection.InsertSymbol(112, "Wingdings")
                End If
                objselection.Font.Italic = 1
                objselection.TypeText("- Gia đình liệt sỹ" & vbTab & vbTab)
                objselection.Font.Italic = 0
                If (String.Compare(_sVal, "1903", True) = 0) Then
                    objselection.InsertSymbol(254, "Wingdings")
                Else
                    objselection.InsertSymbol(112, "Wingdings")
                End If
                objselection.Font.Italic = 1
                objselection.TypeText("- Không" & vbCrLf)
                objselection.Font.Italic = 0

                '30. Ưu tiên Bản thân cán bộ
                objselection.TypeText("30. Ưu tiên bản thân: " & vbTab & vbTab)
                If (dr("IdUT_BThan").ToString().Trim() <> "" And dr("IdUT_BThan").ToString().Trim() <> "0") Then
                    _sVal = dr("IdUT_BThan").ToString().Trim()
                    If (_sVal.Substring(0, 1) = ";") Then   '   Bỏ dấu chấm phẩy ở đầu chuỗi nếu có
                        _sVal = _sVal.Substring(1).Trim()
                    End If
                    While _sVal.EndsWith(";")       '   Bỏ dấu Chấm phẩy (;) ở cuối chuỗi nếu có
                        _sVal = _sVal.Substring(0, _sVal.Length - 1)
                    End While
                    _sVal = _sVal.Trim()
                    Dim arrUTBT() As String = _sVal.Split(";")
                    Dim lengths As Integer = arrUTBT.Length - 1
                    For i As Integer = 0 To lengths
                        _sVal = IIf(arrUTBT(i).ToString() <> "", HS_CanBo.GetNameByCode(String.Format("Select id, Ten_Goi from DanhMuc Where id = {0} And id_goc = 18 and Status = 1", CType(arrUTBT(i).ToString(), Integer))), "")
                        If (i > 1) Then
                            objselection.TypeText(vbCrLf + vbTab + vbTab + vbTab + vbTab)
                        End If
                        objselection.InsertSymbol(254, "Wingdings")
                        objselection.Font.Italic = 1
                        objselection.TypeText("- " + _sVal + vbTab + vbTab)
                        objselection.Font.Italic = 0
                    Next
                Else
                    objselection.Font.Italic = 0
                    objselection.InsertSymbol(112, "Wingdings")
                    objselection.Font.Italic = 1
                    objselection.TypeText("- Thương binh hạng ...." + vbTab + vbTab)

                    objselection.Font.Italic = 0
                    objselection.InsertSymbol(112, "Wingdings")
                    objselection.Font.Italic = 1
                    objselection.TypeText("- Bệnh binh binh hạng ...." + vbCrLf)
                    objselection.TypeText(vbTab + vbTab + vbTab + vbTab)
                    objselection.Font.Italic = 0
                    objselection.InsertSymbol(112, "Wingdings")
                    objselection.Font.Italic = 1
                    objselection.TypeText("- Anh hùng lao động" + vbTab + vbTab)
                    objselection.Font.Italic = 0
                    objselection.InsertSymbol(112, "Wingdings")
                    objselection.Font.Italic = 1
                    objselection.TypeText("- Anh hùng lực lượng vũ trang" + vbCrLf)
                    objselection.TypeText(vbTab + vbTab + vbTab + vbTab)
                    objselection.Font.Italic = 0
                    objselection.InsertSymbol(112, "Wingdings")
                    objselection.Font.Italic = 1
                    objselection.TypeText("- Con liệt sỹ" + vbTab + vbTab + vbTab + vbTab)
                    objselection.Font.Italic = 0
                    objselection.InsertSymbol(112, "Wingdings")
                    objselection.Font.Italic = 1
                    objselection.TypeText("- Không")
                End If
                objselection.Font.Italic = 0
                objselection.TypeText(vbCrLf)
                '31. Nghề nghiệp bản thân trước khi được tuyển dụng (ghi nghề được đào tạo hoặc công nhân (thợ gì), làm ruộng, buôn bán, học sinh,...)
                objselection.TypeText("31. Nghề nghiệp bản thân trước khi được tuyển dụng ")
                objselection.Font.Size = 10
                objselection.TypeText("(ghi nghề được đào tạo hoặc công nhân (thợ gì), làm ruộng, buôn bán, học sinh,...)")
                objselection.Font.Size = 12
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(dr("GhiChu").ToString().Trim <> "", dr("GhiChu").ToString(), "...................") + vbCrLf)
                objselection.Font.Italic = 0

                '_sVal = GetValue(_RowId, 9)
                'objselection.TypeText("32. Ngày được tuyển dụng: " + vbTab + vbTab + " Vào cơ quan nào, ở đâu: " + vbCrLf)

                objselection.TypeText("32. Ngày được tuyển dụng: ")
                objselection.Font.Italic = 1
                If dr("Ngay_ThamNien").ToString().Trim <> "" Then
                    _sVal = IIf(CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_ThamNien").ToString(), DateTime).ToString("dd-MM-yyyy"), vbTab + vbTab)
                Else
                    _sVal = vbTab + vbTab
                End If
                objselection.TypeText(_sVal + vbTab)
                objselection.Font.Italic = 0

                objselection.TypeText("Vào cơ quan nào, ở đâu: ")
                _sVal = GetValue(_RowId, 9)
                objselection.Font.Italic = 1
                objselection.TypeText(IIf(_sVal <> "", _sVal, "................") & vbCrLf)
                objselection.Font.Italic = 0

                objselection.TypeText("33. Ngày vào ngành Ngân hàng: ")
                objselection.Font.Italic = 1
                If dr("Ngay_NH").ToString().Trim <> "" Then
                    _sVal = IIf(CType(dr("Ngay_NH").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(dr("Ngay_NH").ToString(), DateTime).ToString("dd-MM-yyyy"), vbTab + vbTab)
                Else
                    _sVal = vbTab + vbTab
                End If
                objselection.TypeText(_sVal + vbTab)
                objselection.Font.Italic = 0
                If dr("Ngay_VBSP").ToString() <> "" Then _sVal = IIf(CType(dr("Ngay_VBSP").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("Ngay_VBSP").ToString(), DateTime).ToString("dd/MM/yyyy"), "")
                objselection.TypeText("34. Ngày vào NHCSXH: " + IIf(_sVal <> "", _sVal + vbTab + vbTab, vbTab + vbTab + vbTab + vbTab) + ", Tuyển dụng hoặc tiếp nhận: ")
                strSQL = "SELECT b.Ten_Goi,a.* FROM QDNhanSu a, DanhMuc b WHERE a.IdCanBo = '" + _RowId + "' And a.IsKiemNhiem = 0 And"
                strSQL += String.Format(" (a.IdDonvi_moi IN (Select id From ChiNhanh Where id = {0} or id_goc = {0}))", _IdDvParent)
                strSQL += " And (a.IdLoaiQD = b.id And b.id_goc = 15 And b.Status = 1 And b.ma_so IN('1508','1509')) ORDER BY NgayKy_QD ASC"
                Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (_db Is Nothing) Then
                        If (_db.Rows.Count > 0) Then
                            _sVal = _db.Rows(0)("Ten_Goi").ToString().Trim()
                        End If
                    End If
                End Using
                objselection.TypeText(_sVal + vbCrLf)
                objselection.TypeText("35. Sở trường công tác: " + dr("SoTruong_CT").ToString().Trim() + "," + vbTab + " Công việc đã làm lâu nhất: " + dr("CV_Lau").ToString().Trim() + vbCrLf)
                _sVal = ""
                If dr("CM_Ngay").ToString() <> "" Then _sVal = IIf(CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("CM_Ngay").ToString(), DateTime).ToString("dd/MM/yyyy"), "")

                objselection.TypeText("36. Ngày tham gia cách mạng: " + IIf(_sVal <> "", _sVal, vbTab + vbTab) + " Trong tổ chức nào: " + dr("CM_ToChuc").ToString().Trim() + vbCrLf)
                objselection.TypeText("37. Ngày vào Đảng Cộng sản Việt Nam: " + IIf(arrPartys(2) <> "", arrPartys(2) + ", ", vbTab + vbTab))
                objselection.TypeText(" Ngày chính thức: " + arrPartys(3) + vbCrLf)
                '38. Ngày tham gia các tổ chức chính trị, xã hội (Ngày vào Đoàn TNCSHCM, Công đoàn, Hội,...)
                objselection.TypeText("38. Ngày tham gia các tổ chức chính trị, xã hội ")
                objselection.Font.Size = 10
                objselection.TypeText("(Ngày vào Đoàn TNCSHCM, Công đoàn, Hội,...)")
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                _sVal = ""
                If (_sDoanvien <> "") Then
                    If (_sDoanvien.IndexOf(";") >= 0) Then
                        _sVal = _sDoanvien.Substring(0, _sDoanvien.IndexOf(";"))
                    End If
                End If
                objselection.TypeText(": " & _sVal & vbCrLf)
                _sVal = ""
                '39. Ngày nhập ngũ, Ngày xuất ngũ, Quân hàm, chức vụ cao nhất (Năm)
                strSQL = " SELECT a.IdHSLLVT,a.TuNgay,a.DenNgay,a.IdQuanHam, b.Ten_Goi As QuanHam,a.ChucVu,a.DonVi FROM HS_LLVT a, DanhMuc b "
                strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdQuanHam = b.id and b.Status = 1 and b.id_goc = 2) Order by a.TuNgay ASC"
                Dim ngay_nhap As String = ""
                Dim ngay_xuat As String = ""
                Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (_db Is Nothing) Then
                        If (_db.Rows.Count > 0) Then
                            Dim _IdMin As Integer = CType(_db.Rows(0)("IdQuanHam").ToString().Trim(), Integer)
                            Dim _ChuvuQN As String = _db.Rows(0)("ChucVu").ToString().Trim()
                            ngay_nhap = CType(_db.Rows(0)("TuNgay").ToString(), DateTime).ToString("dd/MM/yyyy")
                            For i As Integer = 0 To _db.Rows.Count - 1
                                If (CType(_db.Rows(i)("IdQuanHam").ToString().Trim(), Integer) < _IdMin) Then
                                    _IdMin = CType(_db.Rows(i)("IdQuanHam").ToString().Trim(), Integer)
                                    _ChuvuQN = _db.Rows(i)("ChucVu").ToString().Trim()
                                End If

                                If (i = _db.Rows.Count - 1) Then
                                    ngay_xuat = CType(_db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy")
                                End If
                            Next
                            _sVal = HS_CanBo.GetNameByCode(String.Format("Select id, Ten_Goi from DanhMuc Where id = {0} And id_goc = 2 and Status = 1", _IdMin)) + ", " + _ChuvuQN
                        End If
                    End If
                End Using
                objselection.TypeText("39. Ngày nhập ngũ: " + IIf(ngay_nhap <> "", ngay_nhap, vbTab + vbTab) + ", Ngày xuất ngũ: " + IIf(ngay_xuat <> "", ngay_xuat, vbTab + vbTab) + ", Quân hàm, chức vụ cao nhất (năm): " + _sVal + vbCrLf)

                '40. Xuất dữ liệu: Quá trình đào tạo, Bồi dưỡng về chuyên môn, nghiệp vụ, lý luận chính trị, ngoại ngữ, tin học
                objselection.Font.Bold = 1
                objselection.TypeText("40. Quá trình đào tạo, Bồi dưỡng về chuyên môn, nghiệp vụ, lý luận chính trị, ngoại ngữ, tin học" + vbCrLf)
                objselection.Font.Bold = 0
                objselection.Font.Size = 11
                Dim _rowsTable As Integer = 2
                strSQL = "SELECT CoSo_DT,(Select Ten_Goi From DanhMuc Where IdChuyenNganhDT = Id and id_goc = 11 And Status =1) As ChuyenNganh,"
                strSQL += "NganhHoc,(Convert(Varchar, TuNgay, 103) + ' - ' + Convert(Varchar, DenNgay, 103)) As ThoiGian,"
                strSQL += "(Select Ten_Goi From DanhMuc Where id = IdHinhThucDT And id_goc = 8 And Status = 1) As HinhThuc,"
                strSQL += "(CASE VBCC WHEN 1 THEN 'VB'ELSE 'CC' END) As VBCC,"
                strSQL += "(Select Ten_Goi From DanhMuc Where id = IdTrinhDo And id_goc = 38 And Status = 1) As TrinhDo"
                strSQL += " FROM HS_DTVBCC WHERE IdCanBo = '" + _RowId + "' And HoanThanh = 1 Order By DenNgay Asc"
                Using db_vb As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_vb Is Nothing Or db_vb.Rows.Count = 0) Then
                        _rowsTable = 2
                    Else
                        _rowsTable = db_vb.Rows.Count + 1
                    End If
                    'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 5 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Tên trường")
                    objTable.Cell(1, 2).Range.InsertAfter("Chuyên ngành - Khoa")
                    objTable.Cell(1, 3).Range.InsertAfter("Thời gian học")
                    objTable.Cell(1, 4).Range.InsertAfter("Hình thức học")
                    objTable.Cell(1, 5).Range.InsertAfter("Văn bằng, chứng chỉ, trình độ gì")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    If Not (db_vb Is Nothing) Then
                        If (db_vb.Rows.Count > 0) Then
                            For i As Integer = 0 To db_vb.Rows.Count - 1
                                objTable.Cell(i + 2, 1).Range.InsertAfter(db_vb.Rows(i)("CoSo_DT").ToString().Trim())
                                objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 2).Range.InsertAfter(db_vb.Rows(i)("ChuyenNganh").ToString().Trim() + " - " + db_vb.Rows(i)("NganhHoc").ToString().Trim())
                                objTable.Cell(i + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 3).Range.InsertAfter(db_vb.Rows(i)("ThoiGian").ToString().Trim())
                                objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 4).Range.InsertAfter(db_vb.Rows(i)("HinhThuc").ToString().Trim())
                                objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 5).Range.InsertAfter(db_vb.Rows(i)("VBCC").ToString().Trim() + "-" + db_vb.Rows(i)("TrinhDo").ToString().Trim())
                                objTable.Cell(i + 2, 5).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 23
                    objTable.Columns(2).PreferredWidth = 27
                    objTable.Columns(3).PreferredWidth = 27
                    objTable.Columns(4).PreferredWidth = 13
                    objTable.Columns(5).PreferredWidth = 17
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Font.Size = 10
                objselection.Range.Paragraphs.SpaceAfter = 4
                _sVal = "Ghi chú: Hình thức học: Chính quy, tại chức, chuyên tu, bồi dưỡng,.../ Văn bằng: Tiến sỹ, Phó TS, Thạc sỹ, Cử nhân, kỹ sư,..." + vbCrLf
                objselection.TypeText(_sVal)

                objselection.Font.Size = 12
                '41. Tóm tắt quá trình công tác:
                objselection.Font.Bold = 1
                objselection.TypeText("41. Tóm tắt quá trình công tác:" + vbCrLf)
                objselection.Font.Bold = 0
                objselection.Font.Size = 11
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                Using db_ct As DataTable = listQDNhansu(_RowId, 2)
                    If (db_ct Is Nothing Or db_ct.Rows.Count = 0) Then
                        _rowsTable = 2
                    Else
                        _rowsTable = db_ct.Rows.Count + 1
                    End If
                    'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 2 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 2)
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.PreferredWidth = 100
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Từ ngày tháng năm -" + vbCrLf + " Đến ngày tháng năm")
                    objTable.Cell(1, 2).Range.InsertAfter("Chức danh, chức vụ, đơn vị công tác (Đảng, Chính quyền, Đoàn thể)")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    If Not (db_ct Is Nothing) Then
                        If (db_ct.Rows.Count > 0) Then
                            Dim row As Integer = 0
                            For i As Integer = 0 To db_ct.Rows.Count - 1
                                Dim ThoiGian As String = ""
                                If CInt(db_ct.Rows(i)("IsQD_NHCS")) = 0 Then
                                    If i = 0 Then
                                        row = i
                                        ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        Dim j As Integer = i + 1
                                        While j < db_ct.Rows.Count
                                            Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    Exit While
                                                Case Else
                                                    j = j + 1
                                            End Select
                                        End While
                                        objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                        objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                        objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                        objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                    Else
                                        If Not (db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1504" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1505" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1506" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1511" _
                                                Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1512" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1519" _
                                                Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1520" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1521" Or db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1522") Then
                                            ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                            If Not (db_ct.Rows(i)("DenNgay") Is DBNull.Value) Then
                                                If Not (CType(db_ct.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(db_ct.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy")
                                                    row = row + 1
                                                    objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                                    objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                                    objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                                    objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                                Else

                                                    Dim j As Integer = i + 1
                                                    While j < db_ct.Rows.Count
                                                        Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                            Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                                ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                                Exit While
                                                            Case Else
                                                                j = j + 1
                                                        End Select
                                                    End While
                                                    row = row + 1
                                                    objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                                    objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                                    objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                                    objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F

                                                End If
                                            Else

                                                Dim j As Integer = i + 1
                                                While j < db_ct.Rows.Count
                                                    Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                        Case "1501", "1502", "1503", "1507", "1508", "1509", "1510", "1513", "1514", "1516", "1532", "1534", "1535", "1536", "3701", "3702", "3703", "3704", "3705"
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            Exit While
                                                        Case Else
                                                            j = j + 1
                                                    End Select
                                                End While
                                                row = row + 1
                                                objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                                objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                                objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("Ban Tổng giám đốc", ""), ""))
                                                objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                            End If
                                        End If
                                    End If
                                Else
                                    'ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                    'Dim j As Integer = i + 1
                                    'While j < db_ct.Rows.Count
                                    '    If (CInt(db_ct.Rows(i)("IDDonVi_Moi")) = CInt(db_ct.Rows(j)("IDDonVi_Moi")) And CInt(db_ct.Rows(i)("IdPhong_Moi")) = CInt(db_ct.Rows(j)("IdPhong_Moi")) And CInt(db_ct.Rows(i)("IDChucVu_Moi")) = CInt(db_ct.Rows(j)("IDChucVu_Moi"))) Then
                                    '        j = j + 1
                                    '    Else
                                    '        ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                    '        i = j - 1
                                    '        Exit While
                                    '    End If
                                    'End While
                                    'If i <> 0 Then row = row + 1
                                    'objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                    'objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                    'objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                    'objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString(), ""))
                                    'objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                    'If j = db_ct.Rows.Count Then Exit For
                                    'Đọc các QĐ khác QĐ thôi kiêm nhiệm
                                    If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() <> "1519" Then
                                        ThoiGian = "Từ " & IIf(db_ct.Rows(i)("NgayHL").ToString().Trim() <> "", CType(db_ct.Rows(i)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                        Dim j As Integer = i + 1
                                        While j < db_ct.Rows.Count
                                            Select Case db_ct.Rows(j)("MaLoaiQD").ToString().Trim()
                                                Case "1534", "1535"
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    Exit While
                                                Case "1518"
                                                    'Bỏ qua nếu QĐ tiếp theo là QĐ kiêm nhiệm
                                                    Exit While
                                                Case "1519"
                                                    'Lấy mốc hết thời gian kiêm nhiệm
                                                    ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                    j = j + 1
                                                    Exit While
                                                Case Else
                                                    'TH không đọc tiếp nếu: Nếu đơn vị, phòng ban, chức vụ cũ; hoặc QĐ tiếp theo là QĐ kiêm nhiệm
                                                    If (CInt(db_ct.Rows(i)("IDDonVi_Moi")) = CInt(db_ct.Rows(j)("IDDonVi_Moi")) And CInt(db_ct.Rows(i)("IdPhong_Moi")) = CInt(db_ct.Rows(j)("IdPhong_Moi")) And CInt(db_ct.Rows(i)("IDChucVu_Moi")) = CInt(db_ct.Rows(j)("IDChucVu_Moi"))) Then
                                                        j = j + 1
                                                    Else
                                                        'Nếu QĐ trước đó la QD Kiêm nhiệm mà QĐ tiếp theo ko phải là QĐ Thôi kiêm nhiệm thì bỏ qua
                                                        If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                            Exit While
                                                        Else
                                                            ThoiGian = ThoiGian.Substring(3) & " - " & CType(CType(db_ct.Rows(j)("NgayHL"), Date).AddDays(-1), DateTime).ToString("dd-MM-yyyy")
                                                            i = j - 1
                                                            Exit While
                                                        End If
                                                    End If
                                            End Select
                                        End While
                                        If i <> 0 Then row = row + 1
                                        objTable.Rows(row + 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                        objTable.Cell(row + 2, 1).Range.InsertAfter(ThoiGian)
                                        objTable.Cell(row + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1534" Then
                                            objTable.Cell(row + 2, 2).Range.InsertAfter(db_ct.Rows(i)("LoaiQD").ToString())
                                        Else
                                            If db_ct.Rows(i)("MaLoaiQD").ToString().Trim() = "1518" Then
                                                objTable.Cell(row + 2, 2).Range.InsertAfter(db_ct.Rows(i)("LoaiQD").ToString() & " " & IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace("(CN cấp I & tương đương)", ""), ""))
                                            Else
                                                objTable.Cell(row + 2, 2).Range.InsertAfter(IIf(db_ct.Rows(i)("NoiDung").ToString() <> "", db_ct.Rows(i)("NoiDung").ToString().Replace(", Ban Tổng giám đốc, Hội sở chính", " Ngân hàng Chính sách xã hội Việt Nam").Replace(", Ban Giám đốc (CN cấp I & tương đương)", ""), ""))
                                            End If

                                        End If
                                        objTable.Cell(row + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                        If j = db_ct.Rows.Count Then Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 24
                    objTable.Columns(2).PreferredWidth = 76
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Font.Name = "Times New Roman"
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.Paragraphs.SpaceBefore = 0
                objselection.Paragraphs.SpaceAfter = 4
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.TypeText("42. Đặc điểm, lý lịch bản thân: " + vbCrLf)
                objselection.Font.Size = 11
                objselection.Font.Italic = 1
                If (dr("DacDiem_BT").ToString().Trim() <> "") Then
                    objselection.TypeText(vbTab + dr("DacDiem_BT").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
                Else
                    '_sVal = "      a - Khai rõ: bị bắt, bị tù (từ ngày tháng năm nào đến ngày tháng năm nào, ở đâu), đã khai báo cho ai, những vấn đề gì:" + vbCrLf
                    _sVal = "     .........................................................................................................................................................................." + vbCrLf
                    _sVal += "     .........................................................................................................................................................................." + vbCrLf
                    objselection.TypeText(_sVal)
                End If
                'If (dr("GhiChu").ToString().Trim() <> "") Then
                '    objselection.TypeText(vbTab + "- " + dr("GhiChu").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
                'Else
                '    _sVal = "      b - Bản thân có làm việc trong chế độ cũ (Cơ quan, đơn vị nào, địa điểm, chức danh, chức vụ, thời gian làm việc):" + vbCrLf
                '    _sVal += "     ..........................................................................................................................................................................." + vbCrLf
                '    _sVal += "     ..........................................................................................................................................................................." + vbCrLf
                '    objselection.TypeText(_sVal)
                'End If
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.TypeText("43. Quan hệ với nước ngoài: " + vbCrLf)
                objselection.Font.Size = 11
                objselection.Font.Italic = 1
                If (dr("QuanHe_Nguoi_NN").ToString().Trim() <> "") Then
                    objselection.TypeText(vbTab + "- " + dr("QuanHe_Nguoi_NN").ToString().Trim().Replace(vbCrLf, vbCrLf + vbTab + "- ") + vbCrLf)
                Else
                    _sVal = "      Tham gia hoặc có quan hệ với các tổ chức chính trị, kinh tế, xã hội nào ở nước ngoài (làm gì, tổ chức nào, đặt trụ sở ở đâu?): .................................................................................................................................." + vbCrLf
                    _sVal += "     ............................................................................................................................................................................" + vbCrLf
                    _sVal += "     ............................................................................................................................................................................" + vbCrLf
                    _sVal += "     ............................................................................................................................................................................" + vbCrLf
                    _sVal += "     Có thân nhân (Bố mẹ, vợ, chồng, con, anh chị em ruột) ở nước ngoài (làm gì, địa chỉ?): ....................................................." + vbCrLf
                    _sVal += "     ............................................................................................................................................................................" + vbCrLf
                    _sVal += "     ............................................................................................................................................................................" + vbCrLf
                    _sVal += "     ............................................................................................................................................................................" + vbCrLf
                    objselection.TypeText(_sVal)
                End If
                '44. Thông tin tài khoản ngân hàng
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.TypeText("44. Thông tin tài khoản ngân hàng: " + vbTab)
                objselection.Font.Italic = 1
                objselection.TypeText("- Mã khách hàng: " + dr("NH_MaKH").ToString().Trim() + vbTab + " - Số tài khoản: " + dr("NH_SoTK").ToString().Trim() + vbCrLf)
                objselection.Font.Size = 10
                objselection.TypeText(vbTab + "(Tài khoản trả lương)" + vbTab)
                objselection.Font.Size = 12
                objselection.TypeText(vbTab + "- Tên ngân hàng mở tài khoản: " + dr("NH_TenNH").ToString().Trim() + vbCrLf)
                objselection.Font.Italic = 0
                objselection.TypeText("45. Mã số thế cá nhân: " + dr("MaSoThue").ToString().Trim() + vbCrLf)
                objselection.TypeText("46. Bảo hiểm xã hội: " + vbTab)
                objselection.Font.Italic = 1

                objselection.TypeText("- Số sổ: " + dr("BHXH_SoSo").ToString().Trim() + vbTab)
                _sVal = ""
                If dr("BHXH_NgaySo").ToString() <> "" Then _sVal = IIf(CType(dr("BHXH_NgaySo").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("BHXH_NgaySo").ToString(), DateTime).ToString("dd/MM/yyyy"), "")
                objselection.TypeText("- Ngày cấp sổ: " + _sVal + vbTab)
                _sVal = ""
                If dr("BHXH_NgayBatDau").ToString() <> "" Then _sVal = IIf(CType(dr("BHXH_NgayBatDau").ToString(), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900", CType(dr("BHXH_NgayBatDau").ToString(), DateTime).ToString("dd/MM/yyyy"), "")
                objselection.TypeText("- Ngày bắt đầu tham gia BHXH: " + _sVal + vbCrLf)
                objselection.TypeText(vbTab + vbTab + vbTab + "- Nơi cấp sổ: " + dr("BHXH_NoiLam").ToString().Trim() + vbCrLf)
                objselection.Font.Italic = 0
                '47. Tham gia Đoàn công tác nước ngoài (do NHCSXH cử)
                objselection.TypeText("47. Tham gia Đoàn công tác nước ngoài (do NHCSXH cử)" + vbCrLf)
                strSQL = "SELECT a.SoQD,Convert(Varchar,a.NgayKy_QD,103) As NgayKy_QD,b.Ten_Goi as NuocDen,a.MucDich,Convert(Varchar,a.TuNgay,103) As TG_TuNgay,"
                strSQL += "Convert(Varchar,a.DenNgay,103) As TG_DenNgay,a.NguoiKy_QD FROM HS_XuatNgoai a, QuocGia b "
                strSQL += " WHERE a.IdCanBo = '" + _RowId + "' And (a.IdNuocDen = b.id and b.Status = 1) Order by a.TuNgay Asc"
                objselection.Font.Bold = 0
                objselection.Font.Size = 11
                Using db_xn As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_xn Is Nothing Or db_xn.Rows.Count = 0) Then
                        _rowsTable = 3
                    Else
                        _rowsTable = db_xn.Rows.Count + 2
                    End If
                    'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 4 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 7)
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Columns(1).PreferredWidth = 7
                    objTable.Columns(2).PreferredWidth = 7
                    objTable.Columns(3).PreferredWidth = 14
                    objTable.Columns(4).PreferredWidth = 32
                    objTable.Columns(5).PreferredWidth = 7
                    objTable.Columns(6).PreferredWidth = 7
                    objTable.Columns(7).PreferredWidth = 12
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Quyết định số")
                    objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
                    objTable.Cell(1, 3).Range.InsertAfter("Nước đến")
                    objTable.Cell(1, 4).Range.InsertAfter("Mục đích")
                    objTable.Cell(1, 5).Range.InsertAfter("Thời gian công tác")
                    objTable.Cell(1, 7).Range.InsertAfter("Người ký quyết định")
                    objTable.Cell(2, 5).Range.InsertAfter("Từ ngày, tháng")
                    objTable.Cell(2, 6).Range.InsertAfter("Đến ngày, tháng")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(2).Range.Font.Bold = 1
                    If Not (db_xn Is Nothing) Then
                        If (db_xn.Rows.Count > 0) Then
                            For i As Integer = 0 To db_xn.Rows.Count - 1
                                objTable.Cell(i + 3, 1).Range.InsertAfter(db_xn.Rows(i)("SoQD").ToString().Trim())
                                objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 3, 2).Range.InsertAfter(db_xn.Rows(i)("NgayKy_QD").ToString().Trim())
                                objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 3).Range.InsertAfter(db_xn.Rows(i)("NuocDen").ToString().Trim())
                                objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 4).Range.InsertAfter(db_xn.Rows(i)("MucDich").ToString().Trim())
                                objTable.Cell(i + 3, 4).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 3, 5).Range.InsertAfter(db_xn.Rows(i)("TG_TuNgay").ToString().Trim())
                                objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 6).Range.InsertAfter(db_xn.Rows(i)("TG_DenNgay").ToString().Trim())
                                objTable.Cell(i + 3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 7).Range.InsertAfter(db_xn.Rows(i)("NguoiKy_QD").ToString().Trim())
                                objTable.Cell(i + 3, 7).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                    objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
                    objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
                    objTable.Cell(1, 3).Merge(objTable.Cell(2, 3))
                    objTable.Cell(1, 4).Merge(objTable.Cell(2, 4))
                    objTable.Cell(1, 7).Merge(objTable.Cell(2, 7))
                    objTable.Cell(1, 5).Merge(objTable.Cell(1, 6))
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                '48. Hồ sơ Hộ chiếu của cán bộ
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Name = "Times New Roman"
                objselection.Font.Size = 12
                objselection.Font.Italic = 0
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.TypeText("48. Hộ chiếu" + vbCrLf)
                strSQL = "SELECT So_HoChieu, (Case Loai_HC When 1 Then N'Phổ thông' Else ((Case Loai_HC When 2 Then N'Công vụ' Else N'Ngoại giao' End)) End) As LoaiHC,Convert(Varchar,NgayCap,103) As NgCapHC, NoiCap, Convert(Varchar,NgayHH,103) As NgHetHanHC,"
                strSQL += "(Case TinhTrang When 1 Then N'Còn hiệu lực' Else ((Case TinhTrang When 2 Then N'Hết hạn' Else N'Mất' End)) End) As TinhTrang, GhiChu FROM HS_HoChieu Where IdCanBo = '" + _RowId + "' Order by NgayCap Asc"
                objselection.Font.Bold = 0
                objselection.Font.Size = 11
                Using db_hc As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_hc Is Nothing Or db_hc.Rows.Count = 0) Then
                        _rowsTable = 2
                    Else
                        _rowsTable = db_hc.Rows.Count + 1
                    End If
                    'Tạo một bảng có db_vb.Rows.Count + 1 dòng và 7 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 7)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                    objTable.Columns(1).PreferredWidth = 10
                    objTable.Columns(2).PreferredWidth = 13
                    objTable.Columns(3).PreferredWidth = 12
                    objTable.Columns(4).PreferredWidth = 26
                    objTable.Columns(5).PreferredWidth = 12
                    objTable.Columns(6).PreferredWidth = 13
                    objTable.Columns(7).PreferredWidth = 20
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Số hộ chiếu")
                    objTable.Cell(1, 2).Range.InsertAfter("Loại hộ chiếu")
                    objTable.Cell(1, 3).Range.InsertAfter("Ngày cấp")
                    objTable.Cell(1, 4).Range.InsertAfter("Nơi cấp")
                    objTable.Cell(1, 5).Range.InsertAfter("Ngày hết hạn")
                    objTable.Cell(1, 6).Range.InsertAfter("Tình trạng")
                    objTable.Cell(1, 7).Range.InsertAfter("Ghi chú thêm")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    If Not (db_hc Is Nothing) Then
                        If (db_hc.Rows.Count > 0) Then
                            For i As Integer = 0 To db_hc.Rows.Count - 1
                                objTable.Cell(i + 2, 1).Range.InsertAfter(db_hc.Rows(i)("So_HoChieu").ToString().Trim())
                                objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 2).Range.InsertAfter(db_hc.Rows(i)("LoaiHC").ToString().Trim())
                                objTable.Cell(i + 2, 2).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 3).Range.InsertAfter(db_hc.Rows(i)("NgCapHC").ToString().Trim())
                                objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 4).Range.InsertAfter(db_hc.Rows(i)("NoiCap").ToString().Trim())
                                objTable.Cell(i + 2, 4).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 5).Range.InsertAfter(db_hc.Rows(i)("NgHetHanHC").ToString().Trim())
                                objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 6).Range.InsertAfter(db_hc.Rows(i)("TinhTrang").ToString().Trim())
                                objTable.Cell(i + 2, 6).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 7).Range.InsertAfter(db_hc.Rows(i)("GhiChu").ToString().Trim())
                                objTable.Cell(i + 2, 7).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                End Using

                '49. Quan hệ gia đình
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Color = Word.WdColor.wdColorAutomatic
                objselection.Font.Name = "Times New Roman"
                objselection.Font.Size = 12
                objselection.Font.Bold = 1
                objselection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                objselection.TypeText("49. Quan hệ gia đình:" + vbCrLf)
                objselection.Font.Bold = 0
                objselection.Font.Italic = 1
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText("a. Bố, Mẹ, Vợ (Chồng), các con, anh chị em ruột" + vbCrLf)
                objselection.Font.Size = 11
                objselection.Font.Italic = 0
                strSQL = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2303','2304','2314','2313','2312','2317','2305','2320','2315','2316'))"
                Using db_rows As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_rows Is Nothing Or CType(db_rows.Rows(0)(0), Byte) = 0) Then
                        _rowsTable = 6
                    Else
                        Dim db As DBAccess = New DBAccess
                        Dim ssql As String = ""
                        'Dem So Bo khác bố đẻ
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so in ('2318', '2332'))"
                        If db.getNumber(ssql) = 0 Then
                            _rowsTable = 4
                        Else
                            _rowsTable = 3 + db.getNumber(ssql)
                        End If
                        'Dem So Me khác mẹ đẻ
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so in ('2319', '2333'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        'Dem So con
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2312','2317'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        'If db.getNumber(ssql) = 0 Then
                        '    _rowsTable = 5
                        'Else
                        '    _rowsTable = 4 + db.getNumber(ssql)
                        'End If
                        'Dem So anh chi em
                        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316'))"
                        _rowsTable = _rowsTable + db.getNumber(ssql)
                        '_rowsTable = CType(db_rows.Rows(0)(0), Byte) + 1
                        If _rowsTable <= 6 Then _rowsTable = 6
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
                    objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
                    objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
                    objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
                    objTable.Cell(1, 4).Range.Font.Bold = 0
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(2, 1).Range.InsertAfter("Bố")
                    objTable.Cell(2, 1).Range.Font.Bold = 1
                    objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
                    'objTable.Cell(3, 1).Range.Font.Bold = 1
                    'objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'If (dr("GioiTinh").ToString() = False) Then
                    '    objTable.Cell(4, 1).Range.InsertAfter("Vợ" + vbCrLf)
                    'Else
                    '    objTable.Cell(4, 1).Range.InsertAfter("Chồng" + vbCrLf)
                    'End If
                    'objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'objTable.Cell(4, 1).Range.Font.Bold = 1

                    For i As Byte = 2 To _rowsTable
                        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                    Next

                    If Not (db_rows Is Nothing) Then
                        If (db_rows.Rows.Count > 0) Then
                            'Lấy thông tin Bố đẻ
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2303')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(2, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            Dim _rowGD As Integer = 2

                            'Lấy thông tin Dượng
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2318')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Bố dượng")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Bố nuôi
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2332')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Bố nuôi")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Mẹ đẻ
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2304')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Mẹ")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Mẹ kế
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2319')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Mẹ kế")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Mẹ nuôi
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2333')"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(_rowGD, 1).Range.InsertAfter("Mẹ kế")
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Vợ hoặc chồng
                            Dim rowCVC As Integer = 4  'Danh dau dong Chong/Vo/Con
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2314','2313'))"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowGD = _rowGD + 1
                                        If (dr("GioiTinh").ToString() = False) Then
                                            objTable.Cell(_rowGD, 1).Range.InsertAfter("Vợ" + vbCrLf)
                                        Else
                                            objTable.Cell(_rowGD, 1).Range.InsertAfter("Chồng" + vbCrLf)
                                        End If
                                        objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(0)("ConMat") = False) Then
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(_rowGD, 4).Range.InsertAfter("Đã mất")
                                        End If
                                    Else
                                        objTable.Cell(_rowGD, 2).Range.InsertAfter(vbCrLf)
                                        objTable.Cell(_rowGD, 3).Range.InsertAfter(vbCrLf)
                                        objTable.Cell(_rowGD, 4).Range.InsertAfter(vbCrLf)
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Các con đẻ
                            Dim _RowCon As Byte = 0
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2312','2317')) order by NamSinh"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If db_gd.Rows.Count = 0 Then
                                    _rowGD = _rowGD + 1
                                    objTable.Cell(_rowGD, 1).Range.InsertAfter("Con")
                                    objTable.Cell(_rowGD, 1).Range.Font.Bold = 1
                                    objTable.Cell(_rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                End If
                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _RowCon = db_gd.Rows.Count
                                        _rowGD = _rowGD + 1
                                        For i As Integer = 0 To db_gd.Rows.Count - 1
                                            If (i = 0) Then
                                                objTable.Cell(i + _rowGD, 1).Range.InsertAfter("Các" + vbCrLf + "con")
                                            End If
                                            objTable.Cell(i + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            objTable.Cell(i + _rowGD, 1).Range.Font.Bold = 1
                                            objTable.Cell(i + _rowGD, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _rowGD, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                                            objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                            If (db_gd.Rows(i)("ConMat") = False) Then
                                                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                                                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                                            Else
                                                objTable.Cell(i + _rowGD, 4).Range.InsertAfter("Đã mất")
                                            End If

                                            If (i > 0) Then objTable.Rows(i + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                        Next
                                    End If
                                End If
                            End Using

                            'Lấy thông tin Anh chị em ruột thịt
                            Dim _rowAnhChi As Integer = 0
                            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316')) order by NamSinh"
                            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If db_gd.Rows.Count = 0 Then
                                    If _RowCon = 0 Then
                                        _rowGD = _rowGD + 1
                                        objTable.Cell(6, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                                        objTable.Cell(6, 1).Range.Font.Bold = 1
                                        objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                    Else
                                        objTable.Cell(_RowCon + _rowGD, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                                        objTable.Cell(_RowCon + _rowGD, 1).Range.Font.Bold = 1
                                        objTable.Cell(_RowCon + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                    End If
                                End If

                                If Not (db_gd Is Nothing) Then
                                    If (db_gd.Rows.Count > 0) Then
                                        _rowAnhChi = db_gd.Rows.Count
                                        If _RowCon = 0 Then _RowCon = 1
                                        For i As Integer = 0 To db_gd.Rows.Count - 1
                                            If (i = 0) Then
                                                objTable.Cell(i + _RowCon + _rowGD, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                                                objTable.Rows(i + _RowCon + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                            End If
                                            objTable.Cell(i + _RowCon + _rowGD, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            objTable.Cell(i + _RowCon + _rowGD, 1).Range.Font.Bold = 1
                                            objTable.Cell(i + _RowCon + _rowGD, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _RowCon + _rowGD, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                            objTable.Cell(i + _RowCon + _rowGD, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                                            objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                            If (db_gd.Rows(i)("ConMat") = False) Then
                                                objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                                                objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                                            Else
                                                objTable.Cell(i + _RowCon + _rowGD, 4).Range.InsertAfter("Đã mất")
                                            End If
                                            If (i > 0) Then
                                                Do Until objTable.Rows.Count >= i + _RowCon + _rowGD
                                                    objTable.Rows.Add()
                                                Loop
                                                objTable.Rows(i + _RowCon + _rowGD).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                            End If
                                        Next
                                        objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                        objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                        objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                        objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    End If
                                End If
                            End Using
                            'Merge ô cuối trước
                            'If (_rowAnhChi > 1) Then
                            '    objTable.Cell(_RowCon + _rowGD, 1).Merge(objTable.Cell(_rowsTable, 1))
                            'End If

                            'If (_rowAnhChi > 1) Then
                            '    objTable.Cell(_RowCon + _rowGD + 1, 1).Merge(objTable.Cell(_rowsTable, 1))
                            'End If

                            'If (_rowsTable > 6) Then objTable.Cell(6, 1).Merge(objTable.Cell(_rowsTable, 1))

                            If _RowCon > 1 Then objTable.Cell(_rowGD, 1).Merge(objTable.Cell(_rowGD + _RowCon - 1, 1))


                            'objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                            'objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                            'objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                            'objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                            'Thực hiện Merge Cell Anh chị em ruột



                        End If
                    End If
                    'Thực hiện Merge Cell Anh chị em ruột
                    objTable.Columns(1).PreferredWidth = 9
                    objTable.Columns(2).PreferredWidth = 18
                    objTable.Columns(3).PreferredWidth = 9
                    objTable.Columns(4).PreferredWidth = 59
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                'Tạo một bảng có _rowsTable dòng và 4 cột
                'Tạo một bảng có _rowsTable dòng và 4 cột
                'strSQL = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2303','2304','2314','2313','2312','2317','2305','2320','2315','2316'))"
                'Using db_rows As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '    If (db_rows Is Nothing Or CType(db_rows.Rows(0)(0), Byte) = 0) Then
                '        _rowsTable = 6
                '    Else
                '        Dim db As DBAccess = New DBAccess
                '        Dim ssql As String = ""
                '        'Dem So con
                '        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2312','2317'))"
                '        If db.getNumber(ssql) = 0 Then
                '            _rowsTable = 5
                '        Else
                '            _rowsTable = 4 + db.getNumber(ssql)
                '        End If
                '        'Dem So anh chi em
                '        ssql = "SELECT Count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316'))"
                '        _rowsTable = _rowsTable + db.getNumber(ssql)
                '        '_rowsTable = CType(db_rows.Rows(0)(0), Byte) + 1
                '        If _rowsTable <= 6 Then _rowsTable = 6
                '    End If
                '    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
                '    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                '    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                '    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                '    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                '    objTable.Range.Paragraphs.SpaceBefore = 2
                '    objTable.Range.Paragraphs.SpaceAfter = 2
                '    'Add tiêu đề của các cột trong Bảng
                '    objTable.Rows(1).Range.Font.Bold = 1
                '    objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
                '    objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
                '    objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
                '    objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
                '    objTable.Cell(1, 4).Range.Font.Bold = 0
                '    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    objTable.Cell(2, 1).Range.InsertAfter("Bố")
                '    objTable.Cell(2, 1).Range.Font.Bold = 1
                '    objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
                '    objTable.Cell(3, 1).Range.Font.Bold = 1
                '    objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    If (dr("GioiTinh").ToString() = False) Then
                '        objTable.Cell(4, 1).Range.InsertAfter("Vợ" + vbCrLf)
                '    Else
                '        objTable.Cell(4, 1).Range.InsertAfter("Chồng" + vbCrLf)
                '    End If
                '    objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    objTable.Cell(4, 1).Range.Font.Bold = 1
                '    'If _rowsTable = 6 And CType(db_rows.Rows(0)(0), Byte) = 0 Then
                '    '    objTable.Cell(5, 1).Range.InsertAfter("Các" + vbCrLf + "con")
                '    '    objTable.Cell(5, 1).Range.Font.Bold = 1
                '    '    objTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    '    objTable.Cell(6, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                '    '    objTable.Cell(6, 1).Range.Font.Bold = 1
                '    '    objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    'End If
                '    For i As Byte = 2 To _rowsTable
                '        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                '    Next
                '    If Not (db_rows Is Nothing) Then
                '        If (db_rows.Rows.Count > 0) Then
                '            'Lấy thông tin Bố đẻ
                '            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2303')"
                '            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '                If Not (db_gd Is Nothing) Then
                '                    If (db_gd.Rows.Count > 0) Then
                '                        objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                '                        objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                '                        objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                '                        objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                '                        objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                '                        objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                '                    End If
                '                End If
                '            End Using

                '            'Lấy thông tin Mẹ đẻ
                '            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so = '2304')"
                '            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '                If Not (db_gd Is Nothing) Then
                '                    If (db_gd.Rows.Count > 0) Then
                '                        objTable.Cell(3, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                '                        objTable.Cell(3, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                '                        objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                '                        objTable.Cell(3, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                '                        objTable.Cell(3, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                '                        objTable.Cell(3, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                '                    End If
                '                End If
                '            End Using

                '            'Lấy thông tin Vợ hoặc chồng
                '            Dim rowCVC As Integer = 4  'Danh dau dong Chong/Vo/Con
                '            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2314','2313'))"
                '            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '                If Not (db_gd Is Nothing) Then
                '                    If (db_gd.Rows.Count > 0) Then
                '                        objTable.Cell(4, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                '                        objTable.Cell(4, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                '                        objTable.Cell(4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                        _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                '                        objTable.Cell(4, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                '                        objTable.Cell(4, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                '                        objTable.Cell(4, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                '                    Else
                '                        objTable.Cell(4, 2).Range.InsertAfter(vbCrLf)
                '                        objTable.Cell(4, 3).Range.InsertAfter(vbCrLf)
                '                        objTable.Cell(4, 4).Range.InsertAfter(vbCrLf)
                '                    End If
                '                End If
                '            End Using

                '            'Lấy thông tin Các con đẻ
                '            Dim _RowCon As Byte = 0
                '            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2312','2317')) order by NamSinh"
                '            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '                If db_gd.Rows.Count = 0 Then
                '                    objTable.Cell(5, 1).Range.InsertAfter("Các" + vbCrLf + "con")
                '                    objTable.Cell(5, 1).Range.Font.Bold = 1
                '                    objTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                End If
                '                If Not (db_gd Is Nothing) Then
                '                    If (db_gd.Rows.Count > 0) Then
                '                        _RowCon = db_gd.Rows.Count
                '                        For i As Integer = 0 To db_gd.Rows.Count - 1
                '                            If (i = 0) Then
                '                                objTable.Cell(i + 5, 1).Range.InsertAfter("Các" + vbCrLf + "con")
                '                            End If
                '                            objTable.Cell(i + 5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                            objTable.Cell(i + 5, 1).Range.Font.Bold = 1
                '                            objTable.Cell(i + 5, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                '                            objTable.Cell(i + 5, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                '                            objTable.Cell(i + 5, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                '                            objTable.Cell(i + 5, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                '                            objTable.Cell(i + 5, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                '                            objTable.Cell(i + 5, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                '                            If (i > 0) Then objTable.Rows(i + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                '                        Next
                '                    End If
                '                End If
                '                '                                If _RowCon = db_gd.Rows.Count Then objTable.Rows(_RowCon - 1 + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '            End Using

                '            'Lấy thông tin Anh chị em ruột thịt
                '            Dim _rowAnhChi As Integer = 0
                '            strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2305','2320','2315','2316')) order by NamSinh"
                '            Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '                If db_gd.Rows.Count = 0 Then
                '                    If _RowCon = 0 Then
                '                        objTable.Cell(6, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                '                        objTable.Cell(6, 1).Range.Font.Bold = 1
                '                        objTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                    Else
                '                        objTable.Cell(_RowCon + 5, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                '                        objTable.Cell(_RowCon + 5, 1).Range.Font.Bold = 1
                '                        objTable.Cell(_RowCon + 5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                    End If
                '                End If

                '                If Not (db_gd Is Nothing) Then
                '                    If (db_gd.Rows.Count > 0) Then
                '                        _rowAnhChi = db_gd.Rows.Count
                '                        If _RowCon = 0 Then _RowCon = 1
                '                        For i As Integer = 0 To db_gd.Rows.Count - 1
                '                            If (i = 0) Then
                '                                objTable.Cell(i + _RowCon + 5, 1).Range.InsertAfter("Anh" + vbCrLf + "chị, em" + vbCrLf + "ruột")
                '                                objTable.Rows(i + _RowCon + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '                            End If
                '                            objTable.Cell(i + _RowCon + 5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                            objTable.Cell(i + _RowCon + 5, 1).Range.Font.Bold = 1
                '                            objTable.Cell(i + _RowCon + 5, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                '                            objTable.Cell(i + _RowCon + 5, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                '                            objTable.Cell(i + _RowCon + 5, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                            _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                '                            objTable.Cell(i + _RowCon + 5, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                '                            objTable.Cell(i + _RowCon + 5, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                '                            objTable.Cell(i + _RowCon + 5, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                '                            'If (i > 0) Then objTable.Rows(i + _RowCon + 3).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                '                            If (i > 0) Then objTable.Rows(i + _RowCon + 5).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                '                        Next
                '                        objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '                        objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '                        objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '                        objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '                    End If
                '                End If
                '            End Using
                '            'Merge ô cuối trước
                '            If (_rowAnhChi > 1) Then
                '                'objTable.Cell(_RowCon + 4, 1).Merge(objTable.Cell(_rowsTable, 1))
                '                objTable.Cell(_RowCon + 5, 1).Merge(objTable.Cell(_rowsTable, 1))
                '            End If
                '            'If _RowCon > 1 Then objTable.Cell(4, 1).Merge(objTable.Cell(3 + _RowCon, 1))
                '            If _RowCon > 1 Then objTable.Cell(5, 1).Merge(objTable.Cell(5 + _RowCon - 1, 1))

                '        End If
                '    End If
                '    'Thực hiện Merge Cell Anh chị em ruột
                '    objTable.Columns(1).PreferredWidth = 9
                '    objTable.Columns(2).PreferredWidth = 18
                '    objTable.Columns(3).PreferredWidth = 9
                '    objTable.Columns(4).PreferredWidth = 59
                '    objTable.LeftPadding = 0' -3
                '    objTable.RightPadding = 0' -4
                'End Using

                'Mục 49 b (Bố mẹ, anh chị em ruột bên vợ hoặc chồng)
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.Font.Italic = 1
                objselection.TypeText("b. Bố, mẹ, anh chị em ruột bên vợ hoặc chồng: " + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 11
                objselection.Font.Italic = 0
                strSQL = "SELECT count(IdTVien) FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe IN (Select id From DanhMuc Where ma_so IN ('2326','2327','2328','2329','2330','2331'))"
                Using db_gdvochong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_gdvochong Is Nothing OrElse CType(db_gdvochong.Rows(0)(0), Byte) = 0) Then
                        _rowsTable = 4
                    Else
                        _rowsTable = CType(db_gdvochong.Rows(0)(0), Byte) + 3
                        If _rowsTable <= 4 Then _rowsTable = 4
                    End If
                    'Tạo một bảng có 4 dòng và 4 cột
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 4)
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
                    objTable.Cell(1, 2).Range.InsertAfter("Họ và tên")
                    objTable.Cell(1, 3).Range.InsertAfter("Năm sinh")
                    objTable.Cell(1, 4).Range.InsertAfter("Quê quán, nghề nghiệp, chức danh, chức vụ, đơn vị, công tác, học tập, nơi ở (trong, ngoài nước); thành viên các tổ chức chính trị - xã hội ...")
                    objTable.Cell(1, 4).Range.Font.Bold = 0
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(2, 1).Range.InsertAfter("Bố")
                    objTable.Cell(3, 1).Range.InsertAfter("Mẹ")
                    objTable.Cell(4, 1).Range.InsertAfter("Anh" + vbCrLf + "chị em" + vbCrLf + "ruột")
                    objTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Cell(2, 1).Range.Font.Bold = 1
                    objTable.Cell(3, 1).Range.Font.Bold = 1
                    objTable.Cell(4, 1).Range.Font.Bold = 1
                    For i As Integer = 2 To _rowsTable
                        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                    Next
                    'Lấy thông tin Bố vợ (Bố chồng)
                    strSQL = "SELECT Top 1 * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2322','2324'))"
                    Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db_gd Is Nothing) Then
                            If (db_gd.Rows.Count > 0) Then
                                objTable.Cell(2, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(2, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                objTable.Cell(2, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                If (db_gd.Rows(0)("ConMat") = False) Then
                                    objTable.Cell(2, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                    objTable.Cell(2, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                Else
                                    objTable.Cell(2, 4).Range.InsertAfter("Đã mất")
                                End If
                            End If
                        End If
                    End Using
                    'Lấy thông tin Mẹ vợ (Mẹ chồng)
                    strSQL = "SELECT Top 1 * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2323','2325'))"
                    Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db_gd Is Nothing) Then
                            If (db_gd.Rows.Count > 0) Then
                                objTable.Cell(3, 2).Range.InsertAfter(db_gd.Rows(0)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(3, 3).Range.InsertAfter(db_gd.Rows(0)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                objTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                _sVal = IIf(db_gd.Rows(0)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(0)("IdQueQuan").ToString().Trim(), Integer)), "")
                                objTable.Cell(3, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                If (db_gd.Rows(0)("ConMat") = False) Then
                                    objTable.Cell(3, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(0)("DiaChi").ToString().Trim() + vbCrLf)
                                    objTable.Cell(3, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(0)("NgheNghiep").ToString().Trim())
                                Else
                                    objTable.Cell(3, 4).Range.InsertAfter("Đã mất")
                                End If
                            End If
                        End If
                    End Using
                    If (CType(db_gdvochong.Rows(0)(0), Byte) > 0) Then
                        'Lấy thông tin Anh chị em ruột bên vợ (hoặc bên chồng)
                        strSQL = "SELECT * FROM HS_GDCB WHERE IdCanBo = '" + _RowId + "' And IdQuanHe In (Select id From DanhMuc Where ma_so IN ('2326','2327','2328','2329','2330','2331')) order by NamSinh"
                        Using db_gd As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db_gd Is Nothing) Then
                                If (db_gd.Rows.Count > 0) Then
                                    For i As Integer = 0 To db_gd.Rows.Count - 1
                                        objTable.Cell(i + 4, 2).Range.InsertAfter(db_gd.Rows(i)("HoTen").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(i + 4, 3).Range.InsertAfter(db_gd.Rows(i)("NamSinh").ToString().Trim() + vbCrLf + vbCrLf)
                                        objTable.Cell(i + 4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        _sVal = IIf(db_gd.Rows(i)("IdQueQuan").ToString().Trim() <> "", GetAddress(CType(db_gd.Rows(i)("IdQueQuan").ToString().Trim(), Integer)), "")
                                        objTable.Cell(i + 4, 4).Range.InsertAfter("Quê quán: " + _sVal + vbCrLf)
                                        If (db_gd.Rows(i)("ConMat") = False) Then
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Nơi ở: " + db_gd.Rows(i)("DiaChi").ToString().Trim() + vbCrLf)
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Nghề nghiệp, chức vụ, đơn vị: " + db_gd.Rows(i)("NgheNghiep").ToString().Trim())
                                        Else
                                            objTable.Cell(i + 4, 4).Range.InsertAfter("Đã mất")
                                        End If
                                        If (i > 0) Then
                                            objTable.Rows(i + 4).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot
                                        End If
                                    Next
                                    'objTable.Cell(_rowsTable, 1).Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(1).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(2).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(3).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    objTable.Columns(4).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                                    'Thực hiện Merge Cell Anh chị em ruột
                                    If (_rowsTable > 4) Then objTable.Cell(4, 1).Merge(objTable.Cell(_rowsTable, 1))
                                End If
                            End If
                        End Using
                    End If
                    objTable.Columns(1).PreferredWidth = 9
                    objTable.Columns(2).PreferredWidth = 18
                    objTable.Columns(3).PreferredWidth = 9
                    objTable.Columns(4).PreferredWidth = 59
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                '49c. Hồ sơ Giảm trừ gia cảnh của cán bộ
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.Font.Italic = 1
                objselection.TypeText("c. Hồ sơ giảm trừ gia cảnh" + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 11
                objselection.Font.Italic = 0
                strSQL = "SELECT b.Ten_Goi as QuanHe,a.HoTenNguoiPT,Convert(Varchar,a.TuNgay,103) as Tg_TuNgay,Convert(Varchar,a.DenNgay,103) as Tg_DenNgay,a.GhiChu FROM HS_GTGC a, DanhMuc b Where (a.IdQuanHe = b.id And b.id_goc = 23) And a.IdCanBo = '" + _RowId + "' Order by TuNgay Asc"
                Using db_gtgc As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_gtgc Is Nothing Or db_gtgc.Rows.Count = 0) Then
                        _rowsTable = 3
                    Else
                        _rowsTable = db_gtgc.Rows.Count + 2
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Quan hệ")
                    objTable.Cell(1, 2).Range.InsertAfter("Họ và tên người phụ thuộc")
                    objTable.Cell(1, 3).Range.InsertAfter("Thời gian kê khai")
                    objTable.Cell(2, 3).Range.InsertAfter("Từ ngày")
                    objTable.Cell(2, 4).Range.InsertAfter("Đến ngày")
                    objTable.Cell(1, 5).Range.InsertAfter("Ghi chú")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(2).Range.Font.Bold = 1
                    If Not (db_gtgc Is Nothing) Then
                        If (db_gtgc.Rows.Count > 0) Then
                            For i As Integer = 0 To db_gtgc.Rows.Count - 1
                                objTable.Cell(i + 3, 1).Range.InsertAfter(db_gtgc.Rows(i)("QuanHe").ToString().Trim())
                                objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 3, 2).Range.InsertAfter(db_gtgc.Rows(i)("HoTenNguoiPT").ToString().Trim())
                                objTable.Cell(i + 3, 2).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 3, 3).Range.InsertAfter(db_gtgc.Rows(i)("Tg_TuNgay").ToString().Trim())
                                objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 4).Range.InsertAfter(db_gtgc.Rows(i)("Tg_DenNgay").ToString().Trim())
                                objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 5).Range.InsertAfter(db_gtgc.Rows(i)("GhiChu").ToString().Trim())
                                objTable.Cell(i + 3, 5).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 10
                    objTable.Columns(2).PreferredWidth = 40
                    objTable.Columns(3).PreferredWidth = 10
                    objTable.Columns(4).PreferredWidth = 10
                    objTable.Columns(5).PreferredWidth = 28
                    objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
                    objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
                    objTable.Cell(1, 5).Merge(objTable.Cell(2, 5))
                    objTable.Cell(1, 3).Merge(objTable.Cell(1, 4))
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                ''50. Hồ sơ bổ nhiệm (Nếu có)
                'objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                'objselection.Paragraphs.SpaceBefore = 10
                'objselection.Font.Size = 12
                'objselection.Font.Italic = 1
                'objselection.Font.Bold = 1
                'objselection.TypeText("50. Hồ sơ bổ nhiệm (nếu có)" + vbCrLf)
                'objselection.Paragraphs.SpaceBefore = 3
                'objselection.Paragraphs.SpaceAfter = 3
                'objselection.Font.Size = 11
                'objselection.Font.Italic = 0
                'objselection.Font.Bold = 0
                ''Tạo một bảng gồm 5 cột và _rowsTable dòng
                'strSQL = "SELECT So_QD,Convert(Varchar,NgayKy_QD,103) As NgayKy,"
                'strSQL += "(Select Ten_Goi From DanhMuc Where id = IdChucVu_Moi and id_goc = 14) As ChucDanh,"
                'strSQL += "Convert(Varchar,NgayHL,103) As NgayHL,Convert(Varchar,NgayBoNhiem_TT,103) As NgayBoNhiem_TT,"
                'strSQL += "NguoiKy_QD FROM QDNhanSu WHERE IdLoaiQD IN (Select [id] From DanhMuc Where ma_so IN ('1502','1503','1504')) "
                'strSQL += " And IsKiemNhiem = 0 And IdCanBo = 'CNTT00000000016' Order By NgayKy_QD Asc"
                'Using db_hsbn As DataTable = _SqlHelper.SelectDBRows(strSQL)
                '    If (db_hsbn Is Nothing Or db_hsbn.Rows.Count = 0) Then
                '        _rowsTable = 3
                '    Else
                '        _rowsTable = db_hsbn.Rows.Count + 2
                '    End If
                '    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 6)
                '    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                '    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                '    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                '    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                '    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                '    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                '    objTable.Range.Paragraphs.SpaceBefore = 2
                '    objTable.Range.Paragraphs.SpaceAfter = 2
                '    'Add tiêu đề của các cột trong Bảng
                '    objTable.Cell(1, 1).Range.InsertAfter("Số QĐ")
                '    objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
                '    objTable.Cell(1, 3).Range.InsertAfter("Chức danh")
                '    objTable.Cell(1, 4).Range.InsertAfter("Hiệu lực thi hành")
                '    objTable.Cell(2, 4).Range.InsertAfter("Từ ngày")
                '    objTable.Cell(2, 5).Range.InsertAfter("Đến ngày")
                '    objTable.Cell(1, 6).Range.InsertAfter("Người ký QĐ")
                '    objTable.Rows(1).Range.Font.Bold = 1
                '    objTable.Rows(2).Range.Font.Bold = 1
                '    If Not (db_hsbn Is Nothing) Then
                '        If (db_hsbn.Rows.Count > 0) Then
                '            For i As Integer = 0 To db_hsbn.Rows.Count - 1
                '                objTable.Cell(i + 3, 1).Range.InsertAfter(db_hsbn.Rows(i)("So_QD").ToString().Trim())
                '                objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
                '                objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                '                objTable.Cell(i + 3, 2).Range.InsertAfter(db_hsbn.Rows(i)("NgayKy").ToString().Trim())
                '                objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                objTable.Cell(i + 3, 3).Range.InsertAfter(db_hsbn.Rows(i)("ChucDanh").ToString().Trim())
                '                objTable.Cell(i + 3, 3).Range.Paragraphs.LeftIndent = 3.0F
                '                objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                '                objTable.Cell(i + 3, 4).Range.InsertAfter(db_hsbn.Rows(i)("NgayHL").ToString().Trim())
                '                objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                objTable.Cell(i + 3, 5).Range.InsertAfter(db_hsbn.Rows(i)("NgayBoNhiem_TT").ToString().Trim())
                '                objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                '                objTable.Cell(i + 3, 6).Range.InsertAfter(db_hsbn.Rows(i)("NguoiKy_QD").ToString().Trim())
                '                objTable.Cell(i + 3, 6).Range.Paragraphs.LeftIndent = 3.0F
                '                objTable.Cell(i + 3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                '            Next
                '        End If
                '    End If
                '    objTable.Columns(1).PreferredWidth = 10
                '    objTable.Columns(2).PreferredWidth = 10
                '    objTable.Columns(3).PreferredWidth = 30
                '    objTable.Columns(4).PreferredWidth = 10
                '    objTable.Columns(5).PreferredWidth = 10
                '    objTable.Columns(6).PreferredWidth = 28
                '    objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
                '    objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
                '    objTable.Cell(1, 3).Merge(objTable.Cell(2, 3))
                '    objTable.Cell(1, 6).Merge(objTable.Cell(2, 6))
                '    objTable.Cell(1, 4).Merge(objTable.Cell(1, 5))
                '    objTable.LeftPadding = 0' -3
                '    objTable.RightPadding = 0' -4
                'End Using

                '51. Khen thưởng
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.Font.Bold = 1
                objselection.TypeText("50. Khen thưởng" + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 11
                objselection.Font.Bold = 0
                strSQL = "SELECT SoQD, Convert(Varchar,NgayQD,103) as NgayQD,(Select Ten_Goi From DanhMuc Where id = IdCapKT And id_goc = 4) As ThamQuyenQD,"
                strSQL += "(Select DanhHieu_HinhThuc From ThiDuaKhenThuong Where idTDKT = IdDanhHieuHinhThuc) As DanhHieu,'' As HinhThuc"
                strSQL += " FROM HS_khenthuong t1, HS_KhenThuong_CT t2 where t1.IdKhenThuong=t2.IdKhenthuong AND t2.IdCN_TT = '" + _RowId + "'"
                strSQL += " ORDER BY NgayQD ASC"
                Using db_khenthuong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_khenthuong Is Nothing Or db_khenthuong.Rows.Count = 0) Then
                        _rowsTable = 2
                    Else
                        _rowsTable = db_khenthuong.Rows.Count + 1
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 5)
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Số QĐ")
                    objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
                    objTable.Cell(1, 3).Range.InsertAfter("Thẩm quyền QĐ")
                    objTable.Cell(1, 4).Range.InsertAfter("Danh hiệu" + vbCrLf + "(từ chiến sỹ thi đua cơ sở trở lên)")
                    objTable.Cell(1, 5).Range.InsertAfter("Hình thức" + vbCrLf + "(từ Bằng khen)")
                    objTable.Rows(1).Range.Font.Bold = 1
                    If Not (db_khenthuong Is Nothing) Then
                        If (db_khenthuong.Rows.Count > 0) Then
                            For i As Integer = 0 To db_khenthuong.Rows.Count - 1
                                objTable.Cell(i + 2, 1).Range.InsertAfter(db_khenthuong.Rows(i)("SoQD").ToString().Trim())
                                objTable.Cell(i + 2, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 2).Range.InsertAfter(db_khenthuong.Rows(i)("NgayQD").ToString().Trim())
                                objTable.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 2, 3).Range.InsertAfter(db_khenthuong.Rows(i)("ThamQuyenQD").ToString().Trim())
                                objTable.Cell(i + 2, 3).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 4).Range.InsertAfter(db_khenthuong.Rows(i)("DanhHieu").ToString().Trim())
                                objTable.Cell(i + 2, 4).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 2, 5).Range.InsertAfter(db_khenthuong.Rows(i)("HinhThuc").ToString().Trim())
                                objTable.Cell(i + 2, 5).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 2, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 10
                    objTable.Columns(2).PreferredWidth = 10
                    objTable.Columns(3).PreferredWidth = 20
                    objTable.Columns(4).PreferredWidth = 38
                    objTable.Columns(5).PreferredWidth = 20
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                '52. Kỷ luật
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.Font.Bold = 1
                objselection.TypeText("51. Kỷ luật" + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 11
                objselection.Font.Bold = 0
                strSQL = "SELECT SoQD,Convert(Varchar,NgayQD,103) As NgayKyQD,(Select Ten_Goi From DanhMuc Where id = IdHinhThucKL And id_goc = 5) As HinhThucKL,"
                strSQL += "Convert(Varchar,TuNgay,103) As Tg_TuNgay,Convert(Varchar,DenNgay,103) As Tg_DenNgay, NguoiQD"
                strSQL += " FROM HS_KiLuat Where IdCanBo = '" + _RowId + "' Order By TuNgay"
                Using db_kiluat As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_kiluat Is Nothing Or db_kiluat.Rows.Count = 0) Then
                        _rowsTable = 3
                    Else
                        _rowsTable = db_kiluat.Rows.Count + 2
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, _rowsTable, 6)
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Số QĐ")
                    objTable.Cell(1, 2).Range.InsertAfter("Ngày, tháng")
                    objTable.Cell(1, 3).Range.InsertAfter("Hình thức kỷ luật")
                    objTable.Cell(1, 4).Range.InsertAfter("Hiệu lực thi hành")
                    objTable.Cell(2, 4).Range.InsertAfter("Từ ngày")
                    objTable.Cell(2, 5).Range.InsertAfter("Đến ngày")
                    objTable.Cell(1, 6).Range.InsertAfter("Người ký QĐ")
                    objTable.Rows(1).Range.Font.Bold = 1
                    objTable.Rows(2).Range.Font.Bold = 1
                    If Not (db_kiluat Is Nothing) Then
                        If (db_kiluat.Rows.Count > 0) Then
                            For i As Integer = 0 To db_kiluat.Rows.Count - 1
                                objTable.Cell(i + 3, 1).Range.InsertAfter(db_kiluat.Rows(i)("SoQD").ToString().Trim())
                                objTable.Cell(i + 3, 1).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 3, 2).Range.InsertAfter(db_kiluat.Rows(i)("NgayKyQD").ToString().Trim())
                                objTable.Cell(i + 3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 3).Range.InsertAfter(db_kiluat.Rows(i)("HinhThucKL").ToString().Trim())
                                objTable.Cell(i + 3, 3).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                                objTable.Cell(i + 3, 4).Range.InsertAfter(db_kiluat.Rows(i)("Tg_TuNgay").ToString().Trim())
                                objTable.Cell(i + 3, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 5).Range.InsertAfter(db_kiluat.Rows(i)("Tg_DenNgay").ToString().Trim())
                                objTable.Cell(i + 3, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                objTable.Cell(i + 3, 6).Range.InsertAfter(db_kiluat.Rows(i)("NguoiQD").ToString().Trim())
                                objTable.Cell(i + 3, 6).Range.Paragraphs.LeftIndent = 3.0F
                                objTable.Cell(i + 3, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Next
                        End If
                    End If
                    objTable.Columns(1).PreferredWidth = 10
                    objTable.Columns(2).PreferredWidth = 10
                    objTable.Columns(3).PreferredWidth = 30
                    objTable.Columns(4).PreferredWidth = 10
                    objTable.Columns(5).PreferredWidth = 10
                    objTable.Columns(6).PreferredWidth = 28
                    objTable.Cell(1, 1).Merge(objTable.Cell(2, 1))
                    objTable.Cell(1, 2).Merge(objTable.Cell(2, 2))
                    objTable.Cell(1, 3).Merge(objTable.Cell(2, 3))
                    objTable.Cell(1, 6).Merge(objTable.Cell(2, 6))
                    objTable.Cell(1, 4).Merge(objTable.Cell(1, 5))
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                '53. Kỷ luật
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 10
                objselection.Font.Size = 12
                objselection.TypeText("52. Hoàn cảnh kinh tế gia đình" + vbCrLf)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText("- Quá trình lương của bản thân:" + vbCrLf)
                strSQL = "SELECT (Select Mota from NgachLuong Where IdNgachLuong = b.IdNgachLuong And Status = 1) As NgachLuong,"
                strSQL += " a.IdLuongCB,b.BacLuong As BacLuong,a.HeSoLuong As HeSo,Convert(Varchar, Ngay_Huong, 103) as NgayHuong  From HS_LuongCB a, BacLuong b"
                strSQL += " Where a.IdCanBo = '" + _RowId + "' And (a.IdBacLuong = b.IdBacLuong And b.Status = 1) Order by a.Ngay_Huong "
                'strSQL = "SELECT (Select BacLuong From BacLuong Where IdBacLuong = HS_LuongCB.IdBacLuong) As BacLuong,"
                'strSQL += "(Select NgachLuong + ' ('+Mota+')' From NgachLuong Where IdNgachLuong in (Select IdNgachLuong From BacLuong Where IdBacLuong = HS_LuongCB.IdBacLuong)) As NgachLuong,"
                'strSQL += "CONVERT(VarChar,HeSoLuong) As HeSo,Convert(Varchar,Ngay_Huong,103) As NgayHuong,SoQD,Convert(Varchar,NgayQD,103) As NgayQD,NguoiQD"
                'strSQL += " FROM HS_LuongCB WHERE IdCanBo = '" + _RowId + "' Order By HeSoLuong Asc"
                Using db_qtluong As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (db_qtluong Is Nothing Or db_qtluong.Rows.Count = 0) Then
                        _rowsTable = 10
                    Else
                        _rowsTable = 6
                    End If
                    objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 4, _rowsTable)
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle  'Định dạng đường viền ngoài của bảng
                    objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDot      'Định dạng đường viền trong của bảng
                    objTable.Columns.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
                    objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                    objTable.Range.Paragraphs.SpaceBefore = 2
                    objTable.Range.Paragraphs.SpaceAfter = 2
                    For i As Integer = 1 To 4
                        objTable.Rows(i).Range.Paragraphs.LeftIndent = 3.0F
                    Next
                    objTable.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    'Add tiêu đề của các cột trong Bảng
                    objTable.Cell(1, 1).Range.InsertAfter("Tháng/Năm")
                    objTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    objTable.Cell(2, 1).Range.InsertAfter("Ngạch")
                    objTable.Cell(3, 1).Range.InsertAfter("Bậc lương")
                    objTable.Cell(4, 1).Range.InsertAfter("Hệ số lương")
                    If Not (db_qtluong Is Nothing) Then
                        If (db_qtluong.Rows.Count > 0) Then
                            Dim _columns As Byte = 6  'Dùng biến này làm biến đếm ngược
                            For i As Integer = db_qtluong.Rows.Count - 1 To 0 Step -1
                                objTable.Cell(1, _columns).Range.InsertAfter(db_qtluong.Rows(i)("NgayHuong").ToString().Trim().Substring(3))
                                objTable.Cell(2, _columns).Range.InsertAfter(db_qtluong.Rows(i)("NgachLuong").ToString().Trim())
                                objTable.Cell(3, _columns).Range.InsertAfter(db_qtluong.Rows(i)("BacLuong").ToString().Trim())
                                objTable.Cell(4, _columns).Range.InsertAfter(db_qtluong.Rows(i)("HeSo").ToString().Trim())
                                _columns -= 1
                                If _columns = 1 Then
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    objTable.LeftPadding = 0 ' -3
                    objTable.RightPadding = 0 ' -4
                End Using

                'objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                'objselection.Paragraphs.SpaceBefore = 10
                'objselection.Font.Size = 12
                'objselection.TypeText("- Nguồn thu nhập chính của gia đình (hàng năm):" + vbCrLf)
                'objselection.Paragraphs.SpaceBefore = 3
                'objselection.Paragraphs.SpaceAfter = 3
                'objselection.TypeText(vbTab + vbTab + "+ Lương: ..............................................................................................................." + vbCrLf)
                'objselection.TypeText(vbTab + vbTab + "+ Các nguồn khác: ........................................................................................................." + vbCrLf)
                'objselection.TypeText("- Nhà ở:" + vbTab + "+ Được cấp, được thuê, loại nhà: ....................," + " tổng diện tích sử dụng: .............." + "m2" + vbCrLf)
                'objselection.TypeText(vbTab + vbTab + "+ Nhà tự mua, tự xây, loại nhà: .....................," + " tổng diện tích sử dụng: ................" + "m2" + vbCrLf)
                'objselection.TypeText("- Đất ở:" + vbTab + "+ Đất được cấp: ................m2," + vbTab + vbTab + vbTab + "+ Đất được mua: ....................." + "m2" + vbCrLf)
                'objselection.TypeText("- Đất sản xuất, kinh doanh (tổng diện tích được cấp, tự mua, tự khai phá,...): ......................................" + vbCrLf)
                'objselection.TypeText("........................................................................................................................................................" + vbCrLf + vbCrLf + vbCrLf + vbCrLf)
                objselection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast)
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.Font.Size = 12
                objselection.Font.Bold = 1
                objselection.TypeText("- Nguồn thu nhập chính của gia đình (hàng năm):" + vbCrLf)
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                Dim vLuong As String = "......................................................................................................................."
                Dim vCacNguonKhac As String = "............................................................................................................"
                Dim vNhaO_Cap As String = "..................."
                Dim vNhaO_Cap_DT As String = ".............."
                Dim vNhaO_Mua As String = "..................."
                Dim vNhaO_Mua_DT As String = "................"
                Dim vDatO_Cap As String = "....................."
                Dim vDatO_Mua As String = "....................."
                Dim vDat_KD As String = ".........................................." + vbCrLf
                vDat_KD += ".................................................................................................................................................................."
                'If vNgachLuong <> "" Then vLuong = vNgachLuong
                'strSQL = " SELECT TOP 1 * FROM HS_ThuNhapGD WHERE IDCanBo='" + _RowId + "' and year(NgayKeKhai)= year(getdate()) order by NgayKeKhai desc"
                strSQL = " SELECT TOP 1 * FROM HS_ThuNhapGD WHERE IDCanBo='" + _RowId + "' order by NgayKeKhai desc"
                Using dtTNGD As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If (dtTNGD.Rows.Count > 0) Then
                        vLuong = dtTNGD.Rows(0)("Luong").ToString().Trim()
                        vCacNguonKhac = dtTNGD.Rows(0)("NguonKhac").ToString().Trim()
                        vNhaO_Cap = dtTNGD.Rows(0)("NhaO_DuocCap").ToString().Trim()
                        vNhaO_Cap_DT = dtTNGD.Rows(0)("NhaO_DuocCap_DT").ToString().Trim()
                        vNhaO_Mua = dtTNGD.Rows(0)("NhaO_TuMua").ToString().Trim()
                        vNhaO_Mua_DT = dtTNGD.Rows(0)("NhaO_TuMua_DT").ToString().Trim()
                        vDatO_Cap = dtTNGD.Rows(0)("DatO_DuocCap_DT").ToString().Trim()
                        vDatO_Mua = dtTNGD.Rows(0)("DatO_TuMua_DT").ToString().Trim()
                        vDat_KD = dtTNGD.Rows(0)("DatSXKT").ToString().Trim()
                    End If
                End Using
                objselection.TypeText(vbTab + vbTab + "+ Lương: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vLuong & vbCrLf)
                objselection.Font.Italic = 0
                objselection.TypeText(vbTab + vbTab + "+ Các nguồn khác: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vCacNguonKhac & vbCrLf)
                objselection.Font.Italic = 0

                objselection.Font.Bold = 1
                objselection.TypeText("- Nhà ở:")
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText(vbTab + "+ Được cấp, được thuê, loại nhà: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Cap)
                objselection.Font.Italic = 0
                objselection.TypeText(", tổng diện tích sử dụng: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Cap_DT + "m2" + vbCrLf)
                objselection.Font.Italic = 0
                objselection.TypeText(vbTab + vbTab + "+ Nhà tự mua, tự xây, loại nhà: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Mua)
                objselection.Font.Italic = 0
                objselection.TypeText(", tổng diện tích sử dụng: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vNhaO_Mua_DT + "m2" + vbCrLf)
                objselection.Font.Italic = 0

                objselection.Font.Bold = 1
                objselection.TypeText("- Đất ở:")
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText(vbTab + "+ Đất được cấp: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vDatO_Cap)
                objselection.Font.Italic = 0
                objselection.TypeText("m2," + vbTab + vbTab + vbTab + "+ Đất tự mua: ")
                objselection.Font.Italic = 1
                objselection.TypeText(vDatO_Mua + "m2" + vbCrLf)

                objselection.Font.Bold = 1
                objselection.TypeText("- Đất sản xuất, kinh doanh ")
                objselection.Font.Bold = 0
                objselection.Paragraphs.SpaceBefore = 3
                objselection.Paragraphs.SpaceAfter = 3
                objselection.TypeText("(Tổng diện tích được cấp, tự mua, tự khai phá,...): ")
                objselection.Font.Italic = 1
                objselection.TypeText(vDat_KD)
                objselection.Font.Italic = 0

                objselection.Font.Size = 13
                objTable = objWordApp.ActiveDocument.Tables.Add(objWordApp.Selection.Range, 5, 2)
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                objTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                objTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent
                objTable.PreferredWidth = 100
                objTable.Range.Paragraphs.SpaceBefore = 0
                objTable.Range.Paragraphs.SpaceAfter = 0

                'objTable.Range.Font.Bold = 0
                objTable.Cell(1, 1).Range.InsertAfter("Người khai")
                objTable.Cell(1, 1).Range.Font.Bold = 1
                objTable.Cell(2, 1).Range.InsertAfter("Tôi xin cam đoan những")
                objTable.Cell(3, 1).Range.InsertAfter("lời khai trên đây là đúng sự thật")
                objTable.Cell(4, 1).Range.InsertAfter("(Ký tên)")
                objTable.Cell(5, 1).Range.InsertAfter(vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + dr("HoTen").ToString().Trim().ToUpper())
                objTable.Cell(5, 1).Range.Font.Bold = 1
                _sVal = _sDiaban + ", Ngày " + DateTime.Now.ToString("dd") + " tháng " + DateTime.Now.ToString("MM") + " năm " + DateTime.Now.ToString("yyyy")
                objTable.Cell(2, 2).Range.InsertAfter(_sVal)
                objTable.Cell(2, 2).Range.Font.Italic = 1
                objTable.Cell(3, 2).Range.InsertAfter("Xác nhận của cơ quan quản lý")
                objTable.Cell(3, 2).Range.Font.Bold = 1
                objTable.Cell(5, 2).Range.InsertAfter(vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf)
                objTable.Cell(1, 1).Merge(objTable.Cell(4, 1))
                objTable.Cell(1, 2).Merge(objTable.Cell(4, 2))
                objTable.LeftPadding = 0 ' -3
                objTable.RightPadding = 0 ' -4


                'Định dạng thêm về trang word khi xuất ra
                objWordApp.ActiveDocument.PageSetup.HeaderDistance = 3
                objWordApp.ActiveDocument.PageSetup.FooterDistance = 3
                objWordApp.ActiveDocument.PageSetup.LeftMargin = 25         'Quy ước:   72 Points = 1 Inch.
                objWordApp.ActiveDocument.PageSetup.RightMargin = 20
                objWordApp.ActiveDocument.PageSetup.TopMargin = 25
                objWordApp.ActiveDocument.PageSetup.BottomMargin = 25
                objWordApp.ActiveDocument.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4
                'objWord.ActiveDocument.PageSetup.HeaderDistance = 100
                objWordApp.ActiveDocument.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait   'Định dạng Khổ dấy dọc
                objWordApp.ActiveDocument.ShowGrammaticalErrors = False     'Không cho hiển thị các đường viền check chính tả mầu xanh
                objWordApp.ActiveDocument.ShowSpellingErrors = False        'Không cho hiển thị các đường viền check chính tả mầu đỏ

                'Dim sFileName As String = SaveToFile("Document File|*.doc|XML Document File|*.docx", "Mau01_" + dr("MaCB").ToString().Trim().ToUpper() + "_" + getFullName(_RowId.ToString().Trim(), False))
                'If sFileName <> "" Then
                '    objDocument.SaveAs(sFileName)
                'End If

                SaveToWord(objDocument, "Mau01_" + dr("MaCB").ToString().Trim().ToUpper() + "_" + getFullName(_RowId.ToString().Trim(), False), True)
                Clipboard.Clear()

                objWordApp.Quit(False)
                objWordApp = Nothing
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm trả về danh sách bản ghi theo Id truyền vào
    ''' </summary>
    ''' <param name="_IdCanBo">Id hồ sơ cán bộ</param>
    ''' <param name="_Id">Chỉ số xác định bản ghi</param>
    ''' <param name="state">Chỉ số xác định hồ sơ cần trả dữ liệu
    '''                             1: Hồ sơ công tác
    '''                             2: Hồ sơ xuất ngoại
    '''                             3: Hồ sơ hộ chiếu
    '''                             4: Hồ sơ Tham gia lực lượng vũ trang
    '''                             5: Hồ sơ cũ (Thông tin của cán bộ khi chưa thoát ly)
    '''                             6: Hồ sơ Cán bộ học hàm 
    '''                             7: Hồ sơ Cán bộ học vị 
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDocument(ByVal _IdCanBo As String, ByVal _Id As String, ByVal state As Byte) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand
            Select Case state
                Case 1
                    command = New SqlCommand("HS_CongTac_GetForId", connection)

                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdHSCongTac", SqlDbType.VarChar))
                    command.Parameters("@_IdHSCongTac").Value = _Id
                Case 2
                    command = New SqlCommand("HS_XuatNgoai_GetForId", connection)

                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdXuatNgoai", SqlDbType.VarChar))
                    command.Parameters("@_IdXuatNgoai").Value = _Id
                Case 3
                    command = New SqlCommand("HS_HoChieu_GetForId", connection)

                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdHoChieu", SqlDbType.VarChar))
                    command.Parameters("@_IdHoChieu").Value = _Id

                Case 4
                    command = New SqlCommand("HS_LLVT_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdHSLLVT", SqlDbType.VarChar))
                    command.Parameters("@_IdHSLLVT").Value = _Id
                Case 5
                    command = New SqlCommand("HS_Cu_GetForId", connection)

                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdHSCu", SqlDbType.VarChar))
                    command.Parameters("@_IdHSCu").Value = _Id

                Case 6
                    command = New SqlCommand("CB_HocHam_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdCB_HocHam", SqlDbType.VarChar))
                    command.Parameters("@_IdCB_HocHam").Value = _Id

                Case 7
                    command = New SqlCommand("CB_HocVi_GetForId", connection)

                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                    command.Parameters("@_IdCanBo").Value = _IdCanBo

                    command.Parameters.Add(New SqlParameter("@_IdCB_HocVi", SqlDbType.VarChar))
                    command.Parameters("@_IdCB_HocVi").Value = _Id

                Case Else
            End Select

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "TableRow")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("TableRow")
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
    ''' Hàm thực hiện trả về danh sách lao động của cán bộ tập sự - Theo id cán bộ tập sự truyền vào
    ''' </summary>
    ''' <param name="_ContractId">Id cán bộ tập sự</param>
    ''' <returns>Danh sách thông tin Hợp đồng lao động</returns>
    ''' <remarks></remarks>
    'Public Function GetAll_Contracts(ByVal _ContractId As String) As DataTable
    '    Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
    '    If (connection Is Nothing) Then
    '        MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
    '        Return Nothing
    '    End If

    '    Try
    '        Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_GetForIdCanBo", connection)
    '        command.CommandType = CommandType.StoredProcedure

    '        command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
    '        command.Parameters("@_IdCanBo").Value = _ContractId

    '        Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
    '            Dim ds As DataSet = New DataSet
    '            mydap.Fill(ds, "HDLD_TS")
    '            If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
    '                Return Nothing
    '            End If
    '            Return ds.Tables("HDLD_TS")
    '        End Using
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        If (connection.State = System.Data.ConnectionState.Open) Then
    '            connection.Close()
    '        End If
    '    End Try
    '    Return Nothing
    'End Function

    ' ''' <summary>
    ' ''' Hàm trả về Bản ghi dữ liệu liên quan đến Hợp đồng của cán bộ tập sự theo id truyền vào
    ' ''' </summary>
    ' ''' <param name="_Code">Id của hợp đồng lao động</param>
    ' ''' <returns>Bản ghi cần lấy dữ liệu</returns>
    ' ''' <remarks></remarks>
    'Public Function GetContracts(ByVal _Code As String) As DataRow
    '    Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
    '    If (connection Is Nothing) Then
    '        MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
    '        Return Nothing
    '    End If
    '    Try
    '        Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_GetForId", connection)
    '        command.CommandType = CommandType.StoredProcedure

    '        command.Parameters.Add(New SqlParameter("@_IdCBTS_HDLD", SqlDbType.VarChar))
    '        command.Parameters("@_IdCBTS_HDLD").Value = _Code

    '        Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
    '            Dim ds As DataSet = New DataSet
    '            mydap.Fill(ds, "HDLD_TS")
    '            If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
    '                Return Nothing
    '            End If
    '            Return ds.Tables("HDLD_TS").Rows(0)
    '        End Using
    '    Finally
    '        If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
    '    End Try
    'End Function

    ''' <summary>
    ''' Hàm lấy danh sách bản ghi Hợp đồng lao động của Cán bộ tập sự/Cán bộ ngắn hạn
    ''' </summary>
    ''' <param name="pIdCanBo">Chỉ số xác định Cán bộ ngắn hạn</param>
    ''' <param name="pIdCBTS_HDLD">Chỉ số xác định bản ghi Hợp đồng LĐ ngắn hạn</param>
    ''' <param name="pChiNhanhId">Id Chi nhánh</param>
    ''' <param name="pDonVi_Cd">Mã đơn vị (POS)</param>
    ''' <returns>Danh sách bản ghi</returns>
    ''' <remarks></remarks>
    Public Function GetHSCB_TS_HDLD_GetSearch(ByVal pIdCanBo As String, ByVal pIdCBTS_HDLD As String, ByVal pChiNhanhId As Integer, ByVal pDonVi_Cd As String) As DataTable
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_GetSearch", connection)
            command.CommandType = CommandType.StoredProcedure

            command.Parameters.Add(New SqlParameter("@pIdCanBo", SqlDbType.NVarChar, 16))
            command.Parameters("@pIdCanBo").Value = pIdCanBo

            command.Parameters.Add(New SqlParameter("@pIdCBTS_HDLD", SqlDbType.NVarChar, 16))
            command.Parameters("@pIdCBTS_HDLD").Value = pIdCBTS_HDLD

            command.Parameters.Add(New SqlParameter("@pChiNhanhId", SqlDbType.Int))
            command.Parameters("@pChiNhanhId").Value = pChiNhanhId

            command.Parameters.Add(New SqlParameter("@pDonVi_Cd", SqlDbType.VarChar, 6))
            command.Parameters("@pDonVi_Cd").Value = pDonVi_Cd

            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HSCB_TS_HDLD_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HSCB_TS_HDLD_TMP")
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

#End Region

#Region "---> Fucntions cập nhật dữ liệu Hồ sơ - Và thông tin liên quan <---"
    'Hàm cập nhật thông tin  - Hồ cơ cán bộ
    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu - Hồ sơ cán bộ vào CSDL
    ''' </summary>
    ''' <param name="obj_canbo">Đối tượng Hồ sơ cán bộ</param>
    ''' <remarks></remarks>
    Public Function Insert_Human(ByVal obj_canbo As HS_CanBo) As String
        Dim command As SqlCommand = New SqlCommand("HS_CanBo_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdCanBo", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_MaCB", obj_canbo.MaCB))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_canbo.HoTen))
        command.Parameters.Add(New SqlParameter("@_IdDonVi", obj_canbo.IdDonVi))
        command.Parameters.Add(New SqlParameter("@_TenThuongGoi", obj_canbo.TenThuongGoi))
        command.Parameters.Add(New SqlParameter("@_BiDanh", obj_canbo.BiDanh))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_canbo.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NgaySinh", obj_canbo.NgaySinh))

        command.Parameters.Add(New SqlParameter("@_IdNS_Tinh", obj_canbo.IdNS_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNS_Huyen", obj_canbo.IdNS_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNS_Xa", obj_canbo.IdNS_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNS_Thon", obj_canbo.IdNS_Thon))
        command.Parameters.Add(New SqlParameter("@_NS_DChi", obj_canbo.NS_DChi))

        command.Parameters.Add(New SqlParameter("@_IdNQ_Tinh", obj_canbo.IdNQ_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Huyen", obj_canbo.IdNQ_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Xa", obj_canbo.IdNQ_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Thon", obj_canbo.IdNQ_Thon))
        command.Parameters.Add(New SqlParameter("@_NQ_DChi", obj_canbo.NQ_DChi))

        command.Parameters.Add(New SqlParameter("@_IdThT_Tinh", obj_canbo.IdThT_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdThT_Huyen", obj_canbo.IdThT_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdThT_Xa", obj_canbo.IdThT_Xa))
        command.Parameters.Add(New SqlParameter("@_IdThT_Thon", obj_canbo.IdThT_Thon))
        command.Parameters.Add(New SqlParameter("@_ThT_Diachi", obj_canbo.ThT_Diachi))
        command.Parameters.Add(New SqlParameter("@_ThT_Dienthoai", obj_canbo.ThT_Dienthoai))

        command.Parameters.Add(New SqlParameter("@_IdTTr_Tinh", obj_canbo.IdTTr_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Huyen", obj_canbo.IdTTr_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Xa", obj_canbo.IdTTr_Xa))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Thon", obj_canbo.IdTTr_Thon))
        command.Parameters.Add(New SqlParameter("@_TTr_Diachi", obj_canbo.TTr_Diachi))
        command.Parameters.Add(New SqlParameter("@_TTr_Dienthoai", obj_canbo.TTr_Dienthoai))

        command.Parameters.Add(New SqlParameter("@_IdQuocTich", obj_canbo.IdQuocTich))
        command.Parameters.Add(New SqlParameter("@_IdDanToc", obj_canbo.IdDanToc))
        command.Parameters.Add(New SqlParameter("@_IdTonGiao", obj_canbo.IdTonGiao))
        command.Parameters.Add(New SqlParameter("@_IdThanhPhanGD", obj_canbo.IdThanhPhanGD))

        command.Parameters.Add(New SqlParameter("@_NhomMau", obj_canbo.NhomMau))
        command.Parameters.Add(New SqlParameter("@_CMT_So", obj_canbo.CMT_So))
        command.Parameters.Add(New SqlParameter("@_CMT_NgayCap", obj_canbo.CMT_NgayCap))
        command.Parameters.Add(New SqlParameter("@_CMT_NoiCap", obj_canbo.CMT_NoiCap))

        command.Parameters.Add(New SqlParameter("@_IdUT_GDinh", obj_canbo.IdUT_GDinh))
        command.Parameters.Add(New SqlParameter("@_IdUT_BThan", obj_canbo.IdUT_BThan))

        If (CType(obj_canbo.Ngay_ThamNien, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_ThamNien", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_ThamNien", obj_canbo.Ngay_ThamNien))
        End If

        If (CType(obj_canbo.Ngay_NH, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_NH", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_NH", obj_canbo.Ngay_NH))
        End If

        If (CType(obj_canbo.Ngay_VBSP, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_VBSP", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_VBSP", obj_canbo.Ngay_VBSP))
        End If

        If (CType(obj_canbo.Ngay_BienChe, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_BienChe", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_BienChe", obj_canbo.Ngay_BienChe))
        End If

        If (CType(obj_canbo.Ngay_CQ, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_CQ", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_CQ", obj_canbo.Ngay_CQ))
        End If
        command.Parameters.Add(New SqlParameter("@_Emai", obj_canbo.Email))
        command.Parameters.Add(New SqlParameter("@_DienThoai_NR", obj_canbo.DienThoai_NR))
        command.Parameters.Add(New SqlParameter("@_DienThoai_DD", obj_canbo.DienThoai_DD))
        command.Parameters.Add(New SqlParameter("@_DienThoai_CQ", obj_canbo.DienThoai_CQ))
        command.Parameters.Add(New SqlParameter("@_SoFax", obj_canbo.SoFax))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoVH", obj_canbo.IdTrinhDoVH))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoCT", obj_canbo.IdTrinhDoCT))
        command.Parameters.Add(New SqlParameter("@_HonNhan_Cd", obj_canbo.HonNhan_Cd))
        command.Parameters.Add(New SqlParameter("@_SoTruong_CT", obj_canbo.SoTruong_CT))
        command.Parameters.Add(New SqlParameter("@_CV_Lau", obj_canbo.CV_Lau))

        If (CType(obj_canbo.CM_Ngay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", obj_canbo.CM_Ngay))
        End If

        command.Parameters.Add(New SqlParameter("@_CM_ToChuc", obj_canbo.CM_ToChuc))
        command.Parameters.Add(New SqlParameter("@_DacDiem_BT", obj_canbo.DacDiem_BT))
        command.Parameters.Add(New SqlParameter("@_QuanHe_Nguoi_NN", obj_canbo.QuanHe_Nguoi_NN))

        command.Parameters.Add(New SqlParameter("@_NH_MaKH", obj_canbo.NH_MaKH))
        command.Parameters.Add(New SqlParameter("@_NH_SoTK", obj_canbo.NH_SoTK))
        command.Parameters.Add(New SqlParameter("@_NH_TenNH", obj_canbo.NH_TenNH))
        command.Parameters.Add(New SqlParameter("@_MaSoThue", obj_canbo.MaSoThue))
        command.Parameters.Add(New SqlParameter("@_BHXH_SoSo", obj_canbo.BHXH_SoSo))

        If (CType(obj_canbo.BHXH_NgaySo, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgaySo", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgaySo", obj_canbo.BHXH_NgaySo))
        End If

        If (CType(obj_canbo.BHXH_NgayBatDau, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayBatDau", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayBatDau", obj_canbo.BHXH_NgayBatDau))
        End If

        command.Parameters.Add(New SqlParameter("@_BHXH_NoiLam", obj_canbo.BHXH_NoiLam))
        command.Parameters.Add(New SqlParameter("@_AnhThe", IIf(obj_canbo.AnhThe Is Nothing, System.Data.SqlTypes.SqlBytes.Null, obj_canbo.AnhThe)))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_canbo.GhiChu))
        command.Parameters.Add(New SqlParameter("@_IdNew", obj_canbo.IdNew))
        Try
            _SqlHelper.executeSQL(command)
            obj_canbo.IdCanBo = command.Parameters("@_IdCanBo").Value.ToString
            Return obj_canbo.IdCanBo
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới nhân sự", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu Hồ sơ nhân sự
    ''' </summary>
    ''' <param name="obj_canbo"></param>
    ''' <remarks></remarks>
    Public Sub Update_Human(ByVal obj_canbo As HS_CanBo)
        Dim command As SqlCommand = New SqlCommand("HS_CanBo_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_canbo.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_MaCB", obj_canbo.MaCB))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_canbo.HoTen))
        command.Parameters.Add(New SqlParameter("@_IdDonVi", obj_canbo.IdDonVi))
        command.Parameters.Add(New SqlParameter("@_TenThuongGoi", obj_canbo.TenThuongGoi))
        command.Parameters.Add(New SqlParameter("@_BiDanh", obj_canbo.BiDanh))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_canbo.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NgaySinh", obj_canbo.NgaySinh))

        command.Parameters.Add(New SqlParameter("@_IdNS_Tinh", obj_canbo.IdNS_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNS_Huyen", obj_canbo.IdNS_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNS_Xa", obj_canbo.IdNS_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNS_Thon", obj_canbo.IdNS_Thon))
        command.Parameters.Add(New SqlParameter("@_NS_DChi", obj_canbo.NS_DChi))

        command.Parameters.Add(New SqlParameter("@_IdNQ_Tinh", obj_canbo.IdNQ_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Huyen", obj_canbo.IdNQ_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Xa", obj_canbo.IdNQ_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Thon", obj_canbo.IdNQ_Thon))
        command.Parameters.Add(New SqlParameter("@_NQ_DChi", obj_canbo.NQ_DChi))

        command.Parameters.Add(New SqlParameter("@_IdThT_Tinh", obj_canbo.IdThT_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdThT_Huyen", obj_canbo.IdThT_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdThT_Xa", obj_canbo.IdThT_Xa))
        command.Parameters.Add(New SqlParameter("@_IdThT_Thon", obj_canbo.IdThT_Thon))
        command.Parameters.Add(New SqlParameter("@_ThT_Diachi", obj_canbo.ThT_Diachi))
        command.Parameters.Add(New SqlParameter("@_ThT_Dienthoai", obj_canbo.ThT_Dienthoai))

        command.Parameters.Add(New SqlParameter("@_IdTTr_Tinh", obj_canbo.IdTTr_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Huyen", obj_canbo.IdTTr_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Xa", obj_canbo.IdTTr_Xa))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Thon", obj_canbo.IdTTr_Thon))
        command.Parameters.Add(New SqlParameter("@_TTr_Diachi", obj_canbo.TTr_Diachi))
        command.Parameters.Add(New SqlParameter("@_TTr_Dienthoai", obj_canbo.TTr_Dienthoai))

        command.Parameters.Add(New SqlParameter("@_IdQuocTich", obj_canbo.IdQuocTich))
        command.Parameters.Add(New SqlParameter("@_IdDanToc", obj_canbo.IdDanToc))
        command.Parameters.Add(New SqlParameter("@_IdTonGiao", obj_canbo.IdTonGiao))
        command.Parameters.Add(New SqlParameter("@_IdThanhPhanGD", obj_canbo.IdThanhPhanGD))

        command.Parameters.Add(New SqlParameter("@_NhomMau", obj_canbo.NhomMau))
        command.Parameters.Add(New SqlParameter("@_CMT_So", obj_canbo.CMT_So))
        command.Parameters.Add(New SqlParameter("@_CMT_NgayCap", obj_canbo.CMT_NgayCap))
        command.Parameters.Add(New SqlParameter("@_CMT_NoiCap", obj_canbo.CMT_NoiCap))

        command.Parameters.Add(New SqlParameter("@_IdUT_GDinh", obj_canbo.IdUT_GDinh))
        command.Parameters.Add(New SqlParameter("@_IdUT_BThan", obj_canbo.IdUT_BThan))

        'Ngày tính thâm niên
        If (CType(obj_canbo.Ngay_ThamNien, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_ThamNien", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_ThamNien", obj_canbo.Ngay_ThamNien))
        End If
        'Ngày vào ngành ngân hàng
        If (CType(obj_canbo.Ngay_NH, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_NH", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_NH", obj_canbo.Ngay_NH))
        End If
        'Ngày vào Ngân hàng chính sách xã hội
        If (CType(obj_canbo.Ngay_VBSP, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_VBSP", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_VBSP", obj_canbo.Ngay_VBSP))
        End If
        'Ngày biên chế trong NHCSXH
        If (CType(obj_canbo.Ngay_BienChe, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_BienChe", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_BienChe", obj_canbo.Ngay_BienChe))
        End If
        'Ngày vào cơ quan công tác hiện tại
        If (CType(obj_canbo.Ngay_CQ, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_CQ", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_CQ", obj_canbo.Ngay_CQ))
        End If

        command.Parameters.Add(New SqlParameter("@_Emai", obj_canbo.Email))
        command.Parameters.Add(New SqlParameter("@_DienThoai_NR", obj_canbo.DienThoai_NR))
        command.Parameters.Add(New SqlParameter("@_DienThoai_DD", obj_canbo.DienThoai_DD))
        command.Parameters.Add(New SqlParameter("@_DienThoai_CQ", obj_canbo.DienThoai_CQ))
        command.Parameters.Add(New SqlParameter("@_SoFax", obj_canbo.SoFax))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoVH", obj_canbo.IdTrinhDoVH))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoCT", obj_canbo.IdTrinhDoCT))
        command.Parameters.Add(New SqlParameter("@_HonNhan_Cd", obj_canbo.HonNhan_Cd))
        command.Parameters.Add(New SqlParameter("@_SoTruong_CT", obj_canbo.SoTruong_CT))
        command.Parameters.Add(New SqlParameter("@_CV_Lau", obj_canbo.CV_Lau))

        If (CType(obj_canbo.CM_Ngay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", obj_canbo.CM_Ngay))
        End If
        command.Parameters.Add(New SqlParameter("@_CM_ToChuc", obj_canbo.CM_ToChuc))
        command.Parameters.Add(New SqlParameter("@_DacDiem_BT", obj_canbo.DacDiem_BT))
        command.Parameters.Add(New SqlParameter("@_QuanHe_Nguoi_NN", obj_canbo.QuanHe_Nguoi_NN))

        command.Parameters.Add(New SqlParameter("@_NH_MaKH", obj_canbo.NH_MaKH))
        command.Parameters.Add(New SqlParameter("@_NH_SoTK", obj_canbo.NH_SoTK))
        command.Parameters.Add(New SqlParameter("@_NH_TenNH", obj_canbo.NH_TenNH))
        command.Parameters.Add(New SqlParameter("@_MaSoThue", obj_canbo.MaSoThue))

        command.Parameters.Add(New SqlParameter("@_BHXH_SoSo", obj_canbo.BHXH_SoSo))
        If (CType(obj_canbo.BHXH_NgaySo, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgaySo", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgaySo", obj_canbo.BHXH_NgaySo))
        End If

        If (CType(obj_canbo.BHXH_NgayBatDau, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayBatDau", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayBatDau", obj_canbo.BHXH_NgayBatDau))
        End If
        command.Parameters.Add(New SqlParameter("@_BHXH_NoiLam", obj_canbo.BHXH_NoiLam))
        command.Parameters.Add(New SqlParameter("@_AnhThe", IIf(obj_canbo.AnhThe Is Nothing, System.Data.SqlTypes.SqlBytes.Null, obj_canbo.AnhThe)))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_canbo.GhiChu))
        command.Parameters.Add(New SqlParameter("@_IdNew", obj_canbo.IdNew))
        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật nhân sự", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu Hồ sơ nhân sự
    ''' </summary>
    ''' <param name="obj_canbo"></param>
    ''' <remarks></remarks>
    Public Sub Update_HumanNotFull(ByVal obj_canbo As HS_CanBo)
        Dim command As SqlCommand = New SqlCommand("HS_CanBo_UpdateNotFull")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_canbo.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_MaCB", obj_canbo.MaCB))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_canbo.HoTen))
        command.Parameters.Add(New SqlParameter("@_IdDonVi", obj_canbo.IdDonVi))
        command.Parameters.Add(New SqlParameter("@_TenThuongGoi", obj_canbo.TenThuongGoi))
        command.Parameters.Add(New SqlParameter("@_BiDanh", obj_canbo.BiDanh))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_canbo.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NgaySinh", obj_canbo.NgaySinh))

        command.Parameters.Add(New SqlParameter("@_IdNS_Tinh", obj_canbo.IdNS_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNS_Huyen", obj_canbo.IdNS_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNS_Xa", obj_canbo.IdNS_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNS_Thon", obj_canbo.IdNS_Thon))
        command.Parameters.Add(New SqlParameter("@_NS_DChi", obj_canbo.NS_DChi))

        command.Parameters.Add(New SqlParameter("@_IdNQ_Tinh", obj_canbo.IdNQ_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Huyen", obj_canbo.IdNQ_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Xa", obj_canbo.IdNQ_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Thon", obj_canbo.IdNQ_Thon))
        command.Parameters.Add(New SqlParameter("@_NQ_DChi", obj_canbo.NQ_DChi))

        command.Parameters.Add(New SqlParameter("@_IdThT_Tinh", obj_canbo.IdThT_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdThT_Huyen", obj_canbo.IdThT_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdThT_Xa", obj_canbo.IdThT_Xa))
        command.Parameters.Add(New SqlParameter("@_IdThT_Thon", obj_canbo.IdThT_Thon))
        command.Parameters.Add(New SqlParameter("@_ThT_Diachi", obj_canbo.ThT_Diachi))
        command.Parameters.Add(New SqlParameter("@_ThT_Dienthoai", obj_canbo.ThT_Dienthoai))

        command.Parameters.Add(New SqlParameter("@_IdTTr_Tinh", obj_canbo.IdTTr_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Huyen", obj_canbo.IdTTr_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Xa", obj_canbo.IdTTr_Xa))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Thon", obj_canbo.IdTTr_Thon))
        command.Parameters.Add(New SqlParameter("@_TTr_Diachi", obj_canbo.TTr_Diachi))
        command.Parameters.Add(New SqlParameter("@_TTr_Dienthoai", obj_canbo.TTr_Dienthoai))

        command.Parameters.Add(New SqlParameter("@_IdQuocTich", obj_canbo.IdQuocTich))
        command.Parameters.Add(New SqlParameter("@_IdDanToc", obj_canbo.IdDanToc))
        command.Parameters.Add(New SqlParameter("@_IdTonGiao", obj_canbo.IdTonGiao))
        command.Parameters.Add(New SqlParameter("@_IdThanhPhanGD", obj_canbo.IdThanhPhanGD))

        command.Parameters.Add(New SqlParameter("@_NhomMau", obj_canbo.NhomMau))
        command.Parameters.Add(New SqlParameter("@_CMT_So", obj_canbo.CMT_So))
        command.Parameters.Add(New SqlParameter("@_CMT_NgayCap", obj_canbo.CMT_NgayCap))
        command.Parameters.Add(New SqlParameter("@_CMT_NoiCap", obj_canbo.CMT_NoiCap))

        command.Parameters.Add(New SqlParameter("@_IdUT_GDinh", obj_canbo.IdUT_GDinh))
        command.Parameters.Add(New SqlParameter("@_IdUT_BThan", obj_canbo.IdUT_BThan))

        'Ngày tính thâm niên
        If (CType(obj_canbo.Ngay_ThamNien, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_ThamNien", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_ThamNien", obj_canbo.Ngay_ThamNien))
        End If
        'Ngày vào ngành ngân hàng
        If (CType(obj_canbo.Ngay_NH, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_NH", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_NH", obj_canbo.Ngay_NH))
        End If
        'Ngày vào Ngân hàng chính sách xã hội
        If (CType(obj_canbo.Ngay_VBSP, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_VBSP", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_VBSP", obj_canbo.Ngay_VBSP))
        End If
        'Ngày biên chế trong NHCSXH
        If (CType(obj_canbo.Ngay_BienChe, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_BienChe", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_BienChe", obj_canbo.Ngay_BienChe))
        End If
        'Ngày vào cơ quan công tác hiện tại
        If (CType(obj_canbo.Ngay_CQ, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Ngay_CQ", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Ngay_CQ", obj_canbo.Ngay_CQ))
        End If

        command.Parameters.Add(New SqlParameter("@_Emai", obj_canbo.Email))
        command.Parameters.Add(New SqlParameter("@_DienThoai_NR", obj_canbo.DienThoai_NR))
        command.Parameters.Add(New SqlParameter("@_DienThoai_DD", obj_canbo.DienThoai_DD))
        command.Parameters.Add(New SqlParameter("@_DienThoai_CQ", obj_canbo.DienThoai_CQ))
        command.Parameters.Add(New SqlParameter("@_SoFax", obj_canbo.SoFax))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoVH", obj_canbo.IdTrinhDoVH))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoCT", obj_canbo.IdTrinhDoCT))
        command.Parameters.Add(New SqlParameter("@_SoTruong_CT", obj_canbo.SoTruong_CT))
        command.Parameters.Add(New SqlParameter("@_CV_Lau", obj_canbo.CV_Lau))
        If (CType(obj_canbo.CM_Ngay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", obj_canbo.CM_Ngay))
        End If
        command.Parameters.Add(New SqlParameter("@_CM_ToChuc", obj_canbo.CM_ToChuc))
        command.Parameters.Add(New SqlParameter("@_DacDiem_BT", obj_canbo.DacDiem_BT))
        command.Parameters.Add(New SqlParameter("@_QuanHe_Nguoi_NN", obj_canbo.QuanHe_Nguoi_NN))
        command.Parameters.Add(New SqlParameter("@_BHXH_SoSo", obj_canbo.BHXH_SoSo))
        If (CType(obj_canbo.BHXH_NgaySo, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgaySo", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgaySo", obj_canbo.BHXH_NgaySo))
        End If

        If (CType(obj_canbo.BHXH_NgayBatDau, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayBatDau", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayBatDau", obj_canbo.BHXH_NgayBatDau))
        End If
        command.Parameters.Add(New SqlParameter("@_BHXH_NoiLam", obj_canbo.BHXH_NoiLam))
        command.Parameters.Add(New SqlParameter("@_AnhThe", obj_canbo.AnhThe))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_canbo.GhiChu))
        command.Parameters.Add(New SqlParameter("@_IdNew", obj_canbo.IdNew))
        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật nhân sự", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Id cán bộ truyền vào</param>
    ''' <remarks></remarks>
    Public Sub Delete_Human(ByVal _IdCanBo As String)
        Dim command As SqlCommand = New SqlCommand("HS_CanBo_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", SqlDbType.Int))
        command.Parameters("@_IdCanBo").Value = _IdCanBo
        command.Parameters("@IdBranch").Value = IdDONVI

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá bỏ hồ sơ cán bộ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật chỉ số xác định bản ghi của cán bộ tập sự vào Hồ sơ cán bộ
    ''' </summary>
    ''' <param name="_IdCanBoTS">Chỉ số xác định id cán bộ tập sự</param>
    ''' <param name="_IdCB">Chỉ số xác định id cán bộ</param>
    ''' <param name="_State">Chỉ số trạng thái: 1: Chuyển (Thêm mới) -- 2: Xoá bỏ cán bộ tập sự</param>
    ''' <remarks></remarks>
    Public Sub UpdateIdNew_HS_CanBo(ByVal _IdCanBoTS As String, ByVal _IdCB As String, ByVal _State As Byte)
        Dim strSQL As String = ""
        If (_State = 1) Then
            strSQL = "Update HS_CanBo Set IdNew = @_IdNew, Date_Update = @_Date_Update Where IdCanBo = @_IdCanBo"
        ElseIf (_State = 2) Then    'Trường Hợp Delete HSCB_TS thực hiện update lại cho bảng HS_CanBo
            strSQL = "Update HS_CanBo Set IdNew = '', Date_Update = @_Date_Update Where IdNew = @_IdNew"
        End If

        Dim command As SqlCommand = New SqlCommand(strSQL)
        command.CommandType = CommandType.Text

        command.Parameters.Add(New SqlParameter("@_IdNew", SqlDbType.VarChar))
        command.Parameters("@_IdNew").Value = _IdCanBoTS

        command.Parameters.Add(New SqlParameter("@_Date_Update", SqlDbType.DateTime))
        command.Parameters("@_Date_Update").Value = DateTime.Now

        If (_State = 1) Then
            command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
            command.Parameters("@_IdCanBo").Value = _IdCB
        End If
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật chỉ số xác định bản ghi của cán bộ tập sự vừa thêm vào hồ sơ cán bộ chính: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin  - Hồ sơ cán bộ tập sự
    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu vào - Hồ sơ cán bộ tập sự
    ''' </summary>
    ''' <param name="obj_hscb_ts">Hồ sơ cán bộ tập sự</param>
    ''' <returns>Id cán bộ tập sự vừa cập nhật</returns>
    ''' <remarks></remarks>
    Public Function Insert_HSCB_TS(ByVal obj_hscb_ts As HSCB_TS) As String
        Dim command As SqlCommand = New SqlCommand("HSCB_TS_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_Id", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_MaCB", obj_hscb_ts.MaCB))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_hscb_ts.HoTen))
        command.Parameters.Add(New SqlParameter("@_IdChiNhanh", obj_hscb_ts.IdChiNhanh))
        command.Parameters.Add(New SqlParameter("@_IdPhongBan", obj_hscb_ts.IdPhongBan))
        command.Parameters.Add(New SqlParameter("@_TenThuongGoi", obj_hscb_ts.TenThuongGoi))
        command.Parameters.Add(New SqlParameter("@_BiDanh", obj_hscb_ts.BiDanh))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_hscb_ts.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NgaySinh", obj_hscb_ts.NgaySinh))

        command.Parameters.Add(New SqlParameter("@_IdNS_Tinh", obj_hscb_ts.IdNS_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNS_Huyen", obj_hscb_ts.IdNS_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNS_Xa", obj_hscb_ts.IdNS_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNS_Thon", obj_hscb_ts.IdNS_Thon))
        command.Parameters.Add(New SqlParameter("@_NS_DiaChi", obj_hscb_ts.NS_DiaChi))

        command.Parameters.Add(New SqlParameter("@_IdNQ_Tinh", obj_hscb_ts.IdNQ_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Huyen", obj_hscb_ts.IdNQ_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Xa", obj_hscb_ts.IdNQ_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Thon", obj_hscb_ts.IdNQ_Thon))
        command.Parameters.Add(New SqlParameter("@_NQ_DiaChi", obj_hscb_ts.NQ_DiaChi))

        command.Parameters.Add(New SqlParameter("@_IdThT_Tinh", obj_hscb_ts.IdThT_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdThT_Huyen", obj_hscb_ts.IdThT_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdThT_Xa", obj_hscb_ts.IdThT_Xa))
        command.Parameters.Add(New SqlParameter("@_IdThT_Thon", obj_hscb_ts.IdThT_Thon))
        command.Parameters.Add(New SqlParameter("@_ThT_DiaChi", obj_hscb_ts.ThT_DiaChi))
        command.Parameters.Add(New SqlParameter("@_ThT_DienThoai", obj_hscb_ts.ThT_DienThoai))

        command.Parameters.Add(New SqlParameter("@_IdTTr_Tinh", obj_hscb_ts.IdTTr_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Huyen", obj_hscb_ts.IdTTr_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Xa", obj_hscb_ts.IdTTr_Xa))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Thon", obj_hscb_ts.IdTTr_Thon))
        command.Parameters.Add(New SqlParameter("@_TTr_DiaChi", obj_hscb_ts.TTr_DiaChi))
        command.Parameters.Add(New SqlParameter("@_TTr_DienThoai", obj_hscb_ts.TTr_DienThoai))

        command.Parameters.Add(New SqlParameter("@_IdQuocTich", obj_hscb_ts.IdQuocTich))
        command.Parameters.Add(New SqlParameter("@_IdDanToc", obj_hscb_ts.IdDanToc))
        command.Parameters.Add(New SqlParameter("@_IdTonGiao", obj_hscb_ts.IdTonGiao))
        command.Parameters.Add(New SqlParameter("@_IdThanhPhan_GD", obj_hscb_ts.IdThanhPhan_GD))
        command.Parameters.Add(New SqlParameter("@_NhomMau", obj_hscb_ts.NhomMau))

        command.Parameters.Add(New SqlParameter("@_CMT_So", obj_hscb_ts.CMT_So))
        command.Parameters.Add(New SqlParameter("@_CMT_NgayCap", obj_hscb_ts.CMT_NgayCap))
        command.Parameters.Add(New SqlParameter("@_CMT_NoiCap", obj_hscb_ts.CMT_NoiCap))

        command.Parameters.Add(New SqlParameter("@_IdUT_GD", obj_hscb_ts.IdUT_GD))
        command.Parameters.Add(New SqlParameter("@_IdUT_BT", obj_hscb_ts.IdUT_BT))
        command.Parameters.Add(New SqlParameter("@_Email", obj_hscb_ts.Email))
        command.Parameters.Add(New SqlParameter("@_DienThoai_NR", obj_hscb_ts.DienThoai_NR))
        command.Parameters.Add(New SqlParameter("@_DienThoai_DD", obj_hscb_ts.DienThoai_DD))
        command.Parameters.Add(New SqlParameter("@_DienThoai_CQ", obj_hscb_ts.DienThoai_CQ))
        command.Parameters.Add(New SqlParameter("@_SoFax", obj_hscb_ts.SoFax))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoVH", obj_hscb_ts.IdTrinhDoVH))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoCT", obj_hscb_ts.IdTrinhDoCT))

        command.Parameters.Add(New SqlParameter("@_SoTruong_CT", obj_hscb_ts.SoTruong_CT))
        command.Parameters.Add(New SqlParameter("@_CV_Lau", obj_hscb_ts.CV_Lau))

        If (CType(obj_hscb_ts.CM_Ngay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", obj_hscb_ts.CM_Ngay))
        End If
        command.Parameters.Add(New SqlParameter("@_CM_ToChuc", obj_hscb_ts.CM_ToChuc))

        'Thông tin liên quan đến Hồ sơ đảng
        If (CType(obj_hscb_ts.Dang_NgayVao, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Dang_NgayVao", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Dang_NgayVao", obj_hscb_ts.Dang_NgayVao))
        End If
        If (CType(obj_hscb_ts.Dang_NgayChTh, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Dang_NgayChTh", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Dang_NgayChTh", obj_hscb_ts.Dang_NgayChTh))
        End If

        If (CType(obj_hscb_ts.Dang_NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Dang_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Dang_NgayRa", obj_hscb_ts.Dang_NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_Dang_NoiKN", obj_hscb_ts.Dang_NoiKN))
        command.Parameters.Add(New SqlParameter("@_Dang_SoThe", obj_hscb_ts.Dang_SoThe))
        command.Parameters.Add(New SqlParameter("@_Dang_NGT", obj_hscb_ts.Dang_NGT))
        command.Parameters.Add(New SqlParameter("@_Dang_LyDoRa", obj_hscb_ts.Dang_LyDoRa))

        'Thông tin liên quan đến Tài khoản ngân hàng
        command.Parameters.Add(New SqlParameter("@_NH_MSKH", obj_hscb_ts.NH_MSKH))
        command.Parameters.Add(New SqlParameter("@_NH_SHTK", obj_hscb_ts.NH_SHTK))
        command.Parameters.Add(New SqlParameter("@_NH_Ten_NH", obj_hscb_ts.NH_Ten_NH))
        command.Parameters.Add(New SqlParameter("@_MaSoThue", obj_hscb_ts.MaSoThue))

        'Thông tin liên quan đến Bảo hiểm Xã hội
        command.Parameters.Add(New SqlParameter("@_BHXH_SoSo", obj_hscb_ts.BHXH_SoSo))
        If (CType(obj_hscb_ts.BHXH_NgayLam, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayLam", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayLam", obj_hscb_ts.BHXH_NgayLam))
        End If

        If (CType(obj_hscb_ts.BHXH_NgayDong, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayDong", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayDong", obj_hscb_ts.BHXH_NgayDong))
        End If
        command.Parameters.Add(New SqlParameter("@_BHXH_NoiLam", obj_hscb_ts.BHXH_NoiLam))

        command.Parameters.Add(New SqlParameter("@_IdHocHam", obj_hscb_ts.IdHocHam))
        command.Parameters.Add(New SqlParameter("@_IdHocVi", obj_hscb_ts.IdHocVi))
        command.Parameters.Add(New SqlParameter("@_IdTrdChMon", obj_hscb_ts.IdTrdChMon))
        command.Parameters.Add(New SqlParameter("@_IdTrdNgoaiNgu", obj_hscb_ts.IdTrdNgoaiNgu))
        command.Parameters.Add(New SqlParameter("@_IdTrdTinHoc", obj_hscb_ts.IdTrdTinHoc))
        command.Parameters.Add(New SqlParameter("@_IdChuyenNganhDT", obj_hscb_ts.IdChuyenNganhDT))
        'command.Parameters.Add(New SqlParameter("@_AnhThe", obj_hscb_ts.AnhThe))
        'If obj_hscb_ts.AnhThe Is Nothing Then
        '    command.Parameters.Add(New SqlParameter("@_AnhThe", System.Data.SqlDbType.Image))
        'Else
        '    command.Parameters.Add(New SqlParameter("@_AnhThe", obj_hscb_ts.AnhThe))
        'End If
        command.Parameters.Add(New SqlParameter("@_AnhThe", IIf(obj_hscb_ts.AnhThe Is Nothing, System.Data.SqlTypes.SqlBytes.Null, obj_hscb_ts.AnhThe)))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hscb_ts.GhiChu))
        command.Parameters.Add(New SqlParameter("@_idNew", obj_hscb_ts.idNew))
        command.Parameters.Add(New SqlParameter("@_TrangThai", obj_hscb_ts.TrangThai))
        command.Parameters.Add(New SqlParameter("@_HonNhan_Cd", obj_hscb_ts.HonNhan_Cd))
        Try
            _SqlHelper.executeSQL(command)
            obj_hscb_ts.Id = command.Parameters("@_Id").Value.ToString
            Return obj_hscb_ts.Id
        Catch ex As Exception
            MessageBox.Show("Cập nhật thêm mới hồ sơ cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Hàm cập nhật sửa đổi - Hồ sơ cán bộ tập sự
    ''' </summary>
    ''' <param name="obj_hscb_ts"></param>
    ''' <remarks></remarks>
    Public Sub Update_HSCB_TS(ByVal obj_hscb_ts As HSCB_TS)
        Dim command As SqlCommand = New SqlCommand("HSCB_TS_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_Id", obj_hscb_ts.Id))
        command.Parameters.Add(New SqlParameter("@_MaCB", obj_hscb_ts.MaCB))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_hscb_ts.HoTen))
        command.Parameters.Add(New SqlParameter("@_IdChiNhanh", obj_hscb_ts.IdChiNhanh))
        command.Parameters.Add(New SqlParameter("@_IdPhongBan", obj_hscb_ts.IdPhongBan))
        command.Parameters.Add(New SqlParameter("@_TenThuongGoi", obj_hscb_ts.TenThuongGoi))
        command.Parameters.Add(New SqlParameter("@_BiDanh", obj_hscb_ts.BiDanh))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_hscb_ts.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NgaySinh", obj_hscb_ts.NgaySinh))

        command.Parameters.Add(New SqlParameter("@_IdNS_Tinh", obj_hscb_ts.IdNS_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNS_Huyen", obj_hscb_ts.IdNS_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNS_Xa", obj_hscb_ts.IdNS_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNS_Thon", obj_hscb_ts.IdNS_Thon))
        command.Parameters.Add(New SqlParameter("@_NS_DiaChi", obj_hscb_ts.NS_DiaChi))

        command.Parameters.Add(New SqlParameter("@_IdNQ_Tinh", obj_hscb_ts.IdNQ_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Huyen", obj_hscb_ts.IdNQ_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Xa", obj_hscb_ts.IdNQ_Xa))
        command.Parameters.Add(New SqlParameter("@_IdNQ_Thon", obj_hscb_ts.IdNQ_Thon))
        command.Parameters.Add(New SqlParameter("@_NQ_DiaChi", obj_hscb_ts.NQ_DiaChi))

        command.Parameters.Add(New SqlParameter("@_IdThT_Tinh", obj_hscb_ts.IdThT_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdThT_Huyen", obj_hscb_ts.IdThT_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdThT_Xa", obj_hscb_ts.IdThT_Xa))
        command.Parameters.Add(New SqlParameter("@_IdThT_Thon", obj_hscb_ts.IdThT_Thon))
        command.Parameters.Add(New SqlParameter("@_ThT_DiaChi", obj_hscb_ts.ThT_DiaChi))
        command.Parameters.Add(New SqlParameter("@_ThT_DienThoai", obj_hscb_ts.ThT_DienThoai))

        command.Parameters.Add(New SqlParameter("@_IdTTr_Tinh", obj_hscb_ts.IdTTr_Tinh))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Huyen", obj_hscb_ts.IdTTr_Huyen))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Xa", obj_hscb_ts.IdTTr_Xa))
        command.Parameters.Add(New SqlParameter("@_IdTTr_Thon", obj_hscb_ts.IdTTr_Thon))
        command.Parameters.Add(New SqlParameter("@_TTr_DiaChi", obj_hscb_ts.TTr_DiaChi))
        command.Parameters.Add(New SqlParameter("@_TTr_DienThoai", obj_hscb_ts.TTr_DienThoai))

        command.Parameters.Add(New SqlParameter("@_IdQuocTich", obj_hscb_ts.IdQuocTich))
        command.Parameters.Add(New SqlParameter("@_IdDanToc", obj_hscb_ts.IdDanToc))
        command.Parameters.Add(New SqlParameter("@_IdTonGiao", obj_hscb_ts.IdTonGiao))
        command.Parameters.Add(New SqlParameter("@_IdThanhPhan_GD", obj_hscb_ts.IdThanhPhan_GD))
        command.Parameters.Add(New SqlParameter("@_NhomMau", obj_hscb_ts.NhomMau))
        command.Parameters.Add(New SqlParameter("@_CMT_So", obj_hscb_ts.CMT_So))
        command.Parameters.Add(New SqlParameter("@_CMT_NgayCap", obj_hscb_ts.CMT_NgayCap))
        command.Parameters.Add(New SqlParameter("@_CMT_NoiCap", obj_hscb_ts.CMT_NoiCap))
        command.Parameters.Add(New SqlParameter("@_IdUT_GD", obj_hscb_ts.IdUT_GD))
        command.Parameters.Add(New SqlParameter("@_IdUT_BT", obj_hscb_ts.IdUT_BT))
        command.Parameters.Add(New SqlParameter("@_Email", obj_hscb_ts.Email))
        command.Parameters.Add(New SqlParameter("@_DienThoai_NR", obj_hscb_ts.DienThoai_NR))
        command.Parameters.Add(New SqlParameter("@_DienThoai_DD", obj_hscb_ts.DienThoai_DD))
        command.Parameters.Add(New SqlParameter("@_DienThoai_CQ", obj_hscb_ts.DienThoai_CQ))
        command.Parameters.Add(New SqlParameter("@_SoFax", obj_hscb_ts.SoFax))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoVH", obj_hscb_ts.IdTrinhDoVH))
        command.Parameters.Add(New SqlParameter("@_IdTrinhDoCT", obj_hscb_ts.IdTrinhDoCT))
        command.Parameters.Add(New SqlParameter("@_SoTruong_CT", obj_hscb_ts.SoTruong_CT))
        command.Parameters.Add(New SqlParameter("@_CV_Lau", obj_hscb_ts.CV_Lau))

        If (CType(obj_hscb_ts.CM_Ngay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_CM_Ngay", obj_hscb_ts.CM_Ngay))
        End If
        command.Parameters.Add(New SqlParameter("@_CM_ToChuc", obj_hscb_ts.CM_ToChuc))

        'Thông tin liên quan đến Hồ sơ đảng
        If (CType(obj_hscb_ts.Dang_NgayVao, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Dang_NgayVao", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Dang_NgayVao", obj_hscb_ts.Dang_NgayVao))
        End If
        If (CType(obj_hscb_ts.Dang_NgayChTh, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Dang_NgayChTh", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Dang_NgayChTh", obj_hscb_ts.Dang_NgayChTh))
        End If
        If (CType(obj_hscb_ts.Dang_NgayRa, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_Dang_NgayRa", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_Dang_NgayRa", obj_hscb_ts.Dang_NgayRa))
        End If
        command.Parameters.Add(New SqlParameter("@_Dang_NoiKN", obj_hscb_ts.Dang_NoiKN))
        command.Parameters.Add(New SqlParameter("@_Dang_SoThe", obj_hscb_ts.Dang_SoThe))
        command.Parameters.Add(New SqlParameter("@_Dang_NGT", obj_hscb_ts.Dang_NGT))
        command.Parameters.Add(New SqlParameter("@_Dang_LyDoRa", obj_hscb_ts.Dang_LyDoRa))
        'Thông tin liên quan đến Tài khoản ngân hàng
        command.Parameters.Add(New SqlParameter("@_NH_MSKH", obj_hscb_ts.NH_MSKH))
        command.Parameters.Add(New SqlParameter("@_NH_SHTK", obj_hscb_ts.NH_SHTK))
        command.Parameters.Add(New SqlParameter("@_NH_Ten_NH", obj_hscb_ts.NH_Ten_NH))
        command.Parameters.Add(New SqlParameter("@_MaSoThue", obj_hscb_ts.MaSoThue))
        'Thông tin liên quan đến Bảo hiểm Xã hội
        command.Parameters.Add(New SqlParameter("@_BHXH_SoSo", obj_hscb_ts.BHXH_SoSo))
        If (CType(obj_hscb_ts.BHXH_NgayLam, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayLam", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayLam", obj_hscb_ts.BHXH_NgayLam))
        End If

        If (CType(obj_hscb_ts.BHXH_NgayDong, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayDong", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_BHXH_NgayDong", obj_hscb_ts.BHXH_NgayDong))
        End If
        command.Parameters.Add(New SqlParameter("@_BHXH_NoiLam", obj_hscb_ts.BHXH_NoiLam))
        command.Parameters.Add(New SqlParameter("@_IdHocHam", obj_hscb_ts.IdHocHam))
        command.Parameters.Add(New SqlParameter("@_IdHocVi", obj_hscb_ts.IdHocVi))
        command.Parameters.Add(New SqlParameter("@_IdTrdChMon", obj_hscb_ts.IdTrdChMon))
        command.Parameters.Add(New SqlParameter("@_IdTrdNgoaiNgu", obj_hscb_ts.IdTrdNgoaiNgu))
        command.Parameters.Add(New SqlParameter("@_IdTrdTinHoc", obj_hscb_ts.IdTrdTinHoc))
        command.Parameters.Add(New SqlParameter("@_IdChuyenNganhDT", obj_hscb_ts.IdChuyenNganhDT))
        command.Parameters.Add(New SqlParameter("@_AnhThe", IIf(obj_hscb_ts.AnhThe Is Nothing, System.Data.SqlTypes.SqlBytes.Null, obj_hscb_ts.AnhThe)))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hscb_ts.GhiChu))
        command.Parameters.Add(New SqlParameter("@_idNew", obj_hscb_ts.idNew))
        command.Parameters.Add(New SqlParameter("@_TrangThai", obj_hscb_ts.TrangThai))
        command.Parameters.Add(New SqlParameter("@_HonNhan_Cd", obj_hscb_ts.HonNhan_Cd))
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật sửa đổi hồ sơ cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện - Xoá bỏ hồ sơ cán bộ tập sự
    ''' </summary>
    ''' <param name="_Id">Id cán bộ tập sự cần xoá</param>
    ''' <remarks></remarks>
    Public Sub Delete_HSCB_TS(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HSCB_TS_Delete")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_Id", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_Id").Value = _Id
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hồ sơ cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật các thông tin liên quan đến cán bộ tập sự khi chuyển HSCB Tập sự sang Chính thức
    ''' </summary>
    ''' <param name="_IdCB_TS">Id cán bộ tập sự được chuyển</param>
    ''' <param name="_IdCB_CT">Id cán bộ chính thức được sinh khi chuyển</param>
    ''' <param name="_TrangThai">Trạng thái bản ghi: 0: Chưa chuyển sang hồ sơ chính. 1: Đã chuyển sang hồ sơ chính thức</param>
    ''' <param name="_Flag">Chỉ số xác định. Nếu 0: Cập nhật khi chuyển. Ngược lại cập nhật khi xóa CB Chính thức</param>
    ''' <remarks></remarks>
    Public Sub Update_HSCB_TS_WhenMove(ByVal _IdCB_TS As String, ByVal _IdCB_CT As String, ByVal _TrangThai As Byte, ByVal _Flag As Byte)
        Dim command As SqlCommand = New SqlCommand("HSCB_TS_Update_WhenMove")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_Flag", SqlDbType.VarChar))
        command.Parameters("@_Flag").Value = _Flag
        command.Parameters.Add(New SqlParameter("@_IdTS", SqlDbType.VarChar))
        command.Parameters("@_IdTS").Value = _IdCB_TS
        command.Parameters.Add(New SqlParameter("@_IdCB", SqlDbType.VarChar))
        command.Parameters("@_IdCB").Value = _IdCB_CT
        command.Parameters.Add(New SqlParameter("@_TrangThai", SqlDbType.Bit))
        command.Parameters("@_TrangThai").Value = _TrangThai
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật hồ sơ cán bộ tập sự khi chuyển từ lao động tập sự sang lao động chính thức: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện update lại trường Đến ngày của Hợp đồng lao động TS khi cán bộ này chuyển sang chính thức
    ''' Mà @_DenNgay >= @_NgayHL And @_NgayHL >= @_TuNgay
    ''' </summary>
    ''' <param name="_IdCBTS">Id cán bộ tập sự được chuyển</param>
    ''' <param name="_NgayHL">Ngày hiệu lực của quyết định chính thức</param>
    ''' <remarks></remarks>
    Public Sub Update_HSCB_TS_HDLD_WhenMove(ByVal _IdCBTS As String, ByVal _NgayHL As DateTime)
        Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_Update_WhenMove")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCBTS", SqlDbType.VarChar))
        command.Parameters("@_IdCBTS").Value = _IdCBTS
        command.Parameters.Add(New SqlParameter("@_NgayHL", SqlDbType.DateTime))
        command.Parameters("@_NgayHL").Value = _NgayHL

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật hồ sơ cán bộ tập sự khi chuyển từ lao động tập sự sang lao động chính thức: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin - Hợp đồng lao động của cán bộ tập sự
    Public Function Insert_Update_HSCB_TS_HDLD(ByVal obj_hdld_ts As HSCB_TS_HDLD) As String
        Dim _Ret As String = ""
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "HSCB_TS_HDLD_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.VarChar, 16).Direction = ParameterDirection.Output
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdBranch", IdDONVI, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCBTS_HDLD", obj_hdld_ts.IdCBTS_HDLD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo", obj_hdld_ts.IdCanBo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoQD", obj_hdld_ts.SoQD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgayQD", obj_hdld_ts.NgayQD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NguoiQD", obj_hdld_ts.NguoiQD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCV_Nguoi_QD", obj_hdld_ts.IdCV_Nguoi_QD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tungay", obj_hdld_ts.TuNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DenNgay", obj_hdld_ts.DenNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TuGio", obj_hdld_ts.TuGio, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DenGio", obj_hdld_ts.DenGio, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdHT_TraLuong", obj_hdld_ts.IdHT_TraLuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_HeSo", obj_hdld_ts.HeSo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdBacLuong", obj_hdld_ts.IdBacLuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TyleHuong", obj_hdld_ts.TyleHuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Tien_Luong", obj_hdld_ts.Tien_Luong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdChuyenMon", obj_hdld_ts.IdChuyenMon, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CongViec", obj_hdld_ts.CongViec, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHXH", obj_hdld_ts.BHXH, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_BHYT", obj_hdld_ts.BHYT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Loai", obj_hdld_ts.Loai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_Status", obj_hdld_ts.Status, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_Ngay_HL", obj_hdld_ts.TV_Ngay_HL, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_So_QD", obj_hdld_ts.TV_So_QD, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_NgayKy_QD", obj_hdld_ts.TV_NgayKy_QD, ParameterDirection.Input))

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_IdLyDo", obj_hdld_ts.TV_IdLyDo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_TroCap_ThoiViec", obj_hdld_ts.TV_TroCap_ThoiViec, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_TroCap_Khac", obj_hdld_ts.TV_TroCap_Khac, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_SoTien_BoiThuong", obj_hdld_ts.TV_SoTien_BoiThuong, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TV_SoTien_ThuHoi", obj_hdld_ts.TV_SoTien_ThuHoi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GhiChu", obj_hdld_ts.GhiChu, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ChiNhanhId", obj_hdld_ts.ChiNhanhId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PhongBanDonViId", obj_hdld_ts.PhongBanDonViId, ParameterDirection.Input))
                  
                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = _command.Parameters("@_IdOutPut").Value.ToString()
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi Cập nhật thông tin Hợp đồng Lao động Ngắn hạn/Tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Globals.Logger.Error("Lỗi Cập nhật Hợp đồng Lao động Ngắn hạn/Tập sự Insert_Update_HSCB_TS_HDLD(): " + ex.Message)
                    _Ret = ""
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function


    ' ''' <summary>
    ' ''' Hàm cập nhật Thêm mới - Hợp đồng lao động của cán bộ tập sự
    ' ''' </summary>
    ' ''' <param name="obj_hdld_ts">Hợp đồng lao động của cán bộ tập sự</param>
    ' ''' <returns>Id hợp đồng lao động của cán bộ tập sự</returns>
    ' ''' <remarks></remarks>
    'Public Function Insert_HSCB_TS_HDLD(ByVal obj_hdld_ts As HSCB_TS_HDLD) As String
    '    Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_Insert")
    '    command.CommandType = CommandType.StoredProcedure

    '    command.Parameters.Add("@_IdCBTS_HDLD", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
    '    command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
    '    command.Parameters.Add(New SqlParameter("@_IdCanbo", obj_hdld_ts.IdCanbo))
    '    command.Parameters.Add(New SqlParameter("@_SoQD", obj_hdld_ts.SoQD))
    '    command.Parameters.Add(New SqlParameter("@_NgayQD", obj_hdld_ts.NgayQD))
    '    command.Parameters.Add(New SqlParameter("@_NguoiQD", obj_hdld_ts.NguoiQD))
    '    command.Parameters.Add(New SqlParameter("@_IdCV_Nguoi_QD", obj_hdld_ts.IdCV_Nguoi_QD))

    '    If (CType(obj_hdld_ts.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_Tungay", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_Tungay", obj_hdld_ts.TuNgay))
    '    End If
    '    If (CType(obj_hdld_ts.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_hdld_ts.DenNgay))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_TuGio", obj_hdld_ts.TuGio))
    '    command.Parameters.Add(New SqlParameter("@_DenGio", obj_hdld_ts.DenGio))
    '    command.Parameters.Add(New SqlParameter("@_IdHT_TraLuong", obj_hdld_ts.IdHT_TraLuong))
    '    command.Parameters.Add(New SqlParameter("@_HeSo", obj_hdld_ts.HeSo))
    '    command.Parameters.Add(New SqlParameter("@_IdBacLuong", obj_hdld_ts.IdBacLuong))
    '    command.Parameters.Add(New SqlParameter("@_TyleHuong", obj_hdld_ts.TyleHuong))
    '    command.Parameters.Add(New SqlParameter("@_Tien_Luong", obj_hdld_ts.Tien_Luong))
    '    command.Parameters.Add(New SqlParameter("@_IdChuyenMon", obj_hdld_ts.IdChuyenMon))
    '    command.Parameters.Add(New SqlParameter("@_CongViec", obj_hdld_ts.CongViec))
    '    command.Parameters.Add(New SqlParameter("@_BHXH", obj_hdld_ts.BHXH))
    '    command.Parameters.Add(New SqlParameter("@_BHYT", obj_hdld_ts.BHYT))
    '    command.Parameters.Add(New SqlParameter("@_Loai", obj_hdld_ts.Loai))
    '    command.Parameters.Add(New SqlParameter("@_Status", obj_hdld_ts.Status))
    '    If (CType(obj_hdld_ts.TV_Ngay_HL, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_TV_Ngay_HL", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_TV_Ngay_HL", obj_hdld_ts.TV_Ngay_HL))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_TV_So_QD", obj_hdld_ts.TV_So_QD))
    '    If (CType(obj_hdld_ts.TV_NgayKy_QD, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_TV_NgayKy_QD", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_TV_NgayKy_QD", obj_hdld_ts.TV_NgayKy_QD))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_TV_IdLyDo", obj_hdld_ts.TV_IdLyDo))
    '    command.Parameters.Add(New SqlParameter("@_TV_TroCap_ThoiViec", obj_hdld_ts.TV_TroCap_ThoiViec))
    '    command.Parameters.Add(New SqlParameter("@_TV_TroCap_Khac", obj_hdld_ts.TV_TroCap_Khac))
    '    command.Parameters.Add(New SqlParameter("@_TV_SoTien_BoiThuong", obj_hdld_ts.TV_SoTien_BoiThuong))
    '    command.Parameters.Add(New SqlParameter("@_TV_SoTien_ThuHoi", obj_hdld_ts.TV_SoTien_ThuHoi))
    '    command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hdld_ts.GhiChu))
    '    Try
    '        _SqlHelper.executeSQL(command)
    '        obj_hdld_ts.IdCBTS_HDLD = command.Parameters("@_IdCBTS_HDLD").Value.ToString()
    '        Return obj_hdld_ts.IdCBTS_HDLD
    '    Catch ex As Exception
    '        MessageBox.Show("Cập nhật thêm mới hợp đồng lao động của cán bộ tập sư: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '        Return Nothing
    '    End Try
    'End Function

    ' ''' <summary>
    ' ''' Hàm cập nhật Sửa đổi - Hợp đồng lao động của cán bộ tập sự
    ' ''' </summary>
    ' ''' <param name="obj_hdld_ts">Hợp đồng lao động của cán bộ tập sự</param>
    ' ''' <remarks></remarks>
    'Public Sub Update_HSCB_TS_HDLD(ByVal obj_hdld_ts As HSCB_TS_HDLD)
    '    Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_Update")
    '    command.CommandType = CommandType.StoredProcedure
    '    command.Parameters.Add(New SqlParameter("@_IdCBTS_HDLD", obj_hdld_ts.IdCBTS_HDLD))
    '    command.Parameters.Add(New SqlParameter("@_IdCanbo", obj_hdld_ts.IdCanbo))
    '    command.Parameters.Add(New SqlParameter("@_SoQD", obj_hdld_ts.SoQD))

    '    command.Parameters.Add(New SqlParameter("@_NgayQD", obj_hdld_ts.NgayQD))
    '    command.Parameters.Add(New SqlParameter("@_NguoiQD", obj_hdld_ts.NguoiQD))

    '    command.Parameters.Add(New SqlParameter("@_IdCV_Nguoi_QD", obj_hdld_ts.IdCV_Nguoi_QD))

    '    If (CType(obj_hdld_ts.TuNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_Tungay", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_Tungay", obj_hdld_ts.TuNgay))
    '    End If
    '    If (CType(obj_hdld_ts.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_hdld_ts.DenNgay))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_TuGio", obj_hdld_ts.TuGio))
    '    command.Parameters.Add(New SqlParameter("@_DenGio", obj_hdld_ts.DenGio))
    '    command.Parameters.Add(New SqlParameter("@_IdHT_TraLuong", obj_hdld_ts.IdHT_TraLuong))
    '    command.Parameters.Add(New SqlParameter("@_HeSo", obj_hdld_ts.HeSo))
    '    command.Parameters.Add(New SqlParameter("@_IdBacLuong", obj_hdld_ts.IdBacLuong))
    '    command.Parameters.Add(New SqlParameter("@_TyleHuong", obj_hdld_ts.TyleHuong))
    '    command.Parameters.Add(New SqlParameter("@_Tien_Luong", obj_hdld_ts.Tien_Luong))
    '    command.Parameters.Add(New SqlParameter("@_IdChuyenMon", obj_hdld_ts.IdChuyenMon))
    '    command.Parameters.Add(New SqlParameter("@_CongViec", obj_hdld_ts.CongViec))
    '    command.Parameters.Add(New SqlParameter("@_BHXH", obj_hdld_ts.BHXH))
    '    command.Parameters.Add(New SqlParameter("@_BHYT", obj_hdld_ts.BHYT))
    '    command.Parameters.Add(New SqlParameter("@_Loai", obj_hdld_ts.Loai))
    '    command.Parameters.Add(New SqlParameter("@_Status", obj_hdld_ts.Status))
    '    If (CType(obj_hdld_ts.TV_Ngay_HL, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_TV_Ngay_HL", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_TV_Ngay_HL", obj_hdld_ts.TV_Ngay_HL))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_TV_So_QD", obj_hdld_ts.TV_So_QD))
    '    If (CType(obj_hdld_ts.TV_NgayKy_QD, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_TV_NgayKy_QD", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_TV_NgayKy_QD", obj_hdld_ts.TV_NgayKy_QD))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_TV_IdLyDo", obj_hdld_ts.TV_IdLyDo))
    '    command.Parameters.Add(New SqlParameter("@_TV_TroCap_ThoiViec", obj_hdld_ts.TV_TroCap_ThoiViec))
    '    command.Parameters.Add(New SqlParameter("@_TV_TroCap_Khac", obj_hdld_ts.TV_TroCap_Khac))
    '    command.Parameters.Add(New SqlParameter("@_TV_SoTien_BoiThuong", obj_hdld_ts.TV_SoTien_BoiThuong))
    '    command.Parameters.Add(New SqlParameter("@_TV_SoTien_ThuHoi", obj_hdld_ts.TV_SoTien_ThuHoi))
    '    command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hdld_ts.GhiChu))
    '    Try
    '        _SqlHelper.executeSQL(command)
    '    Catch ex As Exception
    '        MessageBox.Show("Cập nhật sửa đổi hợp đồng lao động của cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub

    ''' <summary>
    ''' Hàm thực hiện Xoá bỏ dữ liệu - Hợp đồng lao động của cán bộ Tập sự
    ''' </summary>
    ''' <param name="_Id">Id Hợp đồng lao động của cán bộ Tập sự</param>
    ''' <remarks></remarks>
    Public Sub Delete_HSCB_TS_HDLD(ByVal _Id As String)
        Dim command As SqlCommand = New SqlCommand("HSCB_TS_HDLD_Delete")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add(New SqlParameter("@_IdCBTS_HDLD", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdCBTS_HDLD").Value = _Id
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ hợp đồng lao động của cán bộ tập sự: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin  - Quyết định Nhân sự
    ''' <summary>
    ''' Hàm thực hiện thêm mới dữ liệu - Quyết định nhân sự
    ''' </summary>
    ''' <param name="obj_quyetdinhns">Đối tượng - Quyết định nhân sự</param>
    ''' <remarks></remarks>
    Public Sub Insert_Decision(ByVal obj_quyetdinhns As QD_NhanSu)
        Dim command As SqlCommand = New SqlCommand("QDNhanSu_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdQDNhanSu", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_So_QD", obj_quyetdinhns.So_QD))
        command.Parameters.Add(New SqlParameter("@_NgayKy_QD", obj_quyetdinhns.NgayKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_quyetdinhns.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_IdLoaiQD", obj_quyetdinhns.IdLoaiQD))
        command.Parameters.Add(New SqlParameter("@_NgayHL", obj_quyetdinhns.NgayHL))
        'Ngày bổ nhiệm tiếp theo - Trong quyết định nhân sự
        If (CType(obj_quyetdinhns.NgayBoNhiem_TT, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", obj_quyetdinhns.NgayBoNhiem_TT))
        End If
        'Ngày thôi lương - trong quyết định nhân sự nếu có
        If (CType(obj_quyetdinhns.NgayThoiLuong, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_NgayThoiLuong", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_NgayThoiLuong", obj_quyetdinhns.NgayThoiLuong))
        End If

        command.Parameters.Add(New SqlParameter("@_NguoiKy_QD", obj_quyetdinhns.NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_Nguoiky_QD", obj_quyetdinhns.IdCV_Nguoiky_QD))

        command.Parameters.Add(New SqlParameter("@_IdDonVi_Cu", obj_quyetdinhns.IdDonVi_Cu))
        command.Parameters.Add(New SqlParameter("@_IdPhong_Cu", obj_quyetdinhns.IdPhong_Cu))
        command.Parameters.Add(New SqlParameter("@_IdChucVu_Cu", obj_quyetdinhns.IdChucVu_Cu))
        command.Parameters.Add(New SqlParameter("@_IdChuyenMon_Cu", obj_quyetdinhns.IdChuyenMon_Cu))

        command.Parameters.Add(New SqlParameter("@_IdDonVi_Moi", obj_quyetdinhns.IdDonVi_Moi))
        command.Parameters.Add(New SqlParameter("@_IdPhong_Moi", obj_quyetdinhns.IdPhong_Moi))
        command.Parameters.Add(New SqlParameter("@_IdChucVu_Moi", obj_quyetdinhns.IdChucVu_Moi))
        command.Parameters.Add(New SqlParameter("@_IdChuyenMon_Moi", obj_quyetdinhns.IdChuyenMon_Moi))

        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_quyetdinhns.GhiChu))
        command.Parameters.Add(New SqlParameter("@_Active", obj_quyetdinhns.Active))
        command.Parameters.Add(New SqlParameter("@_IsKiemNhiem", obj_quyetdinhns.IsKiemNhiem))
        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới quyết định nhân sự", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Quyết định nhân sự
    ''' </summary>
    ''' <param name="obj_quyetdinhns">Đối tượng - Quyết định nhân sự</param>
    ''' <remarks></remarks>
    Public Sub Update_Decision(ByVal obj_quyetdinhns As QD_NhanSu)
        Dim command As SqlCommand = New SqlCommand("QDNhanSu_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdQDNhanSu", obj_quyetdinhns.IdQDNhanSu))
        command.Parameters.Add(New SqlParameter("@_So_QD", obj_quyetdinhns.So_QD))
        command.Parameters.Add(New SqlParameter("@_NgayKy_QD", obj_quyetdinhns.NgayKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_quyetdinhns.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_IdLoaiQD", obj_quyetdinhns.IdLoaiQD))
        command.Parameters.Add(New SqlParameter("@_NgayHL", obj_quyetdinhns.NgayHL))
        command.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", obj_quyetdinhns.NgayBoNhiem_TT))
        command.Parameters.Add(New SqlParameter("@_NgayThoiLuong", obj_quyetdinhns.NgayThoiLuong))

        command.Parameters.Add(New SqlParameter("@_NguoiKy_QD", obj_quyetdinhns.NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_Nguoiky_QD", obj_quyetdinhns.IdCV_Nguoiky_QD))

        command.Parameters.Add(New SqlParameter("@_IdDonVi_Cu", obj_quyetdinhns.IdDonVi_Cu))
        command.Parameters.Add(New SqlParameter("@_IdPhong_Cu", obj_quyetdinhns.IdPhong_Cu))
        command.Parameters.Add(New SqlParameter("@_IdChucVu_Cu", obj_quyetdinhns.IdChucVu_Cu))
        command.Parameters.Add(New SqlParameter("@_IdChuyenMon_Cu", obj_quyetdinhns.IdChuyenMon_Cu))

        command.Parameters.Add(New SqlParameter("@_IdDonVi_Moi", obj_quyetdinhns.IdDonVi_Moi))
        command.Parameters.Add(New SqlParameter("@_IdPhong_Moi", obj_quyetdinhns.IdPhong_Moi))
        command.Parameters.Add(New SqlParameter("@_IdChucVu_Moi", obj_quyetdinhns.IdChucVu_Moi))
        command.Parameters.Add(New SqlParameter("@_IdChuyenMon_Moi", obj_quyetdinhns.IdChuyenMon_Moi))

        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_quyetdinhns.GhiChu))
        command.Parameters.Add(New SqlParameter("@_Active", obj_quyetdinhns.Active))
        command.Parameters.Add(New SqlParameter("@_IsKiemNhiem", obj_quyetdinhns.IsKiemNhiem))
        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật quyết định nhân sự", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Quyết định nhân sự
    ''' </summary>
    ''' <param name="_IdQDNhanSu">Id quyết định nhân sự</param>
    ''' <remarks></remarks>
    Public Sub Delete_Decision(ByVal _IdQDNhanSu As String)
        Dim command As SqlCommand = New SqlCommand("QDNhanSu_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdQDNhanSu", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdQDNhanSu").Value = _IdQDNhanSu
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá bỏ quyết định nhân sự", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin - Hồ sơ Công tác của cán bộ <Employment History>

    ''' <summary>
    ''' Hàm cập nhật thông tin  - Hồ sơ Công tác của cán bộ Employment History
    ''' </summary>
    ''' <param name="obj_emp">Hồ sơ Công tác</param>
    ''' <returns>Chuỗi chỉ số xác định bản ghi HS công tác (Tự sinh)</returns>
    ''' <remarks></remarks>
    Public Function Insert_Employment_History(ByVal obj_emp As HS_CongTac) As String
        Dim command As SqlCommand = New SqlCommand("HS_CongTac_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdHSCongTac", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_emp.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_emp.TuNgay))
        If (CType(obj_emp.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_emp.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_ChucVu", obj_emp.ChucVu))
        command.Parameters.Add(New SqlParameter("@_DiaChi", obj_emp.DiaChi))
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_emp.LyDo))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_emp.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_emp.IdHSCongTac = command.Parameters("@_IdHSCongTac").Value.ToString
            Return obj_emp.IdHSCongTac
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ công tác", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện cập nhật dữ liệu - Hồ sơ công tác
    ''' </summary>
    ''' <param name="obj_emp"></param>
    ''' <remarks></remarks>
    Public Sub Update_Employment_History(ByVal obj_emp As HS_CongTac)
        Dim command As SqlCommand = New SqlCommand("HS_CongTac_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSCongTac", obj_emp.IdHSCongTac))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_emp.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_emp.TuNgay))
        If (CType(obj_emp.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
            command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
        Else
            command.Parameters.Add(New SqlParameter("@_DenNgay", obj_emp.DenNgay))
        End If
        command.Parameters.Add(New SqlParameter("@_ChucVu", obj_emp.ChucVu))
        command.Parameters.Add(New SqlParameter("@_DiaChi", obj_emp.DiaChi))
        command.Parameters.Add(New SqlParameter("@_LyDo", obj_emp.LyDo))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_emp.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ công tác", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ công tác
    ''' </summary>
    ''' <param name="_IdHSCongTac">Id hồ sơ công tác</param>
    ''' <remarks></remarks>
    Public Sub Delete_Employment_History(ByVal _IdHSCongTac As String)
        Dim command As SqlCommand = New SqlCommand("HS_CongTac_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSCongTac", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdHSCongTac").Value = _IdHSCongTac
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Hồ sơ công tác", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin - Hồ sơ xuất ngoại của cán bộ <Abroad>

    ''' <summary>
    ''' Hàm thêm mới dữ liệu - Hồ sơ xuất ngoại của cán bộ
    ''' </summary>
    ''' <param name="obj_hsxn">Hồ sơ xuất ngoại</param>
    ''' <returns>Chuỗi chỉ số xác định bản ghi - tự sinh</returns>
    ''' <remarks></remarks>
    Public Function Insert_Abroad(ByVal obj_hsxn As HS_XuatNgoai) As String
        Dim command As SqlCommand = New SqlCommand("HS_XuatNgoai_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdXuatNgoai", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hsxn.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_hsxn.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_hsxn.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdNuocDen", obj_hsxn.IdNuocDen))
        command.Parameters.Add(New SqlParameter("@_MucDich", obj_hsxn.MucDich))
        command.Parameters.Add(New SqlParameter("@_SoQD", obj_hsxn.SoQD))
        command.Parameters.Add(New SqlParameter("@_NgayKy_QD", obj_hsxn.NgayKy_QD))
        command.Parameters.Add(New SqlParameter("@_NguoiKy_QD", obj_hsxn.NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_NguoiKy_QD", obj_hsxn.IdCV_NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hsxn.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_hsxn.IdXuatNgoai = command.Parameters("@_IdXuatNgoai").Value.ToString
            Return obj_hsxn.IdXuatNgoai
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ xuất ngoại", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm Cập nhật dữ liệu - Hồ sơ xuất ngoại của cán bộ
    ''' </summary>
    ''' <param name="obj_hsxn"></param>
    ''' <remarks></remarks>
    Public Sub Update_Abroad(ByVal obj_hsxn As HS_XuatNgoai)
        Dim command As SqlCommand = New SqlCommand("HS_XuatNgoai_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdXuatNgoai", obj_hsxn.IdXuatNgoai))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hsxn.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_hsxn.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_hsxn.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdNuocDen", obj_hsxn.IdNuocDen))
        command.Parameters.Add(New SqlParameter("@_MucDich", obj_hsxn.MucDich))
        command.Parameters.Add(New SqlParameter("@_SoQD", obj_hsxn.SoQD))
        command.Parameters.Add(New SqlParameter("@_NgayKy_QD", obj_hsxn.NgayKy_QD))
        command.Parameters.Add(New SqlParameter("@_NguoiKy_QD", obj_hsxn.NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_IdCV_NguoiKy_QD", obj_hsxn.IdCV_NguoiKy_QD))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hsxn.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ xuất ngoại", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ xuất ngoại
    ''' </summary>
    ''' <param name="_IdXuatNgoai">Id hồ sơ công tác</param>
    ''' <remarks></remarks>
    Public Sub Delete_Abroad(ByVal _IdXuatNgoai As String)
        Dim command As SqlCommand = New SqlCommand("HS_XuatNgoai_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdXuatNgoai", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdXuatNgoai").Value = _IdXuatNgoai
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Hồ sơ xuất ngoại", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin - Hồ sơ hộ chiếu của cán bộ <Passport>

    ''' <summary>
    ''' Hàm thêm mới dữ liệu - Hồ sơ hộ chiếu của cán bộ Passport History
    ''' </summary>
    ''' <param name="obj_passport">Hồ sơ hộ chiếu</param>
    ''' <returns>Chuỗi chỉ số tự sinh xác định tính duy nhất của bản ghi</returns>
    ''' <remarks></remarks>
    Public Function Insert_Passport(ByVal obj_passport As HS_HoChieu) As String
        Dim command As SqlCommand = New SqlCommand("HS_HoChieu_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdHoChieu", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_passport.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_So_HoChieu", obj_passport.So_HoChieu))
        command.Parameters.Add(New SqlParameter("@_Loai_HC", obj_passport.Loai_HC))
        command.Parameters.Add(New SqlParameter("@_NgayCap", obj_passport.NgayCap))
        command.Parameters.Add(New SqlParameter("@_NoiCap", obj_passport.NoiCap))
        command.Parameters.Add(New SqlParameter("@_NgayHH", obj_passport.NgayHH))
        command.Parameters.Add(New SqlParameter("@_TinhTrang", obj_passport.TinhTrang))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_passport.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_passport.IdHoChieu = command.Parameters("@_IdHoChieu").Value.ToString
            Return obj_passport.IdHoChieu
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ hộ chiếu", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm cập nhật dữ liệu - Hồ sơ hộ chiếu
    ''' </summary>
    ''' <param name="obj_passport"></param>
    ''' <remarks></remarks>
    Public Sub Update_Passport(ByVal obj_passport As HS_HoChieu)
        Dim command As SqlCommand = New SqlCommand("HS_HoChieu_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHoChieu", obj_passport.IdHoChieu))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_passport.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_So_HoChieu", obj_passport.So_HoChieu))
        command.Parameters.Add(New SqlParameter("@_Loai_HC", obj_passport.Loai_HC))
        command.Parameters.Add(New SqlParameter("@_NgayCap", obj_passport.NgayCap))
        command.Parameters.Add(New SqlParameter("@_NoiCap", obj_passport.NoiCap))
        command.Parameters.Add(New SqlParameter("@_NgayHH", obj_passport.NgayHH))
        command.Parameters.Add(New SqlParameter("@_TinhTrang", obj_passport.TinhTrang))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_passport.GhiChu))
        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ hộ chiếu", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ Hộ chiếu
    ''' </summary>
    ''' <param name="_IdHoChieu">Id hồ sơ hộ chiếu</param>
    ''' <remarks></remarks>
    Public Sub Delete_Passport(ByVal _IdHoChieu As String)
        Dim command As SqlCommand = New SqlCommand("HS_HoChieu_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHoChieu", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdHoChieu").Value = _IdHoChieu
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Hồ sơ hộ chiếu", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin - Hồ sơ Lực lượng vũ trang <Armed Force>
    ''' <summary>
    ''' Hàm thêm mới dữ liệu - Hồ sơ lực lượng vũ trang
    ''' </summary>
    ''' <param name="obj_llvt">Hồ sơ lực lượng vũ trang</param>
    ''' <returns>Chuỗi chỉ số tự sinh xác định tính duy nhất của bản ghi</returns>
    ''' <remarks></remarks>
    Public Function Insert_ArmedForce(ByVal obj_llvt As HS_LLVT) As String
        Dim command As SqlCommand = New SqlCommand("HS_LLVT_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdHSLLVT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_llvt.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_llvt.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_llvt.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdLoaiLLVT", obj_llvt.IdLoaiLLVT))
        command.Parameters.Add(New SqlParameter("@_IdQuanHam", obj_llvt.IdQuanHam))
        command.Parameters.Add(New SqlParameter("@_ChucVu", obj_llvt.ChucVu))
        command.Parameters.Add(New SqlParameter("@_DonVi", obj_llvt.DonVi))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_llvt.GhiChu))
        Try
            _SqlHelper.executeSQL(command)
            obj_llvt.IdHSLLVT = command.Parameters("@_IdHSLLVT").Value.ToString
            Return obj_llvt.IdHSLLVT
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ lực lượng vũ trang", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm cập nhật dữ liệu - Hồ sơ lực lượng vũ trang
    ''' </summary>
    ''' <param name="obj_llvt"></param>
    ''' <remarks></remarks>
    Public Sub Update_ArmedForce(ByVal obj_llvt As HS_LLVT)
        Dim command As SqlCommand = New SqlCommand("HS_LLVT_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSLLVT", obj_llvt.IdHSLLVT))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_llvt.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_llvt.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_llvt.DenNgay))
        command.Parameters.Add(New SqlParameter("@_IdLoaiLLVT", obj_llvt.IdLoaiLLVT))
        command.Parameters.Add(New SqlParameter("@_IdQuanHam", obj_llvt.IdQuanHam))
        command.Parameters.Add(New SqlParameter("@_ChucVu", obj_llvt.ChucVu))
        command.Parameters.Add(New SqlParameter("@_DonVi", obj_llvt.DonVi))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_llvt.GhiChu))

        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ lực lượng vũ trang", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try

    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Lực lượng vũ trang
    ''' </summary>
    ''' <param name="_IdHSLLVT">Id hồ sơ Lực lượng vũ trang</param>
    ''' <remarks></remarks>
    Public Sub Delete_ArmedForce(ByVal _IdHSLLVT As String)
        Dim command As SqlCommand = New SqlCommand("HS_LLVT_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSLLVT", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdHSLLVT").Value = _IdHSLLVT
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Lực lượng vũ trang", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin - Hồ sơ cũ <Old document>
    ''' <summary>
    ''' Hàm Thêm mới dữ liệu - Hồ sơ cũ của cán bộ
    ''' </summary>
    ''' <param name="obj_hsOld">Hồ sơ cũ của cán bộ</param>
    ''' <returns>Chuỗi chỉ số xác định tính duy nhất của bản ghi</returns>
    ''' <remarks></remarks>
    Public Function Insert_OldDocumnet(ByVal obj_hsOld As HS_Cu) As String
        Dim command As SqlCommand = New SqlCommand("HS_Cu_Insert")
        command.CommandType = CommandType.StoredProcedure
        command.Parameters.Add("@_IdHSCu", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hsOld.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuThang", obj_hsOld.TuThang))
        command.Parameters.Add(New SqlParameter("@_DenThang", obj_hsOld.DenThang))
        command.Parameters.Add(New SqlParameter("@_NgheNghiep", obj_hsOld.NgheNghiep))
        command.Parameters.Add(New SqlParameter("@_DiaChi", obj_hsOld.DiaChi))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hsOld.GhiChu))
        Try
            _SqlHelper.executeSQL(command)
            obj_hsOld.IdHSCu = command.Parameters("@_IdHSCu").Value.ToString
            Return obj_hsOld.IdHSCu
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ cũ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm cập nhật dữ liệu - Hồ sơ cũ của cán bộ
    ''' </summary>
    ''' <param name="obj_hsOld"></param>
    ''' <remarks></remarks>
    Public Sub Update_OldDocumnet(ByVal obj_hsOld As HS_Cu)
        Dim command As SqlCommand = New SqlCommand("HS_Cu_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSCu", obj_hsOld.IdHSCu))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_hsOld.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_TuThang", obj_hsOld.TuThang))
        command.Parameters.Add(New SqlParameter("@_DenThang", obj_hsOld.DenThang))
        command.Parameters.Add(New SqlParameter("@_NgheNghiep", obj_hsOld.NgheNghiep))
        command.Parameters.Add(New SqlParameter("@_DiaChi", obj_hsOld.DiaChi))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_hsOld.GhiChu))

        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ cũ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ cũ của cán bộ
    ''' </summary>
    ''' <param name="_IdHSCu">Id Hồ sơ cũ của cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Delete_OldDocumnet(ByVal _IdHSCu As String)
        Dim command As SqlCommand = New SqlCommand("HS_Cu_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdHSCu", SqlDbType.VarChar))
        command.Parameters("@_IdHSCu").Value = _IdHSCu
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Hồ sơ cũ", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    'Hàm cập nhật thông tin Học hàm - Học vị của cán bộ <Academic Title - Academic distinction>
    ''' <summary>
    ''' Hàm thêm mới dữ liệu - Hồ sơ học hàm của cán bộ
    ''' </summary>
    ''' <param name="_IdCB_HocHam"></param>
    ''' <param name="_IdCanBo"></param>
    ''' <param name="_IdHocHam"></param>
    ''' <remarks></remarks>
    Public Sub Insert_HocHam(ByVal _IdCB_HocHam As String, ByVal _IdCanBo As String, ByVal _IdHocHam As Int32)
        Dim command As SqlCommand = New SqlCommand("CB_HocHam_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdCB_HocHam", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
        command.Parameters("@_IdCanBo").Value = _IdCanBo

        command.Parameters.Add(New SqlParameter("@_IdHocHam", SqlDbType.Int))
        command.Parameters("@_IdHocHam").Value = _IdHocHam

        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ học hàm", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật dữ liệu - Hồ sơ học hàm của cán bộ
    ''' </summary>
    ''' <param name="_IdCB_HocHam"></param>
    ''' <param name="_IdCanBo"></param>
    ''' <param name="_IdHocHam"></param>
    ''' <remarks></remarks>
    Public Sub Update_HocHam(ByVal _IdCB_HocHam As String, ByVal _IdCanBo As String, ByVal _IdHocHam As Int32)
        Dim command As SqlCommand = New SqlCommand("CB_HocHam_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCB_HocHam", SqlDbType.VarChar))
        command.Parameters("@_IdCB_HocHam").Value = _IdCB_HocHam

        command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
        command.Parameters("@_IdCanBo").Value = _IdCanBo

        command.Parameters.Add(New SqlParameter("@_IdHocHam", SqlDbType.Int))
        command.Parameters("@_IdHocHam").Value = _IdHocHam

        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ học vị", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ lưu học hàm của cán bộ
    ''' </summary>
    ''' <param name="_IdCB_HocHam">Id Hồ sơ lưu học hàm của cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Delete_HocHam(ByVal _IdCB_HocHam As String)
        Dim command As SqlCommand = New SqlCommand("CB_HocHam_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCB_HocHam", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdCB_HocHam").Value = _IdCB_HocHam
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Học hàm", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thêm mới dữ liệu  Hồ sơ học vị
    ''' </summary>
    ''' <param name="_IdCB_HocVi">Id cán bộ học vị</param>
    ''' <param name="_IdCanBo">Id cán bộ</param>
    ''' <param name="_IdHocVi">Id học vị</param>
    ''' <remarks></remarks>
    Public Sub Insert_HocVi(ByVal _IdCB_HocVi As String, ByVal _IdCanBo As String, ByVal _IdHocVi As Int32)
        Dim command As SqlCommand = New SqlCommand("CB_HocVi_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdCB_HocVi", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
        command.Parameters("@_IdCanBo").Value = _IdCanBo

        command.Parameters.Add(New SqlParameter("@_IdHocVi", SqlDbType.Int))
        command.Parameters("@_IdHocVi").Value = _IdHocVi

        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Thêm mới hồ sơ học vị", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm cập nhật dữ liệu - Hồ sơ học vị
    ''' </summary>
    ''' <param name="_IdCB_HocVi">Id cán bộ học vị</param>
    ''' <param name="_IdCanBo">Id cán bộ</param>
    ''' <param name="_IdHocVi">Chỉ số xác định loại học vị</param>
    ''' <remarks></remarks>
    Public Sub Update_HocVi(ByVal _IdCB_HocVi As String, ByVal _IdCanBo As String, ByVal _IdHocVi As Int32)
        Dim command As SqlCommand = New SqlCommand("CB_HocVi_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCB_HocVi", SqlDbType.VarChar))
        command.Parameters("@_IdCB_HocVi").Value = _IdCB_HocVi

        command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
        command.Parameters("@_IdCanBo").Value = _IdCanBo

        command.Parameters.Add(New SqlParameter("@_IdHocVi", SqlDbType.Int))
        command.Parameters("@_IdHocVi").Value = _IdHocVi
        Try
            _SqlHelper.executeSQL(command)

        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Cập nhật hồ sơ học vị", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá bỏ dữ liệu - Hồ sơ lưu học vị của cán bộ
    ''' </summary>
    ''' <param name="_IdCB_HocVi">Id Hồ sơ lưu học vị của cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Delete_HocVi(ByVal _IdCB_HocVi As String)
        Dim command As SqlCommand = New SqlCommand("CB_HocVi_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdCB_HocVi", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@_IdCB_HocVi").Value = _IdCB_HocVi
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Thao tác không hợp lệ: " + ex.Message.ToString(), "Xoá dữ liệu - Học vị", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

#End Region

#Region "---> Functions: Liên quan tới Hồ sơ Gia đình cán bộ <---"

    ''' <summary>
    ''' Hàm Cập nhật Thêm mới dữ liệu - Hồ sơ Gia đình cán bộ
    ''' </summary>
    ''' <param name="obj_gdcb">Gia đình cán bộ</param>
    ''' <returns>Chuỗi Chỉ số xác định bản ghi HS gia đình cán bộ</returns>
    ''' <remarks></remarks>
    Public Function Insert_HS_GDCB(ByVal obj_gdcb As HS_GDCB) As String
        Dim command As SqlCommand = New SqlCommand("HS_GDCB_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdTVien", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_gdcb.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_Ma_TVien", obj_gdcb.Ma_TVien))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_gdcb.HoTen))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_gdcb.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NamSinh", obj_gdcb.NamSinh))
        command.Parameters.Add(New SqlParameter("@_IdQuanHe", obj_gdcb.IdQuanHe))
        command.Parameters.Add(New SqlParameter("@_ConMat", obj_gdcb.ConMat))
        command.Parameters.Add(New SqlParameter("@_NamMat", obj_gdcb.NamMat))
        command.Parameters.Add(New SqlParameter("@_LyDo_Mat", obj_gdcb.LyDo_Mat))
        command.Parameters.Add(New SqlParameter("@_IdUT_BThan", obj_gdcb.IdUT_BThan))
        command.Parameters.Add(New SqlParameter("@_IdQueQuan", obj_gdcb.IdQueQuan))
        command.Parameters.Add(New SqlParameter("@_DienThoai", obj_gdcb.DienThoai))
        command.Parameters.Add(New SqlParameter("@_DiaChi", obj_gdcb.DiaChi))
        command.Parameters.Add(New SqlParameter("@_IdQuocGia", obj_gdcb.IdQuocGia))
        command.Parameters.Add(New SqlParameter("@_NgheNghiep", obj_gdcb.NgheNghiep))

        Try
            _SqlHelper.executeSQL(command)
            obj_gdcb.IdTVien = command.Parameters("@_IdTVien").Value.ToString
            Return obj_gdcb.IdTVien
        Catch ex As Exception
            MessageBox.Show("Thêm mới hồ sơ gia đình cán bộ: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm Cập nhật Sửa đổi dữ liệu - Hồ sơ Gia đình cán bộ
    ''' </summary>
    ''' <param name="obj_gdcb">Hồ sơ Gia đình cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Update_HS_GDCB(ByVal obj_gdcb As HS_GDCB)
        Dim command As SqlCommand = New SqlCommand("HS_GDCB_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdTVien", obj_gdcb.IdTVien))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_gdcb.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_Ma_TVien", obj_gdcb.Ma_TVien))
        command.Parameters.Add(New SqlParameter("@_HoTen", obj_gdcb.HoTen))
        command.Parameters.Add(New SqlParameter("@_GioiTinh", obj_gdcb.GioiTinh))
        command.Parameters.Add(New SqlParameter("@_NamSinh", obj_gdcb.NamSinh))
        command.Parameters.Add(New SqlParameter("@_IdQuanHe", obj_gdcb.IdQuanHe))
        command.Parameters.Add(New SqlParameter("@_ConMat", obj_gdcb.ConMat))
        command.Parameters.Add(New SqlParameter("@_NamMat", obj_gdcb.NamMat))
        command.Parameters.Add(New SqlParameter("@_LyDo_Mat", obj_gdcb.LyDo_Mat))
        command.Parameters.Add(New SqlParameter("@_IdUT_BThan", obj_gdcb.IdUT_BThan))
        command.Parameters.Add(New SqlParameter("@_IdQueQuan", obj_gdcb.IdQueQuan))
        command.Parameters.Add(New SqlParameter("@_DienThoai", obj_gdcb.DienThoai))
        command.Parameters.Add(New SqlParameter("@_DiaChi", obj_gdcb.DiaChi))
        command.Parameters.Add(New SqlParameter("@_IdQuocGia", obj_gdcb.IdQuocGia))
        command.Parameters.Add(New SqlParameter("@_NgheNghiep", obj_gdcb.NgheNghiep))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Cập nhật hồ sơ gia đình cán bộ: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá dữ liệu - Hồ sơ Gia đình cán bộ
    ''' </summary>
    ''' <param name="_IdTVien">Hồ sơ Gia đình cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Delete_HS_GDCB(ByVal _IdTVien As String)
        Dim command As SqlCommand = New SqlCommand("HS_GDCB_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@IdTVien", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@IdTVien").Value = _IdTVien
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu - Hồ sơ gia đình cán bộ: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm Cập nhật Thêm mới dữ liệu - Hồ sơ Khám sức khoẻ Gia đình cán bộ
    ''' </summary>
    ''' <param name="obj_sk_gdcb">Hồ sơ Khám sức khoẻ Gia đình cán bộ</param>
    ''' <returns>Chuỗi chỉ số xác định bản ghi</returns>
    ''' <remarks></remarks>
    Public Function Insert_HS_SKhGDCB(ByVal obj_sk_gdcb As HS_SKhGDCB) As String
        Dim command As SqlCommand = New SqlCommand("HS_SKhGDCB_Insert")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add("@_IdKhamchuaGDCB", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
        command.Parameters.Add(New SqlParameter("@_Nam", obj_sk_gdcb.Nam))
        command.Parameters.Add(New SqlParameter("@_Dot", obj_sk_gdcb.Dot))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_sk_gdcb.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_IdTVien", obj_sk_gdcb.IdTVien))
        command.Parameters.Add(New SqlParameter("@_NoiDungKC", obj_sk_gdcb.NoiDungKC))
        command.Parameters.Add(New SqlParameter("@_NoiKham", obj_sk_gdcb.NoiKham))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_sk_gdcb.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_sk_gdcb.DenNgay))
        command.Parameters.Add(New SqlParameter("@_CanNang", obj_sk_gdcb.CanNang))
        command.Parameters.Add(New SqlParameter("@_ChieuCao", obj_sk_gdcb.ChieuCao))
        command.Parameters.Add(New SqlParameter("@_LoaiSuckhoe", obj_sk_gdcb.LoaiSuckhoe))
        command.Parameters.Add(New SqlParameter("@_KetLuan", obj_sk_gdcb.KetLuan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_sk_gdcb.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
            obj_sk_gdcb.IdKhamchuaGDCB = command.Parameters("@_IdKhamchuaGDCB").Value.ToString
            Return obj_sk_gdcb.IdKhamchuaGDCB
        Catch ex As Exception
            MessageBox.Show("Thêm mới hồ sơ khám chữa bệnh của gia đình cán bộ: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm Cập nhật Sửa đổi dữ liệu - Hồ sơ Khám sức khoẻ Gia đình cán bộ
    ''' </summary>
    ''' <param name="obj_sk_gdcb">Hồ sơ Khám sức khoẻ Gia đình cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Update_HS_SKhGDCB(ByVal obj_sk_gdcb As HS_SKhGDCB)
        Dim command As SqlCommand = New SqlCommand("HS_SKhGDCB_Update")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@_IdKhamchuaGDCB", obj_sk_gdcb.IdKhamchuaGDCB))
        command.Parameters.Add(New SqlParameter("@_Nam", obj_sk_gdcb.Nam))
        command.Parameters.Add(New SqlParameter("@_Dot", obj_sk_gdcb.Dot))
        command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_sk_gdcb.IdCanBo))
        command.Parameters.Add(New SqlParameter("@_IdTVien", obj_sk_gdcb.IdTVien))
        command.Parameters.Add(New SqlParameter("@_NoiDungKC", obj_sk_gdcb.NoiDungKC))
        command.Parameters.Add(New SqlParameter("@_NoiKham", obj_sk_gdcb.NoiKham))
        command.Parameters.Add(New SqlParameter("@_TuNgay", obj_sk_gdcb.TuNgay))
        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_sk_gdcb.DenNgay))
        command.Parameters.Add(New SqlParameter("@_CanNang", obj_sk_gdcb.CanNang))
        command.Parameters.Add(New SqlParameter("@_ChieuCao", obj_sk_gdcb.ChieuCao))
        command.Parameters.Add(New SqlParameter("@_LoaiSuckhoe", obj_sk_gdcb.LoaiSuckhoe))
        command.Parameters.Add(New SqlParameter("@_KetLuan", obj_sk_gdcb.KetLuan))
        command.Parameters.Add(New SqlParameter("@_GhiChu", obj_sk_gdcb.GhiChu))

        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Sửa đổi hồ sơ khám chữa bệnh của gia đình cán bộ: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện xoá dữ liệu - Hồ sơ khám sức khoẻ gia đình cán bộ
    ''' </summary>
    ''' <param name="_IdKhamchuaGDCB">Hồ sơ khám sức khoẻ gia đình cán bộ</param>
    ''' <remarks></remarks>
    Public Sub Delete_HS_SKhGDCB(ByVal _IdKhamchuaGDCB As String)
        Dim command As SqlCommand = New SqlCommand("HS_SKhGDCB_Delete")
        command.CommandType = CommandType.StoredProcedure

        command.Parameters.Add(New SqlParameter("@IdKhamchuaGDCB", SqlDbType.VarChar))
        command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
        command.Parameters("@IdKhamchuaGDCB").Value = _IdKhamchuaGDCB
        Try
            _SqlHelper.executeSQL(command)
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu - Hồ sơ khám sức khoẻ gia đình cán bộ: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ' ''' <summary>
    ' ''' Hàm Cập nhật Thêm mới dữ liệu - Hồ sơ Giảm trừ gia cảnh của cán bộ
    ' ''' </summary>
    ' ''' <param name="obj_gtgc">Hồ sơ Giảm trừ gia cảnh</param>
    ' ''' <returns>Chuỗi chỉ số xác định bản ghi</returns>
    ' ''' <remarks></remarks>
    'Public Function Insert_HS_GTGC(ByVal obj_gtgc As HS_GTGC) As String
    '    Dim command As SqlCommand = New SqlCommand("HS_GTGC_Insert")
    '    command.CommandType = CommandType.StoredProcedure

    '    command.Parameters.Add("@_IdGTGC", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
    '    command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
    '    command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_gtgc.IdCanBo))
    '    command.Parameters.Add(New SqlParameter("@_IdQuanHe", obj_gtgc.IdQuanHe))
    '    command.Parameters.Add(New SqlParameter("@_HoTenNguoiPT", obj_gtgc.HoTenNguoiPT))
    '    command.Parameters.Add(New SqlParameter("@_TuNgay", obj_gtgc.TuNgay))
    '    If (CType(obj_gtgc.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_gtgc.DenNgay))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_GhiChu", obj_gtgc.GhiChu))

    '    Try
    '        _SqlHelper.executeSQL(command)
    '        obj_gtgc.IdGTGC = command.Parameters("@_IdGTGC").Value.ToString
    '        Return obj_gtgc.IdGTGC
    '    Catch ex As Exception
    '        MessageBox.Show("Thêm mới hồ sơ giảm trừ gia cảnh: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '        Return Nothing
    '    End Try
    'End Function

    ' ''' <summary>
    ' ''' Hàm Cập nhật Sửa đổi dữ liệu - Hồ sơ Giảm trừ gia cảnh của cán bộ
    ' ''' </summary>
    ' ''' <param name="obj_gtgc">Hồ sơ Giảm trừ gia cảnh của cán bộ</param>
    ' ''' <remarks></remarks>
    'Public Sub Update_HS_GTGC(ByVal obj_gtgc As HS_GTGC)
    '    Dim command As SqlCommand = New SqlCommand("HS_GTGC_Update")
    '    command.CommandType = CommandType.StoredProcedure

    '    command.Parameters.Add(New SqlParameter("@_IdGTGC", obj_gtgc.IdGTGC))
    '    command.Parameters.Add(New SqlParameter("@_IdCanBo", obj_gtgc.IdCanBo))
    '    command.Parameters.Add(New SqlParameter("@_IdQuanHe", obj_gtgc.IdQuanHe))
    '    command.Parameters.Add(New SqlParameter("@_HoTenNguoiPT", obj_gtgc.HoTenNguoiPT))
    '    command.Parameters.Add(New SqlParameter("@_TuNgay", obj_gtgc.TuNgay))
    '    If (CType(obj_gtgc.DenNgay, DateTime).ToString("dd-MM-yyyy") = "01-01-1900") Then
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
    '    Else
    '        command.Parameters.Add(New SqlParameter("@_DenNgay", obj_gtgc.DenNgay))
    '    End If
    '    command.Parameters.Add(New SqlParameter("@_GhiChu", obj_gtgc.GhiChu))

    '    Try
    '        _SqlHelper.executeSQL(command)
    '    Catch ex As Exception
    '        MessageBox.Show("Sửa đổi hồ sơ Giảm trừ gia cảnh: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub


    Public Function Insert_Update_HS_GTGC(ByVal _HS_GTGC As HsCanBo.HS_GTGC) As String
        Dim _Ret As String = ""
        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "HS_GTGC_Insert_Update"

                    _command.Parameters.Add("@_IdOutPut", SqlDbType.VarChar, 16).Direction = ParameterDirection.Output
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_PosCode", _HS_GTGC.PosCode, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdGTGC", _HS_GTGC.IdGTGC, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdCanBo", _HS_GTGC.IdCanBo, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdQuanHe", _HS_GTGC.IdQuanHe, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_HoTenNguoiPT", _HS_GTGC.HoTenNguoiPT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_IdQuocTich", _HS_GTGC.IdQuocTich, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgaySinh", _HS_GTGC.NgaySinh, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TuNgay", _HS_GTGC.TuNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DenNgay", _HS_GTGC.DenNgay, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_MaSoThueNPT", _HS_GTGC.MaSoThueNPT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DiaChi", _HS_GTGC.DiaChi, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_DienThoai", _HS_GTGC.DienThoai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_SoCMT", _HS_GTGC.SoCMT, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NgayCap", _HS_GTGC.NgayCap, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_NoiCap", _HS_GTGC.NoiCap, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_GhiChu", _HS_GTGC.GhiChu, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_TrangThai", _HS_GTGC.TrangThai, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_CreatedBy", _HS_GTGC.CreatedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _HS_GTGC.ModifiedBy, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _HS_GTGC.IdOutPut = _command.Parameters("@_IdOutPut").Value.ToString()
                        _Ret = _HS_GTGC.IdOutPut
                    End If

                Catch ex As Exception
                    MessageBox.Show("Lỗi Cập nhật thông tin Giảm trừ gia cảnh (Thông tin người phụ thuộc): " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Globals.Logger.Error("Lỗi Cập nhật Giảm trừ gia cảnh của cán bộ Insert_Update_HS_GTGC(): " + ex.Message)
                    _Ret = ""
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    ''' <summary>
    ''' Hàm xóa/đánh dấu xóa Hồ sơ Giảm trừ gia cảnh (Người phụ thuộc của cán bộ)
    ''' </summary>
    ''' <param name="_RowId">Chỉ số khóa cần xóa</param>
    ''' <param name="_ModifiedBy">UserName thực hiện xóa</param>
    ''' <param name="_FlagDelete">Cờ xác định xóa/tạm xóa. Giá trị quy ước:
    '''                          1 - Xóa hẳn dữ liệu theo IdGTGC; 
    '''                          2 - Đánh dấu xóa bản ghi theo IdGTGC
    '''                          3 - Xóa hẳn dữ liệu theo IdCanBo; 
    '''                          4 - Đánh dấu xóa bản ghi theo IdCanBo
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Delete_HS_GTGC(ByVal _RowId As String, ByVal _ModifiedBy As String, _FlagDelete As Byte) As Boolean
        Dim _Ret As Boolean = False

        Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
            Using _command As System.Data.SqlClient.SqlCommand = New SqlCommand()
                Try
                    _command.Connection = conn_obj
                    _command.CommandType = CommandType.StoredProcedure
                    _command.CommandText = "HS_GTGC_Delete"

                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_RowId", _RowId, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_ModifiedBy", _ModifiedBy, ParameterDirection.Input))
                    _command.Parameters.Add(SoftSqlHelper.CreateParameter("@_FlagDelete", _FlagDelete, ParameterDirection.Input))

                    Dim iExecute As Integer = _command.ExecuteNonQuery()
                    If iExecute > 0 Then
                        _Ret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show("Lỗi xẩy ra Hồ sơ giảm trừ gia cảnh (Delete_HS_GTGC): " + ex.Message.ToString(), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    _Ret = False
                Finally
                    DbCommon.CloseConnection(conn_obj)
                End Try
            End Using
        End Using
        Return _Ret
    End Function

    ''' <summary>
    ''' Hàm thực hiện xoá dữ liệu - Hồ sơ Giảm trừ gia cảnh của cán bộ
    ''' </summary>
    ''' <param name="_IdGTGC">Chỉ số xác định mã hiệu - Giảm trừ gia cảnh của cán bộ</param>
    ''' <remarks></remarks>
    'Public Sub Delete_HS_GTGC_Xoa(ByVal _IdGTGC As String)
    '    Dim command As SqlCommand = New SqlCommand("HS_GTGC_Delete")
    '    command.CommandType = CommandType.StoredProcedure

    '    command.Parameters.Add(New SqlParameter("@IdGTGC", SqlDbType.VarChar))
    '    command.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
    '    command.Parameters("@IdGTGC").Value = _IdGTGC
    '    Try
    '        _SqlHelper.executeSQL(command)
    '    Catch ex As Exception
    '        MessageBox.Show("Xoá dữ liệu - Hồ sơ giảm trừ gia cảnh: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub


    ''' <summary>
    ''' Hàm trả về danh sách các hồ sơ liên quan đến Gia đình cán bộ
    ''' </summary>
    ''' <param name="_IdCanBo">Mã hiệu xác định cán bộ</param>
    ''' <param name="_State">Chỉ số phân loại hồ sơ. Với quy ước như sau:
    '''                                     0 - Hồ sơ gia đình cán bộ
    '''                                     1 - Hồ sơ Khám sức khoẻ gia đình cán bộ 
    '''                                     2 - Hồ sơ Giảm trừ gia cảnh
    '''                                     3 - Hồ sơ thong tin thu nhap gia dinh
    ''' </param>
    ''' <returns>Danh sách các bản ghi</returns>
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
                Case 0
                    command = New SqlCommand("HS_GDCB_GetForIdCanBo", connection)
                Case 1
                    command = New SqlCommand("HS_SKhGDCB_GetForIdCanBo", connection)
                Case 2
                    command = New SqlCommand("HS_GTGC_GetSearch", connection)
                Case 3
                    command = New SqlCommand("HS_ThuNhapGD_GetForIdCanBo", connection)
            End Select

            command.CommandType = CommandType.StoredProcedure
            If _State = 2 Then
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViCd", "", ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhongBanCd", "", ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdCanBo", _IdCanBo, ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdGTGC", "", ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdQuanHe", 0, ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pSoCMTNPT", "", ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pMaSoThueNPT", "", ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiemDL", "", ParameterDirection.Input))
                command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIsAll", 1, ParameterDirection.Input))
            Else
                command.Parameters.Add(New SqlParameter("@_IdCanBo", SqlDbType.VarChar))
                command.Parameters("@_IdCanBo").Value = _IdCanBo

            End If
            
            Using mydap As SqlDataAdapter = New SqlDataAdapter(command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "TableRows")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("TableRows")
            End Using
        Catch ex As Exception
            MessageBox.Show("Lấy dữ liệu hồ sơ về gia đình cán bộ: " + ex.Message.ToString(), "Lỗi xẩy ra", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then connection.Close()
        End Try

    End Function

    ''' <summary>
    ''' Hàm trả về bản ghi Hồ sơ liên quan tới Gia đình cán bộ
    ''' </summary>
    ''' <param name="_RecordId">Mã hiệu xác định bản ghi</param>
    ''' <param name="_State">Chỉ số phân loại hồ sơ. Với quy ước như sau:
    '''                                     0 - Hồ sơ gia đình cán bộ
    '''                                     1 - Hồ sơ Khám sức khoẻ gia đình cán bộ 
    '''                                     2 - Hồ sơ Giảm trừ gia cảnh
    '''                                     3 - Hồ sơ thu nhap gia dinh
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRecord(ByVal _RecordId As String, ByVal _State As Byte) As DataRow
        Dim connection As SqlConnection = _SqlHelper.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If
        Try
            Dim command As SqlCommand
            Select Case _State
                Case 0
                    command = New SqlCommand("HS_GDCB_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@_IdTVien", SqlDbType.VarChar))
                    command.Parameters("@_IdTVien").Value = _RecordId
                Case 1
                    command = New SqlCommand("HS_SKhGDCB_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@_IdKhamchuaGDCB", SqlDbType.VarChar))
                    command.Parameters("@_IdKhamchuaGDCB").Value = _RecordId

                Case 2
                    command = New SqlCommand("HS_GTGC_GetSearch", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViCd", "", ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhongBanCd", "", ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdCanBo", "", ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdGTGC", _RecordId, ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdQuanHe", 0, ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pSoCMTNPT", "", ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pMaSoThueNPT", "", ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiemDL", "", ParameterDirection.Input))
                    command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIsAll", 1, ParameterDirection.Input))
                Case 3
                    command = New SqlCommand("HS_ThuNhapGD_GetForId", connection)
                    command.CommandType = CommandType.StoredProcedure

                    command.Parameters.Add(New SqlParameter("@IdThuNhapGD", SqlDbType.VarChar))
                    command.Parameters("@IdThuNhapGD").Value = _RecordId
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


    Public Function GetGTGCSearch(ByVal pDonViCd As String, ByVal pPhongBanCd As String, ByVal pIdCanBo As String, ByVal pIdGTGC As String, ByVal pIdQuanHe As Integer, ByVal pSoCMTNPT As String, ByVal pMaSoThueNPT As String, ByVal pThoiDiemDL As String, ByVal pIsAll As Byte) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("HS_GTGC_GetSearch", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViCd", pDonViCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pPhongBanCd", pPhongBanCd, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdCanBo", pIdCanBo, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdGTGC", pIdGTGC, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIdQuanHe", pIdQuanHe, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pSoCMTNPT", pSoCMTNPT, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pMaSoThueNPT", pMaSoThueNPT, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiemDL", pThoiDiemDL, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pIsAll", pIsAll, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "HS_GTGC_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("HS_GTGC_TMP")
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
#End Region

#Region "---> Functions: Hàm lấy dữ liệu khác <---"
    ''' <summary>
    ''' Hàm thực hiện trả về id phòng ban hiện đang làm của cán bộ theo id cán bộ và id đơn vị truyền vào
    ''' </summary>
    ''' <param name="_HumanId">Chỉ số xác định cán bộ</param>
    ''' <param name="_DonviId">Chỉ số xác định chi nhánh</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDebtId(ByVal _HumanId As String, ByVal _DonviId As Integer) As String
        Dim debtId As Integer = 0
        Dim strSQL As String = String.Format("Select Top 1 * From QDNhanSu as a, HS_CanBo as b Where a.IdCanBo = b.IdCanBo And a.IdCanBo = '{0}' And a.IdDonvi_moi  IN (Select id From ChiNhanh Where id = {1} or id_goc = {1}) ORDER BY NgayKy_QD Desc", _HumanId, _DonviId)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    debtId = IIf(db.Rows(0)("IdPhong_Moi").ToString() <> "", CType(db.Rows(0)("IdPhong_Moi").ToString(), Integer), 0)
                End If
            End If
        End Using
        Return debtId
    End Function

    ''' <summary>
    ''' Hàm thực hiện trả về giá trị lấy từ tham số hệ thống
    ''' </summary>
    Public Function GetVarNam(ByVal _VarName As String) As String
        Dim _result As String = ""
        _result = _SqlHelper.getString(String.Format("Select {1} From SysVar Where ID_DonVi = {0}", IdDONVI, _VarName))
        Return _result
    End Function

    ''' <summary>
    ''' Hàm trả về chức vụ của Giám đốc đơn vị (Nếu là chi nhánh) hoặc Tổng Giám Đốc (Nếu cán bộ thuộc HSC) có cán bộ cần thao tác
    ''' </summary>
    ''' <param name="_IdcanBo">Id cán bộ</param>
    ''' <returns>Chỉ số xác định id của bản ghi chức vụ trong bảng DanhMuc</returns>
    ''' <remarks></remarks>
    Public Function GetDefault_ChucVu(ByVal _IdCanBo As String) As Integer
        Dim strSQL As String = ""
        strSQL = "SELECT a.HoTen, a.IdDonVi, b.ma_so FROM HS_CanBo a, ChiNhanh b Where a.IdCanBo = '" + _IdCanBo + "' And "
        strSQL += " (a.IdDonVi = b.id And b.id_goc <=1 And b.Status = 1) Order By a.IdDonVi Asc"
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If db.Rows.Count > 0 Then
                    strSQL = ""
                    If db.Rows(0)("ma_so").ToString().Trim() <> "" And db.Rows(0)("ma_so").ToString().Trim() = gMaDonViTW Then
                        strSQL = "SELECT * FROM danhmuc WHERE id_goc = 14 And ma_so = '1402' And Status = 1"
                    Else
                        strSQL = "SELECT * FROM danhmuc WHERE id_goc = 14 And ma_so = '1410' And Status = 1"
                    End If
                    Using db_child As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db_child Is Nothing) Then
                            If db_child.Rows.Count > 0 Then
                                If (db_child.Rows(0)("id").ToString().Trim() <> "") Then
                                    Return CType(db_child.Rows(0)("id").ToString().Trim(), Integer)
                                End If
                            End If
                        End If
                    End Using
                End If
            End If
        End Using
        Return 0
    End Function
#End Region

End Class




