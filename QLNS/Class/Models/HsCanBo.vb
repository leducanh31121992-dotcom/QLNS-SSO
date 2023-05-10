Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Public Class HsCanBo
    Public Class DsCanBo
        Private _Id As Integer
        Private _STT As String
        Private _IdDonVi As Integer
        Private _TenDV As String
        Private _IdDonVi_Child As Integer
        Private _TenDV_Child As String
        Private _IdCanBo As String
        Private _HoTen As String
        Private _MaCB As String
        Private _NgaySinh As DateTime
        Private _GioiTinh As String
        Private _CMT_So As String
        Private _CMT_NgayCap As DateTime
        Private _CMT_NoiCap As String
        Private _NS_DiaChi As String
        Private _NQ_DiaChi As String
        Private _TTr_DiaChi As String
        Private _ThT_DiaChi As String
        Private _IdDanToc As Integer
        Private _DanToc As String
        Private _IdTonGiao As Integer
        Private _TonGiao As String
        Private _IdTrinhDoCT As Integer
        Private _TrinhDoCT As String
        Private _DienThoai_DD As String
        Private _Email As String
        Private _QDNS_CV_PB_ChMon As String
        Private _QDNS_IdLoaiQD As Integer
        Private _QDNS_NgayHL As DateTime
        Private _QDNS_IdDonVi As Integer
        Private _QDNS_DonVi As String
        Private _QDNS_IdPhongBan As Integer
        Private _QDNS_PhongBan As String
        Private _QDNS_IdChucVu As Integer
        Private _QDNS_ChucVu As String
        Private _QDNS_IdChuyenMon As Integer
        Private _QDNS_ChuyenMon As String
        Private _CVTruocTuyenDung As String
        Private _NgayTiepNhan As DateTime
        Private _NgayTuyenDung As DateTime
        Private _DVTuyenDung As String
        Private _QDNS_Ngay_BN_BNL As DateTime
        Private _DV_IdDangVien As String
        Private _DV_NgayCThuc As DateTime
        Private _DV_SoThe As String
        Private _DV_NoiCapThe As String
        Private _DT_IdTrinhDo As Integer
        Private _DT_TrinhDo As String
        Private _DT_CoSoDT As String
        Private _DT_IdHeDT As Integer
        Private _DT_HeDT As String
        Private _DT_IdChuyenNganh As Integer
        Private _DT_ChuyenNganh As String
        Private _DT_NamTN As Integer
        Private _DT_NgayHL As DateTime
        Private _LCB_NghiDinhId As Integer
        Private _LCB_NghiDinh As String
        Private _LCB_BangLuongId As Integer
        Private _LCB_BangLuong As String
        Private _LCB_NgachLuongId As Integer
        Private _LCB_NgachLuong As String
        Private _LCB_BacLuongId As Integer
        Private _LCB_BacLuong As String
        Private _LCB_NgayHuong As DateTime
        Private _LCB_HeSoLuong As Double
        Private _PCCV As Double
        Private _PCKV As Double
        Private _PCTN As Double
        Private _PCDH As Double
        Private _PCTH As Double
        Private _PCDang As Double
        Private _HDLD_IdLoaiHD As Integer
        Private _HDLD_LoaiHD As String
        Private _HDLD_So As String
        Private _HDLD_NgayHL As DateTime
        Private _HDLD_NgayKy As DateTime
        Private _HDLD_DVKyHD As String

        Public Sub New()
            _Id = 0
            _STT = ""
            _IdDonVi = 0
            _TenDV = ""
            _IdDonVi_Child = 0
            _TenDV_Child = ""
            _IdCanBo = ""
            _HoTen = ""
            _MaCB = ""
            _NgaySinh = Globals.GetDateTime_ForServerDB
            _GioiTinh = ""
            _CMT_So = ""
            _CMT_NgayCap = Globals.GetDateTime_ForServerDB
            _CMT_NoiCap = ""
            _NS_DiaChi = ""
            _NQ_DiaChi = ""
            _TTr_DiaChi = ""
            _ThT_DiaChi = ""
            _IdDanToc = 0
            _DanToc = ""
            _IdTonGiao = 0
            _TonGiao = ""
            _IdTrinhDoCT = 0
            _TrinhDoCT = ""
            _DienThoai_DD = ""
            _Email = ""
            _QDNS_CV_PB_ChMon = ""
            _QDNS_IdLoaiQD = 0
            _QDNS_NgayHL = Globals.GetDateTime_ForServerDB
            _QDNS_IdDonVi = 0
            _QDNS_DonVi = ""
            _QDNS_IdPhongBan = 0
            _QDNS_PhongBan = ""
            _QDNS_IdChucVu = 0
            _QDNS_ChucVu = ""
            _QDNS_IdChuyenMon = 0
            _QDNS_ChuyenMon = ""
            _CVTruocTuyenDung = ""
            _NgayTiepNhan = Globals.GetDateTime_ForServerDB
            _NgayTuyenDung = Globals.GetDateTime_ForServerDB
            _DVTuyenDung = ""
            _QDNS_Ngay_BN_BNL = Globals.GetDateTime_ForServerDB
            _DV_IdDangVien = ""
            _DV_NgayCThuc = Globals.GetDateTime_ForServerDB
            _DV_SoThe = ""
            _DV_NoiCapThe = ""
            _DT_IdTrinhDo = 0
            _DT_TrinhDo = ""
            _DT_CoSoDT = ""
            _DT_IdHeDT = 0
            _DT_HeDT = ""
            _DT_IdChuyenNganh = 0
            _DT_ChuyenNganh = ""
            _DT_NamTN = 0
            _DT_NgayHL = Globals.GetDateTime_ForServerDB
            _LCB_NghiDinhId = 0
            _LCB_NghiDinh = ""
            _LCB_BangLuongId = 0
            _LCB_BangLuong = ""
            _LCB_NgachLuongId = 0
            _LCB_NgachLuong = ""
            _LCB_BacLuongId = 0
            _LCB_BacLuong = ""
            _LCB_NgayHuong = Globals.GetDateTime_ForServerDB
            _LCB_HeSoLuong = 0
            _PCCV = 0
            _PCKV = 0
            _PCTN = 0
            _PCDH = 0
            _PCTH = 0
            _PCDang = 0
            _HDLD_IdLoaiHD = 0
            _HDLD_LoaiHD = ""
            _HDLD_So = ""
            _HDLD_NgayHL = Globals.GetDateTime_ForServerDB
            _HDLD_NgayKy = Globals.GetDateTime_ForServerDB
            _HDLD_DVKyHD = ""
        End Sub

        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property STT() As String
            Get
                Return _STT
            End Get
            Set(ByVal value As String)
                _STT = value
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

        Public Property TenDV() As String
            Get
                Return _TenDV
            End Get
            Set(ByVal value As String)
                _TenDV = value
            End Set
        End Property

        Public Property IdDonVi_Child() As Integer
            Get
                Return _IdDonVi_Child
            End Get
            Set(ByVal value As Integer)
                _IdDonVi_Child = value
            End Set
        End Property

        Public Property TenDV_Child() As String
            Get
                Return _TenDV_Child
            End Get
            Set(ByVal value As String)
                _TenDV_Child = value
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

        Public Property HoTen() As String
            Get
                Return _HoTen
            End Get
            Set(ByVal value As String)
                _HoTen = value
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

        Public Property NgaySinh() As DateTime
            Get
                Return _NgaySinh
            End Get
            Set(ByVal value As DateTime)
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

        Public Property CMT_So() As String
            Get
                Return _CMT_So
            End Get
            Set(ByVal value As String)
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
                _CMT_NoiCap = value
            End Set
        End Property

        Public Property NS_DiaChi() As String
            Get
                Return _NS_DiaChi
            End Get
            Set(ByVal value As String)
                _NS_DiaChi = value
            End Set
        End Property

        Public Property NQ_DiaChi() As String
            Get
                Return _NQ_DiaChi
            End Get
            Set(ByVal value As String)
                _NQ_DiaChi = value
            End Set
        End Property

        Public Property TTr_DiaChi() As String
            Get
                Return _TTr_DiaChi
            End Get
            Set(ByVal value As String)
                _TTr_DiaChi = value
            End Set
        End Property

        Public Property ThT_DiaChi() As String
            Get
                Return _ThT_DiaChi
            End Get
            Set(ByVal value As String)
                _ThT_DiaChi = value
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

        Public Property DanToc() As String
            Get
                Return _DanToc
            End Get
            Set(ByVal value As String)
                _DanToc = value
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

        Public Property TonGiao() As String
            Get
                Return _TonGiao
            End Get
            Set(ByVal value As String)
                _TonGiao = value
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

        Public Property TrinhDoCT() As String
            Get
                Return _TrinhDoCT
            End Get
            Set(ByVal value As String)
                _TrinhDoCT = value
            End Set
        End Property

        Public Property DienThoai_DD() As String
            Get
                Return _DienThoai_DD
            End Get
            Set(ByVal value As String)
                _DienThoai_DD = value
            End Set
        End Property

        Public Property Email() As String
            Get
                Return _Email
            End Get
            Set(ByVal value As String)
                _Email = value
            End Set
        End Property

        Public Property QDNS_CV_PB_ChMon() As String
            Get
                Return _QDNS_CV_PB_ChMon
            End Get
            Set(ByVal value As String)
                _QDNS_CV_PB_ChMon = value
            End Set
        End Property

        Public Property QDNS_IdLoaiQD() As Integer
            Get
                Return _QDNS_IdLoaiQD
            End Get
            Set(ByVal value As Integer)
                _QDNS_IdLoaiQD = value
            End Set
        End Property

        Public Property QDNS_NgayHL() As DateTime
            Get
                Return _QDNS_NgayHL
            End Get
            Set(ByVal value As DateTime)
                _QDNS_NgayHL = value
            End Set
        End Property

        Public Property QDNS_IdDonVi() As Integer
            Get
                Return _QDNS_IdDonVi
            End Get
            Set(ByVal value As Integer)
                _QDNS_IdDonVi = value
            End Set
        End Property

        Public Property QDNS_DonVi() As String
            Get
                Return _QDNS_DonVi
            End Get
            Set(ByVal value As String)
                _QDNS_DonVi = value
            End Set
        End Property

        Public Property QDNS_IdPhongBan() As Integer
            Get
                Return _QDNS_IdPhongBan
            End Get
            Set(ByVal value As Integer)
                _QDNS_IdPhongBan = value
            End Set
        End Property

        Public Property QDNS_PhongBan() As String
            Get
                Return _QDNS_PhongBan
            End Get
            Set(ByVal value As String)
                _QDNS_PhongBan = value
            End Set
        End Property
         
        Public Property QDNS_IdChucVu() As Integer
            Get
                Return _QDNS_IdChucVu
            End Get
            Set(ByVal value As Integer)
                _QDNS_IdChucVu = value
            End Set
        End Property

        Public Property QDNS_ChucVu() As String
            Get
                Return _QDNS_ChucVu
            End Get
            Set(ByVal value As String)
                _QDNS_ChucVu = value
            End Set
        End Property

        Public Property QDNS_IdChuyenMon() As Integer
            Get
                Return _QDNS_IdChuyenMon
            End Get
            Set(ByVal value As Integer)
                _QDNS_IdChuyenMon = value
            End Set
        End Property

        Public Property QDNS_ChuyenMon() As String
            Get
                Return _QDNS_ChuyenMon
            End Get
            Set(ByVal value As String)
                _QDNS_ChuyenMon = value
            End Set
        End Property

        Public Property CVTruocTuyenDung() As String
            Get
                Return _CVTruocTuyenDung
            End Get
            Set(ByVal value As String)
                _CVTruocTuyenDung = value
            End Set
        End Property

        Public Property NgayTiepNhan() As DateTime
            Get
                Return _NgayTiepNhan
            End Get
            Set(ByVal value As DateTime)
                _NgayTiepNhan = value
            End Set
        End Property

        Public Property NgayTuyenDung() As DateTime
            Get
                Return _NgayTuyenDung
            End Get
            Set(ByVal value As DateTime)
                _NgayTuyenDung = value
            End Set
        End Property

        Public Property DVTuyenDung() As String
            Get
                Return _DVTuyenDung
            End Get
            Set(ByVal value As String)
                _DVTuyenDung = value
            End Set
        End Property

        Public Property QDNS_Ngay_BN_BNL() As DateTime
            Get
                Return _QDNS_Ngay_BN_BNL
            End Get
            Set(ByVal value As DateTime)
                _QDNS_Ngay_BN_BNL = value
            End Set
        End Property
        
        Public Property DV_IdDangVien() As String
            Get
                Return _DV_IdDangVien
            End Get
            Set(ByVal value As String)
                _DV_IdDangVien = value
            End Set
        End Property

        Public Property DV_NgayCThuc() As DateTime
            Get
                Return _DV_NgayCThuc
            End Get
            Set(ByVal value As DateTime)
                _DV_NgayCThuc = value
            End Set
        End Property

        Public Property DV_SoThe() As String
            Get
                Return _DV_SoThe
            End Get
            Set(ByVal value As String)
                _DV_SoThe = value
            End Set
        End Property

        Public Property DV_NoiCapThe() As String
            Get
                Return _DV_NoiCapThe
            End Get
            Set(ByVal value As String)
                _DV_NoiCapThe = value
            End Set
        End Property

        Public Property DT_IdTrinhDo() As Integer
            Get
                Return _DT_IdTrinhDo
            End Get
            Set(ByVal value As Integer)
                _DT_IdTrinhDo = value
            End Set
        End Property

        Public Property DT_TrinhDo() As String
            Get
                Return _DT_TrinhDo
            End Get
            Set(ByVal value As String)
                _DT_TrinhDo = value
            End Set
        End Property

        Public Property DT_CoSoDT() As String
            Get
                Return _DT_CoSoDT
            End Get
            Set(ByVal value As String)
                _DT_CoSoDT = value
            End Set
        End Property

        Public Property DT_IdHeDT() As Integer
            Get
                Return _DT_IdHeDT
            End Get
            Set(ByVal value As Integer)
                _DT_IdHeDT = value
            End Set
        End Property

        Public Property DT_HeDT() As String
            Get
                Return _DT_HeDT
            End Get
            Set(ByVal value As String)
                _DT_HeDT = value
            End Set
        End Property
         
        Public Property DT_IdChuyenNganh() As Integer
            Get
                Return _DT_IdChuyenNganh
            End Get
            Set(ByVal value As Integer)
                _DT_IdChuyenNganh = value
            End Set
        End Property

        Public Property DT_ChuyenNganh() As String
            Get
                Return _DT_ChuyenNganh
            End Get
            Set(ByVal value As String)
                _DT_ChuyenNganh = value
            End Set
        End Property

        Public Property DT_NamTN() As Integer
            Get
                Return _DT_NamTN
            End Get
            Set(ByVal value As Integer)
                _DT_NamTN = value
            End Set
        End Property

        Public Property DT_NgayHL() As DateTime
            Get
                Return _DT_NgayHL
            End Get
            Set(ByVal value As DateTime)
                _DT_NgayHL = value
            End Set
        End Property

        Public Property LCB_NghiDinhId() As Integer
            Get
                Return _LCB_NghiDinhId
            End Get
            Set(ByVal value As Integer)
                _LCB_NghiDinhId = value
            End Set
        End Property

        Public Property LCB_NghiDinh() As String
            Get
                Return _LCB_NghiDinh
            End Get
            Set(ByVal value As String)
                _LCB_NghiDinh = value
            End Set
        End Property

        Public Property LCB_BangLuongId() As Integer
            Get
                Return _LCB_BangLuongId
            End Get
            Set(ByVal value As Integer)
                _LCB_BangLuongId = value
            End Set
        End Property

        Public Property LCB_BangLuong() As String
            Get
                Return _LCB_BangLuong
            End Get
            Set(ByVal value As String)
                _LCB_BangLuong = value
            End Set
        End Property

        Public Property LCB_NgachLuongId() As Integer
            Get
                Return _LCB_NgachLuongId
            End Get
            Set(ByVal value As Integer)
                _LCB_NgachLuongId = value
            End Set
        End Property

        Public Property LCB_NgachLuong() As String
            Get
                Return _LCB_NgachLuong
            End Get
            Set(ByVal value As String)
                _LCB_NgachLuong = value
            End Set
        End Property

        Public Property LCB_BacLuongId() As Integer
            Get
                Return _LCB_BacLuongId
            End Get
            Set(ByVal value As Integer)
                _LCB_BacLuongId = value
            End Set
        End Property

        Public Property LCB_BacLuong() As String
            Get
                Return _LCB_BacLuong
            End Get
            Set(ByVal value As String)
                _LCB_BacLuong = value
            End Set
        End Property

        Public Property LCB_NgayHuong() As DateTime
            Get
                Return _LCB_NgayHuong
            End Get
            Set(ByVal value As DateTime)
                _LCB_NgayHuong = value
            End Set
        End Property

        Public Property LCB_HeSoLuong() As Double
            Get
                Return _LCB_HeSoLuong
            End Get
            Set(ByVal value As Double)
                _LCB_HeSoLuong = value
            End Set
        End Property

        Public Property PCCV() As Double
            Get
                Return _PCCV
            End Get
            Set(ByVal value As Double)
                _PCCV = value
            End Set
        End Property

        Public Property PCKV() As Double
            Get
                Return _PCKV
            End Get
            Set(ByVal value As Double)
                _PCKV = value
            End Set
        End Property

        Public Property PCTN() As Double
            Get
                Return _PCTN
            End Get
            Set(ByVal value As Double)
                _PCTN = value
            End Set
        End Property

        Public Property PCDH() As Double
            Get
                Return _PCDH
            End Get
            Set(ByVal value As Double)
                _PCDH = value
            End Set
        End Property

        Public Property PCTH() As Double
            Get
                Return _PCTH
            End Get
            Set(ByVal value As Double)
                _PCTH = value
            End Set
        End Property

        Public Property PCDang() As Double
            Get
                Return _PCDang
            End Get
            Set(ByVal value As Double)
                _PCDang = value
            End Set
        End Property
         
        Public Property HDLD_IdLoaiHD() As Integer
            Get
                Return _HDLD_IdLoaiHD
            End Get
            Set(ByVal value As Integer)
                _HDLD_IdLoaiHD = value
            End Set
        End Property

        Public Property HDLD_LoaiHD() As String
            Get
                Return _HDLD_LoaiHD
            End Get
            Set(ByVal value As String)
                _HDLD_LoaiHD = value
            End Set
        End Property

        Public Property HDLD_So() As String
            Get
                Return _HDLD_So
            End Get
            Set(ByVal value As String)
                _HDLD_So = value
            End Set
        End Property

        Public Property HDLD_NgayHL() As DateTime
            Get
                Return _HDLD_NgayHL
            End Get
            Set(ByVal value As DateTime)
                _HDLD_NgayHL = value
            End Set
        End Property

        Public Property HDLD_NgayKy() As DateTime
            Get
                Return _HDLD_NgayKy
            End Get
            Set(ByVal value As DateTime)
                _HDLD_NgayKy = value
            End Set
        End Property

        Public Property HDLD_DVKyHD() As String
            Get
                Return _HDLD_DVKyHD
            End Get
            Set(ByVal value As String)
                _HDLD_DVKyHD = value
            End Set
        End Property

        Public Sub GetData(_Reader As System.Data.SqlClient.SqlDataReader)
            Id = IIf(String.IsNullOrEmpty(_Reader("Id").ToString()), 0, CType(_Reader("Id").ToString(), Integer))
            STT = _Reader("STT").ToString()
            IdDonVi = IIf(String.IsNullOrEmpty(_Reader("IdDonVi").ToString()), 0, CType(_Reader("IdDonVi").ToString(), Integer))
            TenDV = _Reader("TenDV").ToString()
            IdDonVi_Child = IIf(String.IsNullOrEmpty(_Reader("IdDonVi_Child").ToString()), 0, CType(_Reader("IdDonVi_Child").ToString(), Integer))
            TenDV_Child = _Reader("TenDV_Child").ToString()
            IdCanBo = _Reader("IdCanBo").ToString()
            HoTen = _Reader("HoTen").ToString()
            MaCB = _Reader("MaCB").ToString()
            NgaySinh = IIf(String.IsNullOrEmpty(_Reader("NgaySinh").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("NgaySinh").ToString(), DateTime))
            GioiTinh = _Reader("GioiTinh").ToString()
            CMT_So = _Reader("CMT_So").ToString()
            CMT_NgayCap = IIf(String.IsNullOrEmpty(_Reader("CMT_NgayCap").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("CMT_NgayCap").ToString(), DateTime))
            CMT_NoiCap = _Reader("CMT_NoiCap").ToString()
            NS_DiaChi = _Reader("NS_DiaChi").ToString()
            NQ_DiaChi = _Reader("NQ_DiaChi").ToString()
            TTr_DiaChi = _Reader("TTr_DiaChi").ToString()
            ThT_DiaChi = _Reader("ThT_DiaChi").ToString()
            IdDanToc = IIf(String.IsNullOrEmpty(_Reader("IdDanToc").ToString()), 0, CType(_Reader("IdDanToc").ToString(), Integer))
            DanToc = _Reader("DanToc").ToString()
            IdTonGiao = IIf(String.IsNullOrEmpty(_Reader("IdTonGiao").ToString()), 0, CType(_Reader("IdTonGiao").ToString(), Integer))
            TonGiao = _Reader("TonGiao").ToString()
            IdTrinhDoCT = IIf(String.IsNullOrEmpty(_Reader("IdTrinhDoCT").ToString()), 0, CType(_Reader("IdTrinhDoCT").ToString(), Integer))
            TrinhDoCT = _Reader("TrinhDoCT").ToString()
            DienThoai_DD = _Reader("DienThoai_DD").ToString()
            Email = _Reader("Email").ToString()
            QDNS_CV_PB_ChMon = _Reader("QDNS_CV_PB_ChMon").ToString()
            QDNS_IdLoaiQD = IIf(String.IsNullOrEmpty(_Reader("QDNS_IdLoaiQD").ToString()), 0, CType(_Reader("QDNS_IdLoaiQD").ToString(), Integer))
            QDNS_NgayHL = IIf(String.IsNullOrEmpty(_Reader("QDNS_NgayHL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("QDNS_NgayHL").ToString(), DateTime))
            QDNS_IdDonVi = IIf(String.IsNullOrEmpty(_Reader("QDNS_IdDonVi").ToString()), 0, CType(_Reader("QDNS_IdDonVi").ToString(), Integer))
            QDNS_DonVi = _Reader("QDNS_DonVi").ToString()
            QDNS_IdPhongBan = IIf(String.IsNullOrEmpty(_Reader("QDNS_IdPhongBan").ToString()), 0, CType(_Reader("QDNS_IdPhongBan").ToString(), Integer))
            QDNS_PhongBan = _Reader("QDNS_PhongBan").ToString()
            QDNS_IdChucVu = IIf(String.IsNullOrEmpty(_Reader("QDNS_IdChucVu").ToString()), 0, CType(_Reader("QDNS_IdChucVu").ToString(), Integer))
            QDNS_ChucVu = _Reader("QDNS_ChucVu").ToString()
            QDNS_IdChuyenMon = IIf(String.IsNullOrEmpty(_Reader("QDNS_IdChuyenMon").ToString()), 0, CType(_Reader("QDNS_IdChuyenMon").ToString(), Integer))
            QDNS_ChuyenMon = _Reader("QDNS_ChuyenMon").ToString()
            CVTruocTuyenDung = _Reader("CVTruocTuyenDung").ToString()
            NgayTiepNhan = IIf(String.IsNullOrEmpty(_Reader("NgayTiepNhan").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("NgayTiepNhan").ToString(), DateTime))
            NgayTuyenDung = IIf(String.IsNullOrEmpty(_Reader("NgayTuyenDung").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("NgayTuyenDung").ToString(), DateTime))
            DVTuyenDung = _Reader("DVTuyenDung").ToString()
            QDNS_Ngay_BN_BNL = IIf(String.IsNullOrEmpty(_Reader("QDNS_Ngay_BN_BNL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("QDNS_Ngay_BN_BNL").ToString(), DateTime))
            DV_IdDangVien = _Reader("DV_IdDangVien").ToString()
            DV_NgayCThuc = IIf(String.IsNullOrEmpty(_Reader("DV_NgayCThuc").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("DV_NgayCThuc").ToString(), DateTime))
            DV_SoThe = _Reader("DV_SoThe").ToString()
            DV_NoiCapThe = _Reader("DV_NoiCapThe").ToString()
            DT_IdTrinhDo = IIf(String.IsNullOrEmpty(_Reader("DT_IdTrinhDo").ToString()), 0, CType(_Reader("DT_IdTrinhDo").ToString(), Integer))
            DT_TrinhDo = _Reader("DT_TrinhDo").ToString()
            DT_CoSoDT = _Reader("DT_CoSoDT").ToString()
            DT_IdHeDT = IIf(String.IsNullOrEmpty(_Reader("DT_IdHeDT").ToString()), 0, CType(_Reader("DT_IdHeDT").ToString(), Integer))
            DT_HeDT = _Reader("DT_HeDT").ToString()
            DT_IdChuyenNganh = IIf(String.IsNullOrEmpty(_Reader("DT_IdChuyenNganh").ToString()), 0, CType(_Reader("DT_IdChuyenNganh").ToString(), Integer))
            DT_ChuyenNganh = _Reader("DT_ChuyenNganh").ToString()
            DT_NamTN = IIf(String.IsNullOrEmpty(_Reader("DT_NamTN").ToString()), 0, CType(_Reader("DT_NamTN").ToString(), Integer))
            DT_NgayHL = IIf(String.IsNullOrEmpty(_Reader("DT_NgayHL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("DT_NgayHL").ToString(), DateTime))
            LCB_NghiDinhId = IIf(String.IsNullOrEmpty(_Reader("LCB_NghiDinhId").ToString()), 0, CType(_Reader("LCB_NghiDinhId").ToString(), Integer))
            LCB_NghiDinh = _Reader("LCB_NghiDinh").ToString()
            LCB_BangLuongId = IIf(String.IsNullOrEmpty(_Reader("LCB_BangLuongId").ToString()), 0, CType(_Reader("LCB_BangLuongId").ToString(), Integer))
            LCB_BangLuong = _Reader("LCB_BangLuong").ToString()
            LCB_NgachLuongId = IIf(String.IsNullOrEmpty(_Reader("LCB_NgachLuongId").ToString()), 0, CType(_Reader("LCB_NgachLuongId").ToString(), Integer))
            LCB_NgachLuong = _Reader("LCB_NgachLuong").ToString()
            LCB_BacLuongId = IIf(String.IsNullOrEmpty(_Reader("LCB_BacLuongId").ToString()), 0, CType(_Reader("LCB_BacLuongId").ToString(), Integer))
            LCB_BacLuong = _Reader("LCB_BacLuong").ToString()
            LCB_NgayHuong = IIf(String.IsNullOrEmpty(_Reader("LCB_NgayHuong").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("LCB_NgayHuong").ToString(), DateTime))
            LCB_HeSoLuong = IIf(String.IsNullOrEmpty(_Reader("LCB_HeSoLuong").ToString()), 0, CType(_Reader("LCB_HeSoLuong").ToString(), Double))
            PCCV = IIf(String.IsNullOrEmpty(_Reader("PCCV").ToString()), 0, CType(_Reader("PCCV").ToString(), Double))
            PCKV = IIf(String.IsNullOrEmpty(_Reader("PCKV").ToString()), 0, CType(_Reader("PCKV").ToString(), Double))
            PCTN = IIf(String.IsNullOrEmpty(_Reader("PCTN").ToString()), 0, CType(_Reader("PCTN").ToString(), Double))
            PCDH = IIf(String.IsNullOrEmpty(_Reader("PCDH").ToString()), 0, CType(_Reader("PCDH").ToString(), Double))
            PCTH = IIf(String.IsNullOrEmpty(_Reader("PCTH").ToString()), 0, CType(_Reader("PCTH").ToString(), Double))
            PCDang = IIf(String.IsNullOrEmpty(_Reader("PCDang").ToString()), 0, CType(_Reader("PCDang").ToString(), Double))
            HDLD_IdLoaiHD = IIf(String.IsNullOrEmpty(_Reader("HDLD_IdLoaiHD").ToString()), 0, CType(_Reader("HDLD_IdLoaiHD").ToString(), Integer))
            HDLD_LoaiHD = _Reader("HDLD_LoaiHD").ToString()
            HDLD_So = _Reader("HDLD_So").ToString()
            HDLD_NgayHL = IIf(String.IsNullOrEmpty(_Reader("HDLD_NgayHL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("HDLD_NgayHL").ToString(), DateTime))
            HDLD_NgayKy = IIf(String.IsNullOrEmpty(_Reader("HDLD_NgayKy").ToString()), DateTimeUtil.MinSqlDateTime, CType(_Reader("HDLD_NgayKy").ToString(), DateTime))
            HDLD_DVKyHD = _Reader("HDLD_DVKyHD").ToString()
        End Sub

        Public Sub GetData(_drow As System.Data.DataRow)
            Id = IIf(String.IsNullOrEmpty(_drow("Id").ToString()), 0, CType(_drow("Id").ToString(), Integer))
            STT = _drow("STT").ToString()
            IdDonVi = IIf(String.IsNullOrEmpty(_drow("IdDonVi").ToString()), 0, CType(_drow("IdDonVi").ToString(), Integer))
            TenDV = _drow("TenDV").ToString()
            IdDonVi_Child = IIf(String.IsNullOrEmpty(_drow("IdDonVi_Child").ToString()), 0, CType(_drow("IdDonVi_Child").ToString(), Integer))
            TenDV_Child = _drow("TenDV_Child").ToString()
            IdCanBo = _drow("IdCanBo").ToString()
            HoTen = _drow("HoTen").ToString()
            MaCB = _drow("MaCB").ToString()
            NgaySinh = IIf(String.IsNullOrEmpty(_drow("NgaySinh").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("NgaySinh").ToString(), DateTime))
            GioiTinh = _drow("GioiTinh").ToString()
            CMT_So = _drow("CMT_So").ToString()
            CMT_NgayCap = IIf(String.IsNullOrEmpty(_drow("CMT_NgayCap").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("CMT_NgayCap").ToString(), DateTime))
            CMT_NoiCap = _drow("CMT_NoiCap").ToString()
            NS_DiaChi = _drow("NS_DiaChi").ToString()
            NQ_DiaChi = _drow("NQ_DiaChi").ToString()
            TTr_DiaChi = _drow("TTr_DiaChi").ToString()
            ThT_DiaChi = _drow("ThT_DiaChi").ToString()
            IdDanToc = IIf(String.IsNullOrEmpty(_drow("IdDanToc").ToString()), 0, CType(_drow("IdDanToc").ToString(), Integer))
            DanToc = _drow("DanToc").ToString()
            IdTonGiao = IIf(String.IsNullOrEmpty(_drow("IdTonGiao").ToString()), 0, CType(_drow("IdTonGiao").ToString(), Integer))
            TonGiao = _drow("TonGiao").ToString()
            IdTrinhDoCT = IIf(String.IsNullOrEmpty(_drow("IdTrinhDoCT").ToString()), 0, CType(_drow("IdTrinhDoCT").ToString(), Integer))
            TrinhDoCT = _drow("TrinhDoCT").ToString()
            DienThoai_DD = _drow("DienThoai_DD").ToString()
            Email = _drow("Email").ToString()
            QDNS_CV_PB_ChMon = _drow("QDNS_CV_PB_ChMon").ToString()
            QDNS_IdLoaiQD = IIf(String.IsNullOrEmpty(_drow("QDNS_IdLoaiQD").ToString()), 0, CType(_drow("QDNS_IdLoaiQD").ToString(), Integer))
            QDNS_NgayHL = IIf(String.IsNullOrEmpty(_drow("QDNS_NgayHL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("QDNS_NgayHL").ToString(), DateTime))
            QDNS_IdDonVi = IIf(String.IsNullOrEmpty(_drow("QDNS_IdDonVi").ToString()), 0, CType(_drow("QDNS_IdDonVi").ToString(), Integer))
            QDNS_DonVi = _drow("QDNS_DonVi").ToString()
            QDNS_IdPhongBan = IIf(String.IsNullOrEmpty(_drow("QDNS_IdPhongBan").ToString()), 0, CType(_drow("QDNS_IdPhongBan").ToString(), Integer))
            QDNS_PhongBan = _drow("QDNS_PhongBan").ToString()
            QDNS_IdChucVu = IIf(String.IsNullOrEmpty(_drow("QDNS_IdChucVu").ToString()), 0, CType(_drow("QDNS_IdChucVu").ToString(), Integer))
            QDNS_ChucVu = _drow("QDNS_ChucVu").ToString()
            QDNS_IdChuyenMon = IIf(String.IsNullOrEmpty(_drow("QDNS_IdChuyenMon").ToString()), 0, CType(_drow("QDNS_IdChuyenMon").ToString(), Integer))
            QDNS_ChuyenMon = _drow("QDNS_ChuyenMon").ToString()
            CVTruocTuyenDung = _drow("CVTruocTuyenDung").ToString()
            NgayTiepNhan = IIf(String.IsNullOrEmpty(_drow("NgayTiepNhan").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("NgayTiepNhan").ToString(), DateTime))
            NgayTuyenDung = IIf(String.IsNullOrEmpty(_drow("NgayTuyenDung").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("NgayTuyenDung").ToString(), DateTime))
            DVTuyenDung = _drow("DVTuyenDung").ToString()
            QDNS_Ngay_BN_BNL = IIf(String.IsNullOrEmpty(_drow("QDNS_Ngay_BN_BNL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("QDNS_Ngay_BN_BNL").ToString(), DateTime))
            DV_IdDangVien = _drow("DV_IdDangVien").ToString()
            DV_NgayCThuc = IIf(String.IsNullOrEmpty(_drow("DV_NgayCThuc").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("DV_NgayCThuc").ToString(), DateTime))
            DV_SoThe = _drow("DV_SoThe").ToString()
            DV_NoiCapThe = _drow("DV_NoiCapThe").ToString()
            DT_IdTrinhDo = IIf(String.IsNullOrEmpty(_drow("DT_IdTrinhDo").ToString()), 0, CType(_drow("DT_IdTrinhDo").ToString(), Integer))
            DT_TrinhDo = _drow("DT_TrinhDo").ToString()
            DT_CoSoDT = _drow("DT_CoSoDT").ToString()
            DT_IdHeDT = IIf(String.IsNullOrEmpty(_drow("DT_IdHeDT").ToString()), 0, CType(_drow("DT_IdHeDT").ToString(), Integer))
            DT_HeDT = _drow("DT_HeDT").ToString()
            DT_IdChuyenNganh = IIf(String.IsNullOrEmpty(_drow("DT_IdChuyenNganh").ToString()), 0, CType(_drow("DT_IdChuyenNganh").ToString(), Integer))
            DT_ChuyenNganh = _drow("DT_ChuyenNganh").ToString()
            DT_NamTN = IIf(String.IsNullOrEmpty(_drow("DT_NamTN").ToString()), 0, CType(_drow("DT_NamTN").ToString(), Integer))
            DT_NgayHL = IIf(String.IsNullOrEmpty(_drow("DT_NgayHL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("DT_NgayHL").ToString(), DateTime))
            LCB_NghiDinhId = IIf(String.IsNullOrEmpty(_drow("LCB_NghiDinhId").ToString()), 0, CType(_drow("LCB_NghiDinhId").ToString(), Integer))
            LCB_NghiDinh = _drow("LCB_NghiDinh").ToString()
            LCB_BangLuongId = IIf(String.IsNullOrEmpty(_drow("LCB_BangLuongId").ToString()), 0, CType(_drow("LCB_BangLuongId").ToString(), Integer))
            LCB_BangLuong = _drow("LCB_BangLuong").ToString()
            LCB_NgachLuongId = IIf(String.IsNullOrEmpty(_drow("LCB_NgachLuongId").ToString()), 0, CType(_drow("LCB_NgachLuongId").ToString(), Integer))
            LCB_NgachLuong = _drow("LCB_NgachLuong").ToString()
            LCB_BacLuongId = IIf(String.IsNullOrEmpty(_drow("LCB_BacLuongId").ToString()), 0, CType(_drow("LCB_BacLuongId").ToString(), Integer))
            LCB_BacLuong = _drow("LCB_BacLuong").ToString()
            LCB_NgayHuong = IIf(String.IsNullOrEmpty(_drow("LCB_NgayHuong").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("LCB_NgayHuong").ToString(), DateTime))
            LCB_HeSoLuong = IIf(String.IsNullOrEmpty(_drow("LCB_HeSoLuong").ToString()), 0, CType(_drow("LCB_HeSoLuong").ToString(), Double))
            PCCV = IIf(String.IsNullOrEmpty(_drow("PCCV").ToString()), 0, CType(_drow("PCCV").ToString(), Double))
            PCKV = IIf(String.IsNullOrEmpty(_drow("PCKV").ToString()), 0, CType(_drow("PCKV").ToString(), Double))
            PCTN = IIf(String.IsNullOrEmpty(_drow("PCTN").ToString()), 0, CType(_drow("PCTN").ToString(), Double))
            PCDH = IIf(String.IsNullOrEmpty(_drow("PCDH").ToString()), 0, CType(_drow("PCDH").ToString(), Double))
            PCTH = IIf(String.IsNullOrEmpty(_drow("PCTH").ToString()), 0, CType(_drow("PCTH").ToString(), Double))
            PCDang = IIf(String.IsNullOrEmpty(_drow("PCDang").ToString()), 0, CType(_drow("PCDang").ToString(), Double))
            HDLD_IdLoaiHD = IIf(String.IsNullOrEmpty(_drow("HDLD_IdLoaiHD").ToString()), 0, CType(_drow("HDLD_IdLoaiHD").ToString(), Integer))
            HDLD_LoaiHD = _drow("HDLD_LoaiHD").ToString()
            HDLD_So = _drow("HDLD_So").ToString()
            HDLD_NgayHL = IIf(String.IsNullOrEmpty(_drow("HDLD_NgayHL").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("HDLD_NgayHL").ToString(), DateTime))
            HDLD_NgayKy = IIf(String.IsNullOrEmpty(_drow("HDLD_NgayKy").ToString()), DateTimeUtil.MinSqlDateTime, CType(_drow("HDLD_NgayKy").ToString(), DateTime))
            HDLD_DVKyHD = _drow("HDLD_DVKyHD").ToString()
        End Sub
    End Class

    Public Class HS_GTGC
        Private _IdOutPut As String
        Private _PosCode As String

        Private _IdGTGC As String
        Private _IdCanBo As String
        Private _IdQuanHe As Integer
        Private _HoTenNguoiPT As String
        Private _IdQuocTich As Integer
        Private _NgaySinh As DateTime
        Private _TuNgay As DateTime
        Private _DenNgay As DateTime
        Private _MaSoThueNPT As String
        Private _DiaChi As String
        Private _DienThoai As String
        Private _SoCMT As String
        Private _NgayCap As DateTime
        Private _NoiCap As String
        Private _GhiChu As String
        Private _TrangThai As Byte
        Private _CreatedBy As String
        Private _Date_Create As DateTime
        Private _ModifiedBy As String
        Private _Date_Update As DateTime

        Public Sub New()
            _IdOutPut = ""
            _PosCode = ""
            _IdGTGC = ""
            _IdCanBo = ""
            _IdQuanHe = 0
            _HoTenNguoiPT = ""
            _IdQuocTich = 0
            _NgaySinh = Globals.GetDateTime_ForServerDB
            _TuNgay = Globals.GetDateTime_ForServerDB
            _DenNgay = Globals.GetDateTime_ForServerDB
            _MaSoThueNPT = ""
            _DiaChi = ""
            _DienThoai = ""
            _SoCMT = ""
            _NgayCap = Globals.GetDateTime_ForServerDB
            _NoiCap = ""
            _GhiChu = ""
            _TrangThai = 1
            _CreatedBy = ""
            _Date_Create = Globals.GetDateTime_ForServerDB
            _ModifiedBy = ""
            _Date_Update = Globals.GetDateTime_ForServerDB
        End Sub

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
        Public Property IdGTGC() As String
            Get
                Return _IdGTGC
            End Get
            Set(ByVal value As String)
                _IdGTGC = value
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
                _HoTenNguoiPT = value
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
        Public Property NgaySinh() As DateTime
            Get
                Return _NgaySinh
            End Get
            Set(ByVal value As DateTime)
                _NgaySinh = value
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
        Public Property MaSoThueNPT() As String
            Get
                Return _MaSoThueNPT
            End Get
            Set(ByVal value As String)
                _MaSoThueNPT = value
            End Set
        End Property
        Public Property DiaChi() As String
            Get
                Return _DiaChi
            End Get
            Set(ByVal value As String)
                _DiaChi = value
            End Set
        End Property
        Public Property DienThoai() As String
            Get
                Return _DienThoai
            End Get
            Set(ByVal value As String)
                _DienThoai = value
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
                _NoiCap = value
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
        Public Property TrangThai() As Byte
            Get
                Return _TrangThai
            End Get
            Set(ByVal value As Byte)
                _TrangThai = value
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
        Public Property Date_Create() As DateTime
            Get
                Return _Date_Create
            End Get
            Set(ByVal value As DateTime)
                _Date_Create = value
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
        Public Property Date_Update() As DateTime
            Get
                Return _Date_Update
            End Get
            Set(ByVal value As DateTime)
                _Date_Update = value
            End Set
        End Property
       
    End Class
End Class
