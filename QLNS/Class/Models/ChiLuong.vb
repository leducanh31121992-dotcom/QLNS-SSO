Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Public Class ChiLuong_TongHop
    Private _IdOutPut As Long
    Private _TongHopId As Long
    Private _DonVi_CL_Cd As String
    Private _PhongBan_CL_Cd As String
    Private _NamBC As Integer
    Private _ThangBC As Integer
    Private _KyBC As Integer
    Private _PhanLoai_Cd As String
    Private _LaoDong_Cd As String
    Private _NgayBC As DateTime
    Private _GhiChu_CL As String
    Private _TrangThai As Byte
    Private _TrangThai_HT As String
    Private _CreatedBy As String
    Private _CreatedDate As DateTime
    Private _ModifiedBy As String
    Private _ModifiedDate As DateTime

    Private _DonVi_CL_HT As String
    Private _PhongBan_CL_HT As String
    Private _PhanLoai_HT As String
    Private _LaoDong_HT As String
    Private _ThongTin_HT As String

    Private _Luong_TTV_Ma As String
    Private _Luong_TTV As Double
    Private _Luong_CSo As Double
    Private _BHXH_CB As Double
    Private _BHXH_DV As Double
    Private _BHYT_CB As Double
    Private _BHYT_DV As Double
    Private _BHTN_CB As Double
    Private _BHTN_DV As Double
    Private _DPCD_CB As Double
    Private _DPCD_LD_NN As Double
    Private _Tinh_Thue_TNCN As Byte
    Private _TNTT_Muc_BanThan As Double
    Private _TNTT_Muc_PhuThuoc As Double
    Private _SoNgayLViec_Thang As Integer
    Private _MucTamUng_V2 As Double
    Private _HeSoLuong_V2 As Double
    Public Sub New()
        _IdOutPut = 0
        _TongHopId = 0
        _DonVi_CL_Cd = ""
        _PhongBan_CL_Cd = ""
        _NamBC = 0
        _ThangBC = 0
        _KyBC = 0
        _PhanLoai_Cd = ""
        _LaoDong_Cd = ""
        _NgayBC = Globals.GetDateTime_ForServerDB
        _GhiChu_CL = ""
        _TrangThai = 0
        _TrangThai_HT = ""
        _CreatedBy = ""
        _CreatedDate = Globals.GetDateTime_ForServerDB
        _ModifiedBy = ""
        _ModifiedDate = Globals.GetDateTime_ForServerDB

        _DonVi_CL_HT = ""
        _PhongBan_CL_HT = ""
        _PhanLoai_HT = ""
        _LaoDong_HT = ""
        _ThongTin_HT = ""

        _Luong_TTV_Ma = ""
        _Luong_TTV = 0
        _Luong_CSo = 0
        _BHXH_CB = 0
        _BHXH_DV = 0
        _BHYT_CB = 0
        _BHYT_DV = 0
        _BHTN_CB = 0
        _BHTN_DV = 0
        _DPCD_CB = 0
        _DPCD_LD_NN = 0
        _Tinh_Thue_TNCN = 0
        _TNTT_Muc_BanThan = 0
        _TNTT_Muc_PhuThuoc = 0
        _SoNgayLViec_Thang = 0
        _MucTamUng_V2 = 0
        _HeSoLuong_V2 = 0
    End Sub
    Public Property IdOutPut() As Long
        Get
            Return _IdOutPut
        End Get
        Set(ByVal value As Long)
            _IdOutPut = value
        End Set
    End Property
    Public Property TongHopId() As Long
        Get
            Return _TongHopId
        End Get
        Set(ByVal value As Long)
            _TongHopId = value
        End Set
    End Property
    Public Property DonVi_CL_Cd() As String
        Get
            Return _DonVi_CL_Cd
        End Get
        Set(ByVal value As String)
            _DonVi_CL_Cd = value
        End Set
    End Property
    Public Property PhongBan_CL_Cd() As String
        Get
            Return _PhongBan_CL_Cd
        End Get
        Set(ByVal value As String)
            _PhongBan_CL_Cd = value
        End Set
    End Property

    Public Property NamBC() As Integer
        Get
            Return _NamBC
        End Get
        Set(ByVal value As Integer)
            _NamBC = value
        End Set
    End Property
    Public Property ThangBC() As Integer
        Get
            Return _ThangBC
        End Get
        Set(ByVal value As Integer)
            _ThangBC = value
        End Set
    End Property
    Public Property KyBC() As Integer
        Get
            Return _KyBC
        End Get
        Set(ByVal value As Integer)
            _KyBC = value
        End Set
    End Property
    Public Property PhanLoai_Cd() As String
        Get
            Return _PhanLoai_Cd
        End Get
        Set(ByVal value As String)
            _PhanLoai_Cd = value
        End Set
    End Property
    Public Property LaoDong_Cd() As String
        Get
            Return _LaoDong_Cd
        End Get
        Set(ByVal value As String)
            _LaoDong_Cd = value
        End Set
    End Property
    Public Property NgayBC() As DateTime
        Get
            Return _NgayBC
        End Get
        Set(ByVal value As DateTime)
            _NgayBC = value
        End Set
    End Property
    Public Property GhiChu_CL() As String
        Get
            Return _GhiChu_CL
        End Get
        Set(ByVal value As String)
            _GhiChu_CL = value
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
    Public Property TrangThai_HT() As String
        Get
            Return _TrangThai_HT
        End Get
        Set(ByVal value As String)
            _TrangThai_HT = value
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

    Public Property DonVi_CL_HT() As String
        Get
            Return _DonVi_CL_HT
        End Get
        Set(ByVal value As String)
            _DonVi_CL_HT = value
        End Set
    End Property
    Public Property PhongBan_CL_HT() As String
        Get
            Return _PhongBan_CL_HT
        End Get
        Set(ByVal value As String)
            _PhongBan_CL_HT = value
        End Set
    End Property
    Public Property PhanLoai_HT() As String
        Get
            Return _PhanLoai_HT
        End Get
        Set(ByVal value As String)
            _PhanLoai_HT = value
        End Set
    End Property
    Public Property LaoDong_HT() As String
        Get
            Return _LaoDong_HT
        End Get
        Set(ByVal value As String)
            _LaoDong_HT = value
        End Set
    End Property
    Public Property ThongTin_HT() As String
        Get
            Return _ThongTin_HT
        End Get
        Set(ByVal value As String)
            _ThongTin_HT = value
        End Set
    End Property

    Public Property Luong_TTV_Ma() As String
        Get
            Return _Luong_TTV_Ma
        End Get
        Set(ByVal value As String)
            _Luong_TTV_Ma = value
        End Set
    End Property
    Public Property Luong_TTV() As Double
        Get
            Return _Luong_TTV
        End Get
        Set(ByVal value As Double)
            _Luong_TTV = value
        End Set
    End Property
    Public Property Luong_CSo() As Double
        Get
            Return _Luong_CSo
        End Get
        Set(ByVal value As Double)
            _Luong_CSo = value
        End Set
    End Property
    Public Property BHXH_CB() As Double
        Get
            Return _BHXH_CB
        End Get
        Set(ByVal value As Double)
            _BHXH_CB = value
        End Set
    End Property
    Public Property BHXH_DV() As Double
        Get
            Return _BHXH_DV
        End Get
        Set(ByVal value As Double)
            _BHXH_DV = value
        End Set
    End Property
    Public Property BHYT_CB() As Double
        Get
            Return _BHYT_CB
        End Get
        Set(ByVal value As Double)
            _BHYT_CB = value
        End Set
    End Property
    Public Property BHYT_DV() As Double
        Get
            Return _BHYT_DV
        End Get
        Set(ByVal value As Double)
            _BHYT_DV = value
        End Set
    End Property
    Public Property BHTN_CB() As Double
        Get
            Return _BHTN_CB
        End Get
        Set(ByVal value As Double)
            _BHTN_CB = value
        End Set
    End Property
    Public Property BHTN_DV() As Double
        Get
            Return _BHTN_DV
        End Get
        Set(ByVal value As Double)
            _BHTN_DV = value
        End Set
    End Property
    Public Property DPCD_CB() As Double
        Get
            Return _DPCD_CB
        End Get
        Set(ByVal value As Double)
            _DPCD_CB = value
        End Set
    End Property
    Public Property DPCD_LD_NN() As Double
        Get
            Return _DPCD_LD_NN
        End Get
        Set(ByVal value As Double)
            _DPCD_LD_NN = value
        End Set
    End Property
    Public Property Tinh_Thue_TNCN() As Byte
        Get
            Return _Tinh_Thue_TNCN
        End Get
        Set(ByVal value As Byte)
            _Tinh_Thue_TNCN = value
        End Set
    End Property
    Public Property TNTT_Muc_BanThan() As Double
        Get
            Return _TNTT_Muc_BanThan
        End Get
        Set(ByVal value As Double)
            _TNTT_Muc_BanThan = value
        End Set
    End Property
    Public Property TNTT_Muc_PhuThuoc() As Double
        Get
            Return _TNTT_Muc_PhuThuoc
        End Get
        Set(ByVal value As Double)
            _TNTT_Muc_PhuThuoc = value
        End Set
    End Property
    Public Property SoNgayLViec_Thang() As Integer
        Get
            Return _SoNgayLViec_Thang
        End Get
        Set(ByVal value As Integer)
            _SoNgayLViec_Thang = value
        End Set
    End Property
    Public Property MucTamUng_V2() As Double
        Get
            Return _MucTamUng_V2
        End Get
        Set(ByVal value As Double)
            _MucTamUng_V2 = value
        End Set
    End Property
    Public Property HeSoLuong_V2() As Double
        Get
            Return _HeSoLuong_V2
        End Get
        Set(ByVal value As Double)
            _HeSoLuong_V2 = value
        End Set
    End Property
End Class

Public Class ChiLuong_BangKeCT
    Private _IdOutPut As Long
    Private _KieuIn As Byte
    Private _OrderNo As Integer
    Private _STT As String
    Private _IdCanBo As String
    Private _MaCB As String
    Private _HoTen As String
    Private _ThongTin_HT As String
    Private _NgaySinh As String
    Private _GioiTinh As String
    Private _SoTK_NHCS As String
    Private _ChiNhanh_Cd As String
    Private _ChiNhanh_HT As String
    Private _DonVi_Cd As String
    Private _DonVi_HT As String
    Private _PhongBan_Cd As String
    Private _PhongBan_HT As String
    Private _ChucVu_Cd As String
    Private _ChucVu_HT As String
    Private _ChuyenMon_Cd As String
    Private _ChuyenMon_HT As String
    Private _LCB_BacLuong_Id As Integer
    Private _LCB_BacLuong_HT As String
    Private _LCB_HeSo As Double
    Private _LCB_NgHuong As DateTime
    Private _PhCap_ChVu_HeSo As Double
    Private _PhCap_ChVu_NgHuong As DateTime
    Private _PhCap_TrNhiem_HeSo As Double
    Private _PhCap_TrNhiem_NgHuong As DateTime
    Private _PhCap_DocHai_HeSo As Double
    Private _PhCap_DocHai_NgHuong As DateTime
    Private _PhCap_ThuHut_HeSo As Double
    Private _PhCap_ThuHut_SoTien As Double
    Private _PhCap_ThuHut_NgHuong As DateTime
    Private _PhCap_KhuVuc_HeSo As Double
    Private _PhCap_KhuVuc_SoTien As Double
    Private _PhCap_KhuVuc_NgHuong As DateTime
    Private _SoNgayNghi_TruLuong As Integer
    Private _SoTienNghi_TruLuong As Double
    Private _Luong_TTV_Ma As String
    Private _Luong_TTV As Double
    Private _Luong_CSo As Double
    Private _Luong_V1 As Double
    Private _Luong_V1_TamUng_K1 As Double
    Private _Luong_V1_TamUng_K2 As Double
    Private _BHXH_CB As Double
    Private _BHXH_CB_SoTien As Double
    Private _BHXH_DV As Double
    Private _BHXH_DV_SoTien As Double
    Private _BHYT_CB As Double
    Private _BHYT_CB_SoTien As Double
    Private _BHYT_DV As Double
    Private _BHYT_DV_SoTien As Double
    Private _BHTN_CB As Double
    Private _BHTN_CB_SoTien As Double
    Private _BHTN_DV As Double
    Private _BHTN_DV_SoTien As Double
    Private _DPCD_CB As Double
    Private _DPCD_CB_SoTien As Double
    Private _DPCD_LD_NN As Double
    Private _DPCD_LD_NN_SoTien As Double
    Private _Tru_Khoan_Khac As Double
    Private _Tru_Khoan_Khac_GhiChu As String
    Private _Tru_UngHo_Khac As Double
    Private _Tru_UngHo_Khac_GhiChu As String
    Private _TNTT_Muc_BanThan As Double
    Private _TNTT_Muc_PhuThuoc As Double
    Private _TNTT_SoTien As Double
    Private _TNTT_TyLe As Double
    Private _TNTT_SoTien_NopThue As Double
    Private _Luong_V1_ThucLinh_K1 As Double
    Private _GhiChu_ChiLuong As String

    Private _Tru_CacKhoan_Tong As Double
    Private _Tru_CacKhoan_GhiChu As String
    Private _SoNguoi_PhuThuoc As Integer

    Private _Luong_V1_ConLai As Double
    Private _Luong_V1_ThucTra As Double
    Private _MucTamUng_V2 As Double
    Private _Luong_V2_TamUng As Double

    Private _Luong_ThucLinh As Double
    Private _SoNgayLViec_Thang As Integer
    Private _TrangThai As Byte
    Private _TrangThai_HT As String
    Private _CreatedBy As String
    Private _CreatedDate As DateTime
    Private _ModifiedBy As String
    Private _ModifiedDate As DateTime
    Private _ChiTietId As Long
    Private _TongHopId As Long

    Private _NamBC As Integer
    Private _ThangBC As Integer
    Private _KyBC As Integer
    Private _PhanLoai_Cd As String
    Private _PhanLoai_HT As String
    Private _LaoDong_Cd As String
    Private _LaoDong_HT As String
    Private _NgayBC As DateTime
    Private _IsBHXH As Byte
    Private _IsBHYT As Byte
    Private _IsBHTN As Byte
    Private _IsDPCD As Byte
    Private _DonGia_LamDem_Gio As Double
    Private _LamDem_SoGio As Double
    Private _LamDem_SoTienPC As Double
    Private _Loai_HDNH As Integer

    Private _LoaiQD_Cd As String
    Private _LoaiQD_HT As String
    Private _LoaiHDLD_Cd As String
    Private _LoaiHDLD_HT As String
    Private _Loai_CB As Byte
    Private _Loai_CB_HT As String
    Private _SoNgay_KhongLV As Integer
    Private _MucHuong As Double
    Private _PhCap_ThuHut_ST_Goc As Double
    Private _PhCap_KhuVuc_ST_Goc As Double

    Public Sub New()
        _IdOutPut = 0
        _KieuIn = 0
        _OrderNo = 0
        _STT = ""
        _IdCanBo = ""
        _MaCB = ""
        _HoTen = ""
        _ThongTin_HT = ""
        _NgaySinh = ""
        _GioiTinh = ""
        _SoTK_NHCS = ""
        _ChiNhanh_Cd = ""
        _ChiNhanh_HT = ""
        _DonVi_Cd = ""
        _DonVi_HT = ""
        _PhongBan_Cd = ""
        _PhongBan_HT = ""
        _ChucVu_Cd = ""
        _ChucVu_HT = ""
        _ChuyenMon_Cd = ""
        _ChuyenMon_HT = ""
        _LCB_BacLuong_Id = 0
        _LCB_BacLuong_HT = ""
        _LCB_HeSo = 0
        _LCB_NgHuong = Globals.GetDateTime_ForServerDB
        _PhCap_ChVu_HeSo = 0
        _PhCap_ChVu_NgHuong = Globals.GetDateTime_ForServerDB
        _PhCap_TrNhiem_HeSo = 0
        _PhCap_TrNhiem_NgHuong = Globals.GetDateTime_ForServerDB
        _PhCap_DocHai_HeSo = 0
        _PhCap_DocHai_NgHuong = Globals.GetDateTime_ForServerDB
        _PhCap_ThuHut_HeSo = 0
        _PhCap_ThuHut_SoTien = 0
        _PhCap_ThuHut_NgHuong = Globals.GetDateTime_ForServerDB
        _PhCap_KhuVuc_HeSo = 0
        _PhCap_KhuVuc_SoTien = 0
        _PhCap_KhuVuc_NgHuong = Globals.GetDateTime_ForServerDB
        _SoNgayNghi_TruLuong = 0
        _SoTienNghi_TruLuong = 0
        _Luong_TTV_Ma = ""
        _Luong_TTV = 0
        _Luong_CSo = 0
        _Luong_V1 = 0
        _Luong_V1_TamUng_K1 = 0
        _Luong_V1_TamUng_K2 = 0
        _BHXH_CB = 0
        _BHXH_CB_SoTien = 0
        _BHXH_DV = 0
        _BHXH_DV_SoTien = 0
        _BHYT_CB = 0
        _BHYT_CB_SoTien = 0
        _BHYT_DV = 0
        _BHYT_DV_SoTien = 0
        _BHTN_CB = 0
        _BHTN_CB_SoTien = 0
        _BHTN_DV = 0
        _BHTN_DV_SoTien = 0
        _DPCD_CB = 0
        _DPCD_CB_SoTien = 0
        _DPCD_LD_NN = 0
        _DPCD_LD_NN_SoTien = 0
        _Tru_Khoan_Khac = 0
        _Tru_Khoan_Khac_GhiChu = ""
        _Tru_UngHo_Khac = 0
        _Tru_UngHo_Khac_GhiChu = ""
        _TNTT_Muc_BanThan = 0
        _TNTT_Muc_PhuThuoc = 0
        _TNTT_SoTien = 0
        _TNTT_TyLe = 0
        _TNTT_SoTien_NopThue = 0
        _Luong_V1_ThucLinh_K1 = 0
        _GhiChu_ChiLuong = ""

        _NamBC = 0
        _ThangBC = 0
        _KyBC = 0
        _PhanLoai_Cd = ""
        _PhanLoai_HT = ""
        _LaoDong_Cd = ""
        _LaoDong_HT = ""
        _NgayBC = Globals.GetDateTime_ForServerDB

        _TrangThai = 0
        _TrangThai_HT = ""
        _CreatedBy = ""
        _CreatedDate = Globals.GetDateTime_ForServerDB
        _ModifiedBy = ""
        _ModifiedDate = Globals.GetDateTime_ForServerDB
        _ChiTietId = 0
        _TongHopId = 0

        _Tru_CacKhoan_Tong = 0
        _Tru_CacKhoan_GhiChu = ""
        _SoNguoi_PhuThuoc = 0
        _Luong_ThucLinh = 0
        _SoNgayLViec_Thang = 0
        _Luong_V1_ConLai = 0
        _Luong_V1_ThucTra = 0
        _MucTamUng_V2 = 0
        _Luong_V2_TamUng = 0
        _IsBHXH = 0
        _IsBHYT = 0
        _IsBHTN = 0
        _IsDPCD = 0
        _DonGia_LamDem_Gio = 0
        _LamDem_SoGio = 0
        _LamDem_SoTienPC = 0
        _Loai_HDNH = 0

        _LoaiQD_Cd = ""
        _LoaiQD_HT = ""
        _LoaiHDLD_Cd = ""
        _LoaiHDLD_HT = ""
        _Loai_CB = 0
        _Loai_CB_HT = ""
        _SoNgay_KhongLV = 0
        _MucHuong = 0
        _PhCap_KhuVuc_ST_Goc = 0
        _PhCap_ThuHut_ST_Goc = 0
    End Sub
    Public Property IdOutPut() As Long
        Get
            Return _IdOutPut
        End Get
        Set(ByVal value As Long)
            _IdOutPut = value
        End Set
    End Property
    Public Property KieuIn() As Byte
        Get
            Return _KieuIn
        End Get
        Set(ByVal value As Byte)
            _KieuIn = value
        End Set
    End Property
    Public Property OrderNo() As Integer
        Get
            Return _OrderNo
        End Get
        Set(ByVal value As Integer)
            _OrderNo = value
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
    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal value As String)
            _IdCanBo = value
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
    Public Property ThongTin_HT() As String
        Get
            Return _ThongTin_HT
        End Get
        Set(ByVal value As String)
            _ThongTin_HT = value
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
    Public Property SoTK_NHCS() As String
        Get
            Return _SoTK_NHCS
        End Get
        Set(ByVal value As String)
            _SoTK_NHCS = value
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
    Public Property ChuyenMon_Cd() As String
        Get
            Return _ChuyenMon_Cd
        End Get
        Set(ByVal value As String)
            _ChuyenMon_Cd = value
        End Set
    End Property
    Public Property ChuyenMon_HT() As String
        Get
            Return _ChuyenMon_HT
        End Get
        Set(ByVal value As String)
            _ChuyenMon_HT = value
        End Set
    End Property
    Public Property LCB_BacLuong_Id() As Integer
        Get
            Return _LCB_BacLuong_Id
        End Get
        Set(ByVal value As Integer)
            _LCB_BacLuong_Id = value
        End Set
    End Property
    Public Property LCB_BacLuong_HT() As String
        Get
            Return _LCB_BacLuong_HT
        End Get
        Set(ByVal value As String)
            _LCB_BacLuong_HT = value
        End Set
    End Property
    Public Property LCB_HeSo() As Double
        Get
            Return _LCB_HeSo
        End Get
        Set(ByVal value As Double)
            _LCB_HeSo = value
        End Set
    End Property
    Public Property LCB_NgHuong() As DateTime
        Get
            Return _LCB_NgHuong
        End Get
        Set(ByVal value As DateTime)
            _LCB_NgHuong = value
        End Set
    End Property
    Public Property PhCap_ChVu_HeSo() As Double
        Get
            Return _PhCap_ChVu_HeSo
        End Get
        Set(ByVal value As Double)
            _PhCap_ChVu_HeSo = value
        End Set
    End Property
    Public Property PhCap_ChVu_NgHuong() As DateTime
        Get
            Return _PhCap_ChVu_NgHuong
        End Get
        Set(ByVal value As DateTime)
            _PhCap_ChVu_NgHuong = value
        End Set
    End Property
    Public Property PhCap_TrNhiem_HeSo() As Double
        Get
            Return _PhCap_TrNhiem_HeSo
        End Get
        Set(ByVal value As Double)
            _PhCap_TrNhiem_HeSo = value
        End Set
    End Property
    Public Property PhCap_TrNhiem_NgHuong() As DateTime
        Get
            Return _PhCap_TrNhiem_NgHuong
        End Get
        Set(ByVal value As DateTime)
            _PhCap_TrNhiem_NgHuong = value
        End Set
    End Property
    Public Property PhCap_DocHai_HeSo() As Double
        Get
            Return _PhCap_DocHai_HeSo
        End Get
        Set(ByVal value As Double)
            _PhCap_DocHai_HeSo = value
        End Set
    End Property
    Public Property PhCap_DocHai_NgHuong() As DateTime
        Get
            Return _PhCap_DocHai_NgHuong
        End Get
        Set(ByVal value As DateTime)
            _PhCap_DocHai_NgHuong = value
        End Set
    End Property
    Public Property PhCap_ThuHut_HeSo() As Double
        Get
            Return _PhCap_ThuHut_HeSo
        End Get
        Set(ByVal value As Double)
            _PhCap_ThuHut_HeSo = value
        End Set
    End Property
    Public Property PhCap_ThuHut_SoTien() As Double
        Get
            Return _PhCap_ThuHut_SoTien
        End Get
        Set(ByVal value As Double)
            _PhCap_ThuHut_SoTien = value
        End Set
    End Property
    Public Property PhCap_ThuHut_NgHuong() As DateTime
        Get
            Return _PhCap_ThuHut_NgHuong
        End Get
        Set(ByVal value As DateTime)
            _PhCap_ThuHut_NgHuong = value
        End Set
    End Property
    Public Property PhCap_KhuVuc_HeSo() As Double
        Get
            Return _PhCap_KhuVuc_HeSo
        End Get
        Set(ByVal value As Double)
            _PhCap_KhuVuc_HeSo = value
        End Set
    End Property
    Public Property PhCap_KhuVuc_SoTien() As Double
        Get
            Return _PhCap_KhuVuc_SoTien
        End Get
        Set(ByVal value As Double)
            _PhCap_KhuVuc_SoTien = value
        End Set
    End Property
    Public Property PhCap_KhuVuc_NgHuong() As DateTime
        Get
            Return _PhCap_KhuVuc_NgHuong
        End Get
        Set(ByVal value As DateTime)
            _PhCap_KhuVuc_NgHuong = value
        End Set
    End Property
    Public Property SoNgayNghi_TruLuong() As Integer
        Get
            Return _SoNgayNghi_TruLuong
        End Get
        Set(ByVal value As Integer)
            _SoNgayNghi_TruLuong = value
        End Set
    End Property
    Public Property SoTienNghi_TruLuong() As Double
        Get
            Return _SoTienNghi_TruLuong
        End Get
        Set(ByVal value As Double)
            _SoTienNghi_TruLuong = value
        End Set
    End Property
    Public Property Luong_TTV_Ma() As String
        Get
            Return _Luong_TTV_Ma
        End Get
        Set(ByVal value As String)
            _Luong_TTV_Ma = value
        End Set
    End Property
    Public Property Luong_TTV() As Double
        Get
            Return _Luong_TTV
        End Get
        Set(ByVal value As Double)
            _Luong_TTV = value
        End Set
    End Property
    Public Property Luong_CSo() As Double
        Get
            Return _Luong_CSo
        End Get
        Set(ByVal value As Double)
            _Luong_CSo = value
        End Set
    End Property
    Public Property Luong_V1() As Double
        Get
            Return _Luong_V1
        End Get
        Set(ByVal value As Double)
            _Luong_V1 = value
        End Set
    End Property
    Public Property Luong_V1_TamUng_K1() As Double
        Get
            Return _Luong_V1_TamUng_K1
        End Get
        Set(ByVal value As Double)
            _Luong_V1_TamUng_K1 = value
        End Set
    End Property
    Public Property Luong_V1_TamUng_K2() As Double
        Get
            Return _Luong_V1_TamUng_K2
        End Get
        Set(ByVal value As Double)
            _Luong_V1_TamUng_K2 = value
        End Set
    End Property
    Public Property BHXH_CB() As Double
        Get
            Return _BHXH_CB
        End Get
        Set(ByVal value As Double)
            _BHXH_CB = value
        End Set
    End Property
    Public Property BHXH_CB_SoTien() As Double
        Get
            Return _BHXH_CB_SoTien
        End Get
        Set(ByVal value As Double)
            _BHXH_CB_SoTien = value
        End Set
    End Property
    Public Property BHXH_DV() As Double
        Get
            Return _BHXH_DV
        End Get
        Set(ByVal value As Double)
            _BHXH_DV = value
        End Set
    End Property
    Public Property BHXH_DV_SoTien() As Double
        Get
            Return _BHXH_DV_SoTien
        End Get
        Set(ByVal value As Double)
            _BHXH_DV_SoTien = value
        End Set
    End Property
    Public Property BHYT_CB() As Double
        Get
            Return _BHYT_CB
        End Get
        Set(ByVal value As Double)
            _BHYT_CB = value
        End Set
    End Property
    Public Property BHYT_CB_SoTien() As Double
        Get
            Return _BHYT_CB_SoTien
        End Get
        Set(ByVal value As Double)
            _BHYT_CB_SoTien = value
        End Set
    End Property
    Public Property BHYT_DV() As Double
        Get
            Return _BHYT_DV
        End Get
        Set(ByVal value As Double)
            _BHYT_DV = value
        End Set
    End Property
    Public Property BHYT_DV_SoTien() As Double
        Get
            Return _BHYT_DV_SoTien
        End Get
        Set(ByVal value As Double)
            _BHYT_DV_SoTien = value
        End Set
    End Property
    Public Property BHTN_CB() As Double
        Get
            Return _BHTN_CB
        End Get
        Set(ByVal value As Double)
            _BHTN_CB = value
        End Set
    End Property
    Public Property BHTN_CB_SoTien() As Double
        Get
            Return _BHTN_CB_SoTien
        End Get
        Set(ByVal value As Double)
            _BHTN_CB_SoTien = value
        End Set
    End Property
    Public Property BHTN_DV() As Double
        Get
            Return _BHTN_DV
        End Get
        Set(ByVal value As Double)
            _BHTN_DV = value
        End Set
    End Property
    Public Property BHTN_DV_SoTien() As Double
        Get
            Return _BHTN_DV_SoTien
        End Get
        Set(ByVal value As Double)
            _BHTN_DV_SoTien = value
        End Set
    End Property
    Public Property DPCD_CB() As Double
        Get
            Return _DPCD_CB
        End Get
        Set(ByVal value As Double)
            _DPCD_CB = value
        End Set
    End Property
    Public Property DPCD_CB_SoTien() As Double
        Get
            Return _DPCD_CB_SoTien
        End Get
        Set(ByVal value As Double)
            _DPCD_CB_SoTien = value
        End Set
    End Property
    Public Property DPCD_LD_NN() As Double
        Get
            Return _DPCD_LD_NN
        End Get
        Set(ByVal value As Double)
            _DPCD_LD_NN = value
        End Set
    End Property
    Public Property DPCD_LD_NN_SoTien() As Double
        Get
            Return _DPCD_LD_NN_SoTien
        End Get
        Set(ByVal value As Double)
            _DPCD_LD_NN_SoTien = value
        End Set
    End Property
    Public Property Tru_Khoan_Khac() As Double
        Get
            Return _Tru_Khoan_Khac
        End Get
        Set(ByVal value As Double)
            _Tru_Khoan_Khac = value
        End Set
    End Property
    Public Property Tru_Khoan_Khac_GhiChu() As String
        Get
            Return _Tru_Khoan_Khac_GhiChu
        End Get
        Set(ByVal value As String)
            _Tru_Khoan_Khac_GhiChu = value
        End Set
    End Property
    Public Property Tru_UngHo_Khac() As Double
        Get
            Return _Tru_UngHo_Khac
        End Get
        Set(ByVal value As Double)
            _Tru_UngHo_Khac = value
        End Set
    End Property
    Public Property Tru_UngHo_Khac_GhiChu() As String
        Get
            Return _Tru_UngHo_Khac_GhiChu
        End Get
        Set(ByVal value As String)
            _Tru_UngHo_Khac_GhiChu = value
        End Set
    End Property
    Public Property TNTT_Muc_BanThan() As Double
        Get
            Return _TNTT_Muc_BanThan
        End Get
        Set(ByVal value As Double)
            _TNTT_Muc_BanThan = value
        End Set
    End Property
    Public Property TNTT_Muc_PhuThuoc() As Double
        Get
            Return _TNTT_Muc_PhuThuoc
        End Get
        Set(ByVal value As Double)
            _TNTT_Muc_PhuThuoc = value
        End Set
    End Property
    Public Property TNTT_SoTien() As Double
        Get
            Return _TNTT_SoTien
        End Get
        Set(ByVal value As Double)
            _TNTT_SoTien = value
        End Set
    End Property
    Public Property TNTT_TyLe() As Double
        Get
            Return _TNTT_TyLe
        End Get
        Set(ByVal value As Double)
            _TNTT_TyLe = value
        End Set
    End Property
    Public Property TNTT_SoTien_NopThue() As Double
        Get
            Return _TNTT_SoTien_NopThue
        End Get
        Set(ByVal value As Double)
            _TNTT_SoTien_NopThue = value
        End Set
    End Property
    Public Property Luong_V1_ThucLinh_K1() As Double
        Get
            Return _Luong_V1_ThucLinh_K1
        End Get
        Set(ByVal value As Double)
            _Luong_V1_ThucLinh_K1 = value
        End Set
    End Property
    Public Property GhiChu_ChiLuong() As String
        Get
            Return _GhiChu_ChiLuong
        End Get
        Set(ByVal value As String)
            _GhiChu_ChiLuong = value
        End Set
    End Property
    Public Property NamBC() As Integer
        Get
            Return _NamBC
        End Get
        Set(ByVal value As Integer)
            _NamBC = value
        End Set
    End Property
    Public Property ThangBC() As Integer
        Get
            Return _ThangBC
        End Get
        Set(ByVal value As Integer)
            _ThangBC = value
        End Set
    End Property
    Public Property KyBC() As Integer
        Get
            Return _KyBC
        End Get
        Set(ByVal value As Integer)
            _KyBC = value
        End Set
    End Property
    Public Property PhanLoai_Cd() As String
        Get
            Return _PhanLoai_Cd
        End Get
        Set(ByVal value As String)
            _PhanLoai_Cd = value
        End Set
    End Property
    Public Property PhanLoai_HT() As String
        Get
            Return _PhanLoai_HT
        End Get
        Set(ByVal value As String)
            _PhanLoai_HT = value
        End Set
    End Property
    Public Property LaoDong_Cd() As String
        Get
            Return _LaoDong_Cd
        End Get
        Set(ByVal value As String)
            _LaoDong_Cd = value
        End Set
    End Property
    Public Property LaoDong_HT() As String
        Get
            Return _LaoDong_HT
        End Get
        Set(ByVal value As String)
            _LaoDong_HT = value
        End Set
    End Property
    Public Property NgayBC() As DateTime
        Get
            Return _NgayBC
        End Get
        Set(ByVal value As DateTime)
            _NgayBC = value
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
    Public Property TrangThai_HT() As String
        Get
            Return _TrangThai_HT
        End Get
        Set(ByVal value As String)
            _TrangThai_HT = value
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
    Public Property ChiTietId() As Long
        Get
            Return _ChiTietId
        End Get
        Set(ByVal value As Long)
            _ChiTietId = value
        End Set
    End Property
    Public Property TongHopId() As Long
        Get
            Return _TongHopId
        End Get
        Set(ByVal value As Long)
            _TongHopId = value
        End Set
    End Property
    Public Property Tru_CacKhoan_Tong() As Double
        Get
            Return _Tru_CacKhoan_Tong
        End Get
        Set(ByVal value As Double)
            _Tru_CacKhoan_Tong = value
        End Set
    End Property
    Public Property Tru_CacKhoan_GhiChu() As String
        Get
            Return _Tru_CacKhoan_GhiChu
        End Get
        Set(ByVal value As String)
            _Tru_CacKhoan_GhiChu = value
        End Set
    End Property
    Public Property SoNguoi_PhuThuoc() As Integer
        Get
            Return _SoNguoi_PhuThuoc
        End Get
        Set(ByVal value As Integer)
            _SoNguoi_PhuThuoc = value
        End Set
    End Property
    Public Property Luong_ThucLinh() As Double
        Get
            Return _Luong_ThucLinh
        End Get
        Set(ByVal value As Double)
            _Luong_ThucLinh = value
        End Set
    End Property
    Public Property SoNgayLViec_Thang() As Integer
        Get
            Return _SoNgayLViec_Thang
        End Get
        Set(ByVal value As Integer)
            _SoNgayLViec_Thang = value
        End Set
    End Property
    Public Property Luong_V1_ConLai() As Double
        Get
            Return _Luong_V1_ConLai
        End Get
        Set(ByVal value As Double)
            _Luong_V1_ConLai = value
        End Set
    End Property
    Public Property Luong_V1_ThucTra() As Double
        Get
            Return _Luong_V1_ThucTra
        End Get
        Set(ByVal value As Double)
            _Luong_V1_ThucTra = value
        End Set
    End Property
    Public Property MucTamUng_V2() As Double
        Get
            Return _MucTamUng_V2
        End Get
        Set(ByVal value As Double)
            _MucTamUng_V2 = value
        End Set
    End Property
    Public Property Luong_V2_TamUng() As Double
        Get
            Return _Luong_V2_TamUng
        End Get
        Set(ByVal value As Double)
            _Luong_V2_TamUng = value
        End Set
    End Property
    Public Property IsBHXH() As Byte
        Get
            Return _IsBHXH
        End Get
        Set(ByVal value As Byte)
            _IsBHXH = value
        End Set
    End Property
    Public Property IsBHYT() As Byte
        Get
            Return _IsBHYT
        End Get
        Set(ByVal value As Byte)
            _IsBHYT = value
        End Set
    End Property
    Public Property IsBHTN() As Byte
        Get
            Return _IsBHTN
        End Get
        Set(ByVal value As Byte)
            _IsBHTN = value
        End Set
    End Property
    Public Property IsDPCD() As Byte
        Get
            Return _IsDPCD
        End Get
        Set(ByVal value As Byte)
            _IsDPCD = value
        End Set
    End Property
    Public Property DonGia_LamDem_Gio() As Double
        Get
            Return _DonGia_LamDem_Gio
        End Get
        Set(ByVal value As Double)
            _DonGia_LamDem_Gio = value
        End Set
    End Property
    Public Property LamDem_SoGio() As Double
        Get
            Return _LamDem_SoGio
        End Get
        Set(ByVal value As Double)
            _LamDem_SoGio = value
        End Set
    End Property
    Public Property LamDem_SoTienPC() As Double
        Get
            Return _LamDem_SoTienPC
        End Get
        Set(ByVal value As Double)
            _LamDem_SoTienPC = value
        End Set
    End Property
    Public Property Loai_HDNH() As Integer
        Get
            Return _Loai_HDNH
        End Get
        Set(ByVal value As Integer)
            _Loai_HDNH = value
        End Set
    End Property

    Public Property LoaiQD_Cd() As String
        Get
            Return _LoaiQD_Cd
        End Get
        Set(ByVal value As String)
            _LoaiQD_Cd = value
        End Set
    End Property
    Public Property LoaiQD_HT() As String
        Get
            Return _LoaiQD_HT
        End Get
        Set(ByVal value As String)
            _LoaiQD_HT = value
        End Set
    End Property
    Public Property LoaiHDLD_Cd() As String
        Get
            Return _LoaiHDLD_Cd
        End Get
        Set(ByVal value As String)
            _LoaiHDLD_Cd = value
        End Set
    End Property
    Public Property LoaiHDLD_HT() As String
        Get
            Return _LoaiHDLD_HT
        End Get
        Set(ByVal value As String)
            _LoaiHDLD_HT = value
        End Set
    End Property
    Public Property Loai_CB() As Byte
        Get
            Return _Loai_CB
        End Get
        Set(ByVal value As Byte)
            _Loai_CB = value
        End Set
    End Property
    Public Property Loai_CB_HT() As String
        Get
            Return _Loai_CB_HT
        End Get
        Set(ByVal value As String)
            _Loai_CB_HT = value
        End Set
    End Property
    Public Property SoNgay_KhongLV() As Integer
        Get
            Return _SoNgay_KhongLV
        End Get
        Set(ByVal value As Integer)
            _SoNgay_KhongLV = value
        End Set
    End Property
    Public Property MucHuong() As Double
        Get
            Return _MucHuong
        End Get
        Set(ByVal value As Double)
            _MucHuong = value
        End Set
    End Property
    Public Property PhCap_KhuVuc_ST_Goc() As Double
        Get
            Return _PhCap_KhuVuc_ST_Goc
        End Get
        Set(ByVal value As Double)
            _PhCap_KhuVuc_ST_Goc = value
        End Set
    End Property
    Public Property PhCap_ThuHut_ST_Goc() As Double
        Get
            Return _PhCap_ThuHut_ST_Goc
        End Get
        Set(ByVal value As Double)
            _PhCap_ThuHut_ST_Goc = value
        End Set
    End Property
End Class
