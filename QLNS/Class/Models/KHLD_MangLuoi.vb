Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections

Public Class KHLD_MangLuoi
    Public Class KhLd_MangLuoi
        Private _Id As Integer
        Private _IdKHML As String
        Private _ThoiDiem As DateTime
        Private _Loai_DL As Byte
        Private _IdDonVi As Integer
        Private _DonVi_Cd As String
        Private _IdPhongBan As Integer
        Private _SoLD_DaiHan As Integer
        Private _SoLD_NganHan As Integer
        Private _So_XaPhuong As Integer
        Private _So_DiemGD As Integer
        Private _So_ToTKVV As Integer
        Private _SoKH_DN As Integer
        Private _TongDN As Double
        Private _DuNo_TH As Double
        Private _DuNo_QH As Double
        Private _DuNo_KH As Double
        Private _SoCV_ThBao As String
        Private _GhiChu As String
        Private _CreatedBy As String
        Private _CreatedDate As DateTime
        Private _ModifiedBy As String
        Private _ModifiedDate As DateTime
        Private _IdOutPut As String

        Private _NhuCauBS_SL As Integer
        Private _NhuCauBS_SoCV As String
        Private _ThBaoTTDangDT_SL As Integer
        Private _ThBaoTTDangDT_SoCV As String

        Public Sub New()
            _Id = 0
            _IdKHML = ""
            _ThoiDiem = Globals.GetDateTime_ForServerDB
            _Loai_DL = 0
            _IdDonVi = 0
            _DonVi_Cd = 0
            _IdPhongBan = 0
            _SoLD_DaiHan = 0
            _SoLD_NganHan = 0
            _So_XaPhuong = 0
            _So_DiemGD = 0
            _So_ToTKVV = 0
            _SoKH_DN = 0
            _TongDN = 0
            _DuNo_TH = 0
            _DuNo_QH = 0
            _DuNo_KH = 0
            _SoCV_ThBao = ""
            _GhiChu = ""
            _CreatedBy = ""
            _CreatedDate = Globals.GetDateTime_ForServerDB
            _ModifiedBy = ""
            _ModifiedDate = Globals.GetDateTime_ForServerDB
            _IdOutPut = ""
            _NhuCauBS_SL = 0
            _NhuCauBS_SoCV = ""
            _ThBaoTTDangDT_SL = 0
            _ThBaoTTDangDT_SoCV = ""
        End Sub
        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property
        Public Property IdKHML() As String
            Get
                Return _IdKHML
            End Get
            Set(ByVal value As String)
                _IdKHML = value
            End Set
        End Property
        Public Property ThoiDiem() As DateTime
            Get
                Return _ThoiDiem
            End Get
            Set(ByVal value As DateTime)
                _ThoiDiem = value
            End Set
        End Property
        Public Property Loai_DL() As Byte
            Get
                Return _Loai_DL
            End Get
            Set(ByVal value As Byte)
                _Loai_DL = value
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
        Public Property DonVi_Cd() As String
            Get
                Return _DonVi_Cd
            End Get
            Set(ByVal value As String)
                _DonVi_Cd = value
            End Set
        End Property
        Public Property IdPhongBan() As Integer
            Get
                Return _IdPhongBan
            End Get
            Set(ByVal value As Integer)
                _IdPhongBan = value
            End Set
        End Property
        Public Property SoLD_DaiHan() As Integer
            Get
                Return _SoLD_DaiHan
            End Get
            Set(ByVal value As Integer)
                _SoLD_DaiHan = value
            End Set
        End Property
        Public Property SoLD_NganHan() As Integer
            Get
                Return _SoLD_NganHan
            End Get
            Set(ByVal value As Integer)
                _SoLD_NganHan = value
            End Set
        End Property
        Public Property So_XaPhuong() As Integer
            Get
                Return _So_XaPhuong
            End Get
            Set(ByVal value As Integer)
                _So_XaPhuong = value
            End Set
        End Property
        Public Property So_DiemGD() As Integer
            Get
                Return _So_DiemGD
            End Get
            Set(ByVal value As Integer)
                _So_DiemGD = value
            End Set
        End Property
        Public Property So_ToTKVV() As Integer
            Get
                Return _So_ToTKVV
            End Get
            Set(ByVal value As Integer)
                _So_ToTKVV = value
            End Set
        End Property
        Public Property SoKH_DN() As Integer
            Get
                Return _SoKH_DN
            End Get
            Set(ByVal value As Integer)
                _SoKH_DN = value
            End Set
        End Property
        Public Property TongDN() As Double
            Get
                Return _TongDN
            End Get
            Set(ByVal value As Double)
                _TongDN = value
            End Set
        End Property
        Public Property DuNo_TH() As Double
            Get
                Return _DuNo_TH
            End Get
            Set(ByVal value As Double)
                _DuNo_TH = value
            End Set
        End Property
        Public Property DuNo_QH() As Double
            Get
                Return _DuNo_QH
            End Get
            Set(ByVal value As Double)
                _DuNo_QH = value
            End Set
        End Property
        Public Property DuNo_KH() As Double
            Get
                Return _DuNo_KH
            End Get
            Set(ByVal value As Double)
                _DuNo_KH = value
            End Set
        End Property
        Public Property SoCV_ThBao() As String
            Get
                Return _SoCV_ThBao
            End Get
            Set(ByVal value As String)
                _SoCV_ThBao = value
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
        Public Property IdOutPut() As String
            Get
                Return _IdOutPut
            End Get
            Set(ByVal value As String)
                _IdOutPut = value
            End Set
        End Property
        Public Property NhuCauBS_SL() As Integer
            Get
                Return _NhuCauBS_SL
            End Get
            Set(ByVal value As Integer)
                _NhuCauBS_SL = value
            End Set
        End Property
        Public Property NhuCauBS_SoCV() As String
            Get
                Return _NhuCauBS_SoCV
            End Get
            Set(ByVal value As String)
                _NhuCauBS_SoCV = value
            End Set
        End Property
        Public Property ThBaoTTDangDT_SL() As Integer
            Get
                Return _ThBaoTTDangDT_SL
            End Get
            Set(ByVal value As Integer)
                _ThBaoTTDangDT_SL = value
            End Set
        End Property
        Public Property ThBaoTTDangDT_SoCV() As String
            Get
                Return _ThBaoTTDangDT_SoCV
            End Get
            Set(ByVal value As String)
                _ThBaoTTDangDT_SoCV = value
            End Set
        End Property
    End Class

End Class
