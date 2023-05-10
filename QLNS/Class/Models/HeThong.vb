Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Public Class HeThong
    Public Class SysVar
        Private _IdOutPut As Integer
        Private _Id As Integer
        Private _Id_DonVi As Integer
        Private _TrucThuoc As Byte
        Private _DiaBan As String
        Private _ALL As Byte
        Private _Max_Ky As Byte
        Private _Tinh_Thue_TNCN As Byte
        Private _IdCanBo_GiamDoc
        Private _GiamDoc As String
        Private _IdCanBo_PhoGD
        Private _PhoGiamDoc As String
        Private _IdCanBo_HCTC
        Private _TruongHCTC As String
        Private _IdCanBo_KeToan
        Private _KeToanTruong As String
        Private _Ten_VT As String
        Private _DonVi_Cd As String
        Private _VungLuongTT As String
        Private _HeSoTamUng_V2 As Double
        Private _MucTamUng_V2 As Double
        Private _NgayHL As DateTime
        Private _NgayHetHL As DateTime
        Private _TrangThai As Byte
        Private _CreatedBy As String
        Private _CreatedDate As DateTime
        Private _ModifiedBy As String
        Private _ModifiedDate As DateTime

        Private _DonVi_HT As String
        Private _TrangThai_HT As String

        Public Sub New()
            _IdOutPut = 0
            _Id = 0
            _Id_DonVi = 0
            _TrucThuoc = 0
            _DiaBan = ""
            _ALL = 0
            _Max_Ky = 0
            _Tinh_Thue_TNCN = 0
            _IdCanBo_GiamDoc = ""
            _GiamDoc = ""
            _IdCanBo_PhoGD = ""
            _PhoGiamDoc = ""
            _IdCanBo_HCTC = ""
            _TruongHCTC = ""
            _IdCanBo_KeToan = ""
            _KeToanTruong = ""
            _Ten_VT = ""
            _DonVi_Cd = ""
            _VungLuongTT = ""
            _HeSoTamUng_V2 = 0
            _MucTamUng_V2 = 0
            _NgayHL = Globals.GetDateTime_ForServerDB
            _NgayHetHL = Globals.GetDateTime_ForServerDB
            _TrangThai = 0
            _CreatedBy = ""
            _CreatedDate = Globals.GetDateTime_ForServerDB
            _ModifiedBy = ""
            _ModifiedDate = Globals.GetDateTime_ForServerDB

            _DonVi_HT = ""
            _TrangThai_HT = ""
        End Sub

        Public Property IdOutPut() As Integer
            Get
                Return _IdOutPut
            End Get
            Set(ByVal value As Integer)
                _IdOutPut = value
            End Set
        End Property
        Public Property Id() As Integer
            Get
                Return _Id
            End Get
            Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property
        Public Property Id_DonVi() As Integer
            Get
                Return _Id_DonVi
            End Get
            Set(ByVal value As Integer)
                _Id_DonVi = value
            End Set
        End Property
        Public Property TrucThuoc() As Byte
            Get
                Return _TrucThuoc
            End Get
            Set(ByVal value As Byte)
                _TrucThuoc = value
            End Set
        End Property
        Public Property DiaBan() As String
            Get
                Return _DiaBan
            End Get
            Set(ByVal value As String)
                _DiaBan = value
            End Set
        End Property
        Public Property ALL() As Byte
            Get
                Return _ALL
            End Get
            Set(ByVal value As Byte)
                _ALL = value
            End Set
        End Property
        Public Property Max_Ky() As Byte
            Get
                Return _Max_Ky
            End Get
            Set(ByVal value As Byte)
                _Max_Ky = value
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
        Public Property IdCanBo_GiamDoc() As String
            Get
                Return _IdCanBo_GiamDoc
            End Get
            Set(ByVal value As String)
                _IdCanBo_GiamDoc = value
            End Set
        End Property
        Public Property GiamDoc() As String
            Get
                Return _GiamDoc
            End Get
            Set(ByVal value As String)
                _GiamDoc = value
            End Set
        End Property
        Public Property IdCanBo_PhoGD() As String
            Get
                Return _IdCanBo_PhoGD
            End Get
            Set(ByVal value As String)
                _IdCanBo_PhoGD = value
            End Set
        End Property
        Public Property PhoGiamDoc() As String
            Get
                Return _PhoGiamDoc
            End Get
            Set(ByVal value As String)
                _PhoGiamDoc = value
            End Set
        End Property
        Public Property IdCanBo_HCTC() As String
            Get
                Return _IdCanBo_HCTC
            End Get
            Set(ByVal value As String)
                _IdCanBo_HCTC = value
            End Set
        End Property
        Public Property TruongHCTC() As String
            Get
                Return _TruongHCTC
            End Get
            Set(ByVal value As String)
                _TruongHCTC = value
            End Set
        End Property
        Public Property IdCanBo_KeToan() As String
            Get
                Return _IdCanBo_KeToan
            End Get
            Set(ByVal value As String)
                _IdCanBo_KeToan = value
            End Set
        End Property
        Public Property KeToanTruong() As String
            Get
                Return _KeToanTruong
            End Get
            Set(ByVal value As String)
                _KeToanTruong = value
            End Set
        End Property
        Public Property Ten_VT() As String
            Get
                Return _Ten_VT
            End Get
            Set(ByVal value As String)
                _Ten_VT = value
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
        Public Property VungLuongTT() As String
            Get
                Return _VungLuongTT
            End Get
            Set(ByVal value As String)
                _VungLuongTT = value
            End Set
        End Property
        Public Property HeSoTamUng_V2() As Double
            Get
                Return _HeSoTamUng_V2
            End Get
            Set(ByVal value As Double)
                _HeSoTamUng_V2 = value
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
        Public Property NgayHL() As DateTime
            Get
                Return _NgayHL
            End Get
            Set(ByVal value As DateTime)
                _NgayHL = value
            End Set
        End Property
        Public Property NgayHetHL() As DateTime
            Get
                Return _NgayHetHL
            End Get
            Set(ByVal value As DateTime)
                _NgayHetHL = value
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
        Public Property DonVi_HT() As String
            Get
                Return _DonVi_HT
            End Get
            Set(ByVal value As String)
                _DonVi_HT = value
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
    End Class

End Class
