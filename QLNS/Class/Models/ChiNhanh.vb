Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections
Public Class ChiNhanh
    Private _Id As Integer
    Private _Id_Goc As Integer
    Private _Ma_So As String
    Private _Ten_Goi As String
    Private _Ten_VT As String
    Private _Dia_Chi As String
    Private _Dien_Thoai As String
    Private _Email As String
    Private _Website As String
    Private _Status As Byte
    Private _SoLg_XaPhuong As Integer
    Private _MaDV As String
    Private _MaCB_Begin As String
    Private _MaCB_End As String
    Private _Ma_So_Old As String
    Private _IP As String
    Public Sub New()
        _Id = 0
        _Id_Goc = 0
        _Ma_So = ""
        _Ten_Goi = ""
        _Ten_VT = ""
        _Dia_Chi = ""
        _Dien_Thoai = ""
        _Email = ""
        _Website = ""
        _Status = 0
        _SoLg_XaPhuong = 0
        _MaDV = ""
        _MaCB_Begin = ""
        _MaCB_End = ""
        _Ma_So_Old = ""
        _IP = ""
    End Sub

    Public Property Id() As Integer
        Get
            Return _Id
        End Get
        Set(ByVal value As Integer)
            _Id = value
        End Set
    End Property

    Public Property Id_Goc() As Integer
        Get
            Return _Id_Goc
        End Get
        Set(ByVal value As Integer)
            _Id_Goc = value
        End Set
    End Property

    Public Property Ma_So() As String
        Get
            Return _Ma_So
        End Get
        Set(ByVal value As String)
            _Ma_So = value
        End Set
    End Property

    Public Property Ten_Goi() As String
        Get
            Return _Ten_Goi
        End Get
        Set(ByVal value As String)
            _Ten_Goi = value
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
  
    Public Property Dia_Chi() As String
        Get
            Return _Dia_Chi
        End Get
        Set(ByVal value As String)
            _Dia_Chi = value
        End Set
    End Property

    Public Property Dien_Thoai() As String
        Get
            Return _Dien_Thoai
        End Get
        Set(ByVal value As String)
            _Dien_Thoai = value
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

    Public Property Website() As String
        Get
            Return _Website
        End Get
        Set(ByVal value As String)
            _Website = value
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

    Public Property SoLg_XaPhuong() As Integer
        Get
            Return _SoLg_XaPhuong
        End Get
        Set(ByVal value As Integer)
            _SoLg_XaPhuong = value
        End Set
    End Property

    Public Property MaDV() As String
        Get
            Return _MaDV
        End Get
        Set(ByVal value As String)
            _MaDV = value
        End Set
    End Property

    Public Property MaCB_Begin() As String
        Get
            Return _MaCB_Begin
        End Get
        Set(ByVal value As String)
            _MaCB_Begin = value
        End Set
    End Property

    Public Property MaCB_End() As String
        Get
            Return _MaCB_End
        End Get
        Set(ByVal value As String)
            _MaCB_End = value
        End Set
    End Property

    Public Property Ma_So_Old() As String
        Get
            Return _Ma_So_Old
        End Get
        Set(ByVal value As String)
            _Ma_So_Old = value
        End Set
    End Property

    Public Property IP() As String
        Get
            Return _IP
        End Get
        Set(ByVal value As String)
            _IP = value
        End Set
    End Property

    Public Sub GetData(_Reader As System.Data.SqlClient.SqlDataReader)
        Id = IIf(String.IsNullOrEmpty(_Reader("Id").ToString()), 0, CType(_Reader("Id").ToString(), Integer))
        Id_Goc = IIf(String.IsNullOrEmpty(_Reader("Id_Goc").ToString()), 0, CType(_Reader("Id_Goc").ToString(), Integer))
        Ma_So = _Reader("Ma_So").ToString()
        Ten_Goi = _Reader("Ten_Goi").ToString()
        Ten_VT = _Reader("Ten_VT").ToString()
        Dia_Chi = _Reader("Dia_Chi").ToString()
        Dien_Thoai = _Reader("Dien_Thoai").ToString()
        Email = _Reader("Email").ToString()
        Website = _Reader("Website").ToString()
        Status = IIf(String.IsNullOrEmpty(_Reader("Status").ToString()), 0, CType(_Reader("Status").ToString(), Byte))
        SoLg_XaPhuong = IIf(String.IsNullOrEmpty(_Reader("SoLg_XaPhuong").ToString()), 0, CType(_Reader("SoLg_XaPhuong").ToString(), Integer))
        MaDV = _Reader("MaDV").ToString()
        MaCB_Begin = _Reader("MaCB_Begin").ToString()
        MaCB_End = _Reader("MaCB_End").ToString()
        Ma_So_Old = _Reader("Ma_So_Old").ToString()
        IP = _Reader("IP").ToString()
    End Sub

    Public Sub GetData(_drow As System.Data.DataRow)
        Id = IIf(String.IsNullOrEmpty(_drow("Id").ToString()), 0, CType(_drow("Id").ToString(), Integer))
        Id_Goc = IIf(String.IsNullOrEmpty(_drow("Id_Goc").ToString()), 0, CType(_drow("Id_Goc").ToString(), Integer))
        Ma_So = _drow("Ma_So").ToString()
        Ten_Goi = _drow("Ten_Goi").ToString()
        Ten_VT = _drow("Ten_VT").ToString()
        Dia_Chi = _drow("Dia_Chi").ToString()
        Dien_Thoai = _drow("Dien_Thoai ").ToString()
        Email = _drow("Email ").ToString()
        Website = _drow("Website ").ToString()
        Status = IIf(String.IsNullOrEmpty(_drow("Status").ToString()), 0, CType(_drow("Status").ToString(), Byte))
        SoLg_XaPhuong = IIf(String.IsNullOrEmpty(_drow("SoLg_XaPhuong").ToString()), 0, CType(_drow("SoLg_XaPhuong").ToString(), Integer))
        MaDV = _drow("MaDV").ToString()
        MaCB_Begin = _drow("MaCB_Begin").ToString()
        MaCB_End = _drow("MaCB_End").ToString()
        Ma_So_Old = _drow("Ma_So_Old").ToString()
        IP = _drow("IP").ToString()
    End Sub
End Class
