Imports System
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Class Đánh giá cán bộ
''' Author: Nguyễn Thị Thuỳ Giang
''' </summary>
''' <remarks></remarks>
Public Class DanhGiaCB
    Dim _IdDGCB As String
    Dim _Nam_DG As Integer
    Dim _Dot_DG As Integer
    Dim _IdCanBo As String
    Dim _CongViec_DN As String
    Dim _KetQua_TH As String
    Dim _PhamChat_CT As String
    Dim _DaoDuc_LS As String
    Dim _NhanXet_TT As String
    Dim _NhanXet_BT As String
    Dim _IdDanhGiaCV As Integer
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdDGCB() As String
        Get
            Return _IdDGCB
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID đánh giá cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdDGCB = Value
        End Set
    End Property

    Public Property Nam_DG() As Integer
        Get
            Return _Nam_DG
        End Get
        Set(ByVal Value As Integer)
            _Nam_DG = Value
        End Set
    End Property

    Public Property Dot_DG() As Integer
        Get
            Return _Dot_DG
        End Get
        Set(ByVal Value As Integer)
            _Dot_DG = Value
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

    Public Property CongViec_DN() As String
        Get
            Return _CongViec_DN
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Công việc đảm nhận có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _CongViec_DN = Value
        End Set
    End Property

    Public Property KetQua_TH() As String
        Get
            Return _KetQua_TH
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Kết quả thực hiện có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KetQua_TH = Value
        End Set
    End Property

    Public Property PhamChat_CT() As String
        Get
            Return _PhamChat_CT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Phẩm chất chính trị có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _PhamChat_CT = Value
        End Set
    End Property

    Public Property DaoDuc_LS() As String
        Get
            Return _DaoDuc_LS
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Đạo đức nối sống có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _DaoDuc_LS = Value
        End Set
    End Property

    Public Property NhanXet_TT() As String
        Get
            Return _NhanXet_TT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Nhận xét tập thể có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NhanXet_TT = Value
        End Set
    End Property

    Public Property NhanXet_BT() As String
        Get
            Return _NhanXet_BT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Nhận xét bản thân có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NhanXet_BT = Value
        End Set
    End Property

    Public Property IdDanhGiaCV() As Integer
        Get
            Return _IdDanhGiaCV
        End Get
        Set(ByVal Value As Integer)
            _IdDanhGiaCV = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_DanhGiaCB_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@Nam_DG", Nam_DG))
            cmd.Parameters.Add(New SqlParameter("@Dot_DG", Dot_DG))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@CongViec_DN", CongViec_DN))
            cmd.Parameters.Add(New SqlParameter("@KetQua_TH", KetQua_TH))
            cmd.Parameters.Add(New SqlParameter("@PhamChat_CT", PhamChat_CT))
            cmd.Parameters.Add(New SqlParameter("@DaoDuc_LS", DaoDuc_LS))
            cmd.Parameters.Add(New SqlParameter("@NhanXet_TT", NhanXet_TT))
            cmd.Parameters.Add(New SqlParameter("@NhanXet_BT", NhanXet_BT))
            cmd.Parameters.Add(New SqlParameter("@IdDanhGiaCV", IdDanhGiaCV))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdDGCB", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdDGCB = cmd.Parameters("@IdDGCB").Value.ToString
        Catch ex As Exception
            IdDGCB = ""
        End Try
        Return IdDGCB
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_DanhGiaCB_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDGCB", IdDGCB))
            cmd.Parameters.Add(New SqlParameter("@Nam_DG", Nam_DG))
            cmd.Parameters.Add(New SqlParameter("@Dot_DG", Dot_DG))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@CongViec_DN", CongViec_DN))
            cmd.Parameters.Add(New SqlParameter("@KetQua_TH", KetQua_TH))
            cmd.Parameters.Add(New SqlParameter("@PhamChat_CT", PhamChat_CT))
            cmd.Parameters.Add(New SqlParameter("@DaoDuc_LS", DaoDuc_LS))
            cmd.Parameters.Add(New SqlParameter("@NhanXet_TT", NhanXet_TT))
            cmd.Parameters.Add(New SqlParameter("@NhanXet_BT", NhanXet_BT))
            cmd.Parameters.Add(New SqlParameter("@IdDanhGiaCV", IdDanhGiaCV))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_DanhGiaCB_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDGCB", IdDGCB))
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
                Dim m_DanhGiaCB As DanhGiaCB = New DanhGiaCB
                m_DanhGiaCB.IdDGCB = smartReader.getString("IdDGCB")
                m_DanhGiaCB.Nam_DG = smartReader.getInt32("Nam_DG")
                m_DanhGiaCB.Dot_DG = smartReader.getInt32("Dot_DG")
                m_DanhGiaCB.IdCanBo = smartReader.getString("IdCanBo")
                m_DanhGiaCB.CongViec_DN = smartReader.getString("CongViec_DN")
                m_DanhGiaCB.KetQua_TH = smartReader.getString("KetQua_TH")
                m_DanhGiaCB.PhamChat_CT = smartReader.getString("PhamChat_CT")
                m_DanhGiaCB.DaoDuc_LS = smartReader.getString("DaoDuc_LS")
                m_DanhGiaCB.NhanXet_TT = smartReader.getString("NhanXet_TT")
                m_DanhGiaCB.NhanXet_BT = smartReader.getString("NhanXet_BT")
                m_DanhGiaCB.IdDanhGiaCV = smartReader.getInt32("IdDanhGiaCV")
                m_DanhGiaCB.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_DanhGiaCB)
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
            strSql = "SELECT * FROM HS_DanhGiaCB Order by Nam_DG desc, Dot_DG desc"
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
            strSql = "SELECT * FROM HS_DanhGiaCB WHERE idCanbo='" & vIdCanbo.Trim & "' order by Nam_DG desc, Dot_DG desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdDGCB As String) As DanhGiaCB
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_DanhGiaCB WHERE IdDGCB='" & vIdDGCB.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), DanhGiaCB)
            Else
                Return New DanhGiaCB
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As DanhGiaCB
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_DanhGiaCB WHERE idCanbo='" & vIdCanbo.Trim & "' order by Nam_DG desc, Dot_DG desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), DanhGiaCB)
            Else
                Return New DanhGiaCB
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

''' <summary>
''' Class Tín nhiệm lãnh đạo
''' </summary>
''' <remarks></remarks>
Public Class TinNhiemLD

    Dim _IdTNhLD As String
    Dim _IdCanBo As String
    Dim _Ngay As Date
    Dim _IdChucDanh As Integer
    Dim _TongPhieu As Integer
    Dim _SoPhieu_TN As Integer
    Dim _TyLe_TN As Double
    Dim _Phieu_KTN_NL As Integer
    Dim _TyLe_KTN_NL As Double
    Dim _Phieu_KTN_DD As Integer
    Dim _TyLe_KTN_DD As Double
    Dim _Phieu_KTN_khac As Integer
    Dim _TyLe_KTN_khac As Double
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdTNhLD() As String
        Get
            Return _IdTNhLD
        End Get
        Set(ByVal Value As String)
            _IdTNhLD = Value
        End Set
    End Property

    Public Property Ngay() As Date
        Get
            Return _Ngay
        End Get
        Set(ByVal Value As Date)
            _Ngay = Value
        End Set
    End Property

    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal Value As String)
            _IdCanBo = Value
        End Set
    End Property

    Public Property IdChucDanh() As Integer
        Get
            Return _IdChucDanh
        End Get
        Set(ByVal Value As Integer)
            _IdChucDanh = Value
        End Set
    End Property

    Public Property TongPhieu() As Integer
        Get
            Return _TongPhieu
        End Get
        Set(ByVal Value As Integer)
            _TongPhieu = Value
        End Set
    End Property

    Public Property SoPhieu_TN() As Integer
        Get
            Return _SoPhieu_TN
        End Get
        Set(ByVal Value As Integer)
            _SoPhieu_TN = Value
        End Set
    End Property

    Public Property TyLe_TN() As Double
        Get
            Return _TyLe_TN
        End Get
        Set(ByVal Value As Double)
            _TyLe_TN = Value
        End Set
    End Property

    Public Property Phieu_KTN_NL() As Integer
        Get
            Return _Phieu_KTN_NL
        End Get
        Set(ByVal Value As Integer)
            _Phieu_KTN_NL = Value
        End Set
    End Property

    Public Property TyLe_KTN_NL() As Double
        Get
            Return _TyLe_KTN_NL
        End Get
        Set(ByVal Value As Double)
            _TyLe_KTN_NL = Value
        End Set
    End Property

    Public Property Phieu_KTN_DD() As Integer
        Get
            Return _Phieu_KTN_DD
        End Get
        Set(ByVal Value As Integer)
            _Phieu_KTN_DD = Value
        End Set
    End Property

    Public Property TyLe_KTN_DD() As Double
        Get
            Return _TyLe_KTN_DD
        End Get
        Set(ByVal Value As Double)
            _TyLe_KTN_DD = Value
        End Set
    End Property

    Public Property Phieu_KTN_khac() As Integer
        Get
            Return _Phieu_KTN_khac
        End Get
        Set(ByVal Value As Integer)
            _Phieu_KTN_khac = Value
        End Set
    End Property

    Public Property TyLe_KTN_khac() As Double
        Get
            Return _TyLe_KTN_khac
        End Get
        Set(ByVal Value As Double)
            _TyLe_KTN_khac = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_TNhLD_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@Ngay", Ngay))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh", IdChucDanh))
            cmd.Parameters.Add(New SqlParameter("@TongPhieu", TongPhieu))
            cmd.Parameters.Add(New SqlParameter("@SoPhieu_TN", SoPhieu_TN))
            cmd.Parameters.Add(New SqlParameter("@TyLe_TN", TyLe_TN))
            cmd.Parameters.Add(New SqlParameter("@Phieu_KTN_NL", Phieu_KTN_NL))
            cmd.Parameters.Add(New SqlParameter("@TyLe_KTN_NL", TyLe_KTN_NL))
            cmd.Parameters.Add(New SqlParameter("@Phieu_KTN_DD", _Phieu_KTN_DD))
            cmd.Parameters.Add(New SqlParameter("@TyLe_KTN_DD", TyLe_KTN_DD))
            cmd.Parameters.Add(New SqlParameter("@Phieu_KTN_khac", Phieu_KTN_khac))
            cmd.Parameters.Add(New SqlParameter("@TyLe_KTN_khac", TyLe_KTN_khac))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdTNhLD", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdTNhLD = cmd.Parameters("@IdTNhLD").Value.ToString
        Catch ex As Exception
            IdTNhLD = ""
        End Try
        Return IdTNhLD
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_TNhLD_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdTNhLD", IdTNhLD))
            cmd.Parameters.Add(New SqlParameter("@Ngay", Ngay))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh", IdChucDanh))
            cmd.Parameters.Add(New SqlParameter("@TongPhieu", TongPhieu))
            cmd.Parameters.Add(New SqlParameter("@SoPhieu_TN", SoPhieu_TN))
            cmd.Parameters.Add(New SqlParameter("@TyLe_TN", TyLe_TN))
            cmd.Parameters.Add(New SqlParameter("@Phieu_KTN_NL", Phieu_KTN_NL))
            cmd.Parameters.Add(New SqlParameter("@TyLe_KTN_NL", TyLe_KTN_NL))
            cmd.Parameters.Add(New SqlParameter("@Phieu_KTN_DD", _Phieu_KTN_DD))
            cmd.Parameters.Add(New SqlParameter("@TyLe_KTN_DD", TyLe_KTN_DD))
            cmd.Parameters.Add(New SqlParameter("@Phieu_KTN_khac", Phieu_KTN_khac))
            cmd.Parameters.Add(New SqlParameter("@TyLe_KTN_khac", TyLe_KTN_khac))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_TNhLD_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdTNhLD", IdTNhLD))
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
                Dim m_TinNhiemLD As TinNhiemLD = New TinNhiemLD
                m_TinNhiemLD.IdTNhLD = smartReader.getString("IdTNhLD")
                m_TinNhiemLD.Ngay = smartReader.getDatetime("Ngay")
                m_TinNhiemLD.IdCanBo = smartReader.getString("IdCanBo")
                m_TinNhiemLD.IdChucDanh = smartReader.getInt32("IdChucDanh")
                m_TinNhiemLD.TongPhieu = smartReader.getInt32("TongPhieu")
                m_TinNhiemLD.SoPhieu_TN = smartReader.getInt32("SoPhieu_TN")
                m_TinNhiemLD.TyLe_TN = smartReader.getFloat("TyLe_TN")
                m_TinNhiemLD.Phieu_KTN_NL = smartReader.getInt32("Phieu_KTN_NL")
                m_TinNhiemLD.TyLe_KTN_NL = smartReader.getFloat("TyLe_KTN_NL")
                m_TinNhiemLD.Phieu_KTN_DD = smartReader.getInt32("Phieu_KTN_DD")
                m_TinNhiemLD.TyLe_KTN_DD = smartReader.getFloat("TyLe_KTN_DD")
                m_TinNhiemLD.Phieu_KTN_khac = smartReader.getInt32("Phieu_KTN_khac")
                m_TinNhiemLD.TyLe_KTN_khac = smartReader.getFloat("TyLe_KTN_khac")
                m_TinNhiemLD.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_TinNhiemLD)
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
            strSql = "SELECT * FROM HS_TNhLD Order by Ngay desc"
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
            strSql = "SELECT * FROM HS_TNhLD WHERE idCanbo='" & vIdCanbo.Trim & "' order by Ngay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdTNhLD As String) As TinNhiemLD
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_TNhLD WHERE IdTNhLD='" & vIdTNhLD.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), TinNhiemLD)
            Else
                Return New TinNhiemLD
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As TinNhiemLD
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_TNhLD WHERE idCanbo='" & vIdCanbo.Trim & "' order by Ngay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), TinNhiemLD)
            Else
                Return New TinNhiemLD
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

''' <summary>
''' Class Nhận xét lãnh đạo
''' </summary>
''' <remarks></remarks>
Public Class NhanXetLD

    Dim _IdNXLD As String
    Dim _Nam As Integer
    Dim _IdCanBo As String
    Dim _TongPhieu As Integer
    Dim _NangLuc_CT As Byte
    Dim _PhamChat_DD As Byte
    Dim _ChieuHuong_PT As Byte
    Dim _DeXuat As Byte
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdNXLD() As String
        Get
            Return _IdNXLD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID nhận xét lãnh đạo có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdNXLD = Value
        End Set
    End Property

    Public Property Nam() As Integer
        Get
            Return _Nam
        End Get
        Set(ByVal Value As Integer)
            _Nam = Value
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

    Public Property TongPhieu() As Integer
        Get
            Return _TongPhieu
        End Get
        Set(ByVal Value As Integer)
            _TongPhieu = Value
        End Set
    End Property

    Public Property NangLuc_CT() As Byte
        Get
            Return _NangLuc_CT
        End Get
        Set(ByVal Value As Byte)
            _NangLuc_CT = Value
        End Set
    End Property

    Public Property PhamChat_DD() As Byte
        Get
            Return _PhamChat_DD
        End Get
        Set(ByVal Value As Byte)
            _PhamChat_DD = Value
        End Set
    End Property

    Public Property ChieuHuong_PT() As Byte
        Get
            Return _ChieuHuong_PT
        End Get
        Set(ByVal Value As Byte)
            _ChieuHuong_PT = Value
        End Set
    End Property

    Public Property DeXuat() As Byte
        Get
            Return _DeXuat
        End Get
        Set(ByVal Value As Byte)
            _DeXuat = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_NXLD_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@TongPhieu", TongPhieu))
            cmd.Parameters.Add(New SqlParameter("@NangLuc_CT", NangLuc_CT))
            cmd.Parameters.Add(New SqlParameter("@PhamChat_DD", PhamChat_DD))
            cmd.Parameters.Add(New SqlParameter("@ChieuHuong_PT", ChieuHuong_PT))
            cmd.Parameters.Add(New SqlParameter("@DeXuat", DeXuat))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdNXLD", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdNXLD = cmd.Parameters("@IdNXLD").Value.ToString
        Catch ex As Exception
            IdNXLD = ""
        End Try
        Return IdNXLD
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_NXLD_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNXLD", IdNXLD))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@TongPhieu", TongPhieu))
            cmd.Parameters.Add(New SqlParameter("@NangLuc_CT", NangLuc_CT))
            cmd.Parameters.Add(New SqlParameter("@PhamChat_DD", PhamChat_DD))
            cmd.Parameters.Add(New SqlParameter("@ChieuHuong_PT", ChieuHuong_PT))
            cmd.Parameters.Add(New SqlParameter("@DeXuat", DeXuat))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_NXLD_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNXLD", IdNXLD))
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
                Dim m_NhanXetLD As NhanXetLD = New NhanXetLD
                m_NhanXetLD.IdNXLD = smartReader.getString("IdNXLD")
                m_NhanXetLD.Nam = smartReader.getInt32("Nam")
                m_NhanXetLD.IdCanBo = smartReader.getString("IdCanBo")
                m_NhanXetLD.TongPhieu = smartReader.getInt32("TongPhieu")
                m_NhanXetLD.NangLuc_CT = smartReader.getByte("NangLuc_CT")
                m_NhanXetLD.PhamChat_DD = smartReader.getByte("PhamChat_DD")
                m_NhanXetLD.ChieuHuong_PT = smartReader.getByte("ChieuHuong_PT")
                m_NhanXetLD.DeXuat = smartReader.getByte("DeXuat")
                m_NhanXetLD.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_NhanXetLD)
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
            strSql = "SELECT * FROM HS_NXLD Order by Nam desc"
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
            strSql = "SELECT * FROM HS_NXLD WHERE idCanbo='" & vIdCanbo.Trim & "' order by Nam desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdNXLD As String) As NhanXetLD
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_NXLD WHERE IdNXLD='" & vIdNXLD.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), NhanXetLD)
            Else
                Return New NhanXetLD
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As NhanXetLD
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_NXLD WHERE idCanbo='" & vIdCanbo.Trim & "' order by Nam desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), NhanXetLD)
            Else
                Return New NhanXetLD
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class
