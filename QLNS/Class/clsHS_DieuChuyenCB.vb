Imports System
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' Author: Nguyễn Thị Thuỳ Giang
''' </summary>
''' <remarks></remarks>

Public Class QDNhanSu
    Private _IdQDNhanSu As String
    Private _So_QD As String
    Private _NgayKy_QD As Date
    Private _IdCanBo As String
    Private _IdLoaiQD As Integer
    Private _NgayHL As Date
    Private _NgayBoNhiem_TT As Date
    Private _NgayThoiLuong As Date
    Private _NguoiKy_QD As String
    Private _idCV_Nguoiky_QD As Integer
    Private _IdDonvi_Cu As Integer
    Private _IdPhong_Cu As Integer
    Private _IdChucvu_Cu As Integer
    Private _IdChuyenMon_Cu As Integer
    Private _IdDonVi_Moi As Integer
    Private _IdPhong_Moi As Integer
    Private _IdChucVu_Moi As Integer
    Private _IdChuyenMon_Moi As Integer
    Private _Active As Boolean
    Private _IsKiemNhiem As Int16
    Private _isQD_NHCS As Boolean
    Private _DenNgay As Date
    Private _NoiDung As String
    Private _DVraQD As String
    Private _CV_NguoiKy_QD As String
    Private _LoaiQD As String
    Private _GhiChu As String

    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdQDNhanSu() As String
        Get
            Return _IdQDNhanSu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID quyết định nhân sự cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdQDNhanSu = Value
        End Set
    End Property

    Public Property So_QD() As String
        Get
            Return _So_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định nhân sự cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _So_QD = Value
        End Set
    End Property

    Public Property NgayKy_QD() As Date
        Get
            Return _NgayKy_QD
        End Get
        Set(ByVal Value As Date)
            _NgayKy_QD = Value
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

    Public Property IdLoaiQD() As Integer
        Get
            Return _IdLoaiQD
        End Get
        Set(ByVal Value As Integer)
            _IdLoaiQD = Value
        End Set
    End Property

    Public Property NgayHL() As Date
        Get
            Return _NgayHL
        End Get
        Set(ByVal Value As Date)
            _NgayHL = Value
        End Set
    End Property

    Public Property NgayBoNhiem_TT() As Date
        Get
            Return _NgayBoNhiem_TT
        End Get
        Set(ByVal Value As Date)
            _NgayBoNhiem_TT = Value
        End Set
    End Property

    Public Property NgayThoiLuong() As Date
        Get
            Return _NgayThoiLuong
        End Get
        Set(ByVal Value As Date)
            _NgayThoiLuong = Value
        End Set
    End Property

    Public Property NguoiKy_QD() As String
        Get
            Return _NguoiKy_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiKy_QD = Value
        End Set
    End Property

    Public Property idCV_Nguoiky_QD() As Integer
        Get
            Return _idCV_Nguoiky_QD
        End Get
        Set(ByVal Value As Integer)
            _idCV_Nguoiky_QD = Value
        End Set
    End Property

    Public Property IdDonvi_Cu() As Integer
        Get
            Return _IdDonvi_Cu
        End Get
        Set(ByVal Value As Integer)
            _IdDonvi_Cu = Value
        End Set
    End Property

    Public Property IdPhong_Cu() As Integer
        Get
            Return _IdPhong_Cu
        End Get
        Set(ByVal Value As Integer)
            _IdPhong_Cu = Value
        End Set
    End Property

    Public Property IdChucvu_Cu() As Integer
        Get
            Return _IdChucvu_Cu
        End Get
        Set(ByVal Value As Integer)
            _IdChucvu_Cu = Value
        End Set
    End Property

    Public Property IdChuyenMon_Cu() As Integer
        Get
            Return _IdChuyenMon_Cu
        End Get
        Set(ByVal Value As Integer)
            _IdChuyenMon_Cu = Value
        End Set
    End Property

    Public Property IdDonvi_Moi() As Integer
        Get
            Return _IdDonVi_Moi
        End Get
        Set(ByVal Value As Integer)
            _IdDonVi_Moi = Value
        End Set
    End Property

    Public Property IdPhong_Moi() As Integer
        Get
            Return _IdPhong_Moi
        End Get
        Set(ByVal Value As Integer)
            _IdPhong_Moi = Value
        End Set
    End Property

    Public Property IdChucvu_Moi() As Integer
        Get
            Return _IdChucVu_Moi
        End Get
        Set(ByVal Value As Integer)
            _IdChucVu_Moi = Value
        End Set
    End Property

    Public Property IdChuyenMon_Moi() As Integer
        Get
            Return _IdChuyenMon_Moi
        End Get
        Set(ByVal Value As Integer)
            _IdChuyenMon_Moi = Value
        End Set
    End Property

    Public Property Active() As Boolean
        Get
            Return _Active
        End Get
        Set(ByVal Value As Boolean)
            _Active = Value
        End Set
    End Property

    Public Property IsKiemNhiem() As Int16
        Get
            Return _IsKiemNhiem
        End Get
        Set(ByVal Value As Int16)
            _IsKiemNhiem = Value
        End Set
    End Property

    Public Property IsQD_NHCS() As Boolean
        Get
            Return _isQD_NHCS
        End Get
        Set(ByVal Value As Boolean)
            _isQD_NHCS = Value
        End Set
    End Property

    Public Property DVraQD() As String
        Get
            Return _DVraQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Đơn vị ra quyết định có độ dài quá 200 kí tự!", Value, Value.ToString())
            End If
            _DVraQD = Value
        End Set
    End Property

    Public Property DenNgay() As Date
        Get
            Return _DenNgay
        End Get
        Set(ByVal Value As Date)
            _DenNgay = Value
        End Set
    End Property

    Public Property NoiDung() As String
        Get
            Return _NoiDung
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 1024) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Nội dung có độ dài quá 1024 kí tự!", Value, Value.ToString())
            End If
            _NoiDung = Value
        End Set
    End Property

    Public Property CV_NguoiKy_QD() As String
        Get
            Return _CV_NguoiKy_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chức danh người ký có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _CV_NguoiKy_QD = Value
        End Set
    End Property

    Public Property LoaiQD() As String
        Get
            Return _LoaiQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Loại quyết định nhân sự có độ dài quá 100 kí tự!", Value, Value.ToString())
            End If
            _LoaiQD = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDNhanSu_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@_So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@_NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@_IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_NgayHL", NgayHL))
            If NgayBoNhiem_TT = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", NgayBoNhiem_TT))
            End If
            If NgayThoiLuong = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_NgayThoiLuong", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_NgayThoiLuong", NgayThoiLuong))
            End If
            cmd.Parameters.Add(New SqlParameter("@_NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_idCV_Nguoiky_QD", idCV_Nguoiky_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdDonvi_Cu", IdDonvi_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdPhong_Cu", IdPhong_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdChucvu_Cu", IdChucvu_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdChuyenMon_Cu", IdChuyenMon_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdDonvi_Moi", IdDonvi_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdPhong_Moi", IdPhong_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdChucvu_Moi", IdChucvu_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdChuyenMon_Moi", IdChuyenMon_Moi))
            cmd.Parameters.Add(New SqlParameter("@_Active", Active))
            cmd.Parameters.Add(New SqlParameter("@_IsKiemNhiem", IsKiemNhiem))
            cmd.Parameters.Add(New SqlParameter("@_isQD_NHCS", IsQD_NHCS))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@_NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@_DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@_CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_LoaiQD", LoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_GhiChu", GhiChu))
            cmd.Parameters.Add("@_IdQDNhanSu", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdQDNhanSu = cmd.Parameters("@_IdQDNhanSu").Value.ToString
        Catch ex As Exception
            IdQDNhanSu = ""
        End Try
        Return IdQDNhanSu
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDNhanSu_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@_IdQDNhanSu", IdQDNhanSu))
            cmd.Parameters.Add(New SqlParameter("@_So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@_NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@_IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_NgayHL", NgayHL))
            If NgayBoNhiem_TT = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", NgayBoNhiem_TT))
            End If
            If NgayThoiLuong = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_NgayThoiLuong", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_NgayThoiLuong", NgayThoiLuong))
            End If
            cmd.Parameters.Add(New SqlParameter("@_NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_idCV_Nguoiky_QD", idCV_Nguoiky_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdDonvi_Cu", IdDonvi_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdPhong_Cu", IdPhong_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdChucvu_Cu", IdChucvu_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdChuyenMon_Cu", IdChuyenMon_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdDonvi_Moi", IdDonvi_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdPhong_Moi", IdPhong_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdChucvu_Moi", IdChucvu_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdChuyenMon_Moi", IdChuyenMon_Moi))
            cmd.Parameters.Add(New SqlParameter("@_Active", Active))
            cmd.Parameters.Add(New SqlParameter("@_IsKiemNhiem", IsKiemNhiem))
            cmd.Parameters.Add(New SqlParameter("@_isQD_NHCS", IsQD_NHCS))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@_NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@_DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@_CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_LoaiQD", LoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdateNotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDNhanSu_UpdateNotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@_IdQDNhanSu", IdQDNhanSu))
            cmd.Parameters.Add(New SqlParameter("@_So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@_NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@_IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_NgayHL", NgayHL))
            If NgayBoNhiem_TT = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_NgayBoNhiem_TT", NgayBoNhiem_TT))
            End If
            If NgayThoiLuong = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_NgayThoiLuong", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_NgayThoiLuong", NgayThoiLuong))
            End If
            cmd.Parameters.Add(New SqlParameter("@_NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_idCV_Nguoiky_QD", idCV_Nguoiky_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdDonvi_Cu", IdDonvi_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdPhong_Cu", IdPhong_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdChucvu_Cu", IdChucvu_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdChuyenMon_Cu", IdChuyenMon_Cu))
            cmd.Parameters.Add(New SqlParameter("@_IdDonvi_Moi", IdDonvi_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdPhong_Moi", IdPhong_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdChucvu_Moi", IdChucvu_Moi))
            cmd.Parameters.Add(New SqlParameter("@_IdChuyenMon_Moi", IdChuyenMon_Moi))
            cmd.Parameters.Add(New SqlParameter("@_Active", Active))
            cmd.Parameters.Add(New SqlParameter("@_IsKiemNhiem", IsKiemNhiem))
            cmd.Parameters.Add(New SqlParameter("@_isQD_NHCS", IsQD_NHCS))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@_DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@_DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@_NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@_DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@_CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_LoaiQD", LoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDNhanSu_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@_IdQDNhanSu", IdQDNhanSu))
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
                Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                m_QDNhanSu.IdQDNhanSu = smartReader.getString("IdQDNhanSu")
                m_QDNhanSu.So_QD = smartReader.getString("So_QD")
                m_QDNhanSu.IdCanBo = smartReader.getString("IdCanBo")
                m_QDNhanSu.IdLoaiQD = smartReader.getInt32("IdLoaiQD")
                m_QDNhanSu.NgayHL = smartReader.getDatetime("NgayHL")
                m_QDNhanSu.NgayBoNhiem_TT = smartReader.getDatetime("NgayBoNhiem_TT")
                m_QDNhanSu.NgayThoiLuong = smartReader.getDatetime("NgayThoiLuong")
                m_QDNhanSu.NgayKy_QD = smartReader.getDatetime("NgayKy_QD")
                m_QDNhanSu.NguoiKy_QD = smartReader.getString("NguoiKy_QD")
                m_QDNhanSu.idCV_Nguoiky_QD = smartReader.getInt32("idCV_Nguoiky_QD")
                m_QDNhanSu.IdDonvi_Cu = smartReader.getInt32("IdDonvi_Cu")
                m_QDNhanSu.IdPhong_Cu = smartReader.getInt32("IdPhong_Cu")
                m_QDNhanSu.IdChucvu_Cu = smartReader.getInt32("IdChucvu_Cu")
                m_QDNhanSu.IdChuyenMon_Cu = smartReader.getInt32("IdChuyenMon_Cu")
                m_QDNhanSu.IdDonvi_Moi = smartReader.getInt32("IdDonvi_Moi")
                m_QDNhanSu.IdPhong_Moi = smartReader.getInt32("IdPhong_Moi")
                m_QDNhanSu.IdChucvu_Moi = smartReader.getInt32("IdChucvu_Moi")
                m_QDNhanSu.IdChuyenMon_Moi = smartReader.getInt32("IdChuyenMon_Moi")
                m_QDNhanSu.Active = smartReader.getBoolean("Active")
                m_QDNhanSu.IsKiemNhiem = smartReader.getInt16("IsKiemNhiem")
                m_QDNhanSu.IsQD_NHCS = smartReader.getBoolean("IsQD_NHCS")
                m_QDNhanSu.DenNgay = smartReader.getDatetime("DenNgay")
                m_QDNhanSu.NoiDung = smartReader.getString("NoiDung")
                m_QDNhanSu.DVraQD = smartReader.getString("DVraQD")
                m_QDNhanSu.CV_NguoiKy_QD = smartReader.getString("CV_NguoiKy_QD")
                m_QDNhanSu.LoaiQD = smartReader.getString("LoaiQD")
                m_QDNhanSu.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_QDNhanSu)
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
            strSql = "SELECT * FROM QDNhanSu Order by NgayHL desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByCanbo(ByVal vIdCanBo As String) As IList
        Try
            Dim strSql As String
            strSql = "Select * From QDNhanSu Where IdCanBo='" & vIdCanBo.Trim & "' Order by NgayHL Desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanBo As String, Optional ByVal vNgayHL As Date = Nothing, Optional ByVal vIsQD_NHCS As Boolean = False, Optional ByVal vIsNotQDCachChuc As Boolean = False, Optional ByVal vIsKiemNhiem As Boolean = False) As QDNhanSu
        Try
            Dim strSql As String
            Dim strIsQD_NHCS As String = ""
            Dim strIsKiemNhiem As String = " And IsKiemNhiem=0 "
            Dim strIsNotQDCachChuc As String = ""
            If vIsNotQDCachChuc Then strIsNotQDCachChuc = " And IdLoaiQD Not In (Select X.Id From DanhMuc X Where X.Ma_So = '1510')"
            If vIsQD_NHCS Then strIsQD_NHCS = " And isQD_NHCS = 1 "
            If vIsKiemNhiem Then strIsKiemNhiem = " And IsKiemNhiem = 1 "
            'Dim sNgayHL As String = vNgayHL.ToString("dd/MMM/yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US"))
            Dim sNgayHL As String = vNgayHL.ToString("yyyy-MM-dd")
            If Not (vNgayHL = Nothing) Then
                strSql = "SELECT TOP 1 * From QDNhanSu Where IdCanBo='" & vIdCanBo.Trim & "' " & strIsQD_NHCS & strIsNotQDCachChuc & strIsKiemNhiem & " And DatedIff(Day,NgayHL,Cast('" & sNgayHL & "' As Date) ) > 0 Order By NgayHL Desc"
            Else
                strSql = "SELECT TOP 1 * From QDNhanSu Where IdCanBo='" & vIdCanBo.Trim & "' " & strIsQD_NHCS & strIsNotQDCachChuc & strIsKiemNhiem & "  Order By NgayHL Desc"
            End If
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), QDNhanSu)
            Else
                Return New QDNhanSu
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdQDNhanSu As String) As QDNhanSu
        Try
            Dim strSql As String
            strSql = "Select * From QDNhanSu Where IdQDNhanSu = '" & vIdQDNhanSu.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), QDNhanSu)
            Else
                Return New QDNhanSu
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getID(ByVal vIdCanBo As String, ByVal vTuNgay As Date, ByVal vSoQD As String, ByVal vNgayQD As Date) As String
        Try
            Dim strSql As String
            Dim sTuNgay As String = vTuNgay.ToString("yyyy-MM-dd")
            strSql = "SELECT IdQDNhanSu From QDNhanSu Where IdCanBo='" & vIdCanBo & "' And DatedIff(Day,NgayHL,Cast('" & sTuNgay & "' As Date) ) = 0 And So_QD = N'" & vSoQD & "' And NgayKy_QD = '" & vNgayQD & "'"
            Return db.getString(strSql)
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

End Class

Public Class QuyHoachCB
    Dim _IdQHCB As String
    Dim _Nam_QH As Integer
    Dim _Dot_QH As Integer
    Dim _IdCanBo As String
    Dim _IdChucDanh_HT As Integer
    Dim _IdTrinhDo_CM_HT As Integer
    Dim _IdTrinhDo_CT_HT As Integer
    Dim _TrinhDo_NN_HT As Integer
    Dim _TrinhDo_TH_HT As Integer
    Dim _IdChucDanh_QH_HT As Integer
    Dim _IdChucDanh_QH_LK As Integer
    Dim _IdChucDanh_QH_5Nam As Integer
    Dim _KQ_KiemPhieu_HT As String
    Dim _KQ_KiemPhieu_LTHT As String
    Dim _KQ_KiemPhieu_LK As String
    Dim _KQ_KiemPhieu_LTLK As String
    Dim _KQ_KiemPhieu_5N As String
    Dim _KQ_KiemPhieu_LT5N As String
    Dim _IdKH_DTBD_CM As Integer
    Dim _IdKH_DTBD_CT As Integer
    Dim _KH_DTBD_NN As Integer
    Dim _KH_DTBD_TH As Integer
    Dim _GhiChu As String
    Dim _DonviQH_HT As String
    Dim _DonviQH_LK As String
    Dim _DonviQH_5Nam As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdQHCB() As String
        Get
            Return _IdQHCB
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID quy hoạch cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdQHCB = Value
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

    Public Property Nam_QH() As Integer
        Get
            Return _Nam_QH
        End Get
        Set(ByVal Value As Integer)
            _Nam_QH = Value
        End Set
    End Property

    Public Property Dot_QH() As Integer
        Get
            Return _Dot_QH
        End Get
        Set(ByVal Value As Integer)
            _Dot_QH = Value
        End Set
    End Property

    Public Property IdChucDanh_HT() As Integer
        Get
            Return _IdChucDanh_HT
        End Get
        Set(ByVal Value As Integer)
            _IdChucDanh_HT = Value
        End Set
    End Property

    Public Property IdTrinhDo_CM_HT() As Integer
        Get
            Return _IdTrinhDo_CM_HT
        End Get
        Set(ByVal Value As Integer)
            _IdTrinhDo_CM_HT = Value
        End Set
    End Property

    Public Property IdTrinhDo_CT_HT() As Integer
        Get
            Return _IdTrinhDo_CT_HT
        End Get
        Set(ByVal Value As Integer)
            _IdTrinhDo_CT_HT = Value
        End Set
    End Property

    Public Property TrinhDo_NN_HT() As Integer
        Get
            Return _TrinhDo_NN_HT
        End Get
        Set(ByVal Value As Integer)
            _TrinhDo_NN_HT = Value
        End Set
    End Property

    Public Property TrinhDo_TH_HT() As Integer
        Get
            Return _TrinhDo_TH_HT
        End Get
        Set(ByVal Value As Integer)
            _TrinhDo_TH_HT = Value
        End Set
    End Property

    Public Property IdChucDanh_QH_HT() As Integer
        Get
            Return _IdChucDanh_QH_HT
        End Get
        Set(ByVal Value As Integer)
            _IdChucDanh_QH_HT = Value
        End Set
    End Property

    Public Property IdChucDanh_QH_LK() As Integer
        Get
            Return _IdChucDanh_QH_LK
        End Get
        Set(ByVal Value As Integer)
            _IdChucDanh_QH_LK = Value
        End Set
    End Property

    Public Property IdChucDanh_QH_5Nam() As Integer
        Get
            Return _IdChucDanh_QH_5Nam
        End Get
        Set(ByVal Value As Integer)
            _IdChucDanh_QH_5Nam = Value
        End Set
    End Property

    Public Property KQ_KiemPhieu_HT() As String
        Get
            Return _KQ_KiemPhieu_HT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết quả kiểm phiếu năm hiện tại có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KQ_KiemPhieu_HT = Value
        End Set
    End Property

    Public Property KQ_KiemPhieu_LTHT() As String
        Get
            Return _KQ_KiemPhieu_LTHT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết quả kiểm phiếu hội nghi liên tịch năm hiện tại có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KQ_KiemPhieu_LTHT = Value
        End Set
    End Property

    Public Property KQ_KiemPhieu_LK() As String
        Get
            Return _KQ_KiemPhieu_LK
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết quả kiểm phiếu năm liền kề có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KQ_KiemPhieu_LK = Value
        End Set
    End Property

    Public Property KQ_KiemPhieu_LTLK() As String
        Get
            Return _KQ_KiemPhieu_LTLK
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết quả kiểm phiếu hội nghi liên tịch năm liền kề có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KQ_KiemPhieu_LTLK = Value
        End Set
    End Property

    Public Property KQ_KiemPhieu_5N() As String
        Get
            Return _KQ_KiemPhieu_5N
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết quả kiểm phiếu 5 năm tiếp theo có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KQ_KiemPhieu_5N = Value
        End Set
    End Property

    Public Property KQ_KiemPhieu_LT5N() As String
        Get
            Return _KQ_KiemPhieu_LT5N
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị kết quả kiểm phiếu hội nghi liên tịch 5 năm tiếp theo có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _KQ_KiemPhieu_LT5N = Value
        End Set
    End Property

    Public Property IdKH_DTBD_CM() As Integer
        Get
            Return _IdKH_DTBD_CM
        End Get
        Set(ByVal Value As Integer)
            _IdKH_DTBD_CM = Value
        End Set
    End Property

    Public Property IdKH_DTBD_CT() As Integer
        Get
            Return _IdKH_DTBD_CT
        End Get
        Set(ByVal Value As Integer)
            _IdKH_DTBD_CT = Value
        End Set
    End Property

    Public Property KH_DTBD_NN() As Integer
        Get
            Return _KH_DTBD_NN
        End Get
        Set(ByVal Value As Integer)
            _KH_DTBD_NN = Value
        End Set
    End Property

    Public Property KH_DTBD_TH() As Integer
        Get
            Return _KH_DTBD_TH
        End Get
        Set(ByVal Value As Integer)
            _KH_DTBD_TH = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

    Public Property DonviQH_HT() As String
        Get
            Return _DonviQH_HT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Phòng/Ban/Đơn vị quy hoạch có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _DonviQH_HT = Value
        End Set
    End Property

    Public Property DonviQH_LK() As String
        Get
            Return _DonviQH_LK
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Phòng/Ban/Đơn vị quy hoạch có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _DonviQH_LK = Value
        End Set
    End Property

    Public Property DonviQH_5Nam() As String
        Get
            Return _DonviQH_5Nam
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Phòng/Ban/Đơn vị quy hoạch có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _DonviQH_5Nam = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_QHCB_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@Nam_QH", Nam_QH))
            cmd.Parameters.Add(New SqlParameter("@Dot_QH", Dot_QH))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_HT", IdChucDanh_HT))
            cmd.Parameters.Add(New SqlParameter("@IdTrinhDo_CM_HT", IdTrinhDo_CM_HT))
            cmd.Parameters.Add(New SqlParameter("@IdTrinhDo_CT_HT", IdTrinhDo_CT_HT))
            cmd.Parameters.Add(New SqlParameter("@TrinhDo_NN_HT", TrinhDo_NN_HT))
            cmd.Parameters.Add(New SqlParameter("@TrinhDo_TH_HT", TrinhDo_TH_HT))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_QH_HT", IdChucDanh_QH_HT))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_QH_LK", IdChucDanh_QH_LK))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_QH_5Nam", IdChucDanh_QH_5Nam))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_HT", KQ_KiemPhieu_HT))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LTHT", KQ_KiemPhieu_LTHT))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LK", KQ_KiemPhieu_LK))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LTLK", KQ_KiemPhieu_LTLK))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_5N", KQ_KiemPhieu_5N))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LT5N", KQ_KiemPhieu_LT5N))
            cmd.Parameters.Add(New SqlParameter("@IdKH_DTBD_CM", IdKH_DTBD_CM))
            cmd.Parameters.Add(New SqlParameter("@IdKH_DTBD_CT", IdKH_DTBD_CT))
            cmd.Parameters.Add(New SqlParameter("@KH_DTBD_NN", KH_DTBD_NN))
            cmd.Parameters.Add(New SqlParameter("@KH_DTBD_TH", KH_DTBD_TH))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add(New SqlParameter("@DonviQH_HT", DonviQH_HT))
            cmd.Parameters.Add(New SqlParameter("@DonviQH_LK", DonviQH_LK))
            cmd.Parameters.Add(New SqlParameter("@DonviQH_5Nam", DonviQH_5Nam))

            cmd.Parameters.Add("@IdQHCB", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdQHCB = cmd.Parameters("@IdQHCB").Value.ToString
        Catch ex As Exception
            IdQHCB = ""
        End Try
        Return IdQHCB
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_QHCB_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdQHCB", IdQHCB))
            cmd.Parameters.Add(New SqlParameter("@Nam_QH", Nam_QH))
            cmd.Parameters.Add(New SqlParameter("@Dot_QH", Dot_QH))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_HT", IdChucDanh_HT))
            cmd.Parameters.Add(New SqlParameter("@IdTrinhDo_CM_HT", IdTrinhDo_CM_HT))
            cmd.Parameters.Add(New SqlParameter("@IdTrinhDo_CT_HT", IdTrinhDo_CT_HT))
            cmd.Parameters.Add(New SqlParameter("@TrinhDo_NN_HT", TrinhDo_NN_HT))
            cmd.Parameters.Add(New SqlParameter("@TrinhDo_TH_HT", TrinhDo_TH_HT))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_QH_HT", IdChucDanh_QH_HT))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_QH_LK", IdChucDanh_QH_LK))
            cmd.Parameters.Add(New SqlParameter("@IdChucDanh_QH_5Nam", IdChucDanh_QH_5Nam))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_HT", KQ_KiemPhieu_HT))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LTHT", KQ_KiemPhieu_LTHT))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LK", KQ_KiemPhieu_LK))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LTLK", KQ_KiemPhieu_LTLK))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_5N", KQ_KiemPhieu_5N))
            cmd.Parameters.Add(New SqlParameter("@KQ_KiemPhieu_LT5N", KQ_KiemPhieu_LT5N))
            cmd.Parameters.Add(New SqlParameter("@IdKH_DTBD_CM", IdKH_DTBD_CM))
            cmd.Parameters.Add(New SqlParameter("@IdKH_DTBD_CT", IdKH_DTBD_CT))
            cmd.Parameters.Add(New SqlParameter("@KH_DTBD_NN", KH_DTBD_NN))
            cmd.Parameters.Add(New SqlParameter("@KH_DTBD_TH", KH_DTBD_TH))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add(New SqlParameter("@DonviQH_HT", DonviQH_HT))
            cmd.Parameters.Add(New SqlParameter("@DonviQH_LK", DonviQH_LK))
            cmd.Parameters.Add(New SqlParameter("@DonviQH_5Nam", DonviQH_5Nam))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_QHCB_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdQHCB", IdQHCB))
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
                Dim m_QuyHoachCB As QuyHoachCB = New QuyHoachCB
                m_QuyHoachCB.IdQHCB = smartReader.getString("IdQHCB")
                m_QuyHoachCB.Nam_QH = smartReader.getInt32("Nam_QH")
                m_QuyHoachCB.Dot_QH = smartReader.getInt32("Dot_QH")
                m_QuyHoachCB.IdCanBo = smartReader.getString("IdCanBo")
                m_QuyHoachCB.IdChucDanh_HT = smartReader.getInt32("IdChucDanh_HT")
                m_QuyHoachCB.IdTrinhDo_CM_HT = smartReader.getInt32("IdTrinhDo_CM_HT")
                m_QuyHoachCB.IdTrinhDo_CT_HT = smartReader.getInt32("IdTrinhDo_CT_HT")
                m_QuyHoachCB.TrinhDo_NN_HT = smartReader.getInt32("TrinhDo_NN_HT")
                m_QuyHoachCB.TrinhDo_TH_HT = smartReader.getInt32("TrinhDo_TH_HT")
                m_QuyHoachCB.IdChucDanh_QH_HT = smartReader.getInt32("IdChucDanh_QH_HT")
                m_QuyHoachCB.IdChucDanh_QH_LK = smartReader.getInt32("IdChucDanh_QH_LK")
                m_QuyHoachCB.IdChucDanh_QH_5Nam = smartReader.getInt32("IdChucDanh_QH_5Nam")
                m_QuyHoachCB.KQ_KiemPhieu_HT = smartReader.getString("KQ_KiemPhieu_HT")
                m_QuyHoachCB.KQ_KiemPhieu_LTHT = smartReader.getString("KQ_KiemPhieu_LTHT")
                m_QuyHoachCB.KQ_KiemPhieu_LK = smartReader.getString("KQ_KiemPhieu_LK")
                m_QuyHoachCB.KQ_KiemPhieu_LTLK = smartReader.getString("KQ_KiemPhieu_LTLK")
                m_QuyHoachCB.KQ_KiemPhieu_5N = smartReader.getString("KQ_KiemPhieu_5N")
                m_QuyHoachCB.KQ_KiemPhieu_LT5N = smartReader.getString("KQ_KiemPhieu_LT5N")
                m_QuyHoachCB.IdKH_DTBD_CM = smartReader.getInt32("IdKH_DTBD_CM")
                m_QuyHoachCB.IdKH_DTBD_CT = smartReader.getInt32("IdKH_DTBD_CT")
                m_QuyHoachCB.KH_DTBD_NN = smartReader.getInt32("KH_DTBD_NN")
                m_QuyHoachCB.KH_DTBD_TH = smartReader.getInt32("KH_DTBD_TH")
                m_QuyHoachCB.GhiChu = smartReader.getString("GhiChu")
                m_QuyHoachCB.DonviQH_HT = smartReader.getString("DonviQH_HT")
                m_QuyHoachCB.DonviQH_LK = smartReader.getString("DonviQH_LK")
                m_QuyHoachCB.DonviQH_5Nam = smartReader.getString("DonviQH_5Nam")
                list.Add(m_QuyHoachCB)
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
            strSql = "SELECT * FROM HS_QHCB Order by Nam_QH desc, Dot_QH desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByCanbo(ByVal vIdCanBo As String) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_QHCB WHERE IdCanBo='" & vIdCanBo.Trim & "' Order by Nam_QH desc, Dot_QH desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdQuyHoachCB As String) As QuyHoachCB
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_QHCB WHERE IdQHCB='" & vIdQuyHoachCB.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), QuyHoachCB)
            Else
                Return New QuyHoachCB
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanBo As String) As QuyHoachCB
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_QHCB WHERE IdCanBo='" & vIdCanBo.Trim & "' Order by Nam_QH desc, Dot_QH desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), QuyHoachCB)
            Else
                Return New QuyHoachCB
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

Public Class CBThoiViec
    Private _IdCBThoiViec As String
    Private _IdCanBo As String
    Private _IdLoaiQD As Integer
    Private _IdLyDo As Integer
    Private _NgayKy_QD As DateTime
    Private _Ngay_HL As DateTime
    Private _NguoiKy_QD As String
    Private _IdCV_NguoKy_QD As Integer
    Private _TroCap_ThoiViec As Long
    Private _TroCap_Khac As Long
    Private _isQD_NHCS As Boolean
    Private _So_QD As String '(20),
    Private _DVraQD As String ' (200),
    Private _SoTien_BoiThuong As Long
    Private _SoTien_ThuHoi As Long
    Private _CV_NguoiKy_QD As String
    Private _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdCBThoiViec() As String
        Get
            Return _IdCBThoiViec
        End Get
        Set(ByVal value As String)
            If Not (value Is Nothing) Then
                If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ thôi việc có độ dài không hợp lệ!", value, value.ToString())
            End If
            _IdCBThoiViec = value
        End Set
    End Property

    Public Property IsQD_NHCS() As Boolean
        Get
            Return _isQD_NHCS
        End Get
        Set(ByVal Value As Boolean)
            _isQD_NHCS = Value
        End Set
    End Property

    Public Property DVraQD() As String
        Get
            Return _DVraQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Đơn vị ra quyết định có độ dài quá 200 kí tự!", Value, Value.ToString())
            End If
            _DVraQD = Value
        End Set
    End Property

    Public Property IdCanBo() As String
        Get
            Return _IdCanBo
        End Get
        Set(ByVal value As String)
            If Not (value Is Nothing) Then
                If (value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị id hồ sơ cán bộ thôi việc có độ dài không hợp lệ!", value, value.ToString())
            End If
            _IdCanBo = value
        End Set
    End Property

    Public Property So_QD() As String
        Get
            Return _So_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định cán bộ có độ dài quá 20 kí tự!", Value, Value.ToString())
            End If
            _So_QD = Value
        End Set
    End Property

    Public Property IdLoaiQD() As Integer
        Get
            Return _IdLoaiQD
        End Get
        Set(ByVal value As Integer)
            _IdLoaiQD = value
        End Set
    End Property

    Public Property IdLyDo() As Integer
        Get
            Return _IdLyDo
        End Get
        Set(ByVal value As Integer)
            _IdLyDo = value
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

    Public Property Ngay_HL() As DateTime
        Get
            Return _Ngay_HL
        End Get
        Set(ByVal value As DateTime)
            _Ngay_HL = value
        End Set
    End Property

    Public Property NguoiKy_QD() As String
        Get
            Return _NguoiKy_QD
        End Get
        Set(ByVal value As String)
            If Not (value Is Nothing) Then
                If (value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người ký có độ dài không hợp lệ!", value, value.ToString())
            End If
            _NguoiKy_QD = value
        End Set
    End Property

    Public Property IdCV_NguoKy_QD() As Integer
        Get
            Return _IdCV_NguoKy_QD
        End Get
        Set(ByVal value As Integer)
            _IdCV_NguoKy_QD = value
        End Set
    End Property

    Public Property TroCap_ThoiViec() As Long
        Get
            Return _TroCap_ThoiViec
        End Get
        Set(ByVal value As Long)
            _TroCap_ThoiViec = value
        End Set
    End Property

    Public Property TroCap_Khac() As Long
        Get
            Return _TroCap_Khac
        End Get
        Set(ByVal value As Long)
            _TroCap_Khac = value
        End Set
    End Property

    Public Property SoTien_BoiThuong() As Long
        Get
            Return _SoTien_BoiThuong
        End Get
        Set(ByVal Value As Long)
            _SoTien_BoiThuong = Value
        End Set
    End Property

    Public Property SoTien_ThuHoi() As Long
        Get
            Return _SoTien_ThuHoi
        End Get
        Set(ByVal Value As Long)
            _SoTien_ThuHoi = Value
        End Set
    End Property

    Public Property CV_NguoiKy_QD() As String
        Get
            Return _CV_NguoiKy_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chức danh người ký có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _CV_NguoiKy_QD = Value
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

#End Region

#Region "Method"
    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_CBThoiViec_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@_IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@_IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_IdLyDo", IdLyDo))
            cmd.Parameters.Add(New SqlParameter("@_NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_Ngay_HL", Ngay_HL))
            cmd.Parameters.Add(New SqlParameter("@_NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdCV_NguoKy_QD", IdCV_NguoKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_TroCap_ThoiViec", TroCap_ThoiViec))
            cmd.Parameters.Add(New SqlParameter("@_TroCap_Khac", TroCap_Khac))
            cmd.Parameters.Add(New SqlParameter("@_isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@_So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@_DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@_SoTien_BoiThuong", SoTien_BoiThuong))
            cmd.Parameters.Add(New SqlParameter("@_SoTien_ThuHoi", SoTien_ThuHoi))
            cmd.Parameters.Add(New SqlParameter("@_CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_GhiChu", GhiChu))
            cmd.Parameters.Add("@_IdCBThoiViec", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdCBThoiViec = cmd.Parameters("@_IdCBThoiViec").Value.ToString
        Catch ex As Exception
            IdCBThoiViec = ""
        End Try
        Return IdCBThoiViec
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_CBThoiViec_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@_IdCBThoiViec", IdCBThoiViec))
            cmd.Parameters.Add(New SqlParameter("@_IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@_IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@_IdLyDo", IdLyDo))
            cmd.Parameters.Add(New SqlParameter("@_NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_Ngay_HL", Ngay_HL))
            cmd.Parameters.Add(New SqlParameter("@_NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_IdCV_NguoKy_QD", IdCV_NguoKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_TroCap_ThoiViec", TroCap_ThoiViec))
            cmd.Parameters.Add(New SqlParameter("@_TroCap_Khac", TroCap_Khac))
            cmd.Parameters.Add(New SqlParameter("@_isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@_So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@_DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@_SoTien_BoiThuong", SoTien_BoiThuong))
            cmd.Parameters.Add(New SqlParameter("@_SoTien_ThuHoi", SoTien_ThuHoi))
            cmd.Parameters.Add(New SqlParameter("@_CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@_GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_CBThoiViec_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@_IdCBThoiViec", IdCBThoiViec))
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
                Dim m_CBThoiViec As CBThoiViec = New CBThoiViec
                m_CBThoiViec.IdCBThoiViec = smartReader.getString("IdCBThoiViec")
                m_CBThoiViec.IdCanBo = smartReader.getString("IdCanBo")
                m_CBThoiViec.IdLoaiQD = smartReader.getInt32("IdLoaiQD")
                m_CBThoiViec.IdLyDo = smartReader.getInt32("IdLyDo")
                m_CBThoiViec.NgayKy_QD = smartReader.getDatetime("NgayKy_QD")
                m_CBThoiViec.Ngay_HL = smartReader.getDatetime("Ngay_HL")
                m_CBThoiViec.NguoiKy_QD = smartReader.getString("NguoiKy_QD")
                m_CBThoiViec.IdCV_NguoKy_QD = smartReader.getInt32("IdCV_NguoKy_QD")
                m_CBThoiViec.TroCap_ThoiViec = smartReader.getInt64("TroCap_ThoiViec")
                m_CBThoiViec.TroCap_Khac = smartReader.getInt64("TroCap_Khac")
                m_CBThoiViec.IsQD_NHCS = smartReader.getBoolean("IsQD_NHCS")
                m_CBThoiViec.So_QD = smartReader.getString("So_QD")
                m_CBThoiViec.DVraQD = smartReader.getString("DVraQD")
                m_CBThoiViec.SoTien_BoiThuong = smartReader.getInt64("SoTien_BoiThuong")
                m_CBThoiViec.SoTien_ThuHoi = smartReader.getInt64("SoTien_ThuHoi")
                m_CBThoiViec.CV_NguoiKy_QD = smartReader.getString("CV_NguoiKy_QD")
                m_CBThoiViec.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_CBThoiViec)
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
            strSql = "SELECT * FROM HS_CBThoiViec Order by Ngay_HL desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByCanbo(ByVal vIdCanBo As String) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_CBThoiViec WHERE IdCanBo='" & vIdCanBo.Trim & "' Order by Ngay_HL desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdCBThoiViec As String) As CBThoiViec
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_CBThoiViec WHERE IdCBThoiViec='" & vIdCBThoiViec.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), CBThoiViec)
            Else
                Return New CBThoiViec
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanBo As String) As CBThoiViec
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_CBThoiViec WHERE IdCanBo='" & vIdCanBo.Trim & "' Order by Ngay_HL desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), CBThoiViec)
            Else
                Return New CBThoiViec
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

Public Class QDKhac
    Private _IdQDKhac As String
    Private _IdCanBo As String
    Private _So_QD As String
    Private _NgayKy_QD As Date
    Private _NguoiKy_QD As String
    Private _CV_Nguoiky_QD As String ' (50)
    Private _DVraQD As String
    Private _NoiDung As String
    Private _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdQDKhac() As String
        Get
            Return _IdQDKhac
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID quyết định khác của cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdQDKhac = Value
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

    Public Property NgayKy_QD() As Date
        Get
            Return _NgayKy_QD
        End Get
        Set(ByVal Value As Date)
            _NgayKy_QD = Value
        End Set
    End Property

    Public Property So_QD() As String
        Get
            Return _So_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _So_QD = Value
        End Set
    End Property

    Public Property NguoiKy_QD() As String
        Get
            Return _NguoiKy_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiKy_QD = Value
        End Set
    End Property

    Public Property CV_Nguoiky_QD() As String
        Get
            Return _CV_Nguoiky_QD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị chức vụ người quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _CV_Nguoiky_QD = Value
        End Set
    End Property

    Public Property DVraQD() As String
        Get
            Return _DVraQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Đơn vị ra quyết định có độ dài quá 200 kí tự!", Value, Value.ToString())
            End If
            _DVraQD = Value
        End Set
    End Property

    Public Property NoiDung() As String
        Get
            Return _NoiDung
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 1024) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Nội dung có độ dài quá 1024 kí tự!", Value, Value.ToString())
            End If
            _NoiDung = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDKhac_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@CV_Nguoiky_QD", CV_Nguoiky_QD))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdQDKhac", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdQDKhac = cmd.Parameters("@IdQDKhac").Value.ToString
        Catch ex As Exception
            IdQDKhac = ""
        End Try
        Return IdQDKhac
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDKhac_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdQDKhac", IdQDKhac))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NgayKy_QD", NgayKy_QD))
            cmd.Parameters.Add(New SqlParameter("@So_QD", So_QD))
            cmd.Parameters.Add(New SqlParameter("@NguoiKy_QD", NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@CV_Nguoiky_QD", CV_Nguoiky_QD))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("QDKhac_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdQDKhac", IdQDKhac))
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
                Dim m_QDKhacCB As QDKhac = New QDKhac
                m_QDKhacCB.IdQDKhac = smartReader.getString("IdQDKhac")
                m_QDKhacCB.IdCanBo = smartReader.getString("IdCanBo")
                m_QDKhacCB.So_QD = smartReader.getString("So_QD")
                m_QDKhacCB.NgayKy_QD = smartReader.getDatetime("NgayKy_QD")
                m_QDKhacCB.NguoiKy_QD = smartReader.getString("NguoiKy_QD")
                m_QDKhacCB.CV_Nguoiky_QD = smartReader.getString("CV_Nguoiky_QD")
                m_QDKhacCB.NoiDung = smartReader.getString("NoiDung")
                m_QDKhacCB.DVraQD = smartReader.getString("DVraQD")
                m_QDKhacCB.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_QDKhacCB)
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
            strSql = "SELECT * FROM QDKhac Order by NgayKy_QD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByCanbo(ByVal vIdCanBo As String) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM QDKhac WHERE IdCanBo='" & vIdCanBo.Trim & "' order by NgayKy_QD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdQDKhac As String) As QDKhac
        Try
            Dim strSql As String
            strSql = "SELECT * FROM QDKhac WHERE IdQDKhac='" & vIdQDKhac.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), QDKhac)
            Else
                Return New QDKhac
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanBo As String) As QDKhac
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM QDKhac WHERE IdCanBo='" & vIdCanBo.Trim & "' order by NgayKy_QD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), QDKhac)
            Else
                Return New QDKhac
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getID(ByVal vIdCanBo As String, ByVal vSo_QD As String, ByVal vNgayKy_QD As Date) As String
        Try
            Dim strSql As String
            strSql = "SELECT IdQDKhac FROM QDKhac WHERE IdCanBo='" & vIdCanBo & "' And So_QD=N'" & vSo_QD & "' And NgayKy_QD='" & vNgayKy_QD & "'"
            Return db.getString(strSql)
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

End Class