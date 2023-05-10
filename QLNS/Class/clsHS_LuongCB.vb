Imports System
Imports System.Data
Imports System.Data.SqlClient
''' <summary>
''' Author: Nguyễn Thị Thuỳ Giang
''' </summary>
''' <remarks></remarks>
Public Class LuongCanBo
    Private _IdLuongCB As String
    Private _IdCanBo As String
    Private _IdBacLuong As Integer
    Private _HeSoLuong As Double
    Private _Ngay_Huong As Date
    Private _NgayLen_DK As Date
    Private _SoQD As String
    Private _NgayQD As Date
    Private _NguoiQD As String
    Private _IdCV_Nguoi_QD As Integer
    Private _IdLoaiQD As Integer
    Private _isQD_NHCS As Boolean
    Private _NoiDung As String
    Private _DVraQD As String
    Private _CV_NguoiKy_QD As String
    Private _LoaiQD As String
    Private _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdLuongCB() As String
        Get
            Return _IdLuongCB
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID lương cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdLuongCB = Value
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

    Public Property IdBacLuong() As Integer
        Get
            Return _IdBacLuong
        End Get
        Set(ByVal Value As Integer)
            _IdBacLuong = Value
        End Set
    End Property

    Public Property HeSoLuong() As Double
        Get
            Return _HeSoLuong
        End Get
        Set(ByVal Value As Double)
            _HeSoLuong = Value
        End Set
    End Property

    Public Property Ngay_Huong() As Date
        Get
            Return _Ngay_Huong
        End Get
        Set(ByVal Value As Date)
            _Ngay_Huong = Value
        End Set
    End Property

    Public Property NgayLen_DK() As Date
        Get
            Return _NgayLen_DK
        End Get
        Set(ByVal Value As Date)
            _NgayLen_DK = Value
        End Set
    End Property

    Public Property SoQD() As String
        Get
            Return _SoQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _SoQD = Value
        End Set
    End Property

    Public Property NgayQD() As Date
        Get
            Return _NgayQD
        End Get
        Set(ByVal Value As Date)
            _NgayQD = Value
        End Set
    End Property

    Public Property NguoiQD() As String
        Get
            Return _NguoiQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiQD = Value
        End Set
    End Property

    Public Property IdCV_Nguoi_QD() As Integer
        Get
            Return _IdCV_Nguoi_QD
        End Get
        Set(ByVal Value As Integer)
            _IdCV_Nguoi_QD = Value
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
                If (Value.Length > 100) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_LuongCB_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdBacLuong", IdBacLuong))
            cmd.Parameters.Add(New SqlParameter("@HeSoLuong", HeSoLuong))
            cmd.Parameters.Add(New SqlParameter("@Ngay_Huong", Ngay_Huong))
            cmd.Parameters.Add(New SqlParameter("@NgayLen_DK", NgayLen_DK))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@LoaiQD", LoaiQD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdLuongCB", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdLuongCB = cmd.Parameters("@IdLuongCB").Value.ToString
        Catch ex As Exception
            IdLuongCB = ""
        End Try
        Return IdLuongCB
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_LuongCB_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdLuongCB", IdLuongCB))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBacLuong", IdBacLuong))
            cmd.Parameters.Add(New SqlParameter("@HeSoLuong", HeSoLuong))
            cmd.Parameters.Add(New SqlParameter("@Ngay_Huong", Ngay_Huong))
            cmd.Parameters.Add(New SqlParameter("@NgayLen_DK", NgayLen_DK))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiQD", IdLoaiQD))
            cmd.Parameters.Add(New SqlParameter("@isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@LoaiQD", LoaiQD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_LuongCB_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdLuongCB", IdLuongCB))
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
                Dim m_LuongCB As LuongCanBo = New LuongCanBo
                m_LuongCB.IdLuongCB = smartReader.getString("IdLuongCB")
                m_LuongCB.IdCanBo = smartReader.getString("IdCanBo")
                m_LuongCB.IdBacLuong = smartReader.getInt32("IdBacLuong")
                m_LuongCB.HeSoLuong = smartReader.getFloat("HeSoLuong")
                m_LuongCB.Ngay_Huong = smartReader.getDatetime("Ngay_Huong")
                m_LuongCB.NgayLen_DK = smartReader.getDatetime("NgayLen_DK")
                m_LuongCB.SoQD = smartReader.getString("SoQD")
                m_LuongCB.NgayQD = smartReader.getDatetime("NgayQD")
                m_LuongCB.NguoiQD = smartReader.getString("NguoiQD")
                m_LuongCB.IdCV_Nguoi_QD = smartReader.getInt32("IdCV_Nguoi_QD")
                m_LuongCB.IdLoaiQD = smartReader.getInt32("IdLoaiQD")
                m_LuongCB.IsQD_NHCS = smartReader.getBoolean("IsQD_NHCS")
                m_LuongCB.NoiDung = smartReader.getString("NoiDung")
                m_LuongCB.DVraQD = smartReader.getString("DVraQD")
                m_LuongCB.CV_NguoiKy_QD = smartReader.getString("CV_NguoiKy_QD")
                m_LuongCB.LoaiQD = smartReader.getString("LoaiQD")
                m_LuongCB.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_LuongCB)
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
            strSql = "SELECT * FROM HS_LuongCB Order by Ngay_Huong desc"
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
            strSql = "SELECT * FROM HS_LuongCB WHERE idCanbo='" & vIdCanbo.Trim & "' order by Ngay_Huong desc, IdLuongCB desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdLuongCB As String) As LuongCanBo
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_LuongCB WHERE IdLuongCB='" & vIdLuongCB.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), LuongCanBo)
            Else
                Return New LuongCanBo
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String, Optional ByVal vNgayHuong As Date = Nothing) As LuongCanBo
        Try
            Dim strSql As String
            If Not (vNgayHuong = Nothing) Then
                strSql = "SELECT TOP 1 * FROM HS_LuongCB WHERE idCanbo='" & vIdCanbo.Trim & "' and datediff(day,Ngay_Huong,'" & vNgayHuong & "')>0 order by Ngay_Huong desc, IdLuongCB desc"
            Else
                strSql = "SELECT TOP 1 * FROM HS_LuongCB WHERE idCanbo='" & vIdCanbo.Trim & "' order by Ngay_Huong desc, IdLuongCB desc"
            End If
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), LuongCanBo)
            Else
                Return New LuongCanBo
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getID(ByVal vIdCanbo As String, ByVal vTuNgay As Date, ByVal vSoQD As String, ByVal vNgayQD As Date) As String
        Try
            Dim strSql As String
            strSql = "SELECT IdLuongCB FROM HS_LuongCB WHERE IdCanbo='" & vIdCanbo & "' and Datediff(day,Ngay_Huong,'" & vTuNgay & "')=0 and SoQD=N'" & vSoQD & "' and NgayQD='" & vNgayQD & "'"
            Return db.getString(strSql)
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

End Class


Public Class Phucap
    Private _IdCB_PhuCap As String
    Private _IdCanBo As String
    Private _IdMucPC As Integer
    Private _TuNgay As Date
    Private _DenNgay As Date
    Private _SoQD As String
    Private _NgayQD As Date
    Private _NguoiQD As String
    Private _IdCV_Nguoi_QD As Integer
    Private _isQD_NHCS As Boolean
    Private _NoiDung As String
    Private _DVraQD As String
    Private _CV_NguoiKy_QD As String
    Private _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdCB_PhuCap() As String
        Get
            Return _IdCB_PhuCap
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID phụ cấp cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdCB_PhuCap = Value
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

    Public Property IdMucPC() As Integer
        Get
            Return _IdMucPC
        End Get
        Set(ByVal Value As Integer)
            _IdMucPC = Value
        End Set
    End Property

    Public Property TuNgay() As Date
        Get
            Return _TuNgay
        End Get
        Set(ByVal Value As Date)
            _TuNgay = Value
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

    Public Property NgayQD() As Date
        Get
            Return _NgayQD
        End Get
        Set(ByVal Value As Date)
            _NgayQD = Value
        End Set
    End Property

    Public Property SoQD() As String
        Get
            Return _SoQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 20) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _SoQD = Value
        End Set
    End Property

    Public Property NguoiQD() As String
        Get
            Return _NguoiQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiQD = Value
        End Set
    End Property

    Public Property IdCV_Nguoi_QD() As Integer
        Get
            Return _IdCV_Nguoi_QD
        End Get
        Set(ByVal Value As Integer)
            _IdCV_Nguoi_QD = Value
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
            Dim cmd As SqlCommand = New SqlCommand("HS_PhuCapCB_Insert")

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdMucPC", IdMucPC))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdCB_PhuCap", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdCB_PhuCap = cmd.Parameters("@IdCB_PhuCap").Value.ToString
        Catch ex As Exception
            IdCB_PhuCap = ""
        End Try
        Return IdCB_PhuCap
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_PhuCapCB_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCB_PhuCap", IdCB_PhuCap))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdMucPC", IdMucPC))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdateNotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_PhuCapCB_UpdateNotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCB_PhuCap", IdCB_PhuCap))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@isQD_NHCS", IsQD_NHCS))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@DVraQD", DVraQD))
            cmd.Parameters.Add(New SqlParameter("@CV_NguoiKy_QD", CV_NguoiKy_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_PhuCapCB_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCB_PhuCap", IdCB_PhuCap))
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
                Dim m_PhucapCB As Phucap = New Phucap
                m_PhucapCB.IdCB_PhuCap = smartReader.getString("IdCB_PhuCap")
                m_PhucapCB.IdCanBo = smartReader.getString("IdCanBo")
                m_PhucapCB.IdMucPC = smartReader.getInt32("IdMucPC")
                m_PhucapCB.TuNgay = smartReader.getDatetime("TuNgay")
                m_PhucapCB.DenNgay = smartReader.getDatetime("DenNgay")
                m_PhucapCB.SoQD = smartReader.getString("SoQD")
                m_PhucapCB.NgayQD = smartReader.getDatetime("NgayQD")
                m_PhucapCB.NguoiQD = smartReader.getString("NguoiQD")
                m_PhucapCB.IdCV_Nguoi_QD = smartReader.getInt32("IdCV_Nguoi_QD")
                m_PhucapCB.IsQD_NHCS = smartReader.getBoolean("IsQD_NHCS")
                m_PhucapCB.NoiDung = smartReader.getString("NoiDung")
                m_PhucapCB.DVraQD = smartReader.getString("DVraQD")
                m_PhucapCB.CV_NguoiKy_QD = smartReader.getString("CV_NguoiKy_QD")
                m_PhucapCB.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_PhucapCB)
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
            strSql = "SELECT * FROM HS_PhucapCB Order by TuNgay desc"
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
            strSql = "SELECT * FROM HS_PhucapCB WHERE idCanbo='" & vIdCanbo.Trim & "' order by TuNgay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdCB_Phucap As String) As Phucap
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_PhucapCB WHERE IdCB_Phucap='" & vIdCB_Phucap.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), Phucap)
            Else
                Return New Phucap
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As Phucap
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_PhucapCB WHERE idCanbo='" & vIdCanbo.Trim & "' order by TuNgay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), Phucap)
            Else
                Return New Phucap
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getID(ByVal vIdCanbo As String, ByVal vTuNgay As Date, ByVal vSoQD As String, ByVal vNgayQD As Date) As String
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 IdCB_PhuCap FROM HS_PhucapCB WHERE IdCanbo='" & vIdCanbo & "' and Datediff(day,TuNgay,'" & vTuNgay & "')=0 and Upper(SoQD)=Upper(N'" & vSoQD & "') and NgayQD='" & vNgayQD & "'"
            Return db.getString(strSql)
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

End Class

Public Class Themgio
    Dim _IdThemGio As String
    Dim _IdDonVi As Integer
    Dim _IdPhong As String
    Dim _IdCanBo As String
    Dim _Thang As Integer
    Dim _Nam As Integer
    Dim _SoGio150 As Double
    Dim _SoGio200 As Double
    Dim _SoGio300 As Double
    Dim _ThucLinh As Long
    Dim _TraThem As Long
    Dim _SoGioLamDemNgayThuong As Double
    Dim _SoGioLamDemTBayCNhat As Double
    Dim _IDHT_ThanhToan As Integer
    Dim _TT_50 As Double
    Dim _TT_100 As Double
    Dim _TT_150 As Double
    Dim _TT_200 As Double
    Dim _TT_300 As Double
    Dim _TT_GioNghibu As Double
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdThemGio() As String
        Get
            Return _IdThemGio
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID thêm giờ cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdThemGio = Value
        End Set
    End Property

    Public Property IdDonVi_TG() As String
        Get
            Return _IdDonVi
        End Get
        Set(ByVal Value As String)
            _IdDonVi = Value
        End Set
    End Property

    Public Property IdPhong() As String
        Get
            Return _IdPhong
        End Get
        Set(ByVal Value As String)
            _IdPhong = Value
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

    Public Property Thang() As Integer
        Get
            Return _Thang
        End Get
        Set(ByVal Value As Integer)
            _Thang = Value
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

    Public Property SoGio150() As Double
        Get
            Return _SoGio150
        End Get
        Set(ByVal Value As Double)
            _SoGio150 = Value
        End Set
    End Property

    Public Property SoGio200() As Double
        Get
            Return _SoGio200
        End Get
        Set(ByVal Value As Double)
            _SoGio200 = Value
        End Set
    End Property

    Public Property SoGio300() As Double
        Get
            Return _SoGio300
        End Get
        Set(ByVal Value As Double)
            _SoGio300 = Value
        End Set
    End Property

    Public Property TraThem() As Long
        Get
            Return _TraThem
        End Get
        Set(ByVal Value As Long)
            _TraThem = Value
        End Set
    End Property

    Public Property ThucLinh() As Long
        Get
            Return _ThucLinh
        End Get
        Set(ByVal Value As Long)
            _ThucLinh = Value
        End Set
    End Property

    Public Property SoGioLamDemNgayThuong() As Double
        Get
            Return _SoGioLamDemNgayThuong
        End Get
        Set(ByVal Value As Double)
            _SoGioLamDemNgayThuong = Value
        End Set
    End Property

    Public Property SoGioLamDemTBayCNhat() As Double
        Get
            Return _SoGioLamDemTBayCNhat
        End Get
        Set(ByVal Value As Double)
            _SoGioLamDemTBayCNhat = Value
        End Set
    End Property

    Public Property IDHT_ThanhToan() As Integer
        Get
            Return _IDHT_ThanhToan
        End Get
        Set(ByVal Value As Integer)
            _IDHT_ThanhToan = Value
        End Set
    End Property

    Public Property TT_50() As Double
        Get
            Return _TT_50
        End Get
        Set(ByVal Value As Double)
            _TT_50 = Value
        End Set
    End Property

    Public Property TT_100() As Double
        Get
            Return _TT_100
        End Get
        Set(ByVal Value As Double)
            _TT_100 = Value
        End Set
    End Property

    Public Property TT_150() As Double
        Get
            Return _TT_150
        End Get
        Set(ByVal Value As Double)
            _TT_150 = Value
        End Set
    End Property

    Public Property TT_200() As Double
        Get
            Return _TT_200
        End Get
        Set(ByVal Value As Double)
            _TT_200 = Value
        End Set
    End Property

    Public Property TT_300() As Double
        Get
            Return _TT_300
        End Get
        Set(ByVal Value As Double)
            _TT_300 = Value
        End Set
    End Property

    Public Property TT_GioNghibu() As Double
        Get
            Return _TT_GioNghibu
        End Get
        Set(ByVal Value As Double)
            _TT_GioNghibu = Value
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

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ThemGio_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_TG))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@SoGio150", SoGio150))
            cmd.Parameters.Add(New SqlParameter("@SoGio200", SoGio200))
            cmd.Parameters.Add(New SqlParameter("@SoGio300", SoGio300))
            cmd.Parameters.Add(New SqlParameter("@TraThem", TraThem))
            cmd.Parameters.Add(New SqlParameter("@ThucLinh", ThucLinh))
            cmd.Parameters.Add(New SqlParameter("@SoGioLamDemNgayThuong", SoGioLamDemNgayThuong))
            cmd.Parameters.Add(New SqlParameter("@SoGioLamDemTBayCNhat", SoGioLamDemTBayCNhat))
            cmd.Parameters.Add(New SqlParameter("@IdHT_ThanhToan", IDHT_ThanhToan))
            cmd.Parameters.Add(New SqlParameter("@TT_50", TT_50))
            cmd.Parameters.Add(New SqlParameter("@TT_100", TT_100))
            cmd.Parameters.Add(New SqlParameter("@TT_150", TT_150))
            cmd.Parameters.Add(New SqlParameter("@TT_200", TT_200))
            cmd.Parameters.Add(New SqlParameter("@TT_300", TT_300))
            cmd.Parameters.Add(New SqlParameter("@TT_GioNghibu", TT_GioNghibu))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdThemGio", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdThemGio = cmd.Parameters("@IdThemGio").Value.ToString
        Catch ex As Exception
            IdThemGio = ""
        End Try
        Return IdThemGio
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ThemGio_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdThemGio", IdThemGio))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_TG))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@SoGio150", SoGio150))
            cmd.Parameters.Add(New SqlParameter("@SoGio200", SoGio200))
            cmd.Parameters.Add(New SqlParameter("@SoGio300", SoGio300))
            cmd.Parameters.Add(New SqlParameter("@TraThem", TraThem))
            cmd.Parameters.Add(New SqlParameter("@ThucLinh", ThucLinh))
            cmd.Parameters.Add(New SqlParameter("@SoGioLamDemNgayThuong", SoGioLamDemNgayThuong))
            cmd.Parameters.Add(New SqlParameter("@SoGioLamDemTBayCNhat", SoGioLamDemTBayCNhat))
            cmd.Parameters.Add(New SqlParameter("@IdHT_ThanhToan", IDHT_ThanhToan))
            cmd.Parameters.Add(New SqlParameter("@TT_50", TT_50))
            cmd.Parameters.Add(New SqlParameter("@TT_100", TT_100))
            cmd.Parameters.Add(New SqlParameter("@TT_150", TT_150))
            cmd.Parameters.Add(New SqlParameter("@TT_200", TT_200))
            cmd.Parameters.Add(New SqlParameter("@TT_300", TT_300))
            cmd.Parameters.Add(New SqlParameter("@TT_GioNghibu", TT_GioNghibu))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ThemGio_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdThemGio", IdThemGio))
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
                Dim m_Themgio As Themgio = New Themgio
                m_Themgio.IdThemGio = smartReader.getString("IdThemGio")
                m_Themgio.IdDonVi_TG = smartReader.getInt32("IdDonVi")
                m_Themgio.IdPhong = smartReader.getInt32("IdPhong")
                m_Themgio.IdCanBo = smartReader.getString("IdCanBo")
                m_Themgio.Thang = smartReader.getInt32("Thang")
                m_Themgio.Nam = smartReader.getInt32("Nam")
                m_Themgio.SoGio150 = smartReader.getFloat("SoGio150")
                m_Themgio.SoGio200 = smartReader.getFloat("SoGio200")
                m_Themgio.SoGio300 = smartReader.getFloat("SoGio300")
                m_Themgio.TraThem = smartReader.getInt64("TraThem")
                m_Themgio.ThucLinh = smartReader.getInt64("ThucLinh")
                m_Themgio.SoGioLamDemNgayThuong = smartReader.getFloat("SoGioLamDemNgayThuong")
                m_Themgio.SoGioLamDemTBayCNhat = smartReader.getFloat("SoGioLamDemTBayCNhat")
                m_Themgio.IDHT_ThanhToan = smartReader.getFloat("IDHT_ThanhToan")
                m_Themgio.TT_50 = smartReader.getFloat("TT_50")
                m_Themgio.TT_100 = smartReader.getFloat("TT_100")
                m_Themgio.TT_150 = smartReader.getFloat("TT_150")
                m_Themgio.TT_200 = smartReader.getFloat("TT_200")
                m_Themgio.TT_300 = smartReader.getFloat("TT_300")
                m_Themgio.TT_GioNghibu = smartReader.getFloat("TT_GioNghibu")
                m_Themgio.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_Themgio)
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
            strSql = "SELECT * FROM HS_Themgio Order by nam desc, thang desc"
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
            strSql = "SELECT * FROM HS_Themgio WHERE idCanbo='" & vIdCanbo.Trim & "' order by nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdThemGio As String) As Themgio
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_Themgio WHERE IdThemGio='" & vIdThemGio.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), Themgio)
            Else
                Return New Themgio
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As Themgio
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_Themgio WHERE idCanbo='" & vIdCanbo.Trim & "' order by nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), Themgio)
            Else
                Return New Themgio
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
#End Region

End Class

Public Class HS_ChiBosung_CT
    Dim _IdHS_ChiBoSung_CT As String
    Dim _IdHS_ChiLuong As String
    Dim _IdLoaiLuongBS As Integer
    Dim _IdPhong As Integer
    Dim _IdCanBo As String
    Dim _NH_TS As Int16
    Dim _IdChucVu As Integer
    Dim _HesoCMNV As Double
    Dim _PC_CV As Double
    Dim _PC_TN As Double
    Dim _PC_KV As Double
    Dim _PC_DH As Double
    Dim _PC_KN As Double
    Dim _PC_TNVK As Double
    Dim _PC_TNN As Double
    Dim _PC_Khac As Double
    Dim _Tong100 As Long
    Dim _Hesochi As Double
    Dim _TongChi As Long
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdHS_ChiBoSung_CT() As String
        Get
            Return _IdHS_ChiBoSung_CT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Lương bổ sung cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiBoSung_CT = Value
        End Set
    End Property

    Public Property IdHS_ChiLuong() As String
        Get
            Return _IdHS_ChiLuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Chi Lương có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiLuong = Value
        End Set
    End Property

    Public Property IdLoaiLuongBS() As Integer
        Get
            Return _IdLoaiLuongBS
        End Get
        Set(ByVal Value As Integer)
            _IdLoaiLuongBS = Value
        End Set
    End Property

    Public Property IdPhong() As Integer
        Get
            Return _IdPhong
        End Get
        Set(ByVal Value As Integer)
            _IdPhong = Value
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

    Public Property NH_TS() As Int16
        Get
            Return _NH_TS
        End Get
        Set(ByVal Value As Int16)
            _NH_TS = Value
        End Set
    End Property

    Public Property IdChucVu() As Integer
        Get
            Return _IdChucVu
        End Get
        Set(ByVal Value As Integer)
            _IdChucVu = Value
        End Set
    End Property

    Public Property HesoCMNV() As Double
        Get
            Return _HesoCMNV
        End Get
        Set(ByVal Value As Double)
            _HesoCMNV = Value
        End Set
    End Property

    Public Property PC_CV() As Double
        Get
            Return _PC_CV
        End Get
        Set(ByVal Value As Double)
            _PC_CV = Value
        End Set
    End Property

    Public Property PC_TN() As Double
        Get
            Return _PC_TN
        End Get
        Set(ByVal Value As Double)
            _PC_TN = Value
        End Set
    End Property

    Public Property PC_KV() As Double
        Get
            Return _PC_KV
        End Get
        Set(ByVal Value As Double)
            _PC_KV = Value
        End Set
    End Property

    Public Property PC_DH() As Double
        Get
            Return _PC_DH
        End Get
        Set(ByVal Value As Double)
            _PC_DH = Value
        End Set
    End Property

    Public Property PC_KN() As Double
        Get
            Return _PC_KN
        End Get
        Set(ByVal Value As Double)
            _PC_KN = Value
        End Set
    End Property

    Public Property PC_TNVK() As Double
        Get
            Return _PC_TNVK
        End Get
        Set(ByVal Value As Double)
            _PC_TNVK = Value
        End Set
    End Property

    Public Property PC_TNN() As Double
        Get
            Return _PC_TNN
        End Get
        Set(ByVal Value As Double)
            _PC_TNN = Value
        End Set
    End Property

    Public Property PC_Khac() As Double
        Get
            Return _PC_Khac
        End Get
        Set(ByVal Value As Double)
            _PC_Khac = Value
        End Set
    End Property

    Public Property Tong100() As Long
        Get
            Return _Tong100
        End Get
        Set(ByVal Value As Long)
            _Tong100 = Value
        End Set
    End Property
   
    Public Property Hesochi() As Double
        Get
            Return _Hesochi
        End Get
        Set(ByVal Value As Double)
            _Hesochi = Value
        End Set
    End Property

    Public Property TongChi() As Long
        Get
            Return _TongChi
        End Get
        Set(ByVal Value As Long)
            _TongChi = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiBoSung_CT_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiLuongBS", IdLoaiLuongBS))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NH_TS", NH_TS))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdChucVu", IdChucVu))
            cmd.Parameters.Add(New SqlParameter("@HesoCMNV", HesoCMNV))
            cmd.Parameters.Add(New SqlParameter("@PC_CV", PC_CV))
            cmd.Parameters.Add(New SqlParameter("@PC_TN", PC_TN))
            cmd.Parameters.Add(New SqlParameter("@PC_KV", PC_KV))
            cmd.Parameters.Add(New SqlParameter("@PC_DH", PC_DH))
            cmd.Parameters.Add(New SqlParameter("@PC_KN", PC_KN))
            cmd.Parameters.Add(New SqlParameter("@PC_TNVK", PC_TNVK))
            cmd.Parameters.Add(New SqlParameter("@PC_TNN", PC_TNN))
            cmd.Parameters.Add(New SqlParameter("@PC_Khac", PC_Khac))
            cmd.Parameters.Add(New SqlParameter("@Tong100", Tong100))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            cmd.Parameters.Add("@IdHS_ChiBoSung_CT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdHS_ChiBoSung_CT = cmd.Parameters("@IdHS_ChiBoSung_CT").Value.ToString
        Catch ex As Exception
            IdHS_ChiBoSung_CT = ""
        End Try
        Return IdHS_ChiBoSung_CT
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiBoSung_CT_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiBoSung_CT", IdHS_ChiBoSung_CT))
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiLuongBS", IdLoaiLuongBS))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NH_TS", NH_TS))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdChucVu", IdChucVu))
            cmd.Parameters.Add(New SqlParameter("@HesoCMNV", HesoCMNV))
            cmd.Parameters.Add(New SqlParameter("@PC_CV", PC_CV))
            cmd.Parameters.Add(New SqlParameter("@PC_TN", PC_TN))
            cmd.Parameters.Add(New SqlParameter("@PC_KV", PC_KV))
            cmd.Parameters.Add(New SqlParameter("@PC_DH", PC_DH))
            cmd.Parameters.Add(New SqlParameter("@PC_KN", PC_KN))
            cmd.Parameters.Add(New SqlParameter("@PC_TNVK", PC_TNVK))
            cmd.Parameters.Add(New SqlParameter("@PC_TNN", PC_TNN))
            cmd.Parameters.Add(New SqlParameter("@PC_Khac", PC_Khac))
            cmd.Parameters.Add(New SqlParameter("@Tong100", Tong100))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdatenotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiBoSung_CT_UpdatenotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiBoSung_CT", IdHS_ChiBoSung_CT))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiBoSung_CT_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiBoSung_CT", IdHS_ChiBoSung_CT))
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
                Dim m_ChiBosung_CT As HS_ChiBosung_CT = New HS_ChiBosung_CT
                m_ChiBosung_CT.IdHS_ChiBoSung_CT = smartReader.getString("IdHS_ChiBoSung_CT")
                m_ChiBosung_CT.IdHS_ChiLuong = smartReader.getString("IdHS_ChiLuong")
                m_ChiBosung_CT.IdLoaiLuongBS = smartReader.getInt32("IdLoaiLuongBS")
                m_ChiBosung_CT.IdCanBo = smartReader.getString("IdCanBo")
                m_ChiBosung_CT.NH_TS = smartReader.getInt16("NH_TS")
                m_ChiBosung_CT.IdPhong = smartReader.getInt32("IdPhong")
                m_ChiBosung_CT.IdChucVu = smartReader.getInt32("IdChucVu")
                m_ChiBosung_CT.HesoCMNV = smartReader.getFloat("HesoCMNV")
                m_ChiBosung_CT.PC_CV = smartReader.getFloat("PC_CV")
                m_ChiBosung_CT.PC_TN = smartReader.getFloat("PC_TN")
                m_ChiBosung_CT.PC_KV = smartReader.getFloat("PC_KV")
                m_ChiBosung_CT.PC_DH = smartReader.getFloat("PC_DH")
                m_ChiBosung_CT.PC_KN = smartReader.getFloat("PC_KN")
                m_ChiBosung_CT.PC_TNVK = smartReader.getFloat("PC_TNVK")
                m_ChiBosung_CT.PC_TNN = smartReader.getFloat("PC_TNN")
                m_ChiBosung_CT.PC_Khac = smartReader.getFloat("PC_Khac")
                m_ChiBosung_CT.Tong100 = smartReader.getInt64("Tong100")
                m_ChiBosung_CT.Hesochi = smartReader.getFloat("Hesochi")
                m_ChiBosung_CT.TongChi = smartReader.getInt64("TongChi")
                list.Add(m_ChiBosung_CT)
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
            strSql = "SELECT * FROM HS_ChiBoSung_CT Order by IdChiLuong desc, IdLoaiLuongBS desc"
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
            strSql = "SELECT * FROM HS_ChiBoSung_CT WHERE idCanbo='" & vIdCanbo.Trim & "' order by nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdHS_ChiBoSung_CT As String) As HS_ChiBosung_CT
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ChiBoSung_CT WHERE IdHS_ChiBoSung_CT='" & vIdHS_ChiBoSung_CT.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), HS_ChiBosung_CT)
            Else
                Return New HS_ChiBosung_CT
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As HS_ChiBosung_CT
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_ChiBoSung_CT WHERE idCanbo='" & vIdCanbo.Trim & "' order by IdChiLuong desc, IdLoaiLuongBS desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), HS_ChiBosung_CT)
            Else
                Return New HS_ChiBosung_CT
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    'Public Function getDup_CB_TS(ByVal vIdDonVi As Integer, ByVal vIdLoaiLuongBS As Integer, ByVal vNam As Integer) As IList
    '    Try
    '        Dim strSql As String
    '        strSql = "SELECT Id, '0' as IdLoaiLuongBS, Thang, Nam, IdCanBo, Hesochi, ThucLinh, GhiChu FROM HS_LuongBosung WHERE IdLoaiLuongBS = " & vIdLoaiLuongBS & " AND Nam=" & vNam & _
    '                 " 						     AND IdCanBo in (SELECT distinct  IdCanBo FROM QDNhansu t1 WHERE t1.idDonvi_Moi= " & vIdDonVi & _
    '                 "						                     AND t1.idcanbo not in (SELECT idcanbo FROM QDNhansu t2 WHERE t2.idCanbo=t1.idCanbo and datediff(second,t1.ngayHL,t2.ngayHL)>0 and year(ngayHL)<=" & vNam & ") " & _
    '                 "						                     AND idcanbo not in (SELECT IdCanbo FROM HS_CBThoiViec WHERE year(ngay_HL)<" & vNam & "))" & _
    '                 " UNION " & _
    '                 " SELECT Id, '1' as IdLoaiLuongBS, Thang, Nam, IdCanBo, Hesochi, ThucLinh, GhiChu FROM HSCB_TS_LuongBoSung  WHERE IdLoaiLuongBS = " & vIdLoaiLuongBS & " AND Nam=" & vNam & " AND IdCanBo in (SELECT Id FROM HSCB_TS WHERE IdChiNhanh=" & vIdDonVi & ")" & _
    '                 " order by Thang desc, IdLoaiLuongBS"
    '        Dim cmd As SqlCommand = New SqlCommand(strSql)
    '        cmd.CommandType = CommandType.Text
    '        Return init(cmd)
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function

#End Region

End Class

Public Class LuongBoSung_CBTS
    Dim _Id As String
    Dim _IdLoaiLuongBS As Integer
    Dim _Thang As Integer
    Dim _Nam As Integer
    Dim _Hesochi As Double
    Dim _IdCanBo As String
    Dim _ThucLinh As Long
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property Id() As String
        Get
            Return _Id
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID lương bổ sung cán bộ tập sự có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _Id = Value
        End Set
    End Property

    Public Property IdLoaiLuongBS() As Integer
        Get
            Return _IdLoaiLuongBS
        End Get
        Set(ByVal Value As Integer)
            _IdLoaiLuongBS = Value
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

    Public Property Thang() As Integer
        Get
            Return _Thang
        End Get
        Set(ByVal Value As Integer)
            _Thang = Value
        End Set
    End Property

    Public Property Hesochi() As Double
        Get
            Return _Hesochi
        End Get
        Set(ByVal Value As Double)
            _Hesochi = Value
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

    Public Property ThucLinh() As Long
        Get
            Return _ThucLinh
        End Get
        Set(ByVal Value As Long)
            _ThucLinh = Value
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
            Dim cmd As SqlCommand = New SqlCommand("HSCB_TS_LuongBoSung_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiLuongBS", IdLoaiLuongBS))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@ThucLinh", ThucLinh))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@Id", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            Id = cmd.Parameters("@Id").Value.ToString
        Catch ex As Exception
            Id = ""
        End Try
        Return Id
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HSCB_TS_LuongBoSung_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@Id", Id))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiLuongBS", IdLoaiLuongBS))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@ThucLinh", ThucLinh))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HSCB_TS_LuongBoSung_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@Id", Id))
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
                Dim m_LuongBosung As LuongBoSung_CBTS = New LuongBoSung_CBTS
                m_LuongBosung.Id = smartReader.getString("Id")
                m_LuongBosung.IdCanBo = smartReader.getString("IdCanBo")
                m_LuongBosung.Thang = smartReader.getInt32("Thang")
                m_LuongBosung.Nam = smartReader.getInt32("Nam")
                m_LuongBosung.IdLoaiLuongBS = smartReader.getInt32("IdLoaiLuongBS")
                m_LuongBosung.Hesochi = smartReader.getFloat("Hesochi")
                m_LuongBosung.ThucLinh = smartReader.getInt64("ThucLinh")
                m_LuongBosung.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_LuongBosung)
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
            strSql = "SELECT * FROM HSCB_TS_LuongBoSung Order by Nam desc, thang desc"
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
            strSql = "SELECT * FROM HSCB_TS_LuongBoSung WHERE idCanbo='" & vIdCanbo.Trim & "' order by nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdLuongBosung As String) As LuongBoSung_CBTS
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HSCB_TS_LuongBoSung WHERE Id='" & vIdLuongBosung.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), LuongBoSung_CBTS)
            Else
                Return New LuongBoSung_CBTS
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As LuongBoSung_CBTS
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HSCB_TS_LuongBoSung WHERE idCanbo='" & vIdCanbo.Trim & "' order by nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), LuongBoSung_CBTS)
            Else
                Return New LuongBoSung_CBTS
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

Public Class HS_ChiLuong
    Dim _IdHS_ChiLuong As String
    Dim _IdDonVi As Integer
    Dim _Nam As Integer
    Dim _Thang As Integer
    Dim _IdLoaiThuChi As Integer
    Dim _Title As String
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdHS_ChiLuong() As String
        Get
            Return _IdHS_ChiLuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Lương bổ sung cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiLuong = Value
        End Set
    End Property

    Public Property IdDonVi_CL() As Integer
        Get
            Return _IdDonVi
        End Get
        Set(ByVal Value As Integer)
            _IdDonVi = Value
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

    Public Property Thang() As Integer
        Get
            Return _Thang
        End Get
        Set(ByVal Value As Integer)
            _Thang = Value
        End Set
    End Property

    Public Property IdLoaiThuChi() As Integer
        Get
            Return _IdLoaiThuChi
        End Get
        Set(ByVal Value As Integer)
            _IdLoaiThuChi = Value
        End Set
    End Property

    Public Property Title() As String
        Get
            Return _Title
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Tiêu đề có độ dài vượt quá 256 kí tự!", Value, Value.ToString())
            End If
            _Title = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 1024) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_CL))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiThuChi", IdLoaiThuChi))
            cmd.Parameters.Add(New SqlParameter("@Title", Title))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdHS_ChiLuong", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdHS_ChiLuong = cmd.Parameters("@IdHS_ChiLuong").Value.ToString
        Catch ex As Exception
            IdHS_ChiLuong = ""
        End Try
        Return IdHS_ChiLuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_CL))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiThuChi", IdLoaiThuChi))
            cmd.Parameters.Add(New SqlParameter("@Title", Title))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub Update_Confirm()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_Update_Confirm")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
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
                Dim m_HS_ChiLuong As HS_ChiLuong = New HS_ChiLuong
                m_HS_ChiLuong.IdHS_ChiLuong = smartReader.getString("IdHS_ChiLuong")
                m_HS_ChiLuong.IdDonVi_CL = smartReader.getString("IdDonVi")
                m_HS_ChiLuong.Thang = smartReader.getInt32("Thang")
                m_HS_ChiLuong.Nam = smartReader.getInt32("Nam")
                m_HS_ChiLuong.IdLoaiThuChi = smartReader.getInt32("IdLoaiThuChi")
                m_HS_ChiLuong.Title = smartReader.getString("Title")
                m_HS_ChiLuong.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_HS_ChiLuong)
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
            strSql = "SELECT * FROM HS_ChiLuong Order by IdDonVi, Nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByDonVi(ByVal vIdDonVi As String) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_chiLuong WHERE IdDonVi='" & vIdDonVi.Trim & "' order by nam desc, thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdHS_ChiLuong As String) As HS_ChiLuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ChiLuong WHERE IdHS_ChiLuong='" & vIdHS_ChiLuong.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), HS_ChiLuong)
            Else
                Return New HS_ChiLuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

Public Class HS_ChiLuong_CT
    Dim _IdHS_ChiLuong_CT As String
    Dim _IdHS_ChiLuong As String
    Dim _Ky As SByte  '  tinyint,
    Dim _IdPhong As String
    Dim _IdCanBo As String
    Dim _NH_TS As Int16
    Dim _IdChucVu As Integer
    Dim _HesoCMNV As Double
    Dim _PC_CV As Double
    Dim _PC_TN As Double
    Dim _PC_KV As Double
    Dim _PC_DH As Double
    Dim _PC_KN As Double
    Dim _PC_TNVK As Double
    Dim _PC_TNN As Double
    Dim _PC_Khac As Double
    Dim _Tong100 As Long
    Dim _TongChi As Long
    Dim _HesoChi As Double
    Dim _T_TNCNtmp As Long
    Dim _T_DFCD As Long
    Dim _T_BHXH As Long
    Dim _T_BHYT As Long
    Dim _T_BHTN As Long
    Dim _T_Nghi As Long
    Dim _T_TruyThu As Long
    Dim _T_TamUng As Long
    Dim _TongTru As Long
    Dim _TruyLinh As Long
    Dim _ThucLinh As Long

    Dim _ThuHut_HesoPC As Double
    Dim _ThuHut_100 As Long
    Dim _ThuHut_80 As Long
    Dim _LuongGio As Long
    Dim _Gio50 As Double
    Dim _Gio100 As Double
    Dim _Gio150 As Double
    Dim _Gio200 As Double
    Dim _Gio300 As Double
    Dim _GioNghiBu As Double
    Dim _TraThemLamDem As Long
    Dim _Gio150_TT As Double
    Dim _Gio200_TT As Double
    Dim _Gio300_TT As Double
    Dim _GioLamDem_NT As Double
    Dim _GioLamDem_NN As Double
    Dim _HeSoCMNV_HH As Double
    Dim _LuongCV As Long
    Dim _LuongCB As Long
    Dim _LuongPC As Long
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdHS_ChiLuong_CT() As String
        Get
            Return _IdHS_ChiLuong_CT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Lương cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiLuong_CT = Value
        End Set
    End Property

    Public Property IdHS_ChiLuong() As String
        Get
            Return _IdHS_ChiLuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Chi Lương có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiLuong = Value
        End Set
    End Property

    Public Property Ky() As Int16
        Get
            Return _Ky
        End Get
        Set(ByVal Value As Int16)
            _Ky = Value
        End Set
    End Property

    Public Property IdPhong() As String
        Get
            Return _IdPhong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 6) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Phòng có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdPhong = Value
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

    Public Property NH_TS() As Int16
        Get
            Return _NH_TS
        End Get
        Set(ByVal Value As Int16)
            _NH_TS = Value
        End Set
    End Property

    Public Property IdChucVu() As Integer
        Get
            Return _IdChucVu
        End Get
        Set(ByVal Value As Integer)
            _IdChucVu = Value
        End Set
    End Property

    Public Property HesoCMNV() As Double
        Get
            Return _HesoCMNV
        End Get
        Set(ByVal Value As Double)
            _HesoCMNV = Value
        End Set
    End Property

    Public Property PC_CV() As Double
        Get
            Return _PC_CV
        End Get
        Set(ByVal Value As Double)
            _PC_CV = Value
        End Set
    End Property

    Public Property PC_TN() As Double
        Get
            Return _PC_TN
        End Get
        Set(ByVal Value As Double)
            _PC_TN = Value
        End Set
    End Property

    Public Property PC_KV() As Double
        Get
            Return _PC_KV
        End Get
        Set(ByVal Value As Double)
            _PC_KV = Value
        End Set
    End Property

    Public Property PC_DH() As Double
        Get
            Return _PC_DH
        End Get
        Set(ByVal Value As Double)
            _PC_DH = Value
        End Set
    End Property

    Public Property PC_KN() As Double
        Get
            Return _PC_KN
        End Get
        Set(ByVal Value As Double)
            _PC_KN = Value
        End Set
    End Property

    Public Property PC_TNVK() As Double
        Get
            Return _PC_TNVK
        End Get
        Set(ByVal Value As Double)
            _PC_TNVK = Value
        End Set
    End Property

    Public Property PC_TNN() As Double
        Get
            Return _PC_TNN
        End Get
        Set(ByVal Value As Double)
            _PC_TNN = Value
        End Set
    End Property

    Public Property PC_Khac() As Double
        Get
            Return _PC_Khac
        End Get
        Set(ByVal Value As Double)
            _PC_Khac = Value
        End Set
    End Property

    Public Property Tong100() As Long
        Get
            Return _Tong100
        End Get
        Set(ByVal Value As Long)
            _Tong100 = Value
        End Set
    End Property

    Public Property TongChi() As Long
        Get
            Return _TongChi
        End Get
        Set(ByVal Value As Long)
            _TongChi = Value
        End Set
    End Property

    Public Property Hesochi() As Double
        Get
            Return _HesoChi
        End Get
        Set(ByVal Value As Double)
            _HesoChi = Value
        End Set
    End Property

    Public Property T_TNCNtmp() As Long
        Get
            Return _T_TNCNtmp
        End Get
        Set(ByVal Value As Long)
            _T_TNCNtmp = Value
        End Set
    End Property

    Public Property T_DFCD() As Long
        Get
            Return _T_DFCD
        End Get
        Set(ByVal Value As Long)
            _T_DFCD = Value
        End Set
    End Property

    Public Property T_BHXH() As Long
        Get
            Return _T_BHXH
        End Get
        Set(ByVal Value As Long)
            _T_BHXH = Value
        End Set
    End Property

    Public Property T_BHYT() As Long
        Get
            Return _T_BHYT
        End Get
        Set(ByVal Value As Long)
            _T_BHYT = Value
        End Set
    End Property

    Public Property T_BHTN() As Long
        Get
            Return _T_BHTN
        End Get
        Set(ByVal Value As Long)
            _T_BHTN = Value
        End Set
    End Property

    Public Property T_Nghi() As Long
        Get
            Return _T_Nghi
        End Get
        Set(ByVal Value As Long)
            _T_Nghi = Value
        End Set
    End Property

    Public Property T_TruyThu() As Long
        Get
            Return _T_TruyThu
        End Get
        Set(ByVal Value As Long)
            _T_TruyThu = Value
        End Set
    End Property

    Public Property T_TamUng() As Long
        Get
            Return _T_TamUng
        End Get
        Set(ByVal Value As Long)
            _T_TamUng = Value
        End Set
    End Property

    Public Property TongTru() As Long
        Get
            Return _TongTru
        End Get
        Set(ByVal Value As Long)
            _TongTru = Value
        End Set
    End Property

    Public Property TruyLinh() As Long
        Get
            Return _TruyLinh
        End Get
        Set(ByVal Value As Long)
            _TruyLinh = Value
        End Set
    End Property

    Public Property ThucLinh() As Long
        Get
            Return _ThucLinh
        End Get
        Set(ByVal Value As Long)
            _ThucLinh = Value
        End Set
    End Property

    Public Property ThuHut_HesoPC() As Double
        Get
            Return _ThuHut_HesoPC
        End Get
        Set(ByVal Value As Double)
            _ThuHut_HesoPC = Value
        End Set
    End Property

    Public Property ThuHut_100() As Long
        Get
            Return _ThuHut_100
        End Get
        Set(ByVal Value As Long)
            _ThuHut_100 = Value
        End Set
    End Property

    Public Property ThuHut_80() As Long
        Get
            Return _ThuHut_80
        End Get
        Set(ByVal Value As Long)
            _ThuHut_80 = Value
        End Set
    End Property

    Public Property LuongGio() As Long
        Get
            Return _LuongGio
        End Get
        Set(ByVal Value As Long)
            _LuongGio = Value
        End Set
    End Property

    Public Property Gio50() As Double
        Get
            Return _Gio50
        End Get
        Set(ByVal Value As Double)
            _Gio50 = Value
        End Set
    End Property

    Public Property Gio100() As Double
        Get
            Return _Gio100
        End Get
        Set(ByVal Value As Double)
            _Gio100 = Value
        End Set
    End Property

    Public Property Gio150() As Double
        Get
            Return _Gio150
        End Get
        Set(ByVal Value As Double)
            _Gio150 = Value
        End Set
    End Property

    Public Property Gio200() As Double
        Get
            Return _Gio200
        End Get
        Set(ByVal Value As Double)
            _Gio200 = Value
        End Set
    End Property

    Public Property Gio300() As Double
        Get
            Return _Gio300
        End Get
        Set(ByVal Value As Double)
            _Gio300 = Value
        End Set
    End Property

    Public Property GioNghiBu() As Double
        Get
            Return _GioNghiBu
        End Get
        Set(ByVal Value As Double)
            _GioNghiBu = Value
        End Set
    End Property

    Public Property TraThemLamDem() As Long
        Get
            Return _TraThemLamDem
        End Get
        Set(ByVal Value As Long)
            _TraThemLamDem = Value
        End Set
    End Property

    Public Property Gio150_TT() As Double
        Get
            Return _Gio150_TT
        End Get
        Set(ByVal Value As Double)
            _Gio150_TT = Value
        End Set
    End Property

    Public Property Gio200_TT() As Double
        Get
            Return _Gio200_TT
        End Get
        Set(ByVal Value As Double)
            _Gio200_TT = Value
        End Set
    End Property

    Public Property Gio300_TT() As Double
        Get
            Return _Gio300_TT
        End Get
        Set(ByVal Value As Double)
            _Gio300_TT = Value
        End Set
    End Property

    Public Property GioLamDem_NT() As Double
        Get
            Return _GioLamDem_NT
        End Get
        Set(ByVal Value As Double)
            _GioLamDem_NT = Value
        End Set
    End Property

    Public Property GioLamDem_NN() As Double
        Get
            Return _GioLamDem_NN
        End Get
        Set(ByVal Value As Double)
            _GioLamDem_NN = Value
        End Set
    End Property

    Public Property HeSoCMNV_HH() As Double
        Get
            Return _HeSoCMNV_HH
        End Get
        Set(ByVal Value As Double)
            _HeSoCMNV_HH = Value
        End Set
    End Property

    Public Property LuongCV() As Long
        Get
            Return _LuongCV
        End Get
        Set(ByVal Value As Long)
            _LuongCV = Value
        End Set
    End Property

    Public Property LuongCB() As Long
        Get
            Return _LuongCB
        End Get
        Set(ByVal Value As Long)
            _LuongCB = Value
        End Set
    End Property

    Public Property LuongPC() As Long
        Get
            Return _LuongPC
        End Get
        Set(ByVal Value As Long)
            _LuongPC = Value
        End Set
    End Property

#End Region

#Region "Method"


    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_CT_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@Ky", Ky))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NH_TS", NH_TS))
            cmd.Parameters.Add(New SqlParameter("@IdChucVu", IdChucVu))
            cmd.Parameters.Add(New SqlParameter("@HesoCMNV", HesoCMNV))
            cmd.Parameters.Add(New SqlParameter("@PC_CV", PC_CV))
            cmd.Parameters.Add(New SqlParameter("@PC_TN", PC_TN))
            cmd.Parameters.Add(New SqlParameter("@PC_KV", PC_KV))
            cmd.Parameters.Add(New SqlParameter("@PC_DH", PC_DH))
            cmd.Parameters.Add(New SqlParameter("@PC_KN", PC_KN))
            cmd.Parameters.Add(New SqlParameter("@PC_TNVK", PC_TNVK))
            cmd.Parameters.Add(New SqlParameter("@PC_TNN", PC_TNN))
            cmd.Parameters.Add(New SqlParameter("@PC_Khac", PC_Khac))
            cmd.Parameters.Add(New SqlParameter("@Tong100", Tong100))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@T_TNCNtmp", T_TNCNtmp))
            cmd.Parameters.Add(New SqlParameter("@T_DFCD", T_DFCD))
            cmd.Parameters.Add(New SqlParameter("@T_BHXH", T_BHXH))
            cmd.Parameters.Add(New SqlParameter("@T_BHYT", T_BHYT))
            cmd.Parameters.Add(New SqlParameter("@T_BHTN", T_BHTN))
            cmd.Parameters.Add(New SqlParameter("@T_Nghi", T_Nghi))
            cmd.Parameters.Add(New SqlParameter("@T_TruyThu", T_TruyThu))
            cmd.Parameters.Add(New SqlParameter("@T_TamUng", T_TamUng))
            cmd.Parameters.Add(New SqlParameter("@TongTru", TongTru))
            cmd.Parameters.Add(New SqlParameter("@TruyLinh", TruyLinh))
            cmd.Parameters.Add(New SqlParameter("@ThucLinh", ThucLinh))

            'Dim _ThuHut_HesoPC As Double
            'Dim _ThuHut_100 As Long
            'Dim _ThuHut_80 As Long
            'Dim _LuongGio As Long
            'Dim _Gio50 As Double
            'Dim _Gio100 As Double
            'Dim _Gio150 As Double
            'Dim _Gio200 As Double
            'Dim _Gio300 As Double
            'Dim _GioNghiBu As Double
            'Dim _TraThemLamDem As Long
            'Dim _Gio150_TT As Double
            'Dim _Gio200_TT As Double
            'Dim _Gio300_TT As Double
            'Dim _GioLamDem_NT As Double
            'Dim _GioLamDem_NN As Double
            'Dim _HeSoCMNV_HH As Double
            'Dim _LuongCV As Long
            'Dim _LuongCB As Long
            'Dim _LuongPC As Long
            cmd.Parameters.Add("@IdHS_ChiLuong_CT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdHS_ChiLuong_CT = cmd.Parameters("@IdHS_ChiBoSung_CT").Value.ToString
        Catch ex As Exception
            IdHS_ChiLuong_CT = ""
        End Try
        Return IdHS_ChiLuong_CT
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_CT_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong_CT", IdHS_ChiLuong_CT))
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@Ky", Ky))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NH_TS", NH_TS))
            cmd.Parameters.Add(New SqlParameter("@IdChucVu", IdChucVu))
            cmd.Parameters.Add(New SqlParameter("@HesoCMNV", HesoCMNV))
            cmd.Parameters.Add(New SqlParameter("@PC_CV", PC_CV))
            cmd.Parameters.Add(New SqlParameter("@PC_TN", PC_TN))
            cmd.Parameters.Add(New SqlParameter("@PC_KV", PC_KV))
            cmd.Parameters.Add(New SqlParameter("@PC_DH", PC_DH))
            cmd.Parameters.Add(New SqlParameter("@PC_KN", PC_KN))
            cmd.Parameters.Add(New SqlParameter("@PC_TNVK", PC_TNVK))
            cmd.Parameters.Add(New SqlParameter("@PC_TNN", PC_TNN))
            cmd.Parameters.Add(New SqlParameter("@PC_Khac", PC_Khac))
            cmd.Parameters.Add(New SqlParameter("@Tong100", Tong100))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            cmd.Parameters.Add(New SqlParameter("@Hesochi", Hesochi))
            cmd.Parameters.Add(New SqlParameter("@T_TNCNtmp", T_TNCNtmp))
            cmd.Parameters.Add(New SqlParameter("@T_DFCD", T_DFCD))
            cmd.Parameters.Add(New SqlParameter("@T_BHXH", T_BHXH))
            cmd.Parameters.Add(New SqlParameter("@T_BHYT", T_BHYT))
            cmd.Parameters.Add(New SqlParameter("@T_BHTN", T_BHTN))
            cmd.Parameters.Add(New SqlParameter("@T_Nghi", T_Nghi))
            cmd.Parameters.Add(New SqlParameter("@T_TruyThu", T_TruyThu))
            cmd.Parameters.Add(New SqlParameter("@T_TamUng", T_TamUng))
            cmd.Parameters.Add(New SqlParameter("@TongTru", TongTru))
            cmd.Parameters.Add(New SqlParameter("@TruyLinh", TruyLinh))
            cmd.Parameters.Add(New SqlParameter("@ThucLinh", ThucLinh))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdatenotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_CT_UpdatenotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong_CT", IdHS_ChiLuong_CT))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@NH_TS", NH_TS))
            cmd.Parameters.Add(New SqlParameter("@Ky", Ky))
            cmd.Parameters.Add(New SqlParameter("@T_TamUng", T_TamUng))
            cmd.Parameters.Add(New SqlParameter("@Tong100", Tong100))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            cmd.Parameters.Add(New SqlParameter("@T_TNCNtmp", T_TNCNtmp))
            cmd.Parameters.Add(New SqlParameter("@T_DFCD", T_DFCD))
            cmd.Parameters.Add(New SqlParameter("@T_BHXH", T_BHXH))
            cmd.Parameters.Add(New SqlParameter("@T_BHYT", T_BHYT))
            cmd.Parameters.Add(New SqlParameter("@T_BHTN", T_BHTN))
            cmd.Parameters.Add(New SqlParameter("@T_Nghi", T_Nghi))
            cmd.Parameters.Add(New SqlParameter("@T_TruyThu", T_TruyThu))
            cmd.Parameters.Add(New SqlParameter("@TruyLinh", TruyLinh))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiLuong_CT_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong_CT", IdHS_ChiLuong_CT))
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
                Dim m_ChiLuong_CT As HS_ChiLuong_CT = New HS_ChiLuong_CT
                m_ChiLuong_CT.IdHS_ChiLuong_CT = smartReader.getString("IdHS_ChiLuong_CT")
                m_ChiLuong_CT.IdHS_ChiLuong = smartReader.getString("IdHS_ChiLuong")
                m_ChiLuong_CT.Ky = smartReader.getInt16("Ky")
                m_ChiLuong_CT.IdPhong = smartReader.getString("IdPhong")
                m_ChiLuong_CT.IdCanBo = smartReader.getString("IdCanBo")
                m_ChiLuong_CT.NH_TS = smartReader.getInt16("NH_TS")
                m_ChiLuong_CT.IdChucVu = smartReader.getInt32("IdChucVu")
                m_ChiLuong_CT.HesoCMNV = smartReader.getFloat("HesoCMNV")
                m_ChiLuong_CT.PC_CV = smartReader.getFloat("PC_CV")
                m_ChiLuong_CT.PC_TN = smartReader.getFloat("PC_TN")
                m_ChiLuong_CT.PC_KV = smartReader.getFloat("PC_KV")
                m_ChiLuong_CT.PC_DH = smartReader.getFloat("PC_DH")
                m_ChiLuong_CT.PC_KN = smartReader.getFloat("PC_KN")
                m_ChiLuong_CT.PC_TNVK = smartReader.getFloat("PC_TNVK")
                m_ChiLuong_CT.PC_TNN = smartReader.getFloat("PC_TNN")
                m_ChiLuong_CT.PC_Khac = smartReader.getFloat("PC_Khac")
                m_ChiLuong_CT.Tong100 = smartReader.getInt64("Tong100")
                m_ChiLuong_CT.TongChi = smartReader.getInt64("TongChi")
                m_ChiLuong_CT.Hesochi = smartReader.getFloat("Hesochi")
                m_ChiLuong_CT.T_TNCNtmp = smartReader.getInt64("T_TNCNtmp")
                m_ChiLuong_CT.T_DFCD = smartReader.getInt64("T_DFCD")
                m_ChiLuong_CT.T_BHXH = smartReader.getInt64("T_BHXH")
                m_ChiLuong_CT.T_BHYT = smartReader.getInt64("T_BHYT")
                m_ChiLuong_CT.T_BHTN = smartReader.getInt64("T_BHTN")
                m_ChiLuong_CT.T_Nghi = smartReader.getInt64("T_Nghi")
                m_ChiLuong_CT.T_TruyThu = smartReader.getInt64("T_TruyThu")
                m_ChiLuong_CT.T_TamUng = smartReader.getInt64("T_TamUng")
                m_ChiLuong_CT.TongTru = smartReader.getInt64("TongTru")
                m_ChiLuong_CT.TruyLinh = smartReader.getInt64("TruyLinh")
                m_ChiLuong_CT.ThucLinh = smartReader.getInt64("ThucLinh")
                list.Add(m_ChiLuong_CT)
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
            strSql = "SELECT * FROM HS_ChiLuong_CT Order by IdPhong, IdChucVu"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllbyIdChiLuong(ByVal vIdChiLuong) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ChiLuong_CT WHERE IdHS_ChiLuong = '" & vIdChiLuong.Trim & "' AND NH_TS=0 Order by IdPhong, IdChucVu"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdHS_ChiLuong_CT As String) As HS_ChiLuong_CT
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ChiLuong_CT WHERE IdHS_ChiLuong_CT='" & vIdHS_ChiLuong_CT.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), HS_ChiLuong_CT)
            Else
                Return New HS_ChiLuong_CT
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class


Public Class HS_ChiThemGio_CT
    Dim _IdHS_ChiThemGio_CT As String
    Dim _IdHS_ChiLuong As String
    Dim _IdCanBo As String
    Dim _IdPhong As String
    Dim _HesoCMNV As Double
    Dim _PC_CV As Double
    Dim _PC_TN As Double
    Dim _PC_KV As Double
    Dim _PC_DH As Double
    Dim _PC_KN As Double
    Dim _PC_TNVK As Double
    Dim _PC_TNN As Double
    Dim _PC_Khac As Double
    Dim _LuongGio As Long
    Dim _Gio150 As Double
    Dim _Gio200 As Double
    Dim _Gio300 As Double
    Dim _LuongLamDem As Long
    Dim _TongGioQD As Double
    Dim _TongChi As Long
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property IdHS_ChiThemGio_CT() As String
        Get
            Return _IdHS_ChiThemGio_CT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Lương thêm giờ cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiThemGio_CT = Value
        End Set
    End Property

    Public Property IdHS_ChiLuong() As String
        Get
            Return _IdHS_ChiLuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Chi Lương có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdHS_ChiLuong = Value
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

    Public Property IdPhong() As String
        Get
            Return _IdPhong
        End Get
        Set(ByVal Value As String)
            _IdPhong = Value
        End Set
    End Property

    Public Property HesoCMNV() As Double
        Get
            Return _HesoCMNV
        End Get
        Set(ByVal Value As Double)
            _HesoCMNV = Value
        End Set
    End Property

    Public Property PC_CV() As Double
        Get
            Return _PC_CV
        End Get
        Set(ByVal Value As Double)
            _PC_CV = Value
        End Set
    End Property

    Public Property PC_TN() As Double
        Get
            Return _PC_TN
        End Get
        Set(ByVal Value As Double)
            _PC_TN = Value
        End Set
    End Property

    Public Property PC_KV() As Double
        Get
            Return _PC_KV
        End Get
        Set(ByVal Value As Double)
            _PC_KV = Value
        End Set
    End Property

    Public Property PC_DH() As Double
        Get
            Return _PC_DH
        End Get
        Set(ByVal Value As Double)
            _PC_DH = Value
        End Set
    End Property

    Public Property PC_KN() As Double
        Get
            Return _PC_KN
        End Get
        Set(ByVal Value As Double)
            _PC_KN = Value
        End Set
    End Property

    Public Property PC_TNVK() As Double
        Get
            Return _PC_TNVK
        End Get
        Set(ByVal Value As Double)
            _PC_TNVK = Value
        End Set
    End Property

    Public Property PC_TNN() As Double
        Get
            Return _PC_TNN
        End Get
        Set(ByVal Value As Double)
            _PC_TNN = Value
        End Set
    End Property

    Public Property PC_Khac() As Double
        Get
            Return _PC_Khac
        End Get
        Set(ByVal Value As Double)
            _PC_Khac = Value
        End Set
    End Property

    Public Property LuongGio() As Long
        Get
            Return _LuongGio
        End Get
        Set(ByVal Value As Long)
            _LuongGio = Value
        End Set
    End Property

    Public Property Gio150() As Double
        Get
            Return _Gio150
        End Get
        Set(ByVal Value As Double)
            _Gio150 = Value
        End Set
    End Property

    Public Property Gio200() As Double
        Get
            Return _Gio200
        End Get
        Set(ByVal Value As Double)
            _Gio200 = Value
        End Set
    End Property

    Public Property Gio300() As Double
        Get
            Return _Gio300
        End Get
        Set(ByVal Value As Double)
            _Gio300 = Value
        End Set
    End Property

    Public Property LuongLamDem() As Long
        Get
            Return _LuongLamDem
        End Get
        Set(ByVal Value As Long)
            _LuongLamDem = Value
        End Set
    End Property

    Public Property TongGioQD() As Double
        Get
            Return _TongGioQD
        End Get
        Set(ByVal Value As Double)
            _TongGioQD = Value
        End Set
    End Property

    Public Property TongChi() As Long
        Get
            Return _TongChi
        End Get
        Set(ByVal Value As Long)
            _TongChi = Value
        End Set
    End Property

#End Region

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiThemGio_CT_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@HesoCMNV", HesoCMNV))
            cmd.Parameters.Add(New SqlParameter("@PC_CV", PC_CV))
            cmd.Parameters.Add(New SqlParameter("@PC_TN", PC_TN))
            cmd.Parameters.Add(New SqlParameter("@PC_KV", PC_KV))
            cmd.Parameters.Add(New SqlParameter("@PC_DH", PC_DH))
            cmd.Parameters.Add(New SqlParameter("@PC_KN", PC_KN))
            cmd.Parameters.Add(New SqlParameter("@PC_TNVK", PC_TNVK))
            cmd.Parameters.Add(New SqlParameter("@PC_TNN", PC_TNN))
            cmd.Parameters.Add(New SqlParameter("@PC_Khac", PC_Khac))
            cmd.Parameters.Add(New SqlParameter("@LuongGio", LuongGio))
            cmd.Parameters.Add(New SqlParameter("@Gio150", Gio150))
            cmd.Parameters.Add(New SqlParameter("@Gio200", Gio200))
            cmd.Parameters.Add(New SqlParameter("@Gio300", Gio300))
            cmd.Parameters.Add(New SqlParameter("@LuongLamDem", LuongLamDem))
            cmd.Parameters.Add(New SqlParameter("@TongGioQD", TongGioQD))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            cmd.Parameters.Add("@IdHS_ChiThemGio_CT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdHS_ChiThemGio_CT = cmd.Parameters("@IdHS_ChiThemGio_CT").Value.ToString
        Catch ex As Exception
            IdHS_ChiThemGio_CT = ""
        End Try
        Return IdHS_ChiThemGio_CT
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiThemGio_CT_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiThemGio_CT", IdHS_ChiThemGio_CT))
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", IdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@HesoCMNV", HesoCMNV))
            cmd.Parameters.Add(New SqlParameter("@PC_CV", PC_CV))
            cmd.Parameters.Add(New SqlParameter("@PC_TN", PC_TN))
            cmd.Parameters.Add(New SqlParameter("@PC_KV", PC_KV))
            cmd.Parameters.Add(New SqlParameter("@PC_DH", PC_DH))
            cmd.Parameters.Add(New SqlParameter("@PC_KN", PC_KN))
            cmd.Parameters.Add(New SqlParameter("@PC_TNVK", PC_TNVK))
            cmd.Parameters.Add(New SqlParameter("@PC_TNN", PC_TNN))
            cmd.Parameters.Add(New SqlParameter("@PC_Khac", PC_Khac))
            cmd.Parameters.Add(New SqlParameter("@LuongGio", LuongGio))
            cmd.Parameters.Add(New SqlParameter("@Gio150", Gio150))
            cmd.Parameters.Add(New SqlParameter("@Gio200", Gio200))
            cmd.Parameters.Add(New SqlParameter("@Gio300", Gio300))
            cmd.Parameters.Add(New SqlParameter("@LuongLamDem", LuongLamDem))
            cmd.Parameters.Add(New SqlParameter("@TongGioQD", TongGioQD))
            cmd.Parameters.Add(New SqlParameter("@TongChi", TongChi))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdatenotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiThemGio_CT_UpdatenotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiThemGio_CT", IdHS_ChiThemGio_CT))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@Gio150", Gio150))
            cmd.Parameters.Add(New SqlParameter("@Gio200", Gio200))
            cmd.Parameters.Add(New SqlParameter("@Gio300", Gio300))
            cmd.Parameters.Add(New SqlParameter("@LuongLamDem", LuongLamDem))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_ChiThemGio_CT_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiThemGio_CT", IdHS_ChiThemGio_CT))
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
                Dim m_ChiThemGio_CT As HS_ChiThemGio_CT = New HS_ChiThemGio_CT
                m_ChiThemGio_CT.IdHS_ChiThemGio_CT = smartReader.getString("IdHS_ChiThemGio_CT")
                m_ChiThemGio_CT.IdHS_ChiLuong = smartReader.getString("IdHS_ChiLuong")
                m_ChiThemGio_CT.IdCanBo = smartReader.getString("IdCanBo")
                m_ChiThemGio_CT.IdPhong = smartReader.getString("IdPhong")
                m_ChiThemGio_CT.HesoCMNV = smartReader.getFloat("HesoCMNV")
                m_ChiThemGio_CT.PC_CV = smartReader.getFloat("PC_CV")
                m_ChiThemGio_CT.PC_TN = smartReader.getFloat("PC_TN")
                m_ChiThemGio_CT.PC_KV = smartReader.getFloat("PC_KV")
                m_ChiThemGio_CT.PC_DH = smartReader.getFloat("PC_DH")
                m_ChiThemGio_CT.PC_KN = smartReader.getFloat("PC_KN")
                m_ChiThemGio_CT.PC_TNVK = smartReader.getFloat("PC_TNVK")
                m_ChiThemGio_CT.PC_TNN = smartReader.getFloat("PC_TNN")
                m_ChiThemGio_CT.PC_Khac = smartReader.getFloat("PC_Khac")
                m_ChiThemGio_CT.LuongGio = smartReader.getInt64("LuongGio")
                m_ChiThemGio_CT.Gio150 = smartReader.getFloat("Gio150")
                m_ChiThemGio_CT.Gio200 = smartReader.getFloat("Gio200")
                m_ChiThemGio_CT.Gio300 = smartReader.getFloat("Gio300")
                m_ChiThemGio_CT.TongGioQD = smartReader.getFloat("TongGioQD")
                m_ChiThemGio_CT.LuongLamDem = smartReader.getInt64("LuongLamDem")
                m_ChiThemGio_CT.TongChi = smartReader.getInt64("TongChi")
                list.Add(m_ChiThemGio_CT)
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
            strSql = "SELECT * FROM HS_ChiThemGio_CT Order by IdPhong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdHS_ChiThemGio_CT As String) As HS_ChiThemGio_CT
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_ChiThemGio_CT WHERE IdHS_ChiThemGio_CT='" & vIdHS_ChiThemGio_CT.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), HS_ChiThemGio_CT)
            Else
                Return New HS_ChiThemGio_CT
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region
End Class


Public Class tbBC05
    Private _Id As String
    Private _IdDonVi As Integer
    Private _All As Boolean
    Private _Nam As Integer
    Private _TB_LDNH As Integer
    Private _TB_LDDH As Integer
    Private _TH_LDNH As Integer
    Private _TH_LDDH As Integer
    Private _L_ChucVu As Long
    Private _L_CapBac As Long
    Private _L_PhuCap As Long
    Private _L_TronGoi As Long
    Private _TN_BHXH As Long
    Private _TN_ThemGio As Long
    Private _TN_BoSung As Long
    Private _DBtmp As Int16
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Property"

    Public Property Id() As String
        Get
            Return _Id
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID BC05 có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _Id = Value
        End Set
    End Property


    Public Property IdDonVi() As Integer
        Get
            Return _IdDonVi
        End Get
        Set(ByVal Value As Integer)
            _IdDonVi = Value
        End Set
    End Property

    Public Property All() As Boolean
        Get
            Return _All
        End Get
        Set(ByVal Value As Boolean)
            _All = Value
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

    Public Property TB_LDNH() As Integer
        Get
            Return _TB_LDNH
        End Get
        Set(ByVal Value As Integer)
            _TB_LDNH = Value
        End Set
    End Property

    Public Property TB_LDDH() As Integer
        Get
            Return _TB_LDDH
        End Get
        Set(ByVal Value As Integer)
            _TB_LDDH = Value
        End Set
    End Property

    Public Property TH_LDNH() As Integer
        Get
            Return _TH_LDNH
        End Get
        Set(ByVal Value As Integer)
            _TH_LDNH = Value
        End Set
    End Property

    Public Property TH_LDDH() As Integer
        Get
            Return _TH_LDDH
        End Get
        Set(ByVal Value As Integer)
            _TH_LDDH = Value
        End Set
    End Property
   
    Public Property L_ChucVu() As Long
        Get
            Return _L_ChucVu
        End Get
        Set(ByVal Value As Long)
            _L_ChucVu = Value
        End Set
    End Property

    Public Property L_CapBac() As Long
        Get
            Return _L_CapBac
        End Get
        Set(ByVal Value As Long)
            _L_CapBac = Value
        End Set
    End Property

    Public Property L_PhuCap() As Long
        Get
            Return _L_PhuCap
        End Get
        Set(ByVal Value As Long)
            _L_PhuCap = Value
        End Set
    End Property

    Public Property L_TronGoi() As Long
        Get
            Return _L_TronGoi
        End Get
        Set(ByVal Value As Long)
            _L_TronGoi = Value
        End Set
    End Property

    Public Property TN_BHXH() As Long
        Get
            Return _TN_BHXH
        End Get
        Set(ByVal Value As Long)
            _TN_BHXH = Value
        End Set
    End Property

    Public Property TN_ThemGio() As Long
        Get
            Return _TN_ThemGio
        End Get
        Set(ByVal Value As Long)
            _TN_ThemGio = Value
        End Set
    End Property

    Public Property TN_BoSung() As Long
        Get
            Return _TN_BoSung
        End Get
        Set(ByVal Value As Long)
            _TN_BoSung = Value
        End Set
    End Property

    Public Property DBtmp() As Int16
        Get
            Return _DBtmp
        End Get
        Set(ByVal Value As Int16)
            _DBtmp = Value
        End Set
    End Property

#End Region

#Region "method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("tbBC05_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi))
            cmd.Parameters.Add(New SqlParameter("@All", All))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@TB_LDNH", TB_LDNH))
            cmd.Parameters.Add(New SqlParameter("@TB_LDDH", TB_LDDH))
            cmd.Parameters.Add(New SqlParameter("@TH_LDNH", TH_LDNH))
            cmd.Parameters.Add(New SqlParameter("@TH_LDDH", TH_LDDH))
            cmd.Parameters.Add(New SqlParameter("@L_ChucVu", L_ChucVu))
            cmd.Parameters.Add(New SqlParameter("@L_CapBac", L_CapBac))
            cmd.Parameters.Add(New SqlParameter("@L_PhuCap", L_PhuCap))
            cmd.Parameters.Add(New SqlParameter("@L_TronGoi", L_TronGoi))
            cmd.Parameters.Add(New SqlParameter("@TN_BHXH", TN_BHXH))
            cmd.Parameters.Add(New SqlParameter("@TN_ThemGio", TN_ThemGio))
            cmd.Parameters.Add(New SqlParameter("@TN_BoSung", TN_BoSung))
            cmd.Parameters.Add(New SqlParameter("@DBtmp", DBtmp))
            cmd.Parameters.Add("@Id", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            Id = cmd.Parameters("@Id").Value.ToString
        Catch ex As Exception
            Id = ""
        End Try
        Return Id
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("tbBC05_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@Id", Id))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi))
            cmd.Parameters.Add(New SqlParameter("@All", All))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@TB_LDNH", TB_LDNH))
            cmd.Parameters.Add(New SqlParameter("@TB_LDDH", TB_LDDH))
            cmd.Parameters.Add(New SqlParameter("@TH_LDNH", TH_LDNH))
            cmd.Parameters.Add(New SqlParameter("@TH_LDDH", TH_LDDH))
            cmd.Parameters.Add(New SqlParameter("@L_ChucVu", L_ChucVu))
            cmd.Parameters.Add(New SqlParameter("@L_CapBac", L_CapBac))
            cmd.Parameters.Add(New SqlParameter("@L_PhuCap", L_PhuCap))
            cmd.Parameters.Add(New SqlParameter("@L_TronGoi", L_TronGoi))
            cmd.Parameters.Add(New SqlParameter("@TN_BHXH", TN_BHXH))
            cmd.Parameters.Add(New SqlParameter("@TN_ThemGio", TN_ThemGio))
            cmd.Parameters.Add(New SqlParameter("@TN_BoSung", TN_BoSung))
            cmd.Parameters.Add(New SqlParameter("@DBtmp", DBtmp))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("tbBC05_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@Id", Id))
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
                Dim m_BC05 As tbBC05 = New tbBC05
                m_BC05.Id = smartReader.getString("Id")
                m_BC05.IdDonVi = smartReader.getInt32("IdDonVi")
                m_BC05.All = smartReader.getBoolean("All")
                m_BC05.Nam = smartReader.getInt32("Nam")
                m_BC05.TB_LDNH = smartReader.getInt32("TB_LDNH")
                m_BC05.TB_LDDH = smartReader.getInt32("TB_LDDH")
                m_BC05.TH_LDNH = smartReader.getInt32("TH_LDNH")
                m_BC05.TH_LDDH = smartReader.getInt32("TH_LDDH")
                m_BC05.L_ChucVu = smartReader.getInt64("L_ChucVu")
                m_BC05.L_CapBac = smartReader.getInt64("L_CapBac")
                m_BC05.L_PhuCap = smartReader.getInt64("L_PhuCap")
                m_BC05.L_TronGoi = smartReader.getInt64("L_TronGoi")
                m_BC05.TN_BHXH = smartReader.getInt64("TN_BHXH")
                m_BC05.TN_ThemGio = smartReader.getInt64("TN_ThemGio")
                m_BC05.TN_BoSung = smartReader.getInt64("TN_BoSung")
                m_BC05.DBtmp = smartReader.getInt16("DBtmp")
                list.Add(m_BC05)
            End While
            smartReader.disposeReader(reader)
            Return list
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            db.closeConnection(conn)
        End Try

    End Function

    Public Function getRecord(ByVal vId As String) As tbBC05
        Try
            Dim strSql As String
            strSql = "SELECT * FROM tbBC05 WHERE Id='" & vId.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), tbBC05)
            Else
                Return New tbBC05
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getID(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short) As String
        Try
            Dim strSql As String
            strSql = "SELECT [ID] FROM tbBC05 WHERE IdDonVi=" & vIdDonVi & " And [All]=" & vAll & " And Nam=" & vYear
            Return db.getString(strSql)
        Catch ex As Exception
            Return ""
        End Try
    End Function

#End Region

End Class