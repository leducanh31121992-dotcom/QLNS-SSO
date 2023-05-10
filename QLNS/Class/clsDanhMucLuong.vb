Imports System
Imports System.Data
Imports System.Data.SqlClient
''' <summary>
''' Author: Nguyễn Thị Thuỳ Giang
''' </summary>
''' <remarks></remarks>
Public Class TienLuong
    Dim _IdTienLuong As Integer
    Dim _LuongCoBan As Decimal
    Dim _HeSoNganh As Double
    Dim _NgayHuong As Date
    Dim _SoQD As String
    Dim _NgayQD As Date
    Dim _NguoiQD As String
    Dim _IdCV_Nguoi_QD As Integer
    Dim _ThongTinChung As String
    Dim _Ghichu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdTienLuong() As Integer
        Get
            Return _IdTienLuong
        End Get
        Set(ByVal Value As Integer)
            _IdTienLuong = Value
        End Set
    End Property

    Public Property LuongCoBan() As Decimal
        Get
            Return _LuongCoBan
        End Get
        Set(ByVal Value As Decimal)
            _LuongCoBan = Value
        End Set
    End Property

    Public Property HeSoNganh() As Double
        Get
            Return _HeSoNganh
        End Get
        Set(ByVal Value As Double)
            _HeSoNganh = Value
        End Set
    End Property

    Public Property NgayHuong() As Date
        Get
            Return _NgayHuong
        End Get
        Set(ByVal Value As Date)
            _NgayHuong = Value
        End Set
    End Property

    Public Property SoQD() As String
        Get
            Return _SoQD
        End Get
        Set(ByVal Value As String)
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

    Public Property ThongTinChung() As String
        Get
            Return _ThongTinChung
        End Get
        Set(ByVal Value As String)
            _ThongTinChung = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _Ghichu
        End Get
        Set(ByVal Value As String)
            _Ghichu = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("TienLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@LuongCoBan", LuongCoBan))
            cmd.Parameters.Add(New SqlParameter("@HeSoNganh", HeSoNganh))
            cmd.Parameters.Add(New SqlParameter("@NgayHuong", NgayHuong))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdTienLuong", SqlDbType.Int).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdTienLuong = cmd.Parameters("@IdTienLuong").Value.ToString
        Catch ex As Exception
            IdTienLuong = -1
        End Try
        Return IdTienLuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("TienLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdTienLuong", IdTienLuong))
            cmd.Parameters.Add(New SqlParameter("@LuongCoBan", LuongCoBan))
            cmd.Parameters.Add(New SqlParameter("@HeSoNganh", HeSoNganh))
            cmd.Parameters.Add(New SqlParameter("@NgayHuong", NgayHuong))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("TienLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdTienLuong", IdTienLuong))
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
                Dim m_TienLuong As TienLuong = New TienLuong
                m_TienLuong.IdTienLuong = smartReader.getInt32("IdTienLuong")
                m_TienLuong.LuongCoBan = smartReader.getDecimal("LuongCoBan")
                m_TienLuong.HeSoNganh = smartReader.getFloat("HeSoNganh")
                m_TienLuong.NgayHuong = smartReader.getDatetime("NgayHuong")
                m_TienLuong.SoQD = smartReader.getString("SoQD")
                m_TienLuong.NgayQD = smartReader.getDatetime("NgayQD")
                m_TienLuong.NguoiQD = smartReader.getString("NguoiQD")
                m_TienLuong.IdCV_Nguoi_QD = smartReader.getInt32("IdCV_Nguoi_QD")
                m_TienLuong.ThongTinChung = "Lương tối thiểu vùng: " & formatMoney(m_TienLuong.LuongCoBan.ToString) & "; Hệ số ngành: " & m_TienLuong.HeSoNganh.ToString & "; Ngày: " & DateTimeUtil.getShortDate(m_TienLuong.NgayHuong)
                m_TienLuong.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_TienLuong)
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
            strSql = "SELECT * FROM TienLuong Order by NgayHuong desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllcurr() As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM TienLuong WHERE status=1 Order by NgayHuong desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdTienLuong As Integer) As TienLuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM TienLuong WHERE IdTienLuong='" & vIdTienLuong.ToString & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), TienLuong)
            Else
                Return New TienLuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

Public Class NghiDinhLuong
    Dim _IdNDLuong As Integer
    Dim _MaND As String
    Dim _SoND As String
    Dim _TenND As String
    Dim _LoaiND As String
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdNDLuong() As Integer
        Get
            Return _IdNDLuong
        End Get
        Set(ByVal Value As Integer)
            _IdNDLuong = Value
        End Set
    End Property

    Public Property MaND() As String
        Get
            Return _MaND
        End Get
        Set(ByVal Value As String)
            _MaND = Value
        End Set
    End Property

    Public Property SoND() As String
        Get
            Return _SoND
        End Get
        Set(ByVal Value As String)
            _SoND = Value
        End Set
    End Property

    Public Property TenND() As String
        Get
            Return _TenND
        End Get
        Set(ByVal Value As String)
            _TenND = Value
        End Set
    End Property

    Public Property LoaiND() As String
        Get
            Return _LoaiND
        End Get
        Set(ByVal Value As String)
            _LoaiND = Value
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

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("NghiDinhLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@MaND", MaND))
            cmd.Parameters.Add(New SqlParameter("@SoND", SoND))
            cmd.Parameters.Add(New SqlParameter("@TenND", TenND))
            cmd.Parameters.Add(New SqlParameter("@LoaiND", LoaiND))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdNDLuong", SqlDbType.Int).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdNDLuong = cmd.Parameters("@IdNDLuong").Value.ToString
        Catch ex As Exception
            IdNDLuong = -1
        End Try
        Return IdNDLuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("NghiDinhLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNDLuong", IdNDLuong))
            cmd.Parameters.Add(New SqlParameter("@MaND", MaND))
            cmd.Parameters.Add(New SqlParameter("@SoND", SoND))
            cmd.Parameters.Add(New SqlParameter("@TenND", TenND))
            cmd.Parameters.Add(New SqlParameter("@LoaiND", LoaiND))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("NghiDinhLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNDLuong", IdNDLuong))
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
                Dim m_NghiDinhLuong As NghiDinhLuong = New NghiDinhLuong
                m_NghiDinhLuong.IdNDLuong = smartReader.getInt32("IdNDLuong")
                m_NghiDinhLuong.MaND = smartReader.getString("MaND")
                m_NghiDinhLuong.SoND = smartReader.getString("SoND")
                m_NghiDinhLuong.TenND = smartReader.getString("TenND")
                m_NghiDinhLuong.LoaiND = smartReader.getString("LoaiND")
                m_NghiDinhLuong.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_NghiDinhLuong)
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
            strSql = "SELECT * FROM NghiDinhLuong Order by MaND"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdNDLuong As Integer) As NghiDinhLuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM NghiDinhLuong WHERE IdNDLuong='" & vIdNDLuong & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), NghiDinhLuong)
            Else
                Return New NghiDinhLuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class BangLuong
    Dim _IdBangLuong As Integer
    Dim _IdND_Luong As Integer
    Dim _BangLuong As String
    Dim _MoTa As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdBangLuong() As Integer
        Get
            Return _IdBangLuong
        End Get
        Set(ByVal Value As Integer)
            _IdBangLuong = Value
        End Set
    End Property

    Public Property IdND_Luong() As Integer
        Get
            Return _IdND_Luong
        End Get
        Set(ByVal Value As Integer)
            _IdND_Luong = Value
        End Set
    End Property

    Public Property BangLuong() As String
        Get
            Return _BangLuong
        End Get
        Set(ByVal Value As String)
            _BangLuong = Value
        End Set
    End Property

    Public Property MoTa() As String
        Get
            Return _MoTa
        End Get
        Set(ByVal Value As String)
            _MoTa = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("BangLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdND_Luong", IdND_Luong))
            cmd.Parameters.Add(New SqlParameter("@BangLuong", BangLuong))
            cmd.Parameters.Add(New SqlParameter("@MoTa", MoTa))
            cmd.Parameters.Add("@IdBangLuong", SqlDbType.Int).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdBangLuong = cmd.Parameters("@IdBangLuong").Value.ToString
        Catch ex As Exception
            IdBangLuong = -1
        End Try
        Return IdBangLuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("BangLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBangLuong", IdBangLuong))
            cmd.Parameters.Add(New SqlParameter("@IdND_Luong", IdND_Luong))
            cmd.Parameters.Add(New SqlParameter("@BangLuong", BangLuong))
            cmd.Parameters.Add(New SqlParameter("@MoTa", MoTa))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("BangLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBangLuong", IdBangLuong))
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
                Dim m_BangLuong As BangLuong = New BangLuong
                m_BangLuong.IdBangLuong = smartReader.getInt32("IdBangLuong")
                m_BangLuong.IdND_Luong = smartReader.getInt32("IdND_Luong")
                m_BangLuong.BangLuong = smartReader.getString("BangLuong")
                m_BangLuong.MoTa = smartReader.getString("MoTa")
                list.Add(m_BangLuong)
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
            strSql = "SELECT * FROM BangLuong Where  Status = 1 Order by IdND_Luong, BangLuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByNghiDinhLuong(ByVal IdNghiDinhLuong As Integer) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM BangLuong WHERE IdND_Luong= " & IdNghiDinhLuong & " Order by BangLuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdBangLuong As Integer) As BangLuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM BangLuong WHERE IdBangLuong='" & vIdBangLuong & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), BangLuong)
            Else
                Return New BangLuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class NgachLuong
    Dim _IdNgachLuong As Integer
    Dim _IdBangLuong As Integer
    Dim _NgachLuong As String
    Dim _Loai As String
    Dim _Mota As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdNgachLuong() As Integer
        Get
            Return _IdNgachLuong
        End Get
        Set(ByVal Value As Integer)
            _IdNgachLuong = Value
        End Set
    End Property

    Public Property IdBangLuong() As Integer
        Get
            Return _IdBangLuong
        End Get
        Set(ByVal Value As Integer)
            _IdBangLuong = Value
        End Set
    End Property

    Public Property NgachLuong() As String
        Get
            Return _NgachLuong
        End Get
        Set(ByVal Value As String)
            _NgachLuong = Value
        End Set
    End Property

    Public Property Loai() As String
        Get
            Return _Loai
        End Get
        Set(ByVal Value As String)
            _Loai = Value
        End Set
    End Property

    Public Property MoTa() As String
        Get
            Return _MoTa
        End Get
        Set(ByVal Value As String)
            _MoTa = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("NgachLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBangLuong", IdBangLuong))
            cmd.Parameters.Add(New SqlParameter("@NgachLuong", NgachLuong))
            cmd.Parameters.Add(New SqlParameter("@Loai", Loai))
            cmd.Parameters.Add(New SqlParameter("@MoTa", MoTa))
            cmd.Parameters.Add("@IdNgachLuong", SqlDbType.Int).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdNgachLuong = cmd.Parameters("@IdNgachLuong").Value.ToString
        Catch ex As Exception
            IdNgachLuong = -1
        End Try
        Return IdNgachLuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("NgachLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNgachLuong", IdNgachLuong))
            cmd.Parameters.Add(New SqlParameter("@IdBangLuong", IdBangLuong))
            cmd.Parameters.Add(New SqlParameter("@NgachLuong", NgachLuong))
            cmd.Parameters.Add(New SqlParameter("@Loai", Loai))
            cmd.Parameters.Add(New SqlParameter("@MoTa", MoTa))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("NgachLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNgachLuong", IdNgachLuong))
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
                Dim m_NgachLuong As NgachLuong = New NgachLuong
                m_NgachLuong.IdNgachLuong = smartReader.getInt32("IdNgachLuong")
                m_NgachLuong.IdBangLuong = smartReader.getInt32("IdBangLuong")
                m_NgachLuong.NgachLuong = smartReader.getString("NgachLuong")
                m_NgachLuong.Loai = smartReader.getString("Loai")
                m_NgachLuong.MoTa = smartReader.getString("MoTa")
                list.Add(m_NgachLuong)
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
            strSql = "SELECT * FROM NgachLuong Order by IdBangLuong, NgachLuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByBangLuong(ByVal IdBangLuong As Integer) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM NgachLuong WHERE IdBangLuong = " & IdBangLuong & " Order by NgachLuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdNgachLuong As Integer) As NgachLuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM NgachLuong WHERE IdNgachLuong='" & vIdNgachLuong & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), NgachLuong)
            Else
                Return New NgachLuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class BacLuong
    Dim _IdBacLuong As Integer
    Dim _IdNgachLuong As Integer
    Dim _BacLuong As Integer
    Dim _Heso As Double
    Dim _Mota As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdBacLuong() As Integer
        Get
            Return _IdBacLuong
        End Get
        Set(ByVal Value As Integer)
            _IdBacLuong = Value
        End Set
    End Property

    Public Property IdNgachLuong() As Integer
        Get
            Return _IdNgachLuong
        End Get
        Set(ByVal Value As Integer)
            _IdNgachLuong = Value
        End Set
    End Property

    Public Property BacLuong() As Integer
        Get
            Return _BacLuong
        End Get
        Set(ByVal Value As Integer)
            _BacLuong = Value
        End Set
    End Property

    Public Property Heso() As Double
        Get
            Return _Heso
        End Get
        Set(ByVal Value As Double)
            _Heso = Value
        End Set
    End Property

    Public Property MoTa() As String
        Get
            Return _Mota
        End Get
        Set(ByVal Value As String)
            _Mota = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("BacLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdNgachLuong", IdNgachLuong))
            cmd.Parameters.Add(New SqlParameter("@BacLuong", BacLuong))
            cmd.Parameters.Add(New SqlParameter("@Heso", Heso))
            cmd.Parameters.Add(New SqlParameter("@MoTa", MoTa))
            cmd.Parameters.Add("@IdBacLuong", SqlDbType.Int).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdBacLuong = cmd.Parameters("@IdBacLuong").Value.ToString
        Catch ex As Exception
            IdBacLuong = -1
        End Try
        Return IdBacLuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("BacLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBacLuong", IdBacLuong))
            cmd.Parameters.Add(New SqlParameter("@IdNgachLuong", IdNgachLuong))
            cmd.Parameters.Add(New SqlParameter("@BacLuong", BacLuong))
            cmd.Parameters.Add(New SqlParameter("@Heso", Heso))
            cmd.Parameters.Add(New SqlParameter("@MoTa", MoTa))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("BacLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBacLuong", IdBacLuong))
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
                Dim m_BacLuong As BacLuong = New BacLuong
                m_BacLuong.IdBacLuong = smartReader.getInt32("IdBacLuong")
                m_BacLuong.IdNgachLuong = smartReader.getInt32("IdNgachLuong")
                m_BacLuong.BacLuong = smartReader.getInt32("BacLuong")
                m_BacLuong.Heso = smartReader.getFloat("Heso")
                m_BacLuong.MoTa = smartReader.getString("MoTa")
                list.Add(m_BacLuong)
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
            strSql = "SELECT * FROM BacLuong Order by IdNgachLuong, BacLuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByNgachLuong(ByVal IdNgachLuong As Integer) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM BacLuong WHERE IdNgachLuong = " & IdNgachLuong & " Order by BacLuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdBacLuong As Integer) As BacLuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM BacLuong WHERE IdBacLuong='" & vIdBacLuong & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), BacLuong)
            Else
                Return New BacLuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class MucPhuCap
    Dim _IdMuc_PhC As Integer
    Dim _IdLoai_PhC As Integer
    Dim _Muc_PhC As Double

    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdMuc_PhC() As Integer
        Get
            Return _IdMuc_PhC
        End Get
        Set(ByVal Value As Integer)
            _IdMuc_PhC = Value
        End Set
    End Property

    Public Property IdLoai_PhC() As Integer
        Get
            Return _IdLoai_PhC
        End Get
        Set(ByVal Value As Integer)
            _IdLoai_PhC = Value
        End Set
    End Property

    Public Property Muc_PhC() As Double
        Get
            Return _Muc_PhC
        End Get
        Set(ByVal Value As Double)
            _Muc_PhC = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("MucPhuCap_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdLoai_PhC", IdLoai_PhC))
            cmd.Parameters.Add(New SqlParameter("@Muc_PhC", Muc_PhC))
            cmd.Parameters.Add("@IdMuc_PhC", SqlDbType.Int).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdMuc_PhC = cmd.Parameters("@IdMuc_PhC").Value.ToString
        Catch ex As Exception
            IdMuc_PhC = -1
        End Try
        Return IdMuc_PhC
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("MucPhuCap_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdMuc_PhC", IdMuc_PhC))
            cmd.Parameters.Add(New SqlParameter("@IdLoai_PhC", IdLoai_PhC))
            cmd.Parameters.Add(New SqlParameter("@Muc_PhC", Muc_PhC))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("MucPhuCap_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdMuc_PhC", IdMuc_PhC))
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
                Dim m_MucPhuCap As MucPhuCap = New MucPhuCap
                m_MucPhuCap.IdMuc_PhC = smartReader.getInt32("IdMuc_PhC")
                m_MucPhuCap.IdLoai_PhC = smartReader.getInt32("IdLoai_PhC")
                m_MucPhuCap.Muc_PhC = smartReader.getFloat("Muc_PhC")
                list.Add(m_MucPhuCap)
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
            strSql = "SELECT * FROM MucPhuCap Order by IdLoai_PhC, Muc_PhC"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdMucPhuCap As Integer) As MucPhuCap
        Try
            Dim strSql As String
            strSql = "SELECT * FROM MucPhuCap WHERE IdMuc_PhC='" & vIdMucPhuCap & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), MucPhuCap)
            Else
                Return New MucPhuCap
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

Public Class LuongDonVi
    Dim _id As String
    Dim _idChiNhanh As Integer
    Dim _idTienLuong As Integer
    Dim _TuNgay As Date
    Dim _SoQD As String
    Dim _NgayQD As Date
    Dim _NguoiQD As String
    Dim _IdCV_Nguoi_QD As Integer
    Dim _Ghichu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property id() As String
        Get
            Return _id
        End Get
        Set(ByVal Value As String)
            _id = Value
        End Set
    End Property

    Public Property idChiNhanh() As Integer
        Get
            Return _idChiNhanh
        End Get
        Set(ByVal Value As Integer)
            _idChiNhanh = Value
        End Set
    End Property

    Public Property idTienLuong() As Integer
        Get
            Return _idTienLuong
        End Get
        Set(ByVal Value As Integer)
            _idTienLuong = Value
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

    Public Property SoQD() As String
        Get
            Return _SoQD
        End Get
        Set(ByVal Value As String)
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

    Public Property GhiChu() As String
        Get
            Return _Ghichu
        End Get
        Set(ByVal Value As String)
            _Ghichu = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("CN_TienLuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@idChiNhanh", idChiNhanh))
            cmd.Parameters.Add(New SqlParameter("@idTienLuong", idTienLuong))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@id", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            id = cmd.Parameters("@id").Value.ToString
        Catch ex As Exception
            id = ""
        End Try
        Return id
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("CN_TienLuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@id", id))
            cmd.Parameters.Add(New SqlParameter("@idChiNhanh", idChiNhanh))
            cmd.Parameters.Add(New SqlParameter("@idTienLuong", idTienLuong))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCV_Nguoi_QD", IdCV_Nguoi_QD))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("CN_TienLuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@id", id))
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
                Dim m_LuongDonVi As LuongDonVi = New LuongDonVi
                m_LuongDonVi.id = smartReader.getString("id")
                m_LuongDonVi.idChiNhanh = smartReader.getInt32("idChiNhanh")
                m_LuongDonVi.idTienLuong = smartReader.getInt32("IdTienLuong")
                m_LuongDonVi.TuNgay = smartReader.getDatetime("TuNgay")
                m_LuongDonVi.SoQD = smartReader.getString("SoQD")
                m_LuongDonVi.NgayQD = smartReader.getDatetime("NgayQD")
                m_LuongDonVi.NguoiQD = smartReader.getString("NguoiQD")
                m_LuongDonVi.IdCV_Nguoi_QD = smartReader.getInt32("IdCV_Nguoi_QD")
                m_LuongDonVi.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_LuongDonVi)
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
            strSql = "SELECT * FROM CN_TienLuong Order by TuNgay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Chi NHánh và các Chi Nhánh con 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAllByDonVi() As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM CN_TienLuong WHERE idChiNhanh=" & IdDONVI.ToString & " or idChiNhanh in (SELECT id FROM ChiNhanh WHERE id_goc=" & IdDONVI.ToString & ")  or idChiNhanh in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & IdDONVI.ToString & ")) Order by TuNgay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Lấy Thông tin Luong Don vi theo Id Đơn vị
    ''' </summary>
    ''' <param name="IdDonVi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAllByDonVi(ByVal IdDonVi As Integer) As LuongDonVi
        Try
            Dim strSql As String
            strSql = "SELECT Top 1 * FROM CN_TienLuong WHERE idChiNhanh=" & IdDonVi.ToString & " order by TuNgay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), LuongDonVi)
            Else
                Return New LuongDonVi
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vId As String) As LuongDonVi
        Try
            Dim strSql As String
            strSql = "SELECT * FROM CN_TienLuong WHERE id='" & vId.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), LuongDonVi)
            Else
                Return New LuongDonVi
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class KeHoachLaoDong
    Dim _IdKHLD As String
    Dim _IdDonVi As Integer
    Dim _IdPhong As Integer
    Dim _Nam As Integer
    Dim _Thang As Integer
    Dim _DaiHan As Integer
    Dim _NganHan As Integer
    Dim _Ghichu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdKHLD() As String
        Get
            Return _IdKHLD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdKHLD = Value
        End Set
    End Property

    Public Property IdDonVi_KH() As Integer
        Get
            Return _IdDonVi
        End Get
        Set(ByVal Value As Integer)
            _IdDonVi = Value
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

    Public Property DaiHan() As Integer
        Get
            Return _DaiHan
        End Get
        Set(ByVal Value As Integer)
            _DaiHan = Value
        End Set
    End Property

    Public Property NganHan() As Integer
        Get
            Return _NganHan
        End Get
        Set(ByVal Value As Integer)
            _NganHan = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _Ghichu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _Ghichu = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("KeHoachLD_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_KH))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@DaiHan", DaiHan))
            cmd.Parameters.Add(New SqlParameter("@NganHan", NganHan))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdKHLD", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdKHLD = cmd.Parameters("@IdKHLD").Value.ToString
        Catch ex As Exception
            IdKHLD = ""
        End Try
        Return IdKHLD
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("KeHoachLD_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKHLD", IdKHLD))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_KH))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@Thang", Thang))
            cmd.Parameters.Add(New SqlParameter("@DaiHan", DaiHan))
            cmd.Parameters.Add(New SqlParameter("@NganHan", NganHan))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("KeHoachLD_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKHLD", IdKHLD))
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
                Dim m_KeHoachLD As KeHoachLaoDong = New KeHoachLaoDong
                m_KeHoachLD.IdKHLD = smartReader.getString("IdKHLD")
                m_KeHoachLD.IdDonVi_KH = smartReader.getInt32("IdDonVi")
                m_KeHoachLD.IdPhong = smartReader.getInt32("IdPhong")
                m_KeHoachLD.Nam = smartReader.getInt32("Nam")
                m_KeHoachLD.Thang = smartReader.getInt32("Thang")
                m_KeHoachLD.DaiHan = smartReader.getInt32("DaiHan")
                m_KeHoachLD.NganHan = smartReader.getInt32("NganHan")
                m_KeHoachLD.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_KeHoachLD)
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
            strSql = "SELECT * FROM KeHoachLD Order By Nam desc, Thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Chi NHánh và các Chi Nhánh con 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAllByDonVi() As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM KeHoachLD WHERE IdDonVi=" & IdDONVI & " or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & IdDONVI & ") Order by IdKHLD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Lấy Thông tin Kế hoạch lao động Don vi theo Id Đơn vị
    ''' </summary>
    ''' <param name="vIdDonVi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAllByDonVi(ByVal vIdDonVi As Integer) As KeHoachLaoDong
        Try
            Dim strSql As String
            strSql = "SELECT Top 1 * FROM KeHoachLD WHERE IdDonVi=" & vIdDonVi & " order by Nam desc, Thang desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), KeHoachLaoDong)
            Else
                Return New KeHoachLaoDong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Lấy Thông tin Kế hoạch lao động Don vi theo Id Đơn vị và năm kế hoạch
    ''' </summary>
    ''' <param name="vIdDonVi"></param>
    ''' <param name="vNam"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAllByDonVi(ByVal vIdDonVi As Integer, ByVal vNam As Integer) As IList
        Try
            Dim strSql As String
            If vIdDonVi = 0 Then
                strSql = " SELECT a.IdKHLD,  a.IdDonVi, a.IdPhong, a.Nam, a.Thang, a.DaiHan, a.NganHan FROM KeHoachLD a, (SELECT IdDonVi, Nam, MAX(Thang) as Thang FROM KeHoachLD WHERE Nam=" & vNam & " Group by IdDonVi, Nam  Having MAX(Thang)>=0) b " & _
                         " WHERE(a.IdDonvi = b.IdDonvi And a.Nam = b.Nam And a.Thang = b.Thang) " & _
                         " AND (a.IdDonVi=" & IdDONVI  & " OR  a.IdDonVi in (SELECT Id FROM ChiNhanh WHERE Id_goc=" & IdDonVi & ")) AND a.Nam = " & vNam & _
                         " Order by IdDonVi asc, IDphong asc "
            Else
                If vIdDonVi = IdDONVI Then
                    strSql = " SELECT a.IdKHLD,  a.IdDonVi, a.IdPhong, a.Nam, a.Thang, a.DaiHan, a.NganHan FROM KeHoachLD a, (SELECT IdDonVi, IdPhong, Nam, MAX(Thang) as Thang FROM KeHoachLD WHERE Nam=" & vNam & " Group by IdDonVi, IdPhong, Nam  Having MAX(Thang)>=0) b " & _
                         " WHERE(a.IdDonvi = b.IdDonvi And a.IdPhong = b.IdPhong And a.Nam = b.Nam And a.Thang = b.Thang) " & _
                         " AND (a.IdDonVi=" & IdDONVI & " OR  a.IdDonVi in (SELECT Id FROM ChiNhanh WHERE Id_goc=" & IdDONVI & ")) AND a.Nam = " & vNam & _
                         " Order by IdDonVi asc, IDphong asc "
                Else
                    strSql = " SELECT a.IdKHLD,  a.IdDonVi, a.IdPhong, a.Nam, a.Thang, a.DaiHan, a.NganHan FROM KeHoachLD a, (SELECT IdDonVi, Nam, MAX(Thang) as Thang FROM KeHoachLD WHERE Nam=" & vNam & " Group by IdDonVi, Nam  Having MAX(Thang)>=0) b " & _
                         " WHERE(a.IdDonvi = b.IdDonvi And a.Nam = b.Nam And a.Thang = b.Thang) " & _
                         " AND a.IdDonVi=" & vIdDonVi & " AND a.Nam = " & vNam & _
                         " Order by IdDonVi asc, IDphong asc "
                End If
            End If

            'If DONVI = gMaDonViTW Then
            '    'strSql = "SELECT IdKHLD, IdDonVi, IdPhong, Nam, Thang, DaiHan, NganHan, GhiChu, Date_create, Date_Update FROM KeHoachLD WHERE IdDonVi in (SELECT Id FROM ChiNhanh WHERE ma_so='" & DONVI & "') AND Nam =" & vNam & _
            '    '" Union " & _
            '    If vIdDonVi = 0 Then
            '        strSql = " SELECT IdKHLD,  IdDonVi, IdPhong, Nam, Thang, DaiHan, NganHan FROM KeHoachLD a, (SELECT IdDonVi, Nam, MAX(Thang) as Thang FROM KeHoachLD WHERE Nam=" & vNam & " Group by IdDonVi, Nam  Having MAX(Thang)>=0) b " & _
            '                 " WHERE(a.IdDonvi = b.IdDonvi And a.Nam = b.Nam And a.Thang = b.Thang) " & _
            '                 " AND (a.IdDonVi=" & IdDonVi & " OR  a.IdDonVi in (SELECT Id FROM ChiNhanh WHERE Id_goc=" & IdDonVi & ")) AND a.Nam = " & vNam & _
            '                 " Order by IdDonVi asc, IDphong asc "
            '    Else
            '        If vIdDonVi = 1 Then
            '            strSql = "SELECT IdKHLD,  IdDonVi, IdPhong, '" & vNam & "' as Nam, '0' as Thang, DaiHan, NganHan FROM KeHoachLD WHERE IdDonvi=1 AND Nam= " & vNam
            '            '" SELECT '' as IdKHLD, [Id] as IdDonVi, 0 as IdPhong, '" & vNam & "' as Nam, '0' as Thang, " & _
            '            '     "  (SELECT sum(DaiHan) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & " ) as DaiHan, " & _
            '            '     "  (SELECT sum(NganHan) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & ") as NganHan,  " & _
            '            '     "  '', (SELECT max(Date_Create) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & ") as Date_Create,  " & _
            '            '     "  (SELECT max(Date_Update) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & ") as Date_Update " & _
            '            '     "  FROM ChiNhanh WHERE id in (SELECT distinct IdDonvi FROM KeHoachLD WHERE Nam=" & vNam & ")" & _
            '            '     " order by Nam desc, Thang desc, IdDonVi asc, IDphong asc "
            '        Else
            '            strSql = " SELECT '' as IdKHLD, [Id] as IdDonVi, 0 as IdPhong, '" & vNam & "' as Nam, '0' as Thang, " & _
            '                     "  (SELECT sum(DaiHan) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & " ) as DaiHan, " & _
            '                     "  (SELECT sum(NganHan) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & ") as NganHan,  " & _
            '                     "  '', (SELECT max(Date_Create) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & ") as Date_Create,  " & _
            '                     "  (SELECT max(Date_Update) FROM KeHoachLD WHERE IdDonvi=[Id] AND Nam =" & vNam & ") as Date_Update " & _
            '                     "  FROM ChiNhanh WHERE id in (SELECT distinct IdDonvi FROM KeHoachLD WHERE Nam=" & vNam & ")" & _
            '                     " order by Nam desc, Thang desc, IdDonVi asc, IDphong asc "
            '        End If

            '    End If
            'Else
            '    strSql = "SELECT * FROM KeHoachLD WHERE (IdDonVi=" & vIdDonVi & " OR  IdDonVi in (SELECT Id FROM ChiNhanh WHERE Id_goc=" & vIdDonVi & ")) AND Nam =" & vNam & " order by Nam desc, Thang desc, IdDonVi asc, IDphong asc"
            'End If
            ''End If
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vId As String) As KeHoachLaoDong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM KeHoachLD WHERE IdKHLD='" & vId.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), KeHoachLaoDong)
            Else
                Return New KeHoachLaoDong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
End Class

Public Class MangLuoiDonVi
    Dim _IdMangLuoi As String
    Dim _IdDonVi As Integer
    Dim _IdPGD As Integer
    Dim _Nam As Integer
    Dim _Quy As Int16
    Dim _SoDiemGD As Integer
    Dim _DuNo As Long
    Dim _GhiChu As String
    Dim _SoXaPhuong As Integer
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

    Public Property IdMangLuoi() As String
        Get
            Return _IdMangLuoi
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Mạng lưới đơn vị có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdMangLuoi = Value
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

    Public Property IdPGD() As Integer
        Get
            Return _IdPGD
        End Get
        Set(ByVal Value As Integer)
            _IdPGD = Value
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

    Public Property Quy() As Integer
        Get
            Return _Quy
        End Get
        Set(ByVal Value As Integer)
            _Quy = Value
        End Set
    End Property

    Public Property SoDiemGD() As Integer
        Get
            Return _SoDiemGD
        End Get
        Set(ByVal Value As Integer)
            _SoDiemGD = Value
        End Set
    End Property

    Public Property DuNo() As Long
        Get
            Return _DuNo
        End Get
        Set(ByVal Value As Long)
            _DuNo = Value
        End Set
    End Property

    Public Property SoXaPhuong() As Integer
        Get
            Return _SoXaPhuong
        End Get
        Set(ByVal Value As Integer)
            _SoXaPhuong = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _Ghichu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("MangLuoiDV_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi))
            cmd.Parameters.Add(New SqlParameter("@IdPGD", IdPGD))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@Quy", Quy))
            cmd.Parameters.Add(New SqlParameter("@SoDiemGD", SoDiemGD))
            cmd.Parameters.Add(New SqlParameter("@DuNo", DuNo))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add(New SqlParameter("@SoXaPhuong", SoXaPhuong))
            cmd.Parameters.Add("@IdMangLuoi", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdMangLuoi = cmd.Parameters("@IdMangLuoi").Value.ToString
        Catch ex As Exception
            IdMangLuoi = ""
        End Try
        Return IdMangLuoi
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("MangLuoiDV_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdMangLuoi", IdMangLuoi))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi))
            cmd.Parameters.Add(New SqlParameter("@IdPGD", IdPGD))
            cmd.Parameters.Add(New SqlParameter("@Nam", Nam))
            cmd.Parameters.Add(New SqlParameter("@Quy", Quy))
            cmd.Parameters.Add(New SqlParameter("@SoDiemGD", SoDiemGD))
            cmd.Parameters.Add(New SqlParameter("@DuNo", DuNo))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add(New SqlParameter("@SoXaPhuong", SoXaPhuong))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("MangLuoiDV_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdMangLuoi", IdMangLuoi))
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
                Dim m_MangLuoiDV As MangLuoiDonVi = New MangLuoiDonVi
                m_MangLuoiDV.IdMangLuoi = smartReader.getString("IdMangLuoi")
                m_MangLuoiDV.IdDonVi = smartReader.getInt32("IdDonVi")
                m_MangLuoiDV.IdPGD = smartReader.getInt32("IdPGD")
                m_MangLuoiDV.Nam = smartReader.getInt32("Nam")
                m_MangLuoiDV.Quy = smartReader.getInt16("Quy")
                m_MangLuoiDV.SoDiemGD = smartReader.getInt32("SoDiemGD")
                m_MangLuoiDV.DuNo = smartReader.getInt64("DuNo")
                m_MangLuoiDV.GhiChu = smartReader.getString("GhiChu")
                m_MangLuoiDV.SoXaPhuong = smartReader.getInt32("SoXaPhuong")
                list.Add(m_MangLuoiDV)
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
            strSql = "SELECT * FROM MangLuoiDV Order By Nam desc, Quy desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Chi NHánh và các Chi Nhánh con 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAllByDonVi(ByVal vIdDonVi As Integer, ByVal vNam As Integer) As IList
        Try
            Dim strSql As String
            If DONVI = gMaDonViTW Then
                strSql = " SELECT IdMangLuoi, IdDonVi, IdPGD, Nam, Quy, SoDiemGD, DuNo, GhiChu, SoXaPhuong, Date_create, Date_Update FROM MangLuoiDV WHERE IdDonVi in (SELECT Id FROM ChiNhanh WHERE ma_so='" & gMaDonViTW & "') AND Nam =" & vNam & _
                " Union " & _
                " SELECT '' as IdMangLuoi, IdDonvi, IdDonvi as IdPGD, '" & vNam & "' as Nam, Quy, sum(SoDiemGD), sum(DuNo),'', sum(SoXaPhuong), (SELECT max(Date_Create) FROM MangLuoiDV WHERE IdDonvi=t.IdDonVi AND Nam =" & vNam & ")as Date_Create, " & _
                 "  (SELECT max(Date_Update) FROM MangLuoiDV WHERE IdDonvi=t.IdDonVi AND Nam =" & vNam & ")as Date_Update  FROM MangLuoiDV t WHERE IdDonVi not in (SELECT Id FROM ChiNhanh WHERE ma_so='" & gMaDonViTW & "') AND Nam =" & vNam & " GROUP BY IdDonVi, Quy " & _
                " order by Nam desc, IdDonVi asc, quy asc "
            Else
                strSql = "SELECT * FROM MangLuoiDV WHERE (IdDonVi=" & vIdDonVi & " or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ")) AND Nam=" & vNam & " Order by Nam desc, Quy desc"
            End If
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vId As String) As MangLuoiDonVi
        Try
            Dim strSql As String
            strSql = "SELECT * FROM MangLuoiDV WHERE IdMangLuoi='" & vId.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), MangLuoiDonVi)
            Else
                Return New MangLuoiDonVi
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class