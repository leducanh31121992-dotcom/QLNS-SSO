Imports System
Imports System.Data
Imports System.Data.SqlClient
''' <summary>
''' Author: Nguyễn Thị Thuỳ Giang
''' </summary>
''' <remarks></remarks>
''' 
Public Class KhenThuong
    Dim _IdKhenThuong As String   '(15),
    Dim _KhenThuong As Boolean
    Dim _DinhKy As Boolean
    Dim _CN_TT As Int16
    Dim _NamKT As Integer
    Dim _SoQD As String ' (30),
    Dim _NgayQD As Date
    Dim _IdCapKT As Integer
    Dim _IdChucVuKyQD As Integer
    Dim _NguoiKyQD As String  '(30)
    Dim _NoiDungKT As String  '(256)
    Dim _KT_ChuyenMon As Boolean
    Dim _GhiChu As String ' (512) 
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdKhenThuong() As String
        Get
            Return _IdKhenThuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID khen thưởng có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdKhenThuong = Value
        End Set
    End Property

    Public Property KhenThuong() As Boolean
        Get
            Return _KhenThuong
        End Get
        Set(ByVal Value As Boolean)
            _KhenThuong = Value
        End Set
    End Property

    Public Property DinhKy() As Boolean
        Get
            Return _DinhKy
        End Get
        Set(ByVal Value As Boolean)
            _DinhKy = Value
        End Set
    End Property

    Public Property CN_TT() As Int16
        Get
            Return _CN_TT
        End Get
        Set(ByVal Value As Int16)
            _CN_TT = Value
        End Set
    End Property

    Public Property NamKT() As Integer
        Get
            Return _NamKT
        End Get
        Set(ByVal Value As Integer)
            _NamKT = Value
        End Set
    End Property

    Public Property SoQD() As String
        Get
            Return _SoQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 30) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị số quyết định có độ dài không hợp lệ!", Value, Value.ToString())
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

    Public Property IdCapKT() As Integer
        Get
            Return _IdCapKT
        End Get
        Set(ByVal Value As Integer)
            _IdCapKT = Value
        End Set
    End Property

    Public Property IdChucVuKyQD() As Integer
        Get
            Return _IdChucVuKyQD
        End Get
        Set(ByVal Value As Integer)
            _IdChucVuKyQD = Value
        End Set
    End Property

    Public Property NguoiKyQD() As String
        Get
            Return _NguoiKyQD
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 30) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiKyQD = Value
        End Set
    End Property

    Public Property NoiDungKT() As String
        Get
            Return _NoiDungKT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị nội dung quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NoiDungKT = Value
        End Set
    End Property

    Public Property KT_ChuyenMon() As Boolean
        Get
            Return _KT_ChuyenMon
        End Get
        Set(ByVal Value As Boolean)
            _KT_ChuyenMon = Value
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
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@KhenThuong", KhenThuong))
            cmd.Parameters.Add(New SqlParameter("@DinhKy", DinhKy))
            cmd.Parameters.Add(New SqlParameter("@CN_TT", CN_TT))
            cmd.Parameters.Add(New SqlParameter("@NamKT", NamKT))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@IdCapKT", IdCapKT))
            cmd.Parameters.Add(New SqlParameter("@IdChucVuKyQD", IdChucVuKyQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiKyQD", NguoiKyQD))
            cmd.Parameters.Add(New SqlParameter("@NoiDungKT", NoiDungKT))
            cmd.Parameters.Add(New SqlParameter("@KT_ChuyenMon", KT_ChuyenMon))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdKhenThuong", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdKhenThuong = cmd.Parameters("@IdKhenThuong").Value.ToString
        Catch ex As Exception
            IdKhenThuong = ""
        End Try
        Return IdKhenThuong
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
            cmd.Parameters.Add(New SqlParameter("@KhenThuong", KhenThuong))
            cmd.Parameters.Add(New SqlParameter("@DinhKy", DinhKy))
            cmd.Parameters.Add(New SqlParameter("@CN_TT", CN_TT))
            cmd.Parameters.Add(New SqlParameter("@NamKT", NamKT))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@IdCapKT", IdCapKT))
            cmd.Parameters.Add(New SqlParameter("@IdChucVuKyQD", IdChucVuKyQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiKyQD", NguoiKyQD))
            cmd.Parameters.Add(New SqlParameter("@NoiDungKT", NoiDungKT))
            cmd.Parameters.Add(New SqlParameter("@KT_ChuyenMon", KT_ChuyenMon))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdatenotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_UpdatenotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
            cmd.Parameters.Add(New SqlParameter("@KhenThuong", KhenThuong))
            cmd.Parameters.Add(New SqlParameter("@DinhKy", DinhKy))
            cmd.Parameters.Add(New SqlParameter("@CN_TT", CN_TT))
            cmd.Parameters.Add(New SqlParameter("@NamKT", NamKT))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@IdCapKT", IdCapKT))
            cmd.Parameters.Add(New SqlParameter("@IdChucVuKyQD", IdChucVuKyQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiKyQD", NguoiKyQD))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
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
                Dim m_KhenThuong As KhenThuong = New KhenThuong
                m_KhenThuong.IdKhenThuong = smartReader.getString("IdKhenThuong")
                m_KhenThuong.KhenThuong = smartReader.getBoolean("KhenThuong")
                m_KhenThuong.DinhKy = smartReader.getBoolean("DinhKy")
                m_KhenThuong.CN_TT = smartReader.getInt16("CN_TT")
                m_KhenThuong.NamKT = smartReader.getInt32("NamKT")
                m_KhenThuong.SoQD = smartReader.getString("SoQD")
                m_KhenThuong.NgayQD = smartReader.getDatetime("NgayQD")
                m_KhenThuong.IdCapKT = smartReader.getInt32("IdCapKT")
                m_KhenThuong.IdChucVuKyQD = smartReader.getInt32("IdChucVuKyQD")
                m_KhenThuong.NguoiKyQD = smartReader.getString("NguoiKyQD")
                m_KhenThuong.NoiDungKT = smartReader.getString("NoiDungKT")
                m_KhenThuong.KT_ChuyenMon = smartReader.getBoolean("KT_ChuyenMon")
                m_KhenThuong.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_KhenThuong)
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
            strSql = "SELECT * FROM HS_KhenThuong Order by NgayQD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdKT As String) As KhenThuong
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_KhenThuong WHERE IdKhenThuong='" & vIdKT.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), KhenThuong)
            Else
                Return New KhenThuong
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    'Public Function getFinalRecord(ByVal vIdCanbo As String) As KhenThuongCB
    '    Try
    '        Dim strSql As String
    '        strSql = "SELECT TOP 1 * FROM HS_KhenThuong WHERE idCanbo='" & vIdCanbo.Trim & "' order by Ngay_QD desc"
    '        Dim cmd As SqlCommand = New SqlCommand(strSql)
    '        cmd.CommandType = CommandType.Text
    '        Dim list As IList
    '        list = init(cmd)
    '        If list.Count = 1 Then
    '            Return CType(list(0), KhenThuongCB)
    '        Else
    '            Return New KhenThuongCB
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function


#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

Public Class KhenThuong_CT
    Dim _IdKhenThuong_CT As String '(15),
    Dim _IdKhenThuong As String 'varchar (15),
    Dim _IdKhenThuong_CT_parent As String 'varchar (15),
    Dim _IdDonVi As Integer
    Dim _IdPhong As Integer
    Dim _IdDanhHieuHinhThuc As Integer
    Dim _CN_TT As Int16
    Dim _IdCN_TT As String '(18),
    Dim _TenCN_TT As String ' (128),
    Dim _DaDuyet As Boolean
    Dim _GhiChu As String ' (256),
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdKhenThuong_CT() As String
        Get
            Return _IdKhenThuong_CT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID khen thưởng chi tiết có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdKhenThuong_CT = Value
        End Set
    End Property

    Public Property IdKhenThuong() As String
        Get
            Return _IdKhenThuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID khen thưởng có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdKhenThuong = Value
        End Set
    End Property

    Public Property IdKhenThuong_CT_parent() As String
        Get
            Return _IdKhenThuong_CT_parent
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID khen thưởng của tập thể nhóm có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdKhenThuong_CT_parent = Value
        End Set
    End Property

    Public Property IdDonVi_KTCT() As Integer
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

    Public Property IdDanhHieuHinhThuc() As Integer
        Get
            Return _IdDanhHieuHinhThuc
        End Get
        Set(ByVal Value As Integer)
            _IdDanhHieuHinhThuc = Value
        End Set
    End Property

    Public Property CN_TT() As Int16
        Get
            Return _CN_TT
        End Get
        Set(ByVal Value As Int16)
            _CN_TT = Value
        End Set
    End Property

    Public Property IdCN_TT() As String
        Get
            Return _IdCN_TT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 18) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị IdCN_TT có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdCN_TT = Value
        End Set
    End Property

    Public Property TenCN_TT() As String
        Get
            Return _TenCN_TT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị TenCN_TT có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _TenCN_TT = Value
        End Set
    End Property

    Public Property DaDuyet() As Boolean
        Get
            Return _DaDuyet
        End Get
        Set(ByVal Value As Boolean)
            _DaDuyet = Value
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
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong_CT_parent", IdKhenThuong_CT_parent))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_KTCT))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdDanhHieuHinhThuc", IdDanhHieuHinhThuc))
            cmd.Parameters.Add(New SqlParameter("@CN_TT", CN_TT))
            cmd.Parameters.Add(New SqlParameter("@IdCN_TT", IdCN_TT))
            cmd.Parameters.Add(New SqlParameter("@TenCN_TT", TenCN_TT))
            cmd.Parameters.Add(New SqlParameter("@DaDuyet", DaDuyet))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdKhenThuong_CT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdKhenThuong_CT = cmd.Parameters("@IdKhenThuong_CT").Value.ToString
        Catch ex As Exception
            IdKhenThuong_CT = ""
        End Try
        Return IdKhenThuong_CT
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong_CT", IdKhenThuong_CT))
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong_CT_parent", IdKhenThuong_CT_parent))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", IdDonVi_KTCT))
            cmd.Parameters.Add(New SqlParameter("@IdPhong", IdPhong))
            cmd.Parameters.Add(New SqlParameter("@IdDanhHieuHinhThuc", IdDanhHieuHinhThuc))
            cmd.Parameters.Add(New SqlParameter("@CN_TT", CN_TT))
            cmd.Parameters.Add(New SqlParameter("@IdCN_TT", IdCN_TT))
            cmd.Parameters.Add(New SqlParameter("@TenCN_TT", TenCN_TT))
            cmd.Parameters.Add(New SqlParameter("@DaDuyet", DaDuyet))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdatenotFull()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_UpdatenotFull")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong_CT", IdKhenThuong_CT))
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
            cmd.Parameters.Add(New SqlParameter("@IdDanhHieuHinhThuc", IdDanhHieuHinhThuc))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdatenotDuyet()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_Update_Duyet")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong_CT", IdKhenThuong_CT))
            cmd.Parameters.Add(New SqlParameter("@DaDuyet", DaDuyet))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong_CT", IdKhenThuong_CT))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            db.executeSQL(cmd)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function DeleteForIdKhenThuong() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_DeleteForIdKhenThuong")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKhenThuong", IdKhenThuong))
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
                Dim m_KhenThuongCT As KhenThuong_CT = New KhenThuong_CT
                m_KhenThuongCT.IdKhenThuong_CT = smartReader.getString("IdKhenThuong_CT")
                m_KhenThuongCT.IdKhenThuong = smartReader.getString("IdKhenThuong")
                m_KhenThuongCT.IdKhenThuong_CT_parent = smartReader.getString("IdKhenThuong_CT_parent")
                m_KhenThuongCT.IdDonVi_KTCT = smartReader.getInt32("IdDonVi")
                m_KhenThuongCT.IdPhong = smartReader.getInt32("IdPhong")
                m_KhenThuongCT.IdDanhHieuHinhThuc = smartReader.getInt32("IdDanhHieuHinhThuc")
                m_KhenThuongCT.CN_TT = smartReader.getInt16("CN_TT")
                m_KhenThuongCT.IdCN_TT = smartReader.getString("IdCN_TT")
                m_KhenThuongCT.TenCN_TT = smartReader.getString("TenCN_TT")
                m_KhenThuongCT.DaDuyet = smartReader.getBoolean("DaDuyet")
                m_KhenThuongCT.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_KhenThuongCT)
            End While
            smartReader.disposeReader(reader)
            Return list
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            db.closeConnection(conn)
        End Try

    End Function

    Public Function getRecord(ByVal vIdKhenThuong_CT As String) As KhenThuong_CT
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_KhenThuong_CT WHERE IdKhenThuong_CT='" & vIdKhenThuong_CT.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), KhenThuong_CT)
            Else
                Return New KhenThuong_CT
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByHSKhenThuong(ByVal vIdKhenThuong As String) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_KhenThuong_CT WHERE idKhenThuong ='" & vIdKhenThuong & "' and (IdKhenThuong_CT_parent is null or IdKhenThuong_CT_parent='') Order by IdDonVi,CN_TT desc"
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
            strSql = "SELECT * FROM HS_KhenThuong_CT WHERE idCN_TT='" & vIdCanbo.Trim & "' order by idKhenThuong"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    'Public Function getAllByDanhHieuTD(ByVal vIdDanhHieuTD As Integer) As IList
    '    Try
    '        Dim strSql As String
    '        strSql = "SELECT * FROM HS_KhenThuong WHERE IdDanhhieuTD =" & vIdDanhHieuTD & " order by Ngay_QD desc"
    '        Dim cmd As SqlCommand = New SqlCommand(strSql)
    '        cmd.CommandType = CommandType.Text
    '        Return init(cmd)
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function

    'Public Function getRecord(ByVal vIdKT As String) As KhenThuongCB
    '    Try
    '        Dim strSql As String
    '        strSql = "SELECT * FROM HS_KhenThuong WHERE IdKhenThuong_CB='" & vIdKT.Trim & "'"
    '        Dim cmd As SqlCommand = New SqlCommand(strSql)
    '        cmd.CommandType = CommandType.Text
    '        Dim list As IList
    '        list = init(cmd)
    '        If list.Count = 1 Then
    '            Return CType(list(0), KhenThuongCB)
    '        Else
    '            Return New KhenThuongCB
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function

    'Public Function getFinalRecord(ByVal vIdCanbo As String) As KhenThuongCB
    '    Try
    '        Dim strSql As String
    '        strSql = "SELECT TOP 1 * FROM HS_KhenThuong WHERE idCanbo='" & vIdCanbo.Trim & "' order by Ngay_QD desc"
    '        Dim cmd As SqlCommand = New SqlCommand(strSql)
    '        cmd.CommandType = CommandType.Text
    '        Dim list As IList
    '        list = init(cmd)
    '        If list.Count = 1 Then
    '            Return CType(list(0), KhenThuongCB)
    '        Else
    '            Return New KhenThuongCB
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function


#End Region

End Class

Public Class KiLuatCB
    Dim _IdKiLuat_CB As String
    Dim _IdCanBo As String
    Dim _IdHinhThucKL As Integer
    Dim _LyDo As String
    Dim _TuNgay As Date
    Dim _DenNgay As Date
    Dim _SoQD As String
    Dim _NgayQD As Date
    Dim _NguoiQD As String
    Dim _IdCapKiLuat As Integer
    Dim _TrachNhiem_VC As String
    Dim _IdLuong As String 'sd: QD ki luat keo dai thoi han nang luong
    Dim _KeoDaiNangLuong As Integer 'sd: QD ki luat keo dai thoi han nang luong
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdKiLuat_CB() As String
        Get
            Return _IdKiLuat_CB
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID kỉ luật cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdKiLuat_CB = Value
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


    Public Property IdHinhThucKL() As Integer
        Get
            Return _IdHinhThucKL
        End Get
        Set(ByVal Value As Integer)
            _IdHinhThucKL = Value
        End Set
    End Property

    Public Property LyDo() As String
        Get
            Return _LyDo
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 200) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị lý do  có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _LyDo = Value
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
                If (Value.Length > 30) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị người quyết định có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiQD = Value
        End Set
    End Property

    Public Property IdCapKiLuat() As Integer
        Get
            Return _IdCapKiLuat
        End Get
        Set(ByVal Value As Integer)
            _IdCapKiLuat = Value
        End Set
    End Property

    Public Property TrachNhiem_VC() As String
        Get
            Return _TrachNhiem_VC
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 50) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị trách nhiệm vật chất có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _TrachNhiem_VC = Value
        End Set
    End Property

    Public Property IdLuong() As String
        Get
            Return _IdLuong
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID lương cán bộ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdLuong = Value
        End Set
    End Property

    Public Property KeoDaiNangLuong() As Integer
        Get
            Return _KeoDaiNangLuong
        End Get
        Set(ByVal Value As Integer)
            _KeoDaiNangLuong = Value
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

#Region "Method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KiLuat_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdHinhThucKL", IdHinhThucKL))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@LyDo", LyDo))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCapKiLuat", IdCapKiLuat))
            cmd.Parameters.Add(New SqlParameter("@TrachNhiem_VC", TrachNhiem_VC))
            cmd.Parameters.Add(New SqlParameter("@IdLuong", IdLuong))
            cmd.Parameters.Add(New SqlParameter("@KeoDaiNangLuong", KeoDaiNangLuong))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdKiLuat_CB", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdKiLuat_CB = cmd.Parameters("@IdKiLuat_CB").Value.ToString
        Catch ex As Exception
            IdKiLuat_CB = ""
        End Try
        Return IdKiLuat_CB
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KiLuat_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKiLuat_CB", IdKiLuat_CB))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdHinhThucKL", IdHinhThucKL))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@LyDo", LyDo))
            cmd.Parameters.Add(New SqlParameter("@SoQD", SoQD))
            cmd.Parameters.Add(New SqlParameter("@NgayQD", NgayQD))
            cmd.Parameters.Add(New SqlParameter("@NguoiQD", NguoiQD))
            cmd.Parameters.Add(New SqlParameter("@IdCapKiLuat", IdCapKiLuat))
            cmd.Parameters.Add(New SqlParameter("@TrachNhiem_VC", TrachNhiem_VC))
            cmd.Parameters.Add(New SqlParameter("@IdLuong", IdLuong))
            cmd.Parameters.Add(New SqlParameter("@KeoDaiNangLuong", KeoDaiNangLuong))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_KiLuat_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdKiLuat_CB", IdKiLuat_CB))
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
                Dim m_KiLuatCB As KiLuatCB = New KiLuatCB
                m_KiLuatCB.IdKiLuat_CB = smartReader.getString("IdKiLuat_CB")
                m_KiLuatCB.IdCanBo = smartReader.getString("IdCanBo")
                m_KiLuatCB.IdHinhThucKL = smartReader.getInt32("IdHinhThucKL")
                m_KiLuatCB.TuNgay = smartReader.getDatetime("TuNgay")
                m_KiLuatCB.DenNgay = smartReader.getDatetime("DenNgay")
                m_KiLuatCB.LyDo = smartReader.getString("LyDo")
                m_KiLuatCB.SoQD = smartReader.getString("SoQD")
                m_KiLuatCB.NgayQD = smartReader.getDatetime("NgayQD")
                m_KiLuatCB.NguoiQD = smartReader.getString("NguoiQD")
                m_KiLuatCB.IdCapKiLuat = smartReader.getInt32("IdCapKiLuat")
                m_KiLuatCB.TrachNhiem_VC = smartReader.getString("TrachNhiem_VC")
                m_KiLuatCB.IdLuong = smartReader.getString("IdLuong")
                m_KiLuatCB.KeoDaiNangLuong = smartReader.getInt32("KeoDaiNangLuong")
                m_KiLuatCB.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_KiLuatCB)
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
            strSql = "SELECT * FROM HS_KiLuat Order by NgayQD desc"
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
            strSql = "SELECT * FROM HS_KiLuat WHERE idCanbo='" & vIdCanbo.Trim & "' order by NgayQD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdKL As String) As KiLuatCB
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_KiLuat WHERE IdKiLuat_CB='" & vIdKL.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), KiLuatCB)
            Else
                Return New KiLuatCB
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As KiLuatCB
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_KiLuat WHERE idCanbo='" & vIdCanbo.Trim & "' order by NgayQD desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), KiLuatCB)
            Else
                Return New KiLuatCB
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class
