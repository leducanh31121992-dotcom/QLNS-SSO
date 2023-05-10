Imports System
Imports System.Data
Imports System.Data.SqlClient
''' <summary>
''' Class Nghiên cứu khoa học
''' Author: Nguyễn Thị Thuỳ Giang
''' </summary>
''' <remarks></remarks>
Public Class NghienCuuKhoaHoc
    Dim _IdCBNCKH As String
    Dim _IdDeTai As String
    Dim _IdCanBo As String
    Dim _ChuNhiem As Boolean
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdCBNCKH() As String
        Get
            Return _IdCBNCKH
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID cán bộ nghiên cứu khoa học có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdCBNCKH = Value
        End Set
    End Property

    Public Property IdDeTai() As String
        Get
            Return _IdDeTai
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Đề tài có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdDeTai = Value
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

    Public Property ChuNhiem() As Boolean
        Get
            Return _ChuNhiem
        End Get
        Set(ByVal Value As Boolean)
            _ChuNhiem = Value
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

#Region "method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_NCKH_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdDeTai", IdDeTai))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@ChuNhiem", ChuNhiem))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdCBNCKH", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdCBNCKH = cmd.Parameters("@IdCBNCKH").Value.ToString
        Catch ex As Exception
            IdCBNCKH = ""
        End Try
        Return IdCBNCKH
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_NCKH_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCBNCKH", IdCBNCKH))
            cmd.Parameters.Add(New SqlParameter("@IdDeTai", IdDeTai))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@ChuNhiem", ChuNhiem))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_NCKH_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCBNCKH", IdCBNCKH))
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
                Dim m_NCKH As NghienCuuKhoaHoc = New NghienCuuKhoaHoc
                m_NCKH.IdCBNCKH = smartReader.getString("IdCBNCKH")
                m_NCKH.IdDeTai = smartReader.getString("IdDeTai")
                m_NCKH.IdCanBo = smartReader.getString("IdCanBo")
                m_NCKH.ChuNhiem = smartReader.getBoolean("ChuNhiem")
                m_NCKH.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_NCKH)
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
            strSql = "SELECT * FROM HS_NCKH"
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
            strSql = "SELECT * FROM HS_NCKH WHERE idCanbo='" & vIdCanbo.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdCBNCKH As String) As NghienCuuKhoaHoc
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_NCKH WHERE IdCBNCKH='" & vIdCBNCKH.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), NghienCuuKhoaHoc)
            Else
                Return New NghienCuuKhoaHoc
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As NghienCuuKhoaHoc
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_NCKH WHERE idCanbo='" & vIdCanbo.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), NghienCuuKhoaHoc)
            Else
                Return New NghienCuuKhoaHoc
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class

Public Class DaoTaoVanBangChungChi
    Dim _IdDTVBCC As String
    Dim _IdCanBo As String
    Dim _IdLoaiVBCC As Integer
    Dim _NamTN As Integer
    Dim _CoSo_DT As String
    Dim _IdHinhThucDT As Integer
    Dim _IdChuyenNganhDT As Integer
    Dim _NganhHoc As String
    Dim _IdTrinhDo As Integer
    Dim _TuNgay As Date
    Dim _DenNgay As Date
    Dim _TyleHuong As Double
    Dim _XepLoai As Byte
    Dim _VBCC As Boolean
    Dim _Lop As String
    Dim _So_VBCC As String
    Dim _Ten_VBCC As String
    Dim _CuDiHoc As Byte
    Dim _IdNuocDT As Integer
    Dim _HoanThanh As Boolean
    Dim _NgayHL As Date
    Dim _NgayHH As Date
    Dim _NgayCap As Date
    Dim _NguoiKy As String
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdDTVBCC() As String
        Get
            Return _IdDTVBCC
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID Đào tạo văn bằng chứng chỉ có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdDTVBCC = Value
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

    Public Property IdLoaiVBCC() As Integer
        Get
            Return _IdLoaiVBCC
        End Get
        Set(ByVal Value As Integer)
            _IdLoaiVBCC = Value
        End Set
    End Property

    Public Property NamTN() As Integer
        Get
            Return _NamTN
        End Get
        Set(ByVal Value As Integer)
            _NamTN = Value
        End Set
    End Property

    Public Property CoSo_DT() As String
        Get
            Return _CoSo_DT
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Cơ sở đào tạo có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _CoSo_DT = Value
        End Set
    End Property

    Public Property IdHinhThucDT() As Integer
        Get
            Return _IdHinhThucDT
        End Get
        Set(ByVal Value As Integer)
            _IdHinhThucDT = Value
        End Set
    End Property

    Public Property IdChuyenNganhDT() As Integer
        Get
            Return _IdChuyenNganhDT
        End Get
        Set(ByVal Value As Integer)
            _IdChuyenNganhDT = Value
        End Set
    End Property
 
    Public Property NganhHoc() As String
        Get
            Return _NganhHoc
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Ngành học có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NganhHoc = Value
        End Set
    End Property

    Public Property IdTrinhDo() As Integer
        Get
            Return _IdTrinhDo
        End Get
        Set(ByVal Value As Integer)
            _IdTrinhDo = Value
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

    Public Property TyleHuong() As Double
        Get
            Return _TyleHuong
        End Get
        Set(ByVal Value As Double)
            _TyleHuong = Value
        End Set
    End Property

    Public Property XepLoai() As Byte
        Get
            Return _XepLoai
        End Get
        Set(ByVal Value As Byte)
            _XepLoai = Value
        End Set
    End Property

    Public Property VBCC() As Boolean
        Get
            Return _VBCC
        End Get
        Set(ByVal Value As Boolean)
            _VBCC = Value
        End Set
    End Property

    Public Property Lop() As String
        Get
            Return _Lop
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Lớp có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _Lop = Value
        End Set
    End Property

    Public Property So_VBCC() As String
        Get
            Return _So_VBCC
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 64) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Số văn bằng có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _So_VBCC = Value
        End Set
    End Property

    Public Property Ten_VBCC() As String
        Get
            Return _Ten_VBCC
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 512) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Tên văn bằng có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _Ten_VBCC = Value
        End Set
    End Property

    Public Property CuDiHoc() As Byte
        Get
            Return _CuDiHoc
        End Get
        Set(ByVal Value As Byte)
            _CuDiHoc = Value
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

    Public Property NgayHH() As Date
        Get
            Return _NgayHH
        End Get
        Set(ByVal Value As Date)
            _NgayHH = Value
        End Set
    End Property

    Public Property NgayCap() As Date
        Get
            Return _NgayCap
        End Get
        Set(ByVal Value As Date)
            _NgayCap = Value
        End Set
    End Property

    Public Property IdNuocDT() As Integer
        Get
            Return _IdNuocDT
        End Get
        Set(ByVal Value As Integer)
            _IdNuocDT = Value
        End Set
    End Property

    Public Property HoanThanh() As Boolean
        Get
            Return _HoanThanh
        End Get
        Set(ByVal Value As Boolean)
            _HoanThanh = Value
        End Set
    End Property
   
    Public Property NguoiKy() As String
        Get
            Return _NguoiKy
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Người ký có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NguoiKy = Value
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
            Dim cmd As SqlCommand = New SqlCommand("HS_DTVBCC_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiVBCC", IdLoaiVBCC))
            cmd.Parameters.Add(New SqlParameter("@NamTN", NamTN))
            cmd.Parameters.Add(New SqlParameter("@CoSo_DT", CoSo_DT))
            cmd.Parameters.Add(New SqlParameter("@IdHinhThucDT", IdHinhThucDT))
            cmd.Parameters.Add(New SqlParameter("@IdChuyenNganhDT", IdChuyenNganhDT))
            cmd.Parameters.Add(New SqlParameter("@NganhHoc", NganhHoc))
            cmd.Parameters.Add(New SqlParameter("@IdTrinhDo", IdTrinhDo))
            If TuNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@TuNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            End If
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@TyleHuong", TyleHuong))
            cmd.Parameters.Add(New SqlParameter("@XepLoai", XepLoai))
            cmd.Parameters.Add(New SqlParameter("@VBCC", VBCC))
            cmd.Parameters.Add(New SqlParameter("@Lop", Lop))
            cmd.Parameters.Add(New SqlParameter("@So_VBCC", So_VBCC))
            cmd.Parameters.Add(New SqlParameter("@Ten_VBCC", Ten_VBCC))
            cmd.Parameters.Add(New SqlParameter("@CuDiHoc", CuDiHoc))
            cmd.Parameters.Add(New SqlParameter("@IdNuocDT", IdNuocDT))
            cmd.Parameters.Add(New SqlParameter("@HoanThanh", HoanThanh))
            If NgayHL = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayHL", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayHL", NgayHL))
            End If
            If NgayHH = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayHH", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayHH", NgayHH))
            End If
            If NgayCap = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayCap", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayCap", NgayCap))
            End If
            cmd.Parameters.Add(New SqlParameter("@NguoiKy", NguoiKy))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdDTVBCC", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdDTVBCC = cmd.Parameters("@IdDTVBCC").Value.ToString
        Catch ex As Exception
            IdDTVBCC = ""
        End Try
        Return IdDTVBCC
    End Function
  
    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_DTVBCC_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDTVBCC", IdDTVBCC))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", IdCanBo))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiVBCC", IdLoaiVBCC))
            cmd.Parameters.Add(New SqlParameter("@NamTN", NamTN))
            cmd.Parameters.Add(New SqlParameter("@CoSo_DT", CoSo_DT))
            cmd.Parameters.Add(New SqlParameter("@IdHinhThucDT", IdHinhThucDT))
            cmd.Parameters.Add(New SqlParameter("@IdChuyenNganhDT", IdChuyenNganhDT))
            cmd.Parameters.Add(New SqlParameter("@NganhHoc", NganhHoc))
            cmd.Parameters.Add(New SqlParameter("@IdTrinhDo", IdTrinhDo))
            If TuNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@TuNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            End If
            If DenNgay = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            End If
            cmd.Parameters.Add(New SqlParameter("@TyleHuong", TyleHuong))
            cmd.Parameters.Add(New SqlParameter("@XepLoai", XepLoai))
            cmd.Parameters.Add(New SqlParameter("@VBCC", VBCC))
            cmd.Parameters.Add(New SqlParameter("@Lop", Lop))
            cmd.Parameters.Add(New SqlParameter("@So_VBCC", So_VBCC))
            cmd.Parameters.Add(New SqlParameter("@Ten_VBCC", Ten_VBCC))
            cmd.Parameters.Add(New SqlParameter("@CuDiHoc", CuDiHoc))
            cmd.Parameters.Add(New SqlParameter("@IdNuocDT", IdNuocDT))
            cmd.Parameters.Add(New SqlParameter("@HoanThanh", HoanThanh))
            If NgayHL = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayHL", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayHL", NgayHL))
            End If
            If NgayHH = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayHH", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayHH", NgayHH))
            End If
            If NgayCap = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayCap", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayCap", NgayCap))
            End If
            cmd.Parameters.Add(New SqlParameter("@NguoiKy", NguoiKy))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("HS_DTVBCC_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDTVBCC", IdDTVBCC))
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
                Dim m_DTVBCC As DaoTaoVanBangChungChi = New DaoTaoVanBangChungChi
                m_DTVBCC.IdDTVBCC = smartReader.getString("IdDTVBCC")
                m_DTVBCC.IdCanBo = smartReader.getString("IdCanBo")
                m_DTVBCC.IdLoaiVBCC = smartReader.getInt32("IdLoaiVBCC")
                m_DTVBCC.NamTN = smartReader.getInt32("NamTN")
                m_DTVBCC.CoSo_DT = smartReader.getString("CoSo_DT")
                m_DTVBCC.IdHinhThucDT = smartReader.getInt32("IdHinhThucDT")
                m_DTVBCC.IdChuyenNganhDT = smartReader.getInt32("IdChuyenNganhDT")
                m_DTVBCC.NganhHoc = smartReader.getString("NganhHoc")
                m_DTVBCC.IdTrinhDo = smartReader.getInt32("IdTrinhDo")
                m_DTVBCC.TuNgay = smartReader.getDatetime("TuNgay")
                m_DTVBCC.DenNgay = smartReader.getDatetime("DenNgay")
                m_DTVBCC.TyleHuong = smartReader.getFloat("TyleHuong")
                m_DTVBCC.XepLoai = smartReader.getByte("XepLoai")
                m_DTVBCC.VBCC = smartReader.getBoolean("VBCC")
                m_DTVBCC.Lop = smartReader.getString("Lop")
                m_DTVBCC.So_VBCC = smartReader.getString("So_VBCC")
                m_DTVBCC.Ten_VBCC = smartReader.getString("Ten_VBCC")
                m_DTVBCC.CuDiHoc = smartReader.getBoolean("CuDiHoc")
                m_DTVBCC.IdNuocDT = smartReader.getInt32("IdNuocDT")
                m_DTVBCC.HoanThanh = smartReader.getBoolean("HoanThanh")
                m_DTVBCC.NgayHL = smartReader.getDatetime("NgayHL")
                m_DTVBCC.NgayHH = smartReader.getDatetime("NgayHH")
                m_DTVBCC.NgayCap = smartReader.getDatetime("NgayCap")
                m_DTVBCC.NguoiKy = smartReader.getString("NguoiKy")
                m_DTVBCC.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_DTVBCC)
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
            strSql = "SELECT * FROM HS_DTVBCC Order by NamTN desc, Tungay desc"
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
            strSql = "SELECT * FROM HS_DTVBCC WHERE idCanbo='" & vIdCanbo.Trim & "' order by Tungay desc"  'NamTN desc
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdDTVBCC As String) As DaoTaoVanBangChungChi
        Try
            Dim strSql As String
            strSql = "SELECT * FROM HS_DTVBCC WHERE IdDTVBCC='" & vIdDTVBCC.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), DaoTaoVanBangChungChi)
            Else
                Return New DaoTaoVanBangChungChi
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCanbo As String) As DaoTaoVanBangChungChi
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM HS_DTVBCC WHERE idCanbo='" & vIdCanbo.Trim & "' order by Tungay desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), DaoTaoVanBangChungChi)
            Else
                Return New DaoTaoVanBangChungChi
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
#End Region

End Class

Public Class DeTaiNCKH
    Dim _IdDeTai As String
    Dim _IdCapDeTai As Integer
    Dim _TenDeTai As String
    Dim _NoiDung As String
    Dim _TuNgay As Date
    Dim _DenNgay As Date
    Dim _NgayNghiemThu As Date
    Dim _DonVi_QL As String
    Dim _GhiChu As String
    Private db As DBAccess

    Public Sub New()
        db = New DBAccess
    End Sub

#Region "Property"

    Public Property IdDeTai() As String
        Get
            Return _IdDeTai
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 15) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị ID đề tài có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _IdDeTai = Value
        End Set
    End Property

    Public Property IdCapDeTai() As Integer
        Get
            Return _IdCapDeTai
        End Get
        Set(ByVal Value As Integer)
            _IdCapDeTai = Value
        End Set
    End Property

    Public Property TenDeTai() As String
        Get
            Return _TenDeTai
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Tên đề tài có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _TenDeTai = Value
        End Set
    End Property

    Public Property NoiDung() As String
        Get
            Return _NoiDung
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Nội dung đề tài có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _NoiDung = Value
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

    Public Property NgayNghiemThu() As Date
        Get
            Return _NgayNghiemThu
        End Get
        Set(ByVal Value As Date)
            _NgayNghiemThu = Value
        End Set
    End Property

    Public Property DonVi_QL() As String
        Get
            Return _DonVi_QL
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 128) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Đơn vị chủ trì có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _DonVi_QL = Value
        End Set
    End Property

    Public Property GhiChu() As String
        Get
            Return _GhiChu
        End Get
        Set(ByVal Value As String)
            If Not (Value Is Nothing) Then
                If (Value.Length > 256) Then Throw New ArgumentOutOfRangeException("Chuỗi giá trị Ghi chú có độ dài không hợp lệ!", Value, Value.ToString())
            End If
            _GhiChu = Value
        End Set
    End Property

#End Region

#Region "method"

    Public Function Add() As String
        Try
            Dim cmd As SqlCommand = New SqlCommand("DeTaiNCKH_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdBranch", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@IdCapDeTai", IdCapDeTai))
            cmd.Parameters.Add(New SqlParameter("@TenDeTai", TenDeTai))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            If NgayNghiemThu = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayNghiemThu", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayNghiemThu", NgayNghiemThu))
            End If
            cmd.Parameters.Add(New SqlParameter("@DonVi_QL", DonVi_QL))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            cmd.Parameters.Add("@IdDeTai", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            IdDeTai = cmd.Parameters("@IdDeTai").Value.ToString
        Catch ex As Exception
            IdDeTai = ""
        End Try
        Return IdDeTai
    End Function

    Public Sub Update()
        Try
            Dim cmd As SqlCommand = New SqlCommand("DeTaiNCKH_Update")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDeTai", IdDeTai))
            cmd.Parameters.Add(New SqlParameter("@IdCapDeTai", IdCapDeTai))
            cmd.Parameters.Add(New SqlParameter("@TenDeTai", TenDeTai))
            cmd.Parameters.Add(New SqlParameter("@NoiDung", NoiDung))
            cmd.Parameters.Add(New SqlParameter("@TuNgay", TuNgay))
            cmd.Parameters.Add(New SqlParameter("@DenNgay", DenNgay))
            If NgayNghiemThu = "12:00:00 AM" Then
                cmd.Parameters.Add(New SqlParameter("@NgayNghiemThu", DBNull.Value))
            Else
                cmd.Parameters.Add(New SqlParameter("@NgayNghiemThu", NgayNghiemThu))
            End If
            cmd.Parameters.Add(New SqlParameter("@DonVi_QL", DonVi_QL))
            cmd.Parameters.Add(New SqlParameter("@GhiChu", GhiChu))
            db.executeSQL(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function Delete() As Boolean
        Try
            Dim cmd As SqlCommand = New SqlCommand("DeTaiNCKH_Delete")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@IdDeTai", IdDeTai))
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
                Dim m_DeTai As DeTaiNCKH = New DeTaiNCKH
                m_DeTai.IdDeTai = smartReader.getString("IdDeTai")
                m_DeTai.IdCapDeTai = smartReader.getInt32("IdCapDeTai")
                m_DeTai.TenDeTai = smartReader.getString("TenDeTai")
                m_DeTai.NoiDung = smartReader.getString("NoiDung")
                m_DeTai.TuNgay = smartReader.getDatetime("TuNgay")
                m_DeTai.DenNgay = smartReader.getDatetime("DenNgay")
                m_DeTai.NgayNghiemThu = smartReader.getDatetime("NgayNghiemThu")
                m_DeTai.DonVi_QL = smartReader.getString("DonVi_QL")
                m_DeTai.GhiChu = smartReader.getString("GhiChu")
                list.Add(m_DeTai)
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
            strSql = "SELECT * FROM DeTaiNCKH Order by Tungay desc, idDeTai desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getAllByCapDeTai(ByVal vCapDeTai As Integer) As IList
        Try
            Dim strSql As String
            strSql = "SELECT * FROM DeTaiNCKH WHERE IdCapDeTai=" & vCapDeTai & " order by Tungay desc, idDeTai desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Return init(cmd)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getRecord(ByVal vIdDeTai As String) As DeTaiNCKH
        Try
            Dim strSql As String
            strSql = "SELECT * FROM DeTaiNCKH WHERE IdDeTai='" & vIdDeTai.Trim & "'"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), DeTaiNCKH)
            Else
                Return New DeTaiNCKH
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getFinalRecord(ByVal vIdCapDeTai As String) As DeTaiNCKH
        Try
            Dim strSql As String
            strSql = "SELECT TOP 1 * FROM DeTaiNCKH WHERE idCapDeTai='" & vIdCapDeTai.Trim & "' order by vIdDeTai desc"
            Dim cmd As SqlCommand = New SqlCommand(strSql)
            cmd.CommandType = CommandType.Text
            Dim list As IList
            list = init(cmd)
            If list.Count = 1 Then
                Return CType(list(0), DeTaiNCKH)
            Else
                Return New DeTaiNCKH
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class
