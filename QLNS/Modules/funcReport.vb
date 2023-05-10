Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Module funcReport

    Public Structure rptBC05
        Public ldtb_daihan As Integer
        Public ldtb_nganhan As Integer
        Public ldth_daihan As Integer
        Public ldth_nganhan As Integer
        Public luong_chucvu As Long
        Public luong_capbac As Long
        Public phucap As Long
        Public tiencong_trongoi As Long
        Public luong_BHXH As Integer
        Public luong_themgio As Long
        Public luong_bosung As Long
        Public tongthunhap As Long
        Public thunhap_binhquan As Long
    End Structure

    Public Function getBC05(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short, ByVal vTbDaiHan As Integer, ByVal vTbNganHan As Integer, ByRef myBC05 As rptBC05) As rptBC05
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BC05")

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add(New SqlParameter("@ldtb_daihan", vTbDaiHan))
        cmd.Parameters.Add(New SqlParameter("@ldtb_nganhan", vTbNganHan))
        cmd.Connection = conn
        Try
            conn.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim smartReader As SmartDataReader = New SmartDataReader(reader)
            While smartReader.Read
                myBC05.ldtb_daihan = smartReader.getInt32("c1")
                myBC05.ldtb_nganhan = smartReader.getInt32("c2")
                myBC05.ldth_daihan = smartReader.getInt32("c3")
                myBC05.ldth_nganhan = smartReader.getInt32("c4")
                myBC05.luong_chucvu = smartReader.getInt32("c5")
                myBC05.luong_capbac = smartReader.getInt64("c6")
                myBC05.phucap = smartReader.getInt64("c7")
                myBC05.tiencong_trongoi = smartReader.getInt64("c8")
                myBC05.luong_BHXH = smartReader.getInt64("c9")
                myBC05.luong_themgio = smartReader.getInt64("c10")
                myBC05.luong_bosung = smartReader.getInt64("c11")
                myBC05.tongthunhap = smartReader.getInt64("c12")
                myBC05.thunhap_binhquan = smartReader.getInt64("c13")
            End While
            smartReader.disposeReader(reader)
            Return myBC05
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getBC06(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vKy As Integer, ByVal vAll As Short) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BC06")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@ky", vKy))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "BC06")
            Return ds.Tables("BC06")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getBC07(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BC07")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "BC07")
            Return ds.Tables("BC07")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getBC08(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vKy As Short, ByVal vAll As Short) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BC08")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@ky", vKy))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "BC08")
            Return ds.Tables("BC08")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getBC09(ByVal vIdDonVi As Integer, ByVal vNgayBaoCao As Date, ByVal vAll As Short, ByRef vN1 As Integer, ByRef vN2 As Integer, ByRef vN3 As Integer, ByRef vN4 As Integer, ByRef vN5 As Integer, ByRef vN6 As Integer, ByRef vN7 As Integer, ByRef vN8 As Integer, ByRef vN9 As Integer, ByRef vN10 As Integer, ByRef vN11 As Integer, ByRef vN12 As Integer, ByRef vN13 As Integer, ByRef vN14 As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BC09")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@ThoiDiem", vNgayBaoCao.Date))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add("@TC3", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC4", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC5", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC6", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC7", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC8", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC9", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC10", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC11", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC12", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC13", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC14", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC15", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC16", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Connection = conn
        cmd.CommandTimeout = 210
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "BC09")
            vN1 = cmd.Parameters("@TC3").Value
            vN2 = cmd.Parameters("@TC4").Value
            vN3 = cmd.Parameters("@TC5").Value
            vN4 = cmd.Parameters("@TC6").Value
            vN5 = cmd.Parameters("@TC7").Value
            vN6 = cmd.Parameters("@TC8").Value
            vN7 = cmd.Parameters("@TC9").Value
            vN8 = cmd.Parameters("@TC10").Value
            vN9 = cmd.Parameters("@TC11").Value
            vN10 = cmd.Parameters("@TC12").Value
            vN11 = cmd.Parameters("@TC13").Value
            vN12 = cmd.Parameters("@TC14").Value
            vN13 = cmd.Parameters("@TC15").Value
            vN14 = cmd.Parameters("@TC16").Value
            Return ds.Tables("BC09")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getBC09_NN(ByVal vIdDonVi As Integer, ByVal vNgayBaoCao As Date, ByVal vAll As Short, ByRef vN1 As Integer, ByRef vN2 As Integer, ByRef vN3 As Integer, ByRef vN4 As Integer, ByRef vN5 As Integer, ByRef vN6 As Integer, ByRef vN7 As Integer, ByRef vN8 As Integer, ByRef vN9 As Integer, ByRef vN10 As Integer, ByRef vN11 As Integer, ByRef vN12 As Integer, ByRef vN13 As Integer, ByRef vN14 As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BC09_NN")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@ThoiDiem", vNgayBaoCao.Date))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add("@TC3", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC4", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC5", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC6", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC7", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC8", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC9", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC10", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC11", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC12", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC13", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC14", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC15", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC16", SqlDbType.Int).Direction = ParameterDirection.Output
        cmd.Connection = conn
        cmd.CommandTimeout = 210
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "BC09_NN")
            vN1 = cmd.Parameters("@TC3").Value
            vN2 = cmd.Parameters("@TC4").Value
            vN3 = cmd.Parameters("@TC5").Value
            vN4 = cmd.Parameters("@TC6").Value
            vN5 = cmd.Parameters("@TC7").Value
            vN6 = cmd.Parameters("@TC8").Value
            vN7 = cmd.Parameters("@TC9").Value
            vN8 = cmd.Parameters("@TC10").Value
            vN9 = cmd.Parameters("@TC11").Value
            vN10 = cmd.Parameters("@TC12").Value
            vN11 = cmd.Parameters("@TC13").Value
            vN12 = cmd.Parameters("@TC14").Value
            vN13 = cmd.Parameters("@TC15").Value
            vN14 = cmd.Parameters("@TC16").Value
            Return ds.Tables("BC09_NN")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    ''' <summary>
    ''' createBangChiLuong
    ''' </summary>
    ''' <param name="vIdHS_ChiLuong"></param>
    ''' <param name="vLoaiCB">0: Lao động dài hạn; 1: Lao động ngắn hạn; 2: Tập sự</param>
    ''' <param name="vIdDonVi"></param>
    ''' <param name="vIdPhongBan"></param>
    ''' <param name="vNam"></param>
    ''' <param name="vThang"></param>
    ''' <param name="vKy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function createBangChiLuong(ByVal vIdHS_ChiLuong As String, ByVal vLoaiCB As Int16, ByVal vIdDonVi As Integer, ByVal vIdPhongBan As String, ByVal vNam As Integer, ByVal vThang As Integer, ByVal vKy As Int16) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        If vLoaiCB = 0 Then
            cmd = New SqlCommand("createBangKeChiLuong")
            cmd.Parameters.Add(New SqlParameter("@Phong", vIdPhongBan))
            cmd.Parameters.Add(New SqlParameter("@Ky", vKy))
        Else
            'If vLoaiCB = 1 Then
            cmd = New SqlCommand("createBangKeChiLuong_NH_TS")
            cmd.Parameters.Add(New SqlParameter("@NH_TS", vLoaiCB))
            'Else
            '    cmd = New SqlCommand("createBangKeChiLuong_TS")
            'End If
        End If
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
        cmd.Parameters.Add(New SqlParameter("@Thang", vThang))
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbChiLuong")
            Return ds.Tables("tbChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    ''' <summary>
    ''' createBangChiThemGio
    ''' </summary>
    ''' <param name="vIdHS_ChiLuong"></param>
    ''' <param name="vIdDonVi"></param>
    ''' <param name="vIdPhongBan"></param>
    ''' <param name="vNam"></param>
    ''' <param name="vThang"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function createBangChiThemGio(ByVal vIdHS_ChiLuong As String, ByVal vIdDonVi As Integer, ByVal vIdPhongBan As String, ByVal vNam As Integer, ByVal vThang As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd = New SqlCommand("createBangKeChiThemGio")
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Phong", vIdPhongBan))
        cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
        cmd.Parameters.Add(New SqlParameter("@Thang", vThang))
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbChiThemGio")
            Return ds.Tables("tbChiThemGio")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function


    Public Function createBangChiKhac(ByVal vIdHS_ChiLuong As String, ByVal vIdDonVi As Integer, ByVal vIdPhongBan As String, ByVal vLoaiCB As Int16, ByVal vListCB As List(Of String), ByVal vNam As Integer, ByVal vThang As Integer, ByVal vIdLoaiChikhac As Integer, ByVal vHeSoChi As Double, ByVal vSoTien As Long) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        Dim ds As New DataSet
        Dim da As SqlDataAdapter
        Dim i As Integer
        Dim arr() As String
        Dim Canbo As String = ""

        If Not vListCB Is Nothing Then
            Try
                conn.Open()
                For i = 0 To vListCB.Count - 1
                    Canbo = vListCB.Item(i)
                    arr = Canbo.Split("_")
                    Dim cmd1 As SqlCommand
                    'Dim loaicb As Int16
                    'Select Case arr(0)
                    '    Case "NH"
                    '        loaicb = 1
                    '    Case "TS"
                    '        loaicb = 2
                    '    Case "CC"
                    '        loaicb = 3
                    '    Case "VH"
                    '        loaicb = 4
                    '    Case Else
                    '        loaicb = 0
                    'End Select
                    cmd1 = New SqlCommand("createBangKeChiKhac_CB")
                    cmd1.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
                    cmd1.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
                    cmd1.Parameters.Add(New SqlParameter("@IdCanBo", arr(1)))
                    cmd1.Parameters.Add(New SqlParameter("@LoaiCB", vLoaiCB))
                    cmd1.Parameters.Add(New SqlParameter("@Nam", vNam))
                    cmd1.Parameters.Add(New SqlParameter("@Thang", vThang))
                    cmd1.Parameters.Add(New SqlParameter("@IdLoaiLuongBS", vIdLoaiChikhac))
                    cmd1.Parameters.Add(New SqlParameter("@HesoChi", vHeSoChi))
                    cmd1.Parameters.Add(New SqlParameter("@SoTien", vSoTien))
                    cmd1.CommandType = CommandType.StoredProcedure
                    cmd1.Connection = conn
                    cmd1.ExecuteNonQuery()
                Next
                If dbconn.getNumber("SELECT count(*) FROM HS_ChiBoSung_CT  WHERE IdHS_ChiLuong = '" & vIdHS_ChiLuong & "'") = 0 Then
                    Dim cmd1 As SqlCommand
                    cmd1 = New SqlCommand("HS_ChiLuong_Delete")
                    cmd1.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
                    cmd1.CommandType = CommandType.StoredProcedure
                    cmd1.ExecuteNonQuery()
                End If
            Catch ex2 As Exception
                Dim cmd2 As SqlCommand
                cmd2 = New SqlCommand("HS_ChiLuong_Delete")
                cmd2.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
                cmd2.CommandType = CommandType.StoredProcedure
                cmd2.ExecuteNonQuery()
            Finally
                dbconn.closeConnection(conn)
            End Try
            'cmd = New SqlCommand("createBangKeChiKhac_CT")
            'cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
            Try
                'Return showBangChiLuong(vIdHS_ChiLuong, 3, vLoaiCB, 0)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                dbconn.closeConnection(conn)
            End Try
        Else
            cmd = New SqlCommand("createBangKeChiKhac_CT")
            cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
            cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
            cmd.Parameters.Add(New SqlParameter("@Phong", vIdPhongBan))
            cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
            cmd.Parameters.Add(New SqlParameter("@Thang", vThang))
            cmd.Parameters.Add(New SqlParameter("@IdLoaiLuongBS", vIdLoaiChikhac))
            cmd.Parameters.Add(New SqlParameter("@HesoChi", vHeSoChi))
            cmd.Parameters.Add(New SqlParameter("@SoTien", vSoTien))
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = conn
            Try
                conn.Open()
                da = New SqlDataAdapter(cmd)
                da.Fill(ds, "tbChiBS")
                Return ds.Tables("tbChiBS")
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                dbconn.closeConnection(conn)
            End Try
        End If

    End Function

    ''' <summary>
    ''' showBangThuChi
    ''' </summary>
    ''' <param name="vIdHS_ChiLuong"></param>
    ''' <param name="vMaThuChi"></param>
    ''' <param name="vLoaiCB"></param>
    ''' <param name="vKy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function showBangThuChi(ByVal vIdHS_ChiLuong As String, ByVal vMaThuChi As String, ByVal vLoaiCB As Int16, ByVal vKy As Int16) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("showBangKeThuChi")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add(New SqlParameter("@MaThuChi", vMaThuChi))
        cmd.Parameters.Add(New SqlParameter("@LoaiCanBo", vLoaiCB))
        cmd.Parameters.Add(New SqlParameter("@Ky", vKy))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbBangChi")
            Return ds.Tables("tbBangChi")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="vIdHS_ChiLuong"></param>
    ''' <param name="vNH_TS">0: CB chinh thuc; 1: LĐ ngắn hạn; 2: LĐ tập sự</param>
    ''' <param name="vKy"></param>
    ''' <param name="vTK"></param>
    ''' <param name="vIdP"></param>
    ''' <param name="tCMNV"></param>
    ''' <param name="tPC_CV"></param>
    ''' <param name="tPC_TN"></param>
    ''' <param name="tPC_KV"></param>
    ''' <param name="tPC_DH"></param>
    ''' <param name="tTong100"></param>
    ''' <param name="tT_Nghi"></param>
    ''' <param name="tTongChi"></param>
    ''' <param name="tT_TNCNtmp"></param>
    ''' <param name="tT_DFCD"></param>
    ''' <param name="tT_BHXH"></param>
    ''' <param name="tT_BHYT"></param>
    ''' <param name="tT_BHTN"></param>
    ''' <param name="tT_TamUng"></param>
    ''' <param name="tT_KhoanThuKhac"></param>
    ''' <param name="tTongTru"></param>
    ''' <param name="tPC_TH_heso"></param>
    ''' <param name="tPC_TH_sotien"></param>
    ''' <param name="tThucLinh"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function printBangChiLuong(ByVal vIdHS_ChiLuong As String, ByVal vNH_TS As Int16, ByVal vKy As Int16, ByVal vTK As Boolean, ByRef vIdP As String, ByRef tCMNV As Double, _
                                            ByRef tPC_CV As Double, ByRef tPC_TN As Double, ByRef tPC_KV As Double, ByRef tPC_DH As Double, _
                                            ByRef tTong100 As Int64, ByRef tT_Nghi As Int64, ByRef tTongChi As Int64, ByRef tT_TNCNtmp As Int64, ByRef tT_DFCD As Int64, _
                                            ByRef tT_BHXH As Int64, ByRef tT_BHYT As Int64, ByRef tT_BHTN As Int64, ByRef tT_TamUng As Int64, _
                                            ByRef tT_KhoanThuKhac As Int64, ByRef tTongTru As Int64, ByRef tPC_TH_heso As Double, ByRef tPC_TH_sotien As Int64, ByRef tThucLinh As Int64) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("printBangKeChiLuong")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add(New SqlParameter("@NH_TS", vNH_TS))
        cmd.Parameters.Add(New SqlParameter("@Ky", vKy))
        cmd.Parameters.Add(New SqlParameter("@TK", vTK))
        cmd.Parameters.Add("@IdP", SqlDbType.VarChar, 6).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tCMNV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_CV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_KV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_TN", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_DH", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTong100", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_Nghi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongChi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_TNCNtmp", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_DFCD", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_BHXH", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_BHYT", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_BHTN", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_TamUng", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_KhoanThuKhac", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongTru", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_TH", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tThuHut_80", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tThucLinh", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ptChiLuong")
            vIdP = cmd.Parameters("@IdP").Value
            tCMNV = cmd.Parameters("@tCMNV").Value
            tPC_CV = cmd.Parameters("@tPC_CV").Value
            tPC_KV = cmd.Parameters("@tPC_KV").Value
            tPC_DH = cmd.Parameters("@tPC_DH").Value
            tPC_TN = cmd.Parameters("@tPC_TN").Value
            tTong100 = cmd.Parameters("@tTong100").Value
            tT_Nghi = cmd.Parameters("@tT_Nghi").Value
            tTongChi = cmd.Parameters("@tTongChi").Value
            tT_TNCNtmp = cmd.Parameters("@tT_TNCNtmp").Value
            tT_DFCD = cmd.Parameters("@tT_DFCD").Value
            tT_BHXH = cmd.Parameters("@tT_BHXH").Value
            tT_BHYT = cmd.Parameters("@tT_BHYT").Value
            tT_BHTN = cmd.Parameters("@tT_BHTN").Value
            tT_TamUng = cmd.Parameters("@tT_TamUng").Value
            tT_KhoanThuKhac = cmd.Parameters("@tT_KhoanThuKhac").Value
            tTongTru = cmd.Parameters("@tTongTru").Value
            tPC_TH_heso = cmd.Parameters("@tPC_TH").Value
            tPC_TH_sotien = cmd.Parameters("@tThuHut_80").Value
            tThucLinh = cmd.Parameters("@tThucLinh").Value
            Return ds.Tables("ptChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function printBangChiLuong_HDNH(ByVal vIdHS_ChiLuong As String, _
                                        ByRef tTong100 As Double, ByRef tTongChi As Int64, ByRef tT_DFCD As Int64, _
                                        ByRef tT_BHXH As Int64, ByRef tT_BHYT As Int64, ByRef tT_BHTN As Int64, ByRef tT_Nghi As Int64, _
                                        ByRef tT_TruyThu As Int64, ByRef tTongTru As Int64, ByRef tTruyLinh As Int64, ByRef tThucLinh As Int64) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("getReportBangKeChiLuong_HDNH")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add("@tTong100", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongChi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_DFCD", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_BHXH", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_BHYT", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_BHTN", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_Nghi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_TruyThu", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongTru", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTruyLinh", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tThucLinh", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiLuong")
            tTong100 = cmd.Parameters("@tTong100").Value
            tTongChi = cmd.Parameters("@tTongChi").Value
            tT_DFCD = cmd.Parameters("@tT_DFCD").Value
            tT_BHXH = cmd.Parameters("@tT_BHXH").Value
            tT_BHYT = cmd.Parameters("@tT_BHYT").Value
            tT_BHTN = cmd.Parameters("@tT_BHTN").Value
            tT_Nghi = cmd.Parameters("@tT_Nghi").Value
            tT_TruyThu = cmd.Parameters("@tT_TruyThu").Value
            tTongTru = cmd.Parameters("@tTongTru").Value
            tTruyLinh = cmd.Parameters("@tTruyLinh").Value
            tThucLinh = cmd.Parameters("@tThucLinh").Value
            Return ds.Tables("ChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getReportBangChiLuong_TS(ByVal vIdHS_ChiLuong As String, ByRef tCMNV As Double, _
                                            ByRef tTong As Double, ByRef tTong100 As Double, ByRef tTongChi As Int64, _
                                            ByRef tT_TruyThu As Int64, ByRef tTruyLinh As Int64, ByRef tThucLinh As Int64) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("getReportBangKeChiLuong_TS")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add("@tCMNV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTong", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTong100", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongChi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tT_TruyThu", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTruyLinh", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tThucLinh", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiLuong")
            tCMNV = cmd.Parameters("@tCMNV").Value
            tTong = cmd.Parameters("@tTong").Value
            tTong100 = cmd.Parameters("@tTong100").Value
            tTongChi = cmd.Parameters("@tTongChi").Value
            tT_TruyThu = cmd.Parameters("@tT_TruyThu").Value
            tTruyLinh = cmd.Parameters("@tTruyLinh").Value
            tThucLinh = cmd.Parameters("@tThucLinh").Value
            Return ds.Tables("ChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function printBangKeChiThemGio(ByVal vIdHS_ChiLuong As String, ByVal vTK As Boolean, ByRef vIdP As String, ByRef tCMNV As Double, ByRef tPC As Double, ByRef tHeso As Double, _
                                            ByRef tGio50 As Double, ByRef tGio100 As Double, ByRef tGio150 As Double, ByRef tGio200 As Double, ByRef tGio300 As Double, ByRef tGioQD As Double, _
                                            ByRef tGioNghiBu As Double, ByRef tTienLamDem As Int64, ByRef tThucLinh As Int64) As DataTable

        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("printBangKeChiThemGio")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add(New SqlParameter("@TK", vTK))
        cmd.Parameters.Add("@IdP", SqlDbType.VarChar, 6).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tCMNV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tHeso", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGio50", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGio100", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGio150", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGio200", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGio300", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGioQD", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tGioNghiBu", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTraThemLamDem", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tThucLinh", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiThemGio")
            vIdP = cmd.Parameters("@IdP").Value
            tCMNV = cmd.Parameters("@tCMNV").Value
            tPC = cmd.Parameters("@tPC").Value
            tHeso = cmd.Parameters("@tHeso").Value
            tGio50 = cmd.Parameters("@tGio50").Value
            tGio100 = cmd.Parameters("@tGio100").Value
            tGio150 = cmd.Parameters("@tGio150").Value
            tGio200 = cmd.Parameters("@tGio200").Value
            tGio300 = cmd.Parameters("@tGio300").Value
            tGioQD = cmd.Parameters("@tGioQD").Value
            tGioNghiBu = cmd.Parameters("@tGioNghiBu").Value
            tTienLamDem = cmd.Parameters("@tTraThemLamDem").Value
            tThucLinh = cmd.Parameters("@tThucLinh").Value
            Return ds.Tables("ChiThemGio")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getReportBangChiBoSung(ByVal vIdHS_ChiLuong As String, ByVal vLoaiCB As Int16, ByRef vIdP As String, ByRef tCMNV As Double, _
                                            ByRef tPC_CV As Double, ByRef tPC_KV As Double, ByRef tPC_KN As Double, ByRef tPC_DH As Double, _
                                            ByRef tPC_TN As Double, ByRef tPC_TNVK As Double, ByRef tPC_TNN As Double, ByRef tPC_Khac As Double, _
                                            ByRef tTong100 As Int64, ByRef tTongChi As Int64, ByRef NoidungChi As String) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("getReportBangKeChiBoSung")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add(New SqlParameter("@LoaiCB", vLoaiCB))
        cmd.Parameters.Add("@IdP", SqlDbType.VarChar, 6).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tCMNV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_CV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_KV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_KN", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_DH", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_TN", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_TNVK", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_TNN", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tPC_Khac", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTong100", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongChi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@NoidungChi", SqlDbType.NVarChar, 1024).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiBoSung")
            vIdP = cmd.Parameters("@IdP").Value
            tCMNV = cmd.Parameters("@tCMNV").Value
            tPC_CV = cmd.Parameters("@tPC_CV").Value
            tPC_KV = cmd.Parameters("@tPC_KV").Value
            tPC_KN = cmd.Parameters("@tPC_KN").Value
            tPC_DH = cmd.Parameters("@tPC_DH").Value
            tPC_TN = cmd.Parameters("@tPC_TN").Value
            tPC_TNVK = cmd.Parameters("@tPC_TNVK").Value
            tPC_TNN = cmd.Parameters("@tPC_TNN").Value
            tPC_Khac = cmd.Parameters("@tPC_Khac").Value
            tTong100 = cmd.Parameters("@tTong100").Value
            tTongChi = cmd.Parameters("@tTongChi").Value
            NoidungChi = cmd.Parameters("@NoidungChi").Value
            Return ds.Tables("ChiBoSung")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getReportBangChiBoSung_NHTS(ByVal vIdHS_ChiLuong As String, ByRef vIdP As String, ByRef tCMNV As Double, _
                                              ByRef tTong100 As Int64, ByRef tTongChi As Int64, ByRef NoidungChi As String) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("getReportBangKeChiBoSung_TS")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdHS_ChiLuong", vIdHS_ChiLuong))
        cmd.Parameters.Add("@IdP", SqlDbType.VarChar, 6).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tCMNV", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTong100", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@tTongChi", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@NoidungChi", SqlDbType.NVarChar, 1024).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiBoSung_TS")
            vIdP = cmd.Parameters("@IdP").Value
            tCMNV = cmd.Parameters("@tCMNV").Value
            tTong100 = cmd.Parameters("@tTong100").Value
            tTongChi = cmd.Parameters("@tTongChi").Value
            NoidungChi = cmd.Parameters("@NoidungChi").Value
            Return ds.Tables("ChiBoSung_TS")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Sub showReportChiLuong(ByVal rptView As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal rptDoc As ReportDocument, ByVal IdChiLuong As String, ByVal vKy As Int16, _
                                    ByVal vIdDonVi As Integer, ByVal vThang As Integer, ByVal vNam As Integer, _
                                    ByVal vLapBieu As String, ByVal vKiemSoat As String, ByVal vGiamDoc As String, ByVal vGhiChu As String)

        Dim tCMNV, tPC_CV, tPC_KV, tPC_KN, tPC_DH, tPC_TN, tPC_TNVK, tPC_TNN, tPC_Khac As String
        Dim vIdPhong As Integer
        Dim tTong100, tTongChi, tT_TNCNtmp, tT_DFCD, tT_BHXH, tT_BHYT, tT_BHTN, tT_TamUng, tT_Nghi, tT_TruyThu, tTongTru, tTruyLinh, tThucLinh As Long
        Dim ds As DataSet = New DataSet
        'ds = getReportBangChiLuong(IdChiLuong, vKy, vIdPhong, tCMNV, tPC_CV, tPC_KV, tPC_KN, tPC_DH, tPC_TN, tPC_TNVK, tPC_TNN, tPC_Khac, tTong100, tTongChi, tT_TNCNtmp, tT_DFCD, tT_BHXH, tT_BHYT, tT_BHTN, tT_TamUng, tT_Nghi, tT_TruyThu, tTongTru, tTruyLinh, tThucLinh)

        rptDoc.SetDataSource(ds)
        rptDoc.SetParameterValue("TC4", tCMNV)
        rptDoc.SetParameterValue("TC5", tPC_CV)
        rptDoc.SetParameterValue("TC6", tPC_KV)
        rptDoc.SetParameterValue("TC7", tPC_KN)
        rptDoc.SetParameterValue("TC8", tPC_DH)
        rptDoc.SetParameterValue("TC9", tPC_TN)
        rptDoc.SetParameterValue("TC10", tPC_TNVK)
        rptDoc.SetParameterValue("TC11", tPC_TNN)
        rptDoc.SetParameterValue("TC12", tPC_Khac)
        rptDoc.SetParameterValue("TN1", tTong100)
        rptDoc.SetParameterValue("TN2", tTongChi)
        rptDoc.SetParameterValue("TN3", tT_TNCNtmp)
        rptDoc.SetParameterValue("TN4", tT_DFCD)
        rptDoc.SetParameterValue("TN5", tT_BHXH)
        rptDoc.SetParameterValue("TN6", tT_BHYT)
        rptDoc.SetParameterValue("TN7", tT_TamUng)
        rptDoc.SetParameterValue("TN8", tT_Nghi)
        rptDoc.SetParameterValue("TN9", tTongTru)
        rptDoc.SetParameterValue("TN10", tThucLinh)
        rptDoc.SetParameterValue("TN11", tT_TruyThu)
        rptDoc.SetParameterValue("TN12", tTruyLinh)
        rptDoc.SetParameterValue("TN13", tT_BHTN)
        rptDoc.SetParameterValue("DonVi", getDonvi(vIdDonVi))
        rptDoc.SetParameterValue("Phong", getPhong(vIdPhong))
        rptDoc.SetParameterValue("tinh", DIABAN)
        rptDoc.SetParameterValue("ng", Now.Day)
        rptDoc.SetParameterValue("th", Now.Month)
        rptDoc.SetParameterValue("nm", Now.Year)
        rptDoc.SetParameterValue("Ky", vKy)
        rptDoc.SetParameterValue("thang", vThang)
        rptDoc.SetParameterValue("nam", vNam)
        If DONVI = gMaDonViTW Then
            rptDoc.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
        Else
            rptDoc.SetParameterValue("labGD", "GIÁM ĐỐC")
        End If
        rptDoc.SetParameterValue("LAPBANG", vLapBieu.Trim)
        rptDoc.SetParameterValue("KIEMSOAT", vKiemSoat.Trim)
        rptDoc.SetParameterValue("GD", vGiamDoc.Trim)
        rptDoc.SetParameterValue("GhiChu", vGhiChu.Trim)
        'frm.rptView.ReportSource = rptDoc
        'frm.rptView.Refresh()s
        'frm.rptView.Zoom(100)
        'frm.Show()
        rptView.ReportSource = rptDoc
        rptView.Zoom(100)
    End Sub

    Public Function getChiLuong(ByVal vIdDonVi As Integer, ByVal vIdPhongBan As Integer, ByVal vKy As Integer, ByVal vMonth As Integer, ByVal vYear As Integer, ByRef vC4 As Double, ByRef vC5 As Double, ByRef vC6 As Double, ByRef vC7 As Double, ByRef vC8 As Double, ByRef vC9 As Double, ByRef vC10 As Double, ByRef vC11 As Double, ByRef vC12 As Double, ByRef vC13 As Double, ByRef vN1 As Integer, ByRef vN2 As Integer, ByRef vN3 As Integer, ByRef vN4 As Integer, ByRef vN5 As Integer, ByRef vN6 As Integer, ByRef vN7 As Integer, ByRef vN8 As Integer, ByRef vN9 As Integer, ByRef vN10 As Integer, ByRef vN11 As Integer, ByRef vN12 As Integer, ByRef vN13 As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BangKeChiLuong")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@IdPhongBan", vIdPhongBan))
        cmd.Parameters.Add(New SqlParameter("@Ky", vKy))
        cmd.Parameters.Add(New SqlParameter("@Thang", vMonth))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add("@TC4", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC5", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC6", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC7", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC8", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC9", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC10", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC11", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC12", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC13", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC14", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC25", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC15", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC16", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC17", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC18", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC19", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC20", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC21", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC22", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC28", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC29", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC30", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiLuong")
            vC4 = cmd.Parameters("@TC4").Value
            vC5 = cmd.Parameters("@TC5").Value
            vC6 = cmd.Parameters("@TC6").Value
            vC7 = cmd.Parameters("@TC7").Value
            vC8 = cmd.Parameters("@TC8").Value
            vC9 = cmd.Parameters("@TC9").Value
            vC10 = cmd.Parameters("@TC10").Value
            vC11 = cmd.Parameters("@TC11").Value
            vC12 = cmd.Parameters("@TC12").Value
            vC13 = cmd.Parameters("@TC25").Value
            vN1 = cmd.Parameters("@TC13").Value
            vN2 = cmd.Parameters("@TC14").Value
            vN3 = cmd.Parameters("@TC15").Value
            vN4 = cmd.Parameters("@TC16").Value
            vN5 = cmd.Parameters("@TC17").Value
            vN6 = cmd.Parameters("@TC18").Value
            vN7 = cmd.Parameters("@TC19").Value
            vN8 = cmd.Parameters("@TC20").Value
            vN9 = cmd.Parameters("@TC21").Value
            vN10 = cmd.Parameters("@TC22").Value
            vN11 = cmd.Parameters("@TC28").Value
            vN12 = cmd.Parameters("@TC29").Value
            vN13 = cmd.Parameters("@TC30").Value
            Return ds.Tables("ChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getChiLuongHDNH(ByVal vIdDonVi As Integer, ByVal vMonth As Integer, ByVal vYear As Integer, ByRef vN1 As Integer, ByRef vN2 As Integer, ByRef vN3 As Integer, ByRef vN4 As Integer, ByRef vN5 As Integer, ByRef vN6 As Integer, ByRef vN7 As Integer, ByRef vN8 As Integer, ByRef vN9 As Integer, ByRef vN10 As Integer, ByRef vN11 As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BangKeChiLuong_HDNH")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Thang", vMonth))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add("@TC3", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC4", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC5", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC6", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC13", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC7", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC9", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC11", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC8", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC12", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC10", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiLuong")
            vN1 = cmd.Parameters("@TC3").Value
            vN2 = cmd.Parameters("@TC4").Value
            vN3 = cmd.Parameters("@TC5").Value
            vN4 = cmd.Parameters("@TC6").Value
            vN10 = cmd.Parameters("@TC13").Value
            vN5 = cmd.Parameters("@TC7").Value
            vN11 = cmd.Parameters("@TC9").Value
            vN8 = cmd.Parameters("@TC11").Value
            vN6 = cmd.Parameters("@TC8").Value
            vN9 = cmd.Parameters("@TC12").Value
            vN7 = cmd.Parameters("@TC10").Value
            Return ds.Tables("ChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getChiLuongTS(ByVal vIdDonVi As Integer, ByVal vMonth As Integer, ByVal vYear As Integer, ByRef vC3 As Double, ByRef vN1 As Integer, ByRef vN2 As Integer, ByRef vN3 As Integer, ByRef vN4 As Integer, ByRef vN5 As Integer, ByRef vN6 As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BangKeChiLuong_TS")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Thang", vMonth))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add("@TC3", SqlDbType.Float, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC4", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC5", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC6", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC8", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC9", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TC10", SqlDbType.Decimal, 18).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiLuong")
            vC3 = cmd.Parameters("@TC3").Value
            vN1 = cmd.Parameters("@TC4").Value
            vN2 = cmd.Parameters("@TC5").Value
            vN3 = cmd.Parameters("@TC6").Value
            vN4 = cmd.Parameters("@TC8").Value
            vN5 = cmd.Parameters("@TC9").Value
            vN6 = cmd.Parameters("@TC10").Value
            Return ds.Tables("ChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Sub getBangKeChiLuong(ByVal vIdDonVi As Integer, ByVal vKy As Integer, ByVal vMonth As Integer, ByVal vYear As Integer, ByRef vDataTable As DataTable, ByRef vColumns As String)
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("BangChiLuongTheoThang")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Ky", vKy))
        cmd.Parameters.Add(New SqlParameter("@Thang", vMonth))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add("@Colunms", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "ChiLuong")
            vColumns = cmd.Parameters("@Colunms").Value.ToString
            vDataTable = ds.Tables("ChiLuong")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Sub

    Public Function getDSKhenThuong(ByVal vIdDonVi As Integer, ByVal vKhenThuong As Int16, ByVal vDinhKi As Int16, ByVal vNam As Integer, ByVal vKT_ChuyenMon As Int16, ByVal vCaNhan As Int16, _
                                    ByRef tC6 As Int32, ByRef tC7 As Int32, ByRef tC8 As Int32, ByRef tC9 As Int32, ByRef tC10 As Int32, ByRef tC11 As Int32, ByRef tC12 As Int32, ByRef tC13 As Int32, _
                                    ByRef tC14 As Int32, ByRef tC15 As Int32, ByRef tC16 As Int32, ByRef tC17 As Int32, ByRef tC18 As Int32, ByRef tC19 As Int32, ByRef tC20 As Int32, ByRef tC21 As Int32, _
                                    ByRef tC22 As Int32, ByRef tC23 As Int32, ByRef tC24 As Int32, ByRef tC25 As Int32, ByRef tC26 As Int32, ByRef tC27 As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd = New SqlCommand("getDSKhenThuong")
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@KhenThuong", vKhenThuong))
        cmd.Parameters.Add(New SqlParameter("@DinhKi", vDinhKi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
        cmd.Parameters.Add(New SqlParameter("@KT_ChuyenMon", vKT_ChuyenMon))
        cmd.Parameters.Add(New SqlParameter("@CN_TT", vCaNhan))
        cmd.Parameters.Add("@TN6", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN7", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN8", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN9", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN10", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN11", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN12", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN13", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN14", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN15", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN16", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN17", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN18", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN19", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN20", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN21", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN22", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN23", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN24", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN25", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN26", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.Parameters.Add("@TN27", SqlDbType.Int, 8).Direction = ParameterDirection.Output
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 360
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbDSKT")
            tC6 = cmd.Parameters("@TN6").Value
            tC7 = cmd.Parameters("@TN7").Value
            tC8 = cmd.Parameters("@TN8").Value
            tC9 = cmd.Parameters("@TN9").Value
            tC10 = cmd.Parameters("@TN10").Value
            tC11 = cmd.Parameters("@TN11").Value
            tC12 = cmd.Parameters("@TN12").Value
            tC13 = cmd.Parameters("@TN13").Value
            tC14 = cmd.Parameters("@TN14").Value
            tC15 = cmd.Parameters("@TN15").Value
            tC16 = cmd.Parameters("@TN16").Value
            tC17 = cmd.Parameters("@TN17").Value
            tC18 = cmd.Parameters("@TN18").Value
            tC19 = cmd.Parameters("@TN19").Value
            tC20 = cmd.Parameters("@TN20").Value
            tC21 = cmd.Parameters("@TN21").Value
            tC22 = cmd.Parameters("@TN22").Value
            tC23 = cmd.Parameters("@TN23").Value
            tC24 = cmd.Parameters("@TN24").Value
            tC25 = cmd.Parameters("@TN25").Value
            tC26 = cmd.Parameters("@TN26").Value
            tC27 = cmd.Parameters("@TN27").Value
            Return ds.Tables("tbDSKT")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getDSDeNghiKT_CN_TT(ByVal vNam As Integer, ByVal vKT_ChuyenMon As Int16, ByVal vDinhKi As Int16, ByVal vCaNhan As Int16, ByVal vIdDoiTuong As String, ByVal vIdDonvi As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd = New SqlCommand("getDSDeNghiKhenThuong")
        cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
        cmd.Parameters.Add(New SqlParameter("@KT_ChuyenMon", vKT_ChuyenMon))
        cmd.Parameters.Add(New SqlParameter("@DinhKi", vDinhKi))
        cmd.Parameters.Add(New SqlParameter("@CN_TT", vCaNhan))
        cmd.Parameters.Add(New SqlParameter("@IdDoiTuong", vIdDoiTuong))
        cmd.Parameters.Add(New SqlParameter("@IdDonvi", vIdDonvi))
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbDSDN")
            Return ds.Tables("tbDSDN")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getTBNangBac(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short, ByVal vKy As Short) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        'alert_NangBacLuong_NEW
        Dim cmd As SqlCommand = Nothing

        If vYear < 2016 Then
            cmd = New SqlCommand("alert_NangBacLuong")
        Else
            cmd = New SqlCommand("alert_NangBacLuong_NEW")
        End If

        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add(New SqlParameter("@Ky", vKy))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "TBNB")
            Return ds.Tables("TBNB")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getTBNangNgach(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("alert_NangNgachLuong")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "TBNN")
            Return ds.Tables("TBNN")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Function getTBHetHanHDLD(ByVal vIdDonVi As Integer, ByVal vAll As Short, ByVal vNgayHetHan As DateTime) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("alert_HetHanHDLD")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add(New SqlParameter("@Dateline", vNgayHetHan))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "TBHH")
            Return ds.Tables("TBHH")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Public Sub bindCboThuocTinhTraCuu(ByVal cbo As ComboBox)
        cbo.Items.Clear()
        cbo.Items.Add("Mã cán bộ")           '0
        cbo.Items.Add("Họ tên")              '1
        cbo.Items.Add("Giới tính")           '2
        cbo.Items.Add("Độ tuổi")             '3
        cbo.Items.Add("Phòng ban")           '4
        cbo.Items.Add("Chức vụ")             '5
        cbo.Items.Add("Hệ số lương")         '6
        cbo.SelectedIndex = 1
    End Sub

    Public Sub bindCboToanTuTraCuu(ByVal cbo As ComboBox)
        cbo.Items.Add("=")          '0
        cbo.Items.Add(">")          '1
        cbo.Items.Add("<")          '2
        'cbo.Items.Add("like")       '3
        cbo.SelectedIndex = 0
    End Sub
End Module
