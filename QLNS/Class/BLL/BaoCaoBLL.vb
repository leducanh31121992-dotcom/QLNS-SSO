Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Syncfusion.XlsIO
Imports Syncfusion.ExcelToPdfConverter
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Collections
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Public Class BaoCaoBLL

    ''' <summary>
    ''' Hàm lấy dữ liệu báo cáo theo yêu cầu
    ''' </summary>
    ''' <param name="pUnitId">Chỉ số xác định Id của Chi nhánh/Đơn vị. Nếu 0 lấy tất</param>
    ''' <param name="pTuNgay">Từ ngày Định dạng yyyy-MM-dd</param>
    ''' <param name="pThoiDiem">Đến ngày/Ngày báo cáo Định dạng yyyy-MM-dd</param>
    ''' <param name="pLoaiBC">QUY ƯỚC:   1: TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG BÌNH QUÂN NĂM
    '''                                  2: TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG BÌNH QUÂN NĂM THEO VÙNG
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDuLieu_BaoCao(ByVal pUnitId As Integer, ByVal pTuNgay As String, ByVal pThoiDiem As String, ByVal pLoaiBC As Integer) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("BaoCao_LaoDongBQNam_05TCCB", connection)
            _command.CommandType = CommandType.StoredProcedure
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pUnitId", pUnitId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pTuNgay", pTuNgay, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiem", pThoiDiem, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pLoaiBC", pLoaiBC, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "DuLieuBC_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("DuLieuBC_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function



    ''' <summary>
    ''' Hàm lấy dữ liệu báo cáo Tình tình lao động - Mạng lưới: 04/BC-TCCB, 03/BC-TCCB, 04/BC-TCCB, 07/BC-TCCB
    ''' </summary>
    ''' <param name="pThoiDiem">Thời điểm báo cáo. Định dang yyyy-MM-dd</param>
    ''' <param name="pDonViId">Đơn vị cần lấy báo cáo. Lấy toàn quốc truyền vào là 0</param>
    ''' <param name="pFlagCall">Cờ xác định báo cáo. Quy ước: 4: Mẫu 04/BC-TCCB. BÁO CÁO TỔNG HỢP TÌNH HÌNH THỰC HIỆN LAO ĐỘNG - MẠNG LƯỚI 
    ''' 
    ''' 
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDLBaoCao_LaoDongMangLuoi(ByVal pThoiDiem As String, ByVal pDonViId As Integer, ByVal pFlagCall As Integer) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("BaoCao_TinhHinhLaoDongMangLuoi", connection)
            _command.CommandType = CommandType.StoredProcedure

            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiem", pThoiDiem, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pDonViId", pDonViId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pFlagCall", pFlagCall, ParameterDirection.Input))

            '_command.Parameters.Add("@_IdOutPut", SqlDbType.VarChar, 16).Direction = ParameterDirection.Output


            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "DuLieuBC_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                '_Ret = _command.Parameters("@_IdOutPut").Value.ToString()
                Return ds.Tables("DuLieuBC_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Hàm lấy dữ liệu báo cáo Tình tình lao động - Mạng lưới: 04/BC-TCCB, 03/BC-TCCB, 04/BC-TCCB, 07/BC-TCCB
    ''' </summary>
    ''' <param name="pThoiDiem">Thời điểm báo cáo. Định dang yyyy-MM-dd</param>
    ''' <param name="pDonViId">Đơn vị cần lấy báo cáo. Lấy toàn quốc truyền vào là 0</param>
    ''' <param name="pFlagCall">Cờ xác định báo cáo. Quy ước: 
    ''' 1: 01/BC-TCCB. TÌNH HÌNH THỰC HIỆN LAO ĐỘNG QUÝ: Tổng hợp theo: 
    '''                      I - Hội sở chính; II - Các chi nhánh NHCSXH Tỉnh/Thành phố; Các chi nhánh NHCSXH; Tổng cộng
    '''                      I - Hội sở tỉnh; II - Các PGD trực thuộc; Các PGD; Tổng cộng
    ''' 2: 01/BC-TCCB. TÌNH HÌNH THỰC HIỆN LAO ĐỘNG QUÝ: Tổng hợp theo: Khu vực => Chi nhánh và Tổng Cộng; Hoặc Các PGD/HST và Tổng cộng
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDLBaoCao_BaoCao_TinhHinhLD_01TCCB(ByVal pThoiDiem As String, ByVal pDonViId As Integer, ByVal pFlagCall As Integer) As DataTable
        Dim connection As SqlConnection = DbCommon.GetSqlConnection()
        If (connection Is Nothing) Then
            MessageBox.Show("Lost Connection !", "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button3)
            Return Nothing
        End If

        Try
            Dim _command As SqlCommand = New SqlCommand("BaoCao_TinhHinhLD_01TCCB", connection)
            _command.CommandType = CommandType.StoredProcedure

            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pThoiDiemCuoiQuy", pThoiDiem, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pUnitId", pDonViId, ParameterDirection.Input))
            _command.Parameters.Add(SoftSqlHelper.CreateParameter("@pLoaiBC", pFlagCall, ParameterDirection.Input))
            Using mydap As SqlDataAdapter = New SqlDataAdapter(_command)
                Dim ds As DataSet = New DataSet
                mydap.Fill(ds, "DuLieuBC_TMP")
                If (ds Is Nothing Or ds.Tables.Count = 0 Or ds.Tables(0).Rows.Count = 0) Then
                    Return Nothing
                End If
                Return ds.Tables("DuLieuBC_TMP")
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            If (connection.State = System.Data.ConnectionState.Open) Then
                connection.Close()
            End If
        End Try
        Return Nothing
    End Function
End Class
