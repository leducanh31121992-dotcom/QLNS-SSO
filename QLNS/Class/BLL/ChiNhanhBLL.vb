Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Public Class ChiNhanhBLL

    ''' <summary>
    ''' Hàm trả về Query lấy danh sách các đơn vị
    ''' </summary>
    ''' <param name="pPosCode"></param>
    ''' <param name="pIsALL">2 - Chỉ lấy danh sách chi nhánh</param>
    ''' <param name="pHoiSoTinh"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetQuerySQL_Branch(ByVal pPosCode As String, ByVal pIsALL As Byte, ByVal pHoiSoTinh As Byte) As String
        Dim sRetSQL As String = ""
        If pIsALL <> 2 Then
            sRetSQL = ""
            sRetSQL = sRetSQL & "Select ZZ.* From "
            sRetSQL = sRetSQL & "       ("
            sRetSQL = sRetSQL & "           Select Id,"
            If pHoiSoTinh = 1 Then
                sRetSQL = sRetSQL & "           (Case When Id_Goc=1 And Id > 5 Then N'Hội sở tỉnh' Else dbo.Replace_BranchName(Ten_Goi) End) TenCN,"
            Else
                sRetSQL = sRetSQL & "           dbo.Replace_BranchName(Ten_Goi) TenCN,"
            End If

            sRetSQL = sRetSQL & "           Id_Goc,Ma_So,Ten_Goi,Ten_VT,"
            sRetSQL = sRetSQL & "                  (Case When Ma_So In ('000100','000101','000196','000197') Then Ma_So Else (Select Top 1 Ma_So From ChiNhanh B Where B.Status=1 And B.Id_Goc=1 and Substring(B.Ma_So,3,2)=Substring(A.Ma_So,3,2)) End) MaCN,"
            sRetSQL = sRetSQL & "                  Ma_SapXep,Dia_Chi,Dien_Thoai,Email,Website,Status,MaDV,MaCB_Begin,MaCB_End,Ma_So_OLD,SoLg_XaPhuong,IP"
            sRetSQL = sRetSQL & "                  From ChiNhanh A Where A.Status=1"
            sRetSQL = sRetSQL & "       ) ZZ Where ZZ.Status=1"
            If pPosCode = "000100" Then
                If pIsALL = 0 Then
                    sRetSQL = sRetSQL & "          And ZZ.Id_Goc In (0,1) "
                End If
            Else
                sRetSQL = sRetSQL & "   And (Id In (Select Id From ChiNhanh Where Ma_So=N'" & pPosCode & "') Or Id_Goc In (Select Id From ChiNhanh Where Ma_So=N'" & pPosCode & "'))"
            End If
            sRetSQL = sRetSQL & "       Order By (Case When ZZ.MaCN='000100' Then 0 Else 1 End),ZZ.MaCN,ZZ.Ma_SapXep"
        Else
            If pPosCode = "000100" Or pPosCode = "000199" Then
                sRetSQL = "Select Id,Ten_Goi,Id_Goc,Ma_So,Ten_VT From ChiNhanh Where Id_Goc In (0,1) And Status=1 Order By Ma_So"
            Else
                sRetSQL = "Select Id,Ten_Goi,Id_Goc,Ma_So,Ten_VT From ChiNhanh Where Id_Goc In (0,1) And Status=1 And Ma_So = '" + pPosCode + "' Order By Ma_So"
            End If
        End If
        Return sRetSQL
    End Function

    Public Function GetChiNhanhs() As IList
        Dim ARL_List As IList = New ArrayList()
        Try
            Dim sSQL As String = ""
            If TRUCTHUOC = 1 Then
                sSQL = String.Format("Select Id,IsNull(Ten_Goi,'') Ten_Goi,IsNull(Id_Goc,0) Id_Goc,IsNull(Ma_So,'') Ma_So,IsNull(Ten_VT,'') Ten_VT,IsNull(Dia_Chi,'') Dia_Chi,IsNull(Dien_Thoai,'') Dien_Thoai,IsNull(Email,'') Email,IsNull(Website,'') Website,Status,IsNull(SoLg_XaPhuong,0) SoLg_XaPhuong,IsNull(MaDV,'') MaDV,IsNull(MaCB_Begin,'') MaCB_Begin,IsNull(MaCB_End,'') MaCB_End,IsNull(Ma_So_Old,'') Ma_So_Old,IsNull(IP,'') IP From ChiNhanh Where Status = 1 And Id_Goc In (0,1)")
            Else
                sSQL = String.Format("Select Id,IsNull(Ten_Goi,'') Ten_Goi,IsNull(Id_Goc,0) Id_Goc,IsNull(Ma_So,'') Ma_So,IsNull(Ten_VT,'') Ten_VT,IsNull(Dia_Chi,'') Dia_Chi,IsNull(Dien_Thoai,'') Dien_Thoai,IsNull(Email,'') Email,IsNull(Website,'') Website,Status,IsNull(SoLg_XaPhuong,0) SoLg_XaPhuong,IsNull(MaDV,'') MaDV,IsNull(MaCB_Begin,'') MaCB_Begin,IsNull(MaCB_End,'') MaCB_End,IsNull(Ma_So_Old,'') Ma_So_Old,IsNull(IP,'') IP From ChiNhanh Where Status = 1 And Ma_So='{0}' And Id_Goc <= 1", DONVI.Trim)
            End If

            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                Dim _ReaderRP As System.Data.SqlClient.SqlDataReader = SoftSqlHelper.ExcuteDataReader(conn_obj, CommandType.Text, sSQL, CType(Nothing, System.Data.SqlClient.SqlParameterCollection))
                While (_ReaderRP.Read())
                    Dim _ChiNhanh As ChiNhanh = New ChiNhanh()
                    _ChiNhanh.GetData(_ReaderRP)
                    ARL_List.Add(_ChiNhanh)
                End While
            End Using
            Return ARL_List
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy dữ liệu Danh mục Chi Nhánh: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    Public Function GetChiNhanh_ForId(ByVal _Id As Int32) As ChiNhanh
        Try
            Dim sSQL As String = ""
            sSQL = String.Format("Select Id,IsNull(Ten_Goi,'') Ten_Goi,IsNull(Id_Goc,0) Id_Goc,IsNull(Ma_So,'') Ma_So,IsNull(Ten_VT,'') Ten_VT,IsNull(Dia_Chi,'') Dia_Chi,IsNull(Dien_Thoai,'') Dien_Thoai,IsNull(Email,'') Email,IsNull(Website,'') Website,Status,IsNull(SoLg_XaPhuong,0) SoLg_XaPhuong,IsNull(MaDV,'') MaDV,IsNull(MaCB_Begin,'') MaCB_Begin,IsNull(MaCB_End,'') MaCB_End,IsNull(Ma_So_Old,'') Ma_So_Old,IsNull(IP,'') IP From ChiNhanh Where Id = {0}", _Id)
            Using conn_obj As System.Data.SqlClient.SqlConnection = DbCommon.GetSqlConnection()
                Dim _ReaderRP As System.Data.SqlClient.SqlDataReader = SoftSqlHelper.ExcuteDataReader(conn_obj, CommandType.Text, sSQL, SoftSqlHelper.CreateParameter("@_Id", _Id, ParameterDirection.Input))
                While (_ReaderRP.Read())
                    Dim _ChiNhanh As ChiNhanh = New ChiNhanh()
                    _ChiNhanh.GetData(_ReaderRP)
                    Return _ChiNhanh
                End While
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy dữ liệu Danh mục Chi Nhánh theo Id truyền vào: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Hàm thực hiện Lấy Dữ liệu Danh sách cần Bind vào ComboBox hoặc ListBox Chi nhánh
    ''' </summary>
    ''' <param name="pLoaiDs">Chỉ số xác định Kiểu danh sách cần lấy. Giá trị quy ước:
    '''                     1 - Danh sách bao gồm: HSC, Các đơn vị TTCNTT, TTĐT, SGD, Chi nhánh và Các đơn vị cấp Quận/Huyện
    '''                     2 - Danh sách bao gồm: HSC; Các đơn vị trực thuộc.
    ''' </param>
    ''' <param name="pNone">Chuỗi giá trị tiêu đề đầu tiên. Ex: "---Danh sách đơn vị ---"</param>
    ''' <returns>Danh sách cần lấy</returns>
    ''' <remarks>Ex: _ChiNhanhBLL.GetListComBo_ChiNhanh(1, "---Danh sách đơn vị ---")</remarks>
    Public Function GetListComBo_ChiNhanh(ByVal pLoaiDs As Byte, ByVal pNone As String) As DataTable
        Dim sSQL As String = ""
        Select Case pLoaiDs
            Case 1      'Danh sách Chi Nhánh bao gồm: Toàn bộ: HSC, Các đơn vị TTCNTT, TTĐT, SGD, Chi nhánh và Các đơn vị cấp Quận/Huyện
                If String.IsNullOrEmpty(pNone) Then
                    'sSQL = "Select Id As Value,(Case When (Id_Goc > 5) Then ('  + ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display"
                    'sSQL = sSQL & " From ChiNhanh Where Status = 1 "
                    'If TRUCTHUOC = 0 Then   'Chạy tại Chi nhánh
                    '    sSQL = sSQL & "      And Ma_So = '" & DONVI.Trim() & "' Or Id_Goc In (Select Id From ChiNhanh Where Ma_So = '" & DONVI.Trim() & "' And Status = 1) "
                    'End If
                    'sSQL = sSQL & " Order By (Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"


                    sSQL = "Select Id As Value,(Case When (Id_Goc > 5) Then ('  + ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display"
                    sSQL = sSQL & " From ("
                    sSQL = sSQL & "        Select Id,Id_Goc,Ma_So,Ten_Goi As Ten_Goi,Status,(Case When A.Id_Goc In (0,1) Then A.Ma_So Else (Select Distinct X.Ma_So From ChiNhanh X Where X.Id=A.Id_Goc And Status=1) End) MaCN From ChiNhanh A Where A.Status=1"
                    If TRUCTHUOC = 0 Then   'Chạy tại Chi nhánh
                        sSQL = sSQL & "      And Ma_So = '" & DONVI.Trim() & "' Or Id_Goc In (Select Id From ChiNhanh Where Ma_So = '" & DONVI.Trim() & "' And Status = 1) "
                    End If
                    sSQL = sSQL & "      ) As Z "
                    sSQL = sSQL & "      Order By MaCN,(Case When MaCN = Ma_So Then 0 Else 1 End),Ma_So"
                Else
                    sSQL = "Select Id As Value,(Case When (Id_Goc > 5) Then ('  + ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display"
                    sSQL = sSQL & " From ("
                    sSQL = sSQL & "        Select 0 As Id,0 Id_Goc,'000000' Ma_So,N'" & pNone & "' As Ten_Goi,1 Status,'000000' MaCN"
                    sSQL = sSQL & "        Union"
                    sSQL = sSQL & "        Select Id,Id_Goc,Ma_So,Ten_Goi As Ten_Goi,Status,(Case When A.Id_Goc In (0,1) Then A.Ma_So Else (Select Distinct X.Ma_So From ChiNhanh X Where X.Id=A.Id_Goc And Status=1) End) MaCN From ChiNhanh A Where A.Status=1"
                    If TRUCTHUOC = 0 Then   'Chạy tại Chi nhánh
                        sSQL = sSQL & "      And Ma_So = '" & DONVI.Trim() & "' Or Id_Goc In (Select Id From ChiNhanh Where Ma_So = '" & DONVI.Trim() & "' And Status = 1) "
                    End If
                    sSQL = sSQL & "      ) As Z "
                    sSQL = sSQL & "      Order By MaCN,(Case When MaCN = Ma_So Then 0 Else 1 End),Ma_So"
                End If

            Case 2      'Danh sách Chi Nhánh bao gồm 2 cấp: Hội sở chính; Các đơn vị trực thuộc
                If String.IsNullOrEmpty(pNone) Then
                    sSQL = " Select 1 As Value, N'Hội sở chính' As Display"
                    sSQL = sSQL & " Union "
                    sSQL = sSQL & " Select 2 As Value, N'Các đơn vị trực thuộc' As Display"
                Else
                    sSQL = "Select 0 As Value, N'" & pNone & "' As Display"
                    sSQL = sSQL & " Union "
                    sSQL = sSQL & " Select 1 As Value, N'Hội sở chính' As Display"
                    sSQL = sSQL & " Union "
                    sSQL = sSQL & " Select 2 As Value, N'Các đơn vị trực thuộc' As Display"
                End If
            Case 3      'Danh sách bao gồm 2 cấp: Hội sở chính; Danh sách các Chi nhánh Tỉnh/TP
                If String.IsNullOrEmpty(pNone) Then
                    sSQL = "Select Id As Value,(Case When (Id_Goc > 5) Then ('  + ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display"
                    sSQL = sSQL & " From ChiNhanh Where Status = 1 And Id_Goc In (0,1) "
                    If TRUCTHUOC = 0 Then   'Chạy tại Chi nhánh
                        sSQL = sSQL & "      And Ma_So = '" & DONVI.Trim() & "' "
                    End If
                    sSQL = sSQL & "Order By (Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"
                Else
                    sSQL = "Select Id As Value,(Case When (Id_Goc > 5) Then ('  + ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display"
                    sSQL = sSQL & " From ("
                    sSQL = sSQL & "        Select 0 As Id,0 Id_Goc,'000000' Ma_So,N'" & pNone & "' As Ten_Goi,1 Status"
                    sSQL = sSQL & "        Union"
                    sSQL = sSQL & "        Select Id,Id_Goc,Ma_So,Ten_Goi As Ten_Goi,Status From ChiNhanh Where Status=1 And Id_Goc In (0,1) "
                    If TRUCTHUOC = 0 Then   'Chạy tại Chi nhánh
                        sSQL = sSQL & "      And Ma_So = '" & DONVI.Trim() & "' "
                    End If
                    sSQL = sSQL & "      ) As Z "
                    sSQL = sSQL & "      Order By (Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"
                End If
        End Select
        Try
            Using db_List As DataTable = SoftSqlHelper.ExecuteForSQL(sSQL)
                Return db_List
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi lấy Danh sách Đơn vị: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return Nothing
        End Try
    End Function

End Class
