Imports System.Data
Imports System.Text
Imports System.Data.SqlClient

Module funcHS_Luong

    ''' <summary>
    ''' Các function dùng chung trong Hồ sơ lương cán bộ
    ''' </summary>
    ''' <remarks></remarks>

    Public Function getSystemVar(ByVal nameVar As String) As String
        Dim strValue As String = ""
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        strSql = "SELECT VarValue FROM sysvar WHERE VarName='" & nameVar & "'"
        strValue = dbconn.getString(strSql).Trim
        Return strValue
    End Function

    ''' <summary>
    ''' Danh sách đơn vị (Chi nhánh)
    ''' </summary>
    ''' <param name="notAll"></param>
    ''' <param name="NoneRow"></param>
    ''' <param name="ChiNhanh_ViewOnly">Chỉ hiển thị danh sách đơn vị cấp tỉnh trở lên</param>
    ''' <param name="isFrmKeHoach">Hiển thị danh sách theo form Kế hoach đơn vị</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listDonvi(Optional ByVal notAll As Boolean = False, Optional ByVal NoneRow As Boolean = False, Optional ByVal ChiNhanh_ViewOnly As Boolean = False, Optional ByVal isFrmKeHoach As Boolean = False, Optional ByVal fullStatus As Boolean = False, Optional ByVal DisplayNoneRow As String = "-----Chọn tất cả Chi Nhánh-----") As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim strStatus As String = ""
        Dim dbconn As DBAccess = New DBAccess
        If fullStatus Then
            strStatus = " (status=1 or status=0)"
        Else
            strStatus = " status=1"
        End If

        Dim strWhere As String = " WHERE " & strStatus
        If (DONVI = "000100") Then
            If ChiNhanh_ViewOnly Then strWhere = strWhere & " AND id_goc in (0,1) "
        Else 'ElseIf notAll Then
            strWhere = strWhere & " AND (ma_so='" & DONVI.Trim & "' or id_goc in (SELECT id FROM ChiNhanh WHERE ma_so='" & DONVI.Trim & "' AND " & strStatus & ")) "
        End If
        'If notAll Then strWhere = strWhere & " AND (ma_so='" & DONVI.Trim & "' or id_goc in (SELECT id FROM ChiNhanh WHERE ma_so='" & DONVI.Trim & "' AND " & strStatus & ")) "
        If isFrmKeHoach Then
            If DONVI = gMaDonViTW Then
                strSql = " SELECT 0 as Value, N'" & DisplayNoneRow & "' as Display UNION " & _
                         " SELECT Id as Value, Ten_Goi as Display FROM chinhanh" & strWhere & " Order by Value"
            Else
                strSql = " SELECT 0 as Value, N'" & DisplayNoneRow & "' as Display UNION " & _
                         " SELECT Id as Value, (CASE WHEN (Id >5 And Id_Goc=1) Then N'Hội sở  tỉnh' ELSE Ten_Goi END) as Display FROM ChiNhanh" & strWhere & " Order by Value"
            End If
        Else
            If NoneRow Then
                ' [dbo].[listDonVi] (@DisplayNoneRow nvarchar(256))
                strSql = " SELECT 0 as Value, N'" & DisplayNoneRow & "' as Display, '     ' As Ma_So UNION " & _
                         " SELECT Id as Value, (Case When (id_goc>5) Then ('  '+ Ten_Goi) Else Ten_Goi End) as Display, Ma_So FROM chinhanh" & strWhere & " Order by Substring(Ma_So,1,4),(Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"
            Else
                strSql = " SELECT Id As Value, (Case When (Id_Goc > 5) Then ('  ' + Replace(Ten_Goi,N'thành phố',N'TP.')) Else Ten_Goi End) As Display From ChiNhanh" & strWhere & " Order by Substring(Ma_So,1,4),(Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"
            End If
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function
    Public Function listDonvi_New(Optional ByVal notAll As Boolean = False, Optional ByVal NoneRow As Boolean = False, Optional ByVal ChiNhanh_ViewOnly As Boolean = False, Optional ByVal isFrmKeHoach As Boolean = False, Optional ByVal fullStatus As Boolean = False, Optional ByVal DisplayNoneRow As String = "-----Chọn tất cả Chi Nhánh-----") As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim strStatus As String = ""
        Dim dbconn As DBAccess = New DBAccess
        If fullStatus Then
            strStatus = " (status=1 or status=0)"
        Else
            strStatus = " status=1"
        End If
        Dim strWhere As String = " WHERE " & strStatus
        If ChiNhanh_ViewOnly Then strWhere = strWhere & " AND id_goc in (0,1) "
        If notAll Then strWhere = strWhere & " AND (ma_so='" & DONVI.Trim & "' or id_goc in (SELECT id FROM ChiNhanh WHERE ma_so='" & DONVI.Trim & "' AND " & strStatus & ")) "
        If isFrmKeHoach Then
            If DONVI = gMaDonViTW Then
                strSql = " SELECT 0 as Value, N'" & DisplayNoneRow & "' as Display UNION " & _
                         " SELECT id as Value, ten_goi as Display FROM chinhanh" & strWhere & " Order by Value"
            Else
                strSql = " SELECT 0 as Value, N'" & DisplayNoneRow & "' as Display UNION " & _
                         " SELECT id as Value, (CASE WHEN (id>5 and id_goc=1) Then N'Hội sở  tỉnh' ELSE ten_goi END) as Display FROM chinhanh" & strWhere & " Order by Value"
            End If
        Else
            If NoneRow Then
                ' [dbo].[listDonVi] (@DisplayNoneRow nvarchar(256))
                strSql = " SELECT 0 as Value, N'" & DisplayNoneRow & "' as Display, '     ' as ma_so UNION " & _
                         " SELECT id as Value, (CASE WHEN (id_goc>5) Then ('  '+ Replace(Ten_Goi,N'thành phố',N'TP.')) ELSE ten_goi END) as Display, ma_so FROM chinhanh" & strWhere & " Order By Substring(Ma_So,1,4),(Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"
            Else
                strSql = " SELECT id as Value, (CASE WHEN (id_goc>5) Then ('  '+ Replace(Ten_Goi,N'thành phố',N'TP.')) ELSE ten_goi END) as Display FROM chinhanh" & strWhere & " Order By Substring(Ma_So,1,4),(Case When Id_Goc In (0,1) Then 0 Else 1 End),Ma_So"
            End If
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function
    ''' <summary>
    ''' Lấy tên chi nhánh
    ''' </summary>
    ''' <param name="vIdDonvi">Id chi nhánh</param>
    ''' <param name="isKH">sử dụng trong form Kế hoạch đơn vị</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDonvi(ByVal vIdDonvi As Integer, Optional ByVal isKH As Boolean = False) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            If isKH Then
                If vIdDonvi = 1 Then
                    Return dbconn.getString("SELECT ten_goi FROM CHINHANH WHERE id=" & vIdDonvi)
                Else
                    Return dbconn.getString("SELECT (CASE WHEN (id>5 and id_goc=1) Then N'Hội sở  tỉnh' ELSE ten_goi END) as ten_goi FROM CHINHANH WHERE id=" & vIdDonvi)
                End If
            Else
                Return dbconn.getString("SELECT ten_goi FROM CHINHANH WHERE id=" & vIdDonvi)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Lấy tên đơn vị + tên đơn vị cấp trên quản lý
    ''' </summary>
    ''' <param name="vIdDonvi">Id Don vị hiển thị tên</param>
    ''' <param name="vIdDonViRoot">Id Đơn vị đang chạy ứng dụng</param>
    ''' <param name="CharJoin">Kí tự dùng để kết nối</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDonviPath(ByVal vIdDonvi As Integer, ByVal vIdDonViRoot As Integer, ByVal CharJoin As String) As String
        Try
            If vIdDonvi = 0 Then
                Return ""
                Exit Function
            End If
            Dim strTmp As String = ""
            Dim dbconn As DBAccess = New DBAccess
            Dim dt As DataTable = New DataTable
            dt = dbconn.SelectDBRows("SELECT Id_goc, Ten_goi FROM Chinhanh WHERE id=" & vIdDonvi)
            If dt.Rows.Count > 0 Then
                strTmp = dt.Rows(0).Item("ten_goi")
                If (dt.Rows(0).Item("Id_goc") <> vIdDonViRoot And vIdDonvi <> vIdDonViRoot) Then
                    strTmp += CharJoin & dbconn.getString("SELECT Ten_goi FROM Chinhanh WHERE id=" & dt.Rows(0).Item("Id_goc"))
                End If
            End If
            Return strTmp
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Lấy mã số chi nhánh
    ''' </summary>
    ''' <param name="vIdDonvi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDonvi_Ma(ByVal vIdDonvi As Integer) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getString("SELECT ma_so FROM CHINHANH WHERE id=" & vIdDonvi)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getDonvi_Name(ByVal vKey As String, ByVal vValue As String) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getString("SELECT ten_goi FROM CHINHANH WHERE " & vKey & "='" & vValue & "'")
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Danh sách các Quốc gia
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listQuocGia() As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess = New DBAccess
        strSql = "SELECT id as Value, ten_goi as Display FROM QuocGia order by ten_goi"
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Lấy tên Tên Quốc gia
    ''' </summary>
    ''' <param name="vIdQuocGia"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getQuocGia(ByVal vIdQuocGia As Integer) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getString("SELECT ten_goi FROM QuocGia WHERE id=" & vIdQuocGia)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Danh sách phòng ban theo đơn vị
    ''' </summary>
    ''' <param name="vIdDonvi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listPhong(ByVal vIdDonvi As String, Optional ByVal NoneRow As Boolean = False, Optional ByVal DisplayPGD As Boolean = False, Optional ByVal only_P_Actived As Boolean = True) As DataTable
        Dim dt As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim currTRUCTHUOC As String
        Dim maDonvi As String
        Dim strSql As String = ""
        Dim strStatusActive As String = " AND status=1 "
        If Not only_P_Actived Then strStatusActive = ""
        maDonvi = dbconn.getString("SELECT ma_so FROM Chinhanh WHERE id=" & vIdDonvi)
        currTRUCTHUOC = getTrucThuoc(maDonvi)
        If NoneRow Then
            If DisplayPGD Then
                'Ap dung cho Truong hop ChiLuong
                If currTRUCTHUOC = "2" Then
                    strSql = " SELECT 'PB_0' as Value, N'0-----Chọn tất cả Phòng ban trực thuộc hội sở tỉnh-----0' as Display " & _
                             " UNION " & _
                             " SELECT 'PB_'+ltrim(str(id)) as Value, ten_phong as Display FROM PHONGBAN WHERE charindex('2',truc_thuoc)>0 " & strStatusActive & _
                             " UNION " & _
                             " SELECT 'DV_'+ltrim(str(Id)) as Value, ten_goi as Display FROM CHINHANH WHERE ID_goc<>1 AND Id_goc=" & vIdDonvi & _
                             " UNION " & _
                             " SELECT 'NH_0' as Value, N'w-----Nhóm cán bộ-----w' as Display " & _
                             " Order by Display"
                Else
                    strSql = " SELECT 'PB_0' as Value, N'0-----Chọn tất cả Phòng ban-----0' as Display " & _
                             " UNION " & _
                             " SELECT 'PB_'+ltrim(str(id)) as Value, ten_phong as Display FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 " & strStatusActive & _
                             " UNION " & _
                             " SELECT 'DV_'+ltrim(str(Id)) as Value, ten_goi as Display FROM CHINHANH WHERE ID_goc<>1 AND Id_goc=" & vIdDonvi & _
                             " UNION " & _
                             " SELECT 'NH_0' as Value, N'w-----Nhóm cán bộ-----w' as Display " & _
                             " Order by Display"
                End If
            Else
                strSql = " SELECT '0' as Value, N'-----Chọn tất cả Phòng ban-----' as Display UNION " & _
                         " SELECT id as Value, ten_phong as Display FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 " & strStatusActive
            End If
        Else
            If DisplayPGD Then
                'Ap dung cho Truong hop ChiLuong
                strSql = " SELECT 'PB_'+ltrim(str(id)) as Value, ten_phong as Display FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 " & strStatusActive & _
                         " UNION " & _
                         " SELECT 'DV_'+ltrim(str(Id)) as Value, ten_goi as Display FROM CHINHANH WHERE ID_goc<>1 AND Id_goc=" & vIdDonvi & _
                         " Order by Display, status desc "
            Else
                strSql = "SELECT id as Value, ten_phong as Display FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 " & strStatusActive & " Order by Ma_so, status desc"
            End If
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Lấy tên phòng ban, loại bỏ phần chú thích trong tên 
    ''' </summary>
    ''' <param name="vIdPhong"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getPhong(ByVal vIdPhong As Integer) As String
        Try
            If vIdPhong = 0 Then
                Return ""
            End If
            Dim strReturn As String = ""
            Dim idx As Integer = 0
            Dim dbconn As DBAccess = New DBAccess
            strReturn = dbconn.getString("SELECT ten_phong FROM PHONGBAN WHERE id=" & vIdPhong)
            idx = strReturn.IndexOf("(")
            If idx > 0 Then
                Return strReturn.Substring(0, idx - 1)
            Else
                Return strReturn
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Lấy tên phòng ban + tên các đơn vị quản lý theo từng cấp trừ cấp đang hiển thị
    ''' </summary>
    ''' <param name="vIdPhong">Id Phòng</param>
    ''' <param name="vIdDonVi">Id đơn vị cấp trực tiếp của phòng</param>
    ''' <param name="vIdDonviRoot">Id Đơn vị đang thực hiện hiển thị danh sách</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getPhongPath(ByVal vIdPhong As Integer, ByVal vIdDonVi As Integer, ByVal vIdDonviRoot As Integer, ByVal CharJoin As String) As String
        Try
            If vIdPhong = 0 Then
                Return ""
            End If
            If vIdDonVi <> vIdDonviRoot Then
                Return getPhong(vIdPhong) & CharJoin & getDonviPath(vIdDonVi, vIdDonviRoot, ", ")
            Else
                Return getPhong(vIdPhong)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Lãnh đạo có quyền quyết định lương cán bộ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listNguoiQDLuong() As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        strSql = " SELECT b1.idCanbo as id, b1.hoten as name, b2.idChucvu_Moi " & _
                 " FROM HS_Canbo b1, QDnhansu b2 " & _
                 " WHERE b1.idCanbo=b2.idCanbo and (idDonvi_Moi =" & IdDONVI & " or idDonvi_Moi in (SELECT id_goc FROM Chinhanh WHERE id=" & IdDONVI & "))" & _
                 "   and b1.idcanbo not in (SELECT idcanbo FROM hs_CBthoiviec WHERE IsQD_NHCS =1 ) " & _
                 "   and b2.idChucvu_Moi in (SELECT id FROM Danhmuc WHERE id_goc=14 and ma_so in ('1401','1402','1403','1404','1410','1411'))"
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Danh sách CHức vụ có quyền kí các quyết đinh đối với cán bộ: QĐ bổ nhiệm,..., lương,..
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listChucVuQuyenRaQD() As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        strSql = "SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE  id_goc=14 and ma_so in ('1401','1402','1403','1404','1410','1411', '1427', '1428', '1440', '1441', '1442', '1421', '1418', '1419', '1420', '1446', '1447', '1448')"
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    Public Function listChucVu(ByVal vIDDonVi As Integer, Optional ByVal NoneRow As Boolean = False) As DataTable
        Dim dt As DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess

        If NoneRow Then
            strSql = " SELECT '0' as Value, '' as Display UNION "
        End If
        If getDonvi_Ma(vIDDonVi) = gMaDonViTW Then
            strSql &= "SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE id_goc=14 and status=1 and ma_so not in ('1428', '1431', '1439') Order by DisPlay"
            'strSql &= "SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE id_goc=14 and status=1 and ma_so not in ('1410', '1428', '1411', '1431', '1439') Order by DisPlay"
        Else
            strSql &= "SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE id_goc=14 and status=1 and  ma_so not in ('1401', '1402', '1403', '1404', '1405', '1406', '1407', '1408', '1409', '1427', '1429', '1430') Order by DisPlay"
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Danh sach chuc vu ky quyet dinh khen thuong
    ''' </summary>
    ''' <param name="idDanhmucGoc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listChucVuQDKhenThuong(ByVal idDanhmucGoc As Integer) As DataTable
        Dim dt As DataTable = New DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        strSql = " SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE id_goc=" & idDanhmucGoc & " and ma_so in('0401','0403','0404','0405','0406','0407','0408','0409','0411','0417') AND status=1"
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Liệt kê danh sách Theo từng Loại;
    ''' Luong Phucap=31;
    ''' Luong Bo sung = 42;
    ''' Loại quyết định lương = 43;
    ''' Chức vụ trong cơ quan = 14;
    ''' Chuyên môn nghiệp vụ =  12;
    ''' Loại Quyết định = 15
    ''' Trình độ chính trị = 36;
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listDanhmuc(ByVal idDanhmucGoc As Integer, Optional ByVal NoneRow As Boolean = False, Optional ByVal OrderBy As Boolean = False, Optional ByVal IdDanhMuc As Integer = 0) As DataTable
        Dim dt As DataTable = New DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        If NoneRow Then
            strSql = " SELECT '0' as Value, '' as Display UNION "
        End If
        strSql &= " SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE  id_goc=" & idDanhmucGoc & "  AND status=1"
        If IdDanhMuc > 0 Then
            strSql &= " AND Id = " & IdDanhMuc
        End If
        If OrderBy Then
            strSql &= " Order by Display"
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Liệt kê danh sách QD Nhân sự (không gồm Qd Cách chức: có ghi but không cho hiển thị và nhập vào trong phần QD nhân sự
    '''                               QD cách chức được nhập trong phần Kỉ luật, sau đó đẩy dữ liệu sang QD Nhân sự    )
    ''' </summary>
    ''' <param name="NoneRow"></param>
    ''' <param name="OrderBy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listDanhmucQDNhanSu(Optional ByVal NoneRow As Boolean = False, Optional ByVal OrderBy As Boolean = False) As DataTable
        Dim dt As DataTable = New DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        If NoneRow Then
            strSql = " SELECT '0' as Value, '' as Display UNION "
        End If
        strSql &= " SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE  id_goc=15 AND ma_so<>'1510' AND status=1"
        If OrderBy Then
            strSql &= " Order by Display"
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Danh sách loại QĐ lương hiển thị trong phần Hồ sơ lương
    ''' </summary>
    ''' <param name="NoneRow"></param>
    ''' <param name="OrderBy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listLoaiQDLuong(Optional ByVal NoneRow As Boolean = False, Optional ByVal OrderBy As Boolean = False) As DataTable
        Dim dt As DataTable = New DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        If NoneRow Then
            strSql = " SELECT '0' as Value, '' as Display UNION "
        End If
        strSql &= "SELECT id as Value, ten_goi as Display  FROM Danhmuc WHERE id_goc=43 AND (substring(ma_so,3,1)='0' or substring(ma_so,3,1)='5') AND status=1"
        If OrderBy Then
            strSql &= " Order by Display"
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    ''' <summary>
    ''' Liet ke danh muc thi dua khen thuong
    ''' </summary>
    ''' <param name="CaNhan"></param>
    ''' <param name="TapThe"></param>
    ''' <param name="vKT_ChuyenMon">=1: danh sach khen thuong theo chuyen mon; =0 danh sach khen thuong theo Cong doan</param>
    ''' <param name="NoneRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listKhenThuong(ByVal CaNhan As Boolean, ByVal TapThe As Boolean, ByVal vKT_ChuyenMon As Int16, Optional ByVal NoneRow As Boolean = False) As DataTable
        Dim dt As DataTable = New DataTable
        Dim strSql As String = ""
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        If NoneRow Then
            strSql = " SELECT '0' as Value, '---------Danh sách khen thưởng---------' as Display UNION "
        End If
        If CaNhan And TapThe Then
            strSql = "SELECT IdTDKT as Value, case when CN_TT=1 then N'Cá nhân => '+DanhHieu_HinhThuc else N'Tập thể => '+DanhHieu_HinhThuc end as Display FROM Thiduakhenthuong WHERE KT_ChuyenMon=" & vKT_ChuyenMon & " and status=1"
        Else
            strSql = "SELECT IdTDKT as Value, case when CN_TT=1 then N'Cá nhân => '+DanhHieu_HinhThuc else N'Tập thể => '+DanhHieu_HinhThuc end as Display FROM Thiduakhenthuong WHERE KT_ChuyenMon=" & vKT_ChuyenMon & " and status=1 and CN_TT=" & IIf(CaNhan, 1, 2)
        End If
        dt = dbconn.SelectDBRows(strSql)
        Return dt
    End Function

    Public Sub getLanhdaobyChucvu(ByVal idChucvu As Integer, ByRef idLanhdao As String, ByRef Hoten As String, Optional ByVal idNhanvien As String = "")
        Dim dt As DataTable
        Dim strSql As String
        Dim _idDonVi As Integer
        Dim dbconn As DBAccess
        dbconn = New DBAccess
        Try
            If idNhanvien <> "" Then
                _idDonVi = dbconn.getNumber("SELECT idDonvi FROM HS_Canbo WHERE idCanbo='" & idNhanvien & "'")
            End If
            strSql = "SELECT top 1 b1.idCanbo, hoten FROM HS_Canbo b1, QDnhansu b2 WHERE b1.idCanbo=b2.idCanbo and idChucvu_moi=" & idChucvu
            If idChucvu = 304 Or idChucvu = 305 Then
                strSql &= " and idDonvi_Moi=" & _idDonVi
            End If
            strSql &= " order by ngayHL desc"
            dt = dbconn.getDataTable(strSql)
            idLanhdao = dt.Rows(0).Item("idCanbo")
            Hoten = dt.Rows(0).Item("hoten")
        Catch ex As Exception
            idLanhdao = 0
            Hoten = ""
        End Try
    End Sub

    Public Function getFullName(ByVal IdCanBo As String, ByVal TS As Boolean) As String
        Dim dbconn As DBAccess = New DBAccess
        Try
            If TS Then
                Return dbconn.getString("SELECT Hoten FROM HSCB_TS WHERE Id='" & IdCanBo.Trim & "'")
            Else
                Return dbconn.getString("SELECT Hoten FROM HS_Canbo WHERE IdCanbo='" & IdCanBo.Trim & "'")
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Lấy tên cán bộ + thuộc phòng ban, đơn vị
    ''' </summary>
    ''' <param name="IdCanBo"></param>
    ''' <param name="TS"></param>
    ''' <param name="vIdPhong">Id Phòng trực thuộc</param>
    ''' <param name="vIdDonvi">Id Đơn vị trực thuộc</param>
    ''' <param name="vIdDonviRoot">Id Đon vị đang chạy ứng dụng</param>
    ''' <param name="CharJoin">Kí tự kết nối</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFullNamePath(ByVal IdCanBo As String, ByVal TS As Boolean, ByVal vIdPhong As Integer, ByVal vIdDonvi As Integer, ByVal vIdDonviRoot As Integer, ByVal CharJoin As String) As String
        Try
            Return getFullName(IdCanBo, TS) & CharJoin & getPhongPath(vIdPhong, vIdDonvi, vIdDonviRoot, ", ")
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Lấy mã cán bộ theo ID
    ''' </summary>
    ''' <param name="vIdCanBo"></param>
    ''' <param name="TS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCanBo_Ma(ByVal vIdCanBo As String, ByVal TS As Boolean) As String
        Dim dbconn As DBAccess = New DBAccess
        Try
            If TS Then
                Return dbconn.getString("SELECT MaCB FROM HSCB_TS WHERE Id='" & vIdCanBo.Trim & "'")
            Else
                Return dbconn.getString("SELECT MaCB FROM HS_Canbo WHERE IdCanbo='" & vIdCanBo.Trim & "'")
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' lẫy ID cán bộ theo mã của cán bộ
    ''' </summary>
    ''' <param name="vMaCanBo"></param>
    ''' <param name="TS"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCanBo_ID(ByVal vMaCanBo As String, ByVal TS As Boolean) As String
        Dim dbconn As DBAccess = New DBAccess
        Try
            If TS Then
                Return dbconn.getString("SELECT [Id] FROM HSCB_TS WHERE MaCB='" & vMaCanBo.Trim & "'")
            Else
                Return dbconn.getString("SELECT [IdCanbo] FROM HS_Canbo WHERE MaCB='" & vMaCanBo.Trim & "'")
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Thông tin Chức vụ của từng cán bộ
    ''' </summary>
    ''' <param name="idCanbo"></param>
    ''' <param name="idChucvu"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getChucVuCanBo(ByVal idCanbo As String, ByVal idChucvu As Integer) As String
        Dim dbconn As DBAccess = New DBAccess
        Try
            Return dbconn.getString("SELECT ten_goi +': ' + (SELECT Hoten FROM HS_Canbo WHERE IdCanbo='" & idCanbo.Trim & "') as CVCB FROM Danhmuc WHERE id_goc=14 and id=" & idChucvu)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getDanhmuc_Name(ByVal idDanhmucGoc As Integer, ByVal id As Integer) As String
        Dim dbconn As DBAccess = New DBAccess
        Try
            Return dbconn.getString("SELECT ten_goi FROM Danhmuc WHERE id_goc=" & idDanhmucGoc & " and id=" & id)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getDanhmuc_ID(ByVal DanhMuc_Maso As String) As Integer
        Try
            Dim db As DBAccess = New DBAccess
            Return db.getNumber("SELECT id FROM Danhmuc WHERE ma_so='" & DanhMuc_Maso & "'")
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function getDanhmuc_MaSo(ByVal DanhMuc_ID As String) As String
        Try
            Dim db As DBAccess = New DBAccess
            Return db.getString("SELECT ma_so FROM Danhmuc WHERE id =" & DanhMuc_ID)
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Hàm trả lại giá trị - Trực thuộc
    ''' </summary>
    ''' <param name="ma_so">Mã số của chi nhánh</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTrucThuoc(ByVal ma_so As String) As String
        Dim strReturn As String = ""
        Dim _SqlHelper As New DBAccess
        Select Case ma_so
            Case gMaDonViTW
                strReturn = 1 'Globals.GetAppSetting("HSC").Trim()
            Case "000196"
                strReturn = 5 'Globals.GetAppSetting("TTCNTT").Trim()
            Case "000197"
                strReturn = 7 'Globals.GetAppSetting("TTDT").Trim()
            Case "000101"
                strReturn = 9 'Globals.GetAppSetting("SGD").Trim()
                'Chữ bổ sung Cơ sở đâò tạo
            Case "002821", "001114", "002734", "003799", "004532", "005399"
                strReturn = 6 'Cơ sở đào tạo
            Case Else
                Dim iCap As Integer = _SqlHelper.getNumber("SELECT id_goc FROM ChiNhanh WHERE ma_so = '" & ma_so & "'")
                If iCap = 1 Then
                    strReturn = 2 'Globals.GetAppSetting("TINH").Trim()
                Else
                    strReturn = 4 'Globals.GetAppSetting("HUYEN").Trim()
                End If
        End Select
        Return strReturn
    End Function

    Public Sub bindTreeThangBangLuong(ByVal treeviewName As TreeView)
        Dim rootNode As TreeNode
        Dim childNode1, childNode2, childNode3, childNode4 As TreeNode
        rootNode = treeviewName.Nodes.Add(1, "Hệ thống thang, bảng lương")
        rootNode.Tag = "1"
        childNode1 = rootNode.Nodes.Add(2, "Nghị định lương")
        childNode1.Tag = "2"
        childNode2 = rootNode.Nodes.Add(3, "Bảng lương")
        childNode2.Tag = "3"
        childNode3 = rootNode.Nodes.Add(4, "Ngạch lương")
        childNode3.Tag = "4"
        childNode4 = rootNode.Nodes.Add(5, "Bậc lương")
        childNode4.Tag = "5"
    End Sub

    Public Function getNodeThangBangLuong(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer) As TreeNode
        Dim i As Integer
        Dim dt As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode
        Dim strsql As String = ""
        Select Case IdParentNode
            Case 2
                strsql = "SELECT IdNDLuong as id, 0 as parent, MaND as ma_so, TenND as ten_goi FROM NGHIDINHLUONG "
            Case 3
                strsql = "SELECT IDBangLuong as id, IDND_Luong as parent, BangLuong as ma_so, Mota as ten_goi FROM BANGLUONG "
            Case 4
                strsql = "SELECT IdNgachLuong as id, IDBangLuong as parent, NgachLuong as ma_so, Mota as ten_goi FROM NGACHLUONG "
            Case 5
                strsql = "SELECT IdBacLuong as id, IdNgachLuong as parent, BacLuong as ma_so, Heso as ten_goi FROM BACLUONG "
        End Select
        dt = db.SelectDBRows(strsql)
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                childNode = parentNode.Nodes.Add(dt.Rows(i).Item("id"), dt.Rows(i).Item("ten_goi"))
                childNode.Tag = "DV_" & dt.Rows(i).Item("id") & "_" & dt.Rows(i).Item("ma_so")
            Next
        End If
        Return parentNode
    End Function

    Public Function getNodeThangBangLuong1(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, ByVal code As Integer) As TreeNode
        Dim i As Integer
        Dim dt As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode
        Dim strsql As String = ""
        Select Case code
            Case 2
                strsql = "SELECT IdNDLuong as id, 0 as parent, MaND as ma_so, TenND as ten_goi FROM NGHIDINHLUONG"
            Case 3
                strsql = "SELECT IDBangLuong as id, IDND_Luong as parent, BangLuong as ma_so, Mota as ten_goi FROM BANGLUONG WHERE IDND_Luong=" & IdParentNode
            Case 4
                strsql = "SELECT IdNgachLuong as id, IDBangLuong as parent, NgachLuong as ma_so, Mota as ten_goi FROM NGACHLUONG WHERE IDBangLuong = " & IdParentNode
            Case 5
                strsql = "SELECT IdBacLuong as id, IdNgachLuong as parent, BacLuong as ma_so, Heso as ten_goi FROM BACLUONG WHERE IdNgachLuong= " & IdParentNode
        End Select
        If strsql <> "" Then
            dt = db.SelectDBRows(strsql)
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(dt.Rows(i).Item("id"), dt.Rows(i).Item("ten_goi"))
                    childNode.Tag = "DV_" & dt.Rows(i).Item("id") & "_" & dt.Rows(i).Item("ma_so")
                    getNodeThangBangLuong1(childNode, dt.Rows(i).Item("id"), code + 1)
                Next
            End If
        End If
        Return parentNode
    End Function

    Public Sub bindTreeview(ByVal treeviewName As TreeView, Optional ByVal CBTS As Boolean = False, Optional ByVal Idphong_viewCB As Integer = 0, Optional ByVal IdDonVi_viewCB As Integer = 0)
        Dim dt As DataTable
        Dim db As DBAccess = New DBAccess
        Dim i As Integer
        Dim idx As Integer = 1
        Dim _All As Boolean = True
        Dim rootNode As TreeNode
        dt = db.SelectDBRows("SELECT * FROM CHINHANH WHERE ma_so='" & DONVI & "' AND status=1")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                rootNode = treeviewName.Nodes.Add(dt.Rows(i).Item("id"), dt.Rows(i).Item("ten_goi"))
                rootNode.Tag = "ROOT_" & dt.Rows(i).Item("id") & "_" & dt.Rows(i).Item("ma_so")
                If Not ALL Then _All = False
                getNode(rootNode, dt.Rows(i).Item("id"), dt.Rows(i).Item("ma_so"), _All, CBTS, Idphong_viewCB, IdDonVi_viewCB)
            Next
        End If
    End Sub

    Public Sub bindTreeviewLuong(ByVal treeviewName As TreeView)
        Dim dt, dtCN As DataTable
        Dim db As DBAccess = New DBAccess
        Dim i As Integer
        Dim idx As Integer = 1
        Dim rootNode, childNode As TreeNode

        dt = db.SelectDBRows("SELECT * FROM CHINHANH WHERE ma_so='" & DONVI & "' AND status=1")
        If dt.Rows.Count > 0 Then
            rootNode = treeviewName.Nodes.Add(dt.Rows(0).Item("id"), dt.Rows(0).Item("ten_goi"))
            rootNode.Tag = "ROOT_" & dt.Rows(0).Item("id") & "_" & dt.Rows(0).Item("ma_so")
            getChiLuong_Nam(rootNode, dt.Rows(0).Item("id"))
            dtCN = db.SelectDBRows("SELECT * FROM CHINHANH WHERE id_goc=" & dt.Rows(0).Item("id") & " AND status=1")
            If dtCN.Rows.Count > 0 Then
                For i = 0 To dtCN.Rows.Count - 1
                    childNode = rootNode.Nodes.Add(dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ten_goi"))
                    childNode.Tag = "DV_" & dtCN.Rows(i).Item("id") & "_" & dtCN.Rows(i).Item("ma_so")
                Next
            End If
        End If
    End Sub

    Public Function getChiLuong_Nam(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, Optional ByVal vYear As Integer = 0) As TreeNode
        Dim i As Integer
        Dim dtNam As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode

        dtNam = db.SelectDBRows("SELECT Nam FROM HS_ChiLuong WHERE IdDonVi=" & IdParentNode & " Group by Nam")
        If dtNam.Rows.Count > 0 Then
            For i = 0 To dtNam.Rows.Count - 1
                childNode = parentNode.Nodes.Add(IdParentNode, "Chi năm " & dtNam.Rows(i).Item("Nam"))
                childNode.Tag = "NAM_" & IdParentNode & "_" & dtNam.Rows(i).Item("Nam")
                If vYear = dtNam.Rows(i).Item("Nam") Then
                    getChiLuong_Loai(childNode, IdParentNode, dtNam.Rows(i).Item("Nam"))
                    childNode.Expand()
                End If
            Next
        End If
        Return parentNode
    End Function

    Public Function getChiLuong_Loai(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, ByVal vYear As Integer) As TreeNode
        Dim i As Integer
        Dim dtNam As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode
        Dim strLoaiThuChi As String = ""
        Dim strKy As String = ""
        Dim Ky As Integer
        Dim LoaiCB As Int16 = 0
        Dim strLoaiCB As String = ""

        While parentNode.Nodes.Count > 0
            parentNode.Nodes.Remove(parentNode.FirstNode)
        End While
        dtNam = db.SelectDBRows("SELECT IdHS_ChiLuong, IdloaiThuChi, IdDonVi, Nam, Thang, tmp FROM HS_ChiLuong WHERE IdDonVi=" & IdParentNode & " AND Nam=" & vYear & " Order by Thang")
        If dtNam.Rows.Count > 0 Then
            For i = 0 To dtNam.Rows.Count - 1
                getInfThuChi(dtNam.Rows(i).Item("IdHS_ChiLuong"), dtNam.Rows(i).Item("IdloaiThuChi"), strLoaiThuChi, Ky, strKy, LoaiCB, strLoaiCB)
                childNode = parentNode.Nodes.Add(dtNam.Rows(i).Item("IdHS_ChiLuong"), strLoaiThuChi & dtNam.Rows(i).Item("Thang") & strKy & strLoaiCB & checkChiLuongTmp(dtNam.Rows(i).Item("tmp")))
                childNode.Tag = "CT_" & dtNam.Rows(i).Item("IdHS_ChiLuong") & "_" & dtNam.Rows(i).Item("IdloaiThuChi") & "_" & dtNam.Rows(i).Item("Thang") & "_" & Ky & "_" & dtNam.Rows(i).Item("tmp") & "_" & dtNam.Rows(i).Item("IdDonVi") & "_" & LoaiCB
            Next
        End If
        Return parentNode
    End Function

    Public Sub bindTreeviewLuong_Loai(ByVal treeviewName As TreeView, ByVal vYear As Integer)
        Dim dt, dtCN As DataTable
        Dim db As DBAccess = New DBAccess
        Dim i As Integer
        Dim idx As Integer = 1
        Dim rootNode, childNode As TreeNode

        treeviewName.Nodes.Clear()
        dt = db.SelectDBRows("SELECT * FROM CHINHANH WHERE ma_so='" & DONVI & "' AND status=1")
        If dt.Rows.Count > 0 Then
            rootNode = treeviewName.Nodes.Add(dt.Rows(0).Item("id"), dt.Rows(0).Item("ten_goi"))
            rootNode.Tag = "ROOT_" & dt.Rows(0).Item("id") & "_" & dt.Rows(0).Item("ma_so")
            getChiLuong_Nam(rootNode, dt.Rows(0).Item("id"), vYear)
            dtCN = db.SelectDBRows("SELECT * FROM CHINHANH WHERE id_goc=" & dt.Rows(0).Item("id") & " AND status=1")
            If dtCN.Rows.Count > 0 Then
                For i = 0 To dtCN.Rows.Count - 1
                    childNode = rootNode.Nodes.Add(dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ten_goi"))
                    childNode.Tag = "DV_" & dtCN.Rows(i).Item("id") & "_" & dtCN.Rows(i).Item("ma_so")
                    'getChiLuong_Nam(rootNode, dt.Rows(0).Item("id"), vYear)
                    getChiLuong_Nam(childNode, dtCN.Rows(i).Item("id"), vYear)
                Next
            End If
            rootNode.Expand()
        End If
    End Sub


    Public Function getNode(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, ByVal CodeParentNode As String, Optional ByVal All As Boolean = True, Optional ByVal CBTS As Boolean = False, Optional ByVal Idphong_viewCB As Integer = 0, Optional ByVal IdDonVi_viewCB As Integer = 0) As TreeNode
        Dim i As Integer
        Dim dtCN, dtPB As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode
        Dim currTRUCTHUOC As String
        Dim strSqlCB As String = ""
        Dim dtCB As DataTable

        currTRUCTHUOC = getTrucThuoc(CodeParentNode)
        dtPB = db.SelectDBRows("SELECT * FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1 Order by Ma_so")
        If dtPB.Rows.Count > 0 Then
            For i = 0 To dtPB.Rows.Count - 1
                childNode = parentNode.Nodes.Add(dtPB.Rows(i).Item("id"), dtPB.Rows(i).Item("ten_phong"))
                childNode.Tag = "PB_" & IdParentNode & "_" & dtPB.Rows(i).Item("id")
                If parentNode.Checked Then
                    childNode.Checked = True
                End If
                If Idphong_viewCB = dtPB.Rows(i).Item("id") And IdParentNode = IdDonVi_viewCB Then
                    getNodeCanbo(childNode, IdParentNode, dtPB.Rows(i).Item("id"))
                End If
            Next
        End If

        If CBTS Then
            strSqlCB = "SELECT IdCanbo, MaCB, Hoten FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh=" & IdParentNode & " and TuNgay<= getdate() and getdate()<=DenNgay"
            dtCB = db.SelectDBRows(strSqlCB)
            If dtCB.Rows.Count > 0 Then
                For i = 0 To dtCB.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                    childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_TS"
                    childNode.ForeColor = Color.DarkViolet
                    If parentNode.Checked Then
                        childNode.Checked = True
                    End If
                Next
            End If
        End If

        If All And Globals.Group <> "S03" And Globals.Group <> "S04" Then
            'Nếu cấp cha là HSC thì kiểm tra quyền quản lý Chi nhánh
            If IdParentNode = 1 Then
                dtCN = db.SelectDBRows(String.Format("Select * from ChiNhanh Where id_goc = {0} and Status = 1 and id in (" & Globals.QuyenQuanLyCN & ")", IdParentNode))
            Else
                dtCN = db.SelectDBRows(String.Format("Select * from ChiNhanh Where id_goc = {0} and Status = 1", IdParentNode))
            End If

            If dtCN.Rows.Count > 0 Then
                For i = 0 To dtCN.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ten_goi"))
                    childNode.Tag = "DV_" & dtCN.Rows(i).Item("id") & "_" & dtCN.Rows(i).Item("ma_so")
                    If DONVI <> gMaDonViTW Then
                        getNode(childNode, dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ma_so"), All, CBTS, Idphong_viewCB, IdDonVi_viewCB)
                    Else
                        Dim IdGoc As Integer = 0
                        IdGoc = db.getNumber("SELECT Id_goc FROM ChiNhanh WHERE id in (SELECT Id_goc FROM ChiNhanh WHERE Id=" & IdDonVi_viewCB & ")")
                        If IdGoc = 1 Then
                            Dim IdGoc_DVview As Integer = 0
                            IdGoc_DVview = db.getNumber("SELECT Id_goc FROM ChiNhanh WHERE Id=" & IdDonVi_viewCB)
                            If IdGoc_DVview = dtCN.Rows(i).Item("id") Then
                                getNode(childNode, dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ma_so"), All, CBTS, Idphong_viewCB, IdDonVi_viewCB)
                            End If
                            If IdGoc_DVview = IdParentNode And IdDonVi_viewCB = dtCN.Rows(i).Item("id") Then
                                getNode(childNode, dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ma_so"), All, CBTS, Idphong_viewCB, IdDonVi_viewCB)
                            End If
                        Else
                            If IdDonVi_viewCB = dtCN.Rows(i).Item("id") Then
                                getNode(childNode, dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ma_so"), All, CBTS, Idphong_viewCB, IdDonVi_viewCB)
                            End If
                        End If
                    End If
                    If parentNode.Checked Then
                        childNode.Checked = True
                    End If
                Next
            End If
        End If

        Return parentNode
    End Function

    ''' <summary>
    ''' Lấy ID trình độ cao nhất theo chuyên môn
    ''' </summary>
    ''' <param name="vIdCanbo"></param>
    ''' <param name="vMa_ChuyenMon">Mã số chuyên môn</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getChinhDo_Max(ByVal vIdCanbo As String, ByVal vMa_ChuyenMon As String) As Integer
        Dim strSQL As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim _result As Integer = 0

        'strSQL = "SELECT (Select ma_so from DanhMuc Where id = a.IdTrinhDo And id_goc = 38 And Status = 1) As ma_so, a.* "
        strSQL = "SELECT IdTrinhDo "
        strSQL += " FROM HS_DTVBCC a, DanhMuc b"
        strSQL += " WHERE a.IdChuyenNganhDT=b.Id and a.IdCanBo = '" + vIdCanbo + "' "
        strSQL += "    and IdChuyenNganhDT in (SELECT ID FROM DanhMuc WHERE ma_so='" & vMa_ChuyenMon & "') And a.HoanThanh = 1"
        strSQL += " Order by Ma_so asc, a.VBCC"

        Using dt As DataTable = dbconn.SelectDBRows(strSQL)
            If Not (dt Is Nothing) Then
                If dt.Rows.Count > 0 Then
                    _result = dt.Rows(0)("IdTrinhDo").ToString().Trim()
                End If
            End If
        End Using

        'If _result = 0 And (vMa_ChuyenMon = "1122") Then
        '    _result = dbconn.getNumber("SELECT IdTrinhDoCT FROM HS_Canbo WHERE IdCanBo='" & vIdCanbo & "'")
        'End If
        Return _result
    End Function

    ''' <summary>
    ''' Lay danh sach can bo trong 1 node
    ''' </summary>
    ''' <param name="vIdDonvi"></param>
    ''' <param name="vidPhongban"></param>
    ''' <param name="isLUONG"></param>
    ''' <param name="isThoiViec"></param>
    ''' <param name="CBCT">Can bo chinh thuc</param>
    ''' <param name="CBNH">Can bo ngan han</param>
    ''' <param name="CBTS">can bo tap su</param>
    ''' <param name="CBCC">Can bo da chuyen cong tac sang DV khac trong he thong</param>
    ''' <param name="CBVH">Can bo ve huu</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getSQL_DSCB(ByVal vIdDonvi As Integer, ByVal vidPhongban As Integer, Optional ByVal isLUONG As Boolean = False, Optional ByVal isThoiViec As Boolean = False, Optional ByVal CBCT As Boolean = False, Optional ByVal CBNH As Boolean = False, Optional ByVal CBTS As Boolean = False, Optional ByVal CBCC As Boolean = False, Optional ByVal CBVH As Boolean = False, Optional ByVal isNghiHuu_ChuyenCC As Boolean = False) As String
        Dim strSqlCB As String = ""
        Dim db As DBAccess = New DBAccess
        Dim VietTat_DVcurr As String = ""
        Dim currTRUCTHUOC As String = ""
        Dim vID_DonVi_goc As Integer = db.getNumber("SELECT id_goc FROM ChiNhanh WHERE [id]=" & vIdDonvi)
        If vID_DonVi_goc = 0 Or vID_DonVi_goc = 1 Then vID_DonVi_goc = vIdDonvi
        'VietTat_DVcurr = TEN_DV_VT
        VietTat_DVcurr = db.getString("SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc)
        Dim sTen_VT = db.getString(String.Format("Select Ten_VT from ChiNhanh Where Id={0}", vID_DonVi_goc))
        
        If isNghiHuu_ChuyenCC Then
            'Lấy danh sách cán bộ đã nghỉ huu hoặc chuyển công tác tới đơn vị khác
            currTRUCTHUOC = getTrucThuoc(db.getString("SELECT ma_so FROM ChiNhanh WHERE id=" & vID_DonVi_goc))
            If DONVI = gMaDonViTW Then
                If vID_DonVi_goc = IdDONVI Then
                    strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew, t3.DBtmp, t3.Rpt_Ngoainganh " & _
                            " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='VBSP' AND IsKiemNhiem = 0  Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                            " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                            "    AND ((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi in (SELECT [id] FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='VBSP' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL ))" & _
                            "          OR (t1.idDonvi_CU=" & vIdDonvi & " AND t1.idPhong_CU in (SELECT [id] FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1 ) AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & ") and Active=1)) " & _
                            "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='VBSP'" & _
                            " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                    'Neu chuyen tu chi nhanh ve thi lay ten viet tat cho nay khong phai VBSP thi sao?

                Else '
                    strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew, t3.DBtmp, t3.Rpt_Ngoainganh " & _
                            " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & ") OR left(IdCanBo,4)='VBSP') AND IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                            " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                            "    AND ((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi in (SELECT [id] FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1 ) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND (left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & ") OR left(IdCanBo,4)='VBSP') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL ))" & _
                            "          OR (t1.idDonvi_Cu=" & vIdDonvi & " AND t1.idPhong_Cu in (SELECT [id] FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1) AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1)) " & _
                            "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') " & _
                            " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                End If
            Else
                'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz
                'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "') AND IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                '           "    AND ((((t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ")) OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1)) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )) " & _
                '           "         OR (t1.idDonvi_Cu in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1)) " & _
                '           "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "','" + sTen_VT + "')" & _
                '           " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "

                'CHUDV: Bỏ từ 12/2022 do có điều kiện bắt theo 4 ký tự đầu của IdCanBo
                'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                '          " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "') AND IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                '          " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                '          "    AND ((((t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ")) ) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )) " & _
                '          "         OR (t1.idDonvi_Cu in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1)) " & _
                '          "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "','" + sTen_VT + "')" & _
                '          " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz
                'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' AND IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                '           "    AND ((((t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ")) OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1)) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )) " & _
                '           "         OR (t1.idDonvi_Cu in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1)) " & _
                '           "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "'" & _
                '           " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "

                '"    AND ((((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi in (SELECT [id] FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1 )) OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idDonvi_Cu=" & vIdDonvi & " and t1.idPhong_Cu=" & vidPhongban & " and Active=1)) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL ))" & _
                '"    AND ((((t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ")) OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idDonvi_Cu=" & vIdDonvi & " and t1.idPhong_Cu=" & vidPhongban & " and Active=1)) AND EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL ))" & _
                '"         OR ((t1.idDonvi_Cu=" & vIdDonvi & " AND t1.idPhong_Cu in (SELECT [id] FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1 ) AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & vIdDonvi & " or id_goc=" & vIdDonvi & ") and Active=1))) " & _

                'CHUDV: Chỉnh sửa lại bỏ điều kiện giới hạn 4 ký tự đầu Viết tắt của chi nhánh
                strSqlCB = "Select T1.IdCanBo, MaCB, HoTen, Login_Username, Id_Nhom, CapQuanLy, T1.IdDonvi_Moi, T1.IdPhong_Moi, T3.IdNew" & _
                          "	   From QDNhanSu T1," & _
                          "	   (" & _
                          "			Select IdCanBo, Max(NgayHL) As NgayHL From QDNhanSu Where IsQD_NHCS = 1 AND IsKiemNhiem = 0 And Cast(NgayHL As Date) <= Cast(GetDate() As Date)" & _
                          "					Group By IdCanBo Having Max(NgayHL)<= Cast(GetDate() As Date)" & _
                          "	   ) T2, HS_CanBo T3" & _
                          "	   Where (T1.IdCanBo = T2.IdCanBo And T1.NgayHL = T2.NgayHL And T2.IdCanBo = T3.IdCanBo)    " & _
                          "		AND (" & _
                          "				(" & _
                          "					(T1.IdDonvi_Moi In (Select X.Id From ChiNhanh X Where X.Id=" & vIdDonvi & " Or X.Id_Goc=" & vIdDonvi & "))" & _
                          "				 And EXISTS " & _
                          "					(" & _
                          "						SELECT T4.IdCanBo From " & _
                          "								(" & _
                          "									SELECT IdCanBo, Max(Ngay_HL) As Ngay_HL From HS_CBThoiviec WHERE IsQD_NHCS = 1 And Cast(Ngay_HL As Date) <= Cast(GetDate() As Date)" & _
                          "											Group By IdCanBo Having Max(Ngay_HL)<= Cast(GetDate() As Date)" & _
                          "								) T4 Where T4.IdCanBo= T2.IdCanBo And DateAdd(Second, 10, T4.Ngay_HL) > Cast(T2.NgayHL As Date)" & _
                          "					)" & _
                          "				)" & _
                          "			 OR (" & _
                          "					T1.IdDonVi_Cu In (Select X.Id From ChiNhanh X Where X.Id=" & vIdDonvi & " Or X.Id_Goc=" & vIdDonvi & ") And T1.IdDonvi_Moi Not In (Select X.Id From ChiNhanh X Where X.Id=" & vIdDonvi & " Or X.Id_Goc=" & vIdDonvi & ") And Active=1" & _
                          "				)" & _
                          "			)     AND T1.IsQD_NHCS = 1" & _
                          "		Order by IdChucVu_Moi, IdChuyenMon_Moi, MaCB, HoTen"
            End If
        Else
            If isLUONG Then
                If (CBCT = False And CBNH = False And CBTS = False And CBCC = False And CBVH = False) Then CBCT = True
                If CBCT Then
                    'If DONVI = gMaDonViTW Then
                    '    If vID_DonVi_goc <> IdDONVI Then
                    '        VietTat_DVcurr = db.getString("SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc)
                    '    End If   in ('" & VietTat_DVcurr & "','" & gMaDonViTW & "')
                    'End If
                    'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                    '                               " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & VietTat_DVcurr & "' and (IsKiemNhiem = 0 or (idDonvi_Moi<>" & vIdDonvi & " AND idDonvi_Cu=" & vIdDonvi & " AND idPhong_Cu=" & vidPhongban & " AND datediff(month,ngayHL, getdate())=0)) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                    '                               " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                    '                               "    AND ((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idDonvi_Cu=" & vIdDonvi & " and t1.idPhong_Cu=" & vidPhongban & " and Active=1))" & _
                    '                               "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & VietTat_DVcurr & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                    '                               "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & VietTat_DVcurr & "'" & _
                    '                               " Order by IdChucvu_moi, idChuyenMon_Moi, Hoten "
                    'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                    '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('" & VietTat_DVcurr & "','VBSP') and (IsKiemNhiem = 0 or (idDonvi_Moi<>" & vIdDonvi & " AND idDonvi_Cu=" & vIdDonvi & " AND idPhong_Cu=" & vidPhongban & " AND datediff(month,ngayHL, getdate())=0)) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                    '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                    '           "    AND ((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idDonvi_Cu=" & vIdDonvi & " and t1.idPhong_Cu=" & vidPhongban & " and Active=1))" & _
                    '           "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('" & VietTat_DVcurr & "','VBSP') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                    '           "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in ('" & VietTat_DVcurr & "','VBSP')" & _
                    '           " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                    strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                                    " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ") and Idcanbo not in (SELECT Idcanbo FROM QDNhanSu WHERE IsKiemNhiem = 2 and idDonvi_Cu=" & vIdDonvi & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                                    " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                                    "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                                    "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (Select Distinct Left(IdCanBo, 4) TenVT From Hs_CanBo Where IdDonVi=" & vID_DonVi_goc & ") GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                                    "    AND IsQD_NHCS =1 " & _
                                    " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                End If
            Else
                If isThoiViec Then
                    strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                               " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " and idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                               " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                               "    AND ((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idPhong_Cu=" & vidPhongban & " and Active=1))" & _
                               "    AND IsQD_NHCS =1 " & _
                               " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "

                Else
                    If DONVI = gMaDonViTW Then
                        If vID_DonVi_goc = IdDONVI Then
                            strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                                    " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                                    " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                                    "    AND (t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & ")" & _
                                    "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                                    "    AND IsQD_NHCS =1 " & _
                                    " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                            'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                            '        " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & ") AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                            '        " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                            '        "    AND (t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & ")" & _
                            '        "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & ") GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                            '        "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & ") " & _
                            '        " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                        Else
                            strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                                   " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                                   " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                                   "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                                   "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                                   "    AND IsQD_NHCS =1 " & _
                                   " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "

                            'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                            '       " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') AND (IsKiemNhiem = 0 or IsKiemNhiem = 1) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                            '       " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                            '       "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                            '       "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                            '       "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') " & _
                            '       " Order by IdChucvu_moi, idChuyenMon_Moi, Hoten "

                            'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                            '       " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                            '       " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                            '       "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                            '       "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                            '       "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4) in (SELECT ten_vt FROM ChiNhanh WHERE id=" & vID_DonVi_goc & "  or ma_so='000100') " & _
                            '       " Order by IdChucvu_moi, idChuyenMon_Moi, Hoten "
                        End If
                    Else
                        'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                        '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                        '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                        '           "    AND ((t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idDonvi_Cu=" & vIdDonvi & " and t1.idPhong_Cu=" & vidPhongban & " and Active=1))" & _
                        '           "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                        '           "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "'" & _
                        '           " Order by IdChucvu_moi, idChuyenMon_Moi, Hoten "

                        'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                        '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                        '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                        '           "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                        '           "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                        '           "    AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "'" & _
                        '           " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "

                        'zzzzzzzzzzzzzzzzzzzzzzzzzzzzz
                        'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                        '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ") and Idcanbo not in (SELECT Idcanbo FROM QDNhanSu WHERE IsKiemNhiem = 2 and idDonvi_Cu=" & vIdDonvi & ")) Group by Idcanbo Having CONVERT(varchar(8),max(NgayHL),112)<=CONVERT(varchar(8),getdate(),112)) t2, HS_Canbo t3 " & _
                        '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                        '           "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                        '           "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                        '           "    AND IsQD_NHCS =1 " & _
                        '           " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                        'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                        '           " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ") and Idcanbo not in (SELECT Idcanbo FROM QDNhanSu WHERE IsKiemNhiem = 2 and idDonvi_Cu=" & vIdDonvi & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                        '           " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                        '           "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                        '           "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                        '           "    AND IsQD_NHCS =1 " & _
                        '           " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "


                        'strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                        '                " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ") and Idcanbo not in (SELECT Idcanbo FROM QDNhanSu WHERE IsKiemNhiem = 2 and idDonvi_Cu=" & vIdDonvi & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                        '                " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                        '                "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                        '                "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in ('VBSP','" & TEN_DV_VT & "') GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                        '                "    AND IsQD_NHCS =1 " & _
                        '                " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "

                        strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, login_Username, ID_Nhom, CapQuanLy, t1.idDonvi_Moi, t1.idPhong_Moi, t3.IdNew " & _
                                        " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 AND (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idDonvi_Moi=" & vIdDonvi & " AND idPhong_Moi=" & vidPhongban & ") and Idcanbo not in (SELECT Idcanbo FROM QDNhanSu WHERE IsKiemNhiem = 2 and idDonvi_Cu=" & vIdDonvi & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                                        " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
                                        "    AND t1.idDonvi_Moi=" & vIdDonvi & " AND t1.idPhong_Moi=" & vidPhongban & _
                                        "    AND NOT EXISTS (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND left(IdCanBo,4) in (Select Distinct Left(IdCanBo, 4) TenVT From Hs_CanBo Where IdDonVi=" & vID_DonVi_goc & ") GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t2.idCanbo AND t4.Ngay_HL >t2.NgayHL )" & _
                                        "    AND IsQD_NHCS =1 " & _
                                        " Order by IdChucvu_moi, idChuyenMon_Moi, MaCB, Hoten "
                        'zzzzzzzzzzzzzzzzzzzzzzzzzzzzz
                    End If
                End If
            End If
        End If

        Return strSqlCB
    End Function

    Public Function getNodeCanbo(ByVal parentNode As TreeNode, ByVal IdDonvi As Integer, ByVal idPhongban As Integer, Optional ByVal isLUONG As Boolean = False, Optional ByVal isThoiViec As Boolean = False, Optional ByVal CBCT As Boolean = False, Optional ByVal CBNH As Boolean = False, Optional ByVal CBTS As Boolean = False, Optional ByVal CBCC As Boolean = False, Optional ByVal CBVH As Boolean = False) As TreeNode
        Dim i, j As Integer
        Dim childNode As TreeNode
        Dim dtCB, dtCBTS As DataTable
        Dim db As DBAccess = New DBAccess
        Dim strSqlCB As String = ""
        Dim strSqlCBTS As String = ""
        'If isLUONG Then
        '    If (CBCT = False And CBNH = False And CBTS = False And CBCC = False And CBVH = False) Then CBCT = True
        '    If CBCT Then
        '        'strSqlCB = " SELECT t2.idCanbo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '        '           " FROM QDNhansu t1, HS_Canbo t2 " & _
        '        '           " WHERE t1.idCanbo = t2.idCanbo" & _
        '        '           "    AND t2.idCanbo not in (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 ) " & _
        '        '           "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR " & _
        '        '           "         (t1.idDonvi_Moi<>" & IdDonvi & " AND t1.idDonvi_Cu=" & IdDonvi & " AND t1.idPhong_Cu=" & idPhongban & " AND datediff(month,t1.ngayHL, getdate())=0)) " & _
        '        '           "    AND t1.idcanbo not in (SELECT idcanbo FROM QDNhansu t3 WHERE t3.idCanbo=t1.idCanbo and datediff(second,t1.ngayHL,t3.ngayHL)>0)" & _
        '        '           "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
        '        '           " Order by idChucvu_moi, Hoten "

        '        strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                   " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and (IsKiemNhiem = 0 or (idDonvi_Moi<>" & IdDonvi & " AND idDonvi_Cu=" & IdDonvi & " AND idPhong_Cu=" & idPhongban & " AND datediff(month,ngayHL, getdate())=0)) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                   " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                   "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & " or id_goc=" & IdDonvi & ") and t1.idDonvi_Cu=" & IdDonvi & " and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                   "    AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 )" & _
        '                   "    AND IsQD_NHCS =1" & _
        '                   " Order by IdChucvu_moi, Hoten "
        '    End If
        'Else
        '    If isThoiViec Then
        '        strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                   " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idPhong_Moi=" & idPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                   " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                   "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & " or id_goc=" & IdDonvi & ") and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                   "    AND IsQD_NHCS =1 " & _
        '                   " Order by IdChucvu_moi, Hoten "
        '        'AND isKiemNhiem = 0 
        '    Else
        '        If DONVI = gMaDonViTW Then
        '            strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                       " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idPhong_Moi=" & idPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                       " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                       "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & ") and t1.idDonvi_Cu=" & IdDonvi & " and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                       "    AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 )" & _
        '                       "    AND IsQD_NHCS =1 " & _
        '                       " Order by IdChucvu_moi, Hoten "
        '            'AND isKiemNhiem = 0  

        '        Else
        '            strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                       " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and (IsKiemNhiem = 0 or (IsKiemNhiem = 1 and idPhong_Moi=" & idPhongban & ")) Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                       " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                       "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & " or id_goc=" & IdDonvi & ") and t1.idDonvi_Cu=" & IdDonvi & " and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                       "    AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 )" & _
        '                       "    AND IsQD_NHCS =1" & _
        '                       " Order by IdChucvu_moi, Hoten "
        '            ' AND isKiemNhiem = 0
        '        End If

        '    End If
        'End If

        strSqlCB = getSQL_DSCB(IdDonvi, idPhongban, isLUONG, isThoiViec, CBCT, CBNH, CBTS, CBCC, CBVH)
        If strSqlCB <> "" Then
            dtCB = db.SelectDBRows(strSqlCB)
            If dtCB.Rows.Count > 0 Then
                For i = 0 To dtCB.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                    childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_" & dtCB.Rows(i).Item("MaCB") & "_" & dtCB.Rows(i).Item("idDonvi_Moi") & "_" & dtCB.Rows(i).Item("idPhong_Moi")
                    If parentNode.Checked Then
                        childNode.Checked = True
                    End If
                Next
            End If
        End If
        If isLUONG Then
            'If CBTS Then
            '    strSqlCBTS = "SELECT IdCanbo, MaCB, Hoten, (CASE WHEN (IdBacLuong is null or IdBacLuong=0) THEN 'NH' ELSE 'TS' END) as loai FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh=" & IdDonvi & " AND left(rtrim(IdPhongBan),2)='PB' AND right(IdPhongBan,len(rtrim(IdPhongBan))-charindex('_',rtrim(IdPhongBan)))=" & idPhongban & " AND TuNgay<= getdate() and getdate()<=DenNgay "
            '    dtCBTS = db.SelectDBRows(strSqlCBTS)
            '    If dtCBTS.Rows.Count > 0 Then
            '        For j = 0 To dtCBTS.Rows.Count - 1
            '            childNode = parentNode.Nodes.Add(dtCBTS.Rows(j).Item("idCanbo"), dtCBTS.Rows(j).Item("hoten"))
            '            childNode.Tag = "CB_" & dtCBTS.Rows(j).Item("idCanbo") & "_" & dtCBTS.Rows(j).Item("loai")
            '            If dtCBTS.Rows(j).Item("loai") = "NH" Then
            '                childNode.ForeColor = Color.Blue
            '            Else
            '                childNode.ForeColor = Color.Green
            '            End If
            '            If parentNode.Checked Then
            '                childNode.Checked = True
            '            End If
            '        Next
            '    End If
            'End If
            If CBTS Then
                strSqlCBTS = "SELECT IdCanbo, MaCB, Hoten FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh=" & IdDonvi & " AND left(rtrim(IdPhongBan),2)='PB' AND right(IdPhongBan,len(rtrim(IdPhongBan))-charindex('_',rtrim(IdPhongBan)))=" & idPhongban & " AND TuNgay<= getdate() and getdate()<=DenNgay  AND not (IdBacLuong is null or IdBacLuong=0)"
                dtCBTS = db.SelectDBRows(strSqlCBTS)
                If dtCBTS.Rows.Count > 0 Then
                    For j = 0 To dtCBTS.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCBTS.Rows(j).Item("idCanbo"), dtCBTS.Rows(j).Item("hoten"))
                        childNode.Tag = "CB_" & dtCBTS.Rows(j).Item("idCanbo") & "_TS"
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If
            If CBNH Then
                strSqlCBTS = "SELECT IdCanbo, MaCB, Hoten FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh=" & IdDonvi & " AND left(rtrim(IdPhongBan),2)='PB' AND right(IdPhongBan,len(rtrim(IdPhongBan))-charindex('_',rtrim(IdPhongBan)))=" & idPhongban & " AND TuNgay<= getdate() and getdate()<=DenNgay AND (IdBacLuong is null or IdBacLuong=0)"
                dtCBTS = db.SelectDBRows(strSqlCBTS)
                If dtCBTS.Rows.Count > 0 Then
                    For j = 0 To dtCBTS.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCBTS.Rows(j).Item("idCanbo"), dtCBTS.Rows(j).Item("hoten"))
                        childNode.Tag = "CB_" & dtCBTS.Rows(j).Item("idCanbo") & "_NH"
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If
            ' Lấy danh sách cán  bộ đã Chuyển Công tác sang đơn vị khác trong cùng hệ thống TRONG NĂM và Trước liền kề
            If CBCC Then
                strSqlCBTS = ""
                strSqlCBTS = " SELECT t1.IdCanBo, MaCB, Hoten FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                             " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                             "       AND (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & " or id_goc=" & IdDonvi & ") and t1.idPhong_Cu=" & idPhongban & " and Active=1)" & _
                             "       AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 ) AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
                             "       AND ( year(t1.NgayHL) = year(getdate()) or year(t1.NgayHL) = year(dateadd(year,-1,getdate())) )" & _
                             " Order by IdChucvu_moi, Hoten "
                dtCBTS = db.SelectDBRows(strSqlCBTS)
                If dtCBTS.Rows.Count > 0 Then
                    For j = 0 To dtCBTS.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCBTS.Rows(j).Item("idCanbo"), dtCBTS.Rows(j).Item("hoten"))
                        childNode.Tag = "CB_" & dtCBTS.Rows(j).Item("idCanbo") & "_CC"
                        'childNode.ForeColor = Color.Purple
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If
            ' Lấy danh sách cán bộ đã về hưu TRONG NĂM và Trước liền kề
            If CBVH Then
                strSqlCBTS = ""
                strSqlCBTS = " SELECT t1.IdCanBo, MaCB, Hoten FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                             " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                             "    AND (t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ")" & _
                             "    AND t1.IdCanBo in (SELECT idCanbo FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND (year(Ngay_HL) = year(getdate()) OR year(Ngay_HL) = year(dateadd(year,-1,getdate()))) AND IdLoaiQD in (SELECT id FROM Danhmuc WHERE ma_so='3703')) " & _
                             "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
                             " Order by IdChucvu_moi, Hoten "
                dtCBTS = db.SelectDBRows(strSqlCBTS)
                If dtCBTS.Rows.Count > 0 Then
                    For j = 0 To dtCBTS.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCBTS.Rows(j).Item("idCanbo"), dtCBTS.Rows(j).Item("hoten"))
                        childNode.Tag = "CB_" & dtCBTS.Rows(j).Item("idCanbo") & "_VH"
                        'childNode.ForeColor = Color.Red
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If
        End If
        Return parentNode
    End Function

    ''' <summary>
    ''' hàm tìm kiếm node trong treeview co tag=vcode
    ''' </summary>
    ''' <param name="vNodes">danh sách các node trong treeview</param>
    ''' <param name="vTag">Tag của Node</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function findNode(ByVal vNodes As TreeNodeCollection, ByVal vTag As String) As TreeNode
        Dim _result As String = ""
        If (vNodes Is Nothing Or vNodes.Count = 0 Or vTag = String.Empty) Then
            Return Nothing
        End If

        For Each node As TreeNode In vNodes
            If Not (node Is Nothing) Then
                _result = node.Tag.ToString()

                'Debug
                If node.Tag = "PB_286_21" Or node.Tag = "DV_286_000304" Then '_result = "PB_282_21" Then
                    Dim i As Integer = 0
                End If

                If (_result = vTag) Then
                    Return node
                End If
            End If

            Dim foundNode As TreeNode = findNode(node.Nodes, vTag)
            'nếu tìm thấy trả về node
            If Not (foundNode Is Nothing) Then
                Return foundNode
            End If
        Next
        Return Nothing
    End Function

    Public Function refreshNodeNOTCanbo(ByRef vTreeView As TreeView, ByVal parentNode As TreeNode, ByVal IdDonvi As Integer, ByVal idPhongban As Integer, Optional ByVal isLUONG As Boolean = False, Optional ByVal isThoiViec As Boolean = False, Optional ByVal vIdCanBo As String = "") As TreeNode
        Dim i As Integer
        Dim childNode As TreeNode
        Dim dtCB As DataTable
        Dim db As DBAccess = New DBAccess
        Dim strSqlCB As String

        'If isLUONG Then
        '    strSqlCB = " SELECT t2.idCanbo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '               " FROM QDNhansu t1, HS_Canbo t2 " & _
        '               " WHERE t1.idCanbo = t2.idCanbo" & _
        '               "    AND t2.idCanbo not in (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 ) " & _
        '               "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR " & _
        '               "         (t1.idDonvi_Moi<>" & IdDonvi & " AND t1.idDonvi_Cu=" & IdDonvi & " AND t1.idPhong_Cu=" & idPhongban & " AND datediff(month,t1.ngayHL, getdate())=0)) " & _
        '               "    AND t1.idcanbo not in (SELECT idcanbo FROM QDNhansu t3 WHERE t3.idCanbo=t1.idCanbo and datediff(second,t1.ngayHL,t3.ngayHL)>0)" & _
        '               "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
        '               " Order by idChucvu_moi, Hoten "
        'Else
        '    If isThoiViec Then
        '        strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                   " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                   " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                   "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & " or id_goc=" & IdDonvi & ") and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                   "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
        '                   " Order by IdChucvu_moi, Hoten "
        '    Else
        '        If DONVI = gMaDonViTW Then
        '            strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                       " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                       " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                       "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & ") and t1.idDonvi_Cu=" & IdDonvi & " and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                       "    AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 )" & _
        '                       "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
        '                       " Order by IdChucvu_moi, Hoten "

        '        Else
        '            strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten, t1.idDonvi_Moi, t1.idPhong_Moi " & _
        '                       " FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
        '                       " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo)" & _
        '                       "    AND ((t1.idDonvi_Moi=" & IdDonvi & " AND t1.idPhong_Moi=" & idPhongban & ") OR (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDonvi & " or id_goc=" & IdDonvi & ") and t1.idDonvi_Cu=" & IdDonvi & " and t1.idPhong_Cu=" & idPhongban & " and Active=1))" & _
        '                       "    AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 )" & _
        '                       "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
        '                       " Order by IdChucvu_moi, Hoten "
        '        End If
        '    End If
        'End If
        'dtCB = db.SelectDBRows(strSqlCB)

        strSqlCB = getSQL_DSCB(IdDonvi, idPhongban, isLUONG, isThoiViec, True)
        If strSqlCB <> "" Then
            dtCB = db.SelectDBRows(strSqlCB)
            If parentNode.Nodes.Count <> dtCB.Rows.Count Then
                While parentNode.Nodes.Count > 0
                    parentNode.Nodes.Remove(parentNode.FirstNode)
                End While
                If dtCB.Rows.Count > 0 Then
                    For i = 0 To dtCB.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                        childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_" & dtCB.Rows(i).Item("MaCB")
                        If parentNode.Checked Then
                            childNode.Checked = True
                        Else
                            If vIdCanBo = dtCB.Rows(i).Item("idCanbo") Then
                                vTreeView.SelectedNode = childNode
                            End If
                        End If
                    Next
                End If
            End If
        End If

        Return parentNode
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="treeviewName"></param>
    ''' <param name="CBCT"> ~ 0 lao dong dai han chinh thuc;</param>
    ''' <param name="CBNH"> ~ 1: ngan han;</param>
    ''' <param name="CBTS"> ~ 2: tap su;</param>
    ''' <param name="CBCC"> ~ 3: da chuyen den DV khac trong NHCS;</param>
    ''' <param name="CBVH"> ~ 4: Da nghi huu</param>
    ''' <remarks></remarks>
    Public Sub bindTreeviewFull(ByVal treeviewName As TreeView, Optional ByVal CBCT As Boolean = True, Optional ByVal CBNH As Boolean = True, Optional ByVal CBTS As Boolean = True, Optional ByVal CBCC As Boolean = True, Optional ByVal CBVH As Boolean = True, Optional ByVal view_DV_PB_only As Boolean = False, Optional ByVal isTS_NH_FrmChonDS_CB As Boolean = False)
        Dim dt As DataTable
        Dim db As DBAccess = New DBAccess
        Dim i As Integer
        Dim idx As Integer = 1
        Dim _All As Boolean = True
        Dim rootNode As TreeNode
        dt = db.SelectDBRows("SELECT * FROM CHINHANH WHERE ma_so='" & DONVI & "' AND status=1")
        If dt.Rows.Count > 0 Then
            'For i = 0 To dt.Rows.Count - 1
            rootNode = treeviewName.Nodes.Add(dt.Rows(i).Item("id"), dt.Rows(i).Item("ten_goi"))
            rootNode.Tag = "ROOT_" & dt.Rows(i).Item("id") & "_" & dt.Rows(i).Item("ma_so")
            If (dt.Rows(i).Item("ma_so") = gMaDonViTW) Then _All = False
            getNodeFull(rootNode, dt.Rows(i).Item("id"), dt.Rows(i).Item("ma_so"), _All, CBCT, CBNH, CBTS, CBCC, CBVH, view_DV_PB_only, isTS_NH_FrmChonDS_CB)
            'Next
        End If
    End Sub

    ''' <summary>
    ''' Lấy toàn bộ danh sách cán bộ
    ''' </summary>
    ''' <param name="parentNode"></param>
    ''' <param name="IdParentNode"></param>
    ''' <param name="CodeParentNode"></param>
    ''' <param name="All"></param>
    ''' <param name="CBTS"></param>
    ''' <param name="CBCC"></param>
    ''' <param name="CBVH"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getNodeFull(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, ByVal CodeParentNode As String, Optional ByVal All As Boolean = True, Optional ByVal CBCT As Boolean = True, Optional ByVal CBNH As Boolean = True, Optional ByVal CBTS As Boolean = True, Optional ByVal CBCC As Boolean = True, Optional ByVal CBVH As Boolean = True, Optional ByVal view_DV_PB_only As Boolean = False, Optional ByVal isTS_NH_FrmChonDS_CB As Boolean = False) As TreeNode
        Dim i As Integer
        Dim dtCN, dtPB As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode
        Dim currTRUCTHUOC As String
        Dim strSqlCB As String = ""
        Dim dtCB As DataTable


        currTRUCTHUOC = getTrucThuoc(CodeParentNode)
        If Not (isTS_NH_FrmChonDS_CB And currTRUCTHUOC = 4) Then
            dtPB = db.SelectDBRows("SELECT * FROM PHONGBAN WHERE charindex('" & currTRUCTHUOC.ToString & "',truc_thuoc)>0 AND status=1 Order by Ma_so")
            If dtPB.Rows.Count > 0 Then
                For i = 0 To dtPB.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(dtPB.Rows(i).Item("id"), dtPB.Rows(i).Item("ten_phong"))
                    childNode.Tag = "PB_" & IdParentNode & "_" & dtPB.Rows(i).Item("id")
                    If parentNode.Checked Then
                        childNode.Checked = True
                    End If
                    If view_DV_PB_only = False Then
                        getNodeCanbo(childNode, parentNode.Name, dtPB.Rows(i).Item("id"), True, False, CBCT, CBNH, CBTS, CBCC, CBVH)
                    End If
                Next
            End If
        End If

        If view_DV_PB_only = False Then
            'If CBTS Then
            '    strSqlCB = "SELECT IdCanbo, MaCB, Hoten, (CASE WHEN (IdBacLuong is null or IdBacLuong=0) THEN 'NH' ELSE 'TS' END) as loai FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh in (SELECT Id_goc FROM ChiNhanh WHERE Id=" & IdParentNode & ") AND left(rtrim(IdPhongBan),2)='DV' AND right(IdPhongBan,len(rtrim(IdPhongBan))-charindex('_',rtrim(IdPhongBan)))=" & IdParentNode & " AND TuNgay<= getdate() and getdate()<=DenNgay"
            '    dtCB = db.SelectDBRows(strSqlCB)
            '    If dtCB.Rows.Count > 0 Then
            '        For i = 0 To dtCB.Rows.Count - 1
            '            childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
            '            childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_" & dtCB.Rows(i).Item("loai")
            '            If dtCB.Rows(i).Item("loai") = "NH" Then
            '                childNode.ForeColor = Color.Blue
            '            Else
            '                childNode.ForeColor = Color.Green
            '            End If
            '            If parentNode.Checked Then
            '                childNode.Checked = True
            '            End If
            '        Next
            '    End If
            'End If
            If CBNH Then
                strSqlCB = "SELECT IdCanbo, MaCB, Hoten FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh in (SELECT Id_goc FROM ChiNhanh WHERE Id=" & IdParentNode & ") AND left(rtrim(IdPhongBan),2)='DV' AND right(IdPhongBan,len(rtrim(IdPhongBan))-charindex('_',rtrim(IdPhongBan)))=" & IdParentNode & " AND TuNgay<= getdate() and getdate()<=DenNgay  AND (IdBacLuong is null or IdBacLuong=0)"
                dtCB = db.SelectDBRows(strSqlCB)
                If dtCB.Rows.Count > 0 Then
                    For i = 0 To dtCB.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                        childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_NH"
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If

            If CBTS Then
                strSqlCB = "SELECT IdCanbo, MaCB, Hoten, (CASE WHEN (IdBacLuong is null or IdBacLuong=0) THEN 'NH' ELSE 'TS' END) as loai FROM V$_HSCB_TS_HDLD WHERE IdChiNhanh in (SELECT Id_goc FROM ChiNhanh WHERE Id=" & IdParentNode & ") AND left(rtrim(IdPhongBan),2)='DV' AND right(IdPhongBan,len(rtrim(IdPhongBan))-charindex('_',rtrim(IdPhongBan)))=" & IdParentNode & " AND TuNgay<= getdate() and getdate()<=DenNgay  AND not (IdBacLuong is null or IdBacLuong=0)"
                dtCB = db.SelectDBRows(strSqlCB)
                If dtCB.Rows.Count > 0 Then
                    For i = 0 To dtCB.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                        childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_TS"
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If

            ' Lấy danh sách cán  bộ đã Chuyển Công tác sang đơn vị khác trong cùng hệ thống TRONG NĂM và Trước liền kề
            If CBCC Then
                strSqlCB = ""
                strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 and IsKiemNhiem = 0 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                             " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                             "       AND (t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and t1.idPhong_Cu=" & IdParentNode & " and Active=1)" & _
                             "       AND NOT EXISTS (SELECT idCanbo FROM HS_CBThoiviec WHERE idCanbo= t2.idCanbo AND IsQD_NHCS =1 ) AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
                             "       AND ( year(t1.NgayHL) = year(getdate()) or year(t1.NgayHL) = year(dateadd(year,-1,getdate())) )" & _
                             " Order by IdChucvu_moi, Hoten "
                dtCB = db.SelectDBRows(strSqlCB)
                If dtCB.Rows.Count > 0 Then
                    For i = 0 To dtCB.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                        childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_CC"
                        'childNode.ForeColor = Color.Purple
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If
            ' Lấy danh sách cán bộ đã về hưu TRONG NĂM và Trước liền kề
            If CBVH Then
                strSqlCB = ""
                strSqlCB = " SELECT t1.IdCanBo, MaCB, Hoten FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                             " WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                             "    AND (t1.idDonvi_Moi=" & IdDONVI & " AND t1.idPhong_Moi=" & IdParentNode & ")" & _
                             "    AND t1.IdCanBo in (SELECT idCanbo FROM HS_CBThoiviec WHERE IsQD_NHCS =1 AND (year(Ngay_HL) = year(getdate()) OR year(Ngay_HL) = year(dateadd(year,-1,getdate()))) AND IdLoaiQD in (SELECT id FROM Danhmuc WHERE ma_so='3703')) " & _
                             "    AND isKiemNhiem = 0  AND IsQD_NHCS =1 " & _
                             " Order by IdChucvu_moi, Hoten "
                dtCB = db.SelectDBRows(strSqlCB)
                If dtCB.Rows.Count > 0 Then
                    For i = 0 To dtCB.Rows.Count - 1
                        childNode = parentNode.Nodes.Add(dtCB.Rows(i).Item("idCanbo"), dtCB.Rows(i).Item("hoten"))
                        childNode.Tag = "CB_" & dtCB.Rows(i).Item("idCanbo") & "_VH"
                        'childNode.ForeColor = Color.Red
                        If parentNode.Checked Then
                            childNode.Checked = True
                        End If
                    Next
                End If
            End If
        End If
        If All Then
            dtCN = db.SelectDBRows("SELECT * FROM CHINHANH WHERE id_goc=" & IdParentNode & " AND status=1")
            If dtCN.Rows.Count > 0 Then
                For i = 0 To dtCN.Rows.Count - 1
                    childNode = parentNode.Nodes.Add(dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ten_goi"))
                    childNode.Tag = "DV_" & dtCN.Rows(i).Item("id") & "_" & dtCN.Rows(i).Item("ma_so")
                    getNodeFull(childNode, dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ma_so"), All, CBCT, CBNH, CBTS, CBCC, CBVH, view_DV_PB_only, isTS_NH_FrmChonDS_CB)
                    If parentNode.Checked Then
                        childNode.Checked = True
                    End If
                Next
            End If
        End If

        Return parentNode
    End Function

    Public Sub bindDSKhenThuong(ByVal treeviewName As TreeView, ByVal vKT_ChuyenMon As Boolean)
        Dim dt, dtCN As DataTable
        Dim db As DBAccess = New DBAccess
        Dim i As Integer
        Dim rootNode, childNode As TreeNode

        dt = db.SelectDBRows("SELECT * FROM CHINHANH WHERE ma_so='" & DONVI & "' AND status=1")
        If dt.Rows.Count > 0 Then
            If (dt.Rows(i).Item("ma_so") = gMaDonViTW) Then
                rootNode = treeviewName.Nodes.Add(dt.Rows(0).Item("id"), "Ngân hàng Chính sách Xã hội")
            Else
                rootNode = treeviewName.Nodes.Add(dt.Rows(0).Item("id"), dt.Rows(0).Item("ten_goi"))
            End If
            rootNode.Tag = "ROOT_" & dt.Rows(0).Item("id") & "_" & dt.Rows(0).Item("ma_so")
            getKhenThuong_Nam(rootNode, dt.Rows(0).Item("id"), vKT_ChuyenMon)
            dtCN = db.SelectDBRows("SELECT * FROM CHINHANH WHERE id_goc in (SELECT id FROM CHINHANH WHERE id=" & dt.Rows(0).Item("id") & " and ma_so='000100') AND status=1")
            If dtCN.Rows.Count > 0 Then
                For i = 0 To dtCN.Rows.Count - 1
                    childNode = rootNode.Nodes.Add(dtCN.Rows(i).Item("id"), dtCN.Rows(i).Item("ten_goi"))
                    childNode.Tag = "DV_" & dtCN.Rows(i).Item("id") & "_" & dtCN.Rows(i).Item("ma_so")
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm lấy danh sách các năm có khen thưởng trong năm
    ''' </summary>
    ''' <param name="parentNode"></param>
    ''' <param name="IdParentNode"></param>
    ''' <param name="vYear"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getKhenThuong_Nam(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, ByVal vKT_ChuyenMon As Int16, Optional ByVal vYear As Integer = 0) As TreeNode
        Dim i As Integer
        Dim childNode As TreeNode
        'Dim dtNam As DataTable
        'Dim db As DBAccess = New DBAccess
        '''''''''''''''''''''''''''''''''''''''''
        'Phần này chạy vẫn OK, but có thể ko cần
        '''''''''''''''''''''''''''''''''''''''''
        'dtNam = db.SelectDBRows(" SELECT a.NamKT as Nam FROM HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong " & _
        '                        " WHERE(b.IdDonVi = " & IdParentNode & ") " & _
        '                        "    or b.IdDonVi in (SELECT id FROM ChiNhanh WHERE Id_goc=" & IdParentNode & ")	" & _
        '                        "    or b.IdDonVi in (SELECT id FROM ChiNhanh WHERE Id_goc in (SELECT id FROM ChiNhanh WHERE Id_goc=" & IdParentNode & ")) " & _
        '                        " Group by NamKT")
        'If dtNam.Rows.Count > 0 Then
        '    For i = 0 To dtNam.Rows.Count - 1
        '        childNode = parentNode.Nodes.Add(IdParentNode, "Năm " & dtNam.Rows(i).Item("Nam"))
        '        childNode.Tag = "NAM_" & IdParentNode & "_" & dtNam.Rows(i).Item("Nam")
        '        If vYear = dtNam.Rows(i).Item("Nam") Then
        '            getKhenThuongList(childNode, IdParentNode, dtNam.Rows(i).Item("Nam"))
        '            childNode.Expand()
        '        End If
        '    Next
        'End If
        '''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''
        'Dùng thế này cho nhanh
        ''''''''''''''''''''''''
        For i = 2002 To Now.Year + 1
            childNode = parentNode.Nodes.Add(IdParentNode, "Năm " & i)
            childNode.Tag = "NAM_" & IdParentNode & "_" & i
            If vYear = i Then
                getKhenThuongList(childNode, IdParentNode, i, vKT_ChuyenMon)
                childNode.Expand()
            End If
        Next
        Return parentNode
    End Function

    ''' <summary>
    ''' Hàm lấy danh sách quyết định Khen thưởng/Đề nghị khen thưởng trong năm
    ''' </summary>
    ''' <param name="parentNode"></param>
    ''' <param name="IdParentNode"></param>
    ''' <param name="vYear"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getKhenThuongList(ByVal parentNode As TreeNode, ByVal IdParentNode As Integer, ByVal vYear As Integer, ByVal vKT_ChuyenMon As Int16) As TreeNode
        Dim i As Integer
        Dim strSQL As String = ""
        Dim dtDS As DataTable
        Dim db As DBAccess = New DBAccess
        Dim childNode As TreeNode

        While parentNode.Nodes.Count > 0
            parentNode.Nodes.Remove(parentNode.FirstNode)
        End While

        If DONVI = gMaDonViTW And getDonvi_Ma(IdParentNode) = gMaDonViTW Then
            strSQL = " SELECT a.IdKhenthuong, a.KhenThuong, a.DinhKy, a.SoQD, a.NgayQD, a.IdCapKT, a.IdChucVuKyQD " & _
                     " FROM HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong " & _
                     " WHERE a.NamKT=" & vYear & " and KT_ChuyenMon=" & vKT_ChuyenMon & " and KhenThuong =0 " & _
                     "   and (  (b.IdDonVi = " & IdParentNode & ") " & _
                     "        or b.IdDonVi in (SELECT id FROM ChiNhanh WHERE Id_goc=" & IdParentNode & ")	" & _
                     "        or b.IdDonVi in (SELECT id FROM ChiNhanh WHERE Id_goc in (SELECT id FROM ChiNhanh WHERE Id_goc=" & IdParentNode & "))) " & _
                     " UNION " & _
                     " SELECT a.IdKhenthuong, a.KhenThuong, a.DinhKy, a.SoQD, a.NgayQD, a.IdCapKT, a.IdChucVuKyQD " & _
                     " FROM HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong " & _
                     " WHERE a.NamKT=" & vYear & " and KT_ChuyenMon=" & vKT_ChuyenMon & " and KhenThuong =1 " & _
                     "   and b.IdDonVi = " & IdParentNode & _
                     " GROUP BY a.IdKhenthuong, a.KhenThuong, a.DinhKy, a.SoQD, a.NgayQD, a.IdCapKT, a.IdChucVuKyQD" & _
                     " order by NgayQD "
        Else
            strSQL = " SELECT a.IdKhenthuong, a.KhenThuong, a.DinhKy, a.SoQD, a.NgayQD, a.IdCapKT, a.IdChucVuKyQD " & _
                     " FROM HS_KhenThuong a inner join HS_KhenThuong_CT b on a.IdKhenThuong=b.IdKhenThuong " & _
                     " WHERE a.NamKT=" & vYear & " and KT_ChuyenMon=" & vKT_ChuyenMon & _
                     "   and (  (b.IdDonVi = " & IdParentNode & ") " & _
                     "        or b.IdDonVi in (SELECT id FROM ChiNhanh WHERE Id_goc=" & IdParentNode & ")	" & _
                     "        or b.IdDonVi in (SELECT id FROM ChiNhanh WHERE Id_goc in (SELECT id FROM ChiNhanh WHERE Id_goc=" & IdParentNode & "))) " & _
                     " GROUP BY a.IdKhenthuong, a.KhenThuong, a.DinhKy, a.SoQD, a.NgayQD, a.IdCapKT, a.IdChucVuKyQD" & _
                     " order by NgayQD "
        End If
        dtDS = db.SelectDBRows(strSQL)
        If dtDS.Rows.Count > 0 Then
            For i = 0 To dtDS.Rows.Count - 1
                childNode = parentNode.Nodes.Add(IdParentNode, IIf(dtDS.Rows(i).Item("KhenThuong"), "QĐ KT: ", "Đề nghị KT: ") & dtDS.Rows(i).Item("SoQD") & " ngày " & DateTimeUtil.getShortDate(CDate(dtDS.Rows(i).Item("NgayQD"))) & " " & getDanhmuc_Name(4, dtDS.Rows(i).Item("IdChucVuKyQD")))
                childNode.Tag = "CT_" & IdParentNode & "_" & dtDS.Rows(i).Item("IdKhenthuong") & "_" & vYear & "_" & dtDS.Rows(i).Item("KhenThuong") & "_" & dtDS.Rows(i).Item("DinhKy") & "_" & dtDS.Rows(i).Item("SoQD") & "_" & dtDS.Rows(i).Item("NgayQD") & "_" & dtDS.Rows(i).Item("IdCapKT") & "_" & dtDS.Rows(i).Item("IdChucVuKyQD")
            Next
        End If
        Return parentNode
    End Function

    ''' <summary>
    ''' Lấy danh sách các cá nhân, tập thể kèm theo quyết định khen thưởng. Vơi các QĐ toàn quốc sẽ đuợc tổng hợp tại HSC
    ''' </summary>
    ''' <param name="vKhenThuong"></param>
    ''' <param name="vDinhKy"></param>
    ''' <param name="vNamKT"></param>
    ''' <param name="vSoQD"></param>
    ''' <param name="vNgayQD"></param>
    ''' <param name="vIdCapKT"></param>
    ''' <param name="vIdChucVuKyQD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getKhenThuongDetail(ByVal vKhenThuong As Boolean, ByVal vDinhKy As Boolean, ByVal vNamKT As Int32, ByVal vSoQD As String, ByVal vNgayQD As String, ByVal vIdCapKT As Int32, ByVal vIdChucVuKyQD As Int32) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("HS_KhenThuong_CT_GetForQD")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@KhenThuong", vKhenThuong))
        cmd.Parameters.Add(New SqlParameter("@DinhKy", vDinhKy))
        cmd.Parameters.Add(New SqlParameter("@NamKT", vNamKT))
        cmd.Parameters.Add(New SqlParameter("@SoQD", vSoQD))
        cmd.Parameters.Add(New SqlParameter("@NgayQD", vNgayQD))
        cmd.Parameters.Add(New SqlParameter("@IdCapKT", vIdCapKT))
        cmd.Parameters.Add(New SqlParameter("@IdChucVuKyQD", vIdChucVuKyQD))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbKT")
            Return ds.Tables("tbKT")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    ''' <summary>
    ''' Lấy danh sách cá nhân khen thưởng
    ''' </summary>
    ''' <param name="vIdCanbo"></param>
    ''' <param name="vKT_ChuyenMon"></param>
    ''' <param name="vNam"></param>
    ''' <param name="vAll">lấy toàn bộ danh sách: khen thuonwgr va đề nghị khen thưởng</param>
    ''' <param name="vKhenThuong">=1: chỉ lấy các QD được khen thưởng</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getKhenThuongDetail_CaNhan(ByVal vIdCanbo As String, ByVal vKT_ChuyenMon As Int16, ByVal vNam As Integer, ByVal vAll As Int16, ByVal vKhenThuong As Int16) As DataSet
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand = New SqlCommand("getDSKhenThuong_CaNhan")
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdCanbo", vIdCanbo))
        cmd.Parameters.Add(New SqlParameter("@KT_ChuyenMon", vKT_ChuyenMon))
        cmd.Parameters.Add(New SqlParameter("@Nam", vNam))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add(New SqlParameter("@KhenThuong", vKhenThuong))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbKTCN")
            'Return ds.Tables("tbKTCN")
            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    'Public Function getKhenthuongCTList(ByVal vIdKhenThuong As String) As DataTable
    '    Dim dbconn As DBAccess = New DBAccess
    '    Try
    '        Return dbconn.SelectDBRows("SELECT CN_TT,IdCN_TT,IdDanhHieuHinhThuc,TenCN_TT,GhiChu FROM HS_KhenThuong_CT WHERE idKhenThuong ='" & vIdKhenThuong & "' and (IdKhenThuong_CT_parent is null or IdKhenThuong_CT_parent='') Order by IdDonVi,CN_TT desc")
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function

    'Public Function getKT_NhomCaNhan(ByVal vIdKhenThuong_CT_parent As String) As DataTable
    '    Dim dbconn As DBAccess = New DBAccess
    '    Try
    '        Return dbconn.SelectDBRows("SELECT IdCN_TT FROM HS_KhenThuong_CT WHERE IdKhenThuong_CT_parent ='" & vIdKhenThuong_CT_parent & "'")
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function

    Public Function getDanhHieu(ByVal IdDanhHieu As Integer) As String
        Try
            Dim strSql As String
            Dim db As DBAccess = New DBAccess
            strSql = "SELECT DanhHieu_HinhThuc FROM ThiDuaKhenThuong WHERE IdTDKT=" & IdDanhHieu
            Return db.getString(strSql)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Check or uncheck all child node order by parent node
    ''' </summary>
    ''' <param name="tNode"></param>
    ''' <param name="nChecked"></param>
    ''' <remarks></remarks>
    Public Sub CheckSelected(ByVal tNode As TreeNode, ByVal nChecked As Boolean)
        Dim childNode As TreeNode
        Try
            For Each childNode In tNode.Nodes
                childNode.Checked = nChecked
                If childNode.Nodes.Count > 0 Then
                    CheckSelected(childNode, nChecked)
                End If
            Next childNode
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Uncheck parent nodes when any child node is uncheck
    ''' </summary>
    ''' <param name="tNode"></param>
    ''' <remarks></remarks>
    Public Sub ParentUnCheckSelected(ByVal tNode As TreeNode)
        If Not tNode.Parent Is Nothing Then
            tNode.Parent.Checked = False
            ParentUnCheckSelected(tNode.Parent)
        End If
    End Sub

    ''' <summary>
    ''' Danh sách các node trong Treeview checked (Các node đã được load)
    ''' </summary>
    ''' <param name="node"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getList(ByVal node As TreeNodeCollection) As List(Of String)
        Dim list As New List(Of String)
        Dim arr() As String
        For Each childNode As TreeNode In node
            If childNode.Checked Then
                arr = childNode.Tag.ToString.Split("_")
                If (arr(0) = "CB") Then
                    Select Case arr(2)
                        Case "TS"
                            list.Add("TS_" & childNode.Name)
                        Case "NH"
                            list.Add("NH_" & childNode.Name)
                        Case "CC"
                            list.Add("CC_" & childNode.Name)
                        Case "VH"
                            list.Add("VH_" & childNode.Name)
                        Case Else
                            list.Add("CB_" & childNode.Name)
                    End Select
                End If
            End If
            list.AddRange(getList(childNode.Nodes))
        Next
        Return list
    End Function

    Public Function getListCB(ByVal node As TreeNodeCollection) As String
        Dim strReturn As String = ""
        Dim arr() As String
        For Each childNode As TreeNode In node
            If childNode.Checked Then
                arr = childNode.Tag.ToString.Split("_")
                If (arr(0) = "CB") Then
                    Select Case arr(2)
                        Case "TS"
                            strReturn = IIf(strReturn = "", "", ",") & "TS_" & childNode.Name
                        Case "NH"
                            strReturn = IIf(strReturn = "", "", ",") & "NH_" & childNode.Name
                        Case "CC"
                            strReturn = IIf(strReturn = "", "", ",") & "CC_" & childNode.Name
                        Case "VH"
                            strReturn = IIf(strReturn = "", "", ",") & "VH_" & childNode.Name
                        Case Else
                            strReturn = IIf(strReturn = "", "", ",") & "CB_" & childNode.Name
                    End Select
                End If
            End If

        Next
        Return strReturn
    End Function

    Public Function listNghiDinhLuong(Optional ByVal NoneRow As Boolean = False) As DataTable
        Try
            Dim strSql As String = ""
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess

            If NoneRow Then
                strSql = " SELECT '0' as Value, '' as Display UNION "
            End If
            strSql &= " SELECT idNDLuong as Value, TenND as Display  FROM NghiDinhLuong"
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Bảng lương
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listBangLuong(ByVal vIdNghiDinhLuong As Integer, Optional ByVal NoneRow As Boolean = False) As DataTable
        Try
            Dim strSql As String = ""
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            If NoneRow Then
                strSql = " SELECT '0' as Value, '' as Display UNION "
            End If
            strSql &= " SELECT idBangluong as Value, Mota as Display  FROM BangLuong WHERE idND_luong ='" & vIdNghiDinhLuong & "'" ' ORDER BY bangluong"
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Ngạch lương
    ''' </summary>
    ''' <param name="vIdBangLuong"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listNgachLuong(ByVal vIdBangLuong As Integer, Optional ByVal NoneRow As Boolean = False) As DataTable
        Try
            Dim strSql As String = ""
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            If NoneRow Then
                strSql = " SELECT '0' as Value, '' as Display UNION "
            End If
            strSql &= " SELECT IdNgachluong as Value, mota as Display  FROM Ngachluong WHERE idBangluong=" & vIdBangLuong  '& "  ORDER BY Ngachluong"
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Danh sách Bậc lương
    ''' </summary>
    ''' <param name="vIdNgachLuong"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listBacLuong(ByVal vIdNgachLuong As Integer, Optional ByVal MaxVar As Integer = 0, Optional ByVal NoneRow As Boolean = False) As DataTable
        Try
            Dim strSql As String = ""
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            If NoneRow Then
                strSql = " SELECT '0' as Value, '' as Display UNION "
            End If
            If MaxVar = 0 Then
                strSql &= " SELECT IdBacluong as Value, bacluong as Display  FROM bacluong WHERE idNgachluong=" & vIdNgachLuong & " ORDER BY Display" ' desc"
            Else
                strSql &= " SELECT IdBacluong as Value, bacluong as Display  FROM bacluong WHERE idNgachluong=" & vIdNgachLuong & " AND BacLuong<=(SELECT BacLuong FROM BacLuong WHERE IdBacLuong=" & MaxVar & ") ORDER BY Display " ' desc"
            End If
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Function getBacLuong(ByVal vIdBacluong As Integer) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getString("SELECT bacluong FROM Bacluong WHERE idbacluong =" & vIdBacluong)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getTimeNangBac(ByVal vIdNgachLuong As Integer) As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getNumber("SELECT TimeNangNgach FROM NgachLuong WHERE IdNgachLuong=" & vIdNgachLuong)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Hệ số lương
    ''' </summary>
    ''' <param name="vIdBacluong"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getHeso(ByVal vIdBacluong As Integer) As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            Return dbconn.getString("SELECT Heso FROM BacLuong WHERE Idbacluong=" & vIdBacluong.ToString)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Danh sách các Mức phụ cấp tương ứng với từng Loại Phụ cấp
    ''' </summary>
    ''' <param name="vIdLoaiPC">Id Loại phụ cấp</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listMucPC(ByVal vIdLoaiPC As Integer, Optional ByVal NoneRow As Boolean = False) As DataTable
        Try
            Dim strSql As String = ""
            Dim dt As DataTable
            Dim dbconn As DBAccess = New DBAccess
            If NoneRow Then
                strSql = " SELECT '0' as Value, '' as Display UNION "
            End If
            strSql &= " SELECT IdMuc_PhC as Value, Muc_PhC as Display  FROM MucPhuCap WHERE idLoai_PhC=" & vIdLoaiPC & " Order by Display"
            dt = dbconn.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Lương Cán bộ của tháng xác định trong năm
    ''' </summary>
    ''' <param name="vIdCanBo"></param>
    ''' <param name="vMonth"></param>
    ''' <param name="vYear"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLuongCanBo(ByVal vIdCanBo As String, ByVal vMonth As Integer, ByVal vYear As Integer, Optional ByVal TS As Boolean = False) As Long
        Dim ReturnValue As Long
        Dim IdCBTS As String
        Dim t1, t2 As Long
        Try
            Dim dbconn As DBAccess = New DBAccess
            If vMonth = 0 Then
                ReturnValue = 0
                Exit Try
            End If
            If vYear = 0 Then
                ReturnValue = 0
                Exit Try
            End If
            If Not TS Then
                t1 = dbconn.getBigNumber("SELECT dbo.getluongCanBo('" & vIdCanBo & "'," & vMonth & "," & vYear & ")")
                IdCBTS = dbconn.getString("SELECT Id FROM HSCB_TS WHERE IdNew ='" & vIdCanBo & "'")
                If Not (IdCBTS Is Nothing) Then
                    t2 = dbconn.getBigNumber("SELECT dbo.getLuongCanBo_TapSu('" & IdCBTS & "'," & vMonth & "," & vYear & ")")
                End If
                ReturnValue = t1 + t2
            Else
                ReturnValue = dbconn.getBigNumber("SELECT dbo.getLuongCanBo_TapSu('" & vIdCanBo & "'," & vMonth & "," & vYear & ")")
            End If
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' Lương tính theo giờ của cán bộ
    ''' </summary>
    ''' <param name="vIdCanBo"></param>
    ''' <param name="vMonth"></param>
    ''' <param name="vYear"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLuongGioCanBo(ByVal vIdCanBo As String, ByVal vMonth As Integer, ByVal vYear As Integer) As Long
        Dim ReturnValue As Long
        Try
            Dim dbconn As DBAccess = New DBAccess
            If vMonth = 0 Then
                ReturnValue = 0
                Exit Try
            End If
            If vYear = 0 Then
                ReturnValue = 0
                Exit Try
            End If
            ReturnValue = dbconn.getBigNumber("SELECT [dbo].[LuongGio]('" & vIdCanBo & "', " & IdDONVI & "," & vMonth & "," & vYear & ")")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return ReturnValue
    End Function

    Public Function getLuongLamThemGioCanBo(ByVal vIdCanBo As String, ByVal vMonth As Integer, ByVal vYear As Integer, ByVal vSoGio As Double, ByVal vLuongTraThem As Long) As Long
        Dim ReturnValue As Long
        Try
            Dim dbconn As DBAccess = New DBAccess
            If vMonth = 0 Then
                ReturnValue = 0
                Exit Try
            End If
            If vYear = 0 Then
                ReturnValue = 0
                Exit Try
            End If
            ReturnValue = dbconn.getBigNumber("SELECT dbo.getLuongLamThemGioCanBo('" & vIdCanBo & "'," & vMonth & "," & vYear & "," & vSoGio & "," & vLuongTraThem & ")")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' Liệt kê danh sách các quyết định nhân sự của cán bộ gồm các QD trong hệ thống NHCSXH và ngoài NHCSXH
    ''' </summary>
    ''' <param name="vIdCanBo">ID Cán bộ</param>
    ''' <param name="vKind">Các hình thức hiển thị:  
    '''                                            0: Full; 
    '''                                            1: DS các quyết định nhân sự trừ QĐ Cách chức; 
    '''                                            2: DS Hồ sơ công tác; 
    '''                                            3: DS các quyết định nhân sự + QĐ thôi việc
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listQDNhansu(ByVal vIdCanBo As String, ByVal vKind As Int16) As DataTable
        Try
            Dim strSql As String = ""
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            Select Case vKind
                Case 0
                    strSql = " SELECT IdQDNhanSu, NgayHL, So_QD, DVraQD,(SELECT ten_goi FROM Danhmuc WHERE id_goc=15 and id= IDLoaiQD)  as 'LoaiQD', " & _
                             "        ((SELECT ten_goi FROM Danhmuc WHERE id_goc=14 and id= IDChucVu_Moi)+', '+ ISNULL((SELECT ten_phong  + ', ' FROM PhongBan WHERE id=IdPhong_Moi and id<>15 and id<>21),'')+(SELECT Ten_goi FROM Chinhanh WHERE id=IDDonVi_Moi)) as 'NoiDung'  " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 " & _
                             " UNION" & _
                             " SELECT IdQDNhanSu, NgayHL, So_QD, DVraQD, LoaiQD, NoiDung " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=0 order by NgayHL desc "
                Case 1
                    Dim IdCachChuc As Integer = db.getNumber("SELECT id FROM DanhMuc WHERE ma_so='1510'")
                    strSql = " SELECT IdQDNhanSu, NgayHL, So_QD, DVraQD,(SELECT ten_goi FROM Danhmuc WHERE id_goc=15 and id= IDLoaiQD)  as 'LoaiQD', " & _
                             "        ((SELECT ten_goi FROM Danhmuc WHERE id_goc=14 and id= IDChucVu_Moi)+' '+ ISNULL((SELECT ten_phong  + ', ' FROM PhongBan WHERE id=IdPhong_Moi and id<>15 and id<>21),'')+(SELECT Ten_goi FROM Chinhanh WHERE id=IDDonVi_Moi)) as 'NoiDung'  " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 AND IDLoaiQD<>" & IdCachChuc & _
                             " UNION" & _
                             " SELECT IdQDNhanSu, NgayHL, So_QD, DVraQD, LoaiQD, NoiDung " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=0 order by NgayHL desc "
                Case 2
                    strSql = " SELECT IdQDNhanSu, NgayHL, So_QD, DVraQD,(SELECT ten_goi FROM Danhmuc WHERE id_goc=15 and id= IDLoaiQD)  as 'LoaiQD', " & _
                             "        ((SELECT ten_goi FROM Danhmuc WHERE id_goc=14 and id= IDChucVu_Moi)+' '+ ISNULL((SELECT ten_phong  + ', ' FROM PhongBan WHERE id=IdPhong_Moi and id<>15 and id<>21),'') + (SELECT Ten_goi FROM Chinhanh WHERE id=IDDonVi_Moi)) as 'NoiDung',  " & _
                             "        (SELECT ma_so FROM DanhMuc WHERE id=IdloaiQD) as MaLoaiQD, DenNgay, " & _
                             "         '1' as IsQD_NHCS, IDDonVi_Moi, IdPhong_Moi, IDChucVu_Moi " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 " & _
                             " UNION" & _
                             " SELECT IdQDNhanSu, NgayHL, So_QD, DVraQD, LoaiQD, NoiDung, (SELECT ma_so FROM DanhMuc WHERE id=IdloaiQD) as MaLoaiQD, DenNgay, " & _
                             "         '0' as IsQD_NHCS, '0' as IDDonVi_Moi, '0' as IdPhong_Moi, '0' as IDChucVu_Moi " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=0 " & _
                             " UNION" & _
                             " SELECT IdCBThoiViec as 'IdQDNhanSu', Ngay_HL as 'NgayHL', So_QD, DVraQD, (SELECT ten_goi FROM DanhMuc WHERE [id]=IdLoaiQD) as 'LoaiQD', (SELECT ten_goi FROM DanhMuc WHERE [id]=IdLyDo) as 'NoiDung', " & _
                             "         (SELECT ma_so FROM DanhMuc WHERE id=IdloaiQD) as MaLoaiQD, '' as DenNgay, '1' as IsQD_NHCS, '0' as IDDonVi_Moi, '0' as IdPhong_Moi, '0' as IDChucVu_Moi" & _
                             "    FROM HS_CBthoiviec WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 order by NgayHL "

                Case 3
                    strSql = " SELECT IdQDNhanSu as 'IdQD', 'QDNhanSu' as table_name, 'IdQDNhanSu' as table_key, NgayHL, So_QD, DVraQD,(SELECT ten_goi FROM Danhmuc WHERE id_goc=15 and id= IDLoaiQD)  as 'LoaiQD', " & _
                             "        ((SELECT ten_goi FROM Danhmuc WHERE id_goc=14 and id= IDChucVu_Moi)+' '+ ISNULL((SELECT ten_phong + ', ' FROM PhongBan WHERE id=IdPhong_Moi and id<>15 and id<>21),'')+(SELECT Ten_goi FROM Chinhanh WHERE id=IDDonVi_Moi)) as 'NoiDung'  " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 " & _
                             " UNION" & _
                             " SELECT IdQDNhanSu as 'IdQD', 'QDNhanSu' as table_name, 'IdQDNhanSu' as table_key, NgayHL, So_QD, DVraQD, LoaiQD, NoiDung " & _
                             "    FROM QDNhanSu WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=0 " & _
                             " UNION" & _
                             " SELECT IdCBThoiViec as 'IdQD', 'HS_CBthoiviec' as table_name, 'IdCBThoiViec' as table_key, Ngay_HL as 'NgayHL', So_QD, DVraQD, (SELECT ten_goi FROM DanhMuc WHERE [id]=IdLoaiQD) as 'LoaiQD', (SELECT ten_goi FROM DanhMuc WHERE [id]=IdLyDo) as 'NoiDung' " & _
                             "    FROM HS_CBthoiviec WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 order by NgayHL desc "
            End Select

            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Liệt kê danh sách các quyết định luong của cán bộ gồm các QD trong hệ thống NHCSXH và ngoài NHCSXH
    ''' </summary>
    ''' <param name="vIdCanBo">ID Cán bộ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listQDLuong(ByVal vIdCanBo As String) As DataTable
        Try
            Dim strSql As String
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            strSql = " SELECT IdLuongCB, Ngay_Huong, SoQD, DVraQD, LoaiQD, NoiDung " & _
                     "       FROM HS_LuongCB WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=0  AND IDLoaiQD=0 " & _
                     " UNION " & _
                     " SELECT IdLuongCB, Ngay_Huong, SoQD, DVraQD, (SELECT ten_goi FROM Danhmuc WHERE id_goc=43 and id= IDLoaiQD)  as 'LoaiQD', NoiDung " & _
                     "       FROM HS_LuongCB WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=0  AND IDLoaiQD>0 " & _
                     " UNION " & _
                     " SELECT IdLuongCB, Ngay_Huong, SoQD, DVraQD,(SELECT ten_goi FROM Danhmuc WHERE id_goc=43 and id= IDLoaiQD)  as 'LoaiQD', " & _
                     "       (N'Hệ số '+ convert(varchar(4),heso) + '; '+ MotaBacLuong + N'; Ngạch '+ MotaNgachLuong  + N'; Bảng lương '+ MotaBangLuong +'; '+ TenND  ) as 'NoiDung'  " & _
                     "       FROM V$_HSLuongCB WHERE idCanbo='" & vIdCanBo & "' AND IsQD_NHCS=1 order by Ngay_Huong desc "
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Liệt kê danh sách các quyết định phu cap của cán bộ gồm các QD trong hệ thống NHCSXH và ngoài NHCSXH
    ''' </summary>
    ''' <param name="vIdCanBo">ID Cán bộ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listQDPhucap(ByVal vIdCanBo As String) As DataTable
        Try
            Dim strSql As String
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            strSql = " SELECT IdCB_PhuCap, TuNgay, SoQD, DVraQD, " & _
                     "        (CASE IsQD_NHCS WHEN 1 THEN" & _
                     "        (SELECT ten_goi + ': ' + convert(varchar(4),Muc_PhC) FROM mucphucap, danhmuc WHERE Idloai_PhC= id AND idMuc_PhC = IDMucPC) " & _
                     "          ELSE Noidung END ) as 'NoiDung' " & _
                     " FROM HS_PhuCapCB WHERE idCanbo='" & vIdCanBo & "' order by TuNgay desc, Date_create desc "
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Liet ke cac quyet dinh nhan su khac cua can bo
    ''' </summary>
    ''' <param name="vIdCanBo">ID Cán bộ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function listQDKhac(ByVal vIdCanBo As String) As DataTable
        Try
            Dim strSql As String
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            strSql = "  SELECT IdQDKhac, NgayKy_QD, So_QD, DVraQD, NoiDung " & _
                     " FROM QDKhac WHERE idCanbo='" & vIdCanBo & "' order by NgayKy_QD desc "
            dt = db.SelectDBRows(strSql)
            Return dt
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Kiểm tra Số Quyết định cho 1 cán bộ đã tồn tại chưa ?
    ''' </summary>
    ''' <param name="vHoso">LUONG, PHUCAP,...</param>
    ''' <param name="vIdCanbo"></param>
    ''' <param name="vSoQD"></param>
    ''' <returns>1: đã tồ tại, 0: chưa tồn tại</returns>
    ''' <remarks></remarks>
    Public Function checkQuyetDinh(ByVal vHoso As String, ByVal vIdCanbo As String, ByVal vSoQD As String) As Boolean
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT dbo.existsQuyetDinh('" & vHoso & "','" & vIdCanbo & "', N'" & vSoQD & "')")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return CBool(ReturnValue)
    End Function

    ''' <summary>
    ''' Hàm kiểm tra Loại quyết định lương có được phép nhập bên Hồ sơ lương không (có 4 loại).  
    ''' </summary>
    ''' <param name="vIdLoaiQDLuong"></param>
    ''' <returns>Nếu không thuôc--> chỉ được phép đọc, không được thay đổi: ReadOnly=true</returns>
    ''' <remarks></remarks>
    Public Function checkQDLuong_ReadOnly(ByVal vIdLoaiQDLuong As Integer) As Boolean
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT Id FROM DanhMuc WHERE Id=" & vIdLoaiQDLuong & " AND ma_so in (SELECT ma_so FROM DanhMuc WHere id_goc=43 and ma_so not in ('4301','4302','4303','4304'))")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return CBool(ReturnValue)
    End Function

    ''' <summary>
    ''' Hàm kiểm tra mã cán bộ có hợp lệ không. 
    ''' </summary>
    ''' <param name="vIdCanBo"></param>
    ''' <param name="vMaCanBo"></param>
    ''' <returns>Hợp lệ:1; Không hợp lê:0</returns>
    ''' <remarks></remarks>
    Public Function checkMaCanBo(ByVal vIdDonVi As Integer, ByVal vIdCanBo As String, ByVal vMaCanBo As String) As Boolean
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT dbo.checkMaCanBo(" & vIdDonVi & ",'" & vIdCanBo & "','" & vMaCanBo & "')")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return CBool(ReturnValue)
    End Function

    ''' <summary>
    ''' Hàm sinh giá trị khóa của các bảng hồ sơ cán bộ
    ''' </summary>
    ''' <param name="vTableName">Tên bảng</param>
    ''' <param name="vKeyName">Tên trường khóa, sinh value</param>
    ''' <returns>Giá trị khóa</returns>
    ''' <remarks></remarks>
    Public Function setKeyValueTable(ByVal vTableName As String, ByVal vKeyName As String) As String
        Dim db As DBAccess = New DBAccess
        Try
            Dim cmd As SqlCommand = New SqlCommand("setKey")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@Id_Donvi", IdDONVI))
            cmd.Parameters.Add(New SqlParameter("@tableName", vTableName))
            cmd.Parameters.Add(New SqlParameter("@fieldKey", vKeyName))
            cmd.Parameters.Add("@idOUT", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            setKeyValueTable = cmd.Parameters("@idOUT").Value.ToString
        Catch ex As Exception
            setKeyValueTable = ""
        End Try
    End Function

    ''' <summary>
    ''' Sinh mã cán bộ tiếp theo
    ''' </summary>
    ''' <param name="vIdDonVi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function setMaCanBo(ByVal vIdDonVi As Integer) As String
        Dim ReturnValue As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getString("SELECT dbo.setMaCanBo(" & vIdDonVi & ")")
        Catch ex As Exception
            ReturnValue = ""
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' Hàm insert table CN_Chuyen_CN
    ''' </summary>
    ''' <param name="vTG_Chuyen">Thời gian hiệu lực chuyển</param>
    ''' <param name="vIDCNA">CN chuyển đi</param>
    ''' <param name="vIDCNB">CN chuyển tới</param>
    ''' <param name="vDi_Den">Di=1; Đến =2</param>
    ''' <param name="vIDCanBo">ID cán bộ được chuyển</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function insertCN_Chuyen_CN(ByVal vTG_Chuyen As Date, ByVal vIDCNA As String, ByVal vIDCNB As String, ByVal vDi_Den As String, ByVal vIDCanBo As String) As String
        Dim db As DBAccess = New DBAccess
        Try
            Dim cmd As SqlCommand = New SqlCommand("CB_Chuyen_CN_Insert")
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@ThoigianChuyen", vTG_Chuyen))
            cmd.Parameters.Add(New SqlParameter("@IdCNA", vIDCNA))
            cmd.Parameters.Add(New SqlParameter("@IdCNB", vIDCNB))
            cmd.Parameters.Add(New SqlParameter("@Di_Den", vDi_Den))
            cmd.Parameters.Add(New SqlParameter("@IdCanBo", vIDCanBo))
            cmd.Parameters.Add("@IDCB_Chuyen_CN", SqlDbType.VarChar, 15).Direction = ParameterDirection.Output
            db.executeSQL(cmd)
            insertCN_Chuyen_CN = cmd.Parameters("@IDCB_Chuyen_CN").Value.ToString
        Catch ex As Exception
            insertCN_Chuyen_CN = ""
        End Try
    End Function

    ''' <summary>
    ''' Kiểm tra Số quyết định khen thưởng đã tồn tại chưa
    ''' </summary>
    ''' <param name="vNam"></param>
    ''' <param name="vSoQD"></param>
    ''' <param name="vCapKT"></param>
    ''' <returns>Neu ton tai tra ve IdKhenThuong</returns>
    ''' <remarks></remarks>
    Public Function checkQuyetDinhKhenThuong(ByVal vNam As Integer, ByVal vSoQD As String, ByVal vCapKT As Integer) As String
        Dim ReturnValue As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getString("SELECT dbo.existsQuyetDinhKhenThuong(" & vNam & ", N'" & vSoQD & "'," & vCapKT & ")")
            If ReturnValue = "0" Then ReturnValue = ""
        Catch ex As Exception
            ReturnValue = ""
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' Kiểm tra Số kế hoạch lao động của đơn vị trong năm đã tồn tại chưa ?
    ''' </summary>
    ''' <param name="vIdDonVi">Chi nhánh</param>
    ''' <param name="vNam">năm</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function checkSoKeHoachLaoDong(ByVal vIdDonVi As Integer, ByVal vNam As Integer, ByVal vThang As Integer) As Boolean
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT [dbo].[existsKeHoachLaoDong_DV](" & vIdDonVi & "," & vNam & ", " & vThang & ")")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return CBool(ReturnValue)
    End Function

    'zzzzzzzzzzzzzzzzzzzzzzzzzzzz
    ''' <summary>
    ''' Check xem Don vi của user hiện tại có quan hệ với Don vi của record đang được chọn hay không
    ''' 
    ''' TẠM THỜI BỎ QUA CHECK
    ''' </summary>
    ''' <param name="vIdCanbo"></param>
    ''' <returns>Dung -> true</returns>
    ''' <remarks></remarks>
    Public Function checkRight_CreateRecord(ByVal vIdCanbo As String) As Boolean
        'Dim dbconn As New DBAccess
        'Dim vIdDonVi As Integer = dbconn.getNumber("SELECT idDonVi FROM HS_CanBo WHERE IdCanBo = '" & vIdCanbo & "'")
        'Try
        '    If vIdDonVi = IdDONVI Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'Catch ex As Exception

        'End Try

        Return True
    End Function
    'zzzzzzzzzzzzzzzzzzzzzzzzzzzzz

    ' ''' <summary>
    ' ''' Check xem Don vi truy cap vao record co phai la don vi tao ban ghi do ko
    ' ''' </summary>
    ' ''' <param name="vIdCanbo"></param>
    ' ''' <returns>Dung -> true</returns>
    ' ''' <remarks></remarks>
    'Public Function checkRight_CreateRecord(ByVal vIdCanbo As String) As Boolean
    '    Try
    '        If vIdCanbo.Substring(0, 4) = TEN_DV_VT Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch ex As Exception

    '    End Try

    'End Function

    ''' <summary>
    ''' Kiểm tra số liệu của mạng lưới hoạt động của đơn vị trong năm, quý đã tồn tại chưa ?
    ''' </summary>
    ''' <param name="vIdDonVi"></param>
    ''' <param name="vNam"></param>
    ''' <param name="vQuy"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function checkMangLuoiDV(ByVal vIdDonVi As Integer, ByVal vNam As Integer, ByVal vQuy As Int16) As Boolean
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT [dbo].[existsMangLuoi_DV](" & vIdDonVi & "," & vNam & ", " & vQuy & ")")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return CBool(ReturnValue)
    End Function

    Public Sub getIdNDBangNgach(ByVal vIdBacLuong As Integer, ByRef vIdNghiDinhLuong As Integer, ByRef vIdBangLuong As Integer, ByRef vIdNgachLuong As Integer)
        Try
            Dim strSql As String
            strSql = "SELECT idNDLuong, idBangLuong, idNgachLuong FROM V$_BacLuong WHERE idBacLuong=" & vIdBacLuong
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            dt = db.SelectDBRows(strSql)
            vIdNghiDinhLuong = dt.Rows(0).Item("idNDLuong")
            vIdBangLuong = dt.Rows(0).Item("idBangLuong")
            vIdNgachLuong = dt.Rows(0).Item("idNgachLuong")
        Catch ex As Exception
        End Try
    End Sub

    Public Sub getIdBangNgachBac(ByVal vIdCanBo As String, ByRef vIdBangLuong As Integer, ByRef vIdNgachLuong As Integer, ByRef vIdBacLuong As Integer, ByRef vHeso As String)
        Try
            If vIdCanBo <> "" Then
                Dim strSql As String
                strSql = "SELECT TOP 1 IdBangLuong, IdNgachLuong, IdBacLuong, HeSo FROM V$_HSLuongCB WHERE  idcanbo ='" & vIdCanBo & "' order by Ngay_Huong desc"
                Dim dt As DataTable
                Dim db As DBAccess = New DBAccess
                dt = db.SelectDBRows(strSql)
                vIdBangLuong = dt.Rows(0).Item("idBangLuong")
                vIdNgachLuong = dt.Rows(0).Item("idNgachLuong")
                vIdBacLuong = dt.Rows(0).Item("IdBacLuong")
                vHeso = dt.Rows(0).Item("HeSo")
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' lay thong tin lien quan den loai va muc phu cap hien huong cua tung can bo
    ''' </summary>
    ''' <param name="vIdCanBo"></param>
    ''' <param name="vIdLoaiPC"></param>
    ''' <remarks></remarks>
    Public Sub getIdLoaiMucPC(ByVal vIdCanBo As String, ByRef vIdLoaiPC As Integer, ByRef vIdMucPC As Integer)
        Try
            If vIdCanBo <> "" Then
                Dim strSql As String
                strSql = "SELECT TOP 1 idLoai_PhC, IdMucPC FROM HS_PhuCapCB a, MucPhucap b WHERE a.IdMucPC=b.IdMuc_PhC and  a.idcanbo ='" & vIdCanBo & "'  order by TuNgay desc"
                Dim dt As DataTable
                Dim db As DBAccess = New DBAccess
                dt = db.SelectDBRows(strSql)
                vIdLoaiPC = dt.Rows(0).Item("idLoai_PhC")
                vIdMucPC = dt.Rows(0).Item("IdMucPC")
            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' get ID Loại phụ cấp theo ID Mức phụ cấp
    ''' </summary>
    ''' <param name="vIdMucPC"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getIdLoaiPC(ByVal vIdMucPC As Integer) As Integer
        Try
            Dim db As DBAccess = New DBAccess
            Return db.getNumber("SELECT idLoai_PhC FROM MucPhucap WHERE idMuc_PhC=" & vIdMucPC)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' get Loại phụ cấp
    ''' </summary>
    ''' <param name="vIdLoaiPC">Id Loại Phụ cấp</param>
    ''' <param name="vIdMucPC">Id Mức phụ cấp</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLoaiPC(Optional ByVal vIdLoaiPC As Integer = 0, Optional ByVal vIdMucPC As Integer = 0) As String
        Try
            Dim strSql As String = ""
            Dim db As DBAccess = New DBAccess
            If vIdLoaiPC > 0 Then strSql = "SELECT ten_goi FROM Danhmuc WHERE id_goc=31 and id=" & vIdLoaiPC
            If vIdMucPC > 0 Then strSql = "SELECT ten_goi FROM Danhmuc WHERE id in (SELECT IdLoai_PhC FROM Mucphucap WHERE IdMuc_PhC =" & vIdMucPC & ")"
            Return db.getString(strSql)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getInfPhuCap(ByVal vIdMucPhuCap As Integer) As String
        Dim strReturn As String = ""
        Try
            Dim db As DBAccess = New DBAccess
            Dim dt As DataTable
            dt = db.SelectDBRows("SELECT ten_goi, Muc_PhC FROM mucphucap, danhmuc WHERE Idloai_PhC= id AND idMuc_PhC=" & vIdMucPhuCap)
            If dt.Rows.Count > 0 Then
                strReturn = dt.Rows(0).Item("ten_goi") & ": " & dt.Rows(0).Item("Muc_PhC")
            End If
        Catch ex As Exception
        End Try
        Return strReturn
    End Function

    Public Sub bindCboBangCap(ByVal cbo As ComboBox)
        'Value: 1: Đại học; 2: Bằng A; 3: Bằng B; 4: Bằng C
        cbo.Items.Add("")
        cbo.Items.Add("Đại học")
        cbo.Items.Add("Bằng A")
        cbo.Items.Add("Bằng B")
        cbo.Items.Add("Bằng C")
        If cbo.SelectedValue = 0 Then
            cbo.SelectedIndex = 1
        End If
    End Sub

    Public Sub bindCboDanhGia(ByVal cbo As ComboBox)
        'Value: 1; TỐT; 2: BÌNH THƯỜNG; 3: CON HẠN CHẾ
        cbo.Items.Add("")
        cbo.Items.Add("Tốt")
        cbo.Items.Add("Bình thường")
        cbo.Items.Add("Còn hạn chế")
        If cbo.SelectedValue = 0 Then
            cbo.SelectedIndex = 1
        End If
    End Sub

    Public Function getBangCap(ByVal value As Byte) As String
        If value = 1 Then
            Return "Đại học"
        ElseIf value = 2 Then
            Return "Bằng A"
        ElseIf value = 3 Then
            Return "Bằng B"
        ElseIf value = 4 Then
            Return "Bằng C"
        Else
            Return ""
        End If
    End Function

    Public Function getDanhgia(ByVal value As Byte) As String
        If value = 1 Then
            Return "Tốt"
        ElseIf value = 2 Then
            Return "Bình thường"
        ElseIf value = 3 Then
            Return "Còn hạn chế"
        Else
            Return ""
        End If
    End Function

    Public Function getID_HS_NghiPhep(ByVal vIdCanbo As String, ByVal vTuNgay As Date, ByVal vTyleHuong As Double, ByVal vNoiNghi As String, ByVal vLydo As String) As String
        Dim ReturnValue As String
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getString("SELECT [dbo].[getID_HS_NgPh] ('" & vIdCanbo & "', '" & vTuNgay & "', " & vTyleHuong & " , '" & vNoiNghi & "', '" & vLydo & "')")
        Catch ex As Exception
            ReturnValue = ""
        End Try
        Return ReturnValue
    End Function

    Public Sub updateSalaryVar_DateEnd(ByVal vIdSalaryVar As Integer, ByVal vNameVar As String, ByVal vDateApply As Date)
        Try
            Dim dbconn As DBAccess = New DBAccess
            dbconn.executeSQL("exec [dbo].[UpdateSalaryVar] " & vIdSalaryVar & ",'" & vNameVar & "', '" & vDateApply & "'")
        Catch ex As Exception
        End Try
    End Sub

    Public Function countworkdays(ByVal vTuNgay As Date, ByVal vDenNgay As Date) As String
        Dim ReturnValue As Integer
        Try
            Dim dbconn As DBAccess = New DBAccess
            ReturnValue = dbconn.getNumber("SELECT [dbo].[countworkdays] ( '" & vTuNgay & "', '" & vDenNgay & "' )")
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return ReturnValue
    End Function

    Public Sub getInfThuChi(ByVal vIdHS_ChiLuong As String, ByVal vLoaiThuChi As Integer, ByRef strLoaiThuChi As String, ByRef vKy As Integer, ByRef strKy As String, ByRef vLoaiCB As Int16, ByRef strLoaiCB As String)
        Dim dbconn As DBAccess = New DBAccess
        Dim strSql As String = ""
        Dim vMaThuChi As String = getDanhmuc_MaSo(vLoaiThuChi)
        Dim dt As DataTable

        Select Case vMaThuChi
            Case "4201"
                strLoaiThuChi = "Lương tháng "
                strSql = "SELECT TOP 1 b.Ky, b.NH_TS FROM HS_ChiLuong a left join HS_ChiLuong_CT b on a.IdHS_ChiLuong=b.IdHS_ChiLuong WHERE a.IdHS_ChiLuong = '" & vIdHS_ChiLuong & "' AND b.Ky is not null"
                dt = dbconn.SelectDBRows(strSql)
                If dt.Rows.Count > 0 Then
                    If dt.Rows(0).Item("Ky") > 0 Then
                        vKy = dt.Rows(0).Item("Ky")
                        strKy = " kỳ " & vKy
                    Else
                        vKy = 0
                        strKy = ""
                    End If
                    If dt.Rows(0).Item("NH_TS") = 0 Then
                        vLoaiCB = 0
                        strLoaiCB = ""
                    Else
                        If dt.Rows(0).Item("NH_TS") = 1 Then
                            vLoaiCB = 1
                            strLoaiCB = " (NH) "
                        End If
                        If dt.Rows(0).Item("NH_TS") = 2 Then
                            vLoaiCB = 2
                            strLoaiCB = " (TS) "
                        End If
                    End If
                Else
                    strLoaiThuChi = ""
                    vKy = 0
                    strKy = ""
                    vLoaiCB = 0
                    strLoaiCB = ""
                End If
            Case "4204"   ' lam them gio
                vKy = 0
                strKy = ""
                strLoaiThuChi = "Thêm giờ tháng "
            Case Else
                vKy = 0
                strKy = ""
                strSql = "SELECT TOP 1 b.NH_TS, b.IdLoaiLuongBS, (SELECT ma_so FROM DanhMuc WHERE id_goc=42 AND id=b.IdLoaiLuongBS) as 'maso' FROM HS_ChiLuong a left join HS_ChiBoSung_CT b on a.IdHS_ChiLuong=b.IdHS_ChiLuong WHERE a.IdHS_ChiLuong = '" & vIdHS_ChiLuong & "' "
                dt = dbconn.SelectDBRows(strSql)
                If dt.Rows.Count > 0 Then
                    vLoaiCB = dt.Rows(0).Item("NH_TS")
                    Select Case dt.Rows(0).Item("maso")
                        Case "4204"
                            strLoaiThuChi = "BHXH "
                        Case "4205"
                            strLoaiThuChi = "Truy lĩnh "
                        Case "4206"
                            strLoaiThuChi = "Truy thu "
                        Case Else
                            strLoaiThuChi = "Chi bổ sung "
                    End Select
                    Select Case dt.Rows(0).Item("NH_TS")
                        Case "1"
                            strLoaiCB = "(NH)"
                        Case "2"
                            strLoaiCB = "(TS)"
                        Case "3"
                            strLoaiCB = "(CC)"
                        Case "4"
                            strLoaiCB = "(VH)"
                        Case Else
                            strLoaiCB = ""
                    End Select
                End If
        End Select
    End Sub

    Public Function checkChiLuongTmp(ByVal vTmp As Integer) As String
        Dim ReturnValue As String = ""
        If vTmp <> 0 Then ReturnValue = " (Chưa kiểm soát)"
        Return ReturnValue
    End Function

    Public Function openApplication(ByVal vDate As Date) As Int16
        Dim ReturnValue As Int16
        Try
            Dim dbconn As DBAccess = New DBAccess
            Dim conn As SqlConnection = dbconn.getConnection
            Dim cmd As SqlCommand = New SqlCommand("openApplication")
            Dim ds As New DataSet
            Dim da As SqlDataAdapter

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@DateOpenApp", vDate))
            cmd.Parameters.Add("@return", SqlDbType.SmallInt).Direction = ParameterDirection.Output
            cmd.Connection = conn
            Try
                conn.Open()
                da = New SqlDataAdapter(cmd)
                da.Fill(ds, "OP")
                ReturnValue = cmd.Parameters("@return").Value
                Return ReturnValue
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                dbconn.closeConnection(conn)
            End Try
        Catch ex As Exception
            ReturnValue = 0
        End Try
        Return ReturnValue

    End Function
End Module
