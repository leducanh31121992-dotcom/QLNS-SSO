Public Class DanhMucForm

#Region "---> Defined parametter and properties <---"
    Dim _DanhMuc As clsHT_DanhMuc = New clsHT_DanhMuc

    ''' <summary>
    ''' Biến lưu cở xác định Form hiển thị dữ liệu của DM nào. 1 - Danh mục chung. 2 - Chi nhánh. 3 - Địa danh
    ''' </summary>
    ''' <remarks></remarks>
    Public _FlagState As Byte

    Private _SqlHelper As DBAccess

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        _SqlHelper = New DBAccess
    End Sub

    ''' <summary>
    ''' Biến lưu tag của node gốc trên treeview chi nhánh
    ''' </summary>
    ''' <remarks></remarks>
    Private _NodeRoot As String = ""

    ''' <summary>
    ''' Chỉ số xác định danh mục cha 
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Private _RootId As Integer = 0

    ''' <summary>
    ''' Chỉ số xác định bản ghi select
    ''' </summary>
    ''' <remarks></remarks>
    Private _RecordId As Integer = 0
    'Cờ bắt khi click node các đơn vị trực thuộc gồm (TTCNTT,TT Đào tạo, sở giao dịch, Văn phòng miền trực thuộc)
    Private _FlagTT As Boolean = False
    'Cờ báo khi click node các phòng ban trực thuộc hội sở chính
    Private _FlagPB As Boolean = False

    'Biến lưu lại Tag của node trên treeview khi select
    Private _NodeTag As String = ""
    Private obj_update As DMKhacForm
#End Region

#Region "---> Functions main: Các hàm chính dùng cho chương trình <---"
    ''' <summary>
    ''' Hàm tìm kiếm dl trong hệ thống danh mục (Hệ thống danh mục chung - Chi nhánh - Địa danh)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Search_Lists()
        Dim _Name = Globals.Find_Replace(edt_tengoi.Text.ToString().Trim())
        Dim _Code = Globals.Find_Replace(edt_maso.Text.ToString().Trim())
        dgv_main.Rows.Clear()
        Dim strSQL As String = ""
        Select Case _FlagState
            Case 1      'TÌM DL - HỆ THỐNG DANH MỤC CHUNG
                If (_RootId = 100) Then     'Tìm kiếm dữ liệu trong danh mục phòng ban
                    Try
                        If (_RecordId > 0) Then
                            strSQL = String.Format("Select * from PhongBan Where id = {0}", _RecordId)
                            If (_Name <> "") Then
                                strSQL += String.Format(" and ten_phong Like N'%{0}%'", _Name)
                            End If
                            If (_Code <> "") Then
                                strSQL += String.Format(" and ma_so Like N'%{0}%'", _Code)
                            End If
                        Else
                            strSQL = String.Format("Select * from PhongBan")
                            If (_Name <> "" And _Code <> "") Then
                                strSQL += String.Format(" Where ten_phong Like N'%{0}%' and ma_so Like N'%{1}%'", _Name, _Code)
                            Else
                                If (_Name <> "") Then
                                    strSQL += String.Format(" Where ten_phong Like N'%{0}%'", _Name)
                                End If
                                If (_Code <> "") Then
                                    strSQL += String.Format(" Where ma_so Like N'%{0}%'", _Code)
                                End If
                            End If
                        End If
                        If (strSQL = "") Then Return
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_phong").ToString() <> "", db.Rows(i)("ten_phong").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString() = "1", " Sử dụng", " Tạm khoá")
                                    Next
                                End If
                            End If
                        End Using
                    Catch ex As Exception
                        MessageBox.Show("Tìm kiếm thông tin phòng ban: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                ElseIf (_RootId = 101) Then     'Tìm kiếm dữ liệu trong danh mục Quốc gia
                    Try
                        If (_RecordId > 0) Then
                            strSQL = String.Format("Select * from QuocGia Where id = {0}", _RecordId)
                            If (_Name <> "") Then
                                strSQL += String.Format(" and ten_goi Like N'%{0}%'", _Name)
                            End If
                            If (_Code <> "") Then
                                strSQL += String.Format(" and ma_so Like N'%{0}%'", _Code)
                            End If
                        Else
                            strSQL = String.Format("Select * from QuocGia")
                            If (_Name <> "" And _Code <> "") Then
                                strSQL += String.Format(" Where ten_goi Like N'%{0}%' and ma_so Like N'%{1}%'", _Name, _Code)
                            Else
                                If (_Name <> "") Then
                                    strSQL += String.Format(" Where ten_goi Like N'%{0}%'", _Name)
                                End If
                                If (_Code <> "") Then
                                    strSQL += String.Format(" Where ma_so Like N'%{0}%'", _Code)
                                End If
                            End If
                        End If
                        If (strSQL = "") Then Return
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString() = "1", " Sử dụng", " Tạm khoá")
                                    Next
                                End If
                            End If
                        End Using
                    Catch ex As Exception
                        MessageBox.Show("Tìm kiếm thông tin Quốc gia: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                Else    'Tìm kiếm dl trong Hệ thống danh mục chung
                    Try
                        If (_RootId <= 0) Then
                            If (_RecordId > 0) Then
                                strSQL = String.Format("Select * from DanhMuc Where id_goc != 0 and id = {0}", _RecordId)
                                If (_Name <> "") Then
                                    strSQL += " and ten_goi Like N'%" + _Name + "%'"
                                End If
                                If (_Code <> "") Then
                                    strSQL += " and ma_so Like N'%" + _Code + "%'"
                                End If
                            End If
                        Else
                            strSQL = String.Format("Select * from DanhMuc Where id_goc != 0 and id_goc = {0}", _RootId)
                            If (_Name <> "") Then
                                strSQL += " and ten_goi Like N'%" + _Name + "%'"
                            End If
                            If (_Code <> "") Then
                                strSQL += " and ma_so Like N'%" + _Code + "%'"
                            End If
                        End If
                        If (strSQL = "") Then Return
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString() = "1", " Sử dụng", " Tạm khoá")
                                    Next
                                End If
                            End If
                        End Using
                    Catch ex As Exception
                        MessageBox.Show("Tìm kiếm thông tin danh mục: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
            Case 2      'TÌM DL - CHI NHÁNH
                _DanhMuc.Create_Frame(dgv_main, _FlagState)
                dgv_main.Rows.Clear()
                If (_NodeRoot <> "") Then       'Trường hợp Node là node gốc "NHCSXH"--> Tìm kiếm thông tin tổng hợp
                    Dim _rowCount As Integer = 0
                    '--> Đầu tiên đi tìm xem có phải dữ liệu muốn tìm là Hội sở chính không ?
                    strSQL = String.Format("Select * from ChiNhanh Where id_goc = 0 and ten_goi Like N'%{0}%' and ma_so Like N'{1}'", _Name, _Code)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Alias").Value = IIf(db.Rows(i)("ten_vt").ToString() <> "", db.Rows(i)("ten_vt").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Address").Value = IIf(db.Rows(i)("dia_chi").ToString() <> "", db.Rows(i)("dia_chi").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Tel").Value = IIf(db.Rows(i)("dien_thoai").ToString() <> "", db.Rows(i)("dien_thoai").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Email").Value = IIf(db.Rows(i)("email").ToString() <> "", db.Rows(i)("email").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Web").Value = IIf(db.Rows(i)("website").ToString() <> "", db.Rows(i)("website").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Villagers").Value = IIf(db.Rows(i)("solg_xaphuong").ToString() <> "", db.Rows(i)("solg_xaphuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm xoá")
                                    _rowCount = _rowCount + 1
                                Next
                            End If
                        End If

                        '--> Tiếp theo tìm kiếm các Chi nhánh cấp tỉnh và trung tâm, sở giao dịch, văn phòng miền trực thuộc
                        If (db Is Nothing Or db.Rows.Count = 0) Then
                            strSQL = String.Format("Select * from ChiNhanh Where id_goc = 0")
                            Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (_db Is Nothing) Then
                                    If (_db.Rows.Count > 0) Then
                                        For i As Integer = 0 To _db.Rows.Count - 1
                                            strSQL = String.Format("Select * from ChiNhanh Where id_goc != 0 and ten_goi Like N'%{0}%' and ma_so Like N'%{1}%' and id_goc = {2}", _Name, _Code, CType(_db.Rows(i)("id").ToString(), Integer))
                                            Using db_cn As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                                If Not (db_cn Is Nothing) Then
                                                    If (db_cn.Rows.Count > 0) Then
                                                        For j As Integer = 0 To db_cn.Rows.Count - 1
                                                            dgv_main.Rows.Add()
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Id").Value = IIf(db_cn.Rows(j)("id").ToString() <> "", db_cn.Rows(j)("id").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_STT").Value = CType((j + _rowCount + 1), String)
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Name").Value = IIf(db_cn.Rows(j)("ten_goi").ToString() <> "", db_cn.Rows(j)("ten_goi").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Code").Value = IIf(db_cn.Rows(j)("ma_so").ToString() <> "", db_cn.Rows(j)("ma_so").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Alias").Value = IIf(db_cn.Rows(j)("ten_vt").ToString() <> "", db_cn.Rows(j)("ten_vt").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Address").Value = IIf(db_cn.Rows(j)("dia_chi").ToString() <> "", db_cn.Rows(j)("dia_chi").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Tel").Value = IIf(db_cn.Rows(j)("dien_thoai").ToString() <> "", db_cn.Rows(j)("dien_thoai").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Email").Value = IIf(db_cn.Rows(j)("email").ToString() <> "", db_cn.Rows(j)("email").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Web").Value = IIf(db_cn.Rows(j)("website").ToString() <> "", db_cn.Rows(j)("website").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Villagers").Value = IIf(db_cn.Rows(j)("solg_xaphuong").ToString() <> "", db_cn.Rows(j)("solg_xaphuong").ToString(), "")
                                                            dgv_main.Rows(j + _rowCount).Cells("cln_Status").Value = IIf(db_cn.Rows(j)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                                        Next
                                                    End If
                                                End If
                                            End Using
                                        Next
                                    End If
                                End If
                            End Using
                        End If

                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    strSQL = String.Format("Select * from ChiNhanh Where id_goc != 0 and ten_goi Like N'%{0}%' and ma_so Like N'%{1}%' and id_goc = {2}", _Name, _Code, CType(db.Rows(i)("id").ToString(), Integer))
                                    Using db_cn As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                        If Not (db_cn Is Nothing) Then
                                            If (db_cn.Rows.Count > 0) Then
                                                For j As Integer = 0 To db_cn.Rows.Count - 1
                                                    dgv_main.Rows.Add()
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Id").Value = IIf(db_cn.Rows(j)("id").ToString() <> "", db_cn.Rows(j)("id").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_STT").Value = CType((j + _rowCount + 1), String)
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Name").Value = IIf(db_cn.Rows(j)("ten_goi").ToString() <> "", db_cn.Rows(j)("ten_goi").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Code").Value = IIf(db_cn.Rows(j)("ma_so").ToString() <> "", db_cn.Rows(j)("ma_so").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Alias").Value = IIf(db_cn.Rows(j)("ten_vt").ToString() <> "", db_cn.Rows(j)("ten_vt").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Address").Value = IIf(db_cn.Rows(j)("dia_chi").ToString() <> "", db_cn.Rows(j)("dia_chi").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Tel").Value = IIf(db_cn.Rows(j)("dien_thoai").ToString() <> "", db_cn.Rows(j)("dien_thoai").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Email").Value = IIf(db_cn.Rows(j)("email").ToString() <> "", db_cn.Rows(j)("email").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Web").Value = IIf(db_cn.Rows(j)("website").ToString() <> "", db_cn.Rows(j)("website").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Villagers").Value = IIf(db_cn.Rows(j)("solg_xaphuong").ToString() <> "", db_cn.Rows(j)("solg_xaphuong").ToString(), "")
                                                    dgv_main.Rows(j + _rowCount).Cells("cln_Status").Value = IIf(db_cn.Rows(j)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                                Next
                                            End If
                                        End If
                                    End Using
                                Next
                            End If
                        End If
                    End Using

                Else                            'Trường hợp không phải node gốc (Tức không phải Click vào Node "NHCSXH")
                    If (_FlagPB = True) Then
                        If (_RecordId > 0) Then
                            strSQL = String.Format("Select * From PhongBan Where Charindex('1',truc_thuoc) > 0 and ten_phong Like N'%{0}%' and ma_so Like N'%{1}%' and id = {2} Order by Ma_so", _Name, _Code, _RecordId)
                        Else
                            strSQL = String.Format("Select * From PhongBan Where Charindex('1',truc_thuoc) > 0 and ten_phong Like N'%{0}%' and ma_so Like N'%{1}%' Order by Ma_so", _Name, _Code)
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_phong").ToString() <> "", db.Rows(i)("ten_phong").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                    Next
                                End If
                            End If
                        End Using
                    Else
                        If (_FlagTT = True) Then    'Khi click vào Node ví dụ "TTCNTT" --> rồi tìm kiếm dl
                            strSQL = String.Format("Select * from ChiNhanh Where id_goc != 0 and id = {0} and ten_goi Like N'%{1}%' and ma_so Like N'%{2}%'", _RootId, _Name, _Code)
                        Else
                            If (_RecordId > 0) Then
                                strSQL = String.Format("Select * from ChiNhanh Where id_goc != 0 and id = {0} and ten_goi Like N'%{1}%' and ma_so Like N'%{2}%'", _RecordId, _Name, _Code)
                            Else
                                strSQL = String.Format("Select * from ChiNhanh Where id_goc != 0 and id_goc = {0} and ten_goi Like N'%{1}%' and ma_so Like N'%{2}%'", _RootId, _Name, _Code)
                            End If
                        End If
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Alias").Value = IIf(db.Rows(i)("ten_vt").ToString() <> "", db.Rows(i)("ten_vt").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Address").Value = IIf(db.Rows(i)("dia_chi").ToString() <> "", db.Rows(i)("dia_chi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Tel").Value = IIf(db.Rows(i)("dien_thoai").ToString() <> "", db.Rows(i)("dien_thoai").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Email").Value = IIf(db.Rows(i)("email").ToString() <> "", db.Rows(i)("email").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Web").Value = IIf(db.Rows(i)("website").ToString() <> "", db.Rows(i)("website").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Villagers").Value = IIf(db.Rows(i)("solg_xaphuong").ToString() <> "", db.Rows(i)("solg_xaphuong").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString() <> "", " Sử dụng", " Tạm xoá")
                                    Next
                                End If
                            End If
                        End Using
                    End If
                End If
            Case 3      'TÌM DL - ĐỊA DANH
                If (_RootId <= 0) Then
                    If (_RecordId > 0) Then
                        strSQL = String.Format("Select * from DiaDanh Where id_goc != 0 and id = {0}", _RecordId)
                        If (_Name <> "") Then
                            strSQL += " and ten_goi Like N'%" + _Name + "%'"
                        End If
                        If (_Code <> "") Then
                            strSQL += " and ma_so Like N'%" + _Code + "%'"
                        End If
                    Else
                        strSQL = String.Format("Select * from DiaDanh Where id_goc = 0")
                        If (_Name <> "") Then
                            strSQL += " and ten_goi Like N'%" + _Name + "%'"
                        End If
                        If (_Code <> "") Then
                            strSQL += " and ma_so Like N'%" + _Code + "%'"
                        End If
                    End If
                    If (strSQL = "") Then Return
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                Dim _rows As Int32 = db.Rows.Count - 1
                                For i As Integer = 0 To _rows
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                        lbl_titleresult.Text = "Kết quả tìm kiếm - Số tỉnh (thành phố) tìm thấy: " + db.Rows.Count.ToString()
                    End Using
                Else
                    Try
                        strSQL = String.Format("Select * From DiaDanh Where id_goc != 0 and id_goc = {0}", _RootId)
                        If (_Name <> "") Then
                            strSQL += " and ten_goi Like N'%" + _Name + "%'"
                        End If
                        If (_Code <> "") Then
                            strSQL += " and ma_so Like N'%" + _Code + "%'"
                        End If
                        If (strSQL = "") Then Return
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    Dim _rows As Int32 = db.Rows.Count - 1
                                    For i As Integer = 0 To _rows
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString() = "1", " Sử dụng", " Tạm khoá")
                                    Next
                                End If
                            End If
                            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00DD") Then
                                lbl_titleresult.Text = "Kết quả tìm kiếm - Số quận (huyện) tìm thấy: " + db.Rows.Count.ToString()
                            Else
                                lbl_titleresult.Text = "Kết quả tìm kiếm - Số xã (phường) tìm thấy: " + db.Rows.Count.ToString()
                            End If
                        End Using
                    Catch ex As Exception
                        MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Tìm kiếm địa danh", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                    End Try
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Hàm trả về Id gốc của chi nhánh khi biết Id bản ghi
    ''' </summary>
    ''' <param name="_idRow">Id bản ghi</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetRootId(ByVal _idRow As Integer) As Integer
        Dim _result As Integer = 0
        Dim strSQL As String = String.Format("Select * from ChiNhanh Where id = {0}", _idRow)
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    _result = CType(db.Rows(0)("id_goc").ToString().Trim(), Integer)
                End If
            End If
        End Using
        Return _result
    End Function

    ''' <summary>
    ''' Hàm thực hiện reset controls khi khởi tạo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetControls()
        _RootId = 0
        _RecordId = 0
        _NodeRoot = ""
        _NodeTag = ""
        _FlagTT = False
        _FlagPB = False
        ckb_chonca.Checked = False
        dgv_main.Rows.Clear()
        edt_maso.Text = ""
        edt_tengoi.Text = ""
    End Sub

    ''' <summary>
    ''' Hàm thực hiện reload lại dữ liệu --> Sau khi giao diện cập nhật tắt
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad()
        Try
            ResetControls()
            If (_FlagState = 3) Then
                ckb_tinh_tp.Visible = True
                ckb_tinh_tp.Checked = False
            Else
                ckb_tinh_tp.Visible = False
            End If
            _DanhMuc.FillData_TreeView(tv_main, _FlagState)
            tv_main.CollapseAll()
            'Thực hiện tìm lại Node thực hiện select lại
            Dim strNode As String = obj_update.Node.ToString().Trim()
            If (strNode <> "") Then
                Dim _nodeCurrent As TreeNode = Nothing
                If (strNode <> "") Then
                    _nodeCurrent = Globals.TreeViewFindNode(tv_main.Nodes, strNode)
                End If
                If Not (_nodeCurrent Is Nothing) Then
                    tv_main.SelectedNode = _nodeCurrent
                    _nodeCurrent.Expand()
                End If
            Else
                If (_FlagState = 3) Then
                    ckb_tinh_tp.Checked = True
                    ckb_tinh_tp_CheckedChanged(Nothing, Nothing)
                End If
            End If
            obj_update.Dispose()
        Catch ex As Exception
            MessageBox.Show("Hiện thị lại dữ liệu: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm kiểm tra thông tin quyền đối với Hệ thống Danh mục
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        Select Case _FlagState
            Case 1      '   HỆ THỐNG DANH MỤC CHUNG
                If (_roles.IndexOf(";55;") < 0) Then
                    tv_main.Enabled = False
                    gb_main.Enabled = False
                End If
                If (_roles.IndexOf(";56;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";57;") < 0) Then
                    btn_edit.Enabled = False
                End If
                If (_roles.IndexOf(";58;") < 0) Then
                    btn_delete.Enabled = False
                    ckb_chonca.Enabled = False
                End If
            Case 2      '   CHI NHÁNH
                If (_roles.IndexOf(";59;") < 0) Then
                    tv_main.Enabled = False
                    gb_main.Enabled = False
                End If
                If (_roles.IndexOf(";60;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";61;") < 0) Then
                    btn_edit.Enabled = False
                End If
                If (_roles.IndexOf(";62;") < 0) Then
                    btn_delete.Enabled = False
                    ckb_chonca.Enabled = False
                End If
            Case 3      '   ĐỊA DANH
                If (_roles.IndexOf(";63;") < 0) Then
                    tv_main.Enabled = False
                    gb_main.Enabled = False
                End If
                If (_roles.IndexOf(";64;") < 0) Then
                    btn_add.Enabled = False
                End If
                If (_roles.IndexOf(";65;") < 0) Then
                    btn_edit.Enabled = False
                End If
                If (_roles.IndexOf(";64;") < 0 And _roles.IndexOf(";65;") < 0) Then
                    ckb_tinh_tp.Enabled = False
                End If
                If (_roles.IndexOf(";66;") < 0) Then
                    btn_delete.Enabled = False
                    ckb_chonca.Enabled = False
                End If
        End Select
    End Sub
#End Region

#Region "---> Events main: Các sự kiện chính dùng cho chương trình <---"
    Private Sub DanhMucForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _DanhMuc.Create_Frame(dgv_main, _FlagState)
        ResetControls()
        If (_FlagState = 3) Then
            ckb_tinh_tp.Visible = True
            ckb_tinh_tp.Checked = False
        Else
            ckb_tinh_tp.Visible = False
        End If
        _DanhMuc.FillData_TreeView(tv_main, _FlagState)
        tv_main.CollapseAll()
        Check_Permits()
        'Thực hiện kiểm tra xem TW hay địa phương sử dụng để ẩn/hiện các chức năng cập nhật
        If (DONVI <> gMaDonViTW) Then
            btn_add.Visible = False
            btn_edit.Visible = False
            btn_delete.Visible = False
            ckb_chonca.Visible = False
        End If
    End Sub

    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        ckb_tinh_tp.Checked = False
        ResetControls()
        Dim strSQL As String = ""
        If (tv_main.Nodes.Count > 0) Then
            _NodeTag = tv_main.SelectedNode.Tag.ToString()
            Select Case _FlagState
                Case 1      '   HỆ THỐNG DANH MỤC CHUNG
                    _DanhMuc.Create_Frame(dgv_main, _FlagState)
                    gb_main.Visible = True
                    'Select các nodes parent của treeview
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00") Then
                        'Nếu chọn dl - Hệ thống danh mục chung
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00DM") Then
                            _RootId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                            'Fill dữ liệu các danh mục mà bạn đang chọn
                            Using db As DataTable = _DanhMuc.GetAll(_RootId, _FlagState, 0)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        For i As Integer = 0 To db.Rows.Count - 1
                                            dgv_main.Rows.Add()
                                            dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                            dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                            If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                            Else
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                            End If
                                        Next
                                    End If
                                End If
                            End Using
                        End If
                        'Nếu chọn dl - Phòng ban

                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00PB") Then
                            _RootId = 100
                            Using db As DataTable = _DanhMuc.GetAll(_RootId, _FlagState, 1)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        For i As Integer = 0 To db.Rows.Count - 1
                                            dgv_main.Rows.Add()
                                            dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                            dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                            If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                            Else
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                            End If
                                        Next
                                    End If
                                End If
                            End Using
                        End If
                        'Nếu chọn dl - Quốc gia
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00QG") Then
                            _RootId = 101
                            Using db As DataTable = _DanhMuc.GetAll(_RootId, _FlagState, 2)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        For i As Integer = 0 To db.Rows.Count - 1
                                            dgv_main.Rows.Add()
                                            dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                            dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                            If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                            Else
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                            End If
                                        Next
                                    End If
                                End If
                            End Using
                        End If
                    End If

                    'Select các nodes child của treeview
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "01") Then
                        _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                        'Nếu chọn dl danh mục
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01DM") Then
                            _RootId = CType(tv_main.SelectedNode.Parent.Tag.ToString().Substring(4), Integer)
                            Dim dr As DataRow
                            dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 0)
                            If Not (dr Is Nothing) Then
                                If (dr.Table.Rows.Count > 0) Then
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("id").ToString() <> "", dr("id").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                                    dgv_main.Rows(0).Cells("cln_Name").Value = IIf(dr("ten_goi").ToString() <> "", dr("ten_goi").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("ma_so").ToString() <> "", dr("ma_so").ToString(), "")
                                    If (dr("Status").ToString().Trim() = "1") Then
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                    Else
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                    End If
                                End If
                            End If
                        End If

                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01PB") Then
                            _RootId = 100
                            Dim dr As DataRow
                            dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 1)
                            If Not (dr Is Nothing) Then
                                If (dr.Table.Rows.Count > 0) Then
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("id").ToString() <> "", dr("id").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                                    dgv_main.Rows(0).Cells("cln_Name").Value = IIf(dr("ten_goi").ToString() <> "", dr("ten_goi").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("ma_so").ToString() <> "", dr("ma_so").ToString(), "")
                                    If (dr("Status").ToString().Trim() = "1") Then
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                    Else
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                    End If
                                End If
                            End If
                        End If
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01QG") Then
                            _RootId = 101
                            Dim dr As DataRow
                            dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 2)
                            If Not (dr Is Nothing) Then
                                If (dr.Table.Rows.Count > 0) Then
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("id").ToString() <> "", dr("id").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                                    dgv_main.Rows(0).Cells("cln_Name").Value = IIf(dr("ten_goi").ToString() <> "", dr("ten_goi").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("ma_so").ToString() <> "", dr("ma_so").ToString(), "")
                                    If (dr("Status").ToString().Trim() = "1") Then
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                    Else
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case 2      '   CHI NHÁNH
                    _DanhMuc.Create_Frame(dgv_main, _FlagState)
                    dgv_main.Rows.Clear()
                    'Fill dữ liệu nếu chọn vào node gốc của chi nhánh (Node: "Ngân hàng chính sách xã hội")
                    If (tv_main.SelectedNode.Tag.ToString() = "NHCSXH") Then
                        _RootId = 1     'Mặc định là trực thuộc
                        _NodeRoot = "NHCSXH"
                        'Fill dữ liệu row - Hội sở chính vào lưới dữ liệu trước
                        strSQL = "SELECT a.*  FROM ChiNhanh a WHERE id_goc = 0 And ma_so = '" & gMaDonViTW & "'"
                        strSQL += " UNION "
                        strSQL += "SELECT a.* FROM ChiNhanh a WHERE id_goc != 0 And id_goc = (SELECT Id FROM ChiNhanh WHERE id_goc = 0) Order by id_goc,Ma_so Asc"
                        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Alias").Value = IIf(db.Rows(i)("ten_vt").ToString() <> "", db.Rows(i)("ten_vt").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Address").Value = IIf(db.Rows(i)("dia_chi").ToString() <> "", db.Rows(i)("dia_chi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Tel").Value = IIf(db.Rows(i)("dien_thoai").ToString() <> "", db.Rows(i)("dien_thoai").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Email").Value = IIf(db.Rows(i)("email").ToString() <> "", db.Rows(i)("email").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Web").Value = IIf(db.Rows(i)("website").ToString() <> "", db.Rows(i)("website").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Villagers").Value = IIf(db.Rows(i)("solg_xaphuong").ToString() <> "", db.Rows(i)("solg_xaphuong").ToString(), "")
                                        If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                            dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                        Else
                                            dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                        End If
                                    Next
                                End If
                            End If
                        End Using
                    End If

                    'Khi click vào cào các chi nhánh trực thuộc (gồm cả các trung tâm, văn phòng miền)
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00CN") Then
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 5) = "00CN1") Then
                            _RootId = 0
                            'Remove columns không cần thiết
                            dgv_main.Columns.Remove("cln_Alias")
                            dgv_main.Columns.Remove("cln_Address")
                            dgv_main.Columns.Remove("cln_Tel")
                            dgv_main.Columns.Remove("cln_Email")
                            dgv_main.Columns.Remove("cln_Web")
                            dgv_main.Columns.Remove("cln_Villagers")
                            _FlagPB = True
                            strSQL = String.Format("Select * From PhongBan Where Charindex('1',truc_thuoc) > 0 Order by Ma_so")
                            Using db_pb As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_pb Is Nothing) Then
                                    If (db_pb.Rows.Count > 0) Then
                                        For i As Integer = 0 To db_pb.Rows.Count - 1
                                            dgv_main.Rows.Add()
                                            dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db_pb.Rows(i)("id").ToString() <> "", db_pb.Rows(i)("id").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                            dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db_pb.Rows(i)("ten_phong").ToString() <> "", db_pb.Rows(i)("ten_phong").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db_pb.Rows(i)("ma_so").ToString() <> "", db_pb.Rows(i)("ma_so").ToString(), "")
                                            If (db_pb.Rows(i)("Status").ToString().Trim() = "1") Then
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                            Else
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                            End If
                                        Next
                                    End If
                                End If
                            End Using
                        End If

                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 5) = "00CN2") Then
                            _RootId = CType(tv_main.SelectedNode.Tag.ToString().Substring(5), Integer)
                            Using db As DataTable = _DanhMuc.GetAll(_RootId, _FlagState, 4)
                                If Not (db Is Nothing) Then
                                    If (db.Rows.Count > 0) Then
                                        For i As Integer = 0 To db.Rows.Count - 1
                                            dgv_main.Rows.Add()
                                            dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                            dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Alias").Value = IIf(db.Rows(i)("ten_vt").ToString() <> "", db.Rows(i)("ten_vt").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Address").Value = IIf(db.Rows(i)("dia_chi").ToString() <> "", db.Rows(i)("dia_chi").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Tel").Value = IIf(db.Rows(i)("dien_thoai").ToString() <> "", db.Rows(i)("dien_thoai").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Email").Value = IIf(db.Rows(i)("email").ToString() <> "", db.Rows(i)("email").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Web").Value = IIf(db.Rows(i)("website").ToString() <> "", db.Rows(i)("website").ToString(), "")
                                            dgv_main.Rows(i).Cells("cln_Villagers").Value = IIf(db.Rows(i)("solg_xaphuong").ToString() <> "", db.Rows(i)("solg_xaphuong").ToString(), "")
                                            If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                            Else
                                                dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                            End If
                                        Next
                                    End If
                                Else
                                    'Vào đây khi người dùng click vào node (TTCNTT,TTĐT, Văn phòng miền trực thuộc) - Ở các node này không có node con
                                    'Lấy mã hiệu chi nhánh
                                    Dim _codeNode As String = ""
                                    strSQL = String.Format("Select * From ChiNhanh Where id_goc <> 0 and id = {0}", _RootId)
                                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                        If Not (_db Is Nothing) Then
                                            If (_db.Rows.Count > 0) Then
                                                _codeNode = _db.Rows(0)("ma_so").ToString().Trim()
                                            End If
                                        End If
                                    End Using
                                    If (_codeNode <> "") Then
                                        If (My_CInt(_codeNode.Substring(0, 4), 0) = 1) Then
                                            _FlagTT = True
                                        End If
                                    End If

                                    Dim dr As DataRow
                                    dr = _DanhMuc.GetDataForId(_RootId, _FlagState, 4)
                                    If Not (dr Is Nothing) Then
                                        If (dr.Table.Rows.Count > 0) Then
                                            dgv_main.Rows.Add()
                                            dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("id").ToString() <> "", dr("id").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                                            dgv_main.Rows(0).Cells("cln_Name").Value = IIf(dr("ten_goi").ToString() <> "", dr("ten_goi").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("ma_so").ToString() <> "", dr("ma_so").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Alias").Value = IIf(dr("ten_vt").ToString() <> "", dr("ten_vt").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Address").Value = IIf(dr("dia_chi").ToString() <> "", dr("dia_chi").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Tel").Value = IIf(dr("dien_thoai").ToString() <> "", dr("dien_thoai").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Email").Value = IIf(dr("email").ToString() <> "", dr("email").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Web").Value = IIf(dr("website").ToString() <> "", dr("website").ToString(), "")
                                            dgv_main.Rows(0).Cells("cln_Villagers").Value = ""
                                            If (dr("Status").ToString().Trim() = "1") Then
                                                dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                            Else
                                                dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                            End If
                                        End If
                                    End If
                                End If
                            End Using
                        End If
                    End If

                    'Khi click vào các node con
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01CN") Then
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 5) = "01CN1") Then
                            _RootId = 0
                            _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(5), Integer)
                            dgv_main.Columns.Remove("cln_Alias")
                            dgv_main.Columns.Remove("cln_Address")
                            dgv_main.Columns.Remove("cln_Tel")
                            dgv_main.Columns.Remove("cln_Email")
                            dgv_main.Columns.Remove("cln_Web")
                            dgv_main.Columns.Remove("cln_Villagers")
                            _FlagPB = True
                            strSQL = String.Format("Select * From PhongBan Where Charindex('1',truc_thuoc) > 0 And id = {0} Order by Ma_so", _RecordId)
                            Using db_pb As DataTable = _SqlHelper.SelectDBRows(strSQL)
                                If Not (db_pb Is Nothing) Then
                                    If (db_pb.Rows.Count > 0) Then
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(0).Cells("cln_Id").Value = IIf(db_pb.Rows(0)("id").ToString() <> "", db_pb.Rows(0)("id").ToString(), "")
                                        dgv_main.Rows(0).Cells("cln_STT").Value = CType((1), String)
                                        dgv_main.Rows(0).Cells("cln_Name").Value = IIf(db_pb.Rows(0)("ten_phong").ToString() <> "", db_pb.Rows(0)("ten_phong").ToString(), "")
                                        dgv_main.Rows(0).Cells("cln_Code").Value = IIf(db_pb.Rows(0)("ma_so").ToString() <> "", db_pb.Rows(0)("ma_so").ToString(), "")
                                        If (db_pb.Rows(0)("Status").ToString().Trim() = "1") Then
                                            dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                        Else
                                            dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                        End If
                                    End If
                                End If
                            End Using
                        End If
                        If (tv_main.SelectedNode.Tag.ToString().Substring(0, 5) = "01CN2") Then
                            _RootId = CType(tv_main.SelectedNode.Parent.Tag.ToString().Substring(5), Integer)
                            _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(5), Integer)
                            Dim dr As DataRow
                            dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 4)
                            If Not (dr Is Nothing) Then
                                If (dr.Table.Rows.Count > 0) Then
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("id").ToString() <> "", dr("id").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                                    dgv_main.Rows(0).Cells("cln_Name").Value = IIf(dr("ten_goi").ToString() <> "", dr("ten_goi").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("ma_so").ToString() <> "", dr("ma_so").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Alias").Value = IIf(dr("ten_vt").ToString() <> "", dr("ten_vt").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Address").Value = IIf(dr("dia_chi").ToString() <> "", dr("dia_chi").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Tel").Value = IIf(dr("dien_thoai").ToString() <> "", dr("dien_thoai").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Email").Value = IIf(dr("email").ToString() <> "", dr("email").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Web").Value = IIf(dr("website").ToString() <> "", dr("website").ToString(), "")
                                    dgv_main.Rows(0).Cells("cln_Villagers").Value = IIf(dr("solg_xaphuong").ToString() <> "", dr("solg_xaphuong").ToString(), "")
                                    If (dr("Status").ToString().Trim() = "1") Then
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Sử dụng"
                                    Else
                                        dgv_main.Rows(0).Cells("cln_Status").Value = " Tạm khoá"
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case 3  'ĐỊA DANH
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00") Then
                        _RootId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                        lbl_titleresult.Text = "DANH SÁCH CÁC QUẬN HUYỆN"
                        Using db As DataTable = _DanhMuc.GetAll(_RootId, _FlagState, 4)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                            dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                        Else
                                            dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                        End If
                                    Next
                                End If
                            End If
                            lbl_titleresult.Text = "DANH SÁCH " + CType(db.Rows.Count - 1, String) + " QUẬN HUYỆN THUỘC " + tv_main.SelectedNode.Text.ToString().Trim().ToUpper()
                        End Using
                    End If
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "01") Then
                        lbl_titleresult.Text = "DANH SÁCH CÁC XÃ PHƯỜNG"
                        _RootId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                        Using db As DataTable = _DanhMuc.GetAll(_RootId, _FlagState, 4)
                            If Not (db Is Nothing) Then
                                If (db.Rows.Count > 0) Then
                                    For i As Integer = 0 To db.Rows.Count - 1
                                        dgv_main.Rows.Add()
                                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                                        If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                            dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                                        Else
                                            dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                                        End If
                                    Next
                                End If
                            End If
                            If tv_main.SelectedNode.Nodes.Count > 0 Then
                                lbl_titleresult.Text = "DANH SÁCH " + CType(db.Rows.Count - 1, String) + " XÃ PHƯỜNG THUỘC " + tv_main.SelectedNode.Text.ToString().Trim().ToUpper()
                            Else
                                lbl_titleresult.Text = "DANH CÁC SÁCH PHƯỜNG THUỘC " + tv_main.SelectedNode.Text.ToString().Trim().ToUpper()
                            End If

                        End Using
                    End If
                    If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "02") Then
                        _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                        Dim dr As DataRow
                        dr = _DanhMuc.GetDataForId(_RecordId, _FlagState, 4)
                        If Not (dr Is Nothing) Then
                            If (dr.Table.Rows.Count > 0) Then
                                dgv_main.Rows.Add()
                                dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("id").ToString() <> "", dr("id").ToString(), "")
                                dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                                dgv_main.Rows(0).Cells("cln_Name").Value = IIf(dr("ten_goi").ToString() <> "", dr("ten_goi").ToString(), "")
                                dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("ma_so").ToString() <> "", dr("ma_so").ToString(), "")
                                dgv_main.Rows(0).Cells("cln_Status").Value = IIf(dr("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                            End If
                        End If
                    End If
            End Select
        End If
    End Sub

    Private Sub edt_tengoi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tengoi.TextChanged
        Search_Lists()
    End Sub

    Private Sub edt_maso_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_maso.TextChanged
        Search_Lists()
    End Sub

    Private Sub ckb_chonca_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_chonca.CheckedChanged
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Globals.Check_All_Items(dgv_main, ckb_chonca.Checked)
            Else
                ckb_chonca.Checked = False
                Return
            End If
        Else
            ckb_chonca.Checked = False
            Return
        End If
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    'Sự kiện chọn vào checked tỉnh - thành phố: Load danh sách các tỉnh thành ra lưới dữ liệu
    Private Sub ckb_tinh_tp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_tinh_tp.CheckedChanged
        _RootId = 0
        _RecordId = 0
        ckb_chonca.Checked = False
        dgv_main.Rows.Clear()
        edt_maso.Text = ""
        edt_tengoi.Text = ""
        If (ckb_tinh_tp.Checked = True) Then
            lbl_titleresult.Text = "DANH SÁCH CÁC TỈNH - THÀNH PHỐ"
            tv_main.CollapseAll()
            _NodeTag = ""
            Using db As DataTable = _SqlHelper.SelectDBRows("Select * From DiaDanh Where id_goc = 0")
                Dim _countALL As Integer = 0
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        For i As Integer = 0 To db.Rows.Count - 1
                            dgv_main.Rows.Add()
                            dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("id").ToString() <> "", db.Rows(i)("id").ToString(), "")
                            dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                            dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_goi").ToString() <> "", db.Rows(i)("ten_goi").ToString(), "")
                            dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_so").ToString() <> "", db.Rows(i)("ma_so").ToString(), "")
                            If (db.Rows(i)("Status").ToString().Trim() = "1") Then
                                dgv_main.Rows(i).Cells("cln_Status").Value = " Sử dụng"
                            Else
                                dgv_main.Rows(i).Cells("cln_Status").Value = " Tạm khoá"
                            End If
                            If db.Rows(i)("ma_so").ToString() <> "99900" Then
                                _countALL += 1
                            End If
                        Next
                    End If
                End If
                lbl_titleresult.Text = "DANH SÁCH " + _countALL.ToString() + " TỈNH - THÀNH PHỐ"
            End Using
        End If
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        obj_update = New DMKhacForm()
        obj_update.ValStatus = False    'False: Xác định trường hợp gọi thêm mới dl
        Select Case _FlagState
            Case 1  'HỆ THỐNG DANH MỤC CHUNG
                If (_RootId = 100) Then         'Xét trường hợp cập nhật Phòng ban
                    obj_update.lbl_danhmuc.Text = " PHÒNG BAN"
                    obj_update._FlagChild = 1
                    obj_update.RootId = _RootId
                    obj_update.RecordId = 0
                ElseIf (_RootId = 101) Then     'Xét trường hợp cập nhật Quốc gia
                    obj_update.lbl_danhmuc.Text = " QUỐC GIA"
                    obj_update._FlagChild = 2
                    obj_update.RootId = _RootId
                    obj_update.RecordId = 0
                Else
                    'Xét trường hợp cập nhật Hệ thống Danh mục chung
                    If (_RecordId > 0) Then
                        Dim _ParentId As Integer = 0
                        _ParentId = CType(tv_main.SelectedNode.Parent.Tag.ToString().Substring(4), Integer)
                        obj_update.lbl_danhmuc.Text = tv_main.SelectedNode.Parent.Text.ToString().ToUpper()
                        obj_update._FlagChild = 0
                        obj_update.RootId = _ParentId
                        obj_update.RecordId = _RecordId
                    Else
                        If (_RootId <= 0) Then
                            MessageBox.Show("Bạn hãy chọn danh mục gốc cần thêm mới dữ liệu trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            Return
                        Else
                            obj_update._FlagChild = 0
                            obj_update.lbl_danhmuc.Text = tv_main.SelectedNode.Text.ToString().ToUpper()
                            obj_update.RootId = _RootId
                            obj_update.RecordId = 0
                        End If
                    End If
                End If
                obj_update._FlagState = _FlagState
                obj_update.Node = _NodeTag
                obj_update.Progress_Changed = New DMKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
                obj_update.ShowDialog()
            Case 2  'DANH MỤC - CHI NHÁNH
                If (_FlagPB = True Or _FlagTT = True Or _RootId <= 0) Then
                    Return
                Else
                    If (_NodeRoot = "NHCSXH") Then
                        obj_update.lbl_danhmuc.Text = "CẬP NHẬT CHI NHÁNH TỈNH - THÀNH PHỐ"
                        obj_update.lbl_dm_title.Text = ""
                    Else
                        If (_RecordId > 0) Then
                            obj_update.lbl_dm_title.Text = ""
                            obj_update.lbl_danhmuc.Text = "Cập nhật PGD thuộc - " + tv_main.SelectedNode.Parent.Text.ToString()
                        Else
                            obj_update.lbl_danhmuc.Text = "Cập nhật PGD thuộc - " + tv_main.SelectedNode.Text.ToString()
                            obj_update.lbl_dm_title.Text = ""
                        End If
                    End If
                    obj_update._FlagState = _FlagState
                    obj_update._FlagChild = 4
                    obj_update.RootId = _RootId
                    obj_update.RecordId = 0
                    obj_update.Node = _NodeTag
                    obj_update.Progress_Changed = New DMKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
                    obj_update.ShowDialog()
                End If

            Case 3  'DANH MỤC - ĐỊA DANH
                If (_RecordId > 0) Then
                    Dim _ParentId As Integer = 0
                    _ParentId = CType(tv_main.SelectedNode.Parent.Tag.ToString().Substring(4), Integer)
                    obj_update.lbl_danhmuc.Text = "THÊM MỚI XÃ PHƯỜNG THUỘC " + tv_main.SelectedNode.Parent.Text.ToString().ToUpper()
                    obj_update.lbl_dm_title.Text = ""
                    obj_update.RootId = _ParentId
                    obj_update.RecordId = 0
                Else
                    If (_RootId = 0) Then
                        obj_update.lbl_danhmuc.Text = "THÊM MỚI ĐỊA DANH TỈNH - THÀNH PHỐ"
                        obj_update.lbl_dm_title.Text = ""
                        obj_update.RootId = 0
                    Else
                        If tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00" Then
                            obj_update.lbl_danhmuc.Text = "THÊM MỚI QUẬN HUYỆN THUỘC " + tv_main.SelectedNode.Text.ToString().ToUpper()
                        Else
                            obj_update.lbl_danhmuc.Text = "THÊM MỚI XÃ PHƯỜNG THUỘC " + tv_main.SelectedNode.Text.ToString().ToUpper()
                        End If
                        obj_update.lbl_dm_title.Text = ""
                        obj_update.RootId = _RootId
                    End If
                End If
                obj_update._FlagChild = 4
                obj_update._FlagState = _FlagState
                obj_update.RecordId = 0
                obj_update.Node = _NodeTag
                obj_update.Progress_Changed = New DMKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
                obj_update.ShowDialog()
        End Select
    End Sub

    Private Sub btn_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit.Click
        obj_update = New DMKhacForm()
        obj_update.ValStatus = True    'True: Xác định trường hợp gọi sửa đổi dl
        Select Case _FlagState
            Case 1  'HỆ THỐNG DANH MỤC CHUNG
                If (dgv_main.Rows.Count <= 0) Then
                    MessageBox.Show("Bạn chưa chọn dữ liệu danh mục cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Return
                End If
                If (_RootId = 100) Then         'Xét trường hợp cập nhật Phòng ban
                    obj_update._FlagState = _FlagState
                    obj_update.lbl_danhmuc.Text = " PHÒNG BAN"
                    obj_update._FlagChild = 1
                ElseIf (_RootId = 101) Then     'Xét trường hợp cập nhật Quốc gia
                    obj_update._FlagState = _FlagState
                    obj_update.lbl_danhmuc.Text = " QUỐC GIA"
                    obj_update._FlagChild = 2
                Else                            'Xét trường hợp cập nhật Hệ thống Danh mục chung
                    If (_RootId <= 0) Then
                        MessageBox.Show("Bạn hãy chọn danh mục gốc cần sửa đổi dữ liệu trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Return
                    Else
                        obj_update._FlagChild = 0
                        obj_update.lbl_danhmuc.Text = tv_main.SelectedNode.Text.ToString().ToUpper()
                        obj_update._FlagState = _FlagState
                    End If
                End If

                obj_update.RootId = _RootId
                obj_update.RecordId = CType(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString(), Integer)
                obj_update.Node = _NodeTag
                obj_update.Progress_Changed = New DMKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
                obj_update.ShowDialog()

            Case 2  'DANH MỤC - CHI NHÁNH
                If (dgv_main.Rows.Count <= 0) Then
                    MessageBox.Show("Bạn chưa chọn dữ liệu chi nhánh cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Return
                End If
                If (_FlagPB = True) Then    'Nếu chọn vào Node Hội sở chính hiện ds các Phòng ban thì bỏ qua <--- ??? 
                    Return
                End If
                Dim _idRow As Integer = 0
                If (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString().Trim() <> "") Then
                    _idRow = CType(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString(), Integer)
                End If
                If (_idRow <= 0) Then Return

                Dim strSQL As String = String.Format("Select * from ChiNhanh Where id = {0}", _idRow)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            _RootId = CType(db.Rows(0)("id_goc").ToString().Trim(), Integer)
                        End If
                    End If
                End Using
                If (_RootId = 0 Or _RootId = 1) Then
                    obj_update.lbl_danhmuc.Text = "CẬP NHẬT CHI NHÁNH TỈNH - THÀNH PHỐ"
                Else
                    If (_RecordId > 0) Then
                        obj_update.lbl_danhmuc.Text = "Cập nhật PGD - " + tv_main.SelectedNode.Parent.Text.ToString()
                    Else
                        obj_update.lbl_danhmuc.Text = "Cập nhật PGD - " + tv_main.SelectedNode.Text.ToString()
                    End If
                End If

                obj_update.RootId = _RootId
                obj_update.lbl_dm_title.Text = ""
                obj_update._FlagChild = 4
                obj_update._FlagState = _FlagState
                obj_update.RecordId = CType(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString(), Integer)
                obj_update.Node = _NodeTag
                obj_update.Progress_Changed = New DMKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
                obj_update.ShowDialog()

            Case 3  'DANH MỤC - ĐỊA DANH
                If (dgv_main.Rows.Count <= 0) Then
                    MessageBox.Show("Bạn chưa chọn dữ liệu địa danh cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Return
                End If
                If (_RecordId > 0) Then
                    obj_update.lbl_danhmuc.Text = "CẬP NHẬT XÃ PHƯỜNG THUỘC " + tv_main.SelectedNode.Parent.Text.ToString().ToUpper()
                    obj_update.RootId = _RootId
                Else
                    If (_RootId = 0) Then
                        obj_update.lbl_danhmuc.Text = "CẬP NHẬT TỈNH - THÀNH PHỐ"
                        obj_update.RootId = 0
                    Else
                        If tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00" Then
                            obj_update.lbl_danhmuc.Text = "CẬP NHẬT QUẬN HUYỆN THUỘC " + tv_main.SelectedNode.Text.ToString().ToUpper()
                        Else
                            obj_update.lbl_danhmuc.Text = "THÊM MỚI XÃ PHƯỜNG THUỘC " + tv_main.SelectedNode.Text.ToString().ToUpper()
                        End If
                        obj_update.RootId = _RootId
                    End If
                End If
                obj_update.lbl_dm_title.Text = ""
                obj_update._FlagChild = 4
                obj_update._FlagState = _FlagState
                obj_update.RecordId = CType(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString(), Integer)
                obj_update.Node = _NodeTag
                obj_update.Progress_Changed = New DMKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
                obj_update.ShowDialog()
        End Select
    End Sub
#End Region

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click

    End Sub
End Class
