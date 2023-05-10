Public Class DMLuongForm

#Region "---> Defined parametter and properties <---"
    Dim _DanhMuc As clsHT_DanhMuc = New clsHT_DanhMuc

    Private _SqlHelper As DBAccess
    Private vNFInfo As System.Globalization.NumberFormatInfo

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        _SqlHelper = New DBAccess

        vNFInfo = New System.Globalization.NumberFormatInfo()
        vNFInfo.NumberDecimalDigits = 2
        vNFInfo.NumberGroupSeparator = " "
    End Sub

    Private _NodeTag As String = ""
    Private _RecordId As Integer = 0

    ''' <summary>
    ''' Chỉ số xác định Danh mục Lương đang thao tác. Với quy định sau
    '''                              _State = 4: Nghị định lương
    '''                              _State = 5: Bảng lương
    '''                              _State = 6: Ngạch lương
    '''                              _State = 7: Bậc lương
    '''                              _State = 8: Lương cơ bản
    '''                              _State = 9: Mức phụ cấp
    ''' </summary>
    ''' <remarks></remarks>
    Private _Status As Byte = 0
    Private strSQL As String = ""
    Private obj_update As DMLKhacForm
#End Region

#Region "---> Functions main: Các hàm chính <---"
    ''' <summary>
    ''' Hàm thực hiện reset controls khi khởi tạo
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetControls()
        _Status = 0
        _RecordId = 0
        _NodeTag = ""
        ckb_chonca.Checked = False
        edt_maso.Text = ""
        edt_tengoi.Text = ""
        dgv_main.Rows.Clear()
    End Sub

    ''' <summary>
    ''' Hàm thực hiện tạo nhẵn - Phù hợp với từng loại danh mục lương
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Create_Labels()
        Select Case _Status
            Case 4      'Nghị định lương
                lbl_name.Text = "(*) Tên nghị định"
                lbl_code.Text = "(*) Số nghị định"
            Case 5      'Bảng lương
                lbl_name.Text = "(*) Mô tả chi tiết"
                lbl_code.Text = "(*) Bảng lương"
            Case 6      'Ngạch lương
                lbl_name.Text = "(*) Mô tả chi tiết"
                lbl_code.Text = "(*) Ngạch lương"
            Case 7      'Bậc lương
                lbl_name.Text = "(*) Mô tả chi tiết"
                lbl_code.Text = "(*) Bậc lương"
            Case Else
                lbl_name.Text = "(*) Tên gọi"
                lbl_code.Text = "(*) Mã số"
        End Select
    End Sub

    ''' <summary>
    ''' Hàm thực hiện tìm kiếm thông tin - Hệ thống danh mục Lương
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Search()
        Try
            Dim _Name = Globals.Find_Replace(edt_tengoi.Text.ToString().Trim())
            Dim _Code = Globals.Find_Replace(edt_maso.Text.ToString().Trim())
            dgv_main.Rows.Clear()
            Select Case _Status
                Case 4          ' Nghị định lương
                    strSQL = "Select * From NghiDinhLuong"
                    If (_Name <> "" And _Code <> "") Then
                        strSQL += String.Format(" Where SoND Like N'%{0}%' and TenND Like N'%{1}%'", _Code, _Name)
                    Else
                        If (_Name <> "") Then
                            strSQL += String.Format(" Where TenND Like N'%{0}%'", _Name)
                        End If
                        If (_Code <> "") Then
                            strSQL += String.Format(" Where SoND Like N'%{0}%'", _Code)
                        End If
                    End If
                    strSQL += " Order by TenND asc"
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdNDLuong").ToString() <> "", db.Rows(i)("IdNDLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("MaND").ToString() <> "", db.Rows(i)("MaND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_SoNgd").Value = IIf(db.Rows(i)("SoND").ToString() <> "", db.Rows(i)("SoND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_TenNd").Value = IIf(db.Rows(i)("TenND").ToString() <> "", db.Rows(i)("TenND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_LoaiNd").Value = IIf(db.Rows(i)("LoaiND").ToString() <> "", db.Rows(i)("LoaiND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using

                Case 5          ' Bảng lương
                    strSQL = "Select a.*, b.TenND as Nd_Luong From BangLuong a, NghiDinhLuong b Where b.IdNDLuong = a.IdND_Luong"
                    If (_Name <> "") Then
                        strSQL += String.Format(" and a.MoTa Like N'%{0}%'", _Name)
                    End If
                    If (_Code <> "") Then
                        strSQL += String.Format(" and a.BangLuong Like N'%{0}%'", _Code)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdBangLuong").ToString() <> "", db.Rows(i)("IdBangLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_NgdLuong").Value = IIf(db.Rows(i)("Nd_Luong").ToString() <> "", db.Rows(i)("Nd_Luong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_BangLuong").Value = IIf(db.Rows(i)("BangLuong").ToString() <> "", db.Rows(i)("BangLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Mota").Value = IIf(db.Rows(i)("MoTa").ToString() <> "", db.Rows(i)("MoTa").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using

                Case 6          ' Ngạch lương
                    strSQL = "Select a.*,b.BangLuong From NgachLuong a, BangLuong b Where a.IdBangLuong = b.IdBangLuong"
                    If (_Name <> "") Then
                        strSQL += String.Format(" and a.Mota Like N'%{0}%'", _Name)
                    End If
                    If (_Code <> "") Then
                        strSQL += String.Format(" and a.NgachLuong Like N'%{0}%'", _Code)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdNgachLuong").ToString() <> "", db.Rows(i)("IdNgachLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_BangLuong").Value = IIf(db.Rows(i)("BangLuong").ToString() <> "", db.Rows(i)("BangLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_NgLuong").Value = IIf(db.Rows(i)("NgachLuong").ToString() <> "", db.Rows(i)("NgachLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Phanloai").Value = IIf(db.Rows(i)("Loai").ToString() <> "", db.Rows(i)("Loai").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Mota").Value = IIf(db.Rows(i)("MoTa").ToString() <> "", db.Rows(i)("MoTa").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using

                Case 7          ' Bậc lương

                    strSQL = "Select a.*,b.NgachLuong From BacLuong a, NgachLuong b Where a.IdNgachLuong = b.IdNgachLuong"
                    If (_Name <> "") Then
                        strSQL += String.Format(" and a.Mota Like N'%{0}%'", _Name)
                    End If
                    If (_Code <> "") Then
                        strSQL += String.Format(" and a.BacLuong Like N'%{0}%'", _Code)
                    End If
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdBacLuong").ToString() <> "", db.Rows(i)("IdBacLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_NgLuong").Value = IIf(db.Rows(i)("NgachLuong").ToString() <> "", db.Rows(i)("NgachLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_BacLuong").Value = IIf(db.Rows(i)("BacLuong").ToString() <> "", db.Rows(i)("BacLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_HeSo").Value = IIf(db.Rows(i)("Heso").ToString() <> "", db.Rows(i)("Heso").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Mota").Value = IIf(db.Rows(i)("MoTa").ToString() <> "", db.Rows(i)("MoTa").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using

            End Select
        Catch ex As Exception
            MessageBox.Show("Tìm kiếm thông tin: " + ex.Message.ToString(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện load lại dữ liệu sau khi cập nhật song
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReLoad()
        Try
            ResetControls()
            _DanhMuc.Fill_Data(tv_main)
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
                    '_nodeCurrent.Expand()
                End If
            End If
            obj_update.Dispose()
        Catch ex As Exception
            MessageBox.Show("Hiện thị lại dữ liệu: " + ex.Message.ToString(), "Error ...", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra quyền thành viên - Cho phép thao tác theo sự phân quyền
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        If (_roles.IndexOf(";67;") < 0) Then
            tv_main.Enabled = False
            gb_main.Enabled = False
        End If
        If (_roles.IndexOf(";68;") < 0) Then
            btn_add.Enabled = False
        End If
        If (_roles.IndexOf(";69;") < 0) Then
            btn_edit.Enabled = False
        End If
        If (_roles.IndexOf(";70;") < 0) Then
            btn_delete.Enabled = False
            ckb_chonca.Visible = False
            dgv_main.Columns("cln_Choice").Visible = False
        End If
    End Sub
#End Region

#Region "---> Events main: Các sự kiện chính <---"
    Private Sub DMLuongForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _DanhMuc.Fill_Data(tv_main)
        _DanhMuc.Create_Frame(dgv_main, 4)
        ResetControls()
        'Xét quyền thành viên thao tác chương trình
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
        ResetControls()
        If (tv_main.Nodes.Count > 0) Then
            _NodeTag = tv_main.SelectedNode.Tag.ToString()

            'A - Parent Node: Select các nodes parent của treeview
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00") Then
                '1 - Nếu chọn dữ liệu - Nghị định lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00ND") Then
                    _Status = 4         'Với giá trị này chỉ số xác định - Nghị định lương
                    _DanhMuc.Create_Frame(dgv_main, 4)
                    'Kiểm tra nếu không có quyền xoá thì bỏ cột chọn xoá trên lưới dữ liệu
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(4)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdNDLuong").ToString() <> "", db.Rows(i)("IdNDLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("MaND").ToString() <> "", db.Rows(i)("MaND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_SoNgd").Value = IIf(db.Rows(i)("SoND").ToString() <> "", db.Rows(i)("SoND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_TenNd").Value = IIf(db.Rows(i)("TenND").ToString() <> "", db.Rows(i)("TenND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_LoaiNd").Value = IIf(db.Rows(i)("LoaiND").ToString() <> "", db.Rows(i)("LoaiND").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("GhiChu").ToString() <> "", db.Rows(i)("GhiChu").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If

                '2 - Nếu chọn dữ liệu - Bảng lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00BL") Then
                    _Status = 5         'Với giá trị này chỉ số xác định - Bảng lương
                    _DanhMuc.Create_Frame(dgv_main, 5)
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(5)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdBangLuong").ToString() <> "", db.Rows(i)("IdBangLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_NgdLuong").Value = IIf(db.Rows(i)("Nd_Luong").ToString() <> "", db.Rows(i)("Nd_Luong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_BangLuong").Value = IIf(db.Rows(i)("BangLuong").ToString() <> "", db.Rows(i)("BangLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Mota").Value = IIf(db.Rows(i)("MoTa").ToString() <> "", db.Rows(i)("MoTa").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If

                '3 - Nếu chọn dữ liệu - Ngạch lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00NL") Then
                    _Status = 6         'Với giá trị này chỉ số xác định - Ngạch lương
                    _DanhMuc.Create_Frame(dgv_main, 6)
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(6)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdNgachLuong").ToString() <> "", db.Rows(i)("IdNgachLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_BangLuong").Value = IIf(db.Rows(i)("BangLuong").ToString() <> "", db.Rows(i)("BangLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_NgLuong").Value = IIf(db.Rows(i)("NgachLuong").ToString() <> "", db.Rows(i)("NgachLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Phanloai").Value = IIf(db.Rows(i)("Loai").ToString() <> "", db.Rows(i)("Loai").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Mota").Value = IIf(db.Rows(i)("MoTa").ToString() <> "", db.Rows(i)("MoTa").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If

                '4 - Nếu chọn dữ liệu - Bậc lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00BA") Then
                    _Status = 7         'Với giá trị này chỉ số xác định - Bậc lương
                    _DanhMuc.Create_Frame(dgv_main, 7)
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(7)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdBacLuong").ToString() <> "", db.Rows(i)("IdBacLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_NgLuong").Value = IIf(db.Rows(i)("NgachLuong").ToString() <> "", db.Rows(i)("NgachLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_BacLuong").Value = IIf(db.Rows(i)("BacLuong").ToString() <> "", db.Rows(i)("BacLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_HeSo").Value = IIf(db.Rows(i)("Heso").ToString() <> "", db.Rows(i)("Heso").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Mota").Value = IIf(db.Rows(i)("MoTa").ToString() <> "", db.Rows(i)("MoTa").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If

                '5 - Nếu chọn dữ liệu - Lương cơ bản
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00CB") Then
                    _Status = 8         'Với giá trị này chỉ số xác định - Lương cơ bản
                    _DanhMuc.Create_Frame(dgv_main, 8)
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(8)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdTienLuong").ToString() <> "", db.Rows(i)("IdTienLuong").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_LuongCb").Value = IIf(db.Rows(i)("LuongCoBan").ToString().Trim() <> "", Double.Parse(db.Rows(i)("LuongCoBan").ToString().Trim(), Globals.cultureNum).ToString("N", vNFInfo), "0")
                                    dgv_main.Rows(i).Cells("cln_HsNganh").Value = IIf(db.Rows(i)("HeSoNganh").ToString() <> "", CType(db.Rows(i)("HeSoNganh"), Double).ToString("N2"), "")
                                    If db.Rows(i)("NgayHuong").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_NgApdung").Value = CType(db.Rows(i)("NgayHuong").ToString(), DateTime).ToString("dd-MM-yyyy")
                                    Else
                                        dgv_main.Rows(i).Cells("cln_NgApdung").Value = ""
                                    End If
                                    dgv_main.Rows(i).Cells("cln_SoQd").Value = IIf(db.Rows(i)("SoQD").ToString() <> "", db.Rows(i)("SoQD").ToString(), "")
                                    If db.Rows(i)("NgayQD").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_NgayQd").Value = CType(db.Rows(i)("NgayQD").ToString(), DateTime).ToString("dd-MM-yyyy")
                                    Else
                                        dgv_main.Rows(i).Cells("cln_NgayQd").Value = ""
                                    End If
                                    dgv_main.Rows(i).Cells("cln_NguoiQd").Value = IIf(db.Rows(i)("NguoiQD").ToString() <> "", db.Rows(i)("NguoiQD").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Chucvu").Value = IIf(db.Rows(i)("Chucvu").ToString() <> "", db.Rows(i)("Chucvu").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Ghichu").Value = IIf(db.Rows(i)("Ghichu").ToString() <> "", db.Rows(i)("Ghichu").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If

                '6 - Nếu chọn dữ liệu - Mức phụ cấp
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00PC") Then
                    _Status = 9         'Với giá trị này chỉ số xác định - Mức phụ cấp
                    _DanhMuc.Create_Frame(dgv_main, 9)
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(9)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdMuc_PhC").ToString() <> "", db.Rows(i)("IdMuc_PhC").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_LoaiPc").Value = IIf(db.Rows(i)("LoaiPhCap").ToString() <> "", db.Rows(i)("LoaiPhCap").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_MucPc").Value = IIf(db.Rows(i)("Muc_PhC").ToString() <> "", CType(db.Rows(i)("Muc_PhC"), Double).ToString("N2"), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If

                '7 - Nếu chọn dữ liệu - Tham số Lương hệ thống
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "00TS") Then
                    _Status = 10         'Với giá trị này chỉ số xác định - Tham số lương hệ thống
                    _DanhMuc.Create_Frame(dgv_main, 11)
                    dgv_main.Rows.Clear()
                    Using db As DataTable = _DanhMuc.GetAll(10)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                For i As Integer = 0 To db.Rows.Count - 1
                                    dgv_main.Rows.Add()
                                    dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("IdSalaryVar").ToString() <> "", db.Rows(i)("IdSalaryVar").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_STT").Value = CType((i + 1), String)
                                    dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("NameVar").ToString() <> "", db.Rows(i)("NameVar").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Value").Value = IIf(db.Rows(i)("Value").ToString() <> "", db.Rows(i)("Value").ToString().Trim().Replace("#", "; "), "")
                                    dgv_main.Rows(i).Cells("cln_Rate").Value = IIf(db.Rows(i)("Rate").ToString() <> "", db.Rows(i)("Rate").ToString().Trim().Replace("#", "; "), "")
                                    If db.Rows(i)("DateApply").ToString().Trim() <> "" Then
                                        dgv_main.Rows(i).Cells("cln_DateApply").Value = CType(db.Rows(i)("DateApply").ToString(), DateTime).ToString("dd-MM-yyyy")
                                    Else
                                        dgv_main.Rows(i).Cells("cln_DateApply").Value = ""
                                    End If
                                    dgv_main.Rows(i).Cells("cln_Discript").Value = IIf(db.Rows(i)("Descript").ToString() <> "", db.Rows(i)("Descript").ToString(), "")
                                    dgv_main.Rows(i).Cells("cln_Status").Value = IIf(db.Rows(i)("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                                Next
                            End If
                        End If
                    End Using
                End If
            End If

            'B - Child Node: Select các nodes parent của treeview
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "01") Then
                '1 - Nếu chọn dữ liệu - Nghị định lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01ND") Then
                    _Status = 4         'Với giá trị này chỉ số xác định - Nghị định lương
                    _DanhMuc.Create_Frame(dgv_main, 4)
                    dgv_main.Rows.Clear()
                    _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                    Dim dr As DataRow
                    dr = _DanhMuc.GetDataForId(_RecordId, 4)
                    If Not (dr Is Nothing) Then
                        If (dr.Table.Rows.Count > 0) Then
                            dgv_main.Rows.Add()
                            dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("IdNDLuong").ToString() <> "", dr("IdNDLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_Code").Value = IIf(dr("MaND").ToString() <> "", dr("MaND").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_SoNgd").Value = IIf(dr("SoND").ToString() <> "", dr("SoND").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_TenNd").Value = IIf(dr("TenND").ToString() <> "", dr("TenND").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_LoaiNd").Value = IIf(dr("LoaiND").ToString() <> "", dr("LoaiND").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_Ghichu").Value = IIf(dr("GhiChu").ToString() <> "", dr("GhiChu").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                            dgv_main.Rows(0).Cells("cln_Status").Value = IIf(dr("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                        End If
                    End If
                End If

                '2 - Nếu chọn dữ liệu - Bảng lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01BL") Then
                    _Status = 5         'Với giá trị này chỉ số xác định - Bảng lương
                    _DanhMuc.Create_Frame(dgv_main, 5)
                    dgv_main.Rows.Clear()
                    _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                    Dim dr As DataRow
                    dr = _DanhMuc.GetDataForId(_RecordId, 5)
                    If Not (dr Is Nothing) Then
                        If (dr.Table.Rows.Count > 0) Then
                            dgv_main.Rows.Add()
                            dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("IdBangLuong").ToString() <> "", dr("IdBangLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                            dgv_main.Rows(0).Cells("cln_NgdLuong").Value = IIf(dr("Nd_Luong").ToString() <> "", dr("Nd_Luong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_BangLuong").Value = IIf(dr("BangLuong").ToString() <> "", dr("BangLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_Mota").Value = IIf(dr("MoTa").ToString() <> "", dr("MoTa").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                            dgv_main.Rows(0).Cells("cln_Status").Value = IIf(dr("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                        End If
                    End If
                End If

                '3 - Nếu chọn dữ liệu - Ngạch lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01NL") Then
                    _Status = 6         'Với giá trị này chỉ số xác định - Ngạch lương
                    _DanhMuc.Create_Frame(dgv_main, 6)
                    dgv_main.Rows.Clear()
                    _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                    Dim dr As DataRow
                    dr = _DanhMuc.GetDataForId(_RecordId, 6)
                    If Not (dr Is Nothing) Then
                        If (dr.Table.Rows.Count > 0) Then
                            dgv_main.Rows.Add()
                            dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("IdNgachLuong").ToString() <> "", dr("IdNgachLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                            dgv_main.Rows(0).Cells("cln_BangLuong").Value = IIf(dr("BangLuong").ToString() <> "", dr("BangLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_NgLuong").Value = IIf(dr("NgachLuong").ToString() <> "", dr("NgachLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_Phanloai").Value = IIf(dr("Loai").ToString() <> "", dr("Loai").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_Mota").Value = IIf(dr("MoTa").ToString() <> "", dr("MoTa").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                            dgv_main.Rows(0).Cells("cln_Status").Value = IIf(dr("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                        End If
                    End If
                End If

                '4 - Nếu chọn dữ liệu - Bậc lương
                If (tv_main.SelectedNode.Tag.ToString().Substring(0, 4) = "01BA") Then
                    _Status = 7         'Với giá trị này chỉ số xác định - Bậc lương
                    _DanhMuc.Create_Frame(dgv_main, 7)
                    dgv_main.Rows.Clear()
                    _RecordId = CType(tv_main.SelectedNode.Tag.ToString().Substring(4), Integer)
                    Dim dr As DataRow
                    dr = _DanhMuc.GetDataForId(_RecordId, 7)
                    If Not (dr Is Nothing) Then
                        If (dr.Table.Rows.Count > 0) Then
                            dgv_main.Rows.Add()
                            dgv_main.Rows(0).Cells("cln_Id").Value = IIf(dr("IdBacLuong").ToString() <> "", dr("IdBacLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_STT").Value = CType(1, String)
                            dgv_main.Rows(0).Cells("cln_NgLuong").Value = IIf(dr("NgachLuong").ToString() <> "", dr("NgachLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_BacLuong").Value = IIf(dr("BacLuong").ToString() <> "", dr("BacLuong").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_HeSo").Value = IIf(dr("Heso").ToString() <> "", dr("Heso").ToString(), "")
                            dgv_main.Rows(0).Cells("cln_Mota").Value = IIf(dr("MoTa").ToString() <> "", dr("MoTa").ToString().Replace(Chr(13) & Chr(10), " - "), "")
                            dgv_main.Rows(0).Cells("cln_Status").Value = IIf(dr("Status").ToString().Trim() = "1", " Sử dụng", " Tạm khoá")
                        End If
                    End If
                End If
            End If
        End If
        Create_Labels()
        If (_Status = 8 Or _Status = 9 Or _Status = 10) Then
            gb_main.Visible = False
            lbl_note.Visible = False
        Else
            gb_main.Visible = True
            lbl_note.Visible = True
        End If
    End Sub

    Private Sub edt_tengoi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_tengoi.TextChanged
        Search()
    End Sub

    Private Sub edt_maso_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edt_maso.TextChanged
        Search()
    End Sub

    Private Sub edt_tengoi_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_tengoi.KeyDown
        If (e.KeyCode = Keys.Tab) Then
            ActiveControl = edt_maso
        End If
    End Sub

    Private Sub edt_maso_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_maso.KeyUp
        If (e.KeyCode = Keys.Up) Then
            ActiveControl = edt_tengoi
        End If
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        If (_Status <= 0 Or _Status > 10) Then
            MessageBox.Show("Bạn hãy chọn danh mục lương cần thêm mới dữ liệu trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        obj_update = New DMLKhacForm()
        obj_update._Status = _Status
        obj_update.RecordId = 0
        obj_update.Node = _NodeTag
        obj_update.Progress_Changed = New DMLKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
        obj_update.ShowDialog()
    End Sub

    Private Sub btn_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_edit.Click
        If (_Status <= 0 Or _Status > 10) Then
            MessageBox.Show("Bạn hãy chọn danh mục lương cần sửa đổi dữ liệu trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        If (dgv_main.Rows.Count <= 0) Then
            MessageBox.Show("Bạn chưa chọn dữ liệu danh mục lương cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If
        obj_update = New DMLKhacForm()
        obj_update._Status = _Status
        obj_update.RecordId = CType(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString(), Integer)
        obj_update.Node = _NodeTag
        obj_update.Progress_Changed = New DMLKhacForm.ProgressChangedEventHandler(AddressOf ReLoad)
        obj_update.ShowDialog()
    End Sub

    Private Sub btn_quayra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub ckb_chonca_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_chonca.CheckedChanged
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
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
#End Region

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click

    End Sub
End Class