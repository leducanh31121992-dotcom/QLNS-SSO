Public Class frmHS_DoanTheQT

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private arr_Chucvu As ArrayList = New ArrayList()
    Private _Globals As Globals = New Globals
    Private _HS_DangDT As clsHS_DangDT = New clsHS_DangDT()
    Private _RowId As String = ""   'Biến lưu Id quá trình sinh hoạt đoàn thể - để lấy id khi sửa đổi
    Private _SqlHelper As DBAccess = New DBAccess()
    'Chuỗi chỉ số xác định Hồ sơ đoàn thể được truyền khi gọi quá trình sinh hoạt
    Private _DocumentId As String
    Public Property DocumentId() As String
        Get
            Return _DocumentId
        End Get
        Set(ByVal value As String)
            _DocumentId = value
        End Set
    End Property
    'Biến lưu chỉ số xác định Hồ sơ: 1-Đảng, 2-Đoàn, 3-Công đoàn
    Public FlagDocument As Byte = 0
    Private strSQL As String = ""
    'Khai báo các thông tin add CheckBox vào cột tiêu đề chọn cả các items Check trên lưới
    Private ckb_ChoiceAll As CheckBox = Nothing   'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0          'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0   'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False  'Cờ báo việc Checkall
#End Region

#Region "---> Functions: Các hàm chính <---"
    ''' <summary>
    ''' Hàm thực hiện reset all controls về trạng thái mặc định
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls()
        _RowId = ""
        dtpk_qt_tungay.Text = DateTime.Now.ToShortDateString()
        dtpk_qt_denngay.Text = DateTime.Now.ToShortDateString()
        cb_qt_chucvu.SelectedIndex = 0
        edt_qt_doanthe.Text = ""
        edt_qt_ghichu.Text = ""
        If (FlagDocument = 1) Then ckb_kiemnhiem.Checked = False
    End Sub

    ''' <summary>
    ''' Hàm fill dữ liệu danh sách quá trình sinh hoạt đoàn thể của cán bộ
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Fill_Data()
        ckb_ChoiceAll.Checked = False
        dgv_main.Rows.Clear()
        Using db As DataTable = _HS_DangDT.GetAll_Process(_DocumentId, FlagDocument)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    Select Case FlagDocument
                        Case 1          'Quá trình sinh hoạt Đảng 
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_main.Rows.Add()
                                dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("IdHSDang").ToString() <> "", db.Rows(i)("IdHSDang").ToString(), "")
                                If db.Rows(i)("TuNgay").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_Tungay").Value = IIf(CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                Else
                                    dgv_main.Rows(i).Cells("cln_Tungay").Value = ""
                                End If
                                If db.Rows(i)("DenNgay").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_Denngay").Value = IIf(CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                Else
                                    dgv_main.Rows(i).Cells("cln_Denngay").Value = ""
                                End If
                                dgv_main.Rows(i).Cells("cln_Chucvu").Value = db.Rows(i)("ChucVu").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_Chibo").Value = db.Rows(i)("ChiBo").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                            Next
                        Case 2          'Quá trình sinh hoạt Đoàn
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_main.Rows.Add()
                                dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("IdHSDoan").ToString() <> "", db.Rows(i)("IdHSDoan").ToString(), "")
                                If db.Rows(i)("TuNgay").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_TuNgay").Value = IIf(CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                Else
                                    dgv_main.Rows(i).Cells("cln_TuNgay").Value = ""
                                End If
                                If db.Rows(i)("DenNgay").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_DenNgay").Value = IIf(CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                Else
                                    dgv_main.Rows(i).Cells("cln_DenNgay").Value = ""
                                End If
                                dgv_main.Rows(i).Cells("cln_ChucVu").Value = db.Rows(i)("ChucVu").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_ChiDoan").Value = db.Rows(i)("ChiDoan").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                            Next
                        Case 3          'Quá trình sinh hoạt Công đoàn
                            For i As Integer = 0 To db.Rows.Count - 1
                                dgv_main.Rows.Add()
                                dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("IdHSCDoan").ToString() <> "", db.Rows(i)("IdHSCDoan").ToString(), "")
                                If db.Rows(i)("TuNgay").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_TuNgay").Value = IIf(CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("TuNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                Else
                                    dgv_main.Rows(i).Cells("cln_TuNgay").Value = ""
                                End If
                                If db.Rows(i)("DenNgay").ToString().Trim() <> "" Then
                                    dgv_main.Rows(i).Cells("cln_DenNgay").Value = IIf(CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy") <> "01-01-1900", CType(db.Rows(i)("DenNgay").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                                Else
                                    dgv_main.Rows(i).Cells("cln_DenNgay").Value = ""
                                End If
                                dgv_main.Rows(i).Cells("cln_ChucVu").Value = db.Rows(i)("ChucVu").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_CoQuan").Value = db.Rows(i)("CoQuan").ToString().Trim()
                                dgv_main.Rows(i).Cells("cln_Ghichu").Value = db.Rows(i)("GhiChu").ToString().Trim()
                            Next
                    End Select
                End If
            End If
        End Using
        TotalCheckBoxes = dgv_main.RowCount
        TotalCheckedCheckBoxes = 0
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra tính hợp lệ của việc cập nhật dữ liệu
    ''' </summary>
    ''' <returns>True: Is success. Reverse -> False</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Select Case FlagDocument
            Case 1  'Quá trình sinh hoạt Đảng
                If (cb_qt_chucvu.SelectedIndex <= 0 And cb_qt_chucvu.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn chức vụ đảng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_qt_chucvu
                    Return False
                End If
                If (edt_qt_doanthe.Text.Trim() = "") Then
                    MessageBox.Show("Chi bộ đảng không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_qt_doanthe
                    Return False
                End If
                If dtpk_qt_tungay.Value >= dtpk_qt_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu quá trình không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_qt_denngay
                    Return False
                End If

                'Kiểm tra xem chọn khoảng thời gian sinh hoạt đã hợp lệ chưa
                strSQL = String.Format("Select * from HS_DangVien Where IdDangVien = '{0}' Order By NgayKN Asc", _DocumentId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (dtpk_qt_tungay.Value < CType(db.Rows(0)("NgayKN").ToString(), DateTime)) Then
                                MessageBox.Show("Ngày bắt đầu quá trình sinh hoạt không trước ngày kết nạp đảng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_qt_tungay
                                Return False
                            End If
                            If (db.Rows(0)("NgayRa").ToString() <> "" AndAlso CType(db.Rows(0)("NgayRa"), DateTime).ToString("dd/MM/yyyy") <> "01/01/1900") Then
                                If (dtpk_qt_denngay.Value > CType(db.Rows(0)("NgayRa").ToString(), DateTime)) Then
                                    MessageBox.Show("Thời gian kết thúc quá trình sinh hoạt không hợp lệ!" + vbCrLf + "Lưu ý: Thời gian kết thúc quá trình sinh hoạt không thể lớn hơn ngày ra khỏi đảng trong hồ sơ đảng", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = dtpk_qt_denngay
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End Using

                'Kiểm tra khoảng thời gian đan xen lẫn nhau
                If (_RowId = "") Then
                    strSQL = String.Format("Select * from HS_Dang Where IdDangVien = '{0}' Order by TuNgay Asc", _DocumentId)
                Else
                    strSQL = String.Format("Select * from HS_Dang Where IdDangVien = '{0}' and IdHSDang <> '{1}' Order by TuNgay Asc", _DocumentId, _RowId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_qt_tungay.Value And dtpk_qt_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_qt_denngay.Value And dtpk_qt_denngay.Value <= _DenNgay) Or (dtpk_qt_tungay.Value <= _TuNgay And dtpk_qt_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc quá trình sinh hoạt này không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_qt_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using
            Case 2  'Quá trình sinh hoạt Đoàn
                If (cb_qt_chucvu.SelectedIndex <= 0 And cb_qt_chucvu.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn chức vụ đoàn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_qt_chucvu
                    Return False
                End If
                If (edt_qt_doanthe.Text.Trim() = "") Then
                    MessageBox.Show("Chi đoàn sinh hoạt không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_qt_doanthe
                    Return False
                End If
                If dtpk_qt_tungay.Value >= dtpk_qt_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu quá trình không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_qt_denngay
                    Return False
                End If

                'Kiểm tra xem chọn khoảng thời gian sinh hoạt đã hợp lệ chưa
                strSQL = String.Format("Select * from HS_DoanVien Where IdDoanVien = '{0}' Order By NgayVao Asc", _DocumentId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (dtpk_qt_tungay.Value < CType(db.Rows(0)("NgayVao").ToString(), DateTime)) Then
                                MessageBox.Show("Ngày bắt đầu quá trình sinh hoạt không trước ngày vào đoàn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_qt_tungay
                                Return False
                            End If
                            If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                If (dtpk_qt_denngay.Value > CType(db.Rows(0)("NgayRa").ToString(), DateTime)) Then
                                    MessageBox.Show("Thời gian kết thúc quá trình sinh hoạt không hợp lệ!" + vbCrLf + "Lưu ý: Thời gian kết thúc quá trình sinh hoạt không thể lớn hơn ngày ra khỏi đoàn trong hồ sơ đoàn", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = dtpk_qt_denngay
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End Using

                'Kiểm tra khoảng thời gian trùng Hoặc đan xen lẫn nhau
                If (_RowId = "") Then
                    strSQL = String.Format("Select * from HS_Doan Where IdDoanVien = '{0}' Order by TuNgay Asc", _DocumentId)
                Else
                    strSQL = String.Format("Select * from HS_Doan Where IdDoanVien = '{0}' and IdHSDoan <> '{1}' Order by TuNgay Asc", _DocumentId, _RowId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_qt_tungay.Value And dtpk_qt_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_qt_denngay.Value And dtpk_qt_denngay.Value <= _DenNgay) Or (dtpk_qt_tungay.Value <= _TuNgay And dtpk_qt_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc quá trình sinh hoạt này không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_qt_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using
            Case 3  'Quá trình sinh hoạt Công đoàn
                If (cb_qt_chucvu.SelectedIndex <= 0 And cb_qt_chucvu.Items.Count <> 0) Then
                    MessageBox.Show("Bạn chưa chọn chức vụ công đoàn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = cb_qt_chucvu
                    Return False
                End If
                If (edt_qt_doanthe.Text.Trim() = "") Then
                    MessageBox.Show("Cơ quan (tổ chức) sinh hoạt không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_qt_doanthe
                    Return False
                End If
                If dtpk_qt_tungay.Value >= dtpk_qt_denngay.Value Then
                    MessageBox.Show("Ngày bắt đầu quá trình không thể lớn hơn hoặc bằng ngày kết thúc!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = dtpk_qt_denngay
                    Return False
                End If

                'Kiểm tra xem chọn khoảng thời gian sinh hoạt đã hợp lệ chưa
                strSQL = String.Format("Select * from HS_CongDoan Where IdCongDoan = '{0}'", _DocumentId)
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (dtpk_qt_tungay.Value < CType(db.Rows(0)("NgayVao").ToString(), DateTime)) Then
                                MessageBox.Show("Ngày bắt đầu quá trình sinh hoạt không trước ngày vào công đoàn!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_qt_tungay
                                Return False
                            End If
                            If (db.Rows(0)("NgayRa").ToString() <> "") Then
                                If (dtpk_qt_denngay.Value > CType(db.Rows(0)("NgayRa").ToString(), DateTime)) Then
                                    MessageBox.Show("Thời gian kết thúc quá trình sinh hoạt không hợp lệ!" + vbCrLf + "Lưu ý: Thời gian kết thúc quá trình sinh hoạt không thể lớn hơn ngày ra khỏi công đoàn trong hồ sơ công đoàn.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                    ActiveControl = dtpk_qt_denngay
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End Using

                'Kiểm tra khoảng thời gian trùng hoặc đan xen lẫn nhau
                If (_RowId = "") Then
                    strSQL = String.Format("Select * from HS_CongDoanQT Where IdCongDoan = '{0}' Order by TuNgay ASC", _DocumentId)
                Else
                    strSQL = String.Format("Select * from HS_CongDoanQT Where IdCongDoan = '{0}' and IdHSCDoan <> '{1}' Order by TuNgay ASC", _DocumentId, _RowId)
                End If
                Dim Exist As Boolean = False
                Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            For i As Integer = 0 To db.Rows.Count - 1
                                Dim _TuNgay As DateTime = CType(db.Rows(i)("TuNgay").ToString(), DateTime)
                                Dim _DenNgay As DateTime = CType(db.Rows(i)("DenNgay").ToString(), DateTime)
                                If ((_TuNgay <= dtpk_qt_tungay.Value And dtpk_qt_tungay.Value <= _DenNgay) Or (_TuNgay <= dtpk_qt_denngay.Value And dtpk_qt_denngay.Value <= _DenNgay) Or (dtpk_qt_tungay.Value <= _TuNgay And dtpk_qt_denngay.Value >= _DenNgay)) Then
                                    Exist = True
                                    Exit For
                                End If
                            Next
                            If (Exist = True) Then
                                MessageBox.Show("Khoảng thời gian ngày bắt đầu và ngày kết thúc quá trình sinh hoạt này không hợp lệ!" + vbCrLf + "Lưu ý: Khoảng thời gian không thể trùng hoặc đan xen lẫn nhau", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                                ActiveControl = dtpk_qt_tungay
                                Return False
                            End If
                        End If
                    End If
                End Using
        End Select
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện select row trên lưới dữ liệu --> Fill dữ liệu vào các controls
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi (Id record) đang select</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String)
        Dim dr As DataRow
        _RowId = _Idnew
        dr = _HS_DangDT.GetRecordProcess(_RowId, FlagDocument)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                If dr("TuNgay").ToString().Trim() <> "" Then
                    dtpk_qt_tungay.Value = CType(dr("TuNgay").ToString(), DateTime)
                End If
                If dr("DenNgay").ToString().Trim() <> "" Then
                    If (CType(dr("DenNgay"), DateTime).ToString("dd/MM/yyyy") = "01/01/1900") Then
                        dtpk_qt_denngay.Checked = False
                    Else
                        dtpk_qt_denngay.Checked = True
                        dtpk_qt_denngay.Value = CType(dr("DenNgay").ToString(), DateTime)
                    End If
                Else
                    dtpk_qt_denngay.Checked = False
                End If
                'If dr("DenNgay").ToString().Trim() <> "" Then
                '    dtpk_qt_denngay.Value = CType(dr("DenNgay").ToString(), DateTime)
                'End If
                edt_qt_ghichu.Text = dr("GhiChu").ToString().Trim()
                If (FlagDocument = 1) Then
                    cb_qt_chucvu.SelectedIndex = CType(arr_Chucvu.IndexOf(dr("IdCVDang").ToString()), Integer)
                    edt_qt_doanthe.Text = dr("ChiBo").ToString().Trim()
                    If dr("IsKiemNhiem").ToString().Trim() <> "" Then
                        If (dr("IsKiemNhiem").ToString().Trim() = True) Then
                            ckb_kiemnhiem.Checked = True
                        Else
                            ckb_kiemnhiem.Checked = False
                        End If
                    Else
                        ckb_kiemnhiem.Checked = False
                    End If
                ElseIf (FlagDocument = 2) Then
                    cb_qt_chucvu.SelectedIndex = CType(arr_Chucvu.IndexOf(dr("IdCVDoan").ToString()), Integer)
                    edt_qt_doanthe.Text = dr("ChiDoan").ToString().Trim()
                ElseIf (FlagDocument = 3) Then
                    cb_qt_chucvu.SelectedIndex = CType(arr_Chucvu.IndexOf(dr("IdChucVu").ToString()), Integer)
                    edt_qt_doanthe.Text = dr("CoQuan").ToString().Trim()
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện kiểm tra quyền thao tác chương trình của người dùng
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        Select Case FlagDocument
            Case 1  'Quá trình sinh hoạt Đảng
                'Nếu không có quyền Xem thông tin với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";163;") < 0) Then
                    dgv_main.Enabled = False
                End If
                'Nếu không có quyền thêm mới với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";164;") < 0) Then
                    btn_add.Enabled = False
                End If
                ''Nếu không có quyền xoá bỏ với cả 3 hồ sơ
                If (Globals.Roles.IndexOf("166;") < 0) Then
                    btn_delete.Enabled = False
                End If
                ''Nếu không có quyền thêm mới và sửa đổi với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";164;") < 0 And Globals.Roles.IndexOf(";165;") < 0) Then
                    btn_accept.Enabled = False
                End If
            Case 2  'Quá trình sinh hoạt Đoàn
                'Nếu không có quyền Xem thông tin với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";220;") < 0) Then
                    dgv_main.Enabled = False
                End If
                'Nếu không có quyền thêm mới với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";221;") < 0) Then
                    btn_add.Enabled = False
                End If
                ''Nếu không có quyền xoá bỏ với cả 3 hồ sơ
                If (Globals.Roles.IndexOf("223;") < 0) Then
                    btn_delete.Enabled = False
                End If
                ''Nếu không có quyền thêm mới và sửa đổi với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";221;") < 0 And Globals.Roles.IndexOf(";222;") < 0) Then
                    btn_accept.Enabled = False
                End If
            Case 3  'Quá trình sinh hoạt Công đoàn
                'Nếu không có quyền Xem thông tin với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";225;") < 0) Then
                    dgv_main.Enabled = False
                End If
                'Nếu không có quyền thêm mới với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";226;") < 0) Then
                    btn_add.Enabled = False
                End If
                ''Nếu không có quyền xoá bỏ với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";228;") < 0) Then
                    btn_delete.Enabled = False
                End If
                ''Nếu không có quyền thêm mới và sửa đổi với cả 3 hồ sơ
                If (Globals.Roles.IndexOf(";226;") < 0 And Globals.Roles.IndexOf(";227;") < 0) Then
                    btn_accept.Enabled = False
                End If
        End Select
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHS_DoanTheQT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _HS_DangDT.Create_Frame(dgv_main, IIf(FlagDocument = 1, 2, IIf(FlagDocument = 2, 4, 6)))
        AddHeaderCheckBox()
        'Gọi một số sự kiện liên quan đến Check all items trên lưới dữ liệu
        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged
        AddHandler dgv_main.CellPainting, AddressOf dgv_main_CellPainting
        AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged

        arr_Chucvu.Clear()
        cb_qt_chucvu.Items.Clear()
        pnl_inputdata.Height = 90
        lbl_div_kn.Visible = False
        pnl_kiemnhiem.Visible = False
        Select Case FlagDocument
            Case 1  'Quá trình sinh hoạt Đảng
                pnl_inputdata.Height = 110
                lbl_div_kn.Visible = True
                pnl_kiemnhiem.Visible = True
                arr_Chucvu = _Globals.Bind_ComBoBox(cb_qt_chucvu, clsHT_DanhMuc.Sql_Cv_dang, "---Chọn chức vụ đảng---")
                lbl_title.Text = " QUÁ TRÌNH SINH HOẠT ĐẢNG"
                lbl_chucvu.Text = "Chức vụ đảng "
                lbl_doanthe.Text = "Chi bộ "
            Case 2  'Quá trình sinh hoạt Đoàn
                arr_Chucvu = _Globals.Bind_ComBoBox(cb_qt_chucvu, clsHT_DanhMuc.Sql_Cv_doan, "---Chọn chức vụ đoàn---")
                lbl_title.Text = " QUÁ TRÌNH SINH HOẠT ĐOÀN"
                lbl_chucvu.Text = "Chức vụ đoàn "
                lbl_doanthe.Text = "Chi đoàn "
            Case 3  'Quá trình sinh hoạt Công đoàn
                arr_Chucvu = _Globals.Bind_ComBoBox(cb_qt_chucvu, clsHT_DanhMuc.Sql_Cv_Congdoan, "---Chọn chức vụ công đoàn---")
                lbl_title.Text = " QUÁ TRÌNH SINH HOẠT CÔNG ĐOÀN"
                lbl_chucvu.Text = "Chức vụ công đoàn "
                lbl_doanthe.Text = "Cơ quan "
        End Select
        ResetAll_Controls()
        Fill_Data()
        dgv_main_CellClick(sender, Nothing)
    End Sub

    Private Sub dgv_main_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_main.CellClick
        ResetAll_Controls()
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                SelectRow(dgv_main.CurrentRow.Cells("cln_Code").Value.ToString())
            End If
        End If
    End Sub

    Private Sub dgv_main_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_add.Click
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
        If (_DocumentId = "") Then Return
        ResetAll_Controls()
        Select Case FlagDocument
            Case 1  'Quá trình sinh hoạt Đảng
                'Set Ngày tiếp theo của quá trình sinh hoạt đảng
                If (dgv_main.Rows.Count = 0) Then
                    strSQL = String.Format("Select * From HS_DangVien Where IdDangVien = '{0}'", _DocumentId)
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If (_db.Rows.Count > 0) Then
                                If (_db.Rows(0)("NgayKN").ToString() <> "") Then
                                    'Ngày bắt đầu sinh hoạt phải trước ngày vào chính thức ít nhất là 1 năm (1 Năm thủ - thách)
                                    Dim dt_ngay_bd As DateTime = New DateTime(CType(_db.Rows(0)("NgayKN").ToString(), DateTime).Year, CType(_db.Rows(0)("NgayKN").ToString(), DateTime).Month, CType(_db.Rows(0)("NgayKN").ToString(), DateTime).Day)
                                    dtpk_qt_tungay.Value = dt_ngay_bd
                                End If
                            End If
                        End If
                    End Using
                Else
                    strSQL = String.Format("Select * From HS_Dang Where IdDangVien = '{0}' Order by DenNgay Desc", _DocumentId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                If (db.Rows(0)("DenNgay").ToString() <> "") Then
                                    Dim dt_ngaytiep As DateTime = CType(db.Rows(0)("DenNgay").ToString(), DateTime).AddDays(1)
                                    dtpk_qt_tungay.Value = dt_ngaytiep
                                End If
                                edt_qt_doanthe.Text = db.Rows(0)("ChiBo").ToString().Trim()
                            End If
                        End If
                    End Using
                End If
                If (cb_qt_chucvu.Items.Count <> 0) Then
                    cb_qt_chucvu.SelectedIndex = cb_qt_chucvu.FindString("Ðảng viên")
                End If
            Case 2  'Quá trình sinh hoạt Đoàn
                'Set Ngày tiếp theo của quá trình sinh hoạt đảng
                If (dgv_main.Rows.Count = 0) Then
                    strSQL = String.Format("Select * from HS_DoanVien Where IdDoanVien = '{0}'", _DocumentId)
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If (_db.Rows.Count > 0) Then
                                If (_db.Rows(0)("NgayVao").ToString() <> "") Then
                                    'Ngày bắt đầu sinh hoạt phải trước ngày vào chính thức ít nhất là 1 năm (1 Năm thủ - thách)
                                    Dim dt_ngaydb As DateTime = New DateTime(CType(_db.Rows(0)("NgayVao").ToString(), DateTime).Year, CType(_db.Rows(0)("NgayVao").ToString(), DateTime).Month, CType(_db.Rows(0)("NgayVao").ToString(), DateTime).Day)
                                    dtpk_qt_tungay.Value = dt_ngaydb
                                End If
                            End If
                        End If
                    End Using
                Else
                    strSQL = String.Format("Select * from HS_Doan Where IdDoanVien = '{0}' Order by DenNgay Desc", _DocumentId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                If (db.Rows(0)("DenNgay").ToString() <> "") Then
                                    Dim dt_ngaytiep As DateTime = CType(db.Rows(0)("DenNgay").ToString(), DateTime).AddDays(1)
                                    dtpk_qt_tungay.Value = dt_ngaytiep
                                End If
                            End If
                        End If
                    End Using
                End If
            Case 3  'Quá trình sinh hoạt Công đoàn
                'Set Ngày tiếp theo của quá trình sinh hoạt đảng
                If (dgv_main.Rows.Count = 0) Then
                    strSQL = String.Format("Select * from HS_CongDoan Where IdCongDoan = '{0}'", _DocumentId)
                    Using _db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (_db Is Nothing) Then
                            If (_db.Rows.Count > 0) Then
                                If (_db.Rows(0)("NgayVao").ToString() <> "") Then
                                    'Ngày bắt đầu sinh hoạt phải trước ngày vào chính thức ít nhất là 1 năm (1 Năm thủ - thách)
                                    Dim dt_ngaydb As DateTime = New DateTime(CType(_db.Rows(0)("NgayVao").ToString(), DateTime).Year, CType(_db.Rows(0)("NgayVao").ToString(), DateTime).Month, CType(_db.Rows(0)("NgayVao").ToString(), DateTime).Day)
                                    dtpk_qt_tungay.Value = dt_ngaydb
                                End If
                            End If
                        End If
                    End Using
                Else
                    strSQL = String.Format("Select * from HS_CongDoanQT Where IdCongDoan = '{0}' Order by DenNgay Desc", _DocumentId)
                    Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
                        If Not (db Is Nothing) Then
                            If (db.Rows.Count > 0) Then
                                If (db.Rows(0)("DenNgay").ToString() <> "") Then
                                    Dim dt_ngaytiep As DateTime = CType(db.Rows(0)("DenNgay").ToString(), DateTime).AddDays(1)
                                    dtpk_qt_tungay.Value = dt_ngaytiep
                                End If
                            End If
                        End If
                    End Using
                End If
        End Select
        ActiveControl = dtpk_qt_tungay
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        If (_DocumentId = "") Then Return
        Dim _currRow As String = ""
        '-->Kiểm tra Quyền thêm mới hoặc sửa đổi Quá trình sinh hoạt
        If (_RowId <> "") Then
            _currRow = _RowId
            Select Case FlagDocument
                Case 1  'Quá trình sinh hoạt Đảng
                    If (Globals.Roles.IndexOf(";165;") < 0) Then
                        MessageBox.Show("Bạn không có quyền thực hiện sửa đổi quá trình sinh hoạt đảng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls()
                        dgv_main.CurrentRow.Selected = False
                        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
                        If (dgv_main.Rows.Count > 0) Then
                            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        ckb_ChoiceAll.Checked = False
                        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                        Return
                    End If
                Case 2  'Quá trình sinh hoạt Đoàn
                    If (Globals.Roles.IndexOf(";227;") < 0) Then
                        MessageBox.Show("Bạn không có quyền thực hiện sửa đổi quá trình sinh hoạt đoàn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls()
                        dgv_main.CurrentRow.Selected = False
                        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
                        If (dgv_main.Rows.Count > 0) Then
                            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        ckb_ChoiceAll.Checked = False
                        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                        Return
                    End If
                Case 3  'Quá trình sinh hoạt Công đoàn
                    If (Globals.Roles.IndexOf(";228;") < 0) Then
                        MessageBox.Show("Bạn không có quyền thực hiện sửa đổi quá trình sinh hoạt công đoàn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls()
                        dgv_main.CurrentRow.Selected = False
                        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
                        If (dgv_main.Rows.Count > 0) Then
                            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                                SelectRow(_currRow)
                            End If
                        End If
                        ckb_ChoiceAll.Checked = False
                        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                        Return
                    End If
            End Select
        End If
        '-->Thực hiện cập nhật dữ liệu quá trình sinh hoạt
        If (IsValid()) Then
            Select Case FlagDocument
                Case 1
                    Dim obj_qt_dang As clsHS_DangDT.HS_Dang = New clsHS_DangDT.HS_Dang()
                    obj_qt_dang.IdDangVien = _DocumentId
                    obj_qt_dang.TuNgay = dtpk_qt_tungay.Value
                    obj_qt_dang.DenNgay = IIf(dtpk_qt_denngay.Checked = True, dtpk_qt_denngay.Value, DateTime.Parse("01/01/1900"))
                    obj_qt_dang.IdCVDang = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_qt_chucvu.SelectedIndex), "0"), Int32)
                    obj_qt_dang.ChiBo = Globals.Find_Replace(edt_qt_doanthe.Text.ToString().Trim())
                    obj_qt_dang.IsKiemNhiem = CType(IIf(ckb_kiemnhiem.Checked = True, 1, 0), Byte)
                    obj_qt_dang.GhiChu = Globals.Find_Replace(edt_qt_ghichu.Text.ToString().Trim())
                    If (_RowId = "") Then
                        _currRow = _HS_DangDT.Insert_PartyProcess(obj_qt_dang)
                    Else
                        _currRow = _RowId
                        obj_qt_dang.IdHSDang = _RowId
                        _HS_DangDT.Update_PartyProcess(obj_qt_dang)
                    End If
                Case 2
                    Dim obj_qt_doan As clsHS_DangDT.HS_Doan = New clsHS_DangDT.HS_Doan()
                    obj_qt_doan.IdDoanVien = _DocumentId
                    obj_qt_doan.TuNgay = dtpk_qt_tungay.Value
                    obj_qt_doan.DenNgay = IIf(dtpk_qt_denngay.Checked = True, dtpk_qt_denngay.Value, DateTime.Parse("01/01/1900"))
                    obj_qt_doan.IdCVDoan = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_qt_chucvu.SelectedIndex), "0"), Int32)
                    obj_qt_doan.ChiDoan = Globals.Find_Replace(edt_qt_doanthe.Text.ToString().Trim())
                    obj_qt_doan.GhiChu = Globals.Find_Replace(edt_qt_ghichu.Text.ToString().Trim())
                    If (_RowId = "") Then
                        _currRow = _HS_DangDT.Insert_UnionMemberProcess(obj_qt_doan)
                    Else
                        _currRow = _RowId
                        obj_qt_doan.IdHSDoan = _RowId
                        _HS_DangDT.Update_UnionMemberProcess(obj_qt_doan)
                    End If
                Case 3
                    Dim obj_qt_congdoan As clsHS_DangDT.HS_CongDoanQT = New clsHS_DangDT.HS_CongDoanQT()
                    obj_qt_congdoan.IdCongDoan = _DocumentId
                    obj_qt_congdoan.TuNgay = dtpk_qt_tungay.Value
                    obj_qt_congdoan.DenNgay = IIf(dtpk_qt_denngay.Checked = True, dtpk_qt_denngay.Value, DateTime.Parse("01/01/1900"))
                    obj_qt_congdoan.IdChucVu = CType(IIf(arr_Chucvu.Count > 0, arr_Chucvu(cb_qt_chucvu.SelectedIndex), "0"), Int32)
                    obj_qt_congdoan.CoQuan = Globals.Find_Replace(edt_qt_doanthe.Text.ToString().Trim())
                    obj_qt_congdoan.GhiChu = Globals.Find_Replace(edt_qt_ghichu.Text.ToString().Trim())
                    If (_RowId = "") Then
                        _currRow = _HS_DangDT.Insert_TradeUnionProcess(obj_qt_congdoan)
                    Else
                        _currRow = _RowId
                        obj_qt_congdoan.IdHSCDoan = _RowId
                        _HS_DangDT.Update_TradeUnionProcess(obj_qt_congdoan)
                    End If
            End Select
            ResetAll_Controls()
            Fill_Data()
            'Tìm lại dòng đang select trước đó
            dgv_main.CurrentRow.Selected = False
            dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
            'Select row trên lưới dữ liệu
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    SelectRow(_currRow)
                End If
            End If
        End If
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        If (dgv_main.Rows.Count <= 0) Then Return
        'Xét quyền được xoá quá trình sinh hoạt của hồ sơ đoàn thể
        Dim vName As String = ""
        Select Case FlagDocument
            Case 1
                vName = "đảng"
                If (Globals.Roles.IndexOf(";166;") < 0) Then
                    MessageBox.Show("Bạn không có quyền thực hiện xoá quá trình sinh hoạt đảng của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                    Return
                End If
            Case 2
                vName = "đoàn"
                If (Globals.Roles.IndexOf(";223;") < 0) Then
                    MessageBox.Show("Bạn không có quyền thực hiện xoá quá trình sinh hoạt đoàn của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                    Return
                End If
            Case 3
                vName = "công đoàn"
                If (Globals.Roles.IndexOf(";228;") < 0) Then
                    MessageBox.Show("Bạn không có quyền thực hiện xoá quá trình sinh hoạt công đoàn của cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button3)
                    Return
                End If
        End Select
        'Thực hiện xoá quá trình sinh hoạt của hồ sơ đoàn thể
        Try
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _Count As Int16 = 0
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    For i As Int32 = 0 To dgv_main.Rows.Count - 1
                        If (dgv_main.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                            If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                _Count = _Count + 1
                                arr_Del.Add(dgv_main.Rows(i).Cells("cln_Code").Value.ToString())
                            End If
                        End If
                    Next
                End If
            End If
            If (_Count > 0) Then
                Dim _mess As String = IIf(_Count = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi quá trình sinh hoạt " + vName + " đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    For i As Int16 = 0 To arr_Del.Count - 1
                        If (FlagDocument = 1) Then
                            _HS_DangDT.Delete_PartyProcess(arr_Del(i).ToString())
                        ElseIf (FlagDocument = 2) Then
                            _HS_DangDT.Delete_UnionMemberProcess(arr_Del(i).ToString())
                        ElseIf (FlagDocument = 3) Then
                            _HS_DangDT.Delete_TradeUnionProcess(arr_Del(i).ToString())
                        End If
                    Next
                    'Thực hiện load thông tin sau khi xoá song
                    Fill_Data()
                    dgv_main_CellClick(sender, Nothing)
                Else    'Nếu không muốn xoá nữa
                    arr_Del.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá thông tin quá trình sinh hoạt " + vName + " :" + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        ckb_ChoiceAll.Checked = False
        HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
        Dim _currRow As String = _RowId
        ResetAll_Controls()
        If (dgv_main.Rows.Count <= 0) Then Return
        If (_currRow = "") Then _currRow = dgv_main.CurrentRow.Cells("cln_Code").Value.ToString()
        dgv_main.CurrentRow.Selected = False
        dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                SelectRow(_currRow)
            End If
        End If
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Close()
    End Sub

    Private Sub dtpk_qt_tungay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_qt_tungay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            dtpk_qt_denngay.Focus()
        End If
    End Sub

    Private Sub dtpk_qt_denngay_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpk_qt_denngay.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            cb_qt_chucvu.Focus()
        End If
    End Sub

    Private Sub cb_qt_chucvu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cb_qt_chucvu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) Then
            edt_qt_doanthe.Focus()
        End If
    End Sub

    Private Sub edt_qt_doanthe_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qt_doanthe.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            edt_qt_ghichu.Focus()
        End If
    End Sub

    Private Sub edt_qt_doanthe_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qt_doanthe.KeyUp
        If (e.KeyCode = Keys.Up) Then cb_qt_chucvu.Focus()
    End Sub

    Private Sub edt_qt_ghichu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qt_ghichu.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            btn_accept.Focus()
        End If
    End Sub

    Private Sub edt_qt_ghichu_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_qt_ghichu.KeyUp
        If (e.KeyCode = Keys.Up) Then edt_qt_doanthe.Focus()
    End Sub
#End Region

#Region "---> Events: Các Hàm và sự kiện liên quan đến Check all các items trên lưới dữ liệu <---"
    Private Sub ckb_ChoiceAll_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Space) Then
            HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
        End If
    End Sub

    Private Sub ckb_ChoiceAll_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
    End Sub

    Private Sub dgv_main_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        'Sự kiện này thực hiện khi click các checked items trên lưới dữ liệu thì sẽ Checked hoặc UnChecked Checkall
        If CType(sender, DataGridView).Columns(e.ColumnIndex).Name = "cln_Choice" And e.RowIndex >= 0 Then
            If Not IsHeaderCheckBoxClicked Then
                Dim _vCellCheck As DataGridViewCheckBoxCell
                _vCellCheck = CType(dgv_main("cln_Choice", e.RowIndex), DataGridViewCheckBoxCell)
                RowCheckBoxClick(_vCellCheck)
            End If
        End If
    End Sub

    Private Sub dgv_main_CurrentCellDirtyStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Dim vCellCheck As DataGridViewCheckBoxCell
                'Checking whether the Datagridview Checkbox column is the first column
                If dgv_main.CurrentCellAddress.X = 0 Then
                    vCellCheck = dgv_main.CurrentRow.Cells("cln_Choice")
                    If (dgv_main.IsCurrentCellDirty) Then 'Checking for dirty cell
                        dgv_main.CommitEdit(DataGridViewDataErrorContexts.Commit) 'If it is dirty, making them to commit
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub dgv_main_CellPainting(ByVal sender As System.Object, ByVal e As DataGridViewCellPaintingEventArgs)
        'Căn chỉnh cho Checkbox nằm vào giữa tiêu đề của cột Chọn xoá
        If (e.RowIndex = -1 And e.ColumnIndex = 0) Then
            ResetHeaderCheckBoxLocation(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    '-------------- Các hàm liên quan --------------'
    ''' <summary>
    ''' Hàm thực hiện add một CheckBox vào Cột tiêu đề chọn tất cả để xóa các bản ghi đã đánh dấu xóa
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddHeaderCheckBox()
        ckb_ChoiceAll = New CheckBox()
        ckb_ChoiceAll.Size = New Size(15, 15)
        'Add the CheckBox into the DataGridView
        ckb_ChoiceAll.Checked = False
        dgv_main.Controls.Add(ckb_ChoiceAll)
    End Sub

    ''' <summary>
    '''  Hàm thực hiện căn chỉnh cho CheckBox nằm giữa cột tiêu đề của lưới dữ liệu
    ''' </summary>
    ''' <param name="_ColumnIndex"></param>
    ''' <param name="_RowIndex"></param>
    ''' <remarks></remarks>
    Private Sub ResetHeaderCheckBoxLocation(ByVal _ColumnIndex As Integer, ByVal _RowIndex As Integer)
        'Get the column header cell bounds
        Dim oRectangle As Rectangle = dgv_main.GetCellDisplayRectangle(_ColumnIndex, _RowIndex, True)
        Dim oPoint As Point = New Point()
        oPoint.X = oRectangle.Location.X + (oRectangle.Width - ckb_ChoiceAll.Width) / 2 + 1
        oPoint.Y = oRectangle.Location.Y + (oRectangle.Height - ckb_ChoiceAll.Height) / 2 + 1
        'Change the location of the CheckBox to make it stay on the header
        ckb_ChoiceAll.Location = oPoint
    End Sub

    ''' <summary>
    '''  Hàm thực hiện các thao tác khi Checkbox (Chọn cả) được Check click
    ''' </summary>
    ''' <param name="ckb_CheckAll"></param>
    ''' <param name="dgv_name"></param>
    ''' <remarks></remarks>
    Private Sub HeaderCheckBoxClick(ByVal ckb_CheckAll As CheckBox, ByVal dgv_name As DataGridView)
        If (dgv_name.Rows.Count <= 0) Then
            ckb_CheckAll.Checked = False
            Return
        End If
        IsHeaderCheckBoxClicked = True
        For i As Integer = 0 To dgv_name.Rows.Count - 1
            dgv_name.Rows(i).Cells("cln_Choice").Value = ckb_CheckAll.Checked
        Next
        dgv_name.RefreshEdit()
        TotalCheckedCheckBoxes = IIf(ckb_CheckAll.Checked, TotalCheckBoxes, 0)
        IsHeaderCheckBoxClicked = False
    End Sub

    ''' <summary>
    ''' Hàm thực hiện các công việc khi item trên lưới dữ liệu được Checked
    ''' </summary>
    ''' <param name="ckb_CellCheck"></param>
    ''' <remarks></remarks>
    Private Sub RowCheckBoxClick(ByVal ckb_CellCheck As DataGridViewCheckBoxCell)
        If Not (ckb_CellCheck Is Nothing) Then
            'Modifiy Counter
            If ((CType(ckb_CellCheck.Value, Boolean) = True) And (TotalCheckedCheckBoxes < TotalCheckBoxes)) Then
                TotalCheckedCheckBoxes += 1
            ElseIf (TotalCheckedCheckBoxes > 0) Then
                TotalCheckedCheckBoxes -= 1
            End If
            'Change state of the header CheckBox
            If (TotalCheckedCheckBoxes < TotalCheckBoxes) Then
                ckb_ChoiceAll.Checked = False
            ElseIf (TotalCheckedCheckBoxes = TotalCheckBoxes) Then
                ckb_ChoiceAll.Checked = True
            End If
        End If
    End Sub
#End Region

End Class