Public Class frmHT_NhomThVien

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Systems As clsHeThong = New clsHeThong()
    Private _SqlHelper As DBAccess = New DBAccess()

    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
    Private _Code As String
    Private _CodeGroup As String
    Public Property CodeGroup() As String
        Get
            Return _CodeGroup
        End Get
        Set(ByVal value As String)
            _CodeGroup = value
        End Set
    End Property

    Private _FlagUpdate As Boolean = False
    Public Property FlagUpdate() As Boolean
        Get
            Return _FlagUpdate
        End Get
        Set(ByVal value As Boolean)
            _FlagUpdate = value
        End Set
    End Property
#End Region

#Region "---> Functions: Các hàm dùng chung <---"
    ''' <summary>
    ''' Hàm thực hiện reset control và thiết lập trạng thái ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetAll_Controls()
        _Code = ""
        edt_code.Text = ""
        edt_name.Text = ""
        edt_note.Text = ""
        ckb_choiceall.Checked = False
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu danh sách Nhóm thành viên
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Fill_Data()
        'Thực hiện fill dữ liệu danh sách nhóm người dùng ra lưới
        Using db As DataTable = _Systems.GetAll_GroupMembers()
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_main.Rows.Add()
                        dgv_main.Rows(i).Cells("cln_Code").Value = IIf(db.Rows(i)("ma_nhom").ToString() <> "", db.Rows(i)("ma_nhom").ToString(), "")
                        dgv_main.Rows(i).Cells("cln_Name").Value = IIf(db.Rows(i)("ten_nhom").ToString() <> "", db.Rows(i)("ten_nhom").ToString(), "")
                        dgv_main.Rows(i).Cells("cln_Note").Value = IIf(db.Rows(i)("ghi_chu").ToString() <> "", db.Rows(i)("ghi_chu").ToString(), "")
                        dgv_main.Rows(i).Cells("cln_CreationDate").Value = IIf(db.Rows(i)("ngay_tao").ToString().Trim() <> "", CType(db.Rows(i)("ngay_tao").ToString(), DateTime).ToString("dd-MM-yyyy"), "")
                    Next
                End If
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Hàm kiểm tra điều kiện cập nhật dữ liệu - Nhóm thành viên
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        If (edt_code.Text.Trim() = "") Then
            MessageBox.Show("Mã hiệu nhóm thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_code
            Return False
        End If
        If (edt_name.Text.Trim() = "") Then
            MessageBox.Show("Tên nhóm thành viên không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_name
            Return False
        End If
        'Kiểm tra điều kiện trùng dữ liệu - Kiểm tra trùng mã hiệu của nhóm
        Dim strSQL As String = ""
        If (_Code = "") Then
            strSQL = String.Format("Select * From HT_NhomTV Where ma_nhom = '{0}'", Globals.Find_Replace(edt_code.Text.ToString().Trim()))
        ElseIf (_Code <> "") Then
            strSQL = String.Format("Select * From HT_NhomTV Where ma_nhom = '{0}' and ma_nhom <> '{1}'", Globals.Find_Replace(edt_code.Text.ToString().Trim()), _Code)
        End If
        Using db As DataTable = _SqlHelper.SelectDBRows(strSQL)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    MessageBox.Show("Mã hiệu của nhóm thành viên này đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                    ActiveControl = edt_code
                    Return False
                End If
            End If
        End Using
        Return True
    End Function

    ''' <summary>
    ''' Hàm thực hiện fill data theo row click trên lưới dữ liệu
    ''' </summary>
    ''' <param name="_Idnew">Chỉ số bản ghi truyền vào</param>
    ''' <remarks></remarks>
    Private Sub SelectRow(ByVal _Idnew As String)
        _Code = _Idnew
        Dim dr As DataRow
        dr = _Systems.GetGroupMember(_Code)
        If Not (dr Is Nothing) Then
            If (dr.Table.Rows.Count > 0) Then
                edt_code.Text = dr("ma_nhom").ToString().Trim()
                edt_name.Text = dr("ten_nhom").ToString().Trim()
                edt_note.Text = dr("ghi_chu").ToString().Trim()
            End If
        End If
    End Sub
#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub frmHT_NhomThVien_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _Systems.Create_Frame(dgv_main, False)
        dgv_main.Rows.Clear()
        Fill_Data()
        ResetAll_Controls()
        'Thiết lập phân quyền thành viên thao tác chương trình
        Dim _roles As String = Globals.Roles

        If (_roles.IndexOf(";45;") < 0) Then
            gb_main.Enabled = False
            dgv_main.Enabled = False
        End If
        If (_roles.IndexOf(";46;") < 0) Then
            btn_add.Visible = False
        End If
        If (_roles.IndexOf(";47;") < 0) Then
            btn_save.Visible = False
        End If
        If (_roles.IndexOf(";48;") < 0) Then
            btn_delete.Visible = False
        End If
    End Sub

    Private Sub ckb_choiceall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_choiceall.CheckedChanged
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                Globals.Check_All_Items(dgv_main, ckb_choiceall.Checked)
            Else
                ckb_choiceall.Checked = False
                Return
            End If
        Else
            ckb_choiceall.Checked = False
            Return
        End If
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
        ResetAll_Controls()
        ActiveControl = edt_code
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Dim _currRow As String = ""
        If (_Code = "") Then
            If (IsValid()) Then
                Dim obj_grmember As clsHeThong.GroupMember = New clsHeThong.GroupMember()
                obj_grmember.ma_nhom = Globals.Find_Replace(edt_code.Text.Trim.ToString())
                obj_grmember.ten_nhom = Globals.Find_Replace(edt_name.Text.Trim.ToString())
                obj_grmember.ghi_chu = Globals.Find_Replace(edt_note.Text.Trim.ToString())
                obj_grmember.ngay_tao = DateTime.Now
                If (_CodeGroup = "") Then
                    _CodeGroup = Globals.Find_Replace(edt_code.Text.Trim.ToString())
                End If
                _Systems.Insert_GroupMember(obj_grmember)
                _currRow = Globals.Find_Replace(edt_code.Text.Trim.ToString())
                ResetAll_Controls()
                dgv_main.Rows.Clear()
                Fill_Data()
                dgv_main.CurrentRow.Selected = False
                dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
                If (dgv_main.Rows.Count > 0) Then
                    If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
                _FlagUpdate = True
            End If
        ElseIf (_Code <> "") Then
            If (Globals.Roles.IndexOf(";47;") < 0) Then
                Return
            End If
            If (IsValid()) Then
                Dim obj_grmember As clsHeThong.GroupMember = New clsHeThong.GroupMember()
                obj_grmember.ma_nhom = Globals.Find_Replace(edt_code.Text.Trim.ToString())
                obj_grmember.ten_nhom = Globals.Find_Replace(edt_name.Text.Trim.ToString())
                obj_grmember.ghi_chu = Globals.Find_Replace(edt_note.Text.Trim.ToString())
                obj_grmember.ngay_tao = DateTime.Now
                _Systems.Update_GroupMember(obj_grmember, _Code)
                _CodeGroup = Globals.Find_Replace(edt_code.Text.Trim.ToString())
                'Fill data - Hồ sơ công tác của cán bộ
                _currRow = Globals.Find_Replace(edt_code.Text.Trim.ToString())
                ResetAll_Controls()
                dgv_main.Rows.Clear()
                Fill_Data()
                dgv_main.CurrentRow.Selected = False
                dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Code")).Selected = True
                If (dgv_main.Rows.Count > 0) Then
                    If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                        SelectRow(_currRow)
                    End If
                End If
                _FlagUpdate = True
            End If
        End If
    End Sub

    Private Sub btn_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_delete.Click
        Try
            If (dgv_main.Rows.Count <= 0) Then
                Return
            End If
            Dim arr_Del As ArrayList = New ArrayList()
            Dim _count As Int16 = 0
            Dim FlagOK As Boolean = False
            Dim DelUserAdmin As Boolean = False
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Code").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Code").Value.ToString() <> "")) Then
                    For i As Int32 = 0 To dgv_main.Rows.Count - 1
                        If (dgv_main.Rows(i).Cells("cln_Code").Value IsNot Nothing) Then
                            If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                If (dgv_main.Rows(i).Cells("cln_Code").Value.ToString().ToUpper = "ADMIN") Or (dgv_main.Rows(i).Cells("cln_Code").Value.ToString().ToUpper = "OPERATOR") Then
                                    DelUserAdmin = True
                                Else
                                    _count += 1
                                    Using db As DataTable = _Systems.GetAll_Members(dgv_main.Rows(i).Cells("cln_Code").Value.ToString())
                                        If Not (db Is Nothing) Then
                                            If (db.Rows.Count > 0) Then
                                                FlagOK = True
                                            End If
                                        Else
                                            arr_Del.Add(dgv_main.Rows(i).Cells("cln_Code").Value.ToString())
                                        End If
                                    End Using
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            If (_count > 0) Then
                Dim _mess As String = IIf(_count = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi nhóm thành viên đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    If (FlagOK = True) Then
                        If (_count = 1) Then
                            MessageBox.Show("Nhóm thành viên bạn chọn hiện đang được dùng, không thể xoá được!" + vbCrLf + "Lưu ý: Muốn xoá dữ liệu nhóm thành viên trước hết bạn hãy thực hiện xoá hết các thành viên thuộc nhóm thành viên này trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Else
                            MessageBox.Show("Danh sách nhóm thành viên bạn chọn hiện đang được dùng, không thể xoá được!" + vbCrLf + "Lưu ý: Muốn xoá dữ liệu nhóm thành viên trước hết bạn hãy thực hiện xoá hết các thành viên thuộc nhóm thành viên này trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        End If
                        arr_Del.Clear()
                        Globals.Check_All_Items(dgv_main, False)
                        ckb_choiceall.Checked = False
                    Else
                        'Thực hiện xoá dữ liệu khi đã chấp thuận
                        For i As Int16 = 0 To arr_Del.Count - 1
                            'Thực hiện xoá dữ liệu nhóm thành viên và Xoá quyền của nhóm thành viên
                            _Systems.Delete_GroupMember(arr_Del(i).ToString())
                        Next
                        'Load lại các thông tin cần thiết sau khi xoá dữ liệu
                        MessageBox.Show("Bạn đã xoá thành công thông tin nhóm thành viên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        ResetAll_Controls()
                    End If
                Else
                    arr_Del.Clear()
                    Globals.Check_All_Items(dgv_main, False)
                    ckb_choiceall.Checked = False
                End If
            Else
                MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                ckb_choiceall.Checked = False
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá bỏ thông tin nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        Dim _currRow As String = _Code
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

        'Thiết lập phân quyền thành viên thao tác chương trình
        Dim _roles As String = Globals.Roles
        If (_roles.IndexOf(";45;") < 0) Then
            gb_main.Enabled = False
            dgv_main.Enabled = False
        End If
        If (_roles.IndexOf(";46;") < 0) Then
            btn_add.Visible = False
        End If
        If (_roles.IndexOf(";48;") < 0) Then
            btn_add.Visible = False
        End If
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Close()
    End Sub

    Private Sub frmHT_NhomThVien_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub frmHT_NhomThVien_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Close()
        End If
    End Sub
#End Region

#Region "---> Bắt các sự kiện Ngoại lệ của người dùng <---"
    Private Sub edt_code_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_code.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_code.Text.Trim() <> "") Then
                edt_name.Focus()
            Else
                edt_code.Focus()
            End If
        End If
    End Sub

    Private Sub edt_name_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_name.KeyDown
        If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Down) Then
            If (edt_name.Text.Trim() <> "") Then
                edt_note.Focus()
            Else
                edt_name.Focus()
            End If
        End If
    End Sub

    Private Sub edt_note_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_note.KeyDown
        If (e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Down) Then
            btn_save.Focus()
        End If
    End Sub

    Private Sub edt_note_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_note.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_name.Focus()
        End If
    End Sub

    Private Sub edt_name_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles edt_name.KeyUp
        If (e.KeyCode = Keys.Up) Then
            edt_code.Focus()
        End If
    End Sub
#End Region

End Class