Public Class frmHT_ThanhVien

#Region "---> Khai báo thuộc tính và khởi tạo đối tượng <---"
    Private _Systems As clsHeThong = New clsHeThong
    Private _Globals As Globals = New Globals
    Private _SqlHelper As DBAccess = New DBAccess()
    Private obj_update As frmHT_NguoiDung
    Private strNode As String = ""
    Private arr_GroupPermit As ArrayList = New ArrayList
    Private arr_Permits As ArrayList = New ArrayList
    Dim strSQL As String = ""
    Private strRoles As String = ""         'Biến lưu chuỗi giá trị tập quyền của nhóm thành viên

    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
#End Region

#Region "---> Functions: Các hàm dùng chung cho chương trình <---"
    Private Sub ResetAll_Controls()
        _Systems.Fill_Data(tv_main)

        clb_nhomquyen.Items.Clear()
        arr_GroupPermit.Clear()
        strSQL = String.Format("Select quyen,ten_quyen From HT_TapQuyen Where nhom_quyen = 0")
        arr_GroupPermit = _Globals.FillData_CheckedListBox(clb_nhomquyen, strSQL)

        ckb_nhquyen.Checked = False
        ckb_tapquyen.Checked = False
        _Globals.ResetItems_CheckedListBox(clb_nhomquyen)
        clb_tapquyen.Items.Clear()
        arr_Permits.Clear()
        strNode = ""
        strRoles = ""
    End Sub

    ''' <summary>
    ''' Hàm thực hiện add items danh sách tập quyền vào CheckedListBox
    ''' </summary>
    ''' <param name="_Items">Current item</param>
    ''' <remarks></remarks>
    Private Sub Add_Items(ByVal _Items As Integer)
        arr_Permits.Clear()
        clb_tapquyen.Items.Clear()
        'Thực hiện lấy chuỗi id các items được Checked trong List của Nhóm quyền
        Dim strItems As String = ""
        strItems = _Globals.GetItems_CheckedListBox(clb_nhomquyen, arr_GroupPermit)
        strItems = strItems + _Items.ToString()
        Dim arrId As ArrayList = New ArrayList()
        arrId = Globals.Splip_Strings(strItems)
        If (arrId.Count > 0) Then
            For i As Int16 = 0 To arrId.Count - 1
                'Lấy được Id của từng nhóm quyền được checked
                Using db As DataTable = _Systems.GetAll_Permits(CType(arrId(i), Int32))
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            If (clb_tapquyen.Items.Count <= 0) Then
                                For j As Integer = 0 To db.Rows.Count - 1
                                    clb_tapquyen.Items.Insert(j, db.Rows(j)("ten_quyen").ToString().Trim())
                                    arr_Permits.Insert(j, db.Rows(j)("quyen").ToString().Trim())
                                Next
                            Else
                                Dim count As Integer = clb_tapquyen.Items.Count + 1
                                For j As Integer = count To (db.Rows.Count + count - 1)
                                    clb_tapquyen.Items.Insert(j - 1, db.Rows(j - count)("ten_quyen").ToString().Trim())
                                    arr_Permits.Insert(j - 1, db.Rows(j - count)("quyen").ToString().Trim())
                                Next
                            End If
                        End If
                    End If
                End Using
            Next
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện remove item permit khi UnChecked nhóm quyền
    ''' </summary>
    ''' <param name="_item">Id nhóm quyền unchecked</param>
    ''' <remarks></remarks>
    Private Sub Remove_Items(ByVal _item As Integer)
        Dim arr_Items As ArrayList = New ArrayList()
        Using db As DataTable = _Systems.GetAll_Permits(_item)
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For j As Integer = 0 To db.Rows.Count - 1
                        arr_Items.Add(db.Rows(j)("quyen").ToString())
                    Next
                End If
            End If
        End Using
        'Thực hiện Kiểm tra so sánh giữa arr_Items với arr_Permits --> Nếu trùng thì remove
        If (arr_Permits.Count > 0 And arr_Items.Count > 0) Then
            For i As Integer = 0 To arr_Items.Count - 1
                For j As Integer = 0 To arr_Permits.Count - 1
                    If (arr_Items(i).ToString() = arr_Permits(j).ToString()) Then
                        arr_Permits.RemoveAt(j)
                        clb_tapquyen.Items.RemoveAt(j)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub

    ''' <summary>
    ''' Hàm thực hiện set items của nhóm thành viên - Thực chất là Checked vào các item tập quyền của thành viên
    ''' </summary>
    ''' <param name="_GroupUser"></param>
    ''' <remarks></remarks>
    Private Sub SetItems_Permits(ByVal _GroupUser As String)
        If (_GroupUser <> "") Then
            Dim dr As DataRow
            dr = _Systems.GetPermit_GroupUser(_GroupUser)
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    'Lấy được chuôi giá trị tập quyền của nhóm thành viên
                    Dim strPermits = dr("tap_quyen").ToString().Trim()
                    'Tìm cách lấy được chuỗi Nhóm quyền của Nhóm thành viên này ?
                    Dim strPermitGroup As String = _Systems.GetString_GroupPermit(strPermits)
                    'Bắt đầu Set các items checked của CheckListBox
                    _Globals.SetItems_CheckedListBox(clb_nhomquyen, strPermitGroup, arr_GroupPermit)
                    _Globals.SetItems_CheckedListBox(clb_tapquyen, strPermits, arr_Permits)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Hàm kiểm tra dữ liệu hợp lệ trước khi cập nhật
    ''' </summary>
    ''' <param name="strPer">Chuỗi quyền nhóm thành viên</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValid(ByVal strPer As String) As Boolean
        If (strPer = "") Then
            MessageBox.Show("Bạn chưa chọn quyền cho nhóm thành viên!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = clb_tapquyen
            Return False
        End If
        If (strNode = "") Then
            MessageBox.Show("Bạn chưa chọn nhóm thành viên cần phân quyền!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = tv_main
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' Hàm kiểm tra thông tin quyền đối với Hệ thống Danh mục
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        'Nhóm thành viên
        If Not (Globals.IsIntersect(";45;46;47;48;", _roles)) Then
            btn_grmember.Visible = False
        End If
        If (_roles.IndexOf(";54;") < 0) Then
            btn_accept.Visible = False
        End If
    End Sub

#End Region

#Region "---> Events: Các sự kiện chính <---"
    Private Sub tv_main_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tv_main.AfterSelect
        Cursor = Cursors.WaitCursor
        ckb_nhquyen.Checked = False
        ckb_tapquyen.Checked = False
        clb_tapquyen.Items.Clear()
        arr_Permits.Clear()
        strNode = ""
        strRoles = ""
        _Globals.ResetItems_CheckedListBox(clb_nhomquyen)
        If (tv_main.Nodes.Count > 0) Then
            If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "00") Then
                strNode = tv_main.SelectedNode.Tag.ToString().Trim()
                'Using db As DataTable = _Systems.GetAll_Members(strNode.Substring(3).ToString().Trim())
                '    If Not (db Is Nothing) Then
                '        If (db.Rows.Count > 0) Then
                '            For i As Integer = 0 To db.Rows.Count - 1
                '                dgv_main.Rows.Add()
                '                dgv_main.Rows(i).Cells("cln_UserName").Value = db.Rows(i)("ten_tv").ToString().Trim()
                '                dgv_main.Rows(i).Cells("cln_FullName").Value = db.Rows(i)("ho_ten").ToString().Trim()
                '                dgv_main.Rows(i).Cells("cln_Email").Value = db.Rows(i)("email").ToString().Trim()
                '                If (db.Rows(i)("ngay_sua").ToString().Trim() <> "") Then
                '                    dgv_main.Rows(i).Cells("cln_ModifiedDate").Value = CType(db.Rows(i)("ngay_sua").ToString(), DateTime).ToString("dd-MM-yyyy")
                '                Else
                '                    dgv_main.Rows(i).Cells("cln_ModifiedDate").Value = ""
                '                End If

                '                If (db.Rows(i)("ngay_tao").ToString().Trim() <> "") Then
                '                    dgv_main.Rows(i).Cells("cln_CreationDate").Value = CType(db.Rows(i)("ngay_tao").ToString(), DateTime).ToString("dd-MM-yyyy")
                '                Else
                '                    dgv_main.Rows(i).Cells("cln_CreationDate").Value = ""
                '                End If
                '                dgv_main.Rows(i).Cells("cln_Note").Value = db.Rows(i)("ghi_chu").ToString().Trim()
                '            Next
                '        End If
                '    End If
                'End Using
            End If
            'If (tv_main.SelectedNode.Tag.ToString().Substring(0, 2) = "01") Then
            '    strNode = tv_main.SelectedNode.Parent.Tag.ToString()
            '    Dim Code As String = tv_main.SelectedNode.Tag.ToString().Substring(3)
            'Dim dr As DataRow
            'dr = _Systems.GetMember(Code)
            'If Not (dr Is Nothing) Then
            '    If (dr.Table.Rows.Count > 0) Then
            '        dgv_main.Rows.Add()
            '        dgv_main.Rows(0).Cells("cln_UserName").Value = dr("ten_tv").ToString().Trim()
            '        dgv_main.Rows(0).Cells("cln_FullName").Value = dr("ho_ten").ToString().Trim()
            '        dgv_main.Rows(0).Cells("cln_Email").Value = dr("email").ToString().Trim()
            '        If (dr("ngay_sua").ToString().Trim() <> "") Then
            '            dgv_main.Rows(0).Cells("cln_ModifiedDate").Value = CType(dr("ngay_sua").ToString(), DateTime).ToString("dd-MM-yyyy")
            '        Else
            '            dgv_main.Rows(0).Cells("cln_ModifiedDate").Value = ""
            '        End If

            '        If (dr("ngay_tao").ToString().Trim() <> "") Then
            '            dgv_main.Rows(0).Cells("cln_CreationDate").Value = CType(dr("ngay_tao").ToString(), DateTime).ToString("dd-MM-yyyy")
            '        Else
            '            dgv_main.Rows(0).Cells("cln_CreationDate").Value = ""
            '        End If
            '        dgv_main.Rows(0).Cells("cln_Note").Value = dr("ghi_chu").ToString().Trim()
            '    End If
            'End If
            'End If
        End If

        'Set item quyền nhóm thành viên

        If (strNode <> "") Then
            Dim dr As DataRow
            dr = _Systems.GetPermit_GroupUser(strNode.Substring(3))
            If Not (dr Is Nothing) Then
                If (dr.Table.Rows.Count > 0) Then
                    clb_tapquyen.BeginUpdate()

                    'Lấy được chuôi giá trị tập quyền của nhóm thành viên
                    strRoles = dr("tap_quyen").ToString().Trim()
                    'Tìm cách lấy được chuỗi Nhóm quyền của Nhóm thành viên này ?
                    Dim strPermitGroup As String = _Systems.GetString_GroupPermit(strRoles)
                    'Bắt đầu Set các items checked của CheckListBox
                    _Globals.SetItems_CheckedListBox(clb_nhomquyen, strPermitGroup, arr_GroupPermit)
                    _Globals.SetItems_CheckedListBox(clb_tapquyen, strRoles, arr_Permits)

                    clb_tapquyen.EndUpdate()
                End If
            End If
        End If
        Cursor = Cursors.Default
    End Sub

    Private Sub frmHT_ThanhVien_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '_Systems.Create_Frame(dgv_main, True)
        'Fill dữ liệu Nhóm quyền của các Thành viên
        ResetAll_Controls()
        Check_Permits()
    End Sub

    Private Sub btn_grmember_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_grmember.Click
        Dim obj_grmembers As New frmHT_NhomThVien
        obj_grmembers.Progress_Changed = New frmHT_NhomThVien.ProgressChangedEventHandler(AddressOf ReLoad)
        obj_grmembers.ShowDialog()
    End Sub

    Private Sub ReLoad()
        ResetAll_Controls()
    End Sub

    Private Sub btn_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        obj_update = New frmHT_NguoiDung
        obj_update.FlagEvent = 1
        obj_update.Node = strNode
        obj_update.Progress_Changed = New frmHT_NguoiDung.ProgressChangedEventHandler(AddressOf ReLoad_Members)
        obj_update.ShowDialog()
    End Sub

    Private Sub ReLoad_Members()
        Try
            ResetAll_Controls()
            'Tìm lại Node trên treeview hiện hành trước khi cập nhật
            Dim node As TreeNode = Nothing
            If (obj_update.Node.ToString().Trim() <> "") Then
                node = Globals.TreeViewFindNode(tv_main.Nodes, obj_update.Node.ToString())
            End If

            If Not (node Is Nothing) Then
                tv_main.SelectedNode = node
            End If
            obj_update.Dispose()
        Catch ex As Exception
            MessageBox.Show("Lỗi xẩy ra: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub ckb_nhquyen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_nhquyen.CheckedChanged
        Globals.CheckAll_CheckListBox(ckb_nhquyen, clb_nhomquyen)
        If (ckb_nhquyen.Checked = False) Then
            ckb_tapquyen.Checked = False
        End If
    End Sub

    Private Sub ckb_tapquyen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckb_tapquyen.CheckedChanged
        Globals.CheckAll_CheckListBox(ckb_tapquyen, clb_tapquyen)
    End Sub

    Private Sub clb_nhomquyen_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles clb_nhomquyen.ItemCheck
        If (clb_nhomquyen.Items.Count > 0) Then
            If (e.NewValue = CheckState.Checked) Then
                'Lấy chuỗi quyền của các items đang được checked - trong danh sách tập quyền của nhóm thành viên
                Dim strValue As String = ""
                If (clb_tapquyen.Items.Count > 0) Then
                    strValue = _Globals.GetItems_CheckedListBox(clb_tapquyen, arr_Permits)
                End If
                Add_Items(CType(arr_GroupPermit(e.Index).ToString(), Integer))
                _Globals.ResetItems_CheckedListBox(clb_tapquyen)
                'Thực hiện Checked lại nếu chuỗi strValue có giá trị
                If (strValue <> "") Then
                    _Globals.SetItems_CheckedListBox(clb_tapquyen, strValue, arr_Permits)
                End If
            Else    'Thực hiện remove items nếu UnChecked
                Dim valGroupPermit As Integer = CType(arr_GroupPermit(e.Index).ToString(), Integer)
                'Thực hiện remove item nhóm quyền (id nhóm quyền) khỏi mảng lưu danh sách nhóm quyền
                Remove_Items(CType(arr_GroupPermit(e.Index).ToString(), Integer))
                'Thực hiện bỏ check tất các item trong checklistbox tập quyền
                _Globals.ResetItems_CheckedListBox(clb_tapquyen)
                'Thực hiện remove id tập quyền trong chuỗi theo nhóm quyền vừa remove.
                Using db As DataTable = _Systems.GetAll_Permits(valGroupPermit)
                    If Not (db Is Nothing) Then
                        If (db.Rows.Count > 0) Then
                            Dim strTemp As String = ""
                            'Cộng chuỗi danh sách tập quyền với dấu (;) ở trước và sau chuỗi
                            If (strRoles <> "") Then
                                If (strRoles.Substring(0, 1) = ";") Then
                                    strTemp = strRoles
                                Else
                                    strTemp = ";" + strRoles
                                End If
                            End If
                            'Tìm và remove tập quyền thuộc nhóm quyền đã unchecked
                            For i As Integer = 0 To db.Rows.Count - 1
                                'Dim strId As String = db.Rows(i)("quyen").ToString() + ";"
                                Dim strId As String = ";" + db.Rows(i)("quyen").ToString() + ";"
                                If (strTemp.IndexOf(strId) >= 0) Then
                                    strTemp = strTemp.Remove(strTemp.IndexOf(strId) + 1, strId.Length - 1)
                                    If (strTemp <> "") Then
                                        If (strTemp.Substring(0, 1) <> ";") Then
                                            strTemp = ";" + strTemp
                                        End If
                                    End If
                                End If
                            Next
                            strRoles = strTemp
                        End If
                    End If
                End Using
                'Kiểm tra xem cuối chuỗi đã có dấu chấm phẩy chưa
                If (strRoles <> "") Then
                    If (strRoles.EndsWith(";") = False) Then
                        strRoles = strRoles + ";"
                    End If
                End If
                'Checked tất các items trong checklistbox tập quyền
                If (clb_tapquyen.Items.Count > 0 And strRoles <> "") Then
                    _Globals.SetItems_CheckedListBox(clb_tapquyen, strRoles.Trim(), arr_Permits)
                End If
            End If
        End If
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        Dim _Node As String = strNode
        strNode = ""
        'dgv_main.Rows.Clear()
        ckb_nhquyen.Checked = False
        ckb_tapquyen.Checked = False
        _Globals.ResetItems_CheckedListBox(clb_nhomquyen)
        clb_tapquyen.Items.Clear()
        arr_Permits.Clear()
        strRoles = ""

        tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
        'Select item của cây dữ liệu
        Dim node As TreeNode = Nothing
        If (_Node <> "") Then
            node = Globals.TreeViewFindNode(tv_main.Nodes, _Node)
            If Not (node Is Nothing) Then
                tv_main.SelectedNode = node
            End If
        End If
        Check_Permits()
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Close()
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        Try
            Dim strPer As String = _Globals.GetItems_CheckedListBox(clb_tapquyen, arr_Permits)
            If (IsValid(strPer)) Then
                Dim _flag As Boolean = False
                Dim strGrUser As String = strNode.Substring(3).Trim()
                Dim dr As DataRow
                dr = _Systems.GetPermit_GroupUser(strGrUser)
                If Not (dr Is Nothing) Then
                    If (dr.Table.Rows.Count > 0) Then
                        _flag = True
                    End If
                End If
                Dim obj_cofig As clsHeThong.PermitGrMember = New clsHeThong.PermitGrMember()
                obj_cofig.nhom_tv = strGrUser
                obj_cofig.tap_quyen = strPer
                obj_cofig.ngay_tao = DateTime.Now
                If (_flag = True) Then  'Sửa đổi
                    _Systems.Update_PermitGrMember(obj_cofig)
                Else
                    'Thêm mới
                    _Systems.Insert_PermitGrMember(obj_cofig)
                End If
                MessageBox.Show("Bạn đã thiết lập thành công quyền của nhóm thành viên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Dim _Node As String = strNode
                strNode = ""
                'dgv_main.Rows.Clear()
                ckb_nhquyen.Checked = False
                ckb_tapquyen.Checked = False
                _Globals.ResetItems_CheckedListBox(clb_nhomquyen)
                clb_tapquyen.Items.Clear()
                arr_Permits.Clear()
                strRoles = ""

                'Cuộn hết các nodes của cây dữ liệu lên
                tv_main.CollapseAll()
                'Select item của cây dữ liệu
                Dim node As TreeNode = Nothing
                If (_Node <> "") Then
                    node = Globals.TreeViewFindNode(tv_main.Nodes, _Node)
                    If Not (node Is Nothing) Then
                        tv_main.SelectedNode = node
                    End If
                End If

                'Lấy lại chuỗi quyền của thành viên đang thao tác chương trình
                Globals.Roles = _Systems.GetRoles(Globals.UserVal)
            End If
        Catch ex As Exception
            MessageBox.Show("Thiết lập quyền cho nhóm thành viên: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub frmHT_ThanhVien_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub
#End Region
End Class