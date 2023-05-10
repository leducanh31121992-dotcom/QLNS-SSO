Public Class frmChonDS_KhenThuong

    'Private _listCB As List(Of String)
    'Private _listTT As List(Of String)
    Private _list As New List(Of String())
    Private _list_tmp As New List(Of String())
    Private _Obj As Int16
    'Biến tạm, 
    Private KT_ChuyenMon As Int16 = 1
    'Khen thuong cua Chuyen mon: KT_ChuyenMon=1
    'Khen thuong cua Cong doan: KT_ChuyenMon=0
    'Private _CaNhan As Boolean
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler


    Public Property list() As List(Of String())
        Get
            Return _list
        End Get
        Set(ByVal value As List(Of String()))
            _list = value
        End Set
    End Property

    Public Property list_tmp() As List(Of String())
        Get
            Return _list_tmp
        End Get
        Set(ByVal value As List(Of String()))
            _list_tmp = value
        End Set
    End Property

    'Public Property CaNhan() As Boolean
    '    Get
    '        Return _CaNhan
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _CaNhan = value
    '    End Set
    'End Property

    Private Sub frmChonDS_KhenThuong_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub frmChonDS_KhenThuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmChonDS_KhenThuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDanhhieuTD.DataSource = listKhenThuong(rdCN.Checked, rdTT.Checked Or rdDV.Checked, KT_ChuyenMon, True)
        'bindTreeviewFull(treeCocau, True, False, False, False, False, True)
        treeCocau.ExpandAll()
        loadKhenThuongList(list, lstKhenThuong)
    End Sub

    Private Function blankList(ByVal listObj As List(Of String())) As List(Of String())
        If Not (listObj Is Nothing) Then listObj.Clear()
        Return listObj
    End Function

    Private Sub bntChoise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntChoise.Click
        ' add cac doi tuong duoc chon vao danh sach chung
        Dim i, j As Integer
        If list_tmp.Count > 0 Then
            ' neu la tap the: doc het cac danh sach CB vao arr(4)
            If _Obj = 2 Then
                Dim v_arr(6) As String
                Dim v_name As String = ""
                For i = 0 To list_tmp.Count - 1
                    v_arr = list_tmp.Item(i)
                    v_name += IIf(v_name = "", v_arr(4), "_" & v_arr(4))
                Next
                v_arr(3) = txtTenNhom.Text
                v_arr(4) = v_name
                v_arr(5) = cboDanhhieuTD.SelectedValue
                v_arr(6) = txtNoiDung_KT.Text
                If (list Is Nothing) Or (list.Count = 0) Then
                    list.Add(v_arr)
                Else
                    If Not (v_arr(0) = list.Item(j)(0) And v_arr(3) = list.Item(j)(3) And v_arr(4) = list.Item(j)(4) And cboDanhhieuTD.SelectedValue = list.Item(j)(5)) Then
                        list.Add(v_arr)
                    End If
                End If
            Else
                For i = 0 To list_tmp.Count - 1
                    Dim v_arr(6) As String
                    Dim v_name As String = ""
                    Dim isExist As Boolean = False

                    v_arr = list_tmp.Item(i)
                    If (list Is Nothing) Or (list.Count = 0) Then
                        v_arr(5) = cboDanhhieuTD.SelectedValue
                        v_arr(6) = txtNoiDung_KT.Text
                        list.Add(v_arr)
                    Else
                        For j = 0 To list.Count - 1
                            If v_arr(0) = list.Item(j)(0) And v_arr(1) = list.Item(j)(1) And v_arr(3) = list.Item(j)(3) And cboDanhhieuTD.SelectedValue = list.Item(j)(5) Then
                                isExist = True
                                Exit For
                            End If
                        Next
                        If isExist = False Then
                            v_arr(5) = cboDanhhieuTD.SelectedValue
                            v_arr(6) = txtNoiDung_KT.Text
                            list.Add(v_arr)
                        End If
                        'If list.Contains(v_arr) = False Then list.Add(v_arr)
                    End If
                Next
            End If
        End If

        'Hiển thị lên màn hình danh sách đã chọn
        lstKhenThuong.ClearSelected()
        lstKhenThuong.Items.Clear()
        For i = 0 To list.Count - 1
            Dim v_arr() As String
            Dim v_name As String = ""
            v_arr = list.Item(i)
            v_name = getDanhHieu(v_arr(5)) & ": "
            Select Case v_arr(0)
                Case 1
                    v_name += getFullNamePath(v_arr(3), False, v_arr(2), v_arr(1), IdDONVI, " _ ")
                Case 2
                    Dim v_arr_TV() As String
                    Dim k As Integer
                    v_name += v_arr(3)
                    v_arr_TV = v_arr(4).Split("_")
                    If v_arr_TV.Length > 0 Then
                        v_name += " ("
                        For k = 0 To v_arr_TV.Length - 1
                            v_name += getFullName(v_arr_TV(k), False) & ", "
                        Next
                        v_name = v_name.Substring(0, v_name.Length - 2) & ")"
                    End If
                Case 3
                    v_name += getPhongPath(v_arr(3), v_arr(1), IdDONVI, " _ ")
                Case 4
                    v_name += getDonviPath(v_arr(3), IdDONVI, " _ ")
            End Select
            lstKhenThuong.Items.Add(v_name)
        Next
        'Reset lại form
        rdDV.Checked = True
        treeCocau.Nodes.Clear()
        bindTreeviewFull(treeCocau, True, False, False, False, False, True)
        treeCocau.ExpandAll()
        blankList(list_tmp)
        txtTenNhom.Text = ""
        txtNoiDung_KT.Text = ""
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub treeCocau_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles treeCocau.NodeMouseClick
        If _Obj = 2 And txtTenNhom.Text.Trim = "" Then
            e.Node.Checked = False
            MessageBox.Show("Phải nhập tên nhóm trước khi chọn các thành viên trong nhóm!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            txtTenNhom.Focus()
            Exit Sub
        End If
        list_tmp = getDoiTuongKT(treeCocau.Nodes, _Obj)
    End Sub

    Private Function getDoiTuongKT(ByVal node As TreeNodeCollection, ByVal Obj As Integer) As List(Of String())
        Dim list As New List(Of String())
        Dim arr() As String
        For Each childNode As TreeNode In node
            If childNode.Checked Then
                Dim arrKT(6) As String
                arr = childNode.Tag.ToString.Split("_")
                arrKT(3) = childNode.Name
                arrKT(4) = ""
                arrKT(5) = "" 'cboDanhhieuTD.SelectedValue
                arrKT(6) = "" 'txtNoiDung_KT.Text
                Select Case arr(0)
                    Case "DV", "ROOT"
                        arrKT(0) = 4
                        arrKT(1) = arr(1) 'Đơn vị cấp trên trực tiếp của đối tượng
                        arrKT(2) = 0
                    Case "PB"
                        arrKT(0) = 3
                        arrKT(1) = arr(1) 'Đơn vị cấp trên trực tiếp của đối tượng
                        arrKT(2) = 0
                    Case "CB"
                        arrKT(1) = arr(3) 'Đơn vị cấp trên trực tiếp của đối tượng
                        arrKT(2) = childNode.Parent.Tag.ToString.Split("_")(2)
                        If _Obj = 2 Then
                            'các thành viên trong nhóm
                            arrKT(0) = 2
                            arrKT(3) = txtTenNhom.Text
                            arrKT(4) = childNode.Name
                        Else
                            arrKT(0) = 1
                        End If
                End Select
                If Obj < 3 Then
                    If Obj = CInt(arrKT(0)) Then list.Add(arrKT)
                Else
                    list.Add(arrKT)
                End If
                'If Obj = CInt(arrKT(0)) Then list.Add(arrKT)
            End If
            list.AddRange(getDoiTuongKT(childNode.Nodes, Obj))
        Next
        Return list
    End Function

    Private Sub loadKhenThuongList(ByVal vList As List(Of String()), ByRef objForm As ListBox)
        Dim i As Integer
        objForm.ClearSelected()
        objForm.Items.Clear()
        For i = 0 To vList.Count - 1
            Dim v_arr() As String
            Dim v_name As String = ""
            v_arr = vList.Item(i)
            v_name = getDanhHieu(v_arr(5)) & ": "
            Select Case v_arr(0)
                Case 1
                    v_name += getFullNamePath(v_arr(3), False, v_arr(2), v_arr(1), IdDONVI, " _ ")
                Case 2
                    Dim v_arr_TV() As String
                    Dim k As Integer
                    v_name = getDanhHieu(v_arr(5)) & ": "
                    v_name += v_arr(3)
                    v_arr_TV = v_arr(4).Split("_")
                    If v_arr_TV.Length > 0 Then
                        v_name += " ("
                        For k = 0 To v_arr_TV.Length - 1
                            v_name += getFullName(v_arr_TV(k), False) & ", "
                        Next
                        v_name = v_name.Substring(0, v_name.Length - 2) & ")"
                    End If
                Case 3
                    v_name += getPhongPath(v_arr(3), v_arr(1), IdDONVI, " _ ")
                Case 4
                    v_name += getDonviPath(v_arr(3), IdDONVI, " _ ")
            End Select
            objForm.Items.Add(v_name)
        Next
    End Sub

    Private Sub rdCN_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdCN.CheckedChanged
        If rdCN.Checked Then
            _Obj = 1
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, True, False, False, False, False)
            treeCocau.ExpandAll()
            cboDanhhieuTD.DataSource = listKhenThuong(True, False, KT_ChuyenMon, True)
            If Not (list_tmp Is Nothing) Then list_tmp.Clear()
            txtTenNhom.Enabled = False
        End If
    End Sub

    Private Sub rdDV_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdDV.CheckedChanged
        If rdDV.Checked Then
            _Obj = 3
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, True, False, False, False, False, True)
            treeCocau.ExpandAll()
            cboDanhhieuTD.DataSource = listKhenThuong(False, True, KT_ChuyenMon, True)
            If Not (list_tmp Is Nothing) Then list_tmp.Clear()
            txtTenNhom.Enabled = False
        End If
    End Sub

    'Private Sub rdPB_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If rdPB.Checked Then
    '        _Obj = 3
    '        treeCocau.Nodes.Clear()
    '        bindTreeviewFull(treeCocau, True, False, False, False, False)
    '        treeCocau.ExpandAll()
    '        cboDanhhieuTD.DataSource = listKhenThuong(False, True, True)
    '        If Not (list_tmp Is Nothing) Then list_tmp.Clear()
    '        txtTenNhom.Enabled = False
    '    End If
    'End Sub

    Private Sub rdTT_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdTT.CheckedChanged
        If rdTT.Checked Then
            _Obj = 2
            treeCocau.Nodes.Clear()
            bindTreeviewFull(treeCocau, True, False, False, False, False)
            treeCocau.ExpandAll()
            cboDanhhieuTD.DataSource = listKhenThuong(False, True, KT_ChuyenMon, True)
            If Not (list_tmp Is Nothing) Then list_tmp.Clear()
            txtTenNhom.Enabled = True
        End If
    End Sub

    Private Sub bntCancelKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancelKT.Click
        If Not (list Is Nothing) Then list.Clear()
        lstKhenThuong.ClearSelected()
        lstKhenThuong.Items.Clear()
    End Sub
End Class