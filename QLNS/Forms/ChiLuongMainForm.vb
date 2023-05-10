Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Syncfusion.XlsIO
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Collections
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class ChiLuongMainForm
    Private _ChiLuongBLL As ChiLuongBLL = New ChiLuongBLL
    Private ARL_LoaiLuong As ArrayList = New ArrayList
    Private _Globals As Globals = New Globals

    Private ckb_ChoiceAll As CheckBox = Nothing                     'Control CheckBox
    Dim TotalCheckBoxes As Integer = 0                              'Tổng số bản ghi trên lưới dữ liệu
    Dim TotalCheckedCheckBoxes As Integer = 0                       'Tổng số các items đang Checked trên lưới dữ liệu
    Dim IsHeaderCheckBoxClicked As Boolean = False                  'Cờ báo việc Checkall
    Dim _ClassNumber As Integer = 6

    Dim _BranchTag As String = ""
    Dim _NodeTag As String = ""
    Private _NodeCurrent As String = ""

    Private chiluongobj As ChiLuongForm
    Private vNFInfo As System.Globalization.NumberFormatInfo

    Private IsCall As Boolean = False
    Private _IndexFlag As Byte = 0

    Private sMain_Choise As String = ""
    Private sPos_Choise As String = ""
    Private sDepartment_Choise As String = ""
    Private sMonth_Choise As String = ""
    Private sProvinceName As String = ""
    Private sBranchName = ""

    Private bFlagCall As Byte = 0

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        vNFInfo = New System.Globalization.NumberFormatInfo()
        'vNFInfo.NumberDecimalDigits = 2
        'vNFInfo.NumberGroupSeparator = " "
    End Sub

    Private Sub AddHeaderCheckBox()
        ckb_ChoiceAll = New CheckBox()
        ckb_ChoiceAll.Size = New Size(15, 15)
        'Add the CheckBox into the DataGridView
        ckb_ChoiceAll.Checked = False
        dgv_main.Controls.Add(ckb_ChoiceAll)
    End Sub

    '230	0	Chi lương định kỳ
    '231	230	Xem thông tin Chi lương
    '232	230	Thêm mới Chi lương
    '233	230	Sửa đổi Chi lương
    '234	230	Xóa bỏ Chi lương
    '235	230	Phê duyệt Chi lương
    Private Sub Check_Permits()
        Dim _roles As String = Globals.Roles
        If Not (Globals.IsIntersect(";231;232;233;234;235;", _roles)) Then
            tv_main.Visible = False
            btn_pheduyet.Visible = False
            btn_themmoi.Visible = False
            btn_timkiem.Visible = False
            btn_suadoi.Visible = False
            btn_xoabo.Visible = False
            btn_export_excel.Visible = False
            pnl_search.Visible = False
            MessageBox.Show("Bạn không được phân quyền thực hiện chức năng này. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Close()
        End If
        'Nếu không có quyền Xem thông tin
        If (Globals.Roles.IndexOf(";231;") < 0) Then
            tv_main.Visible = False
        End If
        'Nếu không có quyền thêm mới
        If (Globals.Roles.IndexOf(";232;") < 0) Then
            btn_themmoi.Visible = False
        End If
        'Nếu không có quyền sửa đổi
        If (Globals.Roles.IndexOf("233;") < 0) Then
            btn_suadoi.Visible = False
        End If
        'Nếu không có quyền xóa bỏ
        If (Globals.Roles.IndexOf("234;") < 0) Then
            btn_xoabo.Visible = False
        End If
        'Nếu không có quyền phê duyệt
        If (Globals.Roles.IndexOf("235;") < 0) Then
            btn_pheduyet.Visible = False
        End If
    End Sub
    'Globals.Group
    Private Sub ChiLuongMainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ARL_LoaiLuong.Clear()
        cb_loaichiluong.Items.Clear()
        ARL_LoaiLuong = _Globals.BindList_ComBoBox_ListBox(cb_loaichiluong, Nothing, IIf(DONVI = "000100" Or DONVI = "000199" Or DONVI = "000196", clsHT_DanhMuc.Sql_LoaiChiLuongALL, clsHT_DanhMuc.Sql_LoaiChiLuongBoEOD), "--- Phân loại ---", 1)
        cb_loaichiluong.SelectedIndex = 0
        rb_tatca.Checked = True
        rb_daihan.Checked = False
        rb_nganhan.Checked = False
        rb_tapsu.Checked = False
        num_kybc.Value = 0
        num_thangbc.Value = 0
        Dim _NgayTMP As DateTime
        _NgayTMP = Globals.GetDateTime_ForServerDB
        num_nambc.Value = CType(_NgayTMP.Year.ToString(), Integer)
        num_nambc_ValueChanged(sender, Nothing)
        dgv_main_CellContentClick(sender, Nothing)

        _ChiLuongBLL.Create_Frame(dgv_main, 10)
        AddHeaderCheckBox()
        Me.dgv_main.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        Me.dgv_main.ColumnHeadersHeight = Me.dgv_main.ColumnHeadersHeight * _ClassNumber
        Me.dgv_main.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        AddHandler Me.dgv_main.CellPainting, AddressOf dgv_main_CellPainting

        AddHandler Me.dgv_main.Paint, AddressOf dgv_main_Paint
        AddHandler Me.dgv_main.Scroll, AddressOf dgv_main_Scroll
        AddHandler Me.dgv_main.ColumnWidthChanged, AddressOf dgv_main_ColumnWidthChanged

        AddHandler ckb_ChoiceAll.KeyUp, AddressOf Me.ckb_ChoiceAll_KeyUp
        AddHandler ckb_ChoiceAll.MouseClick, AddressOf Me.ckb_ChoiceAll_MouseClick
        AddHandler dgv_main.CellValueChanged, AddressOf dgv_main_CellValueChanged

        AddHandler dgv_main.CurrentCellDirtyStateChanged, AddressOf dgv_main_CurrentCellDirtyStateChanged
        Check_Permits()
        'tv_main_AfterSelect(sender, Nothing)
    End Sub

    Private Sub btn_timkiem_Click(sender As Object, e As EventArgs) Handles btn_timkiem.Click

    End Sub

    Private Sub btn_themmoi_Click(sender As Object, e As EventArgs) Handles btn_themmoi.Click
        chiluongobj = New ChiLuongForm()
        chiluongobj.ValUpdate = 1
        _IndexFlag = 1
        chiluongobj.CLTongHopId = 0
        Frm_FitSizeToParent(chiluongobj)
        chiluongobj.Progress_Changed = New ChiLuongForm.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
        chiluongobj.ShowDialog()
    End Sub

    Private Sub btn_suadoi_Click(sender As Object, e As EventArgs) Handles btn_suadoi.Click
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_TongHopId").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "" And dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "0")) Then
                If dgv_main.CurrentRow.Cells("cln_TrangThai").Value.ToString() = "2" Then
                    MessageBox.Show("Bản ghi Chi lương cần sửa đã được phê duyệt, không thực hiện chỉnh sửa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Else
                    chiluongobj = New ChiLuongForm()
                    chiluongobj.ValUpdate = 2
                    _IndexFlag = 2
                    chiluongobj.CLTongHopId = dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString()
                    Frm_FitSizeToParent(chiluongobj)
                    chiluongobj.Progress_Changed = New ChiLuongForm.ProgressChangedEventHandler(AddressOf ReLoad_Infors)
                    chiluongobj.ShowDialog()
                End If
            Else
                MessageBox.Show("Bạn chưa chọn dữ liệu chi lương cần sửa đổi!" + vbCrLf + "Lưu ý: Bản ghi được phép sửa là bản ghi chi tiết và chưa phê duyệt.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If
        Else
            MessageBox.Show("Bạn chưa chọn dữ liệu chi lương cần sửa đổi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub btn_xoabo_Click(sender As Object, e As EventArgs) Handles btn_xoabo.Click
        If (dgv_main.Rows.Count <= 0) Then Return
        Try
            Dim ARL_Choices As ArrayList = New ArrayList()
            Dim _Counts As Integer = 0
            Dim _CountAuthors As Integer = 0
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_TongHopId").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "" And dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "0")) Then
                    If dgv_main.CurrentRow.Cells("cln_KieuIn").Value IsNot Nothing And CType(dgv_main.CurrentRow.Cells("cln_KieuIn").Value, Byte) = 3 Then
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If (dgv_main.Rows(i).Cells("cln_TongHopId").Value IsNot Nothing) Then
                                If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                    If dgv_main.Rows(i).Cells("cln_KieuIn").Value.ToString() = "3" And dgv_main.Rows(i).Cells("cln_TongHopId").Value.ToString() <> "" And dgv_main.Rows(i).Cells("cln_TongHopId").Value.ToString() <> "0" And dgv_main.Rows(i).Cells("cln_TrangThai").Value.ToString() <> "2" Then
                                        _Counts += 1
                                        ARL_Choices.Add(dgv_main.Rows(i).Cells("cln_TongHopId").Value.ToString())
                                    ElseIf dgv_main.Rows(i).Cells("cln_TrangThai").Value.ToString() <> "2" Then
                                        _CountAuthors += 1
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            If (_Counts > 0) Then
                Dim _Mess As String = IIf(_Counts = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn xóa " + _Mess + "bản ghi Chi lương đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    Dim _TongHopIdPd As Long = 0
                    _Counts = 0
                    For i As Int16 = 0 To ARL_Choices.Count - 1
                        _TongHopIdPd = CType(ARL_Choices(i).ToString(), Long)
                        _ChiLuongBLL.Delete_UpdateStatus_ChiLuong_TongHop(_TongHopIdPd, 0, "", "", "", "", "", 0, 0, 0, "", 4)
                        _Counts += 1
                    Next
                    MessageBox.Show("Bạn đã xóa thành công " + _Counts.ToString() + " bản ghi Chi lương đã chọn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

                    'Thực hiện select lại treeview dữ liệu Cơ cấu tổ chức
                    Dim _TagSelect As String = ""
                    Dim _TagParentSelect As String = ""
                    _TagSelect = tv_main.SelectedNode.Tag.ToString().Trim()
                    If (_NodeCurrent.Substring(0, 2) = "TG") Then
                        '_TagSelect = tv_main.SelectedNode.Parent.Tag.ToString().Trim()
                        _TagParentSelect = tv_main.SelectedNode.Parent.Tag.ToString().Trim()
                    Else : _TagParentSelect = _TagSelect
                    End If

                    chiluongobj.bCoLuuDL = True
                    ReLoad_Infors()
                    tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên

                    Dim _IsNode As TreeNode = Globals.TreeViewFindNode(tv_main.Nodes, _TagParentSelect)
                    If Not (_IsNode Is Nothing) Then
                        If _IsNode.Nodes.Count() <= 0 Then
                            tv_main.SelectedNode = _IsNode
                            tv_main.SelectedNode.ForeColor = Color.Blue
                            tv_main.SelectedNode.Expand()
                        Else
                            Dim _IsNodeChild As TreeNode = Globals.TreeViewFindNode(tv_main.Nodes, _TagSelect)
                            If Not (_IsNodeChild Is Nothing) Then
                                tv_main.SelectedNode = _IsNodeChild
                                tv_main.SelectedNode.ForeColor = Color.Maroon
                                tv_main.SelectedNode.Expand()
                            End If
                        End If
                        tv_main_AfterSelect(sender, Nothing)
                    End If
                Else
                    ARL_Choices.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            Else
                If _CountAuthors <= 0 Then
                    MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xóa, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Else
                    MessageBox.Show("Bạn không thể xóa các bản ghi đã phê duyệt, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    ARL_Choices.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi Xóa chi lương: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi Xóa chi lương ChiLuongMainForm\btn_xoabo_Click(): " + ex.Message)
        End Try

    End Sub

    Private Sub btn_pheduyet_Click(sender As Object, e As EventArgs) Handles btn_pheduyet.Click
        If (dgv_main.Rows.Count <= 0) Then Return
        Try
            Dim ARL_Choices As ArrayList = New ArrayList()
            Dim _Counts As Integer = 0
            Dim _CountAuthors As Integer = 0
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_TongHopId").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "" And dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "0")) Then
                    If dgv_main.CurrentRow.Cells("cln_KieuIn").Value IsNot Nothing And CType(dgv_main.CurrentRow.Cells("cln_KieuIn").Value, Byte) = 3 Then
                        For i As Integer = 0 To dgv_main.Rows.Count - 1
                            If (dgv_main.Rows(i).Cells("cln_TongHopId").Value IsNot Nothing) Then
                                If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                                    If dgv_main.Rows(i).Cells("cln_KieuIn").Value.ToString() = "3" And dgv_main.Rows(i).Cells("cln_TongHopId").Value.ToString() <> "" And dgv_main.Rows(i).Cells("cln_TongHopId").Value.ToString() <> "0" And dgv_main.Rows(i).Cells("cln_TrangThai").Value.ToString() <> "2" Then
                                        _Counts += 1
                                        ARL_Choices.Add(dgv_main.Rows(i).Cells("cln_TongHopId").Value.ToString())
                                    ElseIf dgv_main.Rows(i).Cells("cln_TrangThai").Value.ToString() <> "2" Then
                                        _CountAuthors += 1
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            If (_Counts > 0) Then
                Dim _Mess As String = IIf(_Counts = 1, "", "các ")
                If (MessageBox.Show("Bạn có thực sự muốn phê duyệt " + _Mess + "bản ghi Chi lương đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                    Dim _TongHopIdPd As Long = 0
                    _Counts = 0
                    For i As Int16 = 0 To ARL_Choices.Count - 1
                        _TongHopIdPd = CType(ARL_Choices(i).ToString(), Long)
                        _ChiLuongBLL.Delete_UpdateStatus_ChiLuong_TongHop(_TongHopIdPd, 2, Globals.UserVal, "", "", "", "", 0, 0, 0, "", 5)
                        _Counts += 1
                    Next
                    MessageBox.Show("Bạn đã phê duyệt thành công " + _Counts.ToString() + " bản ghi Chi lương đã chọn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    tv_main_AfterSelect(sender, Nothing)
                Else
                    ARL_Choices.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            Else
                If _CountAuthors <= 0 Then
                    MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần phê duyệt, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Else
                    MessageBox.Show("Bạn đánh dấu các bản ghi đã được phê duyệt, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    ARL_Choices.Clear()
                    ckb_ChoiceAll.Checked = False
                    HeaderCheckBoxClick(ckb_ChoiceAll, dgv_main)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi Phê duyệt chi lương: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Globals.Logger.Error("Lỗi Phê duyệt chi lương ChiLuongMainForm\btn_pheduyet_Click(): " + ex.Message)
        End Try
    End Sub


    Private Sub btn_quayra_Click(sender As Object, e As EventArgs) Handles btn_quayra.Click
        Close()
    End Sub

    Private Sub ReLoad_Infors()
        Try
            If chiluongobj.bCoLuuDL Then
                Dim _TagSelect As String = ""
                Dim _TagParentSelect As String = ""
                If Not (tv_main.SelectedNode Is Nothing) Then
                    _TagSelect = tv_main.SelectedNode.Tag.ToString().Trim()
                    If (_NodeCurrent.Substring(0, 2) = "TG") Then
                        _TagParentSelect = tv_main.SelectedNode.Parent.Tag.ToString().Trim()
                    Else : _TagParentSelect = _TagSelect
                    End If
                End If

                IsCall = False
                num_nambc_ValueChanged(Nothing, Nothing)
                tv_main.CollapseAll() 'Cuộn hết các nodes của cây dữ liệu lên
                If _IndexFlag = 1 Or _IndexFlag = 2 Then  'Reload khi thêm mới/Sửa đổi
                    Dim _IsNodeChild As TreeNode = Globals.TreeViewFindNode(tv_main.Nodes, _TagSelect)
                    If Not (_IsNodeChild Is Nothing) Then
                        tv_main.SelectedNode = _IsNodeChild
                        tv_main.SelectedNode.ForeColor = Color.Maroon
                        tv_main.SelectedNode.Expand()
                        tv_main_AfterSelect(Nothing, Nothing)
                    End If
                End If
                _IndexFlag = 0
            End If
        Catch ex As Exception
            Globals.Logger.Error("Lỗi ChiLuongMainForm\ReLoad_Infors(): " + ex.Message)
            MessageBox.Show("Lỗi Reload lại màn hình sau sự kiện ChiLuongMainForm\ReLoad_Infors(): " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

        End Try
    End Sub

#Region "---> Hàm, sự kiện liên quan tới Merge lưới và CheckAll lưới dữ liệu <---"
    Private Sub dgv_main_ColumnWidthChanged(ByVal sender As Object, ByVal e As DataGridViewColumnEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_main.DisplayRectangle
        rtHeader.Height = Me.dgv_main.ColumnHeadersHeight / 2
        Me.dgv_main.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_main_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
        Dim rtHeader As Rectangle = Me.dgv_main.DisplayRectangle
        rtHeader.Height = Me.dgv_main.ColumnHeadersHeight / 3
        Me.dgv_main.Invalidate(rtHeader)
    End Sub

    Private Sub dgv_main_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs)
        If e.RowIndex = -1 AndAlso e.ColumnIndex > -1 Then
            Dim r2 As Rectangle = e.CellBounds
            r2.Y += e.CellBounds.Height / 3
            r2.Height = e.CellBounds.Height / 3
            e.PaintBackground(r2, True)
            e.PaintContent(r2)
            e.Handled = True
        End If

        'Căn chỉnh cho Checkbox nằm vào giữa tiêu đề của cột Chọn xoá
        If (e.RowIndex = -1 And e.ColumnIndex = 1) Then
            ResetHeaderCheckBoxLocation(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    Private Sub dgv_main_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim format As StringFormat = New StringFormat()
        format.Alignment = StringAlignment.Center
        format.LineAlignment = StringAlignment.Center
        Dim rTienLuong As Rectangle = New Rectangle()
        Dim rHeSoLuong As Rectangle = New Rectangle()
        Dim rTrongDo As Rectangle = New Rectangle()
        Dim rPhuCap As Rectangle = New Rectangle()

        Dim rTienCong_NH As Rectangle = New Rectangle()
        Dim rTienCong As Rectangle = New Rectangle()
        Dim rDinhBien As Rectangle = New Rectangle()
        Dim rPhuTro As Rectangle = New Rectangle()

        For j As Integer = 0 To dgv_main.ColumnCount - 1 Step 1
            If j = 7 Then       'TIỀN LƯƠNG V1 (100%) (Tiền lương CBCMNV)
                rTienLuong = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                Dim w6 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 6, -1, True).Width
                Dim w7 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 7, -1, True).Width
                Dim w8 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 8, -1, True).Width
                Dim w9 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 9, -1, True).Width
                Dim w10 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 10, -1, True).Width
                rTienLuong.X += 1
                rTienLuong.Y += 1
                rTienLuong.Width = rTienLuong.Width + w1 + w2 + w3 + w4 + w5 + w6 + w7 + w8 + w9 + w10 - 1

                rTienLuong.Height = (rTienLuong.Height / _ClassNumber - 2) '* 2
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTienLuong)
                e.Graphics.DrawRectangle(Pens.Silver, rTienLuong)
                e.Graphics.DrawString("TIỀN LƯƠNG V1 (100%) (Tiền lương CBCMNV)", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTienLuong, format)
            End If

            If j = 8 Then       'Hệ số lương
                rHeSoLuong = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                rHeSoLuong.X += 1
                rHeSoLuong.Y += rTienLuong.Height + 1
                rHeSoLuong.Width = rHeSoLuong.Width + w1 + w2 + w3 + w4 - 1
                rHeSoLuong.Height = rHeSoLuong.Height / _ClassNumber - 2 '7 lop                    
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rHeSoLuong)
                e.Graphics.DrawRectangle(Pens.Silver, rHeSoLuong)
                e.Graphics.DrawString("Hệ số lương", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rHeSoLuong, format)
            End If

            If j = 9 Then       'Trong đó
                rTrongDo = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                rTrongDo.X += 1
                rTrongDo.Y += rTienLuong.Height + rHeSoLuong.Height + 1
                rTrongDo.Width = rTrongDo.Width + w1 + w2 + w3 - 1
                rTrongDo.Height = rTrongDo.Height / _ClassNumber - 2         '3 la vi 3 lop

                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTrongDo)
                e.Graphics.DrawRectangle(Pens.Silver, rTrongDo)
                e.Graphics.DrawString("Trong đó", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTrongDo, format)
            End If
            If j = 10 Then       'Phụ cấp lương
                rPhuCap = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width

                rPhuCap.X += 1
                rPhuCap.Y += rTienLuong.Height + rHeSoLuong.Height + rTrongDo.Height + 1
                rPhuCap.Width = rPhuCap.Width + w1 + w2 - 1
                rPhuCap.Height = (rPhuCap.Height / _ClassNumber - 2) * 1
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCap)
                e.Graphics.DrawRectangle(Pens.Silver, rPhuCap)
                e.Graphics.DrawString("Phụ cấp", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCap, format)
            End If
            If j = 15 Then       'Phụ cấp khác
                Dim rPhuCapKhac As Rectangle = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width

                rPhuCapKhac.X += 1
                rPhuCapKhac.Y += rTienLuong.Height + 1
                rPhuCapKhac.Width = rPhuCapKhac.Width + w1 - 1
                rPhuCapKhac.Height = (rPhuCapKhac.Height / _ClassNumber - 2)
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuCapKhac)
                e.Graphics.DrawRectangle(Pens.Silver, rPhuCapKhac)
                e.Graphics.DrawString("Phụ cấp khác", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuCapKhac, format)
            End If

            If j = 20 Then       'TIỀN CÔNG (Tiền công lao động làm bảo vệ, lao công, tạp vụ,...)
                rTienCong_NH = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                Dim w4 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 4, -1, True).Width
                Dim w5 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 5, -1, True).Width
                rTienCong_NH.X += 1
                rTienCong_NH.Y += 1
                rTienCong_NH.Width = rTienCong_NH.Width + w1 + w2 + w3 + w4 + w5 + -1
                rTienCong_NH.Height = (rTienCong_NH.Height / _ClassNumber - 2) '* 2
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTienCong_NH)
                e.Graphics.DrawRectangle(Pens.Silver, rTienCong_NH)
                e.Graphics.DrawString("TIỀN CÔNG (TIỀN CÔNG LAO ĐỘNG LÀM BẢO VỆ, LAO CÔNG, TẠP VỤ,...)", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor),
                    rTienCong_NH, format)
            End If

            If j = 20 Then       'Tiền công
                rTienCong = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                rTienCong.X += 1
                rTienCong.Y += rTienCong_NH.Height + 1
                rTienCong.Width = rTienCong.Width + w1 + w2 + w3 - 1
                rTienCong.Height = rTienCong.Height / _ClassNumber - 2 '7 lop                    
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTienCong)
                e.Graphics.DrawRectangle(Pens.Silver, rTienCong)
                e.Graphics.DrawString("Tiền công", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTienCong, format)
            End If

            If j = 20 Then       'LĐ Định biên
                rDinhBien = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                rDinhBien.X += 1
                rDinhBien.Y += rTienCong_NH.Height + rTienCong.Height + 1
                rDinhBien.Width = rDinhBien.Width + w1 - 1
                rDinhBien.Height = (rDinhBien.Height / _ClassNumber - 2) * 2
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rDinhBien)
                e.Graphics.DrawRectangle(Pens.Silver, rDinhBien)
                e.Graphics.DrawString("LĐ Định biên", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rDinhBien, format)
            End If
            If j = 22 Then       'LĐ Phụ trợ
                rPhuTro = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                rPhuTro.X += 1
                rPhuTro.Y += rTienCong_NH.Height + rTienCong.Height + 1
                rPhuTro.Width = rPhuTro.Width + w1 - 1
                rPhuTro.Height = (rPhuTro.Height / _ClassNumber - 2) * 2
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rPhuTro)
                e.Graphics.DrawRectangle(Pens.Silver, rPhuTro)
                e.Graphics.DrawString("LĐ Phụ trợ", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rPhuTro, format)
            End If

            If j = 27 Then       'Lao động nghỉ việc
                rTienCong_NH = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                Dim w2 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 2, -1, True).Width
                Dim w3 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 3, -1, True).Width
                rTienCong_NH.X += 1
                rTienCong_NH.Y += 1
                rTienCong_NH.Width = rTienCong_NH.Width + w1 + w2 + w3 + -1
                rTienCong_NH.Height = (rTienCong_NH.Height / _ClassNumber - 2) * 2
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rTienCong_NH)
                e.Graphics.DrawRectangle(Pens.Silver, rTienCong_NH)
                e.Graphics.DrawString("Lao động nghỉ việc", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rTienCong_NH, format)
            End If
            If j = 27 Then       'Hưởng BHXH (Ốm đau, T.sản trong chế độ)
                rDinhBien = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                rDinhBien.X += 1
                rDinhBien.Y += rTienCong_NH.Height + 1
                rDinhBien.Width = rDinhBien.Width + w1 - 1
                rDinhBien.Height = (rDinhBien.Height / _ClassNumber - 2) * 3
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rDinhBien)
                e.Graphics.DrawRectangle(Pens.Silver, rDinhBien)
                e.Graphics.DrawString("Hưởng BHXH (Ốm đau, T.sản trong chế độ)", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rDinhBien, format)
            End If
            If j = 29 Then       'Ko BHXH (tự túc, Ko lương, Ốm đau, TS ko trong chế độ)
                rDinhBien = Me.dgv_main.GetCellDisplayRectangle(j, -1, True)
                Dim w1 As Integer = Me.dgv_main.GetCellDisplayRectangle(j + 1, -1, True).Width
                rDinhBien.X += 1
                rDinhBien.Y += rTienCong_NH.Height + 1
                rDinhBien.Width = rDinhBien.Width + w1 - 1
                rDinhBien.Height = (rDinhBien.Height / _ClassNumber - 2) * 3
                e.Graphics.FillRectangle(New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.BackColor), rDinhBien)
                e.Graphics.DrawRectangle(Pens.Silver, rDinhBien)
                e.Graphics.DrawString("Không BHXH (tự túc, không lương, Ốm đau, TS không trong chế độ)", Me.dgv_main.ColumnHeadersDefaultCellStyle.Font,
                    New SolidBrush(Me.dgv_main.ColumnHeadersDefaultCellStyle.ForeColor), rDinhBien, format)
            End If

        Next
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
            If ((dgv_main.CurrentRow.Cells("cln_TongHopId").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString() <> "")) Then
                Dim vCellCheck As DataGridViewCheckBoxCell
                'Checking whether the Datagridview Checkbox column is the first column
                If dgv_main.CurrentCellAddress.X = 1 Then
                    vCellCheck = dgv_main.CurrentRow.Cells("cln_Choice")
                    If (dgv_main.IsCurrentCellDirty) Then 'Checking for dirty cell
                        dgv_main.CommitEdit(DataGridViewDataErrorContexts.Commit) 'If it is dirty, making them to commit
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ckb_ChoiceAll_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyCode = Keys.Space) Then
            HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
        End If
    End Sub

    Private Sub ckb_ChoiceAll_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        HeaderCheckBoxClick(CType(sender, CheckBox), dgv_main)
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

    Private Sub num_nambc_ValueChanged(sender As Object, e As EventArgs) Handles num_nambc.ValueChanged
        tv_main.Nodes.Clear()
        _ChiLuongBLL.Fill_Tree(tv_main, CType(num_nambc.Value.ToString(), Integer), CType(num_thangbc.Value.ToString(), Integer), True)
        If IsCall Then
            tv_main_AfterSelect(sender, Nothing)
        End If
    End Sub

    Private Sub tv_main_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tv_main.AfterSelect
        ckb_ChoiceAll.Checked = False
        dgv_main.Rows.Clear()
        btn_xoabo.Enabled = False
        btn_pheduyet.Enabled = False

        _NodeCurrent = ""
        _NodeTag = ""
        _BranchTag = ""
        sProvinceName = ""
        sBranchName = ""
        sDepartment_Choise = ""
        sPos_Choise = ""
        sMain_Choise = ""
        sMonth_Choise = CType(num_thangbc.Value, Byte).ToString()
        Dim _LaoDong_Code As String = ""
        Dim _LoaiChi_Code As String = IIf(ARL_LoaiLuong.Count > 0 And cb_loaichiluong.SelectedIndex >= 0, ARL_LoaiLuong(cb_loaichiluong.SelectedIndex), "")
        _LaoDong_Code = IIf(rb_daihan.Checked, "1", IIf(rb_nganhan.Checked, "2", If(rb_tapsu.Checked, "3", "0")))
        
        Dim _KyBC As Byte = 99
        If CType(num_kybc.Value, Byte) <> 0 Then
            _KyBC = CType(num_kybc.Value, Byte)
        End If

        Dim sColumnCellNull As String = "SoLD;Luong_TTV;TL_PhuCap;PhCap_ThuHut_SoTien;PhCap_KhuVuc_SoTien;Luong_V1;Luong_V2_TamUng;TongLuong_TamUng;Loai_HDNH_DB_SoLD;Loai_HDNH_DB_SoTien;Loai_HDNH_PT_SoLD;Loai_HDNH_PT_SoTien;Tong_TienCong;LamDem_SoTienPC;Tong_TL_TC_TamChi;LdNgViec_CoBHXH_SoLD;LdNgViec_KoBHXH_SoLD"
        Dim ARL_ColNull() As String = Globals.Splip_Strings(sColumnCellNull, ";")

        Dim sColumnCellNullN2 As String = "LCB_Tong_HeSo;LCB_HeSo;PhCap_ChVu_HeSo;PhCap_TrNhiem_HeSo;PhCap_DocHai_HeSo;LdNgViec_CoBHXH_HSLPC;LdNgViec_KoBHXH_HSLPC"
        Dim ARL_ColNullN2() As String = Globals.Splip_Strings(sColumnCellNullN2, ";")

        Dim _ColumnVisible As String = "OrderNo;ChiNhanh_Cd;ChiNhanh_HT;DonVi_CL_Cd;DonVi_CL_HT;PhongBan_CL_Cd;PhongBan_HT;NgayBC;NamBC;ThangBC;PhanLoai_Cd;LaoDong_Cd;Luong_TTV_Ma;Luong_TTV_HT;Luong_CSo;BHXH_CB;BHXH_DV;BHYT_CB;BHYT_DV;BHTN_CB;BHTN_DV;DPCD_CB;DPCD_LD_NN;Tinh_Thue_TNCN;Tinh_Thue_TNCN_HT;TNTT_Muc_BanThan;TNTT_Muc_PhuThuoc;SoNgayLViec_Thang;MucTamUng_V2;HeSoLuong_V2;TrangThai;TrangThai_HT"
        Dim ARL_Cols() As String = Globals.Splip_Strings(_ColumnVisible, ";")
        Dim _KieuIn As Byte = 0
        Dim _TrangThaiSelect As Byte
        Dim ARL_PosCode As String() = Nothing, ARL_MonthCode As String() = Nothing, ARL_Pos As String() = Nothing, ARL_MainPos As String() = Nothing

        Cursor = Cursors.WaitCursor
        If (tv_main.Nodes.Count > 0) Then
            IsCall = True
            If Not IsNothing(tv_main.SelectedNode) Then
                _NodeTag = tv_main.SelectedNode.Tag.ToString()
                _BranchTag = _NodeTag
                _NodeCurrent = tv_main.SelectedNode.Tag.ToString().Trim()
                dgv_main.Columns("cln_Choice").Visible = False
                ckb_ChoiceAll.Visible = False
                If (tv_main.SelectedNode.Tag.ToString().Substring(2, 2) = "CN") Then
                    bFlagCall = 2
                    sPos_Choise = ""
                    ARL_MainPos = tv_main.SelectedNode.Tag.ToString().Split("_")
                    If ARL_MainPos(2).ToString() = "000100" Or ARL_MainPos(2).ToString() = "000199" Then
                        sDepartment_Choise = ""
                        sMain_Choise = "000100"
                        sPos_Choise = ARL_MainPos(2).ToString()
                        sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                        sBranchName = "HỘI SỞ CHÍNH"
                    Else
                        sMain_Choise = IIf(ARL_MainPos.Length > 0, ARL_MainPos(2).ToString(), "")
                        sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                        sBranchName = tv_main.SelectedNode.Text.ToString()
                    End If

                    'If Cap_Nd <> 1 Then
                    '    sMain_Choise = IIf(ARL_MainPos.Length > 0, ARL_MainPos(2).ToString(), "")
                    '    sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                    '    sBranchName = tv_main.SelectedNode.Text.ToString()
                    'ElseIf Cap_Nd = 1 Then
                    '    sDepartment_Choise = ""
                    '    sMain_Choise = "000100"
                    '    sPos_Choise = ARL_MainPos(2).ToString()
                    '    sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                    '    sBranchName = "HỘI SỞ CHÍNH"
                    'End If
                    
                ElseIf (tv_main.SelectedNode.Tag.ToString().Substring(2, 2) = "PB") Then
                    ARL_PosCode = tv_main.SelectedNode.Tag.ToString().Split("_")
                    ARL_MainPos = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
                    sMain_Choise = "000100"
                    sPos_Choise = IIf(ARL_MainPos.Length > 0, ARL_MainPos(2).ToString(), "")
                    sDepartment_Choise = ARL_PosCode(2).ToString()
                    bFlagCall = 2
                    sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                    sBranchName = tv_main.SelectedNode.Text.ToString()
                ElseIf (tv_main.SelectedNode.Tag.ToString().Substring(2, 2) = "DV") Then
                    bFlagCall = 2
                    ARL_PosCode = tv_main.SelectedNode.Tag.ToString().Split("_")
                    If Cap_Nd <> 1 Then
                        If (ARL_PosCode.Length > 0) Then
                            sPos_Choise = ARL_PosCode(2).ToString()
                        End If
                        ARL_MainPos = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
                        If (ARL_MainPos.Length > 0) Then
                            sMain_Choise = ARL_MainPos(2).ToString()
                        End If
                    ElseIf Cap_Nd = 1 Then
                        sDepartment_Choise = ARL_PosCode(2).ToString()
                        sMain_Choise = "000100"
                        ARL_Pos = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
                        sPos_Choise = ARL_Pos(2).ToString()
                    End If
                    sProvinceName = tv_main.SelectedNode.Parent.Text.ToString()
                    sBranchName = tv_main.SelectedNode.Text.ToString()
                ElseIf (tv_main.SelectedNode.Tag.ToString().Substring(2, 2) = "TG") Then
                    dgv_main.Columns("cln_Choice").Visible = True
                    ckb_ChoiceAll.Visible = True
                    btn_xoabo.Enabled = True
                    btn_pheduyet.Enabled = True
                    bFlagCall = 1
                    ARL_PosCode = tv_main.SelectedNode.Parent.Tag.ToString().Split("_")
                    If Cap_Nd <> 1 Then
                        If (ARL_PosCode.Length > 0) Then
                            sPos_Choise = ARL_PosCode(2).ToString()
                        End If
                        ARL_MainPos = tv_main.SelectedNode.Parent.Parent.Tag.ToString().Split("_")
                        If (ARL_MainPos.Length > 0) Then
                            sMain_Choise = ARL_MainPos(2).ToString()
                        End If
                    ElseIf Cap_Nd = 1 Then
                        sDepartment_Choise = ARL_PosCode(2).ToString()
                        sMain_Choise = "000100"
                        ARL_Pos = tv_main.SelectedNode.Parent.Parent.Tag.ToString().Split("_")
                        sPos_Choise = ARL_Pos(2).ToString()
                    End If
                    ARL_MonthCode = tv_main.SelectedNode.Tag.ToString().Split("_")
                    If (ARL_MonthCode.Length > 0) Then
                        sMonth_Choise = ARL_MonthCode(2).ToString()
                    End If
                    sProvinceName = tv_main.SelectedNode.Parent.Parent.Text.ToString()
                    sBranchName = tv_main.SelectedNode.Parent.Text.ToString()
                End If
            End If

            Dim db_BC01TL As DataTable = New DataTable
            db_BC01TL = _ChiLuongBLL.ChiLuong_TongHop_Thang_BC01TL(sMain_Choise, sPos_Choise, sDepartment_Choise, "", CType(num_nambc.Value, Integer), CType(sMonth_Choise, Integer), _KyBC, _LoaiChi_Code, _LaoDong_Code, 0, bFlagCall)
            If Not (db_BC01TL Is Nothing) Then
                If (db_BC01TL.Rows.Count > 0) Then
                    For i As Integer = 0 To db_BC01TL.Rows.Count - 1
                        dgv_main.Rows.Add()
                        dgv_main.Rows(i).Cells("cln_TongHopId").Value = db_BC01TL.Rows(i)("TongHopId").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_KieuIn").Value = db_BC01TL.Rows(i)("KieuIn").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_STT").Value = db_BC01TL.Rows(i)("STT").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_ThongTin_HT").Value = db_BC01TL.Rows(i)("ThongTin_HT").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_PhanLoai_HT").Value = db_BC01TL.Rows(i)("PhanLoai_HT").ToString().Trim()
                        dgv_main.Rows(i).Cells("cln_LaoDong_HT").Value = db_BC01TL.Rows(i)("LaoDong_HT").ToString().Trim()
                        If db_BC01TL.Rows(i)("KyBC").ToString().Trim() = "0" Then
                            dgv_main.Rows(i).Cells("cln_KyBC").Value = ""
                        Else : dgv_main.Rows(i).Cells("cln_KyBC").Value = db_BC01TL.Rows(i)("KyBC").ToString().Trim()
                        End If
                        dgv_main.Rows(i).Cells("cln_GhiChu_CL").Value = db_BC01TL.Rows(i)("GhiChu_CL").ToString().Trim()
                        'Fill dữ liệu nếu NULL hoặc rỗng cho thành [Rỗng]
                        For Each _ValueNull As String In ARL_ColNull
                            If Not String.IsNullOrEmpty(_ValueNull) Then
                                If IsNothing(db_BC01TL.Rows(i)(_ValueNull)) Or String.IsNullOrEmpty(db_BC01TL.Rows(i)(_ValueNull).ToString()) Or db_BC01TL.Rows(i)(_ValueNull).ToString().Trim() = "" Or db_BC01TL.Rows(i)(_ValueNull).ToString().Trim() = "0" Then
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull).Value = ""
                                Else
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull).Value = Double.Parse(db_BC01TL.Rows(i)(_ValueNull).ToString(), Globals.cultureNum).ToString("N0", vNFInfo)
                                End If
                            End If
                        Next
                        For Each _ValueNull2 As String In ARL_ColNullN2
                            If Not String.IsNullOrEmpty(_ValueNull2) Then
                                If IsNothing(db_BC01TL.Rows(i)(_ValueNull2)) Or String.IsNullOrEmpty(db_BC01TL.Rows(i)(_ValueNull2).ToString()) Or db_BC01TL.Rows(i)(_ValueNull2).ToString().Trim() = "" Or db_BC01TL.Rows(i)(_ValueNull2).ToString().Trim() = "0" Then
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull2).Value = ""
                                Else
                                    dgv_main.Rows(i).Cells("cln_" + _ValueNull2).Value = Double.Parse(db_BC01TL.Rows(i)(_ValueNull2).ToString(), Globals.cultureNum).ToString("N2", vNFInfo)
                                End If
                            End If
                        Next
                        For Each _Value As String In ARL_Cols
                            If Not String.IsNullOrEmpty(_Value) Then
                                dgv_main.Rows(i).Cells("cln_" + _Value).Value = db_BC01TL.Rows(i)(_Value).ToString().Trim()
                            End If
                        Next
                        _KieuIn = CType(db_BC01TL.Rows(i)("KieuIn"), Byte)
                        _ChiLuongBLL.SetStyleRowGrid(i, dgv_main, _KieuIn, 0)
                        If IsNothing(db_BC01TL.Rows(i)("TrangThai")) Or String.IsNullOrEmpty(db_BC01TL.Rows(i)("TrangThai").ToString()) Then
                            _TrangThaiSelect = 1
                        Else
                            _TrangThaiSelect = CType(db_BC01TL.Rows(i)("TrangThai"), Byte)
                        End If

                        If _TrangThaiSelect = 2 Then
                            dgv_main.Rows(i).DefaultCellStyle.ForeColor = Color.Green
                        ElseIf _TrangThaiSelect = 0 Then
                            dgv_main.Rows(i).DefaultCellStyle.ForeColor = Color.Red
                        End If
                    Next
                End If
            End If


        End If
        TotalCheckBoxes = dgv_main.RowCount
        TotalCheckedCheckBoxes = 0
        Cursor = Cursors.Default
    End Sub

    Private Sub num_thangbc_ValueChanged(sender As Object, e As EventArgs) Handles num_thangbc.ValueChanged
        tv_main_AfterSelect(sender, Nothing)
    End Sub

    Private Sub num_kybc_ValueChanged(sender As Object, e As EventArgs) Handles num_kybc.ValueChanged
        If IsCall Then
            tv_main_AfterSelect(sender, Nothing)
        End If
    End Sub

    Private Sub cb_loaichiluong_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_loaichiluong.SelectedIndexChanged
        If IsCall Then
            tv_main_AfterSelect(sender, Nothing)
        End If
    End Sub

    Private Sub dgv_main_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_main.CellContentClick
        btn_suadoi.Enabled = False
        If dgv_main.Rows.Count > 1 Then
            Dim _StatusRow As Byte = 0
            If IsNothing(dgv_main.CurrentRow.Cells("cln_TrangThai").Value) Or String.IsNullOrEmpty(dgv_main.CurrentRow.Cells("cln_TrangThai").Value.ToString()) Then
                _StatusRow = 1
            Else
                _StatusRow = CType(dgv_main.CurrentRow.Cells("cln_TrangThai").Value.ToString(), Byte)
            End If

            Dim _TongHopId As Long = 0
            _TongHopId = CType(dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString(), Long)
            If Not (dgv_main.CurrentRow.Cells("cln_KieuIn").Value.ToString() Is Nothing) Then
                If dgv_main.CurrentRow.Cells("cln_KieuIn").Value.ToString() = "3" And _TongHopId <> 0 Then
                    If _StatusRow <> 2 Then
                        btn_suadoi.Enabled = True
                    End If
                End If
            End If

            'If (dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name = "cln_Choice") Then
            '    If dgv_main.CurrentRow.Cells("cln_KieuIn").Value.ToString() = "1" Or dgv_main.CurrentRow.Cells("cln_KieuIn").Value.ToString() = "0" Or _TongHopId = 0 Then
            '        dgv_main.Rows(dgv_main.CurrentCell.RowIndex).ReadOnly = True
            '        Return
            '    Else
            '        dgv_main.Rows(dgv_main.CurrentCell.RowIndex).ReadOnly = False
            '    End If
            'End If
        End If
    End Sub

    Private Sub dgv_main_KeyUp(sender As Object, e As KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellContentClick(sender, Nothing)
    End Sub

    Private Sub dgv_main_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv_main.CellMouseClick
        dgv_main_CellContentClick(sender, Nothing)
    End Sub

    Private Sub dgv_main_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgv_main.CellBeginEdit
        'If dgv_main.Rows.Count > 1 Then
        '    If (dgv_main.Columns(dgv_main.CurrentCell.ColumnIndex).Name = "cln_Choice") Then
        '        Dim _TongHopId As Long = 0
        '        _TongHopId = CType(dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_TongHopId").Value.ToString(), Long)

        '        If dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() = "1" Or dgv_main.Rows(dgv_main.CurrentCell.RowIndex).Cells("cln_KieuIn").Value.ToString().Trim() = "0" Or _TongHopId = 0 Then
        '            e.Cancel = True
        '            dgv_main.Rows(dgv_main.CurrentCell.RowIndex).ReadOnly = True
        '            Return
        '        Else
        '            dgv_main.Rows(dgv_main.CurrentCell.RowIndex).ReadOnly = False
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub dgv_main_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv_main.KeyDown

    End Sub

    Private Sub btn_export_excel_Click(sender As Object, e As EventArgs) Handles btn_export_excel.Click
        If dgv_main.Rows.Count > 1 Then
            Dim _DonVi_TK_Cd As String = ""
            Dim db_luongex As DataTable = New DataTable()
            Dim _TongHopId As Long = CType(dgv_main.CurrentRow.Cells("cln_TongHopId").Value.ToString(), Long)

            Cursor = Cursors.WaitCursor
            lbl_waitting.Text = "Waitting for data export..."
            'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            If Not (dgv_main.CurrentRow.Cells("cln_KieuIn").Value.ToString() Is Nothing) Then
                If dgv_main.CurrentRow.Cells("cln_KieuIn").Value.ToString() = "3" And _TongHopId <> 0 Then
                    'Xuất dữ liệu Chi tiết lương Tháng: Dài hạn (Kỳ 1, 2); Ngắn hạn
                    'Bắt đầu xuất dữ liệu Chi lương ra excel
                    db_luongex = _ChiLuongBLL.ChiLuong_BangKeCT_GetSearch(_TongHopId, 0, "", 0, 0, 99, "", "", 0, 9)
                    If (db_luongex Is Nothing Or db_luongex.Rows.Count <= 0) Then
                        MessageBox.Show("Không có dữ liệu Chi lương để xuất báo cáo file excel. Vui lòng kiểm tra lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Return
                    End If
                   
                    _DonVi_TK_Cd = dgv_main.CurrentRow.Cells("cln_DonVi_CL_Cd").Value.ToString()
                    If _DonVi_TK_Cd = "000100" Or _DonVi_TK_Cd = "000199" Then
                        sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                        sBranchName = dgv_main.CurrentRow.Cells("cln_ThongTin_HT").Value.ToString()
                    Else
                        sProvinceName = dgv_main.CurrentRow.Cells("cln_ChiNhanh_HT").Value.ToString()
                        sBranchName = dgv_main.CurrentRow.Cells("cln_DonVi_CL_HT").Value.ToString().Replace("Chi nhánh NHCSXH", "Hội sở")
                    End If
                    Dim iKyBC As Byte = CType(dgv_main.CurrentRow.Cells("cln_KyBC").Value.ToString(), Byte)
                    Dim iThangBC As Byte = CType(dgv_main.CurrentRow.Cells("cln_ThangBC").Value.ToString(), Integer)
                    Dim _ResultEx As String = _ChiLuongBLL.ExportExcel_CT_LuongThang(_DonVi_TK_Cd, sProvinceName, sBranchName, CType(dgv_main.CurrentRow.Cells("cln_NgayBC").Value, Date), dgv_main.CurrentRow.Cells("cln_PhanLoai_Cd").Value.ToString(), iKyBC, dgv_main.CurrentRow.Cells("cln_LaoDong_Cd").Value.ToString(), db_luongex)

                    Dim sMessEx As String = ""
                    If dgv_main.CurrentRow.Cells("cln_LaoDong_Cd").Value.ToString() = "2" Then
                        sMessEx = "Tháng " + iThangBC.ToString("D2") + " của Lao động Ngăn hạn"
                    ElseIf dgv_main.CurrentRow.Cells("cln_LaoDong_Cd").Value.ToString() = "1" Then
                        sMessEx = "Kỳ " + iKyBC.ToString("D2") + " Tháng " + iThangBC.ToString("D2") + " của Lao động CMNV"
                    End If
                    MessageBox.Show("Báo cáo chi tiết lương [" + sMessEx + "] được xuất ra file excel thành công!" + vbNewLine + "Đường dẫn chứa file báo cáo: [" + _ResultEx + "]", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)

                   
                Else
                    'Xuất excel báo cáo Chi lương mẫu 01/TL
                    Dim _LaoDong_Code As String = IIf(rb_daihan.Checked, "1", IIf(rb_nganhan.Checked, "2", If(rb_tapsu.Checked, "3", "0")))
                    Dim _LoaiChi_Code As String = IIf(ARL_LoaiLuong.Count > 0 And cb_loaichiluong.SelectedIndex >= 0, ARL_LoaiLuong(cb_loaichiluong.SelectedIndex), "")
                    Dim _KyBC As Byte = IIf(CType(num_kybc.Value, Byte) <> 0, CType(num_kybc.Value, Byte), 99)
                    Dim iThangBC As Integer = IIf(CType(num_thangbc.Value, Integer) <> 0, CType(num_thangbc.Value, Integer), 0)
                    Dim iNamBC As Integer = IIf(CType(num_nambc.Value, Integer) >= 2002, CType(num_nambc.Value, Integer), 0)
                    db_luongex = _ChiLuongBLL.ChiLuong_TongHop_Thang_BC01TL(sMain_Choise, sPos_Choise, sDepartment_Choise, "", CType(num_nambc.Value, Integer), CType(sMonth_Choise, Integer), _KyBC, _LoaiChi_Code, _LaoDong_Code, 0, 2)
                    Dim _ChiNhanhRow0 As String = ""
                    If (IsNothing(_DonVi_TK_Cd) Or String.IsNullOrEmpty(_DonVi_TK_Cd)) Then
                        _ChiNhanhRow0 = dgv_main.Rows(0).Cells("cln_ChiNhanh_Cd").Value.ToString()
                        If String.IsNullOrEmpty(dgv_main.Rows(0).Cells("cln_DonVi_CL_Cd").Value.ToString()) Then
                            _DonVi_TK_Cd = dgv_main.Rows(0).Cells("cln_ChiNhanh_Cd").Value.ToString()
                        Else
                            _DonVi_TK_Cd = dgv_main.Rows(0).Cells("cln_DonVi_CL_Cd").Value.ToString()
                        End If
                    End If

                    If sProvinceName = "" Then
                        If _ChiNhanhRow0 = "000199" Or _ChiNhanhRow0 = "000100" Then
                            sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                            If String.IsNullOrEmpty(dgv_main.Rows(0).Cells("cln_PhongBan_CL_Cd").Value.ToString()) Then
                                sBranchName = dgv_main.Rows(0).Cells("cln_DonVi_CL_HT").Value.ToString()
                            Else
                                sBranchName = dgv_main.Rows(0).Cells("cln_PhongBan_CL_HT").Value.ToString()
                            End If
                        ElseIf String.IsNullOrEmpty(dgv_main.Rows(0).Cells("cln_DonVi_CL_Cd").Value.ToString()) Then
                            sProvinceName = "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI"
                            sBranchName = dgv_main.Rows(0).Cells("cln_ChiNhanh_HT").Value.ToString()
                        Else
                            sProvinceName = dgv_main.Rows(0).Cells("cln_ChiNhanh_HT").Value.ToString()
                            sBranchName = dgv_main.Rows(0).Cells("cln_DonVi_CL_HT").Value.ToString()
                        End If
                    End If

                    'sBranchName = dgv_main.Rows(0).Cells("cln_ChiNhanh_HT").Value.ToString()
                    Dim _ResultEx As String = _ChiLuongBLL.ExportExcel_TH_LuongThang_01TL(_DonVi_TK_Cd, sProvinceName, sBranchName, iNamBC, iThangBC, _LoaiChi_Code, _KyBC, _LaoDong_Code, db_luongex)
                    MessageBox.Show("Báo cáo Tình hình thực hiện Lao động - Tiền lương tháng được xuất ra file excel thành công!" + vbNewLine + "Đường dẫn chứa file báo cáo: [" + _ResultEx + "]", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                End If
            End If

            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
            lbl_waitting.Text = ""
            Cursor = Cursors.Default
        End If
    End Sub

   
End Class