Imports CrystalDecisions.CrystalReports
Imports CrystalDecisions.Shared

Public Class frmChiLuong

    Private initpnlBS As Boolean = False
    Private actIdChiLuong As String
    Private actLoaiThuChi As Integer
    Private actIdDonVi As Integer = IdDONVI
    Private actNam As Integer
    Private actThang As Integer
    Private actKy As Int16
    Private actLoaiCB As Int16 = 0
    Private isConfirm As Boolean = False
    Private listCB As List(Of String)
    'Private listCBs As String
    Dim frmUCL As frmUpdateChiLuong
    Dim frmDSCB As frmChonDS_CB

    Private Sub frmChiLuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bindTreeviewLuong(treeChiLuong)
        treeChiLuong.ExpandAll()
        bindCboPhong(IdDONVI)
        blankFrm()
        If DONVI = gMaDonViTW Then
            labGD.Text = "Phó Tổng giám đốc"
        End If
        ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 1)
    End Sub

    Private Sub frmChiLuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub treeChiLuong_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles treeChiLuong.AfterSelect
        If e.Action <> TreeViewAction.Unknown Then
            Dim arr() As String
            arr = treeChiLuong.SelectedNode.Tag.ToString.Split("_")
            labAlertChiLuong.Text = ""
            labGhiChu.Text = "Ghi chú"
            Select Case arr(0)
                Case "DV"
                    ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 2)
                    actIdDonVi = arr(1)
                    If treeChiLuong.SelectedNode.GetNodeCount(True) = 0 Then
                        If Not treeChiLuong.SelectedNode.IsExpanded Then
                            getChiLuong_Nam(treeChiLuong.SelectedNode, arr(1))
                        End If
                    End If
                    treeChiLuong.SelectedNode.Expand()
                Case "NAM"
                    ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 2)
                    getChiLuong_Loai(treeChiLuong.SelectedNode, arr(1), arr(2))
                    actNam = arr(2)
                    treeChiLuong.SelectedNode.Expand()
                Case "CT"
                    setVarValueActived(arr(1), arr(2), arr(6), actNam, arr(3), arr(4), arr(7))
                    labActived.Text = treeChiLuong.SelectedNode.Text
                    showGrid(arr(1))
                    If (arr(5) = 0) Then
                        isConfirm = True
                        ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 4)
                    Else
                        isConfirm = False
                        ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 3)
                    End If
            End Select
        End If
    End Sub

    Private Sub cboLoaiChi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLoaiChi.SelectedIndexChanged
        gridResult.DataSource = Nothing
        If cboLoaiChi.SelectedIndex = 0 Then
            actLoaiThuChi = 1
            pnlLuong.Visible = True
            pnlBoSung.Visible = False
            pnlGroupCB.Enabled = False
            cbGroupCB.Checked = False
            txtGroupCB.Text = ""
            labGhiChu.Text = "Ghi chú"
            listCB = Nothing
            'listCBs = ""
            If MAX_KY = 1 Then
                labKy.Text = ""
                labelKy.Visible = False
                numKy.Visible = False
            Else
                numKy.Minimum = 1
                numKy.Maximum = 2
                labKy.Text = "/ " & MAX_KY.ToString & " KỲ"
            End If
        Else
            labKy.Text = ""
            If cboLoaiChi.SelectedIndex = 2 Then
                actLoaiThuChi = 3
                pnlBoSung.Visible = True
                txtGroupCB.Enabled = False
                lnkChoiseCB.Enabled = False
                pnlGroupCB.Enabled = True
                pnlLuong.Visible = False
                cbGroupCB.Checked = False
                txtGroupCB.Text = ""
                labGhiChu.Text = "Nội dung"
                listCB = Nothing
                'listCBs = ""
            Else
                actLoaiThuChi = 2
                pnlGroupCB.Enabled = False
                pnlLuong.Visible = False
                pnlBoSung.Visible = False
                cbGroupCB.Checked = False
                txtGroupCB.Text = ""
                labGhiChu.Text = "Ghi chú"
                listCB = Nothing
                'listCBs = ""
            End If
        End If

    End Sub

    Private Sub bindCboPhong(ByVal vIdDonvi As Integer)
        Try
            If vIdDonvi > 0 Then
                cboPhong.DataSource = listPhong(vIdDonvi, True, True)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bntNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntNew.Click
        ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 2)
        blankFrm()
    End Sub

    Private Sub bntCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCreate.Click
        Dim m_ChiLuong As New HS_ChiLuong
        Dim Dset As New DataSet
        Dim IDChi As String
        Try
            labAlertChiLuong.Text = ""
            If cboLoaiChi.SelectedIndex = 0 Then
                actLoaiThuChi = 1
            Else
                If cboLoaiChi.SelectedIndex = 2 Then
                    actLoaiThuChi = 3
                Else
                    actLoaiThuChi = 2
                End If
            End If
            If countChiLuong(actLoaiThuChi, getActIdDonVi(cboPhong.SelectedValue), cboPhong.SelectedValue, dtpkNam.Value.Year, numThang.Value, numKy.Value, actLoaiCB, True) > 0 Then
                MessageBox.Show("Đã chi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If actLoaiThuChi = 1 And MAX_KY = 2 And numKy.Value = 2 Then
                If countChiLuong(actLoaiThuChi, actIdDonVi, cboPhong.SelectedValue, dtpkNam.Value.Year, numThang.Value, 1, actLoaiCB, True) = 0 Then
                    MessageBox.Show("Lương kỳ I chưa chi!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                    Exit Sub
                End If
            End If
            If actLoaiThuChi = 3 Then
                If cbGroupCB.Checked Then
                    If Not (listCB Is Nothing) Then
                        If listCB.Count <= 0 Then
                            MessageBox.Show("Hãy chọn lại danh sách cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                            Exit Sub
                        End If
                    Else
                        MessageBox.Show("Hãy chọn lại danh sách cán bộ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        Exit Sub
                    End If
                End If
                If txtHesochi.Text.Trim = "" Then txtHesochi.Text = "0"
                If txtHesochi.Text.Trim = "0" Then
                    If (txtSoTienBS.Text.Trim = "" Or txtSoTienBS.Text.Trim = "0") Then
                        MessageBox.Show("Hãy nhập số tiền thực lĩnh!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        txtSoTienBS.Focus()
                        Exit Sub
                    ElseIf (txtSoTienBS.Text.Trim.Length > 11) Then
                        MessageBox.Show("Hãy nhập lại số tiền thực lĩnh!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                        txtSoTienBS.Focus()
                        Exit Sub
                    End If
                End If
            End If
            If actIdDonVi = 0 Then actIdDonVi = IdDONVI
            m_ChiLuong.IdDonVi_CL = actIdDonVi
            m_ChiLuong.IdLoaiThuChi = actLoaiThuChi
            m_ChiLuong.Nam = dtpkNam.Value.Year
            m_ChiLuong.Thang = numThang.Value
            m_ChiLuong.GhiChu = txtGhiChu.Text
            m_ChiLuong.Add()
            IDChi = m_ChiLuong.IdHS_ChiLuong
            If IDChi = "" Then Exit Try
            If actLoaiThuChi = 1 Then
                If rdDH.Checked Then
                    actLoaiCB = 0
                Else
                    If rdNH.Checked Then
                        actLoaiCB = 1
                    Else
                        actLoaiCB = 2
                    End If
                End If
            End If
            setVarValueActived(IDChi, actLoaiThuChi, actIdDonVi, dtpkNam.Value.Year, numThang.Value, numKy.Value, actLoaiCB)
            bindGrid(IDChi, actLoaiThuChi, m_ChiLuong.IdDonVi_CL, actLoaiCB, cboPhong.SelectedValue, listCB, m_ChiLuong.Nam, m_ChiLuong.Thang, numKy.Value)
            ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 3)
            If actLoaiThuChi = 1 Then
                labActived.Text = "Lương tháng " & actThang & IIf(actKy = 0, "", " kỳ " & actKy) & "(Chưa kiểm soát)"
            Else
                If actLoaiThuChi = 2 Then
                    labActived.Text = "Thêm giờ " & actThang & "(Chưa kiểm soát)"
                Else
                    labActived.Text = "Chi khác " & actThang & IIf(actKy = 0, "", " kỳ " & actKy) & "(Chưa kiểm soát)"
                End If
            End If
            Dim childNode As TreeNode
            If countChiLuong(actLoaiThuChi, actIdDonVi, cboPhong.SelectedValue, dtpkNam.Value.Year, numThang.Value, numKy.Value, actLoaiCB, False) = 1 Then
                bindTreeviewLuong_Loai(treeChiLuong, dtpkNam.Value.Year)
            Else
                For Each childNode In treeChiLuong.Nodes(0).Nodes
                    Dim arr() As String
                    arr = childNode.Tag.ToString.Split("_")
                    If (arr(2) = dtpkNam.Value.Year) Then
                        arr = childNode.Tag.ToString.Split("_")
                        getChiLuong_Loai(childNode, arr(1), arr(2))
                        childNode.Expand()
                        Exit For
                    End If
                Next
            End If
            isConfirm = False
        Catch ex As Exception

        End Try

    End Sub

    Private Sub bntDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntDelete.Click
        Try
            If actIdChiLuong = "" Then
                MessageBox.Show("Hãy chọn lại danh sách lương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            If MessageBox.Show("Bạn có chắc chắn xoá các thông tin này không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                Try
                    Dim m_ChiLuong As HS_ChiLuong = New HS_ChiLuong
                    Dim _node As TreeNode
                    Dim arr() As String
                    m_ChiLuong.IdHS_ChiLuong = actIdChiLuong
                    m_ChiLuong.Delete()
                    Try
                        _node = treeChiLuong.SelectedNode.Parent
                        arr = _node.Tag.ToString.Split("_")
                        If arr(0) = "ROOT" Then
                            _node = treeChiLuong.SelectedNode
                            arr = _node.Tag.ToString.Split("_")
                        End If
                        ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 2)
                        getChiLuong_Loai(_node, arr(1), arr(2))
                        _node.Expand()
                    Catch ex As Exception
                        bindTreeviewLuong_Loai(treeChiLuong, dtpkNam.Value.Year)
                    End Try
                    gridResult.DataSource = Nothing
                    labActived.Text = ""
                    labAlertChiLuong.Text = "Xoá thành công!"
                Catch ex As Exception
                    MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                End Try
            End If
        Catch ex As Exception
            MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntConfirm.Click
        Dim m_ChiLuong As HS_ChiLuong = New HS_ChiLuong
        Dim childNode As TreeNode
        Dim arr() As String

        m_ChiLuong.IdHS_ChiLuong = actIdChiLuong
        m_ChiLuong.Update_Confirm()
        isConfirm = True
        ButtonDisplayed(bntCreate, bntDelete, bntConfirm, bntPrint, 4)

        For Each childNode In treeChiLuong.Nodes(0).Nodes
            arr = childNode.Tag.ToString.Split("_")
            If (arr(2) = actNam) Then
                arr = childNode.Tag.ToString.Split("_")
                getChiLuong_Loai(childNode, arr(1), arr(2))
                childNode.Expand()
                Exit For
            End If
        Next
        labActived.Text = labActived.Text.Substring(0, (labActived.Text).Length - 17)
        ''_node = treeChiLuong.SelectedNode.Parent
        '_node = treeChiLuong.SelectedNode
        'arr = _node.Tag.ToString.Split("_")
        'getChiLuong_Loai(_node, arr(1), arr(2))
        '_node.Expand()
    End Sub

    Private Sub bntPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntPrint.Click
        'Dim frm As frmReportChiLuong = New frmReportChiLuong
        'frm.ShowDialog()
        If actLoaiThuChi = 1 Then
            Dim frm As frmReport_ChiLuong = New frmReport_ChiLuong
            frm.IdChiLuong = actIdChiLuong
            frm.IDLoaiThuChi = actLoaiThuChi
            frm.LoaiCB = actLoaiCB
            frm.Ky = actKy
            frm.ID_DonVi = actIdDonVi
            frm.Thang = actThang
            frm.Nam = actNam
            frm.LapBieu = txtLB.Text
            frm.KiemSoat = txtKS.Text
            frm.GiamDoc = txtGD.Text
            frm.GhiChu = txtGhiChu.Text
            frm.ShowDialog()
        Else
            If actLoaiThuChi = 2 Then
                Dim frm As frmReport_ChiLuong = New frmReport_ChiLuong
                If MessageBox.Show("Có hiển thị thông tin tài khoản không?", "Lựa chọn mẫu biểu", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes Then
                    frm.TK = True
                Else
                    frm.TK = False
                End If
                frm.IdChiLuong = actIdChiLuong
                frm.IDLoaiThuChi = actLoaiThuChi
                frm.LoaiCB = 0
                frm.Ky = 0
                frm.ID_DonVi = actIdDonVi
                frm.Thang = actThang
                frm.Nam = actNam
                frm.LapBieu = txtLB.Text
                frm.KiemSoat = txtKS.Text
                frm.GiamDoc = txtGD.Text
                frm.GhiChu = txtGhiChu.Text
                frm.ShowDialog()
            Else
                Dim frm As frmReport_ChiLuong = New frmReport_ChiLuong
                frm.IdChiLuong = actIdChiLuong
                frm.IDLoaiThuChi = actLoaiThuChi
                frm.LoaiCB = actLoaiCB
                frm.Ky = 0
                frm.ID_DonVi = actIdDonVi
                frm.Thang = actThang
                frm.Nam = actNam
                frm.LapBieu = txtLB.Text
                frm.KiemSoat = txtKS.Text
                frm.GiamDoc = txtGD.Text
                frm.GhiChu = txtGhiChu.Text
                frm.ShowDialog()
            End If
        End If
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub gridResult_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridResult.CellDoubleClick
        If Not isConfirm Then
            frmUCL = New frmUpdateChiLuong
            frmUCL.IdChi_ChiTiet = gridResult.CurrentRow.Cells(0).Value.ToString
            frmUCL.IdChiLuong = gridResult.CurrentRow.Cells(1).Value.ToString
            frmUCL.IdDonVi = actIdDonVi
            frmUCL.IDLoaiChi = actLoaiThuChi
            If actLoaiThuChi = 3 Then
                frmUCL.NH_TS = gridResult.CurrentRow.Cells(2).Value.ToString
            Else
                frmUCL.NH_TS = actLoaiCB
            End If
            frmUCL.Progress_Changed = New frmUpdateChiLuong.ProgressChangedEventHandler(AddressOf ReLoad_frmChiLuong)
            frmUCL.ShowDialog()
        End If
    End Sub

    Private Sub pnlBoSung_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlBoSung.VisibleChanged
        If pnlBoSung.Visible = True Then
            If initpnlBS = False Then
                cboLoaiLuongBS.DataSource = listDanhmuc(42)
                initpnlBS = True
            End If
        End If
    End Sub

    Private Sub txtHesochi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHesochi.TextChanged
        If txtHesochi.Text <> "" Then
            txtHesochi.Text = formatDouble(txtHesochi.Text.Trim)
            txtHesochi.SelectionStart = txtHesochi.Text.Length
        End If
    End Sub

    Private Sub txtSoTienBS_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSoTienBS.TextChanged
        Try
            txtSoTienBS = formatMoneyinTextbox(txtSoTienBS)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtLB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLB.LostFocus
        Try
            txtLB.Text = standardizeName(txtLB.Text)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtKS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKS.LostFocus
        txtKS.Text = standardizeName(txtKS.Text)
    End Sub

    Private Sub txtGD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGD.LostFocus
        txtGD.Text = standardizeName(txtGD.Text)
    End Sub

    Private Sub txtGhiChu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGhiChu.LostFocus
        txtGhiChu.Text = standardizeString(txtGhiChu.Text)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="vIDChi"></param>
    ''' <param name="vIdDonVi"></param>
    ''' <param name="vLoaiCB">0: Lao động dài hạn; 1: Lao động ngắn hạn; 2: Tập sự</param>
    ''' <param name="vIdPhong"></param>
    ''' <param name="vNam"></param>
    ''' <param name="vThang"></param>
    ''' <param name="vKy"></param>
    ''' <remarks></remarks>
    Private Sub bindGrid(ByVal vIDChi As String, ByVal vLoaiThuChi As Integer, ByVal vIdDonVi As Integer, ByVal vLoaiCB As Int16, ByVal vIdPhong As String, ByVal vListCB As List(Of String), ByVal vNam As Integer, ByVal vThang As Integer, ByVal vKy As Int16)
        formatGrid(gridResult)
        If vLoaiThuChi = 1 Then
            gridResult.DataSource = createBangChiLuong(vIDChi, actLoaiCB, vIdDonVi, vIdPhong, vNam, vThang, vKy)
        Else
            If vLoaiThuChi = 2 Then
                gridResult.DataSource = createBangChiThemGio(vIDChi, vIdDonVi, vIdPhong, vNam, vThang)
            Else
                gridResult.DataSource = createBangChiKhac(vIDChi, vIdDonVi, vIdPhong, vLoaiCB, vListCB, vNam, vThang, CInt(cboLoaiLuongBS.SelectedValue), CDbl(txtHesochi.Text.ToString), N2Number(MoneyValue(txtSoTienBS.Text.ToString)))
            End If
        End If
        gridResult.Columns(0).Visible = False
        gridResult.Columns(1).Visible = False
        Dim i As Integer
        If actLoaiThuChi = 1 Then
            If actLoaiCB = 0 Then
                For i = 2 To gridResult.Columns.Count - 1
                    If i > 4 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            Else
                For i = 2 To gridResult.Columns.Count - 1
                    If i > 2 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            End If
        Else
            If actLoaiThuChi = 2 Then
                For i = 2 To gridResult.Columns.Count - 1
                    If i > 3 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            Else
                gridResult.Columns(2).Visible = False
                For i = 3 To gridResult.Columns.Count - 1
                    If i > 5 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            End If
        End If
    End Sub

    Private Sub showGrid(ByVal vIDChi As String)
        formatGrid(gridResult)
        gridResult.DataSource = showBangThuChi(vIDChi, actLoaiThuChi, actLoaiCB, actKy)
        gridResult.Columns(0).Visible = False
        gridResult.Columns(1).Visible = False
        Dim i As Integer
        If actLoaiThuChi = 1 Then
            If actLoaiCB = 0 Then
                For i = 2 To gridResult.Columns.Count - 1
                    If i > 4 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            Else
                For i = 2 To gridResult.Columns.Count - 1
                    If i > 2 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            End If
        Else
            If actLoaiThuChi = 2 Then
                For i = 2 To gridResult.Columns.Count - 1
                    If i > 3 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            Else
                gridResult.Columns(2).Visible = False
                For i = 3 To gridResult.Columns.Count - 1
                    If i > 5 Then
                        gridResult.Columns(i).DefaultCellStyle.Format = "n"
                        gridResult.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    gridResult.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next
            End If
        End If
        
    End Sub

    Private Sub formatGrid(ByRef objGrid As DataGridView)
        objGrid.AutoGenerateColumns = True
        objGrid.AllowUserToAddRows = False
        objGrid.AllowUserToDeleteRows = False
        objGrid.AllowUserToOrderColumns = False
        objGrid.AllowUserToResizeRows = False
        objGrid.RowHeadersVisible = False
        objGrid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken
        objGrid.ColumnHeadersHeight = 100
        objGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        objGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        objGrid.AllowUserToAddRows = False
        objGrid.Columns().Clear()
        objGrid.DataSource = Nothing
        objGrid.Refresh()

        'objGrid.Columns(18).ValueType = Type.GetType("System.Decimal")
        'objGrid.Columns(18).DefaultCellStyle.Format = "n"


    End Sub

    Private Sub ReLoad_frmChiLuong()
        frmUCL.Dispose()
        showGrid(frmUCL.IdChiLuong)
    End Sub

    Private Sub ReLoad_frmChiLuong_DanhSachCB()
        frmDSCB.Dispose()
        actLoaiCB = frmDSCB.LoaiCB
        listCB = frmDSCB.listCB
        setDanhsachCb(frmDSCB.listCB)
        'listCBs = frmDSCB.listCBs
        'setDSCB(frmDSCB.listCBs)
    End Sub

    Private Sub setDSCB(ByVal vListCB As String)
        Dim arr() As String
        Dim tmp As String = ""
        Dim Canbo As String = ""

        If vListCB <> "" Then
            tmp = vListCB & ","
            While tmp <> ""
                Canbo = tmp.Substring(0, tmp.IndexOf(","))
                arr = Canbo.Split("_")
                txtGroupCB.Text = IIf(txtGroupCB.Text.Trim = "", "", txtGroupCB.Text & ", ") & getFullName(arr(1), IIf(arr(0) = "TS", True, False))
                tmp = tmp.Substring(tmp.IndexOf(","))
            End While
        End If
    End Sub

    Private Sub setDanhsachCb(ByVal list As List(Of String))
        Dim i As Integer
        Dim arr() As String
        Dim Canbo As String = ""
        Dim arrDel As ArrayList = New ArrayList()
        Dim arrTS As ArrayList = New ArrayList()
        txtGroupCB.Text = ""
        If Not (list Is Nothing) Then
            If list.Count > 0 Then
                For i = 0 To list.Count - 1
                    Canbo = list.Item(i)
                    arr = Canbo.Split("_")
                    txtGroupCB.Text = IIf(txtGroupCB.Text.Trim = "", "", txtGroupCB.Text & ", ") & getFullName(arr(1), IIf((arr(0) = "TS" Or arr(0) = "NH"), True, False))
                Next
            End If
        End If

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="create"></param>
    ''' <param name="delete"></param>
    ''' <param name="confirm"></param>
    ''' <param name="print"></param>
    ''' <param name="status">1: Load form; 2: Tạo mới; 3: Lập bảng; 4: Kiểm soát</param>
    ''' <remarks></remarks>
    Private Sub ButtonDisplayed(ByRef create As Button, ByRef delete As Button, ByRef confirm As Button, ByRef print As Button, ByVal status As Integer)
        If status = 1 Then
            create.Visible = False
            delete.Visible = False
            confirm.Visible = False
            print.Visible = False
            Exit Sub
        End If
        If status = 2 Then
            create.Visible = True
            delete.Visible = False
            confirm.Visible = False
            print.Visible = False
            Exit Sub
        End If
        If status = 3 Then
            create.Visible = False
            delete.Visible = True
            confirm.Visible = True
            print.Visible = True
            Exit Sub
        End If
        If status = 4 Then
            create.Visible = False
            delete.Visible = False
            confirm.Visible = False
            print.Visible = True
            Exit Sub
        End If
    End Sub

    Private Sub setVarValueActived(ByVal vIdChiLuong As String, ByVal vLoaiThuChi As Integer, ByVal vIdDonVi As Integer, ByVal vNam As Integer, ByVal vThang As Integer, ByVal vKy As Int16, ByVal vLoaiCB As Int16)
        actIdChiLuong = vIdChiLuong
        actLoaiThuChi = vLoaiThuChi
        actIdDonVi = vIdDonVi
        actNam = vNam
        actThang = vThang
        actKy = vKy
        actLoaiCB = vLoaiCB
    End Sub

    Private Sub blankFrm()
        If DONVI = gMaDonViTW Then
            labGD.Text = "Phó Tổng giám đốc"
        End If
        cboPhong.SelectedIndex = 0
        cboLoaiChi.SelectedIndex = 0
        pnlGroupCB.Enabled = False
        pnlBoSung.Visible = False
        pnlLuong.Visible = True
        cboPhong.Enabled = True
        txtLB.Text = NGUOILAPBIEU
        txtKS.Text = KETOANTRUONG
        txtGD.Text = PHOGIAMDOC
        dtpkNam.Value = Now
        labAlertChiLuong.Text = ""
        labActived.Text = ""
        gridResult.DataSource = Nothing
    End Sub

    Private Function getActIdDonVi(ByVal vIdPhong As String) As Integer
        Dim vkindPB As String = ""
        Dim vIDPB As Integer = 0
        If vIdPhong <> "" Then
            vkindPB = vIdPhong.Substring(0, 2)
            vIDPB = vIdPhong.Substring(3)
            If vkindPB = "PB" Then
                actIdDonVi = IdDONVI
            Else
                actIdDonVi = vIDPB
            End If
        End If
        Return actIdDonVi
    End Function

    Private Function countChiLuong(ByVal IdLoaiThuChi As Integer, ByVal vIdDonVi As Integer, ByVal vIdPhong As String, ByVal vNam As Integer, ByVal vThang As Integer, ByVal vKy As Int16, ByVal vNH_TS As Int16, ByVal KiemSoat As Boolean) As Integer
        Dim strSql As String
        Dim returnValue As Integer = 0
        Dim db As DBAccess = New DBAccess
        Try

            If IdLoaiThuChi = 1 Then
                If vIdPhong = "PB_0" Then
                    strSql = "SELECT count(distinct a.IdHS_ChiLuong) FROM HS_ChiLuong a left join HS_ChiLuong_CT b on a.IdHS_ChiLuong=b.IdHS_ChiLuong WHERE a.IdDonVi=" & vIdDonVi & " AND a.Nam=" & vNam & " AND b.idPhong<>'PB_0' " & IIf(KiemSoat, "  AND a.tmp=0 ", "") & " AND a.Thang=" & vThang & " AND b.Ky=" & vKy & " AND b.NH_TS=" & vNH_TS
                Else
                    strSql = "SELECT count(distinct a.IdHS_ChiLuong) FROM HS_ChiLuong a left join HS_ChiLuong_CT b on a.IdHS_ChiLuong=b.IdHS_ChiLuong WHERE a.IdDonVi=" & vIdDonVi & " AND a.Nam=" & vNam & " AND (b.Idphong='" & vIdPhong & "' or b.idPhong='PB_0') AND a.Thang=" & vThang & IIf(KiemSoat, "  AND a.tmp=0 ", "") & " AND b.Ky=" & vKy & " AND b.NH_TS=" & vNH_TS
                End If
            Else
                If IdLoaiThuChi = 1 Then
                    If vIdPhong = "PB_0" Then
                        strSql = "SELECT count(distinct a.IdHS_ChiLuong) FROM HS_ChiLuong a left join HS_ChiThemGio_CT b on a.IdHS_ChiLuong=b.IdHS_ChiLuong WHERE a.IdDonVi=" & vIdDonVi & " AND a.Nam=" & vNam & " AND b.Idphong<>'PB_0' AND a.Thang=" & vThang & IIf(KiemSoat, "  AND a.tmp=0 ", "")
                    Else
                        strSql = "SELECT count(distinct a.IdHS_ChiLuong) FROM HS_ChiLuong a left join HS_ChiThemGio_CT b on a.IdHS_ChiLuong=b.IdHS_ChiLuong WHERE a.IdDonVi=" & vIdDonVi & " AND a.Nam=" & vNam & " AND (b.Idphong='" & vIdPhong & "' or b.idPhong='PB_0') AND a.Thang=" & vThang & IIf(KiemSoat, "  AND a.tmp=0 ", "")
                    End If
                Else
                    strSql = ""
                End If
            End If
            returnValue = db.getNumber(strSql)
            Return returnValue
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Sub rdDH_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdDH.CheckedChanged
        If rdDH.Checked Then actLoaiCB = 0
    End Sub

    Private Sub rdNH_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdNH.CheckedChanged
        If rdNH.Checked Then actLoaiCB = 1
    End Sub

    Private Sub rdTS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdTS.CheckedChanged
        If rdTS.Checked Then actLoaiCB = 2
    End Sub

    Private Sub cbGroupCB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbGroupCB.Click
        If cbGroupCB.Checked Then
            cboPhong.Enabled = False
            txtGroupCB.Enabled = True
            lnkChoiseCB.Enabled = True
        Else
            cboPhong.Enabled = True
            txtGroupCB.Enabled = False
            lnkChoiseCB.Enabled = False
            listCB = Nothing
            'listCBs = ""
        End If
    End Sub

    Private Sub lnkChoiseCB_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkChoiseCB.LinkClicked
        frmDSCB = New frmChonDS_CB
        frmDSCB.Progress_Changed = New frmChonDS_CB.ProgressChangedEventHandler(AddressOf ReLoad_frmChiLuong_DanhSachCB)
        frmDSCB.ShowDialog()
    End Sub

   
End Class