Public Class frmDS_TDKT

    Dim frmUKT As frmUpdateDeNghiKT

    Private Sub frmDS_TDKT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bntUpdate.Visible = False
        bntPrint.Enabled = False
        If DONVI = gMaDonViTW Then
            cboDonVi.DataSource = listDonvi(True, True, True)
        Else
            cboDonVi.DataSource = listDonvi(True, False, True)
        End If
    End Sub

    Private Sub bntCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCreate.Click
        Dim vKhenThuong As Int16 = 0
        Dim vDinhKy As Int16 = 0
        Dim vCaNhan As Int16 = 0
        bntPrint.Enabled = True
        If rdKhenThuong.Checked Then
            vKhenThuong = 1
        Else : vKhenThuong = 0
        End If
        If rdDK.Checked Then
            vDinhKy = 1
        Else : vDinhKy = 0
        End If
        If rdCN.Checked Then
            vCaNhan = 1
        Else : vCaNhan = 0
        End If
        gridResult.DataSource = Nothing
        gridResult.Columns.Clear()
        gridResult.Rows.Clear()
        'gridResult.AllowUserToOrderColumns = False
        'gridResult.AllowUserToAddRows = False
        'gridResult.AllowUserToDeleteRows = False
        'gridResult.AllowUserToResizeRows = False
        'gridResult.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        'If rdDK.Checked Then
        initGrid(gridResult, vDinhKy, CInt(cboDonVi.SelectedValue), vKhenThuong, CInt(dtpkNam.Text), KT_ChuyenMon, vCaNhan)
        'End If
    End Sub

    Private Sub addCol(ByRef vGrid As DataGridView, ByVal vType As Int16, ByVal vHeaderText As String, ByVal vDataPropertyName As String, ByVal vName As String)
        If vType = 1 Then
            Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            ckb_choice.Name = "cln_cbDuyet"
            ckb_choice.HeaderText = "Duyệt"
            ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            ckb_choice.Width = 50
            vGrid.Columns.Add(ckb_choice)
        Else
            Dim colTxt As New DataGridViewTextBoxColumn()
            colTxt.DataPropertyName = vDataPropertyName
            colTxt.HeaderText = vHeaderText
            colTxt.Name = vName
            colTxt.ReadOnly = True
            vGrid.Columns.Add(colTxt)
        End If
    End Sub

    Private Sub initGrid(ByRef gridName As DataGridView, ByVal vDinhKy As Int16, ByVal vIdDonVi As Integer, ByVal vKhenThuong As Int16, ByVal vNam As Integer, ByVal vKT_ChuyenMon As Int16, ByVal vCaNhan As Int16)
        Dim tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27 As Integer
        'If vDinhKy Then
        If vCaNhan Then
            'addCol(gridName, 1, "Duyệt", "C0", "C0")
            addCol(gridName, 0, "ID", "C28", "C28")
            addCol(gridName, 0, "IDDV", "C29", "C29")
            addCol(gridName, 0, "DV_PB", "C30", "C30")
            addCol(gridName, 0, "STT", "C0", "C0")
            addCol(gridName, 0, "Chi Nhánh", "C1", "C1")
            addCol(gridName, 0, "Ông/Bà", "C3", "C3")
            addCol(gridName, 0, "Họ tên", "C2", "C2")
            addCol(gridName, 0, "Chức vụ", "C4", "C4")
            addCol(gridName, 0, "Đơn vị làm việc", "C5", "C5")
            addCol(gridName, 0, "LĐTT", "C6", "C6")
            addCol(gridName, 0, "GK NHCS", "C10", "C10")
            addCol(gridName, 0, "CSTĐ CS", "C7", "C7")
            addCol(gridName, 0, "CSTĐ cấp ngành", "C8", "C8")
            addCol(gridName, 0, "CSTĐ toàn quốc", "C9", "C9")
            addCol(gridName, 0, "Bằng khen TTg", "C11", "C11")
            addCol(gridName, 0, "Bằng khen Thống đốc", "C12", "C12")
            addCol(gridName, 0, "Bằng khen Bộ, Ngành, đoàn thể TW", "C13", "C13")
            addCol(gridName, 0, "Bằng khen UBND tỉnh'", "C14", "C14")
            addCol(gridName, 0, "Kỷ niệm chương ngành", "C19", "C19")
            addCol(gridName, 0, "Kỷ niệm chương khác", "C20", "C20")
            addCol(gridName, 0, "Huân chương lao động hạng 1", "C17", "C17")
            addCol(gridName, 0, "Huân chương lao động hạng 2", "C16", "C16")
            addCol(gridName, 0, "Huân chương lao động hạng 3", "C15", "C15")
            addCol(gridName, 0, "Anh hùng lao động", "C18", "C18")
            addCol(gridName, 0, "Huân chương sao vàng", "C22", "C22")
            addCol(gridName, 0, "Huân chương Hồ Chí Minh", "C23", "C23")
            addCol(gridName, 0, "Huân chương độc lập hạng 1", "C26", "C26")
            addCol(gridName, 0, "Huân chương độc lập hạng 2", "C25", "C25")
            addCol(gridName, 0, "Huân chương độc lập hạng 3", "C24", "C24")
            addCol(gridName, 0, "Huy hiệu", "C21", "C21")
            addCol(gridName, 0, "Khác", "C27", "C27")
            gridName.Columns(0).Visible = False
            gridName.Columns(1).Visible = False
            gridName.Columns(2).Visible = False
            gridName.Columns(3).Width = 40
            If vIdDonVi <> 0 Then
                gridName.Columns(4).Visible = False
            Else
                gridName.Columns(4).Width = 200
            End If
            gridName.Columns(5).Width = 50
            gridName.Columns(6).Width = 180
            gridName.Columns(7).Width = 100
            gridName.Columns(8).Width = 180
            For i As Byte = 9 To 30
                gridName.Columns(i).ReadOnly = True
                gridName.Columns(i).Width = 60
                gridName.Columns(i).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Next
            gridName.DataSource = getDSKhenThuong(vIdDonVi, vKhenThuong, vDinhKy, vNam, vKT_ChuyenMon, vCaNhan, tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27)

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Khoi tao gridDataview

            'gridName.Columns.Add("C0", "STT")
            'gridName.Columns.Add("C1", "Chi Nhánh")
            'gridName.Columns.Add("C3", "Ông/Bà")
            'gridName.Columns.Add("C2", "Họ tên")
            'gridName.Columns.Add("C4", "Chức vụ")
            'gridName.Columns.Add("C5", "Đơn vị làm việc")
            'gridName.Columns.Add("C6", "LĐTT")
            'gridName.Columns.Add("C10", "GK NHCS")
            'gridName.Columns.Add("C7", "CSTĐ CS")
            'gridName.Columns.Add("C8", "CSTĐ cấp ngành")
            'gridName.Columns.Add("C9", "CSTĐ toàn quốc")
            'gridName.Columns.Add("C11", "Bằng khen TTg")
            'gridName.Columns.Add("C12", "Bằng khen Thống đốc")
            'gridName.Columns.Add("C13", "Bằng khen Bộ, Ngành, đoàn thể TW")
            'gridName.Columns.Add("C14", "Bằng khen UBND tỉnh'")
            'gridName.Columns.Add("C19", "Kỷ niệm chương ngành")
            'gridName.Columns.Add("C20", "Kỷ niệm chương khác")
            'gridName.Columns.Add("C17", "Huân chương lao động hạng 1")
            'gridName.Columns.Add("C16", "Huân chương lao động hạng 2")
            'gridName.Columns.Add("C15", "Huân chương lao động hạng 3")
            'gridName.Columns.Add("C18", "Anh hùng lao động")
            'gridName.Columns.Add("C22", "Huân chương sao vàng")
            'gridName.Columns.Add("C23", "Huân chương Hồ Chí Minh")
            'gridName.Columns.Add("C26", "Huân chương độc lập hạng 1")
            'gridName.Columns.Add("C25", "Huân chương độc lập hạng 2")
            'gridName.Columns.Add("C24", "Huân chương độc lập hạng 3")
            'gridName.Columns.Add("C21", "Huy hiệu")
            'gridName.Columns.Add("C27", "Khác")
            ''Dim u As ctlTest = New ctlTest
            ''gridName.Columns("C6").DataGridView.Controls.Add(u)
            ''gridName.Columns("C7").DataGridView.Controls.Add(u)

            'If vIdDonVi <> 0 Then
            '    gridName.Columns(1).Visible = False
            'Else
            '    gridName.Columns(1).Width = 200
            'End If
            'gridName.Columns(2).Width = 80
            'gridName.Columns(3).Width = 180
            'gridName.Columns(4).Width = 100
            'gridName.Columns(5).Width = 150
            'For i As Byte = 6 To 27
            '    gridName.Columns(i).ReadOnly = True
            '    gridName.Columns(i).Width = 60
            '    gridName.Columns(i).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            'Next
            'Dim dt As DataTable = New DataTable
            'dt = getDSKhenThuong(vIdDonVi, vKhenThuong, vNam, vKT_ChuyenMon, vCaNhan, vView)
            'If dt.Rows.Count > 1 Then
            '    Dim ckb_choice As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            '    ckb_choice.Name = "cln_cbDuyet"
            '    ckb_choice.HeaderText = "Duyệt"
            '    ckb_choice.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            '    ckb_choice.CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            '    ckb_choice.Width = 50
            '    Dim u As ctlTest = New ctlTest
            '    gridName.Columns("C20").DataGridView.Controls.Add(u)
            '    Dim tableStyle As DataGridTableStyle = New DataGridTableStyle()
            '    For i As Integer = 0 To dt.Rows.Count - 1
            '        gridName.Rows.Add()
            '        gridName.Rows(i).Cells("C0").Value = dt.Rows(i)("C0").ToString().Trim()
            '        gridName.Rows(i).Cells("C1").Value = dt.Rows(i)("C1").ToString().Trim()
            '        gridName.Rows(i).Cells("C2").Value = dt.Rows(i)("C2").ToString().Trim()
            '        gridName.Rows(i).Cells("C3").Value = dt.Rows(i)("C3").ToString().Trim()
            '        gridName.Rows(i).Cells("C4").Value = dt.Rows(i)("C4").ToString().Trim()
            '        gridName.Rows(i).Cells("C5").Value = dt.Rows(i)("C5").ToString().Trim()


            '        'gridName.Controls.Add(u)
            '        'gridName.Rows(i).Cells("C6").DataGridView.Controls.Add(u)
            '        'IIf(dt.Rows(i)("C6").ToString().Trim() = "!", gridName.Rows(i).Cells("C6")., )
            '        'gridName.Rows(i).Cells("C6").Value = dt.Rows(i)("C6").ToString().Trim()
            '        'gridName.Rows(i).Cells("C7").Value = dt.Rows(i)("C7").ToString().Trim()

            '        'Dim Access As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn
            '        'Access.HeaderText = dt.Rows(i)("C6").ToString.Trim
            '        'Dim cell As DataGridViewCheckBoxCell = New DataGridViewCheckBoxCell

            '        gridName.Rows(i).Cells("C6").DataGridView.Controls.Add(u)
            '        gridName.Rows(i).Cells("C7").DataGridView.Controls.Add(u)
            '        gridName.Rows(i).Cells("C8").DataGridView.Controls.Add(u)
            '        gridName.Rows(i).Cells("C9").DataGridView.Controls.Add(u)
            '        'cell.h = dt.Rows(i)("C6").ToString.Trim
            '        'gridName.Columns.Insert(7, Access)
            '        'gridName.Columns.Add("C6", Access)
            '        'gridName.Columns("C6").MinimumWidth = 40
            '        'gridName.Columns("C6").Name = "Access"


            '        'gridName.Rows(i).Cells("C8").Value = dt.Rows(i)("C8").ToString().Trim()
            '        'gridName.Rows(i).Cells("C9").Value = dt.Rows(i)("C9").ToString().Trim()
            '        'gridName.Rows(i).Cells("C10").Value = dt.Rows(i)("C10").ToString().Trim()

            '        'gridName.Rows(i).Cells("C11").Value = dt.Rows(i)("C11").ToString().Trim()
            '        'gridName.Rows(i).Cells("C12").Value = dt.Rows(i)("C12").ToString().Trim()
            '        'gridName.Rows(i).Cells("C13").Value = dt.Rows(i)("C13").ToString().Trim()
            '        'gridName.Rows(i).Cells("C14").Value = dt.Rows(i)("C14").ToString().Trim()
            '        'gridName.Rows(i).Cells("C15").Value = dt.Rows(i)("c15").ToString().Trim()




            '    Next

            'End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else
            addCol(gridName, 0, "ID", "C28", "C28")
            addCol(gridName, 0, "IDDV", "C29", "C29")
            addCol(gridName, 0, "DV_PB", "C30", "C30")
            addCol(gridName, 0, "STT", "C0", "C0")
            addCol(gridName, 0, "Chi Nhánh", "C1", "C1")
            addCol(gridName, 0, "Tập thể", "C5", "C5")
            addCol(gridName, 0, "LĐTT", "C6", "C6")
            addCol(gridName, 0, "LĐXS", "C7", "C7")
            addCol(gridName, 0, "Đơn vị quyết thắng", "C8", "C8")
            addCol(gridName, 0, "Cờ thi đua ngành NH", "C9", "C9")
            addCol(gridName, 0, "Cờ thi đua của CP", "C10", "C10")
            addCol(gridName, 0, "GK NHCS", "C11", "C11")
            addCol(gridName, 0, "Bằng khen TTg", "C12", "C12")
            addCol(gridName, 0, "Bằng khen Thống đốc", "C13", "C13")
            addCol(gridName, 0, "Bằng khen Bộ, Ngành, đoàn thể TW", "C14", "C14")
            addCol(gridName, 0, "Bằng khen UBND tỉnh'", "C15", "C15")
            addCol(gridName, 0, "Huân chương lao động hạng 1", "C18", "C18")
            addCol(gridName, 0, "Huân chương lao động hạng 2", "C17", "C17")
            addCol(gridName, 0, "Huân chương lao động hạng 3", "C16", "C16")
            addCol(gridName, 0, "Anh hùng lao động", "C19", "C19")
            addCol(gridName, 0, "Huân chương sao vàng", "C20", "C20")
            addCol(gridName, 0, "Huân chương Hồ Chí Minh", "C21", "C21")
            addCol(gridName, 0, "Huân chương độc lập hạng 1", "C24", "C24")
            addCol(gridName, 0, "Huân chương độc lập hạng 2", "C23", "C23")
            addCol(gridName, 0, "Huân chương độc lập hạng 3", "C22", "C22")
            addCol(gridName, 0, "Khác", "C27", "C27")
            gridName.Columns(0).Visible = False
            gridName.Columns(1).Visible = False
            gridName.Columns(2).Visible = False
            gridName.Columns(3).Width = 40
            If vIdDonVi <> 0 Then
                gridName.Columns(4).Visible = False
            Else
                gridName.Columns(4).Width = 200
            End If
            gridName.Columns(5).Width = 240
            For i As Byte = 6 To 25
                gridName.Columns(i).ReadOnly = True
                gridName.Columns(i).Width = 60
                gridName.Columns(i).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Next
            gridName.DataSource = getDSKhenThuong(vIdDonVi, vKhenThuong, vDinhKy, vNam, vKT_ChuyenMon, vCaNhan, tC6, tC7, tC8, tC9, tC10, tC11, tC12, tC13, tC14, tC15, tC16, tC17, tC18, tC19, tC20, tC21, tC22, tC23, tC24, tC25, tC26, tC27)
        End If
        'End If


    End Sub

    Private Sub ReLoad_frmKhenThuong()
        frmUKT.Dispose()
        Dim vKhenThuong As Int16 = 0
        Dim vDinhKy As Int16 = 0
        Dim vCaNhan As Int16 = 0
        If rdKhenThuong.Checked Then
            vKhenThuong = 1
        Else : vKhenThuong = 0
        End If
        If rdDK.Checked Then
            vDinhKy = 1
        Else : vDinhKy = 0
        End If
        If rdCN.Checked Then
            vCaNhan = 1
        Else : vCaNhan = 0
        End If
        gridResult.DataSource = Nothing
        gridResult.Columns.Clear()
        gridResult.Rows.Clear()
        'If rdDK.Checked Then
        initGrid(gridResult, vDinhKy, CInt(cboDonVi.SelectedValue), vKhenThuong, CInt(dtpkNam.Text), KT_ChuyenMon, vCaNhan)
        'End If
    End Sub

    Private Sub bntUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate.Click
        Try
            Dim arrDuyet As ArrayList = New ArrayList()
            If (gridResult.Rows.Count > 0) Then
                For i As Int32 = 0 To gridResult.Rows.Count - 1
                    If (CType(gridResult.Rows(i).Cells("cln_cbDuyet").Value, Boolean) = True) Then
                        arrDuyet.Add(gridResult.Rows(i).Cells("ID").Value.ToString())
                    End If
                Next
            End If
            Dim k As Integer = 0
            'If (arrDuyet.Count > 0) Then
            '    Try
            '        Dim IDDV_Moi, IDP_Moi, IDDV_Cu, IDP_Cu As Integer
            '        Dim m_QDNhanSuFinal As QDNhanSu = New QDNhanSu
            '        m_QDNhanSuFinal = m_QDNhanSuFinal.getFinalRecord(idCanBo)
            '        IDDV_Cu = m_QDNhanSuFinal.IdDonvi_Moi
            '        IDP_Cu = m_QDNhanSuFinal.IdPhong_Moi
            '        Dim i As Integer
            '        For i = 0 To arrDel.Count - 1
            '            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
            '            m_QDNhanSu.IdQDNhanSu = arrDel(i).ToString()
            '            m_QDNhanSu.Delete()
            '        Next
            '        m_QDNhanSuFinal = New QDNhanSu
            '        m_QDNhanSuFinal = m_QDNhanSuFinal.getFinalRecord(idCanBo)
            '        IDDV_Moi = m_QDNhanSuFinal.IdDonvi_Moi
            '        IDP_Moi = m_QDNhanSuFinal.IdPhong_Moi

            '        Dim childNode As TreeNode
            '        Dim _node As TreeNode
            '        Dim isExits As Boolean = False

            '        _node = treeCocau.SelectedNode.Parent
            '        refreshNodeNOTCanbo(treeCocau, _node, IDDV_Cu, IDP_Cu, False, False, idCanBo)
            '        For Each childNode In treeCocau.Nodes(0).Nodes
            '            Dim arr1() As String
            '            Dim arr0() As String
            '            arr1 = childNode.Tag.ToString.Split("_")
            '            arr0 = childNode.Parent.Tag.ToString.Split("_")
            '            If (arr0(1) = IDDV_Moi And arr1(2) = IDP_Moi) Then
            '                refreshNodeNOTCanbo(treeCocau, childNode, arr0(1), arr1(2), False, False, idCanBo)
            '                isExits = True
            '                Exit For
            '            End If
            '        Next
            '        If Not isExits Then
            '            blankOverInfCB()
            '            gridQDNhanSu.DataSource = Nothing
            '        Else
            '            bindGridQDNhansu(idCanBo)
            '            fillQDNhansu(idQDNhansu, idCanBo)
            '        End If
            '        labAlert_QD.Text = "Xoá dữ liệu thành công!"
            '    Catch ex As Exception
            '        MessageBox.Show("Xoá dữ liệu không thành công: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            '    End Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub bntPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntPrint.Click
        Dim frm As frmReport_KhenThuong = New frmReport_KhenThuong
        frm.KhenThuong = IIf(rdKhenThuong.Checked, 1, 0)
        frm.DinhKy = IIf(rdDK.Checked, 1, 0)
        frm.CaNhan = IIf(rdCN.Checked, 1, 0)
        frm.IdDonVi = CInt(cboDonVi.SelectedValue)
        frm.KT_ChuyenMon = KT_ChuyenMon
        frm.Nam = CInt(dtpkNam.Text)
        'frm.LapBieu = txtLB.Text
        'frm.KiemSoat = txtKS.Text
        'frm.GiamDoc = txtGD.Text
        'frm.GhiChu = txtGhiChu.Text
        frm.ShowDialog()
    End Sub

    Private Sub bntCloseKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCloseKT.Click
        Close()
    End Sub

    Private Sub gridResult_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridResult.CellDoubleClick
        If CInt(rdKhenThuong.Checked) = 0 Then
            frmUKT = New frmUpdateDeNghiKT
            Dim vKhenThuong As Int16 = 0
            Dim vDinhKy As Int16 = 0
            If rdKhenThuong.Checked Then
                vKhenThuong = 1
            Else : vKhenThuong = 0
            End If
            If rdDK.Checked Then
                vDinhKy = 1
            Else : vDinhKy = 0
            End If
            If rdCN.Checked Then
                frmUKT.CN_TT = 1
            Else
                frmUKT.CN_TT = gridResult.CurrentRow.Cells("C30").Value.ToString
            End If
            frmUKT.IdDonVi = gridResult.CurrentRow.Cells("C29").Value.ToString
            frmUKT.IdDoiTuong = gridResult.CurrentRow.Cells("C28").Value.ToString
            frmUKT.Nam = CInt(dtpkNam.Text)
            frmUKT.DinhKy = vDinhKy
            frmUKT.Progress_Changed = New frmUpdateDeNghiKT.ProgressChangedEventHandler(AddressOf ReLoad_frmKhenThuong)
            frmUKT.ShowDialog()
        End If
    End Sub
End Class