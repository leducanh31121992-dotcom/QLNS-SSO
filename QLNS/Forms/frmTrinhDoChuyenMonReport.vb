Public Class frmTrinhDoChuyenMonReport

    Public flagReport As Integer
    Private _Reports As clsHS_BaoCao = New clsHS_BaoCao()
    Private Sub frmTrinhDoChuyenMonReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True, False, True)
        txtLapBieu.Text = NGUOILAPBIEU
        txtGD.Text = GIAMDOC
        txtHCTC.Text = TRUONGHCTC
        txtLapBieu.Text = NGUOILAPBIEU
        If DONVI = gMaDonViTW Then
            labGD.Text = "Tổng giám đốc"
            labHCTC.Text = "Giám đốc Ban TCCB"

            cbAll.Visible = True
            cbAll.Enabled = True
        Else
            cbAll.Visible = False
        End If
        dpkThoiDiemBC.Value = "12/31/" & Year(Now)
        cboDonVi.Focus()
    End Sub

    Private Sub frmTrinhDoChuyenMonReport_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
        Try
            Dim lab_Err As String = ""
            Dim all As Short
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            lab_Err = checkRpt()
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                labStatusProcess.Text = "Error"
                Cursor = Cursors.Default
                Exit Sub
            End If
            If cbAll.Checked Then
                all = 1
            Else
                all = 0
            End If
            createReport(CInt(cboDonVi.SelectedValue), dpkThoiDiemBC.Value, all)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub createReport(ByVal vIdDonVi As Integer, ByVal vNgayBaoCao As Date, ByVal vAll As Short)
        'Dim N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14 As Integer
        'rpt_BC09.SetDataSource(getBC09(vIdDonVi, vNgayBaoCao, vAll, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14))


        Dim db_bc09 As DataTable = New DataTable()
        If vAll = 1 And getDonvi_Ma(CInt(cboDonVi.SelectedValue)) = gMaDonViTW Then
            vIdDonVi = 0
        End If
        db_bc09 = _Reports.GetAll_Report09_TKCD(vIdDonVi, 1, vNgayBaoCao)
        db_bc09.TableName = "BC09"
        rpt_BC09.SetDataSource(db_bc09)
        If Not (db_bc09 Is Nothing) Then
            If db_bc09.Rows.Count > 0 Then
                rpt_BC09.SetParameterValue("TC3", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C3").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC4", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C4").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC5", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C5").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC6", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C6").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC7", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C7").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC8", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C8").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC9", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C9").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC10", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C10").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC11", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C11").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC12", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C12").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC13", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C13").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC14", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C14").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC15", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C15").ToString().Trim(), Integer))
                rpt_BC09.SetParameterValue("TC16", CType(db_bc09.Rows(db_bc09.Rows.Count - 1)("C16").ToString().Trim(), Integer))
            End If
        End If


        'Dim N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14 As Integer
        'If flagReport = 1 Then
        '    rpt_BC09.SetDataSource(getBC09(vIdDonVi, vNgayBaoCao, vAll, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14))
        'Else
        '    rpt_BC09.SetDataSource(getBC09_NN(vIdDonVi, vNgayBaoCao, vAll, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13, N14))
        'End If


        'rpt_BC09.SetParameterValue("TC3", N1)
        'rpt_BC09.SetParameterValue("TC4", N2)
        'rpt_BC09.SetParameterValue("TC5", N3)
        'rpt_BC09.SetParameterValue("TC6", N4)
        'rpt_BC09.SetParameterValue("TC7", N5)
        'rpt_BC09.SetParameterValue("TC8", N6)
        'rpt_BC09.SetParameterValue("TC9", N7)
        'rpt_BC09.SetParameterValue("TC10", N8)
        'rpt_BC09.SetParameterValue("TC11", N9)
        'rpt_BC09.SetParameterValue("TC12", N10)
        'rpt_BC09.SetParameterValue("TC13", N11)
        'rpt_BC09.SetParameterValue("TC14", N12)
        'rpt_BC09.SetParameterValue("TC15", N13)
        'rpt_BC09.SetParameterValue("TC16", N14)

        If vAll = 1 And getDonvi_Ma(CInt(cboDonVi.SelectedValue)) = gMaDonViTW Then
            rpt_BC09.SetParameterValue("tenchinhanh", "")
        Else
            rpt_BC09.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
        End If
        rpt_BC09.SetParameterValue("tinh", DIABAN)
        rpt_BC09.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_BC09.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_BC09.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        If DONVI = gMaDonViTW Then
            rpt_BC09.SetParameterValue("labGD", "Tổng giám đốc")
            rpt_BC09.SetParameterValue("labHCTC", "Giám đốc Ban TCCB")
        Else
            If DONVI = "000196" Then
                rpt_BC09.SetParameterValue("labGD", "Giám đốc")
                rpt_BC09.SetParameterValue("labHCTC", "Trưởng phòng Tổng hợp")
            Else
                If DONVI = "000197" Or DONVI = "000101" Then
                    rpt_BC09.SetParameterValue("labGD", "Giám đốc")
                    rpt_BC09.SetParameterValue("labHCTC", "Trưởng phòng HC-NS")
                Else
                    rpt_BC09.SetParameterValue("labGD", "Giám đốc")
                    rpt_BC09.SetParameterValue("labHCTC", "Trưởng phòng HC-TC")
                End If
            End If
        End If
        rpt_BC09.SetParameterValue("LAPBIEU", txtLapBieu.Text)
        rpt_BC09.SetParameterValue("HCTC", txtHCTC.Text)
        rpt_BC09.SetParameterValue("GD", txtGD.Text)
        rpt_View.Show()
        rpt_View.Zoom(95)
        rpt_View.DisplayGroupTree = False
        rpt_View.ReportSource = rpt_BC09
        rpt_View.Refresh()

    End Sub

    Private Function checkRpt() As String
        Dim strReturn As String = ""
        Try
            If txtLapBieu.Text = "" Then
                txtLapBieu.Focus()
                strReturn = "Chưa nhập Người lập biểu!"
                Exit Try
            End If
            If txtHCTC.Text = "" Then
                txtHCTC.Focus()
                strReturn = "Chưa nhập Tên trưởng phòng HC - TC!"
                Exit Try
            End If
            If txtGD.Text = "" Then
                txtGD.Focus()
                strReturn = "Chưa nhập Tên giám đốc!"
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    'Private Sub cboDonVi_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDonVi.SelectedIndexChanged
    '    If cboDonVi.SelectedValue = 1 Then
    '        cbAll.Enabled = True
    '    Else
    '        cbAll.Checked = False
    '        cbAll.Enabled = False
    '    End If
    'End Sub
End Class