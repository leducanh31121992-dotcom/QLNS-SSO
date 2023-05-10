Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms

Public Class frmBangKeChiLuong

    Private Sub frmBangKeChiLuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True)
        bindCboPhong(CInt(cboDonVi.SelectedValue))
        txtThang.Text = Now.Month
        txtNam.Text = Now.Year
        txtLapBieu.Text = NGUOILAPBIEU
        txtGD.Text = GIAMDOC
        txtHCTC.Text = KETOANTRUONG
        cboDonVi.Focus()
        labKy.Text = "/ " & MAX_KY.ToString & " KỲ"
        If DONVI = gMaDonViTW Then
            labGD.Text = "Tổng giám đốc"
        End If
    End Sub

    Private Sub frmBangKeChiLuong_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " + ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub cmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
        Try
            Dim lab_Err As String = ""
            lab_Err = checkRpt()
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            If rdDH.Checked Then
                createReportChiLuong(CInt(cboDonVi.SelectedValue), CInt(cboPhong.SelectedValue), CInt(txtKy.Text), CInt(txtThang.Text), CInt(txtNam.Text))
            ElseIf rdNH.Checked Then
                createReportChiLuongHDNH(CInt(cboDonVi.SelectedValue), CInt(txtThang.Text), CInt(txtNam.Text))
            Else
                createReportChiLuongThuViec(CInt(cboDonVi.SelectedValue), CInt(txtThang.Text), CInt(txtNam.Text))
            End If
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub txtThang_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThang.TextChanged
        txtThang.Text = Val(txtThang.Text.Trim)
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNam.TextChanged
        txtNam.Text = Val(txtNam.Text.Trim)
    End Sub

    Private Sub bindCboPhong(ByVal vIdDonvi As Integer)
        Try
            If vIdDonvi > 0 Then
                cboPhong.DataSource = listPhong(vIdDonvi, True)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub createReportChiLuong(ByVal vIdDonVi As Integer, ByVal vIdPhongBan As Integer, ByVal vKy As Integer, ByVal vMonth As Integer, ByVal vYear As Integer)
        Dim C4, C5, C6, C7, C8, C9, C10, C11, C12, C13 As String
        Dim N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13 As Integer
        Dim dt As DataTable = New DataTable
        dt = getChiLuong(vIdDonVi, vIdPhongBan, vKy, vMonth, vYear, C4, C5, C6, C7, C8, C9, C10, C11, C12, C13, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11, N12, N13)
        If MAX_KY = 1 Then
            rpt_ChiLuong1KY.SetDataSource(dt)
            rpt_ChiLuong1KY.SetParameterValue("TC4", C4)
            rpt_ChiLuong1KY.SetParameterValue("TC5", C5)
            rpt_ChiLuong1KY.SetParameterValue("TC6", C6)
            rpt_ChiLuong1KY.SetParameterValue("TC7", C7)
            rpt_ChiLuong1KY.SetParameterValue("TC8", C8)
            rpt_ChiLuong1KY.SetParameterValue("TC9", C9)
            rpt_ChiLuong1KY.SetParameterValue("TC10", C10)
            rpt_ChiLuong1KY.SetParameterValue("TC11", C11)
            rpt_ChiLuong1KY.SetParameterValue("TC12", C12)
            rpt_ChiLuong1KY.SetParameterValue("TN1", N1)
            rpt_ChiLuong1KY.SetParameterValue("TN2", N2)
            rpt_ChiLuong1KY.SetParameterValue("TN3", N3)
            rpt_ChiLuong1KY.SetParameterValue("TN4", N4)
            rpt_ChiLuong1KY.SetParameterValue("TN5", N5)
            rpt_ChiLuong1KY.SetParameterValue("TN6", N6)
            rpt_ChiLuong1KY.SetParameterValue("TN8", N8)
            rpt_ChiLuong1KY.SetParameterValue("TN9", N9)
            rpt_ChiLuong1KY.SetParameterValue("TN10", N10)
            rpt_ChiLuong1KY.SetParameterValue("TN11", N11)
            rpt_ChiLuong1KY.SetParameterValue("TN12", N12)
            rpt_ChiLuong1KY.SetParameterValue("TN13", N13)
            rpt_ChiLuong1KY.SetParameterValue("DonVi", getDonvi(CInt(cboDonVi.SelectedValue)))
            rpt_ChiLuong1KY.SetParameterValue("Phong", getPhong(CInt(cboPhong.SelectedValue)))
            rpt_ChiLuong1KY.SetParameterValue("tinh", DIABAN)
            rpt_ChiLuong1KY.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
            rpt_ChiLuong1KY.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
            rpt_ChiLuong1KY.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
            rpt_ChiLuong1KY.SetParameterValue("thang", txtThang.Text)
            rpt_ChiLuong1KY.SetParameterValue("nam", txtNam.Text)
            If DONVI = gMaDonViTW Then
                rpt_ChiLuong1KY.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
            Else
                rpt_ChiLuong1KY.SetParameterValue("labGD", "GIÁM ĐỐC")
            End If
            rpt_ChiLuong1KY.SetParameterValue("LAPBANG", txtLapBieu.Text)
            rpt_ChiLuong1KY.SetParameterValue("KIEMSOAT", txtHCTC.Text)
            rpt_ChiLuong1KY.SetParameterValue("GD", txtGD.Text)
            rpt_ChiLuong1KY.SetParameterValue("GhiChu", txtGhiChu.Text)
            rpt_View.ReportSource = rpt_ChiLuong1KY
        Else
            If vKy = 1 Then
                rpt_ChiLuongKY1.SetDataSource(dt)
                rpt_ChiLuongKY1.SetParameterValue("TC4", C4)
                rpt_ChiLuongKY1.SetParameterValue("TC5", C5)
                rpt_ChiLuongKY1.SetParameterValue("TC6", C6)
                rpt_ChiLuongKY1.SetParameterValue("TC7", C7)
                rpt_ChiLuongKY1.SetParameterValue("TC8", C8)
                rpt_ChiLuongKY1.SetParameterValue("TC9", C9)
                rpt_ChiLuongKY1.SetParameterValue("TC10", C10)
                rpt_ChiLuongKY1.SetParameterValue("TC11", C11)
                rpt_ChiLuongKY1.SetParameterValue("TC12", C12)
                rpt_ChiLuongKY1.SetParameterValue("TN1", N1)
                rpt_ChiLuongKY1.SetParameterValue("TN2", N2)
                rpt_ChiLuongKY1.SetParameterValue("TN3", N3)
                rpt_ChiLuongKY1.SetParameterValue("TN4", N4)
                rpt_ChiLuongKY1.SetParameterValue("TN5", N5)
                rpt_ChiLuongKY1.SetParameterValue("TN6", N6)
                rpt_ChiLuongKY1.SetParameterValue("TN9", N9)
                rpt_ChiLuongKY1.SetParameterValue("TN10", N10)
                rpt_ChiLuongKY1.SetParameterValue("TN13", N13)
                rpt_ChiLuongKY1.SetParameterValue("DonVi", getDonvi(CInt(cboDonVi.SelectedValue)))
                rpt_ChiLuongKY1.SetParameterValue("Phong", getPhong(CInt(cboPhong.SelectedValue)))
                rpt_ChiLuongKY1.SetParameterValue("tinh", DIABAN)
                rpt_ChiLuongKY1.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
                rpt_ChiLuongKY1.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
                rpt_ChiLuongKY1.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
                rpt_ChiLuongKY1.SetParameterValue("Ky", txtKy.Text)
                rpt_ChiLuongKY1.SetParameterValue("thang", txtThang.Text)
                rpt_ChiLuongKY1.SetParameterValue("nam", txtNam.Text)
                If DONVI = gMaDonViTW Then
                    rpt_ChiLuongKY1.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                Else
                    rpt_ChiLuongKY1.SetParameterValue("labGD", "GIÁM ĐỐC")
                End If
                rpt_ChiLuongKY1.SetParameterValue("LAPBANG", txtLapBieu.Text)
                rpt_ChiLuongKY1.SetParameterValue("KIEMSOAT", txtHCTC.Text)
                rpt_ChiLuongKY1.SetParameterValue("GD", txtGD.Text)
                rpt_ChiLuongKY1.SetParameterValue("GhiChu", txtGhiChu.Text)
                rpt_View.ReportSource = rpt_ChiLuongKY1
            Else
                rpt_ChiLuongKY2.SetDataSource(dt)
                rpt_ChiLuongKY2.SetParameterValue("TC4", C4)
                rpt_ChiLuongKY2.SetParameterValue("TC5", C5)
                rpt_ChiLuongKY2.SetParameterValue("TC6", C6)
                rpt_ChiLuongKY2.SetParameterValue("TC7", C7)
                rpt_ChiLuongKY2.SetParameterValue("TC8", C8)
                rpt_ChiLuongKY2.SetParameterValue("TC9", C9)
                rpt_ChiLuongKY2.SetParameterValue("TC10", C10)
                rpt_ChiLuongKY2.SetParameterValue("TC11", C11)
                rpt_ChiLuongKY2.SetParameterValue("TC12", C12)
                rpt_ChiLuongKY2.SetParameterValue("TN1", N1)
                rpt_ChiLuongKY2.SetParameterValue("TN2", N2)
                rpt_ChiLuongKY2.SetParameterValue("TN7", N7)
                rpt_ChiLuongKY2.SetParameterValue("TN8", N8)
                rpt_ChiLuongKY2.SetParameterValue("TN9", N9)
                rpt_ChiLuongKY2.SetParameterValue("TN10", N10)
                rpt_ChiLuongKY2.SetParameterValue("TN11", N11)
                rpt_ChiLuongKY2.SetParameterValue("TN12", N12)
                rpt_ChiLuongKY2.SetParameterValue("DonVi", getDonvi(CInt(cboDonVi.SelectedValue)))
                rpt_ChiLuongKY2.SetParameterValue("Phong", getPhong(CInt(cboPhong.SelectedValue)))
                rpt_ChiLuongKY2.SetParameterValue("tinh", DIABAN)
                rpt_ChiLuongKY2.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
                rpt_ChiLuongKY2.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
                rpt_ChiLuongKY2.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
                rpt_ChiLuongKY2.SetParameterValue("Ky", txtKy.Text)
                rpt_ChiLuongKY2.SetParameterValue("thang", txtThang.Text)
                rpt_ChiLuongKY2.SetParameterValue("nam", txtNam.Text)
                If DONVI = gMaDonViTW Then
                    rpt_ChiLuongKY2.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
                Else
                    rpt_ChiLuongKY2.SetParameterValue("labGD", "GIÁM ĐỐC")
                End If
                rpt_ChiLuongKY2.SetParameterValue("LAPBANG", txtLapBieu.Text)
                rpt_ChiLuongKY2.SetParameterValue("KIEMSOAT", txtHCTC.Text)
                rpt_ChiLuongKY2.SetParameterValue("GD", txtGD.Text)
                rpt_ChiLuongKY2.SetParameterValue("GhiChu", txtGhiChu.Text)
                rpt_View.ReportSource = rpt_ChiLuongKY2
            End If
        End If

    End Sub

    Private Sub createReportChiLuongHDNH(ByVal vIdDonVi As Integer, ByVal vMonth As Integer, ByVal vYear As Integer)
        Dim N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11 As Integer
        Dim dt As DataTable = New DataTable
        dt = getChiLuongHDNH(vIdDonVi, vMonth, vYear, N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11)
        rpt_ChiLuongHDNH.SetDataSource(dt)
        rpt_ChiLuongHDNH.SetParameterValue("TN1", N1)
        rpt_ChiLuongHDNH.SetParameterValue("TN2", N2)
        rpt_ChiLuongHDNH.SetParameterValue("TN3", N3)
        rpt_ChiLuongHDNH.SetParameterValue("TN4", N4)
        rpt_ChiLuongHDNH.SetParameterValue("TN5", N5)
        rpt_ChiLuongHDNH.SetParameterValue("TN6", N6)
        rpt_ChiLuongHDNH.SetParameterValue("TN7", N7)
        rpt_ChiLuongHDNH.SetParameterValue("TN8", N8)
        rpt_ChiLuongHDNH.SetParameterValue("TN9", N9)
        rpt_ChiLuongHDNH.SetParameterValue("TN10", N10)
        rpt_ChiLuongHDNH.SetParameterValue("TN11", N11)
        rpt_ChiLuongHDNH.SetParameterValue("DonVi", getDonvi(CInt(cboDonVi.SelectedValue)))
        rpt_ChiLuongHDNH.SetParameterValue("tinh", DIABAN)
        rpt_ChiLuongHDNH.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_ChiLuongHDNH.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_ChiLuongHDNH.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        rpt_ChiLuongHDNH.SetParameterValue("thang", txtThang.Text)
        rpt_ChiLuongHDNH.SetParameterValue("nam", txtNam.Text)
        If DONVI = gMaDonViTW Then
            rpt_ChiLuongHDNH.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
        Else
            rpt_ChiLuongHDNH.SetParameterValue("labGD", "GIÁM ĐỐC")
        End If
        rpt_ChiLuongHDNH.SetParameterValue("LAPBANG", txtLapBieu.Text)
        rpt_ChiLuongHDNH.SetParameterValue("KIEMSOAT", txtHCTC.Text)
        rpt_ChiLuongHDNH.SetParameterValue("GD", txtGD.Text)
        rpt_ChiLuongHDNH.SetParameterValue("GhiChu", txtGhiChu.Text)
        rpt_View.ReportSource = rpt_ChiLuongHDNH

    End Sub

    Private Sub createReportChiLuongThuViec(ByVal vIdDonVi As Integer, ByVal vMonth As Integer, ByVal vYear As Integer)
        Dim C3 As String
        Dim N1, N2, N3, N4, N5, N6 As Integer
        Dim dt As DataTable = New DataTable
        dt = getChiLuongTS(vIdDonVi, vMonth, vYear, C3, N1, N2, N3, N4, N5, N6)
        rpt_ChiLuongTS.SetDataSource(dt)
        rpt_ChiLuongTS.SetParameterValue("TC3", C3)
        rpt_ChiLuongTS.SetParameterValue("TN1", N1)
        rpt_ChiLuongTS.SetParameterValue("TN2", N2)
        rpt_ChiLuongTS.SetParameterValue("TN3", N3)
        rpt_ChiLuongTS.SetParameterValue("TN4", N4)
        rpt_ChiLuongTS.SetParameterValue("TN5", N5)
        rpt_ChiLuongTS.SetParameterValue("TN6", N6)
        rpt_ChiLuongTS.SetParameterValue("DonVi", getDonvi(CInt(cboDonVi.SelectedValue)))
        rpt_ChiLuongTS.SetParameterValue("tinh", DIABAN)
        rpt_ChiLuongTS.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_ChiLuongTS.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_ChiLuongTS.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        rpt_ChiLuongTS.SetParameterValue("thang", txtThang.Text)
        rpt_ChiLuongTS.SetParameterValue("nam", txtNam.Text)
        If DONVI = gMaDonViTW Then
            rpt_ChiLuongTS.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
        Else
            rpt_ChiLuongTS.SetParameterValue("labGD", "GIÁM ĐỐC")
        End If
        rpt_ChiLuongTS.SetParameterValue("LAPBANG", txtLapBieu.Text)
        rpt_ChiLuongTS.SetParameterValue("KIEMSOAT", txtHCTC.Text)
        rpt_ChiLuongTS.SetParameterValue("GD", txtGD.Text)
        rpt_ChiLuongTS.SetParameterValue("GhiChu", txtGhiChu.Text)
        rpt_View.ReportSource = rpt_ChiLuongTS
    End Sub

    Private Function checkRpt() As String
        Dim strReturn As String = ""
        Try
            If txtKy.Text = "" Then
                txtKy.Focus()
                strReturn = "Chưa nhập kỳ lấy báo cáo!"
                Exit Try
            Else
                If CInt(txtKy.Text) > MAX_KY Then
                    txtKy.Focus()
                    strReturn = "Hãy nhập lại kỳ chi lương!"
                    Exit Try
                End If
            End If
            If txtThang.Text = "" Then
                txtThang.Focus()
                strReturn = "Chưa nhập Tháng lấy báo cáo!"
                Exit Try
            Else
                If txtThang.Text > 12 Then
                    txtThang.Focus()
                    strReturn = "Tháng không hợp lệ. Hãy nhập lại!"
                    Exit Try
                End If
            End If
            If txtNam.Text = "" Then
                txtNam.Focus()
                strReturn = "Chưa nhập Năm lấy báo cáo!"
                Exit Try
            Else
                If txtNam.Text < 2003 Then
                    txtNam.Focus()
                    strReturn = "Năm không hợp lệ. Hãy nhập lại!"
                    Exit Try
                End If
            End If
            If txtLapBieu.Text = "" Then
                txtLapBieu.Focus()
                strReturn = "Chưa nhập tên Người lập biểu!"
                Exit Try
            End If
            If txtHCTC.Text = "" Then
                txtHCTC.Focus()
                strReturn = "Chưa nhập tên Trưởng phòng HC - TC!"
                Exit Try
            End If
            If txtGD.Text = "" Then
                txtGD.Focus()
                If DONVI = gMaDonViTW Then
                    labGD.Text = "Chưa nhập tên Tổng giám đốc"
                Else
                    strReturn = "Chưa nhập tên Giám đốc!"
                End If
                Exit Try
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Sub cboDonVi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDonVi.SelectedIndexChanged
        bindCboPhong(CInt(cboDonVi.SelectedValue))
    End Sub

    Private Sub txtKy_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKy.TextChanged
        txtKy.Text = Val(txtKy.Text.Trim)
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub rdNH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdNH.CheckedChanged
        txtKy.Enabled = False
        labKy.Enabled = False
    End Sub

    Private Sub rdTS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdTS.CheckedChanged
        txtKy.Enabled = False
        labKy.Enabled = False
    End Sub

    Private Sub rdDH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdDH.CheckedChanged
        txtKy.Enabled = True
        labKy.Enabled = True
    End Sub

End Class