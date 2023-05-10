Public Class frmChinhSachNguoiLaoDongReport

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
        Try
            Dim all As Short
            Dim lab_Err As String = ""
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
            createReport(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), all)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub frmChinhSachNguoiLaoDongReport_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmChinhSachNguoiLaoDongReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True)
        txtLapBieu.Text = NGUOILAPBIEU
        txtGD.Text = GIAMDOC
        txtHCTC.Text = TRUONGHCTC
        txtNam.Text = Now.Year
        If DONVI = gMaDonViTW Then
            labGD.Text = "Tổng giám đốc"
            labHCTC.Text = "Trưởng phòng TCCB"
        End If
        cboDonVi.Focus()
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNam.TextChanged
        txtNam.Text = Val(txtNam.Text.Trim)
    End Sub

    Private Sub createReport(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short)
        rpt_BC07.SetDataSource(getBC07(vIdDonVi, vYear, vAll))
        If vAll = 1 And getDonvi_Ma(CInt(cboDonVi.SelectedValue)) = gMaDonViTW Then
            rpt_BC07.SetParameterValue("tenchinhanh", "")
        Else
            rpt_BC07.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
        End If
        rpt_BC07.SetParameterValue("tinh", DIABAN)
        rpt_BC07.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_BC07.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_BC07.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        rpt_BC07.SetParameterValue("nam", txtNam.Text)
        If DONVI = gMaDonViTW Then
            rpt_BC07.SetParameterValue("labGD", "Tổng giám đốc")
            rpt_BC07.SetParameterValue("labHCTC", "Trưởng phòng TCCB")
        Else
            rpt_BC07.SetParameterValue("labGD", "Giám đốc")
            rpt_BC07.SetParameterValue("labHCTC", "Trưởng phòng HC-TC")
        End If
        rpt_BC07.SetParameterValue("LAPBIEU", txtLapBieu.Text)
        rpt_BC07.SetParameterValue("HCTC", txtHCTC.Text)
        rpt_BC07.SetParameterValue("GD", txtGD.Text)
        rpt_View.ReportSource = rpt_BC07
    End Sub

    Private Function checkRpt() As String
        Dim strReturn As String = ""
        Try
            If txtNam.Text = "" Then
                txtNam.Focus()
                strReturn = "Chưa nhập Năm lấy báo cáo!"
                Exit Try
            End If
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

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub
End Class