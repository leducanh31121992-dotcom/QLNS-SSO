Public Class frmCanBoBiThiHanhKiLuatReport

    Private Sub frmCanBoBiThiHanhKiLuatReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmCanBoBiThiHanhKiLuatReport_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            lab_Err = checkRpt()
            If lab_Err <> "" Then
                MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                labStatusProcess.Text = "Error"
                Cursor = Cursors.Default
                Exit Sub
            End If
            Dim ky As Short
            Dim all As Short
            If rdKy1.Checked Then
                ky = 1
            Else
                ky = 2
            End If
            If cbAll.Checked Then
                all = 1
            Else
                all = 0
            End If
            createReport(CInt(cboDonVi.SelectedValue), CShort(txtNam.Text), ky, all)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNam.TextChanged
        txtNam.Text = Val(txtNam.Text.Trim)
    End Sub

    Private Sub createReport(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vKy As Short, ByVal vAll As Short)
        Dim thoidiem As String
        rpt_BC08.SetDataSource(getBC08(vIdDonVi, vYear, vKy, vAll))
        If vAll = 1 And getDonvi_Ma(CInt(cboDonVi.SelectedValue)) = gMaDonViTW Then
            rpt_BC08.SetParameterValue("tenchinhanh", "")
        Else
            rpt_BC08.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
        End If
        rpt_BC08.SetParameterValue("tinh", DIABAN)
        rpt_BC08.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_BC08.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_BC08.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        If vKy = 1 Then
            thoidiem = "30/6/" & txtNam.Text
        Else
            thoidiem = "31/12/" & txtNam.Text
        End If
        rpt_BC08.SetParameterValue("Thoidiem", thoidiem)
        If DONVI = gMaDonViTW Then
            rpt_BC08.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
            rpt_BC08.SetParameterValue("labHCTC", "TRƯỞNG PHÒNG TCCB")
        Else
            rpt_BC08.SetParameterValue("labGD", "GIÁM ĐỐC")
            rpt_BC08.SetParameterValue("labHCTC", "TRƯỞNG PHÒNG HC-TC")
        End If
        rpt_BC08.SetParameterValue("LAPBIEU", txtLapBieu.Text)
        rpt_BC08.SetParameterValue("HCTC", txtHCTC.Text)
        rpt_BC08.SetParameterValue("GD", txtGD.Text)
        rpt_View.ReportSource = rpt_BC08
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