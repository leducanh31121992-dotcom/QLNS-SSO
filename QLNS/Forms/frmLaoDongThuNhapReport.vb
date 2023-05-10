Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports
Imports CrystalDecisions.Shared

Public Class frmLaoDongThuNhapReport
    Private da_Search As SqlDataAdapter
    Private _Id_BC05 As String
    Private _IdDonVi_BC05 As Integer
    Private _All_BC05 As Short
    Private _Nam_BC05 As Integer

    Public Property Id_BC05() As String
        Get
            Return _Id_BC05
        End Get
        Set(ByVal value As String)
            _Id_BC05 = value
        End Set
    End Property

    Public Property IdDonVi_BC05() As Integer
        Get
            Return _IdDonVi_BC05
        End Get
        Set(ByVal value As Integer)
            _IdDonVi_BC05 = value
        End Set
    End Property

    Public Property All_BC05() As Short
        Get
            Return _All_BC05
        End Get
        Set(ByVal value As Short)
            _All_BC05 = value
        End Set
    End Property

    Public Property Nam_BC05() As Integer
        Get
            Return _Nam_BC05
        End Get
        Set(ByVal value As Integer)
            _Nam_BC05 = value
        End Set
    End Property

    'Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
    '    Try
    '        Dim all As Short
    '        Dim lab_Err As String = ""
    '        Cursor = Cursors.WaitCursor
    '        labStatusProcess.Text = "Waiting ....."
    '        lab_Err = checkRpt()
    '        If lab_Err <> "" Then
    '            MessageBox.Show(lab_Err, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
    '            labStatusProcess.Text = "Error"
    '            Cursor = Cursors.Default
    '            Exit Sub
    '        End If
    '        If cbAll.Checked Then
    '            all = 1
    '        Else
    '            all = 0
    '        End If
    '        createReport(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), all)
    '        labStatusProcess.Text = "Done"
    '        Cursor = Cursors.Default
    '    Catch ex As Exception
    '        labStatusProcess.Text = "Error"
    '        Cursor = Cursors.Default
    '        MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
    '    End Try
    'End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub cmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
        Dim all As Short
        If CInt(txtNam.Text) < 2000 Then
            MessageBox.Show("Hãy nhập năm báo cáo", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        Else
            Try
                Cursor = Cursors.WaitCursor
                labStatusProcess.Text = "Waiting ....."
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
        End If
    End Sub

    Private Sub cmdCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreate.Click

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

    End Sub

    Private Sub getSoLaoDong(ByVal vIdDonVi As Integer, ByVal vNam As Integer, ByVal vAll As Short, ByRef vSoDaiHan As Integer, ByRef vSoNganHan As Integer)
        Try
            Dim strSql As String
            If vAll = 1 Then
                strSql = "SELECT sum(DaiHan) as DaiHan, sum(NganHan) as NganHan FROM KeHoachLD WHERE (idDonVi=" & vIdDonVi & " or idDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ") or idDonVi in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & "))) AND nam=" & vNam & " AND IdPhong=0  GROUP by IdDonvi, Nam"
            Else
                strSql = "SELECT TOP 1 DaiHan, NganHan FROM KeHoachLD WHERE idDonVi=" & vIdDonVi & " AND nam=" & vNam & " AND IdPhong=0 order by Thang desc"
            End If
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            Dim i As Integer
            dt = db.SelectDBRows(strSql)
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    vSoDaiHan = vSoDaiHan + dt.Rows(i).Item("DaiHan")
                    vSoNganHan = vSoNganHan + dt.Rows(i).Item("NganHan")
                Next
            End If
            
        Catch ex As Exception
        End Try
    End Sub

    Private Sub createReport(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short)
        Dim myBC05 As rptBC05
        Dim vSoKeHoachDaiHan As Integer
        Dim vSoKeHoachNganHan As Integer
        getSoLaoDong(vIdDonVi, vYear, vAll, vSoKeHoachDaiHan, vSoKeHoachNganHan)
        getBC05(vIdDonVi, vYear, vAll, vSoKeHoachDaiHan, vSoKeHoachNganHan, myBC05)
        rpt_BC05.DataDefinition.FormulaFields("ldtb_daihan").Text = myBC05.ldtb_daihan
        rpt_BC05.DataDefinition.FormulaFields("ldtb_nganhan").Text = myBC05.ldtb_nganhan
        rpt_BC05.DataDefinition.FormulaFields("ldth_daihan").Text = myBC05.ldth_daihan
        rpt_BC05.DataDefinition.FormulaFields("ldth_nganhan").Text = myBC05.ldth_nganhan
        rpt_BC05.DataDefinition.FormulaFields("luong_BHXH").Text = myBC05.luong_BHXH
        rpt_BC05.DataDefinition.FormulaFields("luong_bosung").Text = myBC05.luong_bosung
        rpt_BC05.DataDefinition.FormulaFields("luong_capbac").Text = myBC05.luong_capbac
        rpt_BC05.DataDefinition.FormulaFields("luong_chucvu").Text = myBC05.luong_chucvu
        rpt_BC05.DataDefinition.FormulaFields("luong_themgio").Text = myBC05.luong_themgio
        rpt_BC05.DataDefinition.FormulaFields("phucap").Text = myBC05.phucap
        rpt_BC05.DataDefinition.FormulaFields("tiencong_trongoi").Text = myBC05.tiencong_trongoi
        rpt_BC05.DataDefinition.FormulaFields("tongthunhap").Text = myBC05.tongthunhap
        rpt_BC05.DataDefinition.FormulaFields("thunhap_binhquan").Text = myBC05.thunhap_binhquan
        rpt_BC05.SetParameterValue("sobc", IIf(txtSoBC.Text.Trim = "", "        ", txtSoBC.Text.Trim))
        If vAll = 1 And getDonvi_Ma(CInt(cboDonVi.SelectedValue)) = gMaDonViTW Then
            rpt_BC05.SetParameterValue("tenchinhanh", "")
        Else
            rpt_BC05.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
        End If
        rpt_BC05.SetParameterValue("machinhanh", getDonvi_Ma(CInt(cboDonVi.SelectedValue)))
        rpt_BC05.SetParameterValue("tinh", DIABAN)
        rpt_BC05.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_BC05.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_BC05.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        rpt_BC05.SetParameterValue("nam", txtNam.Text)
        If DONVI = gMaDonViTW Then
            rpt_BC05.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
            rpt_BC05.SetParameterValue("labHCTC", "GIÁM ĐỐC BAN TCCB")
        Else
            rpt_BC05.SetParameterValue("labGD", "GIÁM ĐỐC")
            rpt_BC05.SetParameterValue("labHCTC", "TRƯỞNG PHÒNG HC-TC")
        End If
        rpt_BC05.SetParameterValue("lapbieu", txtLapBieu.Text)
        rpt_BC05.SetParameterValue("HCTC", txtHCTC.Text)
        rpt_BC05.SetParameterValue("GD", txtGD.Text)
        rpt_View.ReportSource = rpt_BC05
    End Sub

    Private Sub viewReport(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short)
        Dim myBC05 As rptBC05
        Dim m_bc05 As tbBC05 = New tbBC05
        m_bc05 = m_bc05.getRecord(Id_BC05)
        rpt_BC05.DataDefinition.FormulaFields("ldtb_daihan").Text = m_bc05.TB_LDDH
        rpt_BC05.DataDefinition.FormulaFields("ldtb_nganhan").Text = m_bc05.TB_LDNH
        rpt_BC05.DataDefinition.FormulaFields("ldth_daihan").Text = m_bc05.TH_LDDH
        rpt_BC05.DataDefinition.FormulaFields("ldth_nganhan").Text = m_bc05.TH_LDNH
        rpt_BC05.DataDefinition.FormulaFields("luong_BHXH").Text = m_bc05.TN_BHXH
        rpt_BC05.DataDefinition.FormulaFields("luong_bosung").Text = m_bc05.TN_BoSung
        rpt_BC05.DataDefinition.FormulaFields("luong_capbac").Text = m_bc05.L_CapBac
        rpt_BC05.DataDefinition.FormulaFields("luong_chucvu").Text = m_bc05.L_ChucVu
        rpt_BC05.DataDefinition.FormulaFields("luong_themgio").Text = m_bc05.TN_ThemGio
        rpt_BC05.DataDefinition.FormulaFields("phucap").Text = m_bc05.L_PhuCap
        rpt_BC05.DataDefinition.FormulaFields("tiencong_trongoi").Text = m_bc05.L_TronGoi
        Dim TongTN As Long = (m_bc05.L_CapBac + m_bc05.L_ChucVu + m_bc05.L_PhuCap + m_bc05.L_TronGoi + m_bc05.TN_BHXH + m_bc05.TN_BoSung + m_bc05.TN_ThemGio) * 1000
        rpt_BC05.DataDefinition.FormulaFields("tongthunhap").Text = TongTN
        rpt_BC05.DataDefinition.FormulaFields("thunhap_binhquan").Text = TongTN / ((m_bc05.TH_LDDH + m_bc05.TH_LDNH) * 12)
        rpt_BC05.SetParameterValue("sobc", IIf(txtSoBC.Text.Trim = "", "        ", txtSoBC.Text.Trim))
        If vAll = 1 And getDonvi_Ma(CInt(cboDonVi.SelectedValue)) = gMaDonViTW Then
            rpt_BC05.SetParameterValue("tenchinhanh", "")
        Else
            rpt_BC05.SetParameterValue("tenchinhanh", getDonvi(CInt(cboDonVi.SelectedValue)))
        End If
        rpt_BC05.SetParameterValue("machinhanh", getDonvi_Ma(CInt(cboDonVi.SelectedValue)))
        rpt_BC05.SetParameterValue("tinh", DIABAN)
        rpt_BC05.SetParameterValue("ng", dpkNgayLapBieu.Value.Day)
        rpt_BC05.SetParameterValue("th", dpkNgayLapBieu.Value.Month)
        rpt_BC05.SetParameterValue("nm", dpkNgayLapBieu.Value.Year)
        rpt_BC05.SetParameterValue("nam", txtNam.Text)
        If DONVI = gMaDonViTW Then
            rpt_BC05.SetParameterValue("labGD", "TỔNG GIÁM ĐỐC")
            rpt_BC05.SetParameterValue("labHCTC", "GIÁM ĐỐC BAN TCCB")
        Else
            rpt_BC05.SetParameterValue("labGD", "GIÁM ĐỐC")
            rpt_BC05.SetParameterValue("labHCTC", "TRƯỞNG PHÒNG HC-TC")
        End If
        rpt_BC05.SetParameterValue("lapbieu", txtLapBieu.Text)
        rpt_BC05.SetParameterValue("HCTC", txtHCTC.Text)
        rpt_BC05.SetParameterValue("GD", txtGD.Text)
        rpt_View.ReportSource = rpt_BC05
    End Sub


    Private Function checkRpt() As String
        Dim strReturn As String = ""
        Try
            'If txtSoBC.Text = "" Then
            '    txtSoBC.Focus()
            '    strReturn = "Chưa nhập Số báo cáo!"
            '    Exit Try
            'End If
            If txtNam.Text = "" Then
                txtNam.Focus()
                strReturn = "Chưa nhập Năm lấy báo cáo!"
                Exit Try
            End If
            'If txtDH.Text = "" Then
            '    txtDH.Focus()
            '    strReturn = "Chưa nhập Số lao động hợp đồng dài hạn được thông báo đến 31/12 của năm!"
            '    Exit Try
            'End If
            'If txtNH.Text = "" Then
            '    txtNH.Focus()
            '    strReturn = "Chưa nhập Số lao động hợp đồng ngắn hạn được thông báo đến 31/12 của năm!"
            '    Exit Try
            'End If
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

    Private Sub frmLaoDongThuNhapReport_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmLaoDongThuNhapReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True)
        cboDonVi.SelectedValue = IdDonVi_BC05
        cboDonVi.DroppedDown = False
        cbAll.Checked = All_BC05
        txtNam.Text = Nam_BC05
        txtNam.ReadOnly = True

        txtLapBieu.Text = NGUOILAPBIEU
        txtGD.Text = GIAMDOC
        txtHCTC.Text = TRUONGHCTC
        'txtNam.Text = Now.Year
        If DONVI = gMaDonViTW Then
            labGD.Text = "Tổng giám đốc"
            labHCTC.Text = "Trưởng phòng TCCB"
        End If
        cboDonVi.Focus()

        'viewReport(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), All_BC05)
        viewReport(IdDonVi_BC05, Nam_BC05, All_BC05)
        cmdView.Visible = False
        cmdCreate.Visible = False
        cmdSave.Visible = False
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNam.TextChanged
        txtNam.Text = Val(txtNam.Text.Trim)
    End Sub

    Private Sub cbAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAll.CheckedChanged
        cbAll.Checked = All_BC05
    End Sub

    Private Sub cboDonVi_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDonVi.SelectedValueChanged
        cboDonVi.SelectedValue = IdDonVi_BC05
    End Sub
End Class