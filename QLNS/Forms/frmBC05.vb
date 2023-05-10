Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class frmBC05

    Public flagReport As Integer
    Private IdBC05_cur As String

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
                'labNam_TB.Text = txtNam.Text
                'labNam_TH.Text = txtNam.Text
                lbl_sold_thongbao.Text = "          1. Số lao động được thông báo đến 31/12/" + txtNam.Text
                lbl_sold_thuchien.Text = "          2. Số lao động thực hiện đến 31/12/" + txtNam.Text
                viewBC05(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), all, IdBC05_cur)
                labStatusProcess.Text = "Done"
                Cursor = Cursors.Default
            Catch ex As Exception
                labStatusProcess.Text = "Error"
                Cursor = Cursors.Default
                cmdPrint.Enabled = False
                MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            End Try
        End If

    End Sub

    Private Sub cmdCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreate.Click
        Dim vAll As Short
        cmdPrint.Enabled = False
        If CInt(txtNam.Text) < 2000 Then
            MessageBox.Show("Hãy nhập năm báo cáo", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        Else
            Try
                Cursor = Cursors.WaitCursor
                labStatusProcess.Text = "Waiting ....."
                If cbAll.Checked Then
                    vAll = 1
                Else
                    vAll = 0
                End If
                'labNam_TB.Text = txtNam.Text
                'labNam_TH.Text = txtNam.Text
                lbl_sold_thongbao.Text = "          1. Số lao động được thông báo đến 31/12/" + txtNam.Text
                lbl_sold_thuchien.Text = "          2. Số lao động thực hiện đến 31/12/" + txtNam.Text
                createBC05(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), vAll)
                labStatusProcess.Text = "Done"
                Cursor = Cursors.Default
                cmdSave.Enabled = True
            Catch ex As Exception
                labStatusProcess.Text = "Error"
                Cursor = Cursors.Default
                cmdSave.Enabled = False
                MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            End Try
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim vAll As Short
        If cbAll.Checked Then
            vAll = 1
        Else
            vAll = 0
        End If
        updateBC05(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), vAll, IdBC05_cur)
        If IdBC05_cur = "" Then
            cmdPrint.Enabled = False
        Else
            cmdPrint.Enabled = True
        End If
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Dim frm As frmLaoDongThuNhapReport = New frmLaoDongThuNhapReport
        Dim m_bc05 As tbBC05 = New tbBC05
        Dim vAll As Short
        If cbAll.Checked Then
            vAll = 1
        Else
            vAll = 0
        End If
        frm.Id_BC05 = m_bc05.getID(CInt(cboDonVi.SelectedValue), CInt(txtNam.Text), vAll)
        frm.IdDonVi_BC05 = CInt(cboDonVi.SelectedValue)
        frm.Nam_BC05 = CInt(txtNam.Text)
        frm.All_BC05 = vAll
        frm.Show()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub frmBC05_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmBC05_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cboDonVi.DataSource = listDonvi(True)
        txtNam.Text = Now.Year
        'labNam_TB.Text = txtNam.Text
        'labNam_TH.Text = txtNam.Text
        lbl_sold_thongbao.Text = "          1. Số lao động được thông báo đến 31/12/" + txtNam.Text
        lbl_sold_thuchien.Text = "          2. Số lao động thực hiện đến 31/12/" + txtNam.Text
        cboDonVi.Focus()
        cmdPrint.Enabled = False
        cmdSave.Enabled = False
        IdBC05_cur = ""
    End Sub

    Private Sub txtNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNam.TextChanged
        txtNam.Text = Val(txtNam.Text.Trim)
        lbl_sold_thongbao.Text = "          1. Số lao động được thông báo đến 31/12/" + txtNam.Text
        lbl_sold_thuchien.Text = "          2. Số lao động thực hiện đến 31/12/" + txtNam.Text
    End Sub


    Private Sub viewBC05(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short, ByRef vIdBC05_cur As String)

        Dim dt As DataTable
        Dim db As DBAccess = New DBAccess
        Dim soLD As Integer = db.getNumber("SELECT (TH_LDDH+TH_LDNH) FROM tbBC05 WHERE IdDonVi=" & vIdDonVi & " And [All]=" & vAll & " And Nam=" & vYear & " And DBtmp=" & flagReport)
        If soLD > 0 Then
            dt = db.SelectDBRows("SELECT id, TB_LDDH as 'c1', TB_LDNH as 'c2', TH_LDDH as 'c3', TH_LDNH as 'c4', L_ChucVu as 'c5', L_CapBac as 'c6', L_Phucap as 'c7', L_TronGoi as 'c8', TN_BHXH as 'c9', TN_ThemGio as 'c10', TN_boSung as 'c11', (L_ChucVu+L_CapBac+L_Phucap+L_TronGoi+TN_BHXH+TN_ThemGio+TN_boSung) as 'c12', ((L_ChucVu+L_CapBac+L_Phucap+L_TronGoi+TN_BHXH+TN_ThemGio+TN_boSung)/(" & soLD & "*12)) FROM tbBC05 WHERE IdDonVi=" & vIdDonVi & " And [All]=" & vAll & " And Nam=" & vYear)
            vIdBC05_cur = dt.Rows(0).Item("id")
            txtTB_LDDH.Text = dt.Rows(0).Item("c1")
            txtTB_LDNH.Text = dt.Rows(0).Item("c2")
            txtTH_LDDH.Text = dt.Rows(0).Item("c3")
            txtTH_LDNH.Text = dt.Rows(0).Item("c4")
            txtL_Chucvu.Text = dt.Rows(0).Item("c5")
            txtL_Capbac.Text = dt.Rows(0).Item("c6")
            txtL_Phucap.Text = dt.Rows(0).Item("c7")
            txtL_Trongoi.Text = dt.Rows(0).Item("c8")
            txtL_BHXH.Text = dt.Rows(0).Item("c9")
            txtL_Themgio.Text = dt.Rows(0).Item("c10")
            txtL_Bosung.Text = dt.Rows(0).Item("c11")
            Dim TongTN As Long = dt.Rows(0).Item("c5") + dt.Rows(0).Item("c6") + dt.Rows(0).Item("c7") + dt.Rows(0).Item("c8") + dt.Rows(0).Item("c9") + dt.Rows(0).Item("c10") + dt.Rows(0).Item("c11")
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
            txtL_Chucvu.ReadOnly = False
            txtL_Capbac.ReadOnly = False
            txtL_Phucap.ReadOnly = False
            txtL_Trongoi.ReadOnly = False
            txtL_BHXH.ReadOnly = False
            txtL_Themgio.ReadOnly = False
            txtL_Bosung.ReadOnly = False
            cmdSave.Enabled = True
            cmdPrint.Enabled = True
        Else
            vIdBC05_cur = ""
            txtTB_LDDH.Text = ""
            txtTB_LDNH.Text = ""
            txtTH_LDDH.Text = ""
            txtTH_LDNH.Text = ""
            txtL_Chucvu.Text = ""
            txtL_Capbac.Text = ""
            txtL_Phucap.Text = ""
            txtL_Trongoi.Text = ""
            txtL_BHXH.Text = ""
            txtL_Themgio.Text = ""
            txtL_Bosung.Text = ""
            txtTongThuNhap.Text = ""
            txtThuNhapBQ.Text = ""
            txtL_Chucvu.ReadOnly = True
            txtL_Capbac.ReadOnly = True
            txtL_Phucap.ReadOnly = True
            txtL_Trongoi.ReadOnly = True
            txtL_BHXH.ReadOnly = True
            txtL_Themgio.ReadOnly = True
            txtL_Bosung.ReadOnly = True
            cmdSave.Enabled = False
            cmdPrint.Enabled = False
            MessageBox.Show("Báo cáo 05 năm " & vYear & " chưa được lưu trữ. Hãy tạo lại báo cáo!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If

    End Sub

    Private Sub getSoLaoDong(ByVal vIdDonVi As Integer, ByVal vNam As Integer, ByVal vAll As Short, ByRef vSoDaiHan As Integer, ByRef vSoNganHan As Integer)
        Try
            Dim strSql As String
            Dim dt As DataTable
            Dim db As DBAccess = New DBAccess
            If vAll = 1 Then
                strSql = " SELECT sum(daihan) as 'DaiH', sum(nganhan) as 'NganH' FROM KeHoachLD WHERE (IdDonVi = " & vIdDonVi & "  or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ")) And Nam = " & vNam & " And Thang in (SELECT distinct(max(Thang)) FROM KeHoachLD WHERE (IdDonVi = " & vIdDonVi & "  or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ")) And Nam = " & vNam & ") group by left(IdKHLD,4)"
                dt = db.SelectDBRows(strSql)
            Else
                strSql = " SELECT sum(daihan) as 'DaiH', sum(nganhan) as 'NganH' FROM KeHoachLD WHERE IdDonVi = " & vIdDonVi & " And Nam = " & vNam & " And Thang in (SELECT distinct(max(Thang)) FROM KeHoachLD WHERE (IdDonVi = " & vIdDonVi & "  or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ")) And Nam = " & vNam & ") group by left(IdKHLD,4)"
                'strSql = " SELECT daihan as 'DaiH', nganhan as 'NganH' FROM KeHoachLD WHERE IdDonVi = " & vIdDonVi & " And Nam = " & vNam & "  And Thang<>0 And Thang<=12 group by IdDonVi"
                dt = db.SelectDBRows(strSql)
            End If
            'If vAll = 1 Then
            '    strSql = " SELECT sum(daihan) as 'DaiH', sum(nganhan) as 'NganH' FROM KeHoachLD WHERE (IdDonVi = " & vIdDonVi & "  or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ")) And Nam = " & vNam & " And Thang<>0 And Thang<=12 group by left(IdKHLD,4)"
            '    dt = db.SelectDBRows(strSql)
            '    If dt.Rows.Count = 0 Then
            '        strSql = " SELECT sum(daihan) as 'DaiH', sum(nganhan) as 'NganH' FROM KeHoachLD WHERE (IdDonVi = " & vIdDonVi & "  or IdDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ")) And Nam = " & vNam & " And Thang=0 group by left(IdKHLD,4)"
            '        dt = db.SelectDBRows(strSql)
            '    End If
            'Else
            '    strSql = " SELECT sum(daihan) as 'DaiH', sum(nganhan) as 'NganH' FROM KeHoachLD WHERE IdDonVi = " & vIdDonVi & " And Nam = " & vNam & "  And Thang<>0 And Thang<=12 group by IdDonVi"
            '    dt = db.SelectDBRows(strSql)
            '    If dt.Rows.Count = 0 Then
            '        strSql = " SELECT sum(daihan) as 'DaiH', sum(nganhan) as 'NganH' FROM KeHoachLD WHERE IdDonVi = " & vIdDonVi & " And Nam = " & vNam & " And Thang=0 group by IdDonVi"
            '        dt = db.SelectDBRows(strSql)
            '    End If
            'End If
            'If vAll = 1 Then
            '    strSql = "SELECT sum(DaiHan) as DaiHan, sum(NganHan) as NganHan FROM KeHoachLD WHERE (idDonVi=" & vIdDonVi & " or idDonVi in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & ") or idDonVi in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & vIdDonVi & "))) AND nam=" & vNam & " AND IdPhong=0  GROUP by IdDonvi, Nam"
            'Else
            '    strSql = "SELECT TOP 1 DaiHan, NganHan FROM KeHoachLD WHERE idDonVi=" & vIdDonVi & " AND nam=" & vNam & " AND IdPhong=0 order by Thang desc"
            'End If
            If dt.Rows.Count > 0 Then
                vSoDaiHan = dt.Rows(0).Item("DaiH")
                vSoNganHan = dt.Rows(0).Item("NganH")
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Sub createBC05(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short)
        Dim vSoKeHoachDaiHan As Integer
        Dim vSoKeHoachNganHan As Integer
        Dim dt As DataTable
        Dim db As DBAccess = New DBAccess

        getSoLaoDong(vIdDonVi, vYear, vAll, vSoKeHoachDaiHan, vSoKeHoachNganHan)
        dt = getBC(vIdDonVi, vYear, vAll, vSoKeHoachDaiHan, vSoKeHoachNganHan)
        If (dt.Rows(0).Item("c3") + dt.Rows(0).Item("c4")) > 0 Then
            txtTB_LDDH.Text = dt.Rows(0).Item("c1")
            txtTB_LDNH.Text = dt.Rows(0).Item("c2")
            txtTH_LDDH.Text = dt.Rows(0).Item("c3")
            txtTH_LDNH.Text = dt.Rows(0).Item("c4")
            txtL_Chucvu.Text = dt.Rows(0).Item("c5")
            txtL_Capbac.Text = dt.Rows(0).Item("c6")
            txtL_Phucap.Text = dt.Rows(0).Item("c7")
            txtL_Trongoi.Text = dt.Rows(0).Item("c8")
            txtL_BHXH.Text = dt.Rows(0).Item("c9")
            txtL_Themgio.Text = dt.Rows(0).Item("c10")
            txtL_Bosung.Text = dt.Rows(0).Item("c11")
            'Dim TongTN As Long = dt.Rows(0).Item("c5") + dt.Rows(0).Item("c6") + dt.Rows(0).Item("c7") + dt.Rows(0).Item("c8") + dt.Rows(0).Item("c9") + dt.Rows(0).Item("c10") + dt.Rows(0).Item("c11")
            txtTongThuNhap.Text = dt.Rows(0).Item("c12")  'TongTN
            txtThuNhapBQ.Text = dt.Rows(0).Item("c13")   'TongTN / ((dt.Rows(0).Item("c3") + dt.Rows(0).Item("c4")) * 12)
            txtL_Chucvu.ReadOnly = False
            txtL_Capbac.ReadOnly = False
            txtL_Phucap.ReadOnly = False
            txtL_Trongoi.ReadOnly = False
            txtL_BHXH.ReadOnly = False
            txtL_Themgio.ReadOnly = False
            txtL_Bosung.ReadOnly = False
        Else
            txtTB_LDDH.Text = ""
            txtTB_LDNH.Text = ""
            txtTH_LDDH.Text = ""
            txtTH_LDNH.Text = ""
            txtL_Chucvu.Text = ""
            txtL_Capbac.Text = ""
            txtL_Phucap.Text = ""
            txtL_Trongoi.Text = ""
            txtL_BHXH.Text = ""
            txtL_Themgio.Text = ""
            txtL_Bosung.Text = ""
            txtTongThuNhap.Text = ""
            txtThuNhapBQ.Text = ""
            txtL_Chucvu.ReadOnly = True
            txtL_Capbac.ReadOnly = True
            txtL_Phucap.ReadOnly = True
            txtL_Trongoi.ReadOnly = True
            txtL_BHXH.ReadOnly = True
            txtL_Themgio.ReadOnly = True
            txtL_Bosung.ReadOnly = True
            MessageBox.Show("Báo cáo 05 chưa được lưu trữ. Hãy tạo lại báo cáo!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Function updateBC05(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short, ByRef vIdBC05_cur As String) As Boolean

        If txtL_Chucvu.Text.Trim = "" Then txtL_Chucvu.Text = "0"
        If txtL_Capbac.Text.Trim = "" Then txtL_Capbac.Text = "0"
        If txtL_Phucap.Text.Trim = "" Then txtL_Phucap.Text = "0"
        If txtL_Trongoi.Text.Trim = "" Then txtL_Trongoi.Text = "0"
        If txtL_BHXH.Text.Trim = "" Then txtL_BHXH.Text = "0"
        If txtL_Themgio.Text.Trim = "" Then txtL_Themgio.Text = "0"
        If txtL_Bosung.Text.Trim = "" Then txtL_Bosung.Text = "0"

        If (CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)) > 0 Then

            Dim m_BC05 As tbBC05 = New tbBC05
            Dim IdBC05 As String = m_BC05.getID(vIdDonVi, vYear, vAll)
            m_BC05.Id = IdBC05
            m_BC05.IdDonVi = vIdDonVi
            m_BC05.All = vAll
            m_BC05.Nam = vYear
            m_BC05.TB_LDNH = CInt(txtTB_LDNH.Text.Trim)
            m_BC05.TB_LDDH = CInt(txtTB_LDDH.Text.Trim)
            m_BC05.TH_LDNH = CInt(txtTH_LDNH.Text.Trim)
            m_BC05.TH_LDDH = CInt(txtTH_LDDH.Text.Trim)
            m_BC05.L_ChucVu = N2Number(MoneyValue(txtL_Chucvu.Text.ToString))
            m_BC05.L_CapBac = N2Number(MoneyValue(txtL_Capbac.Text.ToString))
            m_BC05.L_PhuCap = N2Number(MoneyValue(txtL_Phucap.Text.ToString))
            m_BC05.L_TronGoi = N2Number(MoneyValue(txtL_Trongoi.Text.ToString))
            m_BC05.TN_BHXH = N2Number(MoneyValue(txtL_BHXH.Text.ToString))
            m_BC05.TN_ThemGio = N2Number(MoneyValue(txtL_Themgio.Text.ToString))
            m_BC05.TN_BoSung = N2Number(MoneyValue(txtL_Bosung.Text.ToString))
            m_BC05.DBtmp = flagReport
            If m_BC05.Id = "" Then
                vIdBC05_cur = m_BC05.Add()
            Else
                m_BC05.Update()
                vIdBC05_cur = m_BC05.Id
            End If
        End If

    End Function

    Private Function getBC(ByVal vIdDonVi As Integer, ByVal vYear As Integer, ByVal vAll As Short, ByVal vTbDaiHan As Integer, ByVal vTbNganHan As Integer) As DataTable
        Dim dbconn As DBAccess = New DBAccess
        Dim conn As SqlConnection = dbconn.getConnection
        Dim cmd As SqlCommand
        If flagReport = 0 Then
            cmd = New SqlCommand("BC05")
        Else
            cmd = New SqlCommand("BC05_NN")
        End If
        Dim ds As New DataSet
        Dim da As SqlDataAdapter

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@IdDonVi", vIdDonVi))
        cmd.Parameters.Add(New SqlParameter("@Nam", vYear))
        cmd.Parameters.Add(New SqlParameter("@All", vAll))
        cmd.Parameters.Add(New SqlParameter("@ldtb_daihan", vTbDaiHan))
        cmd.Parameters.Add(New SqlParameter("@ldtb_nganhan", vTbNganHan))
        cmd.Connection = conn
        Try
            conn.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(ds, "tbBC")
            Return ds.Tables("tbBC")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            dbconn.closeConnection(conn)
        End Try
    End Function

    Private Sub txtL_Chucvu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_Chucvu_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_Chucvu.TextChanged
        Try
            txtL_Chucvu = formatMoneyinTextbox(txtL_Chucvu)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtL_Capbac_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_Capbac_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_Capbac.TextChanged
        Try
            txtL_Capbac = formatMoneyinTextbox(txtL_Capbac)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtL_Phucap_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_Phucap_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_Phucap.TextChanged
        Try
            txtL_Phucap = formatMoneyinTextbox(txtL_Phucap)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtL_Trongoi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_Trongoi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_Trongoi.TextChanged
        Try
            txtL_Trongoi = formatMoneyinTextbox(txtL_Trongoi)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtL_BHXH_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_BHXH_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_BHXH.TextChanged
        Try
            txtL_BHXH = formatMoneyinTextbox(txtL_BHXH)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtL_Themgio_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_Themgio_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_Themgio.TextChanged
        Try
            txtL_Themgio = formatMoneyinTextbox(txtL_Themgio)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtL_Bosung_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim soLD As Integer = CInt(txtTH_LDDH.Text.Trim) + CInt(txtTH_LDNH.Text.Trim)
        If soLD > 0 Then
            Dim TongTN As Long = CLng(MoneyValue(txtL_Chucvu.Text.ToString)) + CLng(MoneyValue(txtL_Capbac.Text.ToString)) + CLng(MoneyValue(txtL_Phucap.Text.ToString)) + CLng(MoneyValue(txtL_Trongoi.Text.ToString)) + CLng(MoneyValue(txtL_BHXH.Text.ToString)) + CLng(MoneyValue(txtL_Themgio.Text.ToString)) + CLng(MoneyValue(txtL_Bosung.Text.ToString))
            txtTongThuNhap.Text = TongTN * 1000
            txtThuNhapBQ.Text = TongTN * 1000 / (soLD * 12)
        Else
            txtTongThuNhap.Text = 0
            txtThuNhapBQ.Text = 0
        End If
    End Sub

    Private Sub txtL_Bosung_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtL_Bosung.TextChanged
        Try
            txtL_Bosung = formatMoneyinTextbox(txtL_Bosung)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTongThuNhap_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTongThuNhap.TextChanged
        Try
            txtTongThuNhap = formatMoneyinTextbox(txtTongThuNhap)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtThuNhapBQ_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThuNhapBQ.TextChanged
        Try
            txtThuNhapBQ = formatMoneyinTextbox(txtThuNhapBQ)
        Catch ex As Exception
        End Try
    End Sub

End Class