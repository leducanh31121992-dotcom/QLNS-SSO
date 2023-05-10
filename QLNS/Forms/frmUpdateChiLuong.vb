Public Class frmUpdateChiLuong
    Inherits System.Windows.Forms.Form
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler
    Private _IdChiLuong As String
    Private _IdChi_ChiTiet As String
    Private _IdDonVi As String
    Private _IDCB As String
    Private _NH_TS As Int16
    Private _IDLoaiChi As Integer
    Private vKy As Int16 = 0
    Private vIdCB As String
    Private vTamUng As Long

    Public Property IdChiLuong() As String
        Get
            Return _IdChiLuong
        End Get
        Set(ByVal value As String)
            _IdChiLuong = value
        End Set
    End Property

    Public Property IdChi_ChiTiet() As String
        Get
            Return _IdChi_ChiTiet
        End Get
        Set(ByVal value As String)
            _IdChi_ChiTiet = value
        End Set
    End Property

    Public Property IdDonVi() As String
        Get
            Return _IdDonVi
        End Get
        Set(ByVal value As String)
            _IdDonVi = value
        End Set
    End Property

    Public Property IDLoaiChi() As Integer
        Get
            Return _IDLoaiChi
        End Get
        Set(ByVal value As Integer)
            _IDLoaiChi = value
        End Set
    End Property

    Public Property NH_TS() As Int16
        Get
            Return _NH_TS
        End Get
        Set(ByVal value As Int16)
            _NH_TS = value
        End Set
    End Property

    Private Sub frmUpdateChiLuong_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub

    Private Sub frmUpdateChiLuong_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If IDLoaiChi = 1 Then
                Dim m_ChiLuongCT As New HS_ChiLuong_CT
                pnlThemGio.Visible = False
                pnlBosung.Visible = False
                getData_Luong(IdChi_ChiTiet, m_ChiLuongCT)
                If NH_TS = 0 Then
                    If vKy = 0 Then
                        pnlKy1.Visible = True
                        pnlKy2.Visible = True
                    Else
                        If vKy = 1 Then
                            pnlKy1.Visible = True
                            pnlKy2.Visible = False
                        Else
                            pnlKy1.Visible = False
                            pnlKy2.Visible = True
                        End If
                    End If
                Else
                    If NH_TS = 1 Then
                        pnlKy1.Visible = True
                        pnlKy2.Visible = True
                    Else
                        pnlKy1.Visible = False
                        pnlKy2.Visible = True
                        labNghi.Text = "Tổng ngày hưởng lương"
                    End If
                End If
                fillData_Luong(IdChi_ChiTiet, m_ChiLuongCT)
            Else
                If IDLoaiChi = 2 Then
                    pnlKy.Visible = False
                    pnlKy1.Visible = False
                    pnlKy2.Visible = False
                    pnlBosung.Visible = False
                    Dim m_ChiThemGio_CT As New HS_ChiThemGio_CT
                    m_ChiThemGio_CT = m_ChiThemGio_CT.getRecord(IdChi_ChiTiet)
                    vIdCB = m_ChiThemGio_CT.IdCanBo
                    fillData_ThemGio(m_ChiThemGio_CT)
                Else
                    pnlKy.Visible = False
                    pnlKy1.Visible = False
                    pnlKy2.Visible = False
                    pnlThemGio.Visible = False
                    Dim m_ChiBoSung_CT As New HS_ChiBosung_CT
                    m_ChiBoSung_CT = m_ChiBoSung_CT.getRecord(IdChi_ChiTiet)
                    vIdCB = m_ChiBoSung_CT.IdCanBo
                    fillData_BoSung(m_ChiBoSung_CT)
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub getData_Luong(ByVal vIdHS_ChiLuong_CT As String, ByRef objL As HS_ChiLuong_CT)
        objL = objL.getRecord(IdChi_ChiTiet)
        vKy = objL.Ky
        vIdCB = objL.IdCanBo
        vTamUng = objL.T_TamUng
    End Sub

    Private Sub fillData_Luong(ByVal vIdHS_ChiLuong_CT As String, ByVal objL As HS_ChiLuong_CT)
        txtCanBo.Text = getFullName(objL.IdCanBo, IIf(objL.NH_TS = 0, False, True))
        txtTong100.Text = objL.Tong100.ToString
        txtTongChi.Text = objL.TongChi.ToString
        txtTNCN.Text = objL.T_TNCNtmp.ToString
        txtDFCD.Text = objL.T_DFCD.ToString
        txtBHXH.Text = objL.T_BHXH.ToString
        txtBHYT.Text = objL.T_BHYT.ToString
        txtBHTN.Text = objL.T_BHTN.ToString
        'txtTruLPC.Text = objL.T_Nghi.ToString
        txtTruyThu.Text = objL.T_TruyThu.ToString
        txtTruyLinh.Text = objL.TruyLinh.ToString
    End Sub

    Private Sub fillData_ThemGio(ByVal objTG As HS_ChiThemGio_CT)
        txtCanBo.Text = getFullName(objTG.IdCanBo, False)
        txtGio150.Text = objTG.Gio150.ToString
        txtGio200.Text = objTG.Gio200.ToString
        txtGio300.Text = objTG.Gio300.ToString
        txtTraThem.Text = objTG.LuongLamDem.ToString
    End Sub

    Private Sub fillData_BoSung(ByVal objBS As HS_ChiBosung_CT)
        txtCanBo.Text = getFullName(objBS.IdCanBo, IIf(NH_TS = 1 Or NH_TS = 2, True, False))
        txtHesoChi.Text = objBS.Hesochi.ToString
        txtThuclinh.Text = objBS.TongChi.ToString
    End Sub

    Private Sub updateData_Luong()
        Dim m_ChiLuongCT As New HS_ChiLuong_CT
        m_ChiLuongCT.IdHS_ChiLuong_CT = IdChi_ChiTiet
        m_ChiLuongCT.Ky = vKy
        m_ChiLuongCT.IdCanBo = vIdCB
        m_ChiLuongCT.NH_TS = NH_TS
        m_ChiLuongCT.Tong100 = N2Number(MoneyValue(txtTong100.Text.ToString))
        m_ChiLuongCT.TongChi = N2Number(MoneyValue(txtTongChi.Text.ToString))
        If vKy = 0 Or vKy = 1 Then
            m_ChiLuongCT.T_TNCNtmp = N2Number(MoneyValue(txtTNCN.Text.ToString))
            m_ChiLuongCT.T_DFCD = N2Number(MoneyValue(txtDFCD.Text.ToString))
            m_ChiLuongCT.T_BHXH = N2Number(MoneyValue(txtBHXH.Text.ToString))
            m_ChiLuongCT.T_BHYT = N2Number(MoneyValue(txtBHYT.Text.ToString))
            m_ChiLuongCT.T_BHTN = N2Number(MoneyValue(txtBHTN.Text.ToString))
        End If
        If vKy = 0 Or vKy = 2 Then
            m_ChiLuongCT.T_TamUng = vTamUng
            'm_ChiLuongCT.T_Nghi = N2Number(MoneyValue(txtTruLPC.Text.ToString))
            m_ChiLuongCT.T_TruyThu = N2Number(MoneyValue(txtTruyThu.Text.ToString))
            m_ChiLuongCT.TruyLinh = N2Number(MoneyValue(txtTruyLinh.Text.ToString))
        End If
        m_ChiLuongCT.UpdatenotFull()
    End Sub

    Private Sub updateData_ThemGio()
        Dim m_ChiThemGioCT As New HS_ChiThemGio_CT
        m_ChiThemGioCT.IdHS_ChiThemGio_CT = IdChi_ChiTiet
        m_ChiThemGioCT.IdCanBo = vIdCB
        m_ChiThemGioCT.Gio150 = CDec(txtGio150.Text.ToString)
        m_ChiThemGioCT.Gio200 = CDec(txtGio200.Text.ToString)
        m_ChiThemGioCT.Gio300 = CDec(txtGio300.Text.ToString)
        m_ChiThemGioCT.LuongLamDem = N2Number(MoneyValue(txtTraThem.Text.ToString))
        m_ChiThemGioCT.UpdatenotFull()
    End Sub

    Private Sub updateData_BoSung()
        Dim m_ChiBoSungCT As New HS_ChiBosung_CT
        m_ChiBoSungCT.IdHS_ChiBoSung_CT = IdChi_ChiTiet
        m_ChiBoSungCT.IdCanBo = vIdCB
        m_ChiBoSungCT.Hesochi = CDbl(txtHesoChi.Text.ToString)
        m_ChiBoSungCT.TongChi = N2Number(MoneyValue(txtThuclinh.Text.ToString))
        m_ChiBoSungCT.UpdatenotFull()
    End Sub

    Private Function checkFrm()
        Dim returnValue As String = ""
        Return returnValue
    End Function

    Private Sub bntUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntUpdate.Click
        checkFrm()
        If IDLoaiChi = 1 Then
            updateData_Luong()
        Else
            If IDLoaiChi = 2 Then
                updateData_ThemGio()
            Else
                updateData_BoSung()
            End If
        End If
        Close()
    End Sub

    Private Sub bntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel.Click
        Close()
    End Sub

    Private Sub txtTong100_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTong100.TextChanged
        Try
            txtTong100 = formatMoneyinTextbox(txtTong100)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTongChi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTongChi.TextChanged
        Try
            txtTongChi = formatMoneyinTextbox(txtTongChi)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTNCN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTNCN.TextChanged
        Try
            txtTNCN = formatMoneyinTextbox(txtTNCN)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtDFCD_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDFCD.TextChanged
        Try
            txtDFCD = formatMoneyinTextbox(txtDFCD)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtBHXH_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBHXH.TextChanged
        Try
            txtBHXH = formatMoneyinTextbox(txtBHXH)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtBHYT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBHYT.TextChanged
        Try
            txtBHYT = formatMoneyinTextbox(txtBHYT)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtBHTN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBHTN.TextChanged
        Try
            txtBHTN = formatMoneyinTextbox(txtBHTN)
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub txtTruLPC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTruLPC.TextChanged
    '    Try
    '        txtTruLPC = formatMoneyinTextbox(txtTruLPC)
    '    Catch ex As Exception
    '    End Try
    'End Sub

    Private Sub txtTruyThu_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTruyThu.TextChanged
        Try
            txtTruyThu = formatMoneyinTextbox(txtTruyThu)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTruyLinh_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTruyLinh.TextChanged
        Try
            txtTruyLinh = formatMoneyinTextbox(txtTruyLinh)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtTraThem_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTraThem.TextChanged
        Try
            txtTraThem = formatMoneyinTextbox(txtTraThem)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtGio150_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGio150.TextChanged
        txtGio150.Text = formatDouble(txtGio150.Text.Trim)
    End Sub

    Private Sub txtGio200_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGio200.TextChanged
        txtGio200.Text = formatDouble(txtGio200.Text.Trim)
    End Sub

    Private Sub txtGio300_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGio300.TextChanged
        txtGio300.Text = formatDouble(txtGio300.Text.Trim)
    End Sub

    Private Sub txtHesoChi_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHesoChi.TextChanged
        If txtHesoChi.Text <> "" Then
            txtHesoChi.Text = formatDouble(txtHesoChi.Text.Trim)
            txtHesoChi.SelectionStart = txtHesoChi.Text.Length
        End If
    End Sub

    Private Sub txtThuclinh_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThuclinh.TextChanged
        Try
            txtThuclinh = formatMoneyinTextbox(txtThuclinh)
        Catch ex As Exception
        End Try
    End Sub
End Class