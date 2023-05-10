Public Class frmSuaLoi

    Private dbconn As DBAccess = New DBAccess
    Private IDQdNhanSu_Err As String = ""
    Private IDcanbo_Err As String = ""
    Private IdDonViCu As Integer = 0
    Private IdPhongCu As Integer = 0
    Private IdChucVuCu As Integer = 0
    Private IdChuyenMonCu As Integer = 0
    Private IdBangLuong As Integer = 0
    Private SoQDCU As String = ""
    Private NgayHLCU As Date
    Private NgayQDCU As Date
    Private init As Boolean = False
    Private init_L As Boolean = False
    Private init_P As Boolean = False
    Private idLuong As String = ""
    Private idPhucap As String = ""
    Private countDS_Error As Integer = 0

    Private Sub frmSuaLoi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        updateIdDonvi()
        If (Globals.Roles.IndexOf(";101;") < 0) Then
            MessageBox.Show("Bạn không có quyền sửa đổi dữ liệu quyết định nhân sự!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            gridDS_Errors_CellClick(sender, Nothing)
            Return
        End If
        bindDS_Error()
    End Sub

    Private Sub bindDS_Error()
        Dim sSQL As String = ""
        Dim dt As DataTable
        sSQL = "SELECT IdCanBo, MaCB, HoTen, N'' as IDQDNhanSu, N'Chưa có quyết định nhân sự nào của cán bộ. Hãy nhập QĐ nhân sự mới nhất của cán bộ về đơn vị hiện tại' as ThongBao " & _
                "        FROM HS_Canbo WHERE left(IdCanBo,4)='" & TEN_DV_VT & "' AND IdCanbo not in (SELECT distinct IdCanbo FROM QDNhanSu WHERE left(IdCanBo,4)='" & TEN_DV_VT & "')" & _
                " UNION" & _
                " SELECT t1.IDCanbo, MaCB, HoTen, N'' as IDQDNhanSu, N'Chưa có quyết định nhân sự nào của cán bộ về đơn vị hiện tại (QĐ của NHCSXH)' as ThongBao " & _
                "	FROM HS_Canbo t1, " & _
                "	(SELECT IDCanbo, count(*) as SoQDNHCS FROM QDNhansu WHERE left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP by Idcanbo) t2," & _
                "	(SELECT IDCanbo, count(*) as SoQDNHCS FROM QDNhansu WHERE IsQD_NHCS =0 AND left(IdCanBo,4)='" & TEN_DV_VT & "' GROUP by Idcanbo) t3" & _
                " 	WHERE t1.IdCanbo=t2.Idcanbo AND t2.IdCanbo=t3.Idcanbo AND t2.SoQDNHCS=t3.SoQDNHCS" & _
                " UNION" & _
                " SELECT t1.IdCanBo, MaCB, HoTen, IDQDNhanSu, N'Đơn vị hiện đang công tác của cán bộ không thuộc chi nhánh' as ThongBao " & _
                "        FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                "        WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                "           AND t1.idDonvi_Moi not in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and Active=1 " & _
                "           AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "' " & _
                " UNION" & _
                " SELECT t1.IdCanBo, MaCB, HoTen, IDQDNhanSu, N'Phòng hiện đang công tác của cán bộ không tồn tại trong chi nhánh' as ThongBao " & _
                "        FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                "        WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                "           AND t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") and Active=1 " & _
                "           AND t1.idPhong_Moi in (SELECT id FROM PhongBan WHERE id=t1.idPhong_Moi and status=0)" & _
                "           AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "' " & _
                " UNION" & _
                " SELECT t1.IdCanBo, MaCB, HoTen, IDQDNhanSu, N'Phòng/Ban trong quyết định nhân sự mới nhất của cán bộ hiện không tồn tại' as ThongBao " & _
                "        FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                "        WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                "           AND ((not (t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") AND Active=1) " & _
                "           AND t1.idPhong_Moi in (SELECT id FROM PhongBan WHERE status=0))  or t1.idDonvi_Moi=0 or t1.idPhong_Moi=0 ) AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "' " & _
                " UNION " & _
                " SELECT t1.IdCanBo, MaCB, HoTen, IDQDNhanSu, N'Không tích chọn lưu thông tin hồ sơ của cán bộ tại chi nhánh' as ThongBao " & _
                "        FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 Group by Idcanbo Having max(NgayHL)<=getdate()) t2, HS_Canbo t3 " & _
                "        WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                "           AND (t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") Or t1.idDonvi_Cu in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & "))" & _
                "	        AND Active=0 AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "'" & _
                " UNION " & _
                " SELECT t1.IdCanBo, MaCB, HoTen, IDQDNhanSu, N'Ngày hiệu lực của quyết định nhân sự mới nhất của cán bộ lớn hơn ngày hiện tại' as ThongBao " & _
                "        FROM QDNhanSu t1,(SELECT IdCanBo, max(NgayHL) as NgayHL  FROM QDNhanSu WHERE IsQD_NHCS =1 Group by Idcanbo ) t2, HS_Canbo t3 " & _
                "        WHERE(t1.IdCanBo = t2.IdCanBo And t1.NgayHL = t2.NgayHL And t2.IdCanBo = t3.IdCanBo) " & _
                "           AND t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id=" & IdDONVI & " or id_goc=" & IdDONVI & ") AND Active=1 " & _
                "	        AND t2.NgayHL>getdate() " & _
                "           AND IsQD_NHCS =1 AND left(t2.IdCanBo,4)='" & TEN_DV_VT & "'"
        dt = dbconn.SelectDBRows(sSQL)
        gridDS_Errors.DataSource = Nothing
        IDQdNhanSu_Err = ""
        If dt.Rows.Count > 0 Then
            If Not init Then
                init = True
                bindCboLoaiQD()
                cboCVNguoiQD.DataSource = listChucVuQuyenRaQD()
                bindCboDonVi()
                bindCboChuyenMon()
            End If
            countDS_Error = dt.Rows.Count
            gridDS_Errors.AutoGenerateColumns = False
            gridDS_Errors.DataSource = dt
            IDQdNhanSu_Err = dt.Rows(0)("IDQDNhanSu").ToString
        End If

    End Sub

    Private Sub updateIdDonvi()
        Try
            dbconn.executeSQL("UPDATE HS_Canbo Set IdDonVi=" & IdDONVI & " WHERE left(idCanbo,4)='" & TEN_DV_VT & "' AND IdDonVi is null")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub gridDS_Errors_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gridDS_Errors.CellClick
        Try
            IDQdNhanSu_Err = gridDS_Errors.CurrentRow.Cells("IDQDNhanSuE").Value.ToString
            IDcanbo_Err = gridDS_Errors.CurrentRow.Cells("IDCanBoE").Value.ToString
            'tabMain.SelectedTab.Name = "tabRepair"
            'tabFrm_SelectedIndexChanged(sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tabFrm_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabMain.SelectedIndexChanged
        Select Case tabMain.SelectedTab.Name
            Case "tabRepair"
                initTabRepair()
            Case Else
                initTabDS_Errors()
        End Select
    End Sub

    Private Sub initTabDS_Errors()
        bindDS_Error()
    End Sub

    Private Sub bindCboLoaiQD()
        cboLoaiQD.DataSource = listDanhmucQDNhanSu()
    End Sub

    Private Sub bindCboDonVi()
        cboDonviMoi.DataSource = listDonvi(False, False, False, False, True)
    End Sub

    Private Sub bindCboPhongMoi(ByVal vIdDonviMoi As Integer)
        Try
            If vIdDonviMoi > 0 Then
                cboPhongMoi.DataSource = listPhong(vIdDonviMoi, False, False, False)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bindCboChucVu(ByVal vIdDonviMoi As Integer)
        cboChucvuMoi.DataSource = listChucVu(vIdDonviMoi)
    End Sub

    Private Sub bindCboChuyenMon()
        cboChuyenmonMoi.DataSource = listDanhmuc(12, False, True)
    End Sub

    Private Sub initTabRepair()
        If countDS_Error > 0 Then
            fillQDNhansu(IDQdNhanSu_Err)
            bntAdd.Enabled = True
        Else
            bntAdd.Enabled = False
        End If
    End Sub

    Private Sub fillQDNhansu(ByRef vIdQDNhanSu As String)
        Dim QDNhansu_Repair As QDNhanSu = New QDNhanSu
        Dim QDnhansu_Final As QDNhanSu = New QDNhanSu
        Dim m_Luong As LuongCanBo = New LuongCanBo
        Dim m_Phucap As Phucap = New Phucap
        Dim vIdNghiDinhLuong, vIdBangluong, vIdNgachluong, vIdLoaiPC As Integer

        labCanbo.Text = " Mã cán bộ: " & getCanBo_Ma(IDcanbo_Err, False) & "        " & getFullName(IDcanbo_Err, False)

        If vIdQDNhanSu <> "" Then
            getIdBangNgachBac(QDNhansu_Repair.IdCanBo, IdBangLuong, 0, 0, 0)
            QDNhansu_Repair = QDNhansu_Repair.getRecord(vIdQDNhanSu)
            QDnhansu_Final = QDnhansu_Final.getFinalRecord(QDNhansu_Repair.IdCanBo, Nothing, True)
            If QDnhansu_Final.IdQDNhanSu <> QDNhansu_Repair.IdQDNhanSu Then
                QDNhansu_Repair = QDnhansu_Final
            End If
            If IDcanbo_Err <> QDnhansu_Final.IdCanBo Then IDcanbo_Err = QDNhansu_Repair.IdCanBo

            labAlert_QD.Text = ""

            txtSoQD.Text = QDNhansu_Repair.So_QD
            dpkNgayQD.Value = QDNhansu_Repair.NgayKy_QD
            dpkNgayHL.Value = QDNhansu_Repair.NgayHL
            If QDNhansu_Repair.NgayBoNhiem_TT = DateTime.MinValue Then
                dpkNgayBN.Checked = False
            Else
                dpkNgayBN.Checked = True
                dpkNgayBN.Value = QDNhansu_Repair.NgayBoNhiem_TT
            End If
            If QDNhansu_Repair.NgayThoiLuong = DateTime.MinValue Then
                dpkNgayTL.Checked = False
            Else
                dpkNgayTL.Checked = True
                dpkNgayTL.Value = QDNhansu_Repair.NgayThoiLuong
            End If
            txtNguoiQD.Text = QDNhansu_Repair.NguoiKy_QD
            txtDVraQD.Text = QDNhansu_Repair.DVraQD
            cboLoaiQD.SelectedValue = QDNhansu_Repair.IdLoaiQD
            If QDNhansu_Repair.IsQD_NHCS Then
                cbIsQD_NHCS.Checked = True
                cboCVNguoiQD.Visible = True
                grpTTCu.Visible = True
                grpTTCu.Enabled = False
                cboCVNguoiQD.SelectedValue = QDNhansu_Repair.idCV_Nguoiky_QD
                IdDonViCu = QDNhansu_Repair.IdDonvi_Cu
                IdPhongCu = QDNhansu_Repair.IdPhong_Cu
                IdChucVuCu = QDNhansu_Repair.IdChucvu_Cu
                IdChuyenMonCu = QDNhansu_Repair.IdChuyenMon_Cu
                txtDonViCu.Text = getDonvi(IdDonViCu)
                txtPhongCu.Text = getPhong(IdPhongCu)
                txtChucVuCu.Text = getDanhmuc_Name(14, IdChucVuCu)
                txtChuyenMonCu.Text = getDanhmuc_Name(12, IdChuyenMonCu)
                cboDonviMoi.SelectedValue = QDNhansu_Repair.IdDonvi_Moi
                cboPhongMoi.SelectedValue = QDNhansu_Repair.IdPhong_Moi
                cboChucvuMoi.SelectedValue = QDNhansu_Repair.IdChucvu_Moi
                cboChuyenmonMoi.SelectedValue = QDNhansu_Repair.IdChuyenMon_Moi

                Dim dt As DataTable = dbconn.SelectDBRows("SELECT * FROM ChiNhanh WHERE id in (SELECT Id_goc FROM Chinhanh WHERE id=" & QDNhansu_Repair.IdDonvi_Moi & ") AND status=0")
                If dt.Rows.Count > 0 Then
                    labDV.Text = dt.Rows(0)("ten_goi").ToString
                Else
                    labDV.Text = ""
                End If
            End If

            If QDNhansu_Repair.Active Then
                cbActive.Checked = True
            Else : cbActive.Checked = False
            End If
            txtGhichu.Text = QDNhansu_Repair.GhiChu

            NgayHLCU = QDNhansu_Repair.NgayHL
            SoQDCU = QDNhansu_Repair.So_QD
            NgayQDCU = QDNhansu_Repair.NgayKy_QD

            ' fill thông tin lương trong quyết định
            idLuong = m_Luong.getID(QDNhansu_Repair.IdCanBo, NgayHLCU, SoQDCU, NgayQDCU)
            If idLuong <> "" Then
                cbLuong.Checked = True
                m_Luong = m_Luong.getRecord(idLuong)
                txtHeso.Text = m_Luong.HeSoLuong
                getIdNDBangNgach(m_Luong.IdBacLuong, vIdNghiDinhLuong, vIdBangluong, vIdNgachluong)
                cboNghiDinh.SelectedValue = vIdNghiDinhLuong
                cboBang.SelectedValue = vIdBangluong
                cboNgach.SelectedValue = vIdNgachluong
                cboBac.SelectedValue = m_Luong.IdBacLuong
            Else
                cbLuong.Checked = False
                cboNghiDinh.SelectedValue = 0
                cboBang.SelectedValue = 0
                cboNgach.SelectedValue = 0
                cboBac.SelectedValue = 0
            End If

            ' fill thông tin phụ cấp liên quan trong quyết định
            idPhucap = m_Phucap.getID(QDNhansu_Repair.IdCanBo, NgayHLCU, SoQDCU, NgayQDCU)
            If idPhucap <> "" Then
                cbPhucap.Checked = True
                m_Phucap = m_Phucap.getRecord(idPhucap)
                cboMucPC.SelectedValue = m_Phucap.IdMucPC
                vIdLoaiPC = getIdLoaiPC(m_Phucap.IdMucPC)
                cboLoaiPC.SelectedValue = vIdLoaiPC
            Else
                cbPhucap.Checked = False
                cboMucPC.SelectedValue = 0
                cboLoaiPC.SelectedValue = 0
            End If
        Else
            blankQDNhansu()
        End If
    End Sub

    Private Function checkTabQDNhanSu() As String
        Dim strReturn As String = ""
        Try
            If cbIsQD_NHCS.Checked = False Then
                cbIsQD_NHCS.Checked = True
            End If
            If txtSoQD.Text = "" Then
                txtSoQD.Focus()
                strReturn = "Chưa nhập Số quyết định !"
                Exit Try
            Else
                txtSoQD.Text = txtSoQD.Text.Replace("_", "-")
            End If
            If dpkNgayHL.Value.Date > Now.Date Then
                strReturn = "Ngày hiệu lực phải nhỏ hơn hoặc bằng ngày hiện tại. Hãy nhập lại!"
                dpkNgayHL.Focus()
                Exit Try
            End If
            If IDQdNhanSu_Err = "" Then
                If checkQuyetDinh("QDNHANSU", IDcanbo_Err, txtSoQD.Text) Then
                    txtSoQD.Text = ""
                    txtSoQD.Focus()
                    strReturn = "Số quyết định nhân sự của cán bộ đã tồn tại. Hãy nhập lại!"
                    Exit Try
                End If
            End If
        Catch ex As Exception
            strReturn = ""
        End Try
        Return strReturn
    End Function

    Private Function updateQDNhanSu(ByRef vIdQDNhanSu As String) As Boolean
        Try
            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
            Dim dateHL As Date
            Dim dateQD As Date

            dateHL = DateTimeUtil.getDateCurrTime(dpkNgayHL.Text)
            dateQD = DateTimeUtil.getDate(dpkNgayQD.Text)
            m_QDNhanSu.IdQDNhanSu = vIdQDNhanSu
            m_QDNhanSu.IdCanBo = IDcanbo_Err
            m_QDNhanSu.So_QD = txtSoQD.Text.Trim
            m_QDNhanSu.NgayKy_QD = dateQD
            m_QDNhanSu.NguoiKy_QD = standardizeName(txtNguoiQD.Text)
            m_QDNhanSu.NgayHL = dateHL
            If dpkNgayBN.Checked Then
                m_QDNhanSu.NgayBoNhiem_TT = DateTimeUtil.getDate(dpkNgayBN.Text)
            Else
                m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
            End If
            If dpkNgayTL.Checked Then
                m_QDNhanSu.NgayThoiLuong = DateTimeUtil.getDate(dpkNgayTL.Text)
            Else
                m_QDNhanSu.NgayThoiLuong = DateTime.MinValue
            End If
            m_QDNhanSu.DVraQD = standardizeString(txtDVraQD.Text.Trim)
            m_QDNhanSu.IdLoaiQD = CInt(cboLoaiQD.SelectedValue)

            Dim m_QDNSfinal As QDNhanSu = New QDNhanSu
            m_QDNSfinal = m_QDNSfinal.getFinalRecord(IDcanbo_Err, dpkNgayHL.Value)
            m_QDNhanSu.IsQD_NHCS = 1
            m_QDNhanSu.IdDonvi_Cu = m_QDNSfinal.IdDonvi_Moi
            m_QDNhanSu.IdPhong_Cu = m_QDNSfinal.IdPhong_Moi
            m_QDNhanSu.IdChucvu_Cu = m_QDNSfinal.IdChucvu_Moi
            m_QDNhanSu.IdChuyenMon_Cu = m_QDNSfinal.IdChuyenMon_Moi
            m_QDNhanSu.IdDonvi_Moi = CInt(cboDonviMoi.SelectedValue)
            m_QDNhanSu.IdPhong_Moi = CInt(cboPhongMoi.SelectedValue)
            m_QDNhanSu.IdChucvu_Moi = CInt(cboChucvuMoi.SelectedValue)
            m_QDNhanSu.IdChuyenMon_Moi = CInt(cboChuyenmonMoi.SelectedValue)
            m_QDNhanSu.idCV_Nguoiky_QD = CInt(cboCVNguoiQD.SelectedValue)
            m_QDNhanSu.CV_NguoiKy_QD = ""
            m_QDNhanSu.LoaiQD = ""
            m_QDNhanSu.DenNgay = DateTime.MinValue
            m_QDNhanSu.NoiDung = ""

            Select Case getDanhmuc_MaSo(m_QDNhanSu.IdLoaiQD)
                Case "1518"
                    ' Kiem nhiem
                    m_QDNhanSu.IsKiemNhiem = 1
                Case "1519"
                    ' Thoi kiem nhiem
                    m_QDNhanSu.IsKiemNhiem = 2
                Case Else
                    m_QDNhanSu.IsKiemNhiem = 0
            End Select

            If cbActive.Checked Then
                m_QDNhanSu.Active = 1
            Else
                m_QDNhanSu.Active = 0
            End If
            m_QDNhanSu.GhiChu = standardizeString(txtGhichu.Text)

            If vIdQDNhanSu <> "" Then
                m_QDNhanSu.Update()
            Else
                vIdQDNhanSu = m_QDNhanSu.Add()
            End If
            If m_QDNhanSu.IdQDNhanSu <> "" Then
                If m_QDNhanSu.IsKiemNhiem = 2 Then
                    'cap nhat lai QD kiem nhiem truoc do
                    dbconn.executeSQL("UPDATE QDNhanSu SET IsKiemNhiem=2 WHERE IdCanBo='" & IDcanbo_Err & "' and IsKiemNhiem=1 and IdDonvi_Moi=" & m_QDNhanSu.IdDonvi_Moi & " and IdPhong_Moi=" & m_QDNhanSu.IdPhong_Moi & " and IdChucVu_Moi=" & m_QDNhanSu.IdChucvu_Moi)
                End If
                Dim m_Luong As LuongCanBo = New LuongCanBo
                Dim m_PhuCap As Phucap = New Phucap
                Dim maQDNhanSu As String = ""
                Dim IDLoaiQDLuong As String = ""
                Dim db As DBAccess = New DBAccess
                m_Luong.IdLuongCB = idLuong
                m_Luong.IdCanBo = IDcanbo_Err
                If cbLuong.Checked Then
                    m_Luong.Ngay_Huong = dateHL
                    m_Luong.NgayLen_DK = dateHL.AddMonths(getTimeNangBac(cboNgach.SelectedValue))
                    m_Luong.DVraQD = m_QDNhanSu.DVraQD
                    m_Luong.IsQD_NHCS = 1
                    m_Luong.NoiDung = ""
                    m_Luong.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                    m_Luong.CV_NguoiKy_QD = ""
                    m_Luong.IdBacLuong = CInt(cboBac.SelectedValue)
                    m_Luong.HeSoLuong = CDbl(txtHeso.Text)
                    m_Luong.LoaiQD = ""
                    maQDNhanSu = getDanhmuc_MaSo(m_QDNhanSu.IdLoaiQD)
                    IDLoaiQDLuong = db.getNumber("SELECT ID FROM DanhMuc WHERE ma_so ='43" & (CInt(maQDNhanSu.Substring(2, 1)) + 1) & maQDNhanSu.Substring(3, 1) & "' and ma_so in ('4311','4312','4313','4314','4315','4317','4318','4319','4327','4328','4329','4330','4331','4332','4323','4324','4325')")
                    If IDLoaiQDLuong = 0 Then
                        IDLoaiQDLuong = getDanhmuc_ID("4323")
                    End If
                    m_Luong.IdLoaiQD = IDLoaiQDLuong
                    m_Luong.SoQD = m_QDNhanSu.So_QD
                    m_Luong.NgayQD = dateQD
                    m_Luong.NguoiQD = m_QDNhanSu.NguoiKy_QD
                    m_Luong.GhiChu = m_QDNhanSu.GhiChu
                    If idLuong <> "" Then
                        m_Luong.Update()
                    Else
                        ' Kiem tra thông tin lương trong QDNS neu trung với thông tin lương hiện đang hưởng của CB thì ko cho nhập sang HS Luong
                        Dim m_LuongFinal As LuongCanBo = New LuongCanBo
                        m_LuongFinal = m_LuongFinal.getFinalRecord(IDcanbo_Err)
                        If Not (m_LuongFinal.IdBacLuong = m_Luong.IdBacLuong And m_Luong.IsQD_NHCS) Then
                            m_Luong.Add()
                        End If
                    End If
                Else
                    If idLuong <> "" Then
                        m_Luong.Delete()
                    End If
                End If
                m_PhuCap.IdCB_PhuCap = idPhucap
                m_PhuCap.IdCanBo = IDcanbo_Err
                If cbPhucap.Checked Then
                    m_PhuCap.IsQD_NHCS = 1
                    m_PhuCap.NoiDung = ""
                    m_PhuCap.IdMucPC = CInt(cboMucPC.SelectedValue)
                    m_PhuCap.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                    m_PhuCap.CV_NguoiKy_QD = ""
                    m_PhuCap.TuNgay = DateTimeUtil.getDate(dpkNgayHL.Text)  'dateHL
                    m_PhuCap.DenNgay = DateTime.MinValue
                    m_PhuCap.SoQD = m_QDNhanSu.So_QD
                    m_PhuCap.NgayQD = dateQD
                    m_PhuCap.NguoiQD = m_QDNhanSu.NguoiKy_QD
                    m_PhuCap.DVraQD = m_QDNhanSu.DVraQD
                    m_PhuCap.GhiChu = m_QDNhanSu.GhiChu
                    If idPhucap <> "" Then
                        Dim dtPC As DataTable
                        Dim iPC As Integer = 0
                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap FROM HS_PhucapCB WHERE IdCanbo='" & IDcanbo_Err & "' and Datediff(day,TuNgay,'" & NgayHLCU & "')=0 and Upper(SoQD)=Upper(N'" & SoQDCU & "') and NgayQD='" & NgayQDCU & "' and IdMucPC<>" & m_PhuCap.IdMucPC)
                        m_PhuCap.Update()
                        ' Update tiếp với các Phụ cấp khác có cùng số QĐ và ngày
                        If dtPC.Rows.Count > 0 Then
                            For iPC = 0 To dtPC.Rows.Count - 1
                                m_PhuCap.IdCB_PhuCap = dtPC.Rows(iPC).Item("IdCB_PhuCap")
                                m_PhuCap.UpdateNotFull()
                            Next
                        End If
                    Else
                        m_PhuCap.Add()
                    End If
                Else
                    If idPhucap <> "" Then
                        m_PhuCap.Delete()
                    End If
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub blankQDNhansu()
        
        txtDonViCu.Text = ""
        txtPhongCu.Text = ""
        txtChucVuCu.Text = ""
        txtChuyenMonCu.Text = ""

        grpTTCu.Enabled = False
        idLuong = ""
        idPhucap = ""
        txtSoQD.Text = ""
        dpkNgayQD.Value = Date.Now
        dpkNgayHL.Value = Date.Now
        txtDVraQD.Text = getDonvi(IdDONVI)
        txtNguoiQD.Text = GIAMDOC
        cboDonviMoi.SelectedValue = IdDONVI
        If CAP = 1 Then
            cboCVNguoiQD.SelectedValue = getDanhmuc_ID("1402")
        Else
            cboCVNguoiQD.SelectedValue = getDanhmuc_ID("1410")
        End If
        cbLuong.Checked = False
        cbPhucap.Checked = False
        dpkNgayTL.Checked = False
        dpkNgayBN.Checked = False
        cbActive.Checked = True
        cboNghiDinh.SelectedValue = 0
        cboBang.SelectedValue = 0
        cboNgach.SelectedValue = 0
        cboBac.SelectedValue = 0
        txtGhichu.Text = ""
        labAlert_QD.Text = ""
        NgayHLCU = Nothing
        SoQDCU = ""
        NgayQDCU = Nothing
        cbIsQD_NHCS.Focus()
    End Sub

    Private Sub InfLuong(ByVal active As Boolean)
        cboNghiDinh.Enabled = active
        cboBang.Enabled = active
        cboNgach.Enabled = active
        cboBac.Enabled = active
        txtHeso.Enabled = active
        If active And (Not init_L) Then
            init_L = True
            cboNghiDinh.DataSource = listNghiDinhLuong(True)
        End If
    End Sub

    Private Sub InfPhucap(ByVal active As Boolean)
        cboLoaiPC.Enabled = active
        cboMucPC.Enabled = active
        If active And (Not init_P) Then
            init_P = True
            cboLoaiPC.DataSource = listDanhmuc(31, True)
        End If
    End Sub

    Private Sub cboDonviMoi_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDonviMoi.SelectedValueChanged
        If Not (cboDonviMoi.SelectedValue Is Nothing) Then
            bindCboPhongMoi(CInt(cboDonviMoi.SelectedValue))
            bindCboChucVu(CInt(cboDonviMoi.SelectedValue))
            Dim dt As DataTable = dbconn.SelectDBRows("SELECT * FROM ChiNhanh WHERE id in (SELECT Id_goc FROM Chinhanh WHERE id=" & CInt(cboDonviMoi.SelectedValue) & ") AND status=0")
            If dt.Rows.Count > 0 Then
                labDV.Text = dt.Rows(0)("ten_goi").ToString
            Else
                labDV.Text = ""
            End If
        End If
    End Sub

    Private Sub cboNgach_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboNgach.SelectedValueChanged
        If Not (cboNgach.SelectedValue Is Nothing) Then
            If CInt(cboNgach.SelectedValue) > 0 Then cboBac.DataSource = listBacLuong(CInt(cboNgach.SelectedValue), 0, True)
        End If
    End Sub

    Private Sub cboNghiDinh_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboNghiDinh.SelectedValueChanged
        If Not (cboNghiDinh.SelectedValue Is Nothing) Then
            If CInt(cboNghiDinh.SelectedValue) > 0 Then cboBang.DataSource = listBangLuong(CInt(cboNghiDinh.SelectedValue), True)
        End If
    End Sub

    Private Sub cboBang_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBang.SelectedValueChanged
        If Not (cboBang.SelectedValue Is Nothing) Then
            If CInt(cboBang.SelectedValue) > 0 Then cboNgach.DataSource = listNgachLuong(CInt(cboBang.SelectedValue), True)
        End If
    End Sub

    Private Sub cboBac_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBac.SelectedValueChanged
        If Not (cboBac.SelectedValue Is Nothing) Then
            If CInt(cboBac.SelectedValue) > 0 Then
                Try
                    txtHeso.Text = getHeso(CInt(cboBac.SelectedValue))
                Catch ex As Exception
                    txtHeso.Text = ""
                End Try
            End If
        End If
    End Sub

    Private Sub cboLoaiPC_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLoaiPC.SelectedValueChanged
        If Not (cboLoaiPC.SelectedValue Is Nothing) Then
            If CInt(cboLoaiPC.SelectedValue) > 0 Then cboMucPC.DataSource = listMucPC(CInt(cboLoaiPC.SelectedValue))
        End If
    End Sub

    Private Sub cbLuong_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbLuong.CheckedChanged
        InfLuong(cbLuong.Checked)
    End Sub

    Private Sub cbPhucap_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPhucap.CheckedChanged
        InfPhucap(cbPhucap.Checked)
    End Sub

    Private Sub bntAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Try
            Dim lab_ErrQDNS As String = ""
            If IDcanbo_Err = "" Then
                MessageBox.Show("Hãy chọn lại thông tin sửa lỗi từ danh sách đã thông báo !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If

            lab_ErrQDNS = checkTabQDNhanSu()
            If lab_ErrQDNS <> "" Then
                MessageBox.Show(lab_ErrQDNS, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Exit Sub
            End If

            Dim IDDV_Moi, IDP_Moi, IDDV_Cu, IDP_Cu As Integer
            IDDV_Moi = CInt(cboDonviMoi.SelectedValue)
            IDP_Moi = CInt(cboPhongMoi.SelectedValue)
            IDDV_Cu = IdDonViCu
            IDP_Cu = IdPhongCu
            If updateQDNhanSu(IDQdNhanSu_Err) Then
                labAlert_QD.Text = "Ghi dữ liệu thành công!"
            Else
                MessageBox.Show("Ghi dữ liệu không thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntCancel.Click
        If IDQdNhanSu_Err = "" Then
            blankQDNhansu()
        Else
            fillQDNhansu(IDQdNhanSu_Err)
        End If
    End Sub

  
End Class