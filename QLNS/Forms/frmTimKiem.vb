Imports Microsoft.Office.Interop.Excel
'Imports Office = Microsoft.Office.Core

Public Class frmTimKiem

    Private init As Boolean = False
    Private init_P As Boolean = False
    Private init_CV As Boolean = False
    Private init_DT As Boolean = False
    Private init_TG As Boolean = False
    Private ComDset As New DataSet
    Private arrWHERE As ArrayList = New ArrayList()

    Private Sub frmTimKiem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If init = False Then
            cboDonVi.DataSource = listDonvi(True)
            bindCboPhong(CInt(cboDonVi.SelectedValue))
            cboChucvu.DataSource = listDanhmuc(14)
            cboDanToc.DataSource = listDanhmuc(21)
            cboTonGiao.DataSource = listDanhmuc(24)
            'bindCboToanTuTraCuu(cboToanTu)
            Label12.Visible = False
            cbHSL.Visible = False
            cboToanTu.Visible = False
            txtHeSo.Visible = False
            blankFrm()
            init = True
        End If
        AddHandler bntSearch.Click, AddressOf btnUpdateClicked
        AddHandler bntRefresh.Click, AddressOf btnCancelClicked
    End Sub

    Private Sub frmTimKiem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub cbNS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbNS.CheckedChanged
        If cbNS.Checked Then
            dpkTuNgay.Enabled = True
            dpkDenNgay.Enabled = True
        Else
            dpkTuNgay.Enabled = False
            dpkDenNgay.Enabled = False
        End If
    End Sub

    Private Sub cbPhong_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbPhong.CheckedChanged
        If cbPhong.Checked Then
            cboPhong.Enabled = True
        Else
            cboPhong.Enabled = False
        End If
    End Sub

    Private Sub bindCboPhong(ByVal vIdDonvi As Integer)
        Try
            If vIdDonvi > 0 Then
                cboPhong.DataSource = listPhong(vIdDonvi)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub blankFrm()
        cbAll.Checked = False
        txtHoten.Text = ""
        txtMaCB.Text = ""
        cbNam.Checked = False
        cbNu.Checked = False
        cbNS.Checked = False
        cbNgayNHCS.Checked = False
        dpkTuNgay.Enabled = False
        dpkDenNgay.Enabled = False
        dpkNgayNHCS_Tu.Enabled = False
        dpkNgayNHCS_Den.Enabled = False
        cbPhong.Checked = False
        cboPhong.Enabled = False
        cbChucvu.Checked = False
        cboChucvu.Enabled = False
        cbHSL.Checked = False
        cboToanTu.Enabled = False
        txtHeSo.Text = ""
        cbDT.Checked = False
        cboDanToc.Enabled = False
        cbTonGiao.Checked = False
        cboTonGiao.Enabled = False
        cbHuuTri.Checked = False
        dpkHuuTri.Enabled = False
        txtHuuNu.Enabled = False
        txtHuuNam.Enabled = False
        txtHuuNu.Text = "55"
        txtHuuNam.Text = "60"
    End Sub

    Private Sub bindGridResult(ByVal strSQL As String)
        Try
            Try
                Dim db As DBAccess = New DBAccess
                Dim dbconn As DBAccess = New DBAccess
                Dim conn As SqlClient.SqlConnection = dbconn.getConnection
                Dim adp As New SqlClient.SqlDataAdapter(strSQL, conn)
                ComDset.Reset()
                adp.Fill(ComDset, "TTbl")
                If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
                    grpResult.Text = "Kết quả tra cứu _ Số bản ghi tìm thấy: 0"
                    Exit Sub
                End If
                gridResult.AutoGenerateColumns = True
                gridResult.RowHeadersVisible = False
                gridResult.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
                gridResult.AllowUserToAddRows = False

                gridResult.Columns().Clear()
                gridResult.DataSource = ComDset.Tables(0)
                grpResult.Text = "Kết quả tra cứu _ Số bản ghi tìm thấy: " & ComDset.Tables(0).Rows.Count
            Catch ex As Exception
                grpResult.Text = "Kết quả tra cứu _ Số bản ghi tìm thấy: 0"
                MessageBox.Show("Kết quả tra cứu không có: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            End Try
        Catch ex As Exception
            MessageBox.Show("Không load được danh sách: " & ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub cboDonVi_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDonVi.SelectedIndexChanged
        bindCboPhong(CInt(cboDonVi.SelectedValue))
    End Sub

    Private Sub txtHeSo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHeSo.TextChanged
        If txtHeSo.Text <> "" Then
            txtHeSo.Text = formatDouble(txtHeSo.Text.Trim)
            txtHeSo.SelectionStart = txtHeSo.Text.Length
        End If
    End Sub

    Private Sub txtHuuNu_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHuuNu.TextChanged
        txtHuuNu.Text = Val(txtHuuNu.Text.Trim)
    End Sub

    Private Sub txtHuuNam_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHuuNam.TextChanged
        txtHuuNam.Text = Val(txtHuuNam.Text.Trim)
    End Sub

    Private Sub cbChucvu_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbChucvu.CheckedChanged
        If cbChucvu.Checked Then
            cboChucvu.Enabled = True
        Else
            cboChucvu.Enabled = False
        End If
    End Sub

    Private Sub cbHSL_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbHSL.CheckedChanged
        If cbHSL.Checked Then
            cboToanTu.Enabled = True
            txtHeSo.Enabled = True
        Else
            cboToanTu.Enabled = False
            txtHeSo.Enabled = False
        End If
    End Sub

    Private Sub cbDT_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbDT.CheckedChanged
        If cbDT.Checked Then
            cboDanToc.Enabled = True
        Else
            cboDanToc.Enabled = False
        End If
    End Sub

    Private Sub cbTonGiao_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTonGiao.CheckedChanged
        If cbTonGiao.Checked Then
            cboTonGiao.Enabled = True
        Else
            cboTonGiao.Enabled = False
        End If
    End Sub

    Private Sub cbHuuTri_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbHuuTri.CheckedChanged
        If cbHuuTri.Checked Then
            dpkHuuTri.Enabled = True
            txtHuuNu.Enabled = True
            txtHuuNam.Enabled = True
        Else
            dpkHuuTri.Enabled = False
            txtHuuNu.Enabled = False
            txtHuuNam.Enabled = False
        End If
    End Sub

    Private Sub btnUpdateClicked(ByVal source As Object, ByVal e As EventArgs)
        Dim strSQL As String = ""
        Dim strWHERE1 As String = ""
        Dim strWHERE2 As String = ""
        Dim strWHERE1ts As String = ""
        Dim strWHERE As String = ""
        Dim strWHEREts As String = ""
        Dim strTen_vt As String = ""
        Dim strMaDV As String = ""
        Try
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            strMaDV = getDonvi_Ma(cboDonVi.SelectedValue)
            If strMaDV = gMaDonViTW Then
                strWHERE1 = " AND (t.idDonvi_Moi = " & cboDonVi.SelectedValue & ")"
                strWHERE1ts = " AND (ts.idChiNhanh = " & cboDonVi.SelectedValue & ")"
            Else
                strWHERE1 = " AND (t.idDonvi_Moi = " & cboDonVi.SelectedValue & " OR t.idDonvi_Moi in (SELECT [Id] FROM Chinhanh WHERE id_goc=" & cboDonVi.SelectedValue & "))"
                strWHERE1ts = " AND (ts.idChiNhanh = " & cboDonVi.SelectedValue & " OR ts.idChiNhanh in (SELECT [Id] FROM Chinhanh WHERE id_goc=" & cboDonVi.SelectedValue & "))"
            End If
            If cbAll.Checked Then
                strWHERE1 = " AND (t.idDonvi_Moi=" & cboDonVi.SelectedValue & " OR t.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ") OR t.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ")))"
                strWHERE1ts = " AND (ts.idChiNhanh=" & cboDonVi.SelectedValue & " OR ts.idChiNhanh in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ") OR ts.idChiNhanh in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ")))"
            End If
            If Not (strMaDV = gMaDonViTW And cbAll.Checked) Then
                Dim db As DBAccess = New DBAccess
                strTen_vt = db.getString("SELECT ten_vt FROM ChiNhanh WHERE [id]=" & cboDonVi.SelectedValue)
            End If
            'If getDonvi_Ma(cboDonVi.SelectedValue) = gMaDonViTW Then
            '    strWHERE1 = " AND (t1.idDonvi_Moi = " & cboDonVi.SelectedValue & ")"
            '    strWHERE1ts = " AND (ts.idChiNhanh = " & cboDonVi.SelectedValue & ")"
            'Else
            '    strWHERE1 = " AND (t1.idDonvi_Moi = " & cboDonVi.SelectedValue & " OR t1.idDonvi_Moi in (SELECT [Id] FROM Chinhanh WHERE id_goc=" & cboDonVi.SelectedValue & "))"
            '    strWHERE1ts = " AND (ts.idChiNhanh = " & cboDonVi.SelectedValue & " OR ts.idChiNhanh in (SELECT [Id] FROM Chinhanh WHERE id_goc=" & cboDonVi.SelectedValue & "))"
            'End If
            'If cbAll.Checked Then
            '    strWHERE1 = " AND (t1.idDonvi_Moi=" & cboDonVi.SelectedValue & " OR t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ") OR t1.idDonvi_Moi in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ")))"
            '    strWHERE1ts = " AND (ts.idChiNhanh=" & cboDonVi.SelectedValue & " OR ts.idChiNhanh in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ") OR ts.idChiNhanh in (SELECT id FROM ChiNhanh WHERE id_goc in (SELECT id FROM ChiNhanh WHERE id_goc=" & cboDonVi.SelectedValue & ")))"
            'End If
            If txtHoten.Text.Trim <> "" Then
                strWHERE &= " AND (Hoten like N'%" & txtHoten.Text & "%')"
                strWHEREts &= " AND (Hoten like N'%" & txtHoten.Text & "%')"
            End If
            If txtMaCB.Text.Trim <> "" Then
                strWHERE &= " AND (MaCB like N'%" & txtMaCB.Text.Trim & "%')"
                strWHEREts &= " AND (MaCB like N'%" & txtMaCB.Text.Trim & "%')"
            End If
            If Not (cbNam.Checked And cbNu.Checked) Then
                If cbNam.Checked Then
                    strWHERE &= " AND (Gioitinh =0)"
                    strWHEREts &= " AND (Gioitinh =0)"
                End If
                If cbNu.Checked Then
                    strWHERE &= " AND (Gioitinh =1)"
                    strWHEREts &= " AND (Gioitinh =1)"
                End If
            End If
            If cbNS.Checked Then
                strWHERE &= " AND ( NgaySinh >= '" & dpkTuNgay.Value & "' AND NgaySinh <='" & dpkDenNgay.Value & "')"
                strWHEREts &= " AND ( NgaySinh >= '" & dpkTuNgay.Value & "' AND NgaySinh <='" & dpkDenNgay.Value & "')"
            End If
            If cbNgayNHCS.Checked Then
                strWHERE &= " AND ( Ngay_VBSP >= '" & dpkNgayNHCS_Tu.Value & "' AND Ngay_VBSP <='" & dpkNgayNHCS_Den.Value & "')"
                strWHEREts &= " AND ( Ngay_VBSP >= '" & dpkNgayNHCS_Tu.Value & "' AND Ngay_VBSP <='" & dpkNgayNHCS_Den.Value & "')"
            End If
            If cbPhong.Checked Then
                strWHERE &= " AND (t.IdPhong_Moi = " & cboPhong.SelectedValue & ")"
                strWHEREts &= " AND (substring(ts.idPhongBan,4,len(ts.idPhongBan)-3)) = " & cboPhong.SelectedValue & ")"
            End If
            If cbChucvu.Checked Then
                strWHERE &= " AND (t.IdChucVu_Moi = " & cboChucvu.SelectedValue & ")"
                strWHEREts &= " AND (SoTruong_CT = '" & cboChucvu.SelectedValue & "')"
            End If
            'If cbHSL.Checked Then
            '    strWHERE2 &= " AND (hesoluong " & cboToanTu.SelectedItem & txtHeSo.Text.Trim & ")"
            'End If
            If cbDT.Checked Then
                strWHERE &= " AND (IdDanToc = " & cboDanToc.SelectedValue & ")"
                strWHEREts &= " AND (IdDanToc = " & cboDanToc.SelectedValue & ")"
            End If
            If cbTonGiao.Checked Then
                strWHERE &= " AND (IdTonGiao = " & cboTonGiao.SelectedValue & ")"
                strWHEREts &= " AND (IdTonGiao = " & cboTonGiao.SelectedValue & ")"
            End If
            If cbHuuTri.Checked Then
                strWHERE &= " AND ((DATEDIFF(yyyy, NgaySinh,'" & dpkHuuTri.Value & "')>=" & txtHuuNu.Text & " AND (Gioitinh =1)) OR (DATEDIFF(yyyy, NgaySinh,'" & dpkHuuTri.Value & "')>=" & txtHuuNam.Text & " AND (Gioitinh =0)))"
                strWHEREts &= " AND ((DATEDIFF(yyyy, NgaySinh,'" & dpkHuuTri.Value & "')>=" & txtHuuNu.Text & " AND (Gioitinh =1)) OR (DATEDIFF(yyyy, NgaySinh,'" & dpkHuuTri.Value & "')>=" & txtHuuNam.Text & " AND (Gioitinh =0)))"
            End If
            If cbSoBHXH.Checked Then
                strWHERE &= " AND (BHXH_SoSo ='' or BHXH_SoSo is null)"
            End If
            If strTen_vt.Trim <> "" Then
                strWHERE &= " AND (left(t1.Idcanbo,4)='" & strTen_vt & "')"
            End If

            strSQL = "SELECT MaCB as N'Mã Cán bộ', HoTen as N'Họ tên',(case gioitinh when 0 then N'Nam' when 1 then N'Nữ' end) as N'Giới tính', convert(varchar, NgaySinh, 103) as N'Ngày sinh'," & _
                          " t1.NQ_DChi as N'Nguyên quán', t1.TTr_DiaChi as N'Địa chỉ thường trú' , t1.CMT_So as 'Số CMT' , convert(varchar, t1.CMT_NgayCap, 103) as N'Ngày cấp', t1.CMT_noicap as N'Nơi cấp'," & _
                          "((SELECT ten_goi FROM ChiNhanh WHERE id=t.idDonvi_Moi) + (select case when t.idDonvi_Moi<70 then '' else ', ' +(select ten_goi from chinhanh where id in (SELECT id_goc FROM ChiNhanh WHERE id=t.idDonvi_Moi)) end)) as N'Đơn vị',  " & _
                          "(SELECT Ten_phong FROM PhongBan WHERE id=t.idPhong_Moi) as N'Phòng ban', " & _
                          "(SELECT Ten_goi FROM DanhMuc WHERE id=t.IdChucVu_Moi ) as N'Chức vụ', " & _
                          "(SELECT TOP 1 hesoluong FROM hs_luongcb WHERE IdCanbo=t.Idcanbo " & strWHERE2 & " order by Ngay_Huong desc) as N'Hệ số', " & _
                          "(SELECT TOP 1 convert(varchar, Ngay_Huong, 103) FROM hs_luongcb WHERE IdCanbo=t.Idcanbo " & strWHERE2 & " order by Ngay_Huong desc) as N'Ngày hưởng lương', " & _
                          "(SELECT ten_goi FROM DanhMuc where id_goc=21 and [id]=IdDanToc) as N'Dân tộc', (SELECT ten_goi FROM DanhMuc where id_goc=36 and [id]=IdTrinhDoCT) as N'Trình độ chính trị', " & _
                          "convert(varchar, Ngay_NH, 103) as N'Ngày vào NH', convert(varchar, Ngay_VBSP, 103) as N'Ngày vào VBSP', " & _
                          "(SELECT TOP 1 convert(varchar, NgayVao, 103) FROM HS_DangVien WHERE IdCanBo=t1.idCanbo  AND (NgayRa Is Null Or NgayRa <= getdate())  order by NgayRa desc) as N'Ngày vào Đảng'," & _
                          "DienThoai_NR as 'Phone', DienThoai_DD as 'Mobile', NH_SoTK as N'Số tài khoản', NH_TenNH as N'Ngân hàng', MaSoThue as N'Mã số thuế', BHXH_SoSo as N'Số sổ BHXH',  " & _
                          "(SELECT TOP 1 (SELECT ten_goi FROM DanhMuc WHERE Id=IdTrinhDo) FROM HS_DTVBCC  WHERE IdTrinhDo in (SELECT [id] FROM DanhMuc Where ma_so in (SELECT min(ma_so) FROM HS_DTVBCC s1, DanhMuc s2, HS_CanBo s3 WHERE s1.IdTrinhDo= s2.Id and s1.idCanBo=s3.idCanBo and t1.idCanbo=s1.idCanBo group by s1.idCanBo)) AND IdCanBo=t1.idCanbo order by IdChuyenNganhDT, TuNgay desc) as N'Trình độ', " & _
                          "(SELECT TOP 1 CoSo_DT FROM HS_DTVBCC  WHERE IdTrinhDo in (SELECT [id] FROM DanhMuc Where ma_so in (SELECT min(ma_so) FROM HS_DTVBCC s1, DanhMuc s2, HS_CanBo s3 WHERE s1.IdTrinhDo= s2.Id and s1.idCanBo=s3.idCanBo and t1.idCanbo=s1.idCanBo group by s1.idCanBo)) AND IdCanBo=t1.idCanbo order by IdChuyenNganhDT, TuNgay desc) as N'Cơ sở đào tạo', " & _
                          "(SELECT TOP 1 (SELECT ten_goi FROM DanhMuc WHERE Id=IdHinhThucDT) FROM HS_DTVBCC  WHERE IdTrinhDo in (SELECT [id] FROM DanhMuc Where ma_so in (SELECT min(ma_so) FROM HS_DTVBCC s1, DanhMuc s2, HS_CanBo s3 WHERE s1.IdTrinhDo= s2.Id and s1.idCanBo=s3.idCanBo and t1.idCanbo=s1.idCanBo group by s1.idCanBo)) AND IdCanBo=t1.idCanbo order by IdChuyenNganhDT, TuNgay desc) as N'Hệ đào tạo'," & _
                          "(SELECT TOP 1 (SELECT ten_goi FROM DanhMuc WHERE Id=IdChuyenNganhDT) FROM HS_DTVBCC  WHERE IdTrinhDo in (SELECT [id] FROM DanhMuc Where ma_so in (SELECT min(ma_so) FROM HS_DTVBCC s1, DanhMuc s2, HS_CanBo s3 WHERE s1.IdTrinhDo= s2.Id and s1.idCanBo=s3.idCanBo and t1.idCanbo=s1.idCanBo group by s1.idCanBo)) AND IdCanBo=t1.idCanbo order by IdChuyenNganhDT, TuNgay desc) as N'Chuyên ngành đào tạo'," & _
                          "(SELECT TOP 1 NamTN FROM HS_DTVBCC  WHERE IdTrinhDo in (SELECT [id] FROM DanhMuc Where ma_so in (SELECT min(ma_so) FROM HS_DTVBCC s1, DanhMuc s2, HS_CanBo s3 WHERE s1.IdTrinhDo= s2.Id and s1.idCanBo=s3.idCanBo and t1.idCanbo=s1.idCanBo group by s1.idCanBo)) AND IdCanBo=t1.idCanbo order by IdChuyenNganhDT, TuNgay desc) as N'Năm tốt nghiệp'," & _
                          "(SELECT TOP 1 (SELECT ten_goi FROM DanhMuc WHERE Id=IdLoaiHD) FROM hs_hdld WHERE idcanbo=t1.idCanbo order by Ngay_HL desc) as N'Loại HĐLĐ'," & _
                          "(SELECT TOP 1 SoHD FROM hs_hdld WHERE idcanbo=t1.idCanbo order by Ngay_HL desc) as N'Số HĐLĐ'," & _
                          "(SELECT TOP 1 convert(varchar, Ngay_HL, 103) FROM hs_hdld WHERE idcanbo=t1.idCanbo order by Ngay_HL desc) as N'Ngày hiệu lực HĐLĐ'," & _
                          "(SELECT TOP 1 convert(varchar, NgayKy_HD, 103) FROM hs_hdld WHERE idcanbo=t1.idCanbo order by Ngay_HL desc) as N'Ngày ký HĐLĐ'," & _
                          "(SELECT TOP 1 DVKyHDLD FROM hs_hdld WHERE idcanbo=t1.idCanbo order by Ngay_HL desc) as N'Đơn vị ký HĐLĐ'" & _
                    " FROM QDNhansu t, HS_Canbo t1, (SELECT Idcanbo, max(NgayHL) as NgayHL from QDNhansu WHERE isKiemNhiem=0 AND IsQD_NHCS =1 group by idcanbo having max(NgayHL)<=getdate()) t2 " & _
                    " WHERE(t.idCanbo = t1.idCanbo And t.idcanbo = t2.idcanbo And t.NgayHL = t2.NgayHL " & strWHERE1 & ")" & _
                    "              AND t1.idCanbo not in (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t1.idCanbo AND t4.Ngay_HL >t.NgayHL ) "
            strSQL = strSQL & strWHERE & _
                   "Order by t.idDonvi_Moi, t.idPhong_Moi, t.idChucvu_moi, Hoten "

            '---------------------------------------------
            'strSQL = "SELECT MaCB as N'Mã Cán bộ', HoTen as N'Họ tên',(case gioitinh when 0 then N'Nam' when 1 then N'Nữ' end) as N'Giới tính', convert(varchar, NgaySinh, 3) as N'Ngày sinh'," & _
            '              "(SELECT ten_goi FROM ChiNhanh WHERE id=t.idDonvi_Moi) as N'Đơn vị',  " & _
            '              "(SELECT Ten_phong FROM PhongBan WHERE id=t.idPhong_Moi) as N'Phòng ban', " & _
            '              "(SELECT Ten_goi FROM DanhMuc WHERE id=t.IdChucVu_Moi ) as N'Chức vụ', " & _
            '              "(SELECT TOP 1 hesoluong FROM hs_luongcb WHERE IdCanbo=t.Idcanbo " & strWHERE2 & " order by Ngay_Huong desc) as N'Hệ số', " & _
            '              "(SELECT TOP 1 convert(varchar, Ngay_Huong, 3) FROM hs_luongcb WHERE IdCanbo=t.Idcanbo " & strWHERE2 & " order by Ngay_Huong desc) as N'Ngày hưởng lương', " & _
            '              "convert(varchar, Ngay_VBSP, 3) as N'Ngày vào NH', DienThoai_NR as 'Phone', DienThoai_DD as 'Mobile', NH_SoTK as N'Số tài khoản', NH_TenNH as N'Ngân hàng', MaSoThue as N'Mã số thuế', BHXH_SoSo as N'Số sổ BHXH'  " & _
            '         "FROM HS_Canbo t2, " & _
            '         "	   (SELECT t1.idCanbo, t1.idDonvi_Moi, t1.idPhong_Moi, t1.IdChucVu_Moi " & _
            '         "	    FROM QDNhansu t1 " & _
            '         "      WHERE(t1.isKiemNhiem = 0 AND t1.IsQD_NHCS=1) " & _
            '         "	      AND t1.idCanbo not in (SELECT t4.IdCanbo FROM (SELECT IdCanBo, max(Ngay_HL) as Ngay_HL FROM HS_CBThoiviec WHERE IsQD_NHCS =1 GROUP BY Idcanbo Having max(Ngay_HL)<=getdate()) t4 WHERE t4.idCanbo= t1.idCanbo AND t4.Ngay_HL >t1.NgayHL) " & _
            '         "	      AND t1.idcanbo not in (SELECT idcanbo FROM QDNhansu t5 WHERE t5.idCanbo=t1.idCanbo AND datediff(second,t1.ngayHL,t5.ngayHL)>0)  "
            'strSQL = strSQL & strWHERE1 & "	    ) t " & _
            '         "WHERE (t.idCanbo = t2.idCanbo)" & strWHERE & _
            '         "Order by idDonvi_Moi, idPhong_Moi, idChucvu_moi, Hoten "
            '---------------------------------------------
            '" UNION " & _
            '"SELECT MaCB as N'Mã Cán bộ', HoTen as N'Họ tên',(case gioitinh when 0 then N'Nam' when 1 then N'Nữ' end) as N'Giới tính', convert(varchar, NgaySinh, 3) as N'Ngày sinh'," & _
            '"(SELECT ten_goi FROM ChiNhanh WHERE id=ts.idChiNhanh) as idDonvi_Moi,  " & _
            '"(SELECT Ten_phong FROM PhongBan WHERE id=substring(ts.idPhongBan,4,len(ts.idPhongBan)-3)) as idPhong_Moi, " & _
            '"'LĐNH' as idChucvu_moi, " & _
            '"'0' as N'Hệ số', " & _
            '"convert(varchar, TuNgay, 3) as N'Ngày hưởng lương', " & _
            '"convert(varchar, TuNgay, 3) as N'Ngày vào NH', DienThoai_NR as 'Phone', DienThoai_DD as 'Mobile', NH_SHTK as N'Số tài khoản', NH_Ten_NH as N'Ngân hàng', MaSoThue as N'Mã số thuế', BHXH_SoSo as N'Số sổ BHXH'  " & _
            '"FROM V$_HSCB_TS_HDLD ts WHERE TuNgay<= getdate() AND getdate()<=DenNgay AND (IdNew is null or IdNew='')" & strWHERE1ts & strWHEREts & _


            bindGridResult(strSQL)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub btnCancelClicked(ByVal source As Object, ByVal e As EventArgs)
        blankFrm()
    End Sub

    Private Sub bntExportFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExportFile.Click
        Try
            Cursor = Cursors.WaitCursor
            labStatusProcess.Text = "Waiting ....."
            Load_Excel_Details(ComDset)
            labStatusProcess.Text = "Done"
            Cursor = Cursors.Default
        Catch ex As Exception
            labStatusProcess.Text = "Error"
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
        End Try
    End Sub

    Private Sub bntClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntClose.Click
        Close()
    End Sub

    Private Sub Load_Excel_Details(ByVal ds As DataSet)
        Dim col, row As Integer
        If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
            Exit Sub
        End If
        Dim Excel As Object = CreateObject("Excel.Application")
        If Excel Is Nothing Then
            MessageBox.Show("Excel chưa được cài đặt trên máy. Để tiếp tục yêu cầu cài đặt MS Excel", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End If

        Try
            'Thay đổi setting Regional và trả lại sau khi kết thúc công việc
            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            With Excel
                'không cho hiện ứng dụng Excel lên để tránh gây đơ máy
                .Visible = False
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()
                '.Worksheets(1).Columns.AutoFit()
                '.Worksheets(1).Columns.EntireColumn.AutoFit()

                '                caption.Select();
                'caption.FormulaR1C1 = tieude;
                '//căn lề cho tiêu đề
                'caption.HorizontalAlignment = Excel.Constants.xlCenter;

                Dim i As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    .cells(1, i).value = ComDset.Tables(0).Columns(col).ColumnName
                    .cells(1, i).Font.Size = 12
                    .cells(1, i).Font.Bold = True
                    '.cells(1, i).DegreeAlignment = 90
                    '.cells(1, i).VerticalAlignment = -4108
                    '.cells(1, i).HorizontalAlignment = -4108
                    'wb.Worksheets("Sheet1").Columns("B:B").NumberFormat = "m/d/yyyy;@
                    .cells(1, i).EntireColumn.AutoFit()


                    'If i = 1 Then
                    '    .cells(1, i).ColumnWidth = 15
                    'Else
                    '    If i = 2 Then
                    '        .cells(1, i).ColumnWidth = 25
                    '    Else
                    '        .cells(1, i).ColumnWidth = 20
                    '    End If
                    'End If
                    i += 1
                Next
                i = 2
                Dim k As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    i = 2
                    For row = 0 To ComDset.Tables(0).Rows.Count - 1
                        .Cells(i, k).Value = ComDset.Tables(0).Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next

                SaveToExcel(.ActiveCell.Worksheet, "TraCuu_" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls", True)

            End With
            'Trả lại thiết lập cũ
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            MsgBox(ex.Message)
        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub cbNgayNHCS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbNgayNHCS.CheckedChanged
        If cbNgayNHCS.Checked Then
            dpkNgayNHCS_Tu.Enabled = True
            dpkNgayNHCS_Den.Enabled = True
        Else
            dpkNgayNHCS_Tu.Enabled = False
            dpkNgayNHCS_Den.Enabled = False
        End If
    End Sub

   
End Class