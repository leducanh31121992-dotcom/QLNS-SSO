Imports System.Configuration

Public Class frmHT_ThamSo
#Region "---> Defined parametter and properties <---"
    Private strSQL As String = ""
    Private _RecordId As Integer = 0
    Private arr_Branch As ArrayList = New ArrayList
    Private _Globals As Globals = New Globals
    Private _ChiNhanhBLL As ChiNhanhBLL = New ChiNhanhBLL()

    Private _ValSys As String = ""
    Dim _SQLHelper As New DBAccess

    Dim _DanhMuc As clsHT_DanhMuc = New clsHT_DanhMuc
    Dim _HeThongBLL As HeThongBLL = New HeThongBLL
    Private ARL_DonVi As ArrayList = New ArrayList
    Private ARL_GiamDoc As ArrayList = New ArrayList
    Private ARL_PhoGiamDoc As ArrayList = New ArrayList
    Private ARL_HcToChuc As ArrayList = New ArrayList
    Private ARL_KeToan As ArrayList = New ArrayList
    Private ARL_LuongTTV As ArrayList = New ArrayList
#End Region

#Region "---> Functions and events <---"
    ''' <summary>
    ''' Hàm thực hiện reset controls về trạng thái ban đầu
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Reset_Controls()
        _RecordId = 0
        dtpk_ngayapdung.Text = Globals.GetDateTime_ForServerDB()
        numr_sokyluong.Value = 2
        rb_cotamtinh.Checked = True
        lbl_donvi_tenvt.Text = ""
        edt_giamdoc.Text = ""
        edt_phogiamdoc.Text = ""
        edt_tphanhchinhtc.Text = ""
        edt_ketoan.Text = ""
        edt_sotien_luongvtt.Text = "0"
        edt_tendiaban.Text = ""
        edt_muctamung_v2.Text = "0"
        edt_hesotamung_v2.Text = "0"

        If cb_donvi.Items.Count <> 0 Then
            cb_donvi.SelectedIndex = 0
        End If
        If cb_giamdoc.Items.Count <> 0 Then
            cb_giamdoc.SelectedIndex = 0
        End If
        If cb_phogiamdoc.Items.Count <> 0 Then
            cb_phogiamdoc.SelectedIndex = 0
        End If
        If cb_hctochuc.Items.Count <> 0 Then
            cb_hctochuc.SelectedIndex = 0
        End If
        If cb_ketoan.Items.Count <> 0 Then
            cb_ketoan.SelectedIndex = 0
        End If
        If cb_vungluongtt.Items.Count <> 0 Then
            cb_vungluongtt.SelectedIndex = 0
        End If

        rb_cotamtinh.Checked = True
        rb_khongtamtinh.Checked = False
    End Sub

    ''' <summary>
    ''' Hàm thực hiện fill dữ liệu vào các controls
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Bind_DataGridView()
        dgv_main.Rows.Clear()
        'Dim _PosCdSearch = IIf(Cap_Nd = 1, "", IIf(Cap_Nd = 2, DONVI.Substring(0, 4), DONVI))
        Dim _PosCdSearch As String = ""
        If Cap_Nd = 3 Then
            _PosCdSearch = ""
        ElseIf (Cap_Nd = 2 And DONVI <> "000196" And DONVI <> "000197" And DONVI <> "000199" And DONVI <> "000101") Then
            _PosCdSearch = DONVI.Substring(0, 4)
        Else
            _PosCdSearch = DONVI
        End If
        Using db As DataTable = _HeThongBLL.GetListSysVarSearch(0, 0, _PosCdSearch, "", "", "", CType("0", Byte))
            If Not (db Is Nothing) Then
                If (db.Rows.Count > 0) Then
                    For i As Integer = 0 To db.Rows.Count - 1
                        dgv_main.Rows.Add()
                        dgv_main.Rows(i).Cells("cln_Id").Value = IIf(db.Rows(i)("Id").ToString() <> "", db.Rows(i)("Id").ToString(), "")
                        dgv_main.Rows(i).Cells("cln_Id_DonVi").Value = IIf(db.Rows(i)("Id_DonVi").ToString() <> "", db.Rows(i)("Id_DonVi").ToString(), "")
                        dgv_main.Rows(i).Cells("cln_TrucThuoc").Value = IIf(db.Rows(i)("TrucThuoc").ToString() <> "", db.Rows(i)("TrucThuoc").ToString(), "")
                        dgv_main.Rows(i).Cells("cln_STT").Value = db.Rows(i)("STT").ToString()
                        dgv_main.Rows(i).Cells("cln_DonVi_Cd").Value = db.Rows(i)("DonVi_Cd").ToString()
                        dgv_main.Rows(i).Cells("cln_DonVi_HT").Value = db.Rows(i)("DonVi_HT").ToString()
                        dgv_main.Rows(i).Cells("cln_Ten_VT").Value = db.Rows(i)("Ten_VT").ToString()
                        dgv_main.Rows(i).Cells("cln_DiaBan").Value = db.Rows(i)("DiaBan").ToString()
                        dgv_main.Rows(i).Cells("cln_Max_Ky").Value = db.Rows(i)("Max_Ky").ToString()
                        dgv_main.Rows(i).Cells("cln_Tinh_Thue_TNCN").Value = db.Rows(i)("Tinh_Thue_TNCN_HT").ToString()
                        dgv_main.Rows(i).Cells("cln_VungLuongTT").Value = db.Rows(i)("VungLuongTT_HT").ToString()
                        dgv_main.Rows(i).Cells("cln_HeSoTamUng_V2").Value = db.Rows(i)("HeSoTamUng_V2").ToString()
                        dgv_main.Rows(i).Cells("cln_MucTamUng_V2").Value = db.Rows(i)("MucTamUng_V2").ToString()
                        dgv_main.Rows(i).Cells("cln_GiamDoc").Value = db.Rows(i)("GiamDoc").ToString()
                        dgv_main.Rows(i).Cells("cln_PhoGiamDoc").Value = db.Rows(i)("PhoGiamDoc").ToString()
                        dgv_main.Rows(i).Cells("cln_TruongHCTC").Value = db.Rows(i)("TruongHCTC").ToString()
                        dgv_main.Rows(i).Cells("cln_TruongKT").Value = db.Rows(i)("KeToanTruong").ToString()
                        dgv_main.Rows(i).Cells("cln_NgayHL").Value = db.Rows(i)("NgayHL_HT").ToString()
                        If (CType(db.Rows(i)("TrangThai").ToString(), Byte) = 0) Then
                            dgv_main.Rows(i).DefaultCellStyle.ForeColor = System.Drawing.Color.Red
                            dgv_main.Rows(i).DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Red
                        Else
                            dgv_main.Rows(i).DefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        End If
                    Next
                End If
            End If
        End Using
    End Sub

    Private Sub SelectRowGrid(ByVal _RowId As Integer)
        Dim dr As DataRow
        _RecordId = _RowId
        If _RowId > 0 Then
            Using db As DataTable = _HeThongBLL.GetListSysVarSearch(_RowId, 0, "", "", "", "", CType("0", Byte))
                If Not (db Is Nothing) Then
                    If (db.Rows.Count > 0) Then
                        If db.Rows(0)("NgayHL").ToString().Trim() <> "" Then
                            dtpk_ngayapdung.Value = CType(db.Rows(0)("NgayHL").ToString(), DateTime)
                        End If
                        cb_donvi.SelectedIndex = CType(ARL_DonVi.IndexOf(db.Rows(0)("Id_DonVi").ToString()), Integer)
                        cb_donvi_SelectedIndexChanged(Nothing, Nothing)
                        lbl_donvi_tenvt.Text = db.Rows(0)("Ten_VT").ToString()
                        edt_tendiaban.Text = db.Rows(0)("DiaBan").ToString()

                        cb_giamdoc.SelectedIndex = CType(ARL_GiamDoc.IndexOf(db.Rows(0)("IdCanBo_GiamDoc").ToString()), Integer)
                        edt_giamdoc.Text = db.Rows(0)("GiamDoc").ToString()

                        cb_phogiamdoc.SelectedIndex = CType(ARL_PhoGiamDoc.IndexOf(db.Rows(0)("IdCanBo_PhoGD").ToString()), Integer)
                        edt_phogiamdoc.Text = db.Rows(0)("PhoGiamDoc").ToString()

                        cb_hctochuc.SelectedIndex = CType(ARL_HcToChuc.IndexOf(db.Rows(0)("IdCanBo_HCTC").ToString()), Integer)
                        edt_tphanhchinhtc.Text = db.Rows(0)("TruongHCTC").ToString()

                        cb_ketoan.SelectedIndex = CType(ARL_KeToan.IndexOf(db.Rows(0)("IdCanBo_KeToan").ToString()), Integer)
                        edt_ketoan.Text = db.Rows(0)("KeToanTruong").ToString()

                        cb_vungluongtt.SelectedIndex = CType(ARL_LuongTTV.IndexOf(db.Rows(0)("VungLuongTT").ToString()), Integer)
                        cb_vungluongtt.Text = db.Rows(0)("VungLuongTT_HT").ToString()

                        If CType(db.Rows(0)("Tinh_Thue_TNCN").ToString(), Byte) = 1 Then
                            rb_cotamtinh.Checked = True
                            rb_khongtamtinh.Checked = False
                        Else
                            rb_cotamtinh.Checked = False
                            rb_khongtamtinh.Checked = True
                        End If
                        numr_sokyluong.Value = CType(db.Rows(0)("Max_Ky").ToString(), Byte)

                        edt_hesotamung_v2.Text = db.Rows(0)("HeSoTamUng_V2").ToString().Replace(",", "").Trim()
                        edt_muctamung_v2.Text = db.Rows(0)("MucTamUng_V2").ToString().Replace(",", "").Trim()
                    End If
                End If
            End Using
        End If
        

    End Sub

    ''' <summary>
    ''' Hàm kiểm tra tính hợp lệ của Dữ liệu cập nhật
    ''' </summary>
    ''' <returns>True: Hợp lệ. False: Không hợp lệ</returns>
    ''' <remarks></remarks>
    Private Function IsValid() As Boolean
        Dim _Result = True
        Dim _CountExist As Integer = 0
        If (cb_donvi.SelectedIndex <= 0 And cb_donvi.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn đơn vị cần khai báo tham số hệ thống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_donvi
            _Result = False
            Return False
        End If

        Dim _DonViIdTMP As Integer = CType(IIf(ARL_DonVi.Count > 0, ARL_DonVi(cb_donvi.SelectedIndex), "0"), Integer)
        If dtpk_ngayapdung.Value.Date > DateTimeUtil.StringToDateTime("01/01/2020", "dd/MM/yyyy") Then
            If (_RecordId <> 0) Then
                strSQL = String.Format("Select Count(*) From SysVar Where TrangThai = 1 And Id_DonVi = {0} And NgayHL <= '2020-01-01' And Id <> {1} ", _DonViIdTMP, _RecordId)
            Else
                strSQL = String.Format("Select Count(*) From SysVar Where TrangThai = 1 And Id_DonVi = {0} And NgayHL <= '2020-01-01' ", _DonViIdTMP)
            End If
            _CountExist = SoftSqlHelper.GetNumber(strSQL, 0)
            If _CountExist <= 0 Then
                MessageBox.Show(String.Format("Tham số của đơn vị [{0}] chưa được khai báo từ thời điểm 01-01-2020. Vui lòng kiểm tra lại!" + vbCrLf + "Lưu ý: Đầu tiên phải khai báo tham số từ đầu năm 2020, sau đó tiếp tục thêm mới tham số cho đơn vị nếu có sự thay đổi (Khó đó Ngày hiệu lực sẽ chọn là ngày có sự thay đổi).", cb_donvi.SelectedItem, DateTimeUtil.DateTimeToString(dtpk_ngayapdung.Value, "dd/MM/yyyy")), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
                ActiveControl = dtpk_ngayapdung
                Return False
            End If
        End If
       

        If (edt_giamdoc.Text.Trim() = "" And cb_giamdoc.Items.Count > 1) Then
            MessageBox.Show("Thông tin tham số Giám đốc không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_giamdoc
            _Result = False
            Return False
        End If
        If (edt_phogiamdoc.Text.Trim() = "" And cb_phogiamdoc.Items.Count > 1) Then
            MessageBox.Show("Thông tin tham số phó giám đốc không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_phogiamdoc
            _Result = False
            Return False
        End If
        If (edt_tphanhchinhtc.Text.Trim() = "" And cb_hctochuc.Items.Count > 1) Then
            MessageBox.Show("Thông tin tham số trưởng phòng Hành chính - Tổ chức không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_tphanhchinhtc
            _Result = False
            Return False
        End If
        If (edt_ketoan.Text.Trim() = "" And cb_ketoan.Items.Count > 1) Then
            MessageBox.Show("Thông tin tham số Kế toán trưởng (TP. Kế toán) không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_ketoan
            _Result = False
            Return False
        End If
        
        If (cb_vungluongtt.SelectedIndex <= 0 And cb_vungluongtt.Items.Count <> 0) Then
            MessageBox.Show("Bạn chưa chọn Lương vùng tối thiểu áp dụng cho đơn vị cần khai báo tham số hệ thống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = cb_vungluongtt
            _Result = False
            Return False
        End If
        If (edt_tendiaban.Text.Trim() = "") Then
            MessageBox.Show("Tên địa bàn không được để trống!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = edt_tendiaban
            _Result = False
            Return False
        End If

        'HeSoTamUng_V2,MucTamUng_V2
        _CountExist = 0
        If (_RecordId <> 0) Then
            strSQL = String.Format("Select Count(*) From SysVar Where TrangThai = 1 And Id_DonVi = {0} And NgayHL = '{1}' And Id <> {2} ", _DonViIdTMP, DateTimeUtil.DateTimeToString(dtpk_ngayapdung.Value, "yyyy-MM-dd"), _RecordId)
        Else
            strSQL = String.Format("Select Count(*) From SysVar Where TrangThai = 1 And Id_DonVi = {0} And NgayHL = '{1}' ", _DonViIdTMP, DateTimeUtil.DateTimeToString(dtpk_ngayapdung.Value, "yyyy-MM-dd"))
        End If
        _CountExist = SoftSqlHelper.GetNumber(strSQL, 0)
        If _CountExist <> 0 Then
            MessageBox.Show(String.Format("Tham số của đơn vị [{0}] ngày hiệu lực [{1}] đã tồn tại. Vui lòng kiểm tra lại!", cb_donvi.SelectedItem, DateTimeUtil.DateTimeToString(dtpk_ngayapdung.Value, "dd/MM/yyyy")), "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3)
            ActiveControl = dtpk_ngayapdung
            Return False
        End If

        Return _Result
    End Function
#End Region

#Region "---> Các sự kiện người dùng <---"
    Private Sub frmHT_ThamSo_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not CheckSettings() AndAlso My_MessageBox("Thông số chương trình chưa đầy đủ, thiết lập lại tham số?") = vbYes Then
            e.Cancel = True
        End If
    End Sub

    Private Sub frmHT_ThamSo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If DONVI = gMaDonViTW Then
            lbl_giamdoc.Text = "Tổng giám đốc "
            lbl_phogiamdoc.Text = "P.Tổng giám đốc "
        Else
            lbl_giamdoc.Text = "Giám đốc "
            lbl_phogiamdoc.Text = "Phó giám đốc "
        End If
        
        'Check permit của thành viên đang thao tác
        If (Globals.Roles.IndexOf(";43;") < 0) Then
            'gb_main.Enabled = False
        End If
        If (Globals.Roles.IndexOf(";44;") < 0) Then
            'btn_accept.Visible = False
        End If

        If Cap_Nd = 1 Then
            strSQL = clsHT_DanhMuc.Sql_UnitAll
        ElseIf Cap_Nd = 2 Then
            If (DONVI = "000199" Or DONVI = "000196" Or DONVI = "000197" Or DONVI = "000101") Then
                strSQL = String.Format("Select Id,dbo.Replace_BranchName(Ten_Goi) Name,Ma_So Code From ChiNhanh Where Status=1 And Ma_So Like N'{0}%' Order By Substring(Ma_So,1,4), Id_Goc,Ma_So", DONVI)
            Else
                strSQL = String.Format("Select Id,dbo.Replace_BranchName(Ten_Goi) Name,Ma_So Code From ChiNhanh Where Status=1 And Ma_So Like N'{0}%' Order By Substring(Ma_So,1,4), Id_Goc,Ma_So", DONVI.Substring(0, 4))
            End If
        ElseIf Cap_Nd = 3 Then
            strSQL = String.Format("Select Id,dbo.Replace_BranchName(Ten_Goi) Name,Ma_So Code From ChiNhanh Where Status=1 And Ma_So Like N'{0}' Order By Substring(Ma_So,1,4), Id_Goc,Ma_So", DONVI)
        End If
        ARL_DonVi.Clear()
        cb_donvi.Items.Clear()
        ARL_DonVi = _Globals.Bind_ComBoBox(cb_donvi, strSQL, "---Đơn vị---")
        cb_donvi_SelectedIndexChanged(sender, Nothing)

        strSQL = "Select ValueKey,ValueDesc From SysListValue Where ListKey ='VUNGLUONGTT' And Status=1 Order By OrderNo"
        ARL_LuongTTV.Clear()
        cb_vungluongtt.Items.Clear()
        ARL_LuongTTV = _Globals.Bind_ComBoBox(cb_vungluongtt, strSQL, "---Chọn vùng lương tối thiểu đơn vị---")

        'Reset controls trên form cho mặc định trạng thái ban đầu
        Reset_Controls()
        _DanhMuc.Create_Frame(dgv_main, 12)
        Bind_DataGridView()
        dgv_main_CellClick(sender, Nothing)
    End Sub

    Private Sub btn_cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_cancel.Click
        'Reset controls trên form cho mặc định trạng thái ban đầu
        Reset_Controls()
        'Fill data ra các controls tham số hệ thống
        dgv_main_CellClick(sender, Nothing)
    End Sub

    Private Sub btn_accept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_accept.Click
        Dim _currRow As Integer = 0
        If (IsValid()) Then
            Dim _SysVar As HeThong.SysVar = New HeThong.SysVar()

            _SysVar.Id_DonVi = CType(IIf(ARL_DonVi.Count > 0, ARL_DonVi(cb_donvi.SelectedIndex), "0"), Integer)
            _SysVar.DiaBan = Globals.Find_Replace(edt_tendiaban.Text)
            _SysVar.TrucThuoc = 0
            _SysVar.ALL = 0
            _SysVar.Max_Ky = CType(numr_sokyluong.Value, Byte)
            _SysVar.Tinh_Thue_TNCN = IIf(rb_cotamtinh.Checked = True, 1, 0)
            _SysVar.IdCanBo_GiamDoc = IIf(ARL_GiamDoc.Count > 0, ARL_GiamDoc(cb_giamdoc.SelectedIndex), "")
            _SysVar.GiamDoc = edt_giamdoc.Text
            _SysVar.IdCanBo_PhoGD = IIf(ARL_PhoGiamDoc.Count > 0, ARL_PhoGiamDoc(cb_phogiamdoc.SelectedIndex), "")
            _SysVar.PhoGiamDoc = edt_phogiamdoc.Text
            _SysVar.IdCanBo_HCTC = IIf(ARL_HcToChuc.Count > 0, ARL_HcToChuc(cb_hctochuc.SelectedIndex), "")
            _SysVar.TruongHCTC = edt_tphanhchinhtc.Text
            _SysVar.IdCanBo_KeToan = IIf(ARL_KeToan.Count > 0, ARL_KeToan(cb_ketoan.SelectedIndex), "")
            _SysVar.KeToanTruong = edt_ketoan.Text
            _SysVar.Ten_VT = lbl_donvi_tenvt.Text
            _SysVar.DiaBan = edt_tendiaban.Text
            _SysVar.DonVi_Cd = ""
            _SysVar.VungLuongTT = IIf(ARL_LuongTTV.Count > 0, ARL_LuongTTV(cb_vungluongtt.SelectedIndex), "")
            If String.IsNullOrEmpty(edt_muctamung_v2.Text.Replace(",", "").Trim()) Then
                _SysVar.MucTamUng_V2 = 0
            Else
                _SysVar.MucTamUng_V2 = CType(edt_muctamung_v2.Text.Replace(",", "").Trim(), Double)
            End If
            If String.IsNullOrEmpty(edt_hesotamung_v2.Text.Replace(",", "").Trim()) Then
                _SysVar.HeSoTamUng_V2 = 0
            Else
                _SysVar.HeSoTamUng_V2 = CType(edt_hesotamung_v2.Text.Replace(",", "").Trim(), Double)
            End If
            _SysVar.CreatedBy = Globals.UserVal
            _SysVar.ModifiedBy = Globals.UserVal
            _SysVar.TrangThai = 1
            _SysVar.NgayHL = dtpk_ngayapdung.Value
            _SysVar.NgayHetHL = DateTimeUtil.StringToDateTime("31/12/2050", "dd/MM/yyyy")
            _SysVar.Id = _RecordId
            Dim iRetId As Integer
            iRetId = _HeThongBLL.Insert_Update_SysVar(_SysVar)

            Reset_Controls()
            Bind_DataGridView()
            dgv_main.CurrentRow.Selected = False
            dgv_main.Rows(clsHT_DanhMuc.GetRowIndex(_currRow, dgv_main, "cln_Id")).Selected = True
            'Select row trên lưới dữ liệu
            If (dgv_main.Rows.Count > 0) Then
                If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                    SelectRowGrid(_currRow)
                End If
            End If

            'Thực hiện truyền lại các tham số hệ thống cho biến chương trình
            initPublicVar()
        End If
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        Close()
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        Dim ARL_Dels As ArrayList = New ArrayList()
        Dim _count As Integer = 0
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                For i As Integer = 0 To dgv_main.Rows.Count - 1
                    If (dgv_main.Rows(i).Cells("cln_Id").Value IsNot Nothing) Then
                        If (CType(dgv_main.Rows(i).Cells("cln_Choice").Value, Boolean) = True) Then
                            _count += 1
                            ARL_Dels.Add(dgv_main.Rows(i).Cells("cln_Id").Value.ToString())
                        End If
                    End If
                Next
            End If
        End If
        If (_count > 0) Then
            Dim _mess As String = IIf(_count = 1, "", "các ")
            If (MessageBox.Show("Bạn có thực sự muốn xoá " + _mess + "bản ghi Tham số hệ thống đã chọn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) = Windows.Forms.DialogResult.Yes) Then
                For i As Integer = 0 To ARL_Dels.Count - 1
                    _HeThongBLL.Delete_SysVar(CType(ARL_Dels(i).ToString(), Integer), Globals.UserVal, CType("2", Byte))
                Next
                MessageBox.Show("Bạn đã xoá thành công " + _mess + "bản ghi Tham số hệ thống đã chọn!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
                dgv_main_CellClick(sender, Nothing)
            Else
                ARL_Dels.Clear()
                Globals.Check_All_Items(dgv_main, False)
            End If
        Else
            MessageBox.Show("Bạn chưa đánh dấu các bản ghi cần xoá, thao tác bị huỷ bỏ!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
        End If
    End Sub

    Private Sub cb_donvi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_donvi.SelectedIndexChanged
        cb_giamdoc.Items.Clear()
        ARL_GiamDoc.Clear()

        cb_phogiamdoc.Items.Clear()
        ARL_PhoGiamDoc.Clear()

        cb_hctochuc.Items.Clear()
        ARL_HcToChuc.Clear()

        cb_ketoan.Items.Clear()
        ARL_KeToan.Clear()

        lbl_donvi_tenvt.Text = ""

        Dim _TieuDeGiamDoc As String = "", _TieuDePhoGiamDoc As String = "", _TieuDeHCTC As String = "", _TieuDeKeToan As String = "", _PosCodeSelect As String = ""
        Dim _IdGoc As Integer = 0
        Dim sNgayAP As String = dtpk_ngayapdung.Value.ToString("dd/MM/yyyy")
        If (cb_donvi.SelectedIndex > 0 And cb_donvi.Items.Count <> 0) Then
            Dim _DonViId As Int32 = CType(ARL_DonVi(IIf(cb_donvi.SelectedIndex > 0, cb_donvi.SelectedIndex, "0")), Int32)
            If _DonViId > 0 Then
                Dim _ChiNhanhModel As ChiNhanh = New ChiNhanh()
                _ChiNhanhModel = _ChiNhanhBLL.GetChiNhanh_ForId(_DonViId)
                If (Not _ChiNhanhModel Is Nothing) Then
                    lbl_donvi_tenvt.Text = _ChiNhanhModel.Ten_VT
                    _IdGoc = _ChiNhanhModel.Id_Goc
                    _PosCodeSelect = _ChiNhanhModel.Ma_So
                End If
                SoftSqlHelper.GetNumber(String.Format("Select Id_Goc From ChiNhanh Where Id={0}", _DonViId), 0)
                If _IdGoc = 0 Or _IdGoc = 1 Then
                    If DONVI = gMaDonViTW Then
                        _TieuDeGiamDoc = "Tổng giám đốc "
                        _TieuDePhoGiamDoc = "Phó Tổng giám đốc "
                        _TieuDeHCTC = "Giám đốc Ban TCCB "
                        _TieuDeKeToan = "Giám đốc Ban KT-QLTC "
                    Else
                        _TieuDeGiamDoc = "Giám đốc "
                        _TieuDePhoGiamDoc = "Phó giám đốc "
                        _TieuDeHCTC = "TP. Hành chính - Tổ chức "
                        _TieuDeKeToan = "TP. Kế toán - Tài chính "

                    End If
                Else
                    _TieuDeGiamDoc = "Giám đốc "
                    _TieuDePhoGiamDoc = "Phó giám đốc "
                    _TieuDeHCTC = "Tổ trưởng tổ Tín dụng/Tổng hợp "
                    _TieuDeKeToan = "Trưởng kế toán "
                End If
                lbl_giamdoc.Text = _TieuDeGiamDoc
                lbl_phogiamdoc.Text = _TieuDePhoGiamDoc
                lbl_hctc.Text = _TieuDeHCTC
                lbl_ketoan.Text = _TieuDeKeToan

                'Tổng Giám đốc / Giám đốc đơn vị
                strSQL = HeThongBLL.GetQuery_ListCanBos(_DonViId, sNgayAP, "'1402','1403','1406','1410','1430','1411','1431','1428'", "'02','14','20'")
                ARL_GiamDoc = _Globals.Bind_ComBoBox(cb_giamdoc, strSQL, "--- Chọn " + _TieuDeGiamDoc + "---")
                cb_giamdoc_SelectedIndexChanged(sender, Nothing)
                'Phó Tổng Giám đốc / Phó Giám đốc đơn vị
                strSQL = HeThongBLL.GetQuery_ListCanBos(_DonViId, sNgayAP, "'1404','1430','1431','1407','1409','1411','1428'", "'02','14','20'")
                ARL_PhoGiamDoc = _Globals.Bind_ComBoBox(cb_phogiamdoc, strSQL, "--- Chọn " + _TieuDePhoGiamDoc + "---")
                cb_phogiamdoc_SelectedIndexChanged(sender, Nothing)

                'Giám đốc Ban TCCB/Trưởng phòng HC-TC
                If _PosCodeSelect = "001114" Or _PosCodeSelect = "002734" Or _PosCodeSelect = "002821" Or _PosCodeSelect = "003799" Or _PosCodeSelect = "005399" Then
                    strSQL = HeThongBLL.GetQuery_ListCanBos(_DonViId, sNgayAP, "'1411','1412','1406','1430','1432','1414','1450','1439','1444','1426','1415','1443','1428'", "'06','18','22','23','31','26','14','20'")
                Else
                    strSQL = HeThongBLL.GetQuery_ListCanBos(_DonViId, sNgayAP, "'1412','1413','1406','1430','1432','1414','1450','1439','1444','1426','1415','1443','1428'", "'06','18','22','23','31','26'")
                End If
                ARL_HcToChuc = _Globals.Bind_ComBoBox(cb_hctochuc, strSQL, "--- Chọn " + _TieuDeHCTC + "---")
                cb_hctochuc_SelectedIndexChanged(sender, Nothing)

                'Giám đốc Ban KTTC/Trưởng phòng Kế toán
                If _PosCodeSelect = "001114" Or _PosCodeSelect = "002734" Or _PosCodeSelect = "002821" Or _PosCodeSelect = "003799" Or _PosCodeSelect = "005399" Then
                    strSQL = HeThongBLL.GetQuery_ListCanBos(_DonViId, sNgayAP, "'1411','1412','1406','1430','1432','1414','1450','1439','1444','1426','1415','1443','1428'", "'09','15','21','29','34','14','20'")
                Else
                    strSQL = HeThongBLL.GetQuery_ListCanBos(_DonViId, sNgayAP, "'1412','1413','1406','1430','1432','1414','1450','1439','1444','1426','1415','1443','1428'", "'09','15','21','29','34'")
                End If

                ARL_KeToan = _Globals.Bind_ComBoBox(cb_ketoan, strSQL, "--- Chọn " + _TieuDeKeToan + "---")
                cb_ketoan_SelectedIndexChanged(sender, Nothing)
            End If
        End If
    End Sub

    Private Sub cb_giamdoc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_giamdoc.SelectedIndexChanged
        edt_giamdoc.Text = ""
        If (cb_giamdoc.SelectedIndex > 0 And cb_giamdoc.Items.Count <> 0) Then
            Dim sNameSelect As String = cb_giamdoc.SelectedItem
            edt_giamdoc.Text = sNameSelect.Substring(0, sNameSelect.IndexOf(",")).Replace("Ông ", "").Replace("Bà ", "")
        End If
    End Sub

    Private Sub cb_phogiamdoc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_phogiamdoc.SelectedIndexChanged
        edt_phogiamdoc.Text = ""
        If (cb_phogiamdoc.SelectedIndex > 0 And cb_phogiamdoc.Items.Count <> 0) Then
            Dim sNameSelect As String = cb_phogiamdoc.SelectedItem
            edt_phogiamdoc.Text = sNameSelect.Substring(0, sNameSelect.IndexOf(",")).Replace("Ông ", "").Replace("Bà ", "")
        End If
    End Sub

    Private Sub cb_hctochuc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_hctochuc.SelectedIndexChanged
        edt_tphanhchinhtc.Text = ""
        If (cb_hctochuc.SelectedIndex > 0 And cb_hctochuc.Items.Count <> 0) Then
            Dim sNameSelect As String = cb_hctochuc.SelectedItem
            edt_tphanhchinhtc.Text = sNameSelect.Substring(0, sNameSelect.IndexOf(",")).Replace("Ông ", "").Replace("Bà ", "")
        End If
    End Sub

    Private Sub cb_ketoan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_ketoan.SelectedIndexChanged
        edt_ketoan.Text = ""
        If (cb_ketoan.SelectedIndex > 0 And cb_ketoan.Items.Count <> 0) Then
            Dim sNameSelect As String = cb_ketoan.SelectedItem
            edt_ketoan.Text = sNameSelect.Substring(0, sNameSelect.IndexOf(",")).Replace("Ông ", "").Replace("Bà ", "")
        End If
    End Sub

    Private Sub cb_vungluongtt_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb_vungluongtt.SelectedIndexChanged
        edt_sotien_luongvtt.Text = ""
        If (cb_vungluongtt.SelectedIndex > 0 And cb_vungluongtt.Items.Count <> 0) Then
            Dim _VungLuong As String = ARL_LuongTTV(IIf(cb_vungluongtt.SelectedIndex > 0, cb_vungluongtt.SelectedIndex, "0"))
            Dim dLuongTTV As Double = SoftSqlHelper.GetNumberDouble(String.Format("Select Top 1 Cast(IsNull(ThSo_GiaTri,'0') As Float) From HT_ThSoLuong Where ThSo_TenGoi Like '{0}' And ThSo_Loai=1 And ThSo_TrangThai=1 Order By ThSo_NgayHL Desc", "LUONG_TTV_" + _VungLuong), 0)
            edt_sotien_luongvtt.Text = dLuongTTV.ToString()
        Else
            edt_sotien_luongvtt.Text = "0"
        End If
    End Sub

    Private Sub btn_addvar_Click(sender As Object, e As EventArgs) Handles btn_addvar.Click
        Reset_Controls()
        ActiveControl = dtpk_ngayapdung
    End Sub

    Private Sub dgv_main_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_main.CellClick
        Reset_Controls()
        If (dgv_main.Rows.Count > 0) Then
            If ((dgv_main.CurrentRow.Cells("cln_Id").Value IsNot Nothing) And (dgv_main.CurrentRow.Cells("cln_Id").Value.ToString() <> "")) Then
                Dim _SelectId As Integer = CType(dgv_main.CurrentRow.Cells("cln_Id").Value.ToString(), Integer)
                SelectRowGrid(_SelectId)
            End If
        End If
    End Sub

    Private Sub dgv_main_KeyUp(sender As Object, e As KeyEventArgs) Handles dgv_main.KeyUp
        dgv_main_CellClick(sender, Nothing)
    End Sub
#End Region

#Region "---> Events: Các sự kiện ngoại lệ người dùng <---"
    Private Sub edt_sotien_luongvtt_TextChanged(sender As Object, e As EventArgs) Handles edt_sotien_luongvtt.TextChanged
        edt_sotien_luongvtt = formatMoneyinTextbox(edt_sotien_luongvtt)
    End Sub

    Private Sub edt_sotien_luongvtt_Leave(sender As Object, e As EventArgs) Handles edt_sotien_luongvtt.Leave
        If edt_sotien_luongvtt.Text.Trim() = "" Then
            edt_sotien_luongvtt.Text = "0"
        End If
    End Sub

    Private Sub edt_muctamung_v2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_muctamung_v2.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If

        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_hesotamung_v2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles edt_hesotamung_v2.KeyPress
        Dim KeyAscii As Integer
        KeyAscii = Asc(e.KeyChar)
        If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13) Then
            KeyAscii = 0
        End If
        If KeyAscii = 0 Then
            e.Handled = True
        Else
            e.Handled = False
        End If
    End Sub

    Private Sub edt_hesotamung_v2_MouseClick(sender As Object, e As MouseEventArgs) Handles edt_hesotamung_v2.MouseClick
        edt_hesotamung_v2.SelectAll()
    End Sub

    Private Sub edt_muctamung_v2_MouseClick(sender As Object, e As MouseEventArgs) Handles edt_muctamung_v2.MouseClick
        edt_muctamung_v2.SelectAll()
    End Sub
#End Region

End Class