Option Explicit On
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Text
Imports System
Imports System.Reflection

Public Class frmImportExcel_CanBoInf

    Private Sub frmImportExcel_CanBoInf_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmImportExcel_CanBoInf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HideProgressBar()
        If DONVI = gMaDonViTW Then
            bntImportMaCB.Visible = True
        Else
            bntImportMaCB.Visible = False
        End If
        cboLoaiHS.SelectedIndex = 0
    End Sub

    Private Sub HideProgressBar()
        ProgressBar1.Visible = False
    End Sub

    Private Sub ShowProgressBar()
        ProgressBar1.Visible = True
    End Sub

    Private Sub SetProgress(ByVal iVal As Integer)
        ProgressBar1.Value = iVal
        Me.Refresh()
    End Sub

    Private Sub bntImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImport.Click
        Try
            If txtPath.Text.Trim = "" Then
                MsgBox("Phải chọn file trước khi import.")
            Else
                Dim bLoaiHS As Integer = CInt(cboLoaiHS.SelectedIndex)
                ShowProgressBar()
                ProgressBar1.Maximum = 10
                SetProgress(1)
                Me.Refresh()

                initGrid(bLoaiHS, gridDS_CB)
                SetProgress(3)
                ImportExcel(bLoaiHS, txtPath.Text, gridDS_CB)
                'ImportExcel_DaiMa(txtPath.Text)
            End If
            
        Catch ex As Exception

        End Try
    End Sub

    Private Sub initGrid(ByVal sLoaiHS As Integer, ByVal myDataGrid As DataGridView)
        myDataGrid.Columns().Clear()
        myDataGrid.Columns.Add("STT", "STT")
        myDataGrid.Columns.Add("MaCB", "Mã Cán bộ")
        myDataGrid.Columns.Add("HoTen", "Họ tên")
        Select Case sLoaiHS
            Case 0
                myDataGrid.Columns.Add("GioiTinh", "Giới tính")
                myDataGrid.Columns.Add("NgaySinh", "Ngày sinh")
                myDataGrid.Columns.Add("CMT_So", "Số CMT")
                myDataGrid.Columns.Add("ChiNhanh", "Đơn vị")
                myDataGrid.Columns.Add("Phong", "Phòng")

                myDataGrid.Columns("ChiNhanh").Width = 188
                myDataGrid.Columns("Phong").Width = 188
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 1
                'QD nhan su
                myDataGrid.Columns.Add("NgayHL", "Ngày hiệu lực")
                myDataGrid.Columns.Add("So_QD", "Số QĐ")
                myDataGrid.Columns.Add("LoaiQD", "Loại quyết định")
                myDataGrid.Columns.Add("NoiDung", "Nội dung quyết định")

                myDataGrid.Columns("LoaiQD").Width = 100
                myDataGrid.Columns("NoiDung").Width = 208
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 2
                'Quyết định lương
                myDataGrid.Columns.Add("NgayHL", "Ngày hiệu lực")
                myDataGrid.Columns.Add("So_QD", "Số QĐ")
                myDataGrid.Columns.Add("LoaiQD", "Loại quyết định")
                myDataGrid.Columns.Add("NoiDung", "Nội dung quyết định")

                myDataGrid.Columns("LoaiQD").Width = 100
                myDataGrid.Columns("NoiDung").Width = 208
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 3
                'HS_DTVBCC
                myDataGrid.Columns.Add("NamTN", "Năm tốt nghiệp")
                myDataGrid.Columns.Add("TrinhDo", "Trình độ")
                myDataGrid.Columns.Add("ChuyenNganh", "Chuyên ngành đào tạo")
                myDataGrid.Columns.Add("NganhHoc", "Ngành học")

                myDataGrid.Columns("ChuyenNganh").Width = 150
                myDataGrid.Columns("NganhHoc").Width = 215
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 4
                'HS_GDCB
                myDataGrid.Columns.Add("TVien", "Họ tên thành viên GĐCB")
                myDataGrid.Columns.Add("QuanHe", "Quan hệ với CB")
                myDataGrid.Columns.Add("NgheNghiep", "Nghề nghiệp, nơi công tác")
                myDataGrid.Columns.Add("DiaChi", "Địa chỉ nơi ở")

                myDataGrid.Columns("TVien").Width = 150
                myDataGrid.Columns("NgheNghiep").Width = 215
                myDataGrid.Columns("DiaChi").Width = 150
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 5
                'Các văn bản khác: 'NgayKy_QD, So_QD, DVraQD, NoiDung
                myDataGrid.Columns.Add("NgayKy_QD", "Ngày quyết định")
                myDataGrid.Columns.Add("So_QD", "Số QĐ")
                myDataGrid.Columns.Add("DVraQD", "Đơn vị ra quyết định")
                myDataGrid.Columns.Add("NoiDung", "Nội dung quyết định")

                myDataGrid.Columns("DVraQD").Width = 100
                myDataGrid.Columns("NoiDung").Width = 208
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 6
                'Hợp đồng lao động  
                myDataGrid.Columns.Add("DVKyHDLD", "Đơn vị ký HĐLĐ")
                myDataGrid.Columns.Add("NoiLamViec", "Nơi làm việc")
                myDataGrid.Columns.Add("LoaiHD", "Loại HĐLĐ")

                myDataGrid.Columns("DVKyHDLD").Width = 100
                myDataGrid.Columns("NoiLamViec").Width = 208
                myDataGrid.Columns("LoaiHD").Width = 208
                myDataGrid.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
            Case 7
                'Thông tin lực lượng vũ trang
            Case 8
                'HS ki luat
        End Select
        myDataGrid.Columns("STT").Width = 50
        myDataGrid.Columns("MaCB").Width = 100
        myDataGrid.Columns("HoTen").Width = 150
        myDataGrid.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        myDataGrid.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        myDataGrid.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub ImportExcel(ByVal sLoaiHS As Integer, ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)
        Select Case sLoaiHS
            Case 0
                'Thông tin chung của cán bộ
                ImportExcel_CBInf(txtPath.Text, gridDS_CB)
            Case 1
                'QD nhan su
                ImportExcel_QDNhansu(txtPath.Text, gridDS_CB, 3)
                ImportExcel_QDTiepNhan(txtPath.Text, gridDS_CB, 11)
            Case 2
                'Quyết định lương
                ImportExcel_QDLuong(txtPath.Text, gridDS_CB, 4)
            Case 3
                'HS_DTVBCC
                ImportExcel_DaoTao(txtPath.Text, gridDS_CB, 5)
            Case 4
                'HS_GDCB
                ImportExcel_GiaDinhCB(txtPath.Text, gridDS_CB, 6)
            Case 5
                'QD khác
                ImportExcel_QDKhac(txtPath.Text, gridDS_CB, 7)
            Case 6
                'Hợp đồng lao động
                ImportExcel_HDLD(txtPath.Text, gridDS_CB, 8)
            Case 7
                'Thông tin thu nhập gia đình
                MsgBox("Chức năng này đang cập nhật")
            Case 8
                'HS ki luat
                MsgBox("Chức năng này đang cập nhật")
        End Select

    End Sub

    'Đọc nội dung từ file excel vao datagrid
    Private Sub ImportExcel_CBInf(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
        Dim dsCanBo As String = ""
        Dim db As DBAccess = New DBAccess
        Dim strErr As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(1).Select()

                Dim vIdCB As String = ""
                row = 2
                While .cells(row, "A").value.ToString.Trim <> ""
                    If .cells(row, "A").value.ToString.Trim <> "" And .cells(row, "B").value.ToString.Trim <> "" And .cells(row, "C").value.ToString.Trim <> "" Then
                        Dim MaCB_Cu As String = ""
                        's = s & .cells(1, "A").value & ";" & .cells(1, "B").value & ";" & .cells(1, "B").value
                        'Kiem tra neu Mã can bo và CMT neu da co thi ko insert vso nua
                        'If db.getString("SELECT CMT_so FROM HS_canbo WHERE CMT_so='" & .cells(row, "C").value.ToString.Trim & "'") <> "" Then
                        '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").value.ToString.Trim & " đã tồn tại." & vbCrLf
                        '    GoTo nex
                        'End If

                        'If db.getString("SELECT MaCB FROM HS_canbo WHERE MaCB='" & .cells(row, "A").value.ToString.Trim & "'") <> "" Then
                        '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").value.ToString.Trim & " có mã cán bộ đã tồn tại." & vbCrLf
                        '    GoTo nex
                        'End If
                        If Not (.cells(row, "AK").value Is Nothing) Then
                            MaCB_Cu = .cells(row, "AK").value.ToString.Trim
                        End If
                        vIdCB = ""
                        '''''Insert HS can bo'''''
                        Try
                            Dim strCode As String = ""
                            Dim m_HSCanBo As clsHS_CanBo.HS_CanBo = New clsHS_CanBo.HS_CanBo()
                            m_HSCanBo.MaCB = .cells(row, "A").value.ToString.Trim
                            m_HSCanBo.HoTen = standardizeName(.cells(row, "B").value.ToString.Trim)
                            m_HSCanBo.TenThuongGoi = m_HSCanBo.HoTen
                            m_HSCanBo.BiDanh = m_HSCanBo.HoTen
                            If .cells(row, "D").value Is Nothing Then
                                m_HSCanBo.GioiTinh = 0
                            Else
                                m_HSCanBo.GioiTinh = .cells(row, "D").value
                            End If
                            'If Excel.GetType 
                            If .cells(row, "N").value Is Nothing Then
                                m_HSCanBo.NgaySinh = DateTime.Parse("01/01/1900")
                            Else
                                If .cells(row, "N").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "N").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                    m_HSCanBo.NgaySinh = .cells(row, "N").value
                                Else
                                    m_HSCanBo.NgaySinh = DateTimeUtil.getDate(.cells(row, "N").value)
                                End If
                            End If
                            m_HSCanBo.IdDonVi = IdDONVI
                            m_HSCanBo.IdQuocTich = 1
                            If .cells(row, "S").value Is Nothing Then
                                m_HSCanBo.IdDanToc = 0
                            Else
                                If .cells(row, "S").value.ToString.Trim = "" Then
                                    m_HSCanBo.IdDanToc = 0
                                Else
                                    m_HSCanBo.IdDanToc = CInt(.cells(row, "S").value)
                                End If
                            End If
                            If .cells(row, "T").value Is Nothing Then
                                m_HSCanBo.IdTonGiao = 0
                            Else
                                If .cells(row, "T").value.ToString.Trim = "" Then
                                    m_HSCanBo.IdTonGiao = 0
                                Else
                                    m_HSCanBo.IdTonGiao = CInt(.cells(row, "T").value)
                                End If
                            End If

                            m_HSCanBo.CMT_So = .cells(row, "C").value.ToString.Trim
                            If .cells(row, "AI").value Is Nothing Then
                                m_HSCanBo.CMT_NgayCap = DateTime.Parse("01/01/1900")
                            Else
                                If .cells(row, "AI").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AI").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                    m_HSCanBo.CMT_NgayCap = .cells(row, "AI").value
                                Else
                                    m_HSCanBo.CMT_NgayCap = DateTimeUtil.getDate(.cells(row, "AI").value)
                                End If
                            End If
                            If .cells(row, "AJ").value Is Nothing Then
                                m_HSCanBo.CMT_NoiCap = ""
                            Else
                                m_HSCanBo.CMT_NoiCap = .cells(row, "AJ").value.ToString.Trim
                            End If
                            m_HSCanBo.IdNS_Tinh = 0
                            m_HSCanBo.IdNS_Huyen = 0
                            If .cells(row, "O").value Is Nothing Then
                                m_HSCanBo.NS_DChi = ""
                            Else
                                m_HSCanBo.NS_DChi = .cells(row, "O").value.ToString.Trim
                            End If

                            m_HSCanBo.IdNQ_Tinh = 0
                            m_HSCanBo.IdNQ_Huyen = 0
                            If .cells(row, "P").value Is Nothing Then
                                m_HSCanBo.NQ_DChi = ""
                            Else
                                m_HSCanBo.NQ_DChi = .cells(row, "P").value.ToString.Trim
                                If m_HSCanBo.NS_DChi = "" Then m_HSCanBo.NS_DChi = m_HSCanBo.NQ_DChi
                            End If

                            m_HSCanBo.IdThT_Tinh = 0
                            m_HSCanBo.IdThT_Huyen = 0
                            If .cells(row, "Q").value Is Nothing Then
                                m_HSCanBo.ThT_Diachi = ""
                            Else
                                m_HSCanBo.ThT_Diachi = .cells(row, "Q").value.ToString.Trim
                            End If
                            If .cells(row, "R").value Is Nothing Then
                                m_HSCanBo.ThT_Dienthoai = ""
                            Else
                                m_HSCanBo.ThT_Dienthoai = .cells(row, "R").value.ToString.Trim
                            End If
                            m_HSCanBo.IdTTr_Huyen = 0
                            m_HSCanBo.TTr_Diachi = m_HSCanBo.ThT_Diachi
                            m_HSCanBo.TTr_Dienthoai = ""
                            m_HSCanBo.DienThoai_CQ = ""
                            m_HSCanBo.DienThoai_DD = ""
                            m_HSCanBo.DienThoai_NR = ""
                            m_HSCanBo.SoFax = ""
                            m_HSCanBo.Email = ""
                            m_HSCanBo.NhomMau = ""
                            m_HSCanBo.IdUT_GDinh = 0

                            If .cells(row, "U").value Is Nothing Then
                                m_HSCanBo.IdUT_BThan = ""
                            Else
                                m_HSCanBo.IdUT_BThan = .cells(row, "U").value.ToString.Trim & ";"
                            End If
                            If .cells(row, "V").value Is Nothing Then
                                m_HSCanBo.IdThanhPhanGD = 0
                            Else
                                If .cells(row, "V").value.ToString.Trim = "" Then
                                    m_HSCanBo.IdThanhPhanGD = 0
                                Else
                                    m_HSCanBo.IdThanhPhanGD = CInt(.cells(row, "V").value)
                                End If
                            End If
                            If .cells(row, "W").value Is Nothing Then
                                m_HSCanBo.DacDiem_BT = ""
                            Else
                                m_HSCanBo.DacDiem_BT = .cells(row, "W").value.ToString.Trim
                            End If
                            If .cells(row, "X").value Is Nothing Then
                                m_HSCanBo.QuanHe_Nguoi_NN = ""
                            Else
                                m_HSCanBo.QuanHe_Nguoi_NN = .cells(row, "X").value.ToString.Trim
                            End If

                            If .cells(row, "Z").value Is Nothing Then
                                m_HSCanBo.IdTrinhDoVH = 0
                            Else
                                If .cells(row, "Z").value.ToString.Trim = "" Then
                                    m_HSCanBo.IdTrinhDoVH = 0
                                Else
                                    m_HSCanBo.IdTrinhDoVH = CInt(.cells(row, "Z").value)
                                End If
                            End If
                            If .cells(row, "AC").value Is Nothing Then
                                m_HSCanBo.IdTrinhDoCT = 0
                            Else
                                If .cells(row, "AC").value.ToString.Trim = "" Then
                                    m_HSCanBo.IdTrinhDoCT = 0
                                Else
                                    m_HSCanBo.IdTrinhDoCT = CInt(.cells(row, "AC").value)
                                End If
                            End If

                            If .cells(row, "AG").value Is Nothing Then
                                m_HSCanBo.Ngay_NH = DateTime.Parse("01/01/1900")
                            Else
                                If .cells(row, "AG").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AG").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                    m_HSCanBo.Ngay_NH = .cells(row, "AG").value
                                Else
                                    m_HSCanBo.Ngay_NH = DateTimeUtil.getDate(.cells(row, "AG").value)
                                End If
                            End If
                            If .cells(row, "AH").value Is Nothing Then
                                m_HSCanBo.Ngay_ThamNien = DateTime.Parse("01/01/1900")
                                m_HSCanBo.Ngay_VBSP = DateTime.Parse("01/01/1900")
                                m_HSCanBo.Ngay_CQ = DateTime.Parse("01/01/1900")
                            Else
                                If .cells(row, "AH").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AH").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                    m_HSCanBo.Ngay_VBSP = .cells(row, "AH").value
                                Else
                                    m_HSCanBo.Ngay_VBSP = DateTimeUtil.getDate(.cells(row, "AH").value)
                                End If
                                m_HSCanBo.Ngay_ThamNien = m_HSCanBo.Ngay_VBSP
                                m_HSCanBo.Ngay_CQ = m_HSCanBo.Ngay_VBSP
                            End If
                            m_HSCanBo.Ngay_BienChe = DateTimeUtil.getDate(DateTime.Parse("01/01/1900"))
                            If .cells(row, "Y").value Is Nothing Then
                                m_HSCanBo.CM_Ngay = DateTime.Parse("01/01/1900")
                            Else
                                If .cells(row, "Y").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "Y").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                    m_HSCanBo.CM_Ngay = .cells(row, "Y").value
                                Else
                                    m_HSCanBo.CM_Ngay = DateTimeUtil.getDate(.cells(row, "Y").value)
                                End If
                            End If
                            m_HSCanBo.CM_ToChuc = ""
                            m_HSCanBo.SoTruong_CT = ""
                            m_HSCanBo.CV_Lau = ""
                            m_HSCanBo.NH_MaKH = ""
                            m_HSCanBo.NH_SoTK = ""
                            m_HSCanBo.NH_TenNH = ""
                            If .cells(row, "AD").value Is Nothing Then
                                m_HSCanBo.MaSoThue = 0
                            Else
                                m_HSCanBo.MaSoThue = .cells(row, "AD").value.ToString.Trim
                            End If
                            If .cells(row, "AE").value Is Nothing Then
                                m_HSCanBo.BHXH_SoSo = 0
                            Else
                                m_HSCanBo.BHXH_SoSo = .cells(row, "AE").value.ToString.Trim
                            End If

                            m_HSCanBo.BHXH_NgaySo = DateTime.Parse("01/01/1900")
                            m_HSCanBo.BHXH_NgayBatDau = DateTime.Parse("01/01/1900")
                            If .cells(row, "AF").value Is Nothing Then
                                m_HSCanBo.BHXH_NoiLam = ""
                            Else
                                m_HSCanBo.BHXH_NoiLam = .cells(row, "AF").value.ToString.Trim
                            End If

                            m_HSCanBo.GhiChu = ""
                            m_HSCanBo.IdNew = ""
                            m_HSCanBo.AnhThe = Nothing

                            If MaCB_Cu <> "" Then
                                Try
                                    Dim IDCB_Cu As String = ""
                                    IDCB_Cu = getCanBo_ID(MaCB_Cu, False)
                                    m_HSCanBo.IdCanBo = IDCB_Cu
                                    If IDCB_Cu = "" Then
                                        strCode = _HS_CanBo.Insert_Human(m_HSCanBo)
                                    Else
                                        _HS_CanBo.Update_HumanNotFull(m_HSCanBo)
                                    End If
                                    vIdCB = m_HSCanBo.IdCanBo
                                Catch ex As Exception
                                End Try
                            Else
                                strCode = _HS_CanBo.Insert_Human(m_HSCanBo)
                                vIdCB = m_HSCanBo.IdCanBo
                                If vIdCB = "" Then
                                    strErr = strErr & m_HSCanBo.HoTen & vbCr
                                End If
                            End If

                            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu

                            ' Neu insert vao HSCanbo Ok thi insert tiep vao HS QD Nhan su
                            If vIdCB <> "" Then

                                'If .cells(row, "E").value Is Nothing Then
                                '    'strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").value.ToString.Trim & " nhập sai Ngày vào đơn vị hiện đang công tác." & vbCrLf
                                '    '_HS_CanBo.Delete_Human(vIdCB)
                                '    'GoTo nex
                                '    .cells(row, "E").value = Now.Date
                                'Else

                                'End If
                                'If .cells(row, "F").value Is Nothing Then
                                '    'strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " nhập sai Loại quyết định nhân sự. (ID Loại quyết định = 321/322/323/657/658)." & vbCrLf
                                '    '_HS_CanBo.Delete_Human(vIdCB)
                                '    'GoTo nex
                                '    .cells(row, "F").value = 321
                                'End If
                                'If .cells(row, "G").value Is Nothing Then
                                '    'strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chưa nhập số Quyết định vào đơn vị hiện tại." & vbCrLf
                                '    '_HS_CanBo.Delete_Human(vIdCB)
                                '    'GoTo nex
                                '    .cells(row, "G").value = "ChuaCoSo"
                                'End If
                                If .cells(row, "K").value Is Nothing Or .cells(row, "L").value Is Nothing Or .cells(row, "M").value Is Nothing Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chưa nhập đầy đủ ID Phòng(tổ)/Chức vụ/Chuyên môn." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                ' Kiem tra Loai quyet dinh co thuoc danh sach QD khong
                                'If Not (CInt(.cells(row, "F").value) = 321 Or CInt(.cells(row, "F").value) = 322 Or CInt(.cells(row, "F").value) = 323 Or CInt(.cells(row, "F").value) = 657 Or CInt(.cells(row, "F").value) = 658) Then
                                '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai Loại quyết định nhân sự. (ID Loại quyết định = 321/322/323/657/658)" & vbCrLf
                                '    _HS_CanBo.Delete_Human(vIdCB)
                                '    GoTo nex
                                'End If
                                'If db.getNumber("SELECT [ID] FROM Chinhanh WHERE (id=" & IdDONVI & " or id_goc=" & IdDONVI & ") AND id=" & CInt(.cells(row, "J").value) & " AND status=1") <= 0 Then
                                '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Hội sở tỉnh hoặc ID của PGD hiện đang công tác." & vbCrLf
                                '    _HS_CanBo.Delete_Human(vIdCB)
                                '    GoTo nex
                                'End If
                                'If Not ((IdDONVI >= 6 And 15 <= CInt(.cells(row, "K").value) And CInt(.cells(row, "K").value) <= 24) Or (IdDONVI <= 5) And ((1 <= CInt(.cells(row, "K").value) And CInt(.cells(row, "K").value) <= 14) Or (25 <= CInt(.cells(row, "K").value) And CInt(.cells(row, "K").value) <= 37))) Then
                                '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Phòng hoặc tổ hiện đang công tác." & vbCrLf
                                '    _HS_CanBo.Delete_Human(vIdCB)
                                '    GoTo nex
                                'End If
                                'If Not ((304 <= CInt(.cells(row, "L").value) And CInt(.cells(row, "L").value) <= 309) Or CInt(.cells(row, "L").value) = 320 Or CInt(.cells(row, "L").value) = 706) Then
                                '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Chức vụ hiện đang công tác." & vbCrLf
                                '    _HS_CanBo.Delete_Human(vIdCB)
                                '    GoTo nex
                                'End If
                                'If Not ((240 <= CInt(.cells(row, "M").value) And CInt(.cells(row, "M").value) <= 263) Or CInt(.cells(row, "M").value) = 271 Or CInt(.cells(row, "M").value) = 282) Then
                                '    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Chuyên môn hiện tại." & vbCrLf
                                '    _HS_CanBo.Delete_Human(vIdCB)
                                '    GoTo nex
                                'End If

                                m_QDNhanSu.IdQDNhanSu = ""
                                m_QDNhanSu.IdCanBo = vIdCB
                                If .cells(row, "G").value Is Nothing Then
                                    m_QDNhanSu.So_QD = "ChuaCoSo"
                                Else
                                    m_QDNhanSu.So_QD = .cells(row, "G").value.ToString.Trim
                                End If

                                If .cells(row, "E").value Is Nothing Then
                                    m_QDNhanSu.NgayKy_QD = Now.Date
                                Else
                                    If .cells(row, "E").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "E").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                        m_QDNhanSu.NgayKy_QD = .cells(row, "E").value
                                    Else
                                        m_QDNhanSu.NgayKy_QD = DateTimeUtil.getDate(.cells(row, "E").value)
                                    End If
                                End If

                                If .cells(row, "H").value Is Nothing Then
                                    m_QDNhanSu.NguoiKy_QD = ""
                                Else
                                    m_QDNhanSu.NguoiKy_QD = standardizeName(.cells(row, "H").value.ToString.Trim)
                                End If
                                If .cells(row, "F").value Is Nothing Then
                                    m_QDNhanSu.IdLoaiQD = 321
                                Else
                                    If .cells(row, "F").value.ToString.Trim = "" Then
                                        m_QDNhanSu.IdLoaiQD = 321
                                    Else
                                        m_QDNhanSu.IdLoaiQD = CInt(.cells(row, "F").value)
                                    End If
                                End If

                                m_QDNhanSu.LoaiQD = ""
                                m_QDNhanSu.NgayHL = m_QDNhanSu.NgayKy_QD
                                If .cells(row, "I").value Is Nothing Then
                                    m_QDNhanSu.idCV_Nguoiky_QD = 296
                                Else
                                    If .cells(row, "I").value.ToString.Trim = "" Then
                                        m_QDNhanSu.idCV_Nguoiky_QD = 296
                                    Else
                                        m_QDNhanSu.idCV_Nguoiky_QD = .cells(row, "I").value
                                    End If
                                End If
                                m_QDNhanSu.CV_NguoiKy_QD = ""
                                m_QDNhanSu.NgayBoNhiem_TT = DateTime.Parse("01/01/1900")
                                m_QDNhanSu.NgayThoiLuong = DateTime.Parse("01/01/1900")
                                m_QDNhanSu.IdDonvi_Cu = 0
                                m_QDNhanSu.IdPhong_Cu = 0
                                m_QDNhanSu.IdChucvu_Cu = 0
                                m_QDNhanSu.IdChuyenMon_Cu = 0
                                m_QDNhanSu.IdDonvi_Moi = CInt(.cells(row, "J").value)
                                m_QDNhanSu.IdPhong_Moi = CInt(.cells(row, "K").value)
                                m_QDNhanSu.IdChucvu_Moi = CInt(.cells(row, "L").value)
                                m_QDNhanSu.IdChuyenMon_Moi = CInt(.cells(row, "M").value)
                                m_QDNhanSu.IsQD_NHCS = 1
                                m_QDNhanSu.DVraQD = getDonvi(m_QDNhanSu.IdDonvi_Moi)
                                m_QDNhanSu.DenNgay = DateTime.Parse("01/01/1900")
                                m_QDNhanSu.NoiDung = ""
                                m_QDNhanSu.Active = 1
                                m_QDNhanSu.GhiChu = ""
                                Dim vID_QDNhansuCB As String
                                vID_QDNhansuCB = db.getString("SELECT TOP 1 IdQDNhanSu from qdnhansu where idcanbo='" & vIdCB & "' and IsQD_NHCS=1 order by NgayHL desc")
                                If vID_QDNhansuCB = "" Then
                                    m_QDNhanSu.Add()
                                Else
                                    m_QDNhanSu.IdQDNhanSu = vID_QDNhansuCB
                                    m_QDNhanSu.UpdateNotFull()
                                End If

                            End If
                            ' Insert tiep vao HS Dang

                            ' OK
                            If m_QDNhanSu.IdQDNhanSu <> "" Then
                                countCB = countCB + 1
                                If dsCanBo = "" Then
                                    dsCanBo = "'" & vIdCB & "'"
                                Else
                                    dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                End If
                            Else
                                _HS_CanBo.Delete_Human(vIdCB)
                            End If

                        Catch ex As Exception
                            _HS_CanBo.Delete_Human(vIdCB)
                        End Try
                    End If
nex:
                    row = row + 1
                End While
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then

            Dim arrDel() As String
            arrDel = dsCanBo.Split(",")

            Dim sqlDSCanBo As String = "SELECT * FROM HS_CanBo WHERE IdCanbo in (" & dsCanBo & ")"
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = db.SelectDBRows(sqlDSCanBo)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                myDataGrid.Rows(j).Cells("GioiTinh").Value = IIf(dt.Rows(j)("GioiTinh").ToString() = False, "Nam", "Nữ")
                If dt.Rows(j)("NgaySinh").ToString().Trim() <> "" Then
                    myDataGrid.Rows(j).Cells("NgaySinh").Value = CType(dt.Rows(j)("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy")
                Else
                    myDataGrid.Rows(j).Cells("NgaySinh").Value = ""
                End If
                myDataGrid.Rows(j).Cells("CMT_So").Value = dt.Rows(j)("CMT_So").ToString()

                Dim dt1 As DataTable
                dt1 = db.SelectDBRows("SELECT TOP 1 (SELECT ten_goi FROM ChiNhanh where Id=IdDonvi_Moi) as Chinhanh, (SELECT ten_phong FROM PhongBan WHERE Id=IdPhong_Moi) as Phong FROM QDNhansu WHERE IdCanBo='" & dt.Rows(j)("IdCanBo").ToString() & "' order by NgayHL desc")
                myDataGrid.Rows(j).Cells("Chinhanh").Value = dt1.Rows(0)("Chinhanh").ToString()
                myDataGrid.Rows(j).Cells("Phong").Value = dt1.Rows(0)("Phong").ToString()
            Next
            labDS.Text = "Danh sách cán bộ được import vào CSDL: " & countCB.ToString & " cán bộ."
        Else
            labDS.Text = "Danh sách cán bộ được import vào CSDL: 0 cán bộ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách cán bộ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_QDNhansu(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "QDNHANSU" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "E").value Is Nothing) Then
                                If .cells(row, "E").value.ToString.Trim <> "" Then
                                    Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                                    Dim m_LuongOld As LuongCanBo = New LuongCanBo
                                    Dim m_PhucapOld As Phucap = New Phucap
                                    Dim vIdQDNhanSu As String = ""
                                    Dim vIdLuong As String = ""
                                    Dim vIdPhucap As String = ""
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then
                                        'If row = 94 Then
                                        '    MsgBox("debug")
                                        'End If
                                        m_QDNhanSu.IdQDNhanSu = ""
                                        m_QDNhanSu.IdCanBo = vIdCB
                                        If .cells(row, "D").value Is Nothing Then
                                            m_QDNhanSu.IsQD_NHCS = 0
                                        Else
                                            m_QDNhanSu.IsQD_NHCS = CInt(.cells(row, "D").value)
                                        End If

                                        If .cells(row, "E").value Is Nothing Then
                                            m_QDNhanSu.So_QD = "ChuaCoSo"
                                        Else
                                            m_QDNhanSu.So_QD = .cells(row, "E").value.ToString.Trim
                                        End If

                                        'kiem tra mot so dieu kien
                                        If .cells(row, "I").value Is Nothing Then
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định số " & m_QDNhanSu.So_QD & " Chưa nhập ID loại quyết định nhân sự." & vbCrLf
                                            GoTo nex
                                        End If
                                        If .cells(row, "I").value.ToString.Trim <> "" Then
                                            If m_QDNhanSu.IsQD_NHCS Then
                                                If Not (.cells(row, "L").value Is Nothing) And Not (.cells(row, "N").value Is Nothing) And Not (.cells(row, "P").value Is Nothing) And Not (.cells(row, "R").value Is Nothing) Then
                                                    Dim _value As Integer = 0
                                                    _value = CInt(.cells(row, "I").value.ToString.Trim)
                                                    If Not ((_value >= 321 And _value <= 327) Or _value = 657 Or _value = 658 Or _value = 688 Or _value = 711 Or (_value >= 714 And _value <= 717) Or _value = 748 Or _value = 768 Or (_value >= 770 And _value <= 774) Or _value = 808 Or _value = 809 Or _value = 812 Or _value = 813) Then
                                                        strErr = strErr & "Giá trị ID Loại quyết định nhân sự không đúng. Hãy kiểm tra lại."
                                                        GoTo nex
                                                    End If
                                                    _value = CInt(.cells(row, "L").value.ToString.Trim)
                                                    If Not ((295 <= _value And _value <= 320) Or _value = 691 Or _value = 706 Or (738 <= _value And _value <= 747) Or _value <= 765) Then
                                                        strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_QDNhanSu.So_QD & " nhập sai ID Chức vụ." & vbCrLf
                                                        GoTo nex
                                                    End If
                                                    _value = CInt(.cells(row, "N").value.ToString.Trim)
                                                    If Not ((240 <= _value And _value <= 262) Or _value = 266 Or _value = 268 Or _value = 269 Or (272 <= _value And _value <= 282) Or (762 <= _value And _value <= 764)) Then
                                                        strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_QDNhanSu.So_QD & " nhập sai ID của Chuyên môn." & vbCrLf
                                                        GoTo nex
                                                    End If
                                                Else
                                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_QDNhanSu.So_QD & " chưa nhập đủ các thông tin cần thiết của Quyết định nhân sự." & vbCrLf
                                                    GoTo nex
                                                End If
                                            Else
                                                If Not (.cells(row, "H").value Is Nothing) And Not (.cells(row, "J").value Is Nothing) And Not (.cells(row, "AH").value Is Nothing) Then
                                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc không thuộc NHCS số " & m_QDNhanSu.So_QD & " chưa nhập đủ các thông tin cần thiết của Quyết định nhân sự." & vbCrLf
                                                    GoTo nex
                                                End If
                                            End If
                                        Else
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Loại Quyết định nhân sự." & vbCrLf
                                            GoTo nex
                                        End If
                                        ' tiep tuc thuc hien
                                        If .cells(row, "F").value Is Nothing Then
                                            m_QDNhanSu.NgayKy_QD = Now.Date
                                        Else
                                            If .cells(row, "F").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "F").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_QDNhanSu.NgayKy_QD = .cells(row, "F").value
                                            Else
                                                m_QDNhanSu.NgayKy_QD = DateTimeUtil.getDate(.cells(row, "F").value)
                                            End If
                                        End If

                                        If .cells(row, "AC").value Is Nothing Then
                                            m_QDNhanSu.NguoiKy_QD = ""
                                        Else
                                            m_QDNhanSu.NguoiKy_QD = standardizeName(.cells(row, "AC").value.ToString.Trim)
                                        End If

                                        If .cells(row, "J").value Is Nothing Then
                                            m_QDNhanSu.NgayHL = m_QDNhanSu.NgayKy_QD
                                        Else
                                            If .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_QDNhanSu.NgayHL = .cells(row, "J").value
                                            Else
                                                m_QDNhanSu.NgayHL = DateTimeUtil.getDate(.cells(row, "J").value)
                                            End If
                                        End If
                                        m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
                                        m_QDNhanSu.NgayThoiLuong = DateTime.MinValue

                                        If .cells(row, "G").value Is Nothing Then
                                            m_QDNhanSu.DVraQD = ""
                                        Else
                                            m_QDNhanSu.DVraQD = standardizeString(.cells(row, "G").value.ToString.Trim)
                                        End If

                                        If .cells(row, "I").value Is Nothing Then
                                            m_QDNhanSu.IdLoaiQD = 321
                                        Else
                                            If .cells(row, "I").value.ToString.Trim = "" Then
                                                m_QDNhanSu.IdLoaiQD = 321
                                            Else
                                                m_QDNhanSu.IdLoaiQD = CInt(.cells(row, "I").value)
                                            End If
                                        End If

                                        If m_QDNhanSu.IsQD_NHCS Then
                                            Dim m_QDNSfinal As QDNhanSu = New QDNhanSu
                                            m_QDNSfinal = m_QDNSfinal.getFinalRecord(vIdCB, m_QDNhanSu.NgayHL)
                                            m_QDNhanSu.IdDonvi_Cu = m_QDNSfinal.IdDonvi_Moi
                                            m_QDNhanSu.IdPhong_Cu = m_QDNSfinal.IdPhong_Moi
                                            m_QDNhanSu.IdChucvu_Cu = m_QDNSfinal.IdChucvu_Moi
                                            m_QDNhanSu.IdChuyenMon_Cu = m_QDNSfinal.IdChuyenMon_Moi
                                            m_QDNhanSu.IdDonvi_Moi = CInt(.cells(row, "R").value)
                                            m_QDNhanSu.IdPhong_Moi = CInt(.cells(row, "P").value)
                                            m_QDNhanSu.IdChucvu_Moi = CInt(.cells(row, "L").value)
                                            m_QDNhanSu.IdChuyenMon_Moi = CInt(.cells(row, "N").value)
                                            If .cells(row, "AE").value Is Nothing Then
                                                m_QDNhanSu.idCV_Nguoiky_QD = 296
                                            Else
                                                If .cells(row, "AE").value.ToString.Trim = "" Then
                                                    m_QDNhanSu.idCV_Nguoiky_QD = 296
                                                Else
                                                    m_QDNhanSu.idCV_Nguoiky_QD = CInt(.cells(row, "AE").value)
                                                End If
                                            End If
                                            m_QDNhanSu.CV_NguoiKy_QD = ""
                                            m_QDNhanSu.LoaiQD = ""
                                            m_QDNhanSu.DenNgay = DateTime.MinValue
                                            m_QDNhanSu.NoiDung = ""
                                        Else
                                            m_QDNhanSu.IdDonvi_Cu = 0
                                            m_QDNhanSu.IdPhong_Cu = 0
                                            m_QDNhanSu.IdChucvu_Cu = 0
                                            m_QDNhanSu.IdChuyenMon_Cu = 0
                                            m_QDNhanSu.IdDonvi_Moi = 0
                                            m_QDNhanSu.IdPhong_Moi = 0
                                            m_QDNhanSu.IdChucvu_Moi = 0
                                            If .cells(row, "AG").value Is Nothing Then
                                                m_QDNhanSu.DenNgay = m_QDNhanSu.NgayKy_QD
                                            Else
                                                If .cells(row, "AG").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AG").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                    m_QDNhanSu.DenNgay = .cells(row, "AG").value
                                                Else
                                                    m_QDNhanSu.DenNgay = DateTimeUtil.getDate(.cells(row, "AG").value)
                                                End If
                                            End If
                                            If .cells(row, "AD").value Is Nothing Then
                                                m_QDNhanSu.CV_NguoiKy_QD = ""
                                            Else
                                                m_QDNhanSu.CV_NguoiKy_QD = standardizeString(.cells(row, "AD").value.ToString.Trim)
                                            End If
                                            m_QDNhanSu.idCV_Nguoiky_QD = 0
                                            If .cells(row, "H").value Is Nothing Then
                                                m_QDNhanSu.LoaiQD = ""
                                            Else
                                                m_QDNhanSu.LoaiQD = standardizeString(.cells(row, "H").value.ToString.Trim)
                                            End If
                                            If .cells(row, "AH").value Is Nothing Then
                                                If .cells(row, "K").value Is Nothing And .cells(row, "Q").value Is Nothing Then
                                                    m_QDNhanSu.NoiDung = ""
                                                Else
                                                    If .cells(row, "K").value Is Nothing Then
                                                        m_QDNhanSu.NoiDung = ""
                                                    Else
                                                        m_QDNhanSu.NoiDung = .cells(row, "K").value.ToString.Trim
                                                    End If
                                                    If Not .cells(row, "Q").value Is Nothing Then
                                                        m_QDNhanSu.NoiDung = m_QDNhanSu.NoiDung & .cells(row, "Q").value.ToString.Trim
                                                    End If
                                                End If
                                            Else
                                                m_QDNhanSu.NoiDung = standardizeString(.cells(row, "AH").value.ToString.Trim)
                                            End If
                                        End If

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

                                        m_QDNhanSu.Active = 1
                                        If .cells(row, "AF").value Is Nothing Then
                                            m_QDNhanSu.GhiChu = ""
                                        Else
                                            m_QDNhanSu.GhiChu = standardizeString(.cells(row, "AF").value.ToString.Trim)
                                        End If

                                        vIdLuong = m_LuongOld.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)
                                        vIdPhucap = m_PhucapOld.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)

                                        'If row = 128 Then
                                        '    MsgBox("Debug")
                                        'End If
                                        '
                                        'If Not checkQuyetDinh("QDNHANSU", vIdCB, m_QDNhanSu.So_QD) Then
                                        vIdQDNhanSu = dbconn.getString("SELECT IdQDNhanSu FROM QDNhanSu WHERE IdCanBo = '" & vIdCB & "' and Upper(So_QD)=Upper(N'" & m_QDNhanSu.So_QD & "')")
                                        If vIdQDNhanSu = "" Then
                                            vIdQDNhanSu = m_QDNhanSu.Add()
                                            'Else
                                            '    vIdQDNhanSu = m_QDNhanSu.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)
                                        End If

                                        If vIdQDNhanSu <> "" Then
                                            If m_QDNhanSu.IsKiemNhiem = 2 Then
                                                'cap nhat lai QD kiem nhiem truoc do
                                                dbconn.executeSQL("UPDATE QDNhanSu SET IsKiemNhiem=2 WHERE IdCanBo='" & vIdCB & "' and IsKiemNhiem=1 and IdDonvi_Moi=" & m_QDNhanSu.IdDonvi_Moi & " and IdPhong_Moi=" & m_QDNhanSu.IdPhong_Moi & " and IdChucVu_Moi=" & m_QDNhanSu.IdChucvu_Moi)
                                            End If
                                            Dim m_Luong As LuongCanBo = New LuongCanBo
                                            Dim m_PhuCap As Phucap = New Phucap
                                            Dim maQDNhanSu As String = ""
                                            Dim IDLoaiQDLuong As String = ""
                                            Dim db As DBAccess = New DBAccess
                                            m_Luong.IdLuongCB = vIdLuong
                                            m_Luong.IdCanBo = vIdCB

                                            If Not (.cells(row, "V").value Is Nothing) Then
                                                If Not (.cells(row, "V").value.ToString.Trim = "" Or .cells(row, "V").value.ToString.Trim = "0") Then
                                                    m_Luong.Ngay_Huong = DateTimeUtil.getDate(m_QDNhanSu.NgayHL)
                                                    m_Luong.DVraQD = m_QDNhanSu.DVraQD
                                                    If m_QDNhanSu.IsQD_NHCS Then
                                                        'QD lương kem theo QD nhan su thuoc NHCS
                                                        If .cells(row, "AA").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Bậc lương." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "AA").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Bậc lương." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        Dim dt As DataTable = New System.Data.DataTable
                                                        Dim IdNgach As Integer = 0
                                                        Dim Heso As Double = 0
                                                        dt = dbconn.SelectDBRows("SELECT IdNgachLuong, Heso FROM V$_BacLuong WHERE Idbacluong=" & CInt(.cells(row, "AA").value))
                                                        If dt.Rows.Count > 0 Then
                                                            IdNgach = dt.Rows(0).Item("IdNgachLuong")
                                                            Heso = dt.Rows(0).Item("Heso")
                                                        End If
                                                        m_Luong.NgayLen_DK = m_QDNhanSu.NgayHL.AddMonths(getTimeNangBac(IdNgach))
                                                        m_Luong.IsQD_NHCS = 1
                                                        m_Luong.NoiDung = ""
                                                        m_Luong.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                                                        m_Luong.CV_NguoiKy_QD = ""
                                                        m_Luong.IdBacLuong = CInt(.cells(row, "AA").value)
                                                        m_Luong.HeSoLuong = Heso
                                                        m_Luong.LoaiQD = ""
                                                    Else
                                                        'QD lương kem theo QD nhan su không thuoc NHCS
                                                        If .cells(row, "AB").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập thông tin về lương trong QĐ nhân sự." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "AB").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập thông tin về lương trong QĐ nhân sự." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        m_Luong.NgayLen_DK = Date.MinValue
                                                        m_Luong.IdCV_Nguoi_QD = 0
                                                        m_Luong.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                                                        m_Luong.NoiDung = .cells(row, "AB").value.ToString.Trim
                                                        m_Luong.IsQD_NHCS = 0
                                                        m_Luong.LoaiQD = m_QDNhanSu.LoaiQD
                                                    End If
                                                    maQDNhanSu = getDanhmuc_MaSo(m_QDNhanSu.IdLoaiQD)
                                                    IDLoaiQDLuong = db.getNumber("SELECT ID FROM DanhMuc WHERE ma_so ='43" & (CInt(maQDNhanSu.Substring(2, 1)) + 1) & maQDNhanSu.Substring(3, 1) & "' and ma_so in ('4311','4312','4313','4314','4315','4317','4318','4319','4327','4328','4329','4330','4331','4332','4323','4324','4325')")
                                                    If IDLoaiQDLuong = 0 Then
                                                        IDLoaiQDLuong = getDanhmuc_ID("4323")
                                                    End If
                                                    m_Luong.IdLoaiQD = IDLoaiQDLuong
                                                    m_Luong.SoQD = m_QDNhanSu.So_QD
                                                    m_Luong.NgayQD = m_QDNhanSu.NgayKy_QD
                                                    m_Luong.NguoiQD = m_QDNhanSu.NguoiKy_QD
                                                    m_Luong.GhiChu = m_QDNhanSu.GhiChu
                                                    If vIdLuong <> "" Then
                                                        m_Luong.Update()
                                                    Else
                                                        ' Kiem tra thông tin lương trong QDNS neu trung với thông tin lương hiện đang hưởng của CB thì ko cho nhập sang HS Luong
                                                        Dim m_LuongFinal As LuongCanBo = New LuongCanBo
                                                        m_LuongFinal = m_LuongFinal.getFinalRecord(vIdCB)
                                                        If Not (m_LuongFinal.IdBacLuong = m_Luong.IdBacLuong And m_Luong.IsQD_NHCS) Then
                                                            m_Luong.Add()
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If vIdLuong <> "" Then
                                                    m_Luong.Delete()
                                                End If
                                            End If

                                            m_PhuCap.IdCB_PhuCap = vIdPhucap
                                            m_PhuCap.IdCanBo = vIdCB
                                            If Not (.cells(row, "S").value Is Nothing) Then
                                                If Not (.cells(row, "S").value.ToString.Trim = "" Or .cells(row, "S").value.ToString.Trim = "0") Then
                                                    If m_QDNhanSu.IsQD_NHCS Then
                                                        If .cells(row, "T").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Mức phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "T").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Mức phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        m_PhuCap.IsQD_NHCS = 1
                                                        m_PhuCap.NoiDung = ""
                                                        m_PhuCap.IdMucPC = CInt(.cells(row, "T").value)
                                                        m_PhuCap.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                                                        m_PhuCap.CV_NguoiKy_QD = ""
                                                    Else
                                                        If .cells(row, "U").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập Nội dung hưởng phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "U").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập Nội dung hưởng phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        m_PhuCap.IdCV_Nguoi_QD = 0
                                                        m_PhuCap.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                                                        m_PhuCap.IsQD_NHCS = 0
                                                        m_PhuCap.NoiDung = .cells(row, "U").value.ToString.Trim
                                                    End If
                                                    m_PhuCap.TuNgay = DateTimeUtil.getDate(m_QDNhanSu.NgayHL)  'dateHL
                                                    m_PhuCap.DenNgay = DateTime.MinValue
                                                    m_PhuCap.SoQD = m_QDNhanSu.So_QD
                                                    m_PhuCap.NgayQD = m_QDNhanSu.NgayKy_QD
                                                    m_PhuCap.NguoiQD = m_QDNhanSu.NguoiKy_QD
                                                    m_PhuCap.DVraQD = m_QDNhanSu.DVraQD
                                                    m_PhuCap.GhiChu = m_QDNhanSu.GhiChu
                                                    If vIdPhucap <> "" Then
                                                        Dim dtPC As DataTable
                                                        Dim iPC As Integer = 0
                                                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap FROM HS_PhucapCB WHERE IdCanbo='" & vIdCB & "' and Datediff(day,TuNgay,'" & m_QDNhanSu.NgayHL & "')=0 and Upper(SoQD)=Upper(N'" & m_QDNhanSu.So_QD & "') and NgayQD='" & m_QDNhanSu.NgayKy_QD & "' and IdMucPC<>" & m_PhuCap.IdMucPC)
                                                        m_PhuCap.Update()
                                                        ' Update tiếp với các Phụ cấp khác có cùng số QĐ và ngày
                                                        If dtPC.Rows.Count > 0 Then
                                                            For iPC = 0 To dtPC.Rows.Count - 1
                                                                m_PhuCap.IdCB_PhuCap = dtPC.Rows(iPC).Item("IdCB_PhuCap")
                                                                m_PhuCap.UpdateNotFull()
                                                            Next
                                                        End If
                                                    Else
                                                        Dim dtPC As DataTable
                                                        m_PhuCap.Add()
                                                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap, DenNgay FROM HS_PhucapCB t1, MucPhuCap t2 WHERE t1.IdMucPC=t2.IdMuc_PhC and idcanbo='" & vIdCB & "' and IsQD_NHCS=1 and IdLoai_PhC=" & dbconn.getNumber("SELECT IdLoai_PhC FROM Mucphucap WHERE IdMuc_phC=" & m_PhuCap.IdMucPC) & " order by TuNgay desc")
                                                        If dtPC.Rows.Count >= 2 Then
                                                            If dtPC.Rows(1).Item("DenNgay") Is DBNull.Value Then
                                                                db.executeSQL("UPDATE HS_PhuCapCB Set DenNgay='" & m_PhuCap.TuNgay.AddDays(-1) & "' WHERE IdCB_PhuCap= '" & dtPC.Rows(1).Item("IdCB_PhuCap") & "'")
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If vIdPhucap <> "" Then
                                                    m_PhuCap.Delete()
                                                End If
                                            End If
                                        Else
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": số QĐ " & m_QDNhanSu.So_QD & " không cập nhật được." & vbCrLf
                                            GoTo nex
                                        End If
                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If
                                    If m_QDNhanSu.IdQDNhanSu <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If
                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Quyết định nhân sự không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then
            Dim strSql As String = " SELECT t1.idCanbo, MaCB, HoTen, NgayHL, So_QD, (SELECT ten_goi FROM Danhmuc WHERE id_goc=15 and id= IDLoaiQD)  as 'LoaiQD', " & _
                                        "        ((SELECT ten_goi FROM Danhmuc WHERE id_goc=14 and id= IDChucVu_Moi)+', '+ (SELECT ten_phong FROM PhongBan WHERE id=IdPhong_Moi)+', '+(SELECT Ten_goi FROM Chinhanh WHERE id=IDDonVi_Moi)) as 'NoiDung'  " & _
                                        " FROM QDNhanSu t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=1 " & _
                                        " UNION" & _
                                        " SELECT t1.idCanbo, MaCB, HoTen, NgayHL, So_QD, LoaiQD, NoiDung " & _
                                        " FROM QDNhanSu t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=0 order by t1.idCanbo, NgayHL desc "
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                If dt.Rows(j)("NgayHL").ToString().Trim() <> "" Then
                    myDataGrid.Rows(j).Cells("NgayHL").Value = CType(dt.Rows(j)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy")
                Else
                    myDataGrid.Rows(j).Cells("NgayHL").Value = ""
                End If
                myDataGrid.Rows(j).Cells("So_QD").Value = dt.Rows(j)("So_QD").ToString()
                myDataGrid.Rows(j).Cells("LoaiQD").Value = dt.Rows(j)("LoaiQD").ToString()
                myDataGrid.Rows(j).Cells("NoiDung").Value = dt.Rows(j)("NoiDung").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_QDLuong(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "QDLUONG" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "E").value Is Nothing) Then
                                If .cells(row, "E").value.ToString.Trim <> "" Then
                                    Dim m_LuongCB As LuongCanBo = New LuongCanBo
                                    Dim vIdLuong As String = ""
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then
                                        m_LuongCB.IdLuongCB = ""
                                        m_LuongCB.IdCanBo = vIdCB
                                        If .cells(row, "D").value Is Nothing Then
                                            m_LuongCB.IsQD_NHCS = 0
                                        Else
                                            m_LuongCB.IsQD_NHCS = CInt(.cells(row, "D").value)
                                        End If

                                        If .cells(row, "E").value Is Nothing Then
                                            m_LuongCB.SoQD = "ChuaCoSo"
                                        Else
                                            m_LuongCB.SoQD = .cells(row, "E").value.ToString.Trim
                                        End If

                                        'kiem tra mot so dieu kien
                                        If .cells(row, "I").value Is Nothing Then
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định số " & m_LuongCB.SoQD & " Chưa nhập ID loại quyết định lương." & vbCrLf
                                            GoTo nex
                                        End If
                                        If .cells(row, "I").value.ToString.Trim <> "" Then
                                            Dim _value As Integer = 0
                                            If m_LuongCB.IsQD_NHCS Then
                                                If Not (.cells(row, "K").value Is Nothing) Then
                                                    If .cells(row, "K").value.ToString.Trim <> "" Then
                                                        m_LuongCB.IdBacLuong = CInt(.cells(row, "K").value)
                                                    Else
                                                        If Not (.cells(row, "J").value Is Nothing) Then
                                                            If .cells(row, "J").value.ToString.Trim <> "" Then
                                                                _value = CInt(.cells(row, "J").value.ToString.Trim)
                                                                m_LuongCB.IdBacLuong = dbconn.getNumber("SELECT TOP 1 IdBacLuong FROM V$_bacLuong WHERE Heso=" & _value & " order by IdBacLuong")
                                                            Else
                                                                strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_LuongCB.SoQD & " chưa nhập ID Bậc lương." & vbCrLf
                                                                GoTo nex
                                                            End If
                                                        Else
                                                            strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_LuongCB.SoQD & " chưa nhập ID Bậc lương." & vbCrLf
                                                            GoTo nex
                                                        End If
                                                    End If
                                                Else
                                                    If Not (.cells(row, "J").value Is Nothing) Then
                                                        If .cells(row, "J").value.ToString.Trim <> "" Then
                                                            _value = CInt(.cells(row, "J").value.ToString.Trim)
                                                            m_LuongCB.IdBacLuong = dbconn.getNumber("SELECT TOP 1 IdBacLuong FROM V$_bacLuong WHERE Heso=" & _value & " order by IdBacLuong")
                                                        Else
                                                            strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_LuongCB.SoQD & " chưa nhập ID Bậc lương." & vbCrLf
                                                            GoTo nex
                                                        End If
                                                    Else
                                                        strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_LuongCB.SoQD & " chưa nhập ID Bậc lương." & vbCrLf
                                                        GoTo nex
                                                    End If
                                                End If

                                                _value = CInt(.cells(row, "I").value.ToString.Trim)
                                                If Not ((_value >= 719 And _value <= 733) Or (_value >= 753 And _value <= 759) Or _value = 769 Or (_value >= 775 And _value <= 781) Or _value <= 784) Then
                                                    strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID Loại quyết định Lương không đúng. Hãy kiểm tra lại."
                                                    GoTo nex
                                                End If
                                            Else
                                                m_LuongCB.IdBacLuong = 0
                                                'If Not (.cells(row, "J").value Is Nothing) And Not (.cells(row, "M").value Is Nothing) And Not (.cells(row, "N").value Is Nothing) And Not (.cells(row, "O").value Is Nothing) And Not (.cells(row, "P").value Is Nothing) Then
                                                If .cells(row, "J").value Is Nothing And .cells(row, "M").value Is Nothing And .cells(row, "N").value Is Nothing And .cells(row, "O").value Is Nothing And .cells(row, "P").value Is Nothing Then
                                                    strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc không thuộc NHCS số " & m_LuongCB.SoQD & " chưa nhập đủ các thông tin cần thiết của Quyết định nhân sự." & vbCrLf
                                                    GoTo nex
                                                End If
                                            End If
                                        Else
                                            strErr = strErr & vbCrLf & "Dòng " & row & ", Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Loại Quyết định nhân sự." & vbCrLf
                                            GoTo nex
                                        End If

                                        'tiep tuc
                                        If .cells(row, "F").value Is Nothing Then
                                            m_LuongCB.NgayQD = Now.Date
                                        Else
                                            If .cells(row, "F").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "F").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_LuongCB.NgayQD = .cells(row, "F").value
                                            Else
                                                m_LuongCB.NgayQD = DateTimeUtil.getDate(.cells(row, "F").value)
                                            End If
                                        End If

                                        If .cells(row, "V").value Is Nothing Then
                                            m_LuongCB.Ngay_Huong = m_LuongCB.NgayQD
                                        Else
                                            If .cells(row, "V").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_LuongCB.Ngay_Huong = .cells(row, "V").value
                                            Else
                                                m_LuongCB.Ngay_Huong = DateTimeUtil.getDate(.cells(row, "V").value)
                                            End If
                                        End If
                                        If .cells(row, "G").value Is Nothing Then
                                            m_LuongCB.DVraQD = ""
                                        Else
                                            m_LuongCB.DVraQD = standardizeName(.cells(row, "G").value.ToString.Trim)
                                        End If

                                        If .cells(row, "S").value Is Nothing Then
                                            m_LuongCB.NguoiQD = ""
                                        Else
                                            m_LuongCB.NguoiQD = standardizeName(.cells(row, "S").value.ToString.Trim)
                                        End If

                                        If .cells(row, "U").value Is Nothing Then
                                            m_LuongCB.GhiChu = ""
                                        Else
                                            m_LuongCB.GhiChu = standardizeName(.cells(row, "U").value.ToString.Trim)
                                        End If

                                        m_LuongCB.NgayLen_DK = DateTime.Parse("01/01/1900")

                                        If m_LuongCB.IsQD_NHCS Then
                                            If .cells(row, "W").value Is Nothing Then
                                                m_LuongCB.IdCV_Nguoi_QD = 296
                                            Else
                                                If .cells(row, "W").value.ToString.Trim = "" Then
                                                    m_LuongCB.IdCV_Nguoi_QD = 296
                                                Else
                                                    m_LuongCB.IdCV_Nguoi_QD = CInt(.cells(row, "W").value)
                                                End If
                                            End If
                                            m_LuongCB.CV_NguoiKy_QD = ""
                                            If .cells(row, "I").value Is Nothing Then
                                                m_LuongCB.IdLoaiQD = 719
                                            Else
                                                If .cells(row, "I").value.ToString.Trim = "" Then
                                                    m_LuongCB.IdLoaiQD = 719
                                                Else
                                                    m_LuongCB.IdLoaiQD = CInt(.cells(row, "I").value)
                                                End If
                                            End If
                                            m_LuongCB.HeSoLuong = dbconn.getDouble("SELECT HeSo FROM V$_BacLuong WHERE IdBacLuong=" & m_LuongCB.IdBacLuong)
                                            m_LuongCB.NoiDung = ""
                                            m_LuongCB.LoaiQD = ""
                                        Else
                                            m_LuongCB.HeSoLuong = 0
                                            m_LuongCB.IdCV_Nguoi_QD = 0
                                            If .cells(row, "T").value Is Nothing Then
                                                m_LuongCB.CV_NguoiKy_QD = ""
                                            Else
                                                m_LuongCB.CV_NguoiKy_QD = standardizeString(.cells(row, "T").value.ToString.Trim)
                                            End If
                                            m_LuongCB.IdLoaiQD = 0
                                            If .cells(row, "H").value Is Nothing Then
                                                m_LuongCB.LoaiQD = ""
                                            Else
                                                m_LuongCB.LoaiQD = standardizeString(.cells(row, "H").value.ToString.Trim)
                                            End If

                                            If .cells(row, "J").value Is Nothing Then
                                                m_LuongCB.NoiDung = ""
                                            Else
                                                m_LuongCB.NoiDung = standardizeString(.cells(row, "J").value.ToString.Trim)
                                            End If
                                            If Not (.cells(row, "P").value Is Nothing) Then
                                                m_LuongCB.NoiDung += "; Bậc " & standardizeString(.cells(row, "P").value.ToString.Trim)
                                            End If
                                            If Not (.cells(row, "O").value Is Nothing) Then
                                                m_LuongCB.NoiDung += "; Ngạch " & standardizeString(.cells(row, "O").value.ToString.Trim)
                                            End If
                                            If Not (.cells(row, "N").value Is Nothing) Then
                                                m_LuongCB.NoiDung += "; Bảng " & standardizeString(.cells(row, "N").value.ToString.Trim)
                                            End If
                                            If Not (.cells(row, "M").value Is Nothing) Then
                                                m_LuongCB.NoiDung += "; Nghị định " & standardizeString(.cells(row, "M").value.ToString.Trim)
                                            End If
                                        End If
                                        If Not checkQuyetDinh("LUONG", vIdCB, m_LuongCB.SoQD) Then
                                            vIdLuong = m_LuongCB.Add
                                        End If
                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If
                                    If vIdLuong <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If
                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Quyết lương không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then
            Dim strSql As String = " SELECT t1.idCanbo, MaCB, HoTen, Ngay_Huong as 'NgayHL', SoQD as 'So_QD', DVraQD, LoaiQD, NoiDung " & _
                                      "       FROM HS_LuongCB t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=0  AND IDLoaiQD=0 " & _
                                      " UNION " & _
                                      " SELECT t1.idCanbo, MaCB, HoTen, Ngay_Huong as 'NgayHL', SoQD as 'So_QD', DVraQD, (SELECT ten_goi FROM Danhmuc WHERE id_goc=43 and id= IDLoaiQD)  as 'LoaiQD', NoiDung " & _
                                      "       FROM HS_LuongCB t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=0  AND IDLoaiQD>0 " & _
                                      " UNION " & _
                                      " SELECT t1.idCanbo, MaCB, HoTen, Ngay_Huong as 'NgayHL', SoQD as 'So_QD', DVraQD,(SELECT ten_goi FROM Danhmuc WHERE id_goc=43 and id= IDLoaiQD)  as 'LoaiQD', " & _
                                      "       (N'Hệ số '+ convert(varchar(4),heso) + '; '+ MotaBacLuong + N'; Ngạch '+ MotaNgachLuong  + N'; Bảng lương '+ MotaBangLuong +'; '+ TenND  ) as 'NoiDung'  " & _
                                      "       FROM V$_HSLuongCB t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=1 order by t1.idCanbo, Ngay_Huong desc "
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                If dt.Rows(j)("NgayHL").ToString().Trim() <> "" Then
                    myDataGrid.Rows(j).Cells("NgayHL").Value = CType(dt.Rows(j)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy")
                Else
                    myDataGrid.Rows(j).Cells("NgayHL").Value = ""
                End If
                myDataGrid.Rows(j).Cells("So_QD").Value = dt.Rows(j)("So_QD").ToString()
                myDataGrid.Rows(j).Cells("LoaiQD").Value = dt.Rows(j)("LoaiQD").ToString()
                myDataGrid.Rows(j).Cells("NoiDung").Value = dt.Rows(j)("NoiDung").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_DaoTao(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "DAOTAO" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "E").value Is Nothing) Then
                                If .cells(row, "E").value.ToString.Trim <> "" Then
                                    Dim m_DTVBCC As DaoTaoVanBangChungChi = New DaoTaoVanBangChungChi
                                    Dim idDTVBCC As String = ""
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then
                                        ' Kiểm tra một số trường thông tin
                                        Dim _value As Integer = 0
                                        If Not (.cells(row, "F").value Is Nothing) Then
                                            _value = CInt(.cells(row, "F").value.ToString.Trim)
                                            If Not ((_value >= 610 And _value <= 615) Or _value = 783) Then
                                                strErr = strErr & vbCrLf & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID Loại văn bằng chứng chỉ không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "M").value Is Nothing) Then
                                            _value = CInt(.cells(row, "M").value.ToString.Trim)
                                            If Not ((_value >= 200 And _value <= 232) Or _value = 237 Or _value = 238 Or _value = 239 Or _value = 766 Or _value = 767) Then
                                                strErr = strErr & vbCrLf & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID Chuyên ngành đào tạo không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "H").value Is Nothing) Then
                                            _value = CInt(.cells(row, "H").value.ToString.Trim)
                                            If Not (_value >= 167 And _value <= 178) Then
                                                strErr = strErr & vbCrLf & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID Hệ đào tạo không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "K").value Is Nothing) Then
                                            _value = CInt(.cells(row, "K").value.ToString.Trim)
                                            If Not (_value >= 1 And _value <= 173) Then
                                                strErr = strErr & vbCrLf & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID Nước đào tạo không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "O").value Is Nothing) Then
                                            _value = CInt(.cells(row, "O").value.ToString.Trim)
                                            If Not ((_value >= 636 And _value <= 638) Or (_value >= 640 And _value <= 650)) Then
                                                strErr = strErr & vbCrLf & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID Trình độ đào tạo không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "S").value Is Nothing) Then
                                            _value = CInt(.cells(row, "S").value.ToString.Trim)
                                            If Not (_value >= 1 And _value <= 6) Then
                                                strErr = strErr & vbCrLf & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Giá trị ID xếp loại VBCC không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If

                                        ' Tiếp tục
                                        m_DTVBCC.IdDTVBCC = ""
                                        m_DTVBCC.IdCanBo = vIdCB
                                        If .cells(row, "D").value Is Nothing Then
                                            m_DTVBCC.CuDiHoc = 0
                                            m_DTVBCC.TyleHuong = 0
                                        Else
                                            m_DTVBCC.CuDiHoc = CInt(.cells(row, "D").value)
                                            If .cells(row, "Y").value Is Nothing Then
                                                m_DTVBCC.TyleHuong = 0
                                            Else
                                                m_DTVBCC.TyleHuong = CDbl(.cells(row, "Y").value)
                                            End If
                                        End If
                                        If .cells(row, "E").value Is Nothing Then
                                            m_DTVBCC.VBCC = 0
                                        Else
                                            m_DTVBCC.VBCC = CInt(.cells(row, "E").value)
                                        End If
                                        If .cells(row, "F").value Is Nothing Then
                                            m_DTVBCC.IdLoaiVBCC = 610
                                        Else
                                            If .cells(row, "F").value.ToString.Trim = "" Then
                                                m_DTVBCC.IdLoaiVBCC = 610
                                            Else
                                                m_DTVBCC.IdLoaiVBCC = CInt(.cells(row, "F").value)
                                            End If
                                        End If
                                        If .cells(row, "H").value Is Nothing Then
                                            m_DTVBCC.IdHinhThucDT = 178
                                        Else
                                            If .cells(row, "H").value.ToString.Trim = "" Then
                                                m_DTVBCC.IdHinhThucDT = 178
                                            Else
                                                m_DTVBCC.IdHinhThucDT = CInt(.cells(row, "H").value)
                                            End If
                                        End If
                                        If .cells(row, "I").value Is Nothing Then
                                            m_DTVBCC.CoSo_DT = ""
                                        Else
                                            m_DTVBCC.CoSo_DT = standardizeString(.cells(row, "I").value.ToString.Trim)
                                        End If
                                        If .cells(row, "K").value Is Nothing Then
                                            m_DTVBCC.IdNuocDT = 1
                                        Else
                                            If .cells(row, "K").value.ToString.Trim = "" Then
                                                m_DTVBCC.IdNuocDT = 1
                                            Else
                                                m_DTVBCC.IdNuocDT = CInt(.cells(row, "K").value)
                                            End If
                                        End If
                                        If .cells(row, "M").value Is Nothing Then
                                            m_DTVBCC.IdChuyenNganhDT = 239
                                        Else
                                            If .cells(row, "M").value.ToString.Trim = "" Then
                                                m_DTVBCC.IdChuyenNganhDT = 239
                                            Else
                                                m_DTVBCC.IdChuyenNganhDT = CInt(.cells(row, "M").value)
                                            End If
                                        End If
                                        If .cells(row, "O").value Is Nothing Then
                                            m_DTVBCC.IdTrinhDo = 650
                                        Else
                                            If .cells(row, "O").value.ToString.Trim = "" Then
                                                m_DTVBCC.IdTrinhDo = 650
                                            Else
                                                m_DTVBCC.IdTrinhDo = CInt(.cells(row, "O").value)
                                            End If
                                        End If
                                        If .cells(row, "L").value Is Nothing Then
                                            m_DTVBCC.NganhHoc = ""
                                        Else
                                            m_DTVBCC.NganhHoc = standardizeString(.cells(row, "L").value.ToString.Trim)
                                        End If
                                        If .cells(row, "AB").value Is Nothing Then
                                            m_DTVBCC.Lop = ""
                                        Else
                                            m_DTVBCC.Lop = standardizeString(.cells(row, "AB").value.ToString.Trim)
                                        End If
                                        If .cells(row, "P").value Is Nothing Then
                                            m_DTVBCC.TuNgay = DateTime.MinValue
                                        Else
                                            If .cells(row, "P").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "P").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_DTVBCC.TuNgay = .cells(row, "P").value
                                            Else
                                                m_DTVBCC.TuNgay = DateTimeUtil.getDate(.cells(row, "P").value)
                                            End If
                                        End If
                                        If .cells(row, "Q").value Is Nothing Then
                                            m_DTVBCC.DenNgay = DateTime.MinValue
                                        Else
                                            If .cells(row, "Q").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "Q").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_DTVBCC.DenNgay = .cells(row, "Q").value
                                            Else
                                                m_DTVBCC.DenNgay = DateTimeUtil.getDate(.cells(row, "Q").value)
                                            End If
                                        End If
                                        If .cells(row, "T").value Is Nothing Then
                                            m_DTVBCC.HoanThanh = 0
                                            m_DTVBCC.NamTN = 0
                                        Else
                                            If .cells(row, "T").value.ToString.Trim = "" Then
                                                m_DTVBCC.HoanThanh = 0
                                                m_DTVBCC.NamTN = 0
                                            Else
                                                m_DTVBCC.HoanThanh = 1
                                                m_DTVBCC.NamTN = CInt(.cells(row, "T").value)
                                            End If
                                        End If
                                        If .cells(row, "AC").value Is Nothing Then
                                            m_DTVBCC.Ten_VBCC = ""
                                        Else
                                            m_DTVBCC.Ten_VBCC = standardizeString(.cells(row, "AC").value.ToString.Trim)
                                        End If
                                        If .cells(row, "S").value Is Nothing Then
                                            m_DTVBCC.XepLoai = 6
                                        Else
                                            If .cells(row, "S").value.ToString.Trim = "" Then
                                                m_DTVBCC.XepLoai = 6
                                            Else
                                                m_DTVBCC.XepLoai = CByte(.cells(row, "S").value)
                                            End If
                                        End If
                                        If .cells(row, "U").value Is Nothing Then
                                            m_DTVBCC.So_VBCC = ""
                                        Else
                                            m_DTVBCC.So_VBCC = standardizeString(.cells(row, "U").value.ToString.Trim)
                                        End If
                                        If .cells(row, "V").value Is Nothing Then
                                            m_DTVBCC.NgayCap = DateTime.MinValue
                                        Else
                                            If .cells(row, "V").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "V").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_DTVBCC.NgayCap = .cells(row, "V").value
                                            Else
                                                m_DTVBCC.NgayCap = DateTimeUtil.getDate(.cells(row, "V").value)
                                            End If
                                        End If
                                        If .cells(row, "W").value Is Nothing Then
                                            m_DTVBCC.NguoiKy = ""
                                        Else
                                            m_DTVBCC.NguoiKy = standardizeName(.cells(row, "W").value.ToString.Trim)
                                        End If
                                        If .cells(row, "AD").value Is Nothing Then
                                            m_DTVBCC.NgayHL = m_DTVBCC.NgayCap
                                        Else
                                            If .cells(row, "AD").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AD").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_DTVBCC.NgayHL = .cells(row, "AD").value
                                            Else
                                                m_DTVBCC.NgayHL = DateTimeUtil.getDate(.cells(row, "AD").value)
                                            End If
                                        End If
                                        If .cells(row, "AE").value Is Nothing Then
                                            m_DTVBCC.NgayHH = DateTime.MinValue
                                        Else
                                            If .cells(row, "AE").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AE").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_DTVBCC.NgayHH = .cells(row, "AE").value
                                            Else
                                                m_DTVBCC.NgayHH = DateTimeUtil.getDate(.cells(row, "AE").value)
                                            End If
                                        End If
                                        If .cells(row, "Z").value Is Nothing Then
                                            m_DTVBCC.GhiChu = ""
                                        Else
                                            m_DTVBCC.GhiChu = standardizeString(.cells(row, "Z").value.ToString.Trim)
                                        End If

                                        idDTVBCC = m_DTVBCC.Add()
                                        If idDTVBCC <> "" Then
                                            Dim MaChuyenMon As String = getDanhmuc_MaSo(m_DTVBCC.IdChuyenNganhDT)
                                            If MaChuyenMon = "1122" Then
                                                Dim IdTrinhDoCT_MAX As Integer = 0
                                                Dim MaTrinhDo_MAX As String = ""
                                                Dim IdTrinhDoCT_Moi As Integer = 0
                                                IdTrinhDoCT_MAX = getChinhDo_Max(m_DTVBCC.IdCanBo, "1122")
                                                MaTrinhDo_MAX = getDanhmuc_MaSo(IdTrinhDoCT_MAX)
                                                If MaTrinhDo_MAX = "3807" Then
                                                    IdTrinhDoCT_Moi = getDanhmuc_ID("3601")
                                                ElseIf MaTrinhDo_MAX = "3808" Then
                                                    IdTrinhDoCT_Moi = getDanhmuc_ID("3602")
                                                ElseIf MaTrinhDo_MAX = "3809" Then
                                                    IdTrinhDoCT_Moi = getDanhmuc_ID("3603")
                                                End If
                                                If IdTrinhDoCT_Moi <> 0 Then dbconn.executeSQL("UPDATE HS_Canbo SET IdTrinhDoCT=" & IdTrinhDoCT_Moi & " WHERE IdCanBo='" & m_DTVBCC.IdCanBo & "'")
                                            End If

                                        End If
                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If

                                    If idDTVBCC <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If

                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Văn bằng, chứng chỉ của cán bộ không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then
            Dim strSql As String = "SELECT t1.idCanbo, MaCB, HoTen, NamTN, (SELECT ten_goi FROM DanhMuc WHERE Id=IdTrinhDo) as 'TrinhDo', (SELECT ten_goi FROM DanhMuc WHERE Id=IdChuyenNganhDT) as 'ChuyenNganh', NganhHoc FROM HS_DTVBCC t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") order by t1.idcanbo, NamTN desc"
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                myDataGrid.Rows(j).Cells("NamTN").Value = dt.Rows(j)("NamTN").ToString()
                myDataGrid.Rows(j).Cells("TrinhDo").Value = dt.Rows(j)("TrinhDo").ToString()
                myDataGrid.Rows(j).Cells("ChuyenNganh").Value = dt.Rows(j)("ChuyenNganh").ToString()
                myDataGrid.Rows(j).Cells("NganhHoc").Value = dt.Rows(j)("NganhHoc").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_GiaDinhCB(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""
        Dim _PersonnelFile As clsHS_CanBo = New clsHS_CanBo()
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "GiaDinhCB" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "D").value Is Nothing) Then
                                If .cells(row, "D").value.ToString.Trim <> "" Then
                                    Dim obj_gdcb As clsHS_CanBo.HS_GDCB = New clsHS_CanBo.HS_GDCB()
                                    Dim IdTVien As String = ""
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then
                                        ' Kiểm tra một số trường thông tin
                                        If .cells(row, "D").value Is Nothing Then
                                            strErr = strErr & "Họ tên thành viên gia đình cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập. Hãy kiểm tra lại."
                                            GoTo nex
                                        End If
                                        Dim _value As Integer = 0
                                        If Not (.cells(row, "G").value Is Nothing) Then
                                            _value = CInt(.cells(row, "G").value.ToString.Trim)
                                            If Not ((_value >= 447 And _value <= 479) Or _value = 695 Or _value = 696 Or _value = 761) Then
                                                strErr = strErr & "Giá trị ID Quan hệ với cán bộ không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "O").value Is Nothing) Then
                                            _value = CInt(.cells(row, "O").value.ToString.Trim)
                                            If Not (_value >= 1 And _value <= 173) Then
                                                strErr = strErr & "Giá trị ID Quốc tịch không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If
                                        If Not (.cells(row, "T").value Is Nothing) Then
                                            _value = CInt(.cells(row, "T").value.ToString.Trim)
                                            If Not ((_value >= 360 And _value <= 367) Or _value = 690) Then
                                                strErr = strErr & "Giá trị ID Ưu tiên bản thân không đúng. Hãy kiểm tra lại."
                                                GoTo nex
                                            End If
                                        End If

                                        ' Tiếp tục
                                        obj_gdcb.IdTVien = ""
                                        obj_gdcb.IdCanBo = vIdCB
                                        obj_gdcb.Ma_TVien = ""
                                        If .cells(row, "D").value Is Nothing Then
                                            obj_gdcb.HoTen = ""
                                        Else
                                            obj_gdcb.HoTen = standardizeName(.cells(row, "D").value.ToString.Trim)
                                        End If
                                        If .cells(row, "G").value Is Nothing Then
                                            obj_gdcb.IdQuanHe = 696
                                        Else
                                            If .cells(row, "G").value.ToString.Trim = "" Then
                                                obj_gdcb.IdQuanHe = 696
                                            Else
                                                obj_gdcb.IdQuanHe = CInt(.cells(row, "G").value)
                                            End If
                                        End If
                                        If .cells(row, "Q").value Is Nothing Then
                                            obj_gdcb.GioiTinh = 0
                                        Else
                                            obj_gdcb.GioiTinh = CByte(.cells(row, "Q").value)
                                        End If
                                        If .cells(row, "E").value Is Nothing Then
                                            obj_gdcb.NamSinh = 1945
                                        Else
                                            If .cells(row, "E").value.ToString.Trim = "" Then
                                                obj_gdcb.NamSinh = 1945
                                            Else
                                                obj_gdcb.NamSinh = CInt(.cells(row, "E").value)
                                            End If
                                        End If
                                        If .cells(row, "R").value Is Nothing Then
                                            obj_gdcb.ConMat = 0
                                        Else
                                            obj_gdcb.ConMat = CByte(.cells(row, "R").value)
                                        End If
                                        obj_gdcb.NamMat = 0
                                        obj_gdcb.LyDo_Mat = ""
                                        If .cells(row, "T").value Is Nothing Then
                                            obj_gdcb.IdUT_BThan = 690
                                        Else
                                            If .cells(row, "T").value.ToString.Trim = "" Then
                                                obj_gdcb.IdUT_BThan = 690
                                            Else
                                                obj_gdcb.IdUT_BThan = CInt(.cells(row, "T").value)
                                            End If
                                        End If
                                        If .cells(row, "O").value Is Nothing Then
                                            obj_gdcb.IdQuocGia = 1
                                        Else
                                            If .cells(row, "O").value.ToString.Trim = "" Then
                                                obj_gdcb.IdQuocGia = 1
                                            Else
                                                obj_gdcb.IdQuocGia = CInt(.cells(row, "O").value)
                                            End If
                                        End If
                                        obj_gdcb.IdQueQuan = dbconn.getNumber("SELECT IdNQ_Huyen FROM HS_CanBo WHERE IdCanBo='" & vIdCB & "'")
                                        obj_gdcb.DienThoai = ""
                                        If .cells(row, "L").value Is Nothing Then
                                            obj_gdcb.DiaChi = ""
                                        Else
                                            obj_gdcb.DiaChi = standardizeString(.cells(row, "L").value.ToString.Trim)
                                        End If
                                        If .cells(row, "K").value Is Nothing Then
                                            obj_gdcb.NgheNghiep = ""
                                        Else
                                            obj_gdcb.NgheNghiep = standardizeString(.cells(row, "K").value.ToString.Trim)
                                        End If

                                        IdTVien = _PersonnelFile.Insert_HS_GDCB(obj_gdcb)

                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If

                                    If IdTVien <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If

                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Thành viên gia đình cán bộ không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then
            Dim strSql As String = "SELECT t2.idCanbo, MaCB, t2.HoTen, t1.HoTen as 'TVien', (SELECT ten_goi FROM DanhMuc WHERE Id=IdQuanHe) as 'QuanHe', t1.NgheNghiep, t1.DiaChi FROM hs_gdcb t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t2.idCanbo in (" & dsCanBo & ") Order by  t1.idcanbo"
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                myDataGrid.Rows(j).Cells("TVien").Value = dt.Rows(j)("TVien").ToString()
                myDataGrid.Rows(j).Cells("QuanHe").Value = dt.Rows(j)("QuanHe").ToString()
                myDataGrid.Rows(j).Cells("NgheNghiep").Value = dt.Rows(j)("NgheNghiep").ToString()
                myDataGrid.Rows(j).Cells("DiaChi").Value = dt.Rows(j)("DiaChi").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_QDTiepNhan(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""
        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "TIEPNHAN" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "E").value Is Nothing) Then
                                If .cells(row, "E").value.ToString.Trim <> "" Then
                                    Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                                    Dim m_LuongOld As LuongCanBo = New LuongCanBo
                                    Dim m_PhucapOld As Phucap = New Phucap
                                    Dim vIdQDNhanSu As String = ""
                                    Dim vIdLuong As String = ""
                                    Dim vIdPhucap As String = ""
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then

                                        m_QDNhanSu.IdQDNhanSu = ""
                                        m_QDNhanSu.IdCanBo = vIdCB
                                        If .cells(row, "D").value Is Nothing Then
                                            m_QDNhanSu.IsQD_NHCS = 0
                                        Else
                                            m_QDNhanSu.IsQD_NHCS = CInt(.cells(row, "D").value)
                                        End If

                                        If .cells(row, "E").value Is Nothing Then
                                            m_QDNhanSu.So_QD = "ChuaCoSo"
                                        Else
                                            m_QDNhanSu.So_QD = .cells(row, "E").value.ToString.Trim
                                        End If

                                        'kiem tra mot so dieu kien
                                        If .cells(row, "I").value Is Nothing Then
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định số " & m_QDNhanSu.So_QD & " Chưa nhập ID loại quyết định nhân sự." & vbCrLf
                                            GoTo nex
                                        End If
                                        If .cells(row, "I").value.ToString.Trim <> "" Then
                                            If m_QDNhanSu.IsQD_NHCS Then
                                                If Not (.cells(row, "L").value Is Nothing) And Not (.cells(row, "N").value Is Nothing) And Not (.cells(row, "P").value Is Nothing) And Not (.cells(row, "R").value Is Nothing) Then
                                                    Dim _value As Integer = 0
                                                    _value = CInt(.cells(row, "I").value.ToString.Trim)
                                                    If Not ((_value >= 321 And _value <= 327) Or _value = 657 Or _value = 658 Or _value = 688 Or _value = 711 Or (_value >= 714 And _value <= 717) Or _value = 748 Or _value = 768 Or (_value >= 770 And _value <= 774)) Then
                                                        strErr = strErr & "Giá trị ID Loại quyết định nhân sự không đúng. Hãy kiểm tra lại."
                                                        GoTo nex
                                                    End If
                                                    _value = CInt(.cells(row, "L").value.ToString.Trim)
                                                    If Not ((295 <= _value And _value <= 320) Or _value = 691 Or _value = 706 Or (738 <= _value And _value <= 747) Or _value <= 765) Then
                                                        strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_QDNhanSu.So_QD & " nhập sai ID Chức vụ." & vbCrLf
                                                        GoTo nex
                                                    End If
                                                    _value = CInt(.cells(row, "N").value.ToString.Trim)
                                                    If Not ((240 <= _value And _value <= 262) Or _value = 266 Or _value = 268 Or _value = 269 Or (272 <= _value And _value <= 282) Or (762 <= _value And _value <= 764)) Then
                                                        strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_QDNhanSu.So_QD & " nhập sai ID của Chuyên môn." & vbCrLf
                                                        GoTo nex
                                                    End If
                                                Else
                                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc NHCS số " & m_QDNhanSu.So_QD & " chưa nhập đủ các thông tin cần thiết của Quyết định nhân sự." & vbCrLf
                                                    GoTo nex
                                                End If
                                            Else
                                                If Not (.cells(row, "H").value Is Nothing) And Not (.cells(row, "J").value Is Nothing) And Not (.cells(row, "AH").value Is Nothing) Then
                                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định thuộc không thuộc NHCS số " & m_QDNhanSu.So_QD & " chưa nhập đủ các thông tin cần thiết của Quyết định nhân sự." & vbCrLf
                                                    GoTo nex
                                                End If
                                            End If
                                        Else
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Loại Quyết định nhân sự." & vbCrLf
                                            GoTo nex
                                        End If
                                        ' tiep tuc thuc hien
                                        If .cells(row, "F").value Is Nothing Then
                                            m_QDNhanSu.NgayKy_QD = Now.Date
                                        Else
                                            If .cells(row, "F").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "F").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_QDNhanSu.NgayKy_QD = .cells(row, "F").value
                                            Else
                                                m_QDNhanSu.NgayKy_QD = DateTimeUtil.getDate(.cells(row, "F").value)
                                            End If
                                        End If

                                        If .cells(row, "AC").value Is Nothing Then
                                            m_QDNhanSu.NguoiKy_QD = ""
                                        Else
                                            m_QDNhanSu.NguoiKy_QD = standardizeName(.cells(row, "AC").value.ToString.Trim)
                                        End If

                                        If .cells(row, "J").value Is Nothing Then
                                            m_QDNhanSu.NgayHL = m_QDNhanSu.NgayKy_QD
                                        Else
                                            If .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_QDNhanSu.NgayHL = .cells(row, "J").value
                                            Else
                                                'If row = 14 Then
                                                '    MsgBox("debug")
                                                'End If
                                                m_QDNhanSu.NgayHL = DateTimeUtil.getDate(.cells(row, "J").value)
                                            End If
                                        End If
                                        m_QDNhanSu.NgayBoNhiem_TT = DateTime.Parse("01/01/1900")
                                        m_QDNhanSu.NgayThoiLuong = DateTime.Parse("01/01/1900")

                                        If .cells(row, "G").value Is Nothing Then
                                            m_QDNhanSu.DVraQD = ""
                                        Else
                                            m_QDNhanSu.DVraQD = standardizeString(.cells(row, "G").value.ToString.Trim)
                                        End If

                                        If .cells(row, "I").value Is Nothing Then
                                            m_QDNhanSu.IdLoaiQD = 321
                                        Else
                                            If .cells(row, "I").value.ToString.Trim = "" Then
                                                m_QDNhanSu.IdLoaiQD = 321
                                            Else
                                                m_QDNhanSu.IdLoaiQD = CInt(.cells(row, "I").value)
                                            End If
                                        End If

                                        If m_QDNhanSu.IsQD_NHCS Then
                                            Dim m_QDNSfinal As QDNhanSu = New QDNhanSu
                                            m_QDNSfinal = m_QDNSfinal.getFinalRecord(vIdCB, m_QDNhanSu.NgayHL)
                                            m_QDNhanSu.IdDonvi_Cu = m_QDNSfinal.IdDonvi_Moi
                                            m_QDNhanSu.IdPhong_Cu = m_QDNSfinal.IdPhong_Moi
                                            m_QDNhanSu.IdChucvu_Cu = m_QDNSfinal.IdChucvu_Moi
                                            m_QDNhanSu.IdChuyenMon_Cu = m_QDNSfinal.IdChuyenMon_Moi
                                            m_QDNhanSu.IdDonvi_Moi = CInt(.cells(row, "R").value)
                                            m_QDNhanSu.IdPhong_Moi = CInt(.cells(row, "P").value)
                                            m_QDNhanSu.IdChucvu_Moi = CInt(.cells(row, "L").value)
                                            m_QDNhanSu.IdChuyenMon_Moi = CInt(.cells(row, "N").value)
                                            If .cells(row, "AE").value Is Nothing Then
                                                m_QDNhanSu.idCV_Nguoiky_QD = 296
                                            Else
                                                If .cells(row, "AE").value.ToString.Trim = "" Then
                                                    m_QDNhanSu.idCV_Nguoiky_QD = 296
                                                Else
                                                    m_QDNhanSu.idCV_Nguoiky_QD = CInt(.cells(row, "AE").value)
                                                End If
                                            End If
                                            m_QDNhanSu.CV_NguoiKy_QD = ""
                                            m_QDNhanSu.LoaiQD = ""
                                            m_QDNhanSu.DenNgay = DateTime.MinValue
                                            m_QDNhanSu.NoiDung = ""
                                        Else
                                            m_QDNhanSu.IdDonvi_Cu = 0
                                            m_QDNhanSu.IdPhong_Cu = 0
                                            m_QDNhanSu.IdChucvu_Cu = 0
                                            m_QDNhanSu.IdChuyenMon_Cu = 0
                                            m_QDNhanSu.IdDonvi_Moi = 0
                                            m_QDNhanSu.IdPhong_Moi = 0
                                            m_QDNhanSu.IdChucvu_Moi = 0
                                            If .cells(row, "AG").value Is Nothing Then
                                                m_QDNhanSu.DenNgay = m_QDNhanSu.NgayKy_QD
                                            Else
                                                If .cells(row, "AG").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "AG").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                    m_QDNhanSu.DenNgay = .cells(row, "AG").value
                                                Else
                                                    m_QDNhanSu.DenNgay = DateTimeUtil.getDate(.cells(row, "AG").value)
                                                End If
                                            End If
                                            If .cells(row, "AD").value Is Nothing Then
                                                m_QDNhanSu.CV_NguoiKy_QD = ""
                                            Else
                                                m_QDNhanSu.CV_NguoiKy_QD = standardizeString(.cells(row, "AD").value.ToString.Trim)
                                            End If
                                            m_QDNhanSu.idCV_Nguoiky_QD = 0
                                            If .cells(row, "H").value Is Nothing Then
                                                m_QDNhanSu.LoaiQD = ""
                                            Else
                                                m_QDNhanSu.LoaiQD = standardizeString(.cells(row, "H").value.ToString.Trim)
                                            End If
                                            If .cells(row, "AH").value Is Nothing Then
                                                If .cells(row, "K").value Is Nothing And .cells(row, "Q").value Is Nothing Then
                                                    m_QDNhanSu.NoiDung = ""
                                                Else
                                                    m_QDNhanSu.NoiDung = .cells(row, "K").value.ToString.Trim & " " & .cells(row, "Q").value.ToString.Trim
                                                End If
                                            Else
                                                m_QDNhanSu.NoiDung = standardizeString(.cells(row, "AH").value.ToString.Trim)
                                            End If
                                        End If

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

                                        m_QDNhanSu.Active = 1
                                        If .cells(row, "AF").value Is Nothing Then
                                            m_QDNhanSu.GhiChu = ""
                                        Else
                                            m_QDNhanSu.GhiChu = standardizeString(.cells(row, "AF").value.ToString.Trim)
                                        End If

                                        vIdLuong = m_LuongOld.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)
                                        vIdPhucap = m_PhucapOld.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)
                                        If Not checkQuyetDinh("QDNHANSU", vIdCB, m_QDNhanSu.So_QD) Then
                                            vIdQDNhanSu = m_QDNhanSu.Add()
                                        End If

                                        If vIdQDNhanSu <> "" Then
                                            If m_QDNhanSu.IsKiemNhiem = 2 Then
                                                'cap nhat lai QD kiem nhiem truoc do
                                                dbconn.executeSQL("UPDATE QDNhanSu SET IsKiemNhiem=2 WHERE IdCanBo='" & vIdCB & "' and IsKiemNhiem=1 and IdDonvi_Moi=" & m_QDNhanSu.IdDonvi_Moi & " and IdPhong_Moi=" & m_QDNhanSu.IdPhong_Moi & " and IdChucVu_Moi=" & m_QDNhanSu.IdChucvu_Moi)
                                            End If
                                            Dim m_Luong As LuongCanBo = New LuongCanBo
                                            Dim m_PhuCap As Phucap = New Phucap
                                            Dim maQDNhanSu As String = ""
                                            Dim IDLoaiQDLuong As String = ""
                                            Dim db As DBAccess = New DBAccess
                                            m_Luong.IdLuongCB = vIdLuong
                                            m_Luong.IdCanBo = vIdCB

                                            If Not (.cells(row, "V").value Is Nothing) Then
                                                If Not (.cells(row, "V").value.ToString.Trim = "" Or .cells(row, "V").value.ToString.Trim = "0") Then
                                                    m_Luong.Ngay_Huong = DateTimeUtil.getDate(m_QDNhanSu.NgayHL)
                                                    m_Luong.DVraQD = m_QDNhanSu.DVraQD
                                                    If m_QDNhanSu.IsQD_NHCS Then
                                                        'QD lương kem theo QD nhan su thuoc NHCS
                                                        If .cells(row, "AA").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Bậc lương." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "AA").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Bậc lương." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        Dim dt As System.Data.DataTable = New System.Data.DataTable
                                                        Dim IdNgach As Integer = 0
                                                        Dim Heso As Double = 0
                                                        dt = dbconn.SelectDBRows("SELECT IdNgachLuong, Heso FROM V$_BacLuong WHERE Idbacluong=" & CInt(.cells(row, "AA").value))
                                                        If dt.Rows.Count > 0 Then
                                                            IdNgach = dt.Rows(0).Item("IdNgachLuong")
                                                            Heso = dt.Rows(0).Item("Heso")
                                                        End If
                                                        m_Luong.NgayLen_DK = m_QDNhanSu.NgayHL.AddMonths(getTimeNangBac(IdNgach))
                                                        m_Luong.IsQD_NHCS = 1
                                                        m_Luong.NoiDung = ""
                                                        m_Luong.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                                                        m_Luong.CV_NguoiKy_QD = ""
                                                        m_Luong.IdBacLuong = CInt(.cells(row, "AA").value)
                                                        m_Luong.HeSoLuong = Heso
                                                        m_Luong.LoaiQD = ""
                                                    Else
                                                        'QD lương kem theo QD nhan su không thuoc NHCS
                                                        If .cells(row, "AB").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập thông tin về lương trong QĐ nhân sự." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "AB").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập thông tin về lương trong QĐ nhân sự." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        m_Luong.NgayLen_DK = Date.MinValue
                                                        m_Luong.IdCV_Nguoi_QD = 0
                                                        m_Luong.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                                                        m_Luong.NoiDung = .cells(row, "AB").value.ToString.Trim
                                                        m_Luong.IsQD_NHCS = 0
                                                        m_Luong.LoaiQD = m_QDNhanSu.LoaiQD
                                                    End If
                                                    maQDNhanSu = getDanhmuc_MaSo(m_QDNhanSu.IdLoaiQD)
                                                    IDLoaiQDLuong = db.getNumber("SELECT ID FROM DanhMuc WHERE ma_so ='43" & (CInt(maQDNhanSu.Substring(2, 1)) + 1) & maQDNhanSu.Substring(3, 1) & "' and ma_so in ('4311','4312','4313','4314','4315','4317','4318','4319','4327','4328','4329','4330','4331','4332','4323','4324','4325')")
                                                    If IDLoaiQDLuong = 0 Then
                                                        IDLoaiQDLuong = getDanhmuc_ID("4323")
                                                    End If
                                                    m_Luong.IdLoaiQD = IDLoaiQDLuong
                                                    m_Luong.SoQD = m_QDNhanSu.So_QD
                                                    m_Luong.NgayQD = m_QDNhanSu.NgayKy_QD
                                                    m_Luong.NguoiQD = m_QDNhanSu.NguoiKy_QD
                                                    m_Luong.GhiChu = m_QDNhanSu.GhiChu
                                                    If vIdLuong <> "" Then
                                                        m_Luong.Update()
                                                    Else
                                                        ' Kiem tra thông tin lương trong QDNS neu trung với thông tin lương hiện đang hưởng của CB thì ko cho nhập sang HS Luong
                                                        Dim m_LuongFinal As LuongCanBo = New LuongCanBo
                                                        m_LuongFinal = m_LuongFinal.getFinalRecord(vIdCB)
                                                        If Not (m_LuongFinal.IdBacLuong = m_Luong.IdBacLuong And m_Luong.IsQD_NHCS) Then
                                                            m_Luong.Add()
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If vIdLuong <> "" Then
                                                    m_Luong.Delete()
                                                End If
                                            End If

                                            m_PhuCap.IdCB_PhuCap = vIdPhucap
                                            m_PhuCap.IdCanBo = vIdCB
                                            If Not (.cells(row, "S").value Is Nothing) Then
                                                If Not (.cells(row, "S").value.ToString.Trim = "" Or .cells(row, "S").value.ToString.Trim = "0") Then
                                                    If m_QDNhanSu.IsQD_NHCS Then
                                                        If .cells(row, "T").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Mức phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "T").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Mức phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        m_PhuCap.IsQD_NHCS = 1
                                                        m_PhuCap.NoiDung = ""
                                                        m_PhuCap.IdMucPC = CInt(.cells(row, "T").value)
                                                        m_PhuCap.IdCV_Nguoi_QD = m_QDNhanSu.idCV_Nguoiky_QD
                                                        m_PhuCap.CV_NguoiKy_QD = ""
                                                    Else
                                                        If .cells(row, "U").value Is Nothing Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập Nội dung hưởng phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        If Not (.cells(row, "U").value.ToString.Trim <> "") Then
                                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập Nội dung hưởng phụ cấp." & vbCrLf
                                                            m_QDNhanSu.Delete()
                                                            GoTo nex
                                                        End If
                                                        m_PhuCap.IdCV_Nguoi_QD = 0
                                                        m_PhuCap.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                                                        m_PhuCap.IsQD_NHCS = 0
                                                        m_PhuCap.NoiDung = .cells(row, "U").value.ToString.Trim
                                                    End If
                                                    m_PhuCap.TuNgay = DateTimeUtil.getDate(m_QDNhanSu.NgayHL)  'dateHL
                                                    m_PhuCap.DenNgay = DateTime.MinValue
                                                    m_PhuCap.SoQD = m_QDNhanSu.So_QD
                                                    m_PhuCap.NgayQD = m_QDNhanSu.NgayKy_QD
                                                    m_PhuCap.NguoiQD = m_QDNhanSu.NguoiKy_QD
                                                    m_PhuCap.DVraQD = m_QDNhanSu.DVraQD
                                                    m_PhuCap.GhiChu = m_QDNhanSu.GhiChu
                                                    If vIdPhucap <> "" Then
                                                        Dim dtPC As DataTable
                                                        Dim iPC As Integer = 0
                                                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap FROM HS_PhucapCB WHERE IdCanbo='" & vIdCB & "' and Datediff(day,TuNgay,'" & m_QDNhanSu.NgayHL & "')=0 and Upper(SoQD)=Upper(N'" & m_QDNhanSu.So_QD & "') and NgayQD='" & m_QDNhanSu.NgayKy_QD & "' and IdMucPC<>" & m_PhuCap.IdMucPC)
                                                        m_PhuCap.Update()
                                                        ' Update tiếp với các Phụ cấp khác có cùng số QĐ và ngày
                                                        If dtPC.Rows.Count > 0 Then
                                                            For iPC = 0 To dtPC.Rows.Count - 1
                                                                m_PhuCap.IdCB_PhuCap = dtPC.Rows(iPC).Item("IdCB_PhuCap")
                                                                m_PhuCap.UpdateNotFull()
                                                            Next
                                                        End If
                                                    Else
                                                        Dim dtPC As DataTable
                                                        m_PhuCap.Add()
                                                        dtPC = db.SelectDBRows("SELECT IdCB_PhuCap, DenNgay FROM HS_PhucapCB t1, MucPhuCap t2 WHERE t1.IdMucPC=t2.IdMuc_PhC and idcanbo='" & vIdCB & "' and IsQD_NHCS=1 and IdLoai_PhC=" & dbconn.getNumber("SELECT IdLoai_PhC FROM Mucphucap WHERE IdMuc_phC=" & m_PhuCap.IdMucPC) & " order by TuNgay desc")
                                                        If dtPC.Rows.Count >= 2 Then
                                                            If dtPC.Rows(1).Item("DenNgay") Is DBNull.Value Then
                                                                db.executeSQL("UPDATE HS_PhuCapCB Set DenNgay='" & m_PhuCap.TuNgay.AddDays(-1) & "' WHERE IdCB_PhuCap= '" & dtPC.Rows(1).Item("IdCB_PhuCap") & "'")
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If vIdPhucap <> "" Then
                                                    m_PhuCap.Delete()
                                                End If
                                            End If
                                        Else
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": số QĐ " & m_QDNhanSu.So_QD & " không cập nhật được." & vbCrLf
                                            GoTo nex
                                        End If
                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If
                                    If m_QDNhanSu.IdQDNhanSu <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If
                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Quyết định nhân sự không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then

            'Dim arrDel() As String
            'arrDel = dsCanBo.Split(",")
            Dim strSql As String = " SELECT t1.idCanbo, MaCB, HoTen, NgayHL, So_QD, (SELECT ten_goi FROM Danhmuc WHERE id_goc=15 and id= IDLoaiQD)  as 'LoaiQD', " & _
                                        "        ((SELECT ten_goi FROM Danhmuc WHERE id_goc=14 and id= IDChucVu_Moi)+', '+ (SELECT ten_phong FROM PhongBan WHERE id=IdPhong_Moi)+', '+(SELECT Ten_goi FROM Chinhanh WHERE id=IDDonVi_Moi)) as 'NoiDung'  " & _
                                        " FROM QDNhanSu t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=1 " & _
                                        " UNION" & _
                                        " SELECT t1.idCanbo, MaCB, HoTen, NgayHL, So_QD, LoaiQD, NoiDung " & _
                                        " FROM QDNhanSu t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") AND IsQD_NHCS=0 order by t1.idCanbo, NgayHL desc "
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()  'NgayHL, So_QD, DVraQD, LoaiQD, NoiDung
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                If dt.Rows(j)("NgayHL").ToString().Trim() <> "" Then
                    myDataGrid.Rows(j).Cells("NgayHL").Value = CType(dt.Rows(j)("NgayHL").ToString(), DateTime).ToString("dd-MM-yyyy")
                Else
                    myDataGrid.Rows(j).Cells("NgayHL").Value = ""
                End If
                myDataGrid.Rows(j).Cells("So_QD").Value = dt.Rows(j)("So_QD").ToString()
                myDataGrid.Rows(j).Cells("LoaiQD").Value = dt.Rows(j)("LoaiQD").ToString()
                myDataGrid.Rows(j).Cells("NoiDung").Value = dt.Rows(j)("NoiDung").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_QDKhac(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""

        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "QDKHAC" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "D").value Is Nothing) Then
                                If .cells(row, "D").value.ToString.Trim <> "" Then
                                    Dim m_QDKhac As QDKhac = New QDKhac
                                    Dim IdQdKhac As String = ""
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then
                                        m_QDKhac.IdQDKhac = ""
                                        m_QDKhac.IdCanBo = vIdCB
                                        If .cells(row, "D").value Is Nothing Then
                                            m_QDKhac.So_QD = "ChuaCoSo"
                                        Else
                                            m_QDKhac.So_QD = .cells(row, "D").value.ToString.Trim
                                        End If

                                        If .cells(row, "E").value Is Nothing Then
                                            m_QDKhac.NgayKy_QD = Now.Date
                                        Else
                                            If .cells(row, "E").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "E").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_QDKhac.NgayKy_QD = .cells(row, "E").value
                                            Else
                                                m_QDKhac.NgayKy_QD = DateTimeUtil.getDate(.cells(row, "E").value)
                                            End If
                                        End If

                                        If .cells(row, "F").value Is Nothing Then
                                            m_QDKhac.DVraQD = ""
                                        Else
                                            m_QDKhac.DVraQD = standardizeString(.cells(row, "F").value.ToString.Trim)
                                        End If

                                        If .cells(row, "G").value Is Nothing Then
                                            m_QDKhac.NoiDung = ""
                                        Else
                                            m_QDKhac.NoiDung = standardizeString(.cells(row, "G").value.ToString.Trim)
                                        End If

                                        If .cells(row, "H").value Is Nothing Then
                                            m_QDKhac.NguoiKy_QD = ""
                                        Else
                                            m_QDKhac.NguoiKy_QD = standardizeName(.cells(row, "H").value.ToString.Trim)
                                        End If

                                        If .cells(row, "I").value Is Nothing Then
                                            m_QDKhac.CV_Nguoiky_QD = ""
                                        Else
                                            m_QDKhac.CV_Nguoiky_QD = standardizeString(.cells(row, "I").value.ToString.Trim)
                                        End If

                                        If .cells(row, "J").value Is Nothing Then
                                            m_QDKhac.GhiChu = ""
                                        Else
                                            m_QDKhac.GhiChu = standardizeString(.cells(row, "J").value.ToString.Trim)
                                        End If

                                        IdQdKhac = m_QDKhac.Add()
                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If

                                    If IdQdKhac <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If

                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Quyết định khác của cán bộ không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then
            Dim strSql As String = " SELECT t1.idCanbo, MaCB, HoTen, NgayKy_QD, So_QD, DVraQD, NoiDung " & _
                                           " FROM QDKhac t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") order by t1.idCanbo, NgayKy_QD desc "
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                If dt.Rows(j)("NgayKy_QD").ToString().Trim() <> "" Then
                    myDataGrid.Rows(j).Cells("NgayKy_QD").Value = CType(dt.Rows(j)("NgayKy_QD").ToString(), DateTime).ToString("dd-MM-yyyy")
                Else
                    myDataGrid.Rows(j).Cells("NgayKy_QD").Value = ""
                End If
                myDataGrid.Rows(j).Cells("So_QD").Value = dt.Rows(j)("So_QD").ToString()
                myDataGrid.Rows(j).Cells("DVraQD").Value = dt.Rows(j)("DVraQD").ToString()
                myDataGrid.Rows(j).Cells("NoiDung").Value = dt.Rows(j)("NoiDung").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub ImportExcel_HDLD(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView, ByVal vSheet As Integer)

        Dim Excel As Object = CreateObject("Excel.Application")
        Dim row As Integer
        Dim countCB As Integer = 0
        Dim dsCanBo As String = ""
        Dim dbconn As DBAccess = New DBAccess
        Dim strErr As String = ""
        Dim vIDCB_previous As String = ""
        Dim _Labour As clsHS_Hdld = New clsHS_Hdld()

        Try
            With Excel
                .Workbooks.Open(PathExcelFile)
                .SheetsInNewWorkbook = 1
                .Worksheets(vSheet).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "HDLD" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            If Not (.cells(row, "F").value Is Nothing) Then
                                If .cells(row, "F").value.ToString.Trim <> "" Then
                                    Dim m_HDLD As clsHS_Hdld.HS_HDLD = New clsHS_Hdld.HS_HDLD()
                                    Dim IdHDLD As String = ""
                                    'If row = 14 Then
                                    '    MsgBox("debug")
                                    'End If
                                    vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                                    If vIdCB <> "" Then
                                        If .cells(row, "J").value Is Nothing Then
                                            strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": chưa nhập đầy đủ ngày bắt đầu hợp đồng." & vbCrLf
                                            GoTo nex
                                        Else
                                            If .cells(row, "J").value.ToString.Trim = "" Then
                                                strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": chưa nhập đầy đủ ngày bắt đầu hợp đồng." & vbCrLf
                                                GoTo nex
                                            End If
                                        End If

                                        m_HDLD.IdCB_HDLD = ""
                                        m_HDLD.IdCanBo = vIdCB
                                        If .cells(row, "D").value Is Nothing Then
                                            m_HDLD.SoHD = "ChuaCoSo"
                                        Else
                                            m_HDLD.SoHD = .cells(row, "D").value.ToString.Trim
                                        End If

                                        m_HDLD.IdLoaiHD = CInt(.cells(row, "F").value)

                                        If .cells(row, "G").value Is Nothing Then
                                            m_HDLD.NgayKy_HD = Now.Date
                                        Else
                                            If .cells(row, "G").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "G").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_HDLD.NgayKy_HD = .cells(row, "G").value
                                            Else
                                                m_HDLD.NgayKy_HD = DateTimeUtil.getDate(.cells(row, "G").value)
                                            End If
                                        End If

                                        If .cells(row, "H").value Is Nothing Then
                                            m_HDLD.DVKyHDLD = ""
                                        Else
                                            m_HDLD.DVKyHDLD = standardizeString(.cells(row, "H").value.ToString.Trim)
                                        End If

                                        If .cells(row, "I").value Is Nothing Then
                                            m_HDLD.NoiLamViec = ""
                                        Else
                                            m_HDLD.NoiLamViec = standardizeString(.cells(row, "I").value.ToString.Trim)
                                        End If

                                        If .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "J").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                            m_HDLD.Ngay_HL = .cells(row, "J").value
                                        Else
                                            m_HDLD.Ngay_HL = DateTimeUtil.getDate(.cells(row, "J").value)
                                        End If
                                        m_HDLD.TuNgay = m_HDLD.Ngay_HL

                                        If .cells(row, "K").value Is Nothing Then
                                            m_HDLD.DenNgay = DateTime.Parse("01/01/1900")
                                        Else
                                            If .cells(row, "K").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "K").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                                m_HDLD.DenNgay = .cells(row, "K").value
                                            Else
                                                m_HDLD.DenNgay = DateTimeUtil.getDate(.cells(row, "K").value)
                                            End If
                                        End If

                                        If .cells(row, "L").value Is Nothing Then
                                            m_HDLD.IdHT_TraLuong = 654
                                        Else
                                            m_HDLD.IdHT_TraLuong = CInt(.cells(row, "L").value)
                                        End If

                                        If .cells(row, "O").value Is Nothing Then
                                            m_HDLD.IdBacLuong = 0
                                        Else
                                            m_HDLD.IdBacLuong = CInt(.cells(row, "O").value)
                                        End If

                                        If m_HDLD.IdBacLuong > 0 Then
                                            If .cells(row, "M").value Is Nothing Then
                                                m_HDLD.HeSo = 0
                                            Else
                                                m_HDLD.HeSo = CDbl(.cells(row, "M").value.ToString.Trim)
                                            End If
                                        Else
                                            If .cells(row, "M").value Is Nothing Then
                                                m_HDLD.GhiChu = ""
                                            Else
                                                m_HDLD.GhiChu = .cells(row, "M").value.ToString.Trim
                                            End If
                                        End If

                                        If .cells(row, "W").value Is Nothing Then
                                            m_HDLD.TyleHuong = 0
                                        Else
                                            m_HDLD.TyleHuong = CDbl(.cells(row, "W").value.ToString.Trim)
                                        End If

                                        If .cells(row, "X").value Is Nothing Then
                                            m_HDLD.NguoKy_QD = ""
                                        Else
                                            m_HDLD.NguoKy_QD = standardizeName(.cells(row, "X").value.ToString.Trim)
                                        End If

                                        If .cells(row, "Z").value Is Nothing Then
                                            m_HDLD.IdCV_Nguoiky_QD = 0
                                        Else
                                            m_HDLD.IdCV_Nguoiky_QD = CInt(.cells(row, "Z").value)
                                        End If

                                        If Not (.cells(row, "AA").value Is Nothing) Then
                                            m_HDLD.GhiChu = m_HDLD.GhiChu & " " & standardizeString(.cells(row, "AA").value.ToString.Trim)
                                        End If
                                       
                                        IdHDLD = _Labour.Insert_LaborContract(m_HDLD)

                                    Else
                                        strErr = strErr & "Cán bộ có mã " & .cells(row, "B").value.ToString.Trim & " chưa nhập vào hệ thống."
                                        GoTo nex
                                    End If

                                    If IdHDLD <> "" Then
                                        countCB = countCB + 1
                                        If dsCanBo = "" Then
                                            vIDCB_previous = vIdCB
                                            dsCanBo = "'" & vIdCB & "'"
                                        Else
                                            If vIDCB_previous <> vIdCB Then dsCanBo = dsCanBo & ",'" & vIdCB & "'"
                                        End If
                                        vIDCB_previous = vIdCB
                                    End If

                                Else
                                    GoTo nex
                                End If
                            End If
nex:
                            row = row + 1
                        End If
                    End While
                Else
                    strErr = "Thông tin trong Sheet sử dụng để nhập các Quyết định khác của cán bộ không đúng" & vbCrLf
                End If
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
            MsgBox(ex.Message)
        End Try

        SetProgress(8)
        If dsCanBo <> "" Then
            Dim strSql As String = " SELECT t1.idCanbo, MaCB, HoTen, DVKyHDLD, NoiLamViec, (SELECT ten_goi FROM DanhMuc WHERE Id=IdLoaiHD) as 'LoaiHD' " & _
                                           " FROM HS_HDLD t1,  HS_CanBo t2 WHERE t1.idcanbo= t2.idCanbo and t1.idCanbo in (" & dsCanBo & ") order by t1.idCanbo, TuNgay desc "
            Dim dt As DataTable
            Dim j As Integer = 0
            dt = dbconn.SelectDBRows(strSql)
            For j = 0 To dt.Rows.Count - 1
                myDataGrid.Rows.Add()
                myDataGrid.Rows(j).Cells("STT").Value = j + 1
                myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
                myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
                myDataGrid.Rows(j).Cells("DVKyHDLD").Value = dt.Rows(j)("DVKyHDLD").ToString()
                myDataGrid.Rows(j).Cells("NoiLamViec").Value = dt.Rows(j)("NoiLamViec").ToString()
                myDataGrid.Rows(j).Cells("LoaiHD").Value = dt.Rows(j)("LoaiHD").ToString()
            Next
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: " & countCB.ToString & " hồ sơ."
        Else
            labDS.Text = "Danh sách Hồ sơ được import vào CSDL: 0 hồ sơ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show("Danh sách Hồ sơ không import được: " & vbCr & strErr, "Thông báo")
        End If

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    Private Sub cmdChonTM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChonTM.Click
        Dim openFile As New OpenFileDialog
        HideProgressBar()
        Try
            openFile.Filter = "Excel Files (*.xls)|*.xls"
            openFile.ShowDialog()
            txtPath.Text = openFile.FileName
        Catch ex As Exception

        End Try
    End Sub

   
End Class