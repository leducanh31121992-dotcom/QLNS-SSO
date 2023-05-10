Option Explicit On
'Option Strict On

Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Text
Imports System
Imports System.Reflection

Public Class frmImportExcelCB

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


    'Đọc nội dung từ file excel vao datagrid
    Private Sub ImportExcel_CU(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)

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
                        's = s & .cells(1, "A").value & ";" & .cells(1, "B").value & ";" & .cells(1, "B").value
                        'Kiem tra neu Mã can bo và CMT neu da co thi ko insert vso nua
                        If db.getString("SELECT CMT_so FROM HS_canbo WHERE CMT_so='" & .cells(row, "C").value.ToString.Trim & "'") <> "" Then
                            strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").value.ToString.Trim & " đã tồn tại." & vbCrLf
                            GoTo nex
                        End If

                        If db.getString("SELECT MaCB FROM HS_canbo WHERE MaCB='" & .cells(row, "A").value.ToString.Trim & "'") <> "" Then
                            strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").value.ToString.Trim & " có mã cán bộ đã tồn tại." & vbCrLf
                            GoTo nex
                        End If
                        vIdCB = ""
                        '''''Insert HS can bo'''''
                        Try

                            Dim m_HSCanBo As clsHS_CanBo.HS_CanBo = New clsHS_CanBo.HS_CanBo()
                            m_HSCanBo.MaCB = .cells(row, "A").value.ToString.Trim
                            m_HSCanBo.HoTen = standardizeName(.cells(row, "B").value.ToString.Trim)
                            m_HSCanBo.TenThuongGoi = m_HSCanBo.HoTen
                            m_HSCanBo.BiDanh = m_HSCanBo.HoTen
                            If .cells(row, "D").value Is Nothing Then
                                m_HSCanBo.GioiTinh = 1
                            Else
                                m_HSCanBo.GioiTinh = .cells(row, "D").value
                            End If
                            If .cells(row, "N").value Is Nothing Then
                                m_HSCanBo.NgaySinh = DateTime.Parse("01/01/1900")
                            Else
                                m_HSCanBo.NgaySinh = .cells(row, "N").value
                            End If
                            m_HSCanBo.IdDonVi = IdDONVI
                            m_HSCanBo.IdQuocTich = 1
                            If .cells(row, "S").value Is Nothing Then
                                m_HSCanBo.IdDanToc = 0
                            Else
                                m_HSCanBo.IdDanToc = CInt(.cells(row, "S").value)
                            End If
                            If .cells(row, "T").value Is Nothing Then
                                m_HSCanBo.IdTonGiao = 0
                            Else
                                m_HSCanBo.IdTonGiao = CInt(.cells(row, "T").value)
                            End If

                            m_HSCanBo.CMT_So = .cells(row, "C").value.ToString.Trim
                            If .cells(row, "AI").value Is Nothing Then
                                m_HSCanBo.CMT_NgayCap = DateTime.Parse("01/01/1900")
                            Else
                                m_HSCanBo.CMT_NgayCap = .cells(row, "AI").value
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
                            m_HSCanBo.TTr_Diachi = ""
                            m_HSCanBo.TTr_Dienthoai = ""
                            m_HSCanBo.DienThoai_CQ = ""
                            m_HSCanBo.DienThoai_DD = ""
                            m_HSCanBo.DienThoai_NR = ""
                            m_HSCanBo.SoFax = ""
                            m_HSCanBo.Email = ""
                            m_HSCanBo.NhomMau = ""
                            m_HSCanBo.IdThanhPhanGD = 0

                            If .cells(row, "U").value Is Nothing Then
                                m_HSCanBo.IdUT_BThan = ""
                            Else
                                m_HSCanBo.IdUT_BThan = .cells(row, "U").value.ToString.Trim & ";"
                            End If
                            If .cells(row, "V").value Is Nothing Then
                                m_HSCanBo.IdUT_GDinh = 0
                            Else
                                If .cells(row, "V").value.ToString.Trim = "" Then
                                    m_HSCanBo.IdUT_GDinh = 0
                                Else
                                    m_HSCanBo.IdUT_GDinh = CInt(.cells(row, "V").value)
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
                                m_HSCanBo.Ngay_NH = .cells(row, "AG").value
                            End If
                            If .cells(row, "AH").value Is Nothing Then
                                m_HSCanBo.Ngay_ThamNien = DateTime.Parse("01/01/1900")
                                m_HSCanBo.Ngay_VBSP = DateTime.Parse("01/01/1900")
                                m_HSCanBo.Ngay_CQ = DateTime.Parse("01/01/1900")
                            Else
                                m_HSCanBo.Ngay_VBSP = .cells(row, "AH").value
                                m_HSCanBo.Ngay_ThamNien = m_HSCanBo.Ngay_VBSP
                                m_HSCanBo.Ngay_CQ = m_HSCanBo.Ngay_VBSP
                            End If
                            m_HSCanBo.Ngay_BienChe = DateTime.Parse("01/01/1900")
                            If .cells(row, "Y").value Is Nothing Then
                                m_HSCanBo.CM_Ngay = DateTime.Parse("01/01/1900")
                            Else
                                m_HSCanBo.CM_Ngay = .cells(row, "Y").value
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
                            Dim strCode As String = ""
                            strCode = _HS_CanBo.Insert_Human(m_HSCanBo)
                            vIdCB = m_HSCanBo.IdCanBo
                            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu

                            ' Neu insert vao HSCanbo Ok thi insert tiep vao HS QD Nhan su
                            If vIdCB <> "" Then
                                If .cells(row, "E").value Is Nothing Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").value.ToString.Trim & " nhập sai Ngày vào đơn vị hiện đang công tác." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If .cells(row, "F").value Is Nothing Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " nhập sai Loại quyết định nhân sự. (ID Loại quyết định = 321/322/323/657/658)." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If .cells(row, "G").value Is Nothing Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chưa nhập số Quyết định vào đơn vị hiện tại." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If .cells(row, "J").value Is Nothing Or .cells(row, "K").value Is Nothing Or .cells(row, "L").value Is Nothing Or .cells(row, "M").value Is Nothing Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chưa nhập đầy đủ ID đơn vị(PGD)/Phòng(tổ)/Chức vụ/Chuyên môn." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                ' Kiem tra Loai quyet dinh co thuoc danh sach QD khong
                                If Not (CInt(.cells(row, "F").value) = 321 Or CInt(.cells(row, "F").value) = 322 Or CInt(.cells(row, "F").value) = 323 Or CInt(.cells(row, "F").value) = 657 Or CInt(.cells(row, "F").value) = 658) Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai Loại quyết định nhân sự. (ID Loại quyết định = 321/322/323/657/658)" & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If db.getNumber("SELECT [ID] FROM Chinhanh WHERE (id=" & IdDONVI & " or id_goc=" & IdDONVI & ") AND id=" & CInt(.cells(row, "J").value) & " AND status=1") <= 0 Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Hội sở tỉnh hoặc ID của PGD hiện đang công tác." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If Not ((IdDONVI >= 6 And 15 <= CInt(.cells(row, "K").value) And CInt(.cells(row, "K").value) <= 24) Or (IdDONVI <= 5) And ((1 <= CInt(.cells(row, "K").value) And CInt(.cells(row, "K").value) <= 14) Or (25 <= CInt(.cells(row, "K").value) And CInt(.cells(row, "K").value) <= 37))) Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Phòng hoặc tổ hiện đang công tác." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If Not ((304 <= CInt(.cells(row, "L").value) And CInt(.cells(row, "L").value) <= 309) Or CInt(.cells(row, "L").value) = 320 Or CInt(.cells(row, "L").value) = 706) Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Chức vụ hiện đang công tác." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If
                                If Not ((240 <= CInt(.cells(row, "M").value) And CInt(.cells(row, "M").value) <= 263) Or CInt(.cells(row, "M").value) = 271 Or CInt(.cells(row, "M").value) = 282) Then
                                    strErr = strErr & "Cán bộ có số CMT: " & .cells(row, "C").ToString.Trim & " chọn sai ID của Chuyên môn hiện tại." & vbCrLf
                                    _HS_CanBo.Delete_Human(vIdCB)
                                    GoTo nex
                                End If

                                m_QDNhanSu.IdQDNhanSu = ""
                                m_QDNhanSu.IdCanBo = vIdCB
                                m_QDNhanSu.So_QD = .cells(row, "G").value.ToString.Trim
                                m_QDNhanSu.NgayKy_QD = .cells(row, "E").value
                                m_QDNhanSu.NguoiKy_QD = standardizeName(.cells(row, "H").value.ToString.Trim)
                                m_QDNhanSu.IdLoaiQD = CInt(.cells(row, "F").value)
                                m_QDNhanSu.LoaiQD = ""
                                m_QDNhanSu.NgayHL = .cells(row, "E").value
                                If .cells(row, "I").value Is Nothing Then
                                    m_QDNhanSu.idCV_Nguoiky_QD = 0
                                Else
                                    m_QDNhanSu.idCV_Nguoiky_QD = .cells(row, "I").value
                                End If
                                m_QDNhanSu.CV_NguoiKy_QD = ""
                                m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
                                m_QDNhanSu.NgayThoiLuong = DateTime.MinValue
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
                                m_QDNhanSu.DenNgay = DateTime.MinValue
                                m_QDNhanSu.NoiDung = ""
                                m_QDNhanSu.Active = 1
                                m_QDNhanSu.GhiChu = ""
                                m_QDNhanSu.Add()
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
                dt1 = db.SelectDBRows("SELECT (SELECT ten_goi FROM ChiNhanh where Id=IdDonvi_Moi) as Chinhanh, (SELECT ten_phong FROM PhongBan WHERE Id=IdPhong_Moi) as Phong FROM QDNhansu WHERE IdCanBo='" & dt.Rows(j)("IdCanBo").ToString() & "'")
                myDataGrid.Rows(j).Cells("Chinhanh").Value = dt1.Rows(0)("Chinhanh").ToString()
                myDataGrid.Rows(j).Cells("Phong").Value = dt1.Rows(0)("Phong").ToString()
            Next
            labDS.Text = "Danh sách cán bộ được import vào CSDL: " & countCB.ToString & " cán bộ."
        Else
            labDS.Text = "Danh sách cán bộ được import vào CSDL: 0 cán bộ."
        End If

        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show(strErr, "Thông báo")
        End If

        'Dim exsheet As Excel.Worksheets("fdf").Select()

        'exsheet = EXL.Workbooks.Open("C: est.xls").Worksheets("Sheet1")
        'exsheet = EXL.ActiveSheet
        'Dim cell As String
        'Dim x As Integer
        'x = 1
        ''Place contents of column into a ListBox 
        'Do While EXL.ActiveSheet.Range("B" & x).Value() > 0
        '    cell = EXL.ActiveSheet.Range("B" & x).Value()
        '    ListBox1.AddItem(cell)
        '    x = x + 1

        'Loop
        ''Close Excel   
        'EXL.Quit()
        ''Close the spreadsheet, otherwise it will be locked  
        'EXL = Nothing
        'exsheet = Nothing


        '        Dim MyConnection As System.Data.OleDb.OleDbConnection
        '        Try
        '            Dim DtSet As System.Data.DataSet
        '            Dim DtTable As System.Data.DataTable
        '            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

        '            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
        '                            "data source='" & PathExcelFile & " '; " & "Extended Properties=Excel 8.0;")

        '            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [DSCanBo$] ", MyConnection)
        '            MyCommand.TableMappings.Add("Table", "Attendence")
        '            DtSet = New System.Data.DataSet
        '            MyCommand.Fill(DtSet)
        '            SetProgress(6)
        '            'myDataGrid.DataSource = DtSet.Tables(0)
        '            DtTable = DtSet.Tables(0)
        '            MyConnection.Close()

        '            Dim i As Integer
        '            Dim countCB As Integer = 0
        '            Dim _HS_CanBo As clsHS_CanBo = New clsHS_CanBo
        '            Dim dsCanBo As String = ""
        '            Dim db As DBAccess = New DBAccess
        '            Dim strErr As String = ""
        '            Dim vIdCB As String = ""

        '            For i = 0 To DtTable.Rows.Count - 1
        '                If DtTable.Rows(i)(0).ToString.Trim <> "" And DtTable.Rows(i)(1).ToString.Trim <> "" And DtTable.Rows(i)(2).ToString.Trim <> "" Then
        '                    'Kiem tra neu Mã can bo và CMT neu da co thi ko insert vso nua
        '                    If db.getString("SELECT CMT_so FROM HS_canbo WHERE CMT_so='" & DtTable.Rows(i)(2).ToString.Trim & "'") <> "" Then
        '                        strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " đã tồn tại." & vbCrLf
        '                        GoTo nex
        '                    End If

        '                    If db.getString("SELECT MaCB FROM HS_canbo WHERE MaCB='" & DtTable.Rows(i)(2).ToString.Trim & "'") <> "" Then
        '                        strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " có mã cán bộ đã tồn tại." & vbCrLf
        '                        GoTo nex
        '                    End If
        '                    vIdCB = ""
        '                    '''''Insert HS can bo'''''
        '                    Try

        '                        Dim m_HSCanBo As clsHS_CanBo.HS_CanBo = New clsHS_CanBo.HS_CanBo()
        '                        m_HSCanBo.MaCB = DtTable.Rows(i)(0)
        '                        m_HSCanBo.HoTen = standardizeName(DtTable.Rows(i)(1))
        '                        m_HSCanBo.TenThuongGoi = m_HSCanBo.HoTen
        '                        m_HSCanBo.BiDanh = m_HSCanBo.HoTen
        '                        If DtTable.Rows(i)(3) Is DBNull.Value Then
        '                            m_HSCanBo.GioiTinh = 1
        '                        Else
        '                            m_HSCanBo.GioiTinh = DtTable.Rows(i)(3)
        '                        End If
        '                        If DtTable.Rows(i)(4) Is DBNull.Value Then
        '                            m_HSCanBo.NgaySinh = DateTime.Parse("01/01/1900")
        '                        Else
        '                            m_HSCanBo.NgaySinh = DtTable.Rows(i)(4)
        '                        End If
        '                        m_HSCanBo.IdDonVi = IdDONVI
        '                        m_HSCanBo.IdQuocTich = 1
        '                        If DtTable.Rows(i)(9) Is DBNull.Value Then
        '                            m_HSCanBo.IdDanToc = 0
        '                        Else
        '                            m_HSCanBo.IdDanToc = DtTable.Rows(i)(9)
        '                        End If
        '                        If DtTable.Rows(i)(10) Is DBNull.Value Then
        '                            m_HSCanBo.IdTonGiao = 0
        '                        Else
        '                            m_HSCanBo.IdTonGiao = DtTable.Rows(i)(10)
        '                        End If

        '                        m_HSCanBo.CMT_So = DtTable.Rows(i)(2)
        '                        If DtTable.Rows(i)(11) Is DBNull.Value Then
        '                            m_HSCanBo.CMT_NgayCap = DateTime.Parse("01/01/1900")
        '                        Else
        '                            m_HSCanBo.CMT_NgayCap = DtTable.Rows(i)(11)
        '                        End If
        '                        If DtTable.Rows(i)(12) Is DBNull.Value Then
        '                            m_HSCanBo.CMT_NoiCap = ""
        '                        Else
        '                            m_HSCanBo.CMT_NoiCap = DtTable.Rows(i)(12).ToString.Trim
        '                        End If

        '                        m_HSCanBo.IdNS_Tinh = 0
        '                        m_HSCanBo.IdNS_Huyen = 0
        '                        If DtTable.Rows(i)(5) Is DBNull.Value Then
        '                            m_HSCanBo.NS_DChi = ""
        '                        Else
        '                            m_HSCanBo.NS_DChi = DtTable.Rows(i)(5).ToString.Trim
        '                        End If

        '                        m_HSCanBo.IdNQ_Tinh = 0
        '                        m_HSCanBo.IdNQ_Huyen = 0
        '                        If DtTable.Rows(i)(6) Is DBNull.Value Then
        '                            m_HSCanBo.NQ_DChi = ""
        '                        Else
        '                            m_HSCanBo.NQ_DChi = DtTable.Rows(i)(6).ToString.Trim
        '                        End If

        '                        m_HSCanBo.IdThT_Tinh = 0
        '                        m_HSCanBo.IdThT_Huyen = 0
        '                        If DtTable.Rows(i)(7) Is DBNull.Value Then
        '                            m_HSCanBo.ThT_Diachi = ""
        '                        Else
        '                            m_HSCanBo.ThT_Diachi = DtTable.Rows(i)(7).ToString.Trim
        '                        End If
        '                        If DtTable.Rows(i)(8) Is DBNull.Value Then
        '                            m_HSCanBo.ThT_Dienthoai = ""
        '                        Else
        '                            m_HSCanBo.ThT_Dienthoai = DtTable.Rows(i)(8).ToString.Trim
        '                        End If

        '                        m_HSCanBo.IdTTr_Huyen = 0
        '                        m_HSCanBo.TTr_Diachi = ""
        '                        m_HSCanBo.TTr_Dienthoai = ""
        '                        m_HSCanBo.DienThoai_CQ = ""
        '                        m_HSCanBo.DienThoai_DD = ""
        '                        m_HSCanBo.DienThoai_NR = ""
        '                        m_HSCanBo.SoFax = ""
        '                        m_HSCanBo.Email = ""
        '                        m_HSCanBo.NhomMau = ""
        '                        m_HSCanBo.IdThanhPhanGD = 0
        '                        If DtTable.Rows(i)(13) Is DBNull.Value Then
        '                            m_HSCanBo.IdUT_BThan = ""
        '                        Else
        '                            m_HSCanBo.IdUT_BThan = DtTable.Rows(i)(13).ToString.Trim & ";"
        '                        End If
        '                        If DtTable.Rows(i)(14) Is DBNull.Value Then
        '                            m_HSCanBo.IdUT_GDinh = 0
        '                        Else
        '                            m_HSCanBo.IdUT_GDinh = DtTable.Rows(i)(14)
        '                        End If
        '                        If DtTable.Rows(i)(15) Is DBNull.Value Then
        '                            m_HSCanBo.DacDiem_BT = 0
        '                        Else
        '                            m_HSCanBo.DacDiem_BT = DtTable.Rows(i)(15).ToString.Trim
        '                        End If
        '                        If DtTable.Rows(i)(16) Is DBNull.Value Then
        '                            m_HSCanBo.QuanHe_Nguoi_NN = 0
        '                        Else
        '                            m_HSCanBo.QuanHe_Nguoi_NN = DtTable.Rows(i)(16).ToString.Trim
        '                        End If

        '                        If DtTable.Rows(i)(20) Is DBNull.Value Then
        '                            m_HSCanBo.IdTrinhDoVH = 0
        '                        Else
        '                            m_HSCanBo.IdTrinhDoVH = DtTable.Rows(i)(20)
        '                        End If
        '                        If DtTable.Rows(i)(23) Is DBNull.Value Then
        '                            m_HSCanBo.IdTrinhDoCT = 0
        '                        Else
        '                            m_HSCanBo.IdTrinhDoCT = DtTable.Rows(i)(23)
        '                        End If

        '                        If DtTable.Rows(i)(27) Is DBNull.Value Then
        '                            m_HSCanBo.Ngay_ThamNien = DateTime.Parse("01/01/1900")
        '                        Else
        '                            'm_HSCanBo.Ngay_ThamNien = DateTimeUtil.getDate(DtTable.Rows(i)(27))
        '                            m_HSCanBo.Ngay_ThamNien = DtTable.Rows(i)(27)
        '                        End If
        '                        If DtTable.Rows(i)(28) Is DBNull.Value Then
        '                            m_HSCanBo.Ngay_NH = DateTime.Parse("01/01/1900")
        '                        Else
        '                            'm_HSCanBo.Ngay_NH = DateTimeUtil.getDate(DtTable.Rows(i)(27))
        '                            m_HSCanBo.Ngay_NH = DtTable.Rows(i)(27)
        '                        End If
        '                        If DtTable.Rows(i)(29) Is DBNull.Value Then
        '                            m_HSCanBo.Ngay_VBSP = DateTime.Parse("01/01/1900")
        '                        Else
        '                            'm_HSCanBo.Ngay_VBSP = DateTimeUtil.getDate(DtTable.Rows(i)(29))
        '                            m_HSCanBo.Ngay_VBSP = DtTable.Rows(i)(29)
        '                        End If
        '                        m_HSCanBo.Ngay_BienChe = DateTime.Parse("01/01/1900")
        '                        'm_HSCanBo.Ngay_CQ = DateTimeUtil.getDate(DtTable.Rows(i)(30))
        '                        m_HSCanBo.Ngay_CQ = DtTable.Rows(i)(30)
        '                        If DtTable.Rows(i)(17) Is DBNull.Value Then
        '                            m_HSCanBo.CM_Ngay = DateTime.Parse("01/01/1900")
        '                        Else
        '                            'm_HSCanBo.CM_Ngay = DateTimeUtil.getDate(DtTable.Rows(i)(17))
        '                            m_HSCanBo.CM_Ngay = DtTable.Rows(i)(17)
        '                        End If
        '                        m_HSCanBo.CM_ToChuc = ""
        '                        m_HSCanBo.SoTruong_CT = ""
        '                        m_HSCanBo.CV_Lau = ""
        '                        m_HSCanBo.NH_MaKH = ""
        '                        m_HSCanBo.NH_SoTK = ""
        '                        m_HSCanBo.NH_TenNH = ""
        '                        If DtTable.Rows(i)(24) Is DBNull.Value Then
        '                            m_HSCanBo.MaSoThue = 0
        '                        Else
        '                            m_HSCanBo.MaSoThue = DtTable.Rows(i)(24).ToString.Trim
        '                        End If
        '                        If DtTable.Rows(i)(25) Is DBNull.Value Then
        '                            m_HSCanBo.BHXH_SoSo = 0
        '                        Else
        '                            m_HSCanBo.BHXH_SoSo = DtTable.Rows(i)(25).ToString.Trim
        '                        End If

        '                        m_HSCanBo.BHXH_NgaySo = DateTime.Parse("01/01/1900")
        '                        m_HSCanBo.BHXH_NgayBatDau = DateTime.Parse("01/01/1900")
        '                        If DtTable.Rows(i)(26) Is DBNull.Value Then
        '                            m_HSCanBo.BHXH_NoiLam = ""
        '                        Else
        '                            m_HSCanBo.BHXH_NoiLam = DtTable.Rows(i)(26).ToString.Trim
        '                        End If

        '                        m_HSCanBo.GhiChu = ""
        '                        m_HSCanBo.IdNew = ""
        '                        m_HSCanBo.AnhThe = ""
        '                        Dim strCode As String = ""
        '                        strCode = _HS_CanBo.Insert_Human(m_HSCanBo)
        '                        vIdCB = m_HSCanBo.IdCanBo
        '                        Dim m_QDNhanSu As QDNhanSu = New QDNhanSu

        '                        ' Neu insert vao HSCanbo Ok thi insert tiep vao HS QD Nhan su
        '                        'If vIdCB <> "" And CInt(DtTable.Rows(i)(31)) >= 0 And DtTable.Rows(i)(32).ToString.Trim <> "" And CInt(DtTable.Rows(i)(35)) > 0 And CInt(DtTable.Rows(i)(36)) > 0 And CInt(DtTable.Rows(i)(37)) > 0 And CInt(DtTable.Rows(i)(38)) > 0 Then
        '                        If vIdCB <> "" Then

        '                            If DtTable.Rows(i)(30) Is DBNull.Value Then
        '                                strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " nhập sai Ngày vào đơn vị hiện đang công tác." & vbCrLf
        '                                _HS_CanBo.Delete_Human(vIdCB)
        '                                GoTo nex
        '                            End If
        '                            If DtTable.Rows(i)(31) Is DBNull.Value Then
        '                                strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " nhập sai Loại quyết định nhân sự. (ID Loại quyết định = 321/322/323/657/658)." & vbCrLf
        '                                _HS_CanBo.Delete_Human(vIdCB)
        '                                GoTo nex
        '                            End If
        '                            If DtTable.Rows(i)(32) Is DBNull.Value Or DtTable.Rows(i)(32).ToString.Trim = "" Then
        '                                strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chưa nhập số Quyết định vào đơn vị hiện tại." & vbCrLf
        '                                _HS_CanBo.Delete_Human(vIdCB)
        '                                GoTo nex
        '                            End If
        '                            If DtTable.Rows(i)(35) Is DBNull.Value Or DtTable.Rows(i)(36) Is DBNull.Value Or DtTable.Rows(i)(37) Is DBNull.Value Or DtTable.Rows(i)(38) Is DBNull.Value Then
        '                                strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chưa nhập đầy đủ ID đơn vị(PGD)/Phòng(tổ)/Chức vụ/Chuyên môn." & vbCrLf
        '                                _HS_CanBo.Delete_Human(vIdCB)
        '                                GoTo nex
        '                            End If
        '                            If Not (DtTable.Rows(i)(31) Is DBNull.Value And DtTable.Rows(i)(32) Is DBNull.Value And DtTable.Rows(i)(35) Is DBNull.Value And DtTable.Rows(i)(36) Is DBNull.Value And DtTable.Rows(i)(37) Is DBNull.Value And DtTable.Rows(i)(38) Is DBNull.Value) Then
        '                                ' Kiem tra Loai quyet dinh co thuoc danh sach QD khong
        '                                If Not (CInt(DtTable.Rows(i)(31)) = 321 Or CInt(DtTable.Rows(i)(31)) = 322 Or CInt(DtTable.Rows(i)(31)) = 323 Or CInt(DtTable.Rows(i)(31)) = 657 Or CInt(DtTable.Rows(i)(31)) = 658) Then
        '                                    strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chọn sai Loại quyết định nhân sự. (ID Loại quyết định = 321/322/323/657/658)" & vbCrLf
        '                                    _HS_CanBo.Delete_Human(vIdCB)
        '                                    GoTo nex
        '                                End If
        '                                If db.getNumber("SELECT [ID] FROM Chinhanh WHERE (id=" & IdDONVI & " or id_goc=" & IdDONVI & ") AND id=" & CInt(DtTable.Rows(i)(35)) & " AND status=1") <= 0 Then
        '                                    strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chọn sai ID của Hội sở tỉnh hoặc ID của PGD hiện đang công tác." & vbCrLf
        '                                    _HS_CanBo.Delete_Human(vIdCB)
        '                                    GoTo nex
        '                                End If
        '                                If Not ((IdDONVI >= 6 And 15 <= CInt(DtTable.Rows(i)(36)) And CInt(DtTable.Rows(i)(36)) <= 24) Or (IdDONVI <= 5) And ((1 <= CInt(DtTable.Rows(i)(36)) And CInt(DtTable.Rows(i)(36)) <= 14) Or (25 <= CInt(DtTable.Rows(i)(36)) And CInt(DtTable.Rows(i)(36)) <= 37))) Then
        '                                    strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chọn sai ID của Phòng hoặc tổ hiện đang công tác." & vbCrLf
        '                                    _HS_CanBo.Delete_Human(vIdCB)
        '                                    GoTo nex
        '                                End If
        '                                If Not ((304 <= CInt(DtTable.Rows(i)(37)) And CInt(DtTable.Rows(i)(37)) <= 309) Or CInt(DtTable.Rows(i)(37)) = 320 Or CInt(DtTable.Rows(i)(37)) = 706) Then
        '                                    strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chọn sai ID của Chức vụ hiện đang công tác." & vbCrLf
        '                                    _HS_CanBo.Delete_Human(vIdCB)
        '                                    GoTo nex
        '                                End If
        '                                If Not ((240 <= CInt(DtTable.Rows(i)(38)) And CInt(DtTable.Rows(i)(38)) <= 263) Or CInt(DtTable.Rows(i)(38)) = 271 Or CInt(DtTable.Rows(i)(38)) = 282) Then
        '                                    strErr = strErr & "Cán bộ có số CMT: " & DtTable.Rows(i)(2).ToString.Trim & " chọn sai ID của Chuyên môn hiện tại." & vbCrLf
        '                                    _HS_CanBo.Delete_Human(vIdCB)
        '                                    GoTo nex
        '                                End If


        '                                'Dim dateHL As Date
        '                                'dateHL = DateTimeUtil.getDate(DtTable.Rows(i)(30))
        '                                m_QDNhanSu.IdQDNhanSu = ""
        '                                m_QDNhanSu.IdCanBo = vIdCB
        '                                m_QDNhanSu.So_QD = DtTable.Rows(i)(32)
        '                                m_QDNhanSu.NgayKy_QD = DtTable.Rows(i)(30)
        '                                m_QDNhanSu.NguoiKy_QD = standardizeName(DtTable.Rows(i)(32))
        '                                m_QDNhanSu.IdLoaiQD = CInt(DtTable.Rows(i)(31))
        '                                m_QDNhanSu.NgayHL = DtTable.Rows(i)(30)
        '                                If DtTable.Rows(i)(34) Is DBNull.Value Then
        '                                    m_QDNhanSu.idCV_Nguoiky_QD = 0
        '                                Else
        '                                    m_QDNhanSu.idCV_Nguoiky_QD = DtTable.Rows(i)(34)
        '                                End If
        '                                m_QDNhanSu.NgayBoNhiem_TT = DateTime.MinValue
        '                                m_QDNhanSu.NgayThoiLuong = DateTime.MinValue
        '                                m_QDNhanSu.IdDonvi_Cu = 0
        '                                m_QDNhanSu.IdPhong_Cu = 0
        '                                m_QDNhanSu.IdChucvu_Cu = 0
        '                                m_QDNhanSu.IdChuyenMon_Cu = 0
        '                                m_QDNhanSu.IdDonvi_Moi = CInt(DtTable.Rows(i)(35))
        '                                m_QDNhanSu.IdPhong_Moi = CInt(DtTable.Rows(i)(36))
        '                                m_QDNhanSu.IdChucvu_Moi = CInt(DtTable.Rows(i)(37))
        '                                m_QDNhanSu.IdChuyenMon_Moi = CInt(DtTable.Rows(i)(38))
        '                                m_QDNhanSu.IsQD_NHCS = 1
        '                                m_QDNhanSu.DVraQD = getDonvi(m_QDNhanSu.IdDonvi_Moi)
        '                                m_QDNhanSu.DenNgay = DateTime.MinValue
        '                                m_QDNhanSu.NoiDung = ""
        '                                m_QDNhanSu.Active = 1
        '                                m_QDNhanSu.GhiChu = ""
        '                                m_QDNhanSu.Add()
        '                            End If
        '                        End If
        '                        ' Insert tiep vao HS Dang

        '                        ' OK
        '                        If m_QDNhanSu.IdQDNhanSu <> "" Then
        '                            countCB = countCB + 1
        '                            If dsCanBo = "" Then
        '                                dsCanBo = "'" & vIdCB & "'"
        '                            Else
        '                                dsCanBo = dsCanBo & ",'" & vIdCB & "'"
        '                            End If
        '                        Else
        '                            _HS_CanBo.Delete_Human(vIdCB)
        '                        End If

        '                    Catch ex As Exception
        '                        'Dim arrDel() As String
        '                        'arrDel = dsCanBo.Split(",")
        '                        'Dim idxDel As Integer = 0
        '                        'For idxDel = 0 To arrDel.Length - 1
        '                        '    _HS_CanBo.Delete_Human(arrDel(idxDel).Substring(1, 15))
        '                        'Next
        '                        'dsCanBo = ""
        '                        _HS_CanBo.Delete_Human(vIdCB)
        '                    End Try
        '                End If
        'nex:
        '            Next
        '            SetProgress(8)
        '            If dsCanBo <> "" Then

        '                Dim arrDel() As String
        '                arrDel = dsCanBo.Split(",")

        '                Dim sqlDSCanBo As String = "SELECT * FROM HS_CanBo WHERE IdCanbo in (" & dsCanBo & ")"
        '                Dim dt As DataTable
        '                Dim j As Integer = 0
        '                dt = db.SelectDBRows(sqlDSCanBo)
        '                'myDataGrid.Rows.Add()
        '                For j = 0 To dt.Rows.Count - 1
        '                    myDataGrid.Rows.Add()
        '                    myDataGrid.Rows(j).Cells("STT").Value = j + 1
        '                    myDataGrid.Rows(j).Cells("MaCB").Value = IIf(dt.Rows(j)("MaCB").ToString() <> "", dt.Rows(j)("MaCB").ToString(), "")
        '                    myDataGrid.Rows(j).Cells("HoTen").Value = IIf(dt.Rows(j)("HoTen").ToString() <> "", dt.Rows(j)("HoTen").ToString(), "")
        '                    myDataGrid.Rows(j).Cells("GioiTinh").Value = IIf(dt.Rows(j)("GioiTinh").ToString() = False, "Nam", "Nữ")
        '                    If dt.Rows(j)("NgaySinh").ToString().Trim() <> "" Then
        '                        myDataGrid.Rows(j).Cells("NgaySinh").Value = CType(dt.Rows(j)("NgaySinh").ToString(), DateTime).ToString("dd-MM-yyyy")
        '                    Else
        '                        myDataGrid.Rows(j).Cells("NgaySinh").Value = ""
        '                    End If
        '                    myDataGrid.Rows(j).Cells("CMT_So").Value = dt.Rows(j)("CMT_So").ToString()

        '                    Dim dt1 As DataTable
        '                    dt1 = db.SelectDBRows("SELECT (SELECT ten_goi FROM ChiNhanh where Id=IdDonvi_Moi) as Chinhanh, (SELECT ten_phong FROM PhongBan WHERE Id=IdPhong_Moi) as Phong FROM QDNhansu WHERE IdCanBo='" & dt.Rows(j)("IdCanBo").ToString() & "'")
        '                    myDataGrid.Rows(j).Cells("Chinhanh").Value = dt1.Rows(0)("Chinhanh").ToString()
        '                    myDataGrid.Rows(j).Cells("Phong").Value = dt1.Rows(0)("Phong").ToString()
        '                    'If dt.Rows(j)("CMT_NgayCap").ToString().Trim() <> "" Then
        '                    '    myDataGrid.Rows(j).Cells("CMT_NgayCap").Value = CType(dt.Rows(j)("CMT_NgayCap").ToString(), DateTime).ToString("dd-MM-yyyy")
        '                    'Else
        '                    '    myDataGrid.Rows(j).Cells("CMT_NgayCap").Value = ""
        '                    'End If
        '                    'myDataGrid.Rows(j).Cells("CMT_NoiCap").Value = dt.Rows(j)("CMT_NoiCap").ToString()
        '                Next
        '                labDS.Text = "Danh sách cán bộ được import vào CSDL: " & countCB.ToString & " cán bộ."
        '            End If

        '            SetProgress(10)
        '            If strErr <> "" Then
        '                MessageBox.Show(strErr, "Thông báo")
        '            End If

        '        Catch ex As Exception
        '            MyConnection.Close()
        '        End Try

        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each i As Process In pro
            i.Kill()
        Next

    End Sub

    'Đọc nội dung từ file excel vao datagrid
    Private Sub ImportExcel(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)

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

    Private Sub ImportExcel_QDNS_TruocNHCS(ByVal PathExcelFile As String, ByVal myDataGrid As DataGridView)

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
                .Worksheets(1).Select()

                Dim vIdCB As String = ""
                If .cells(1, "A").value.ToString.Trim = "QDNHANSU" Then
                    row = 3
                    While Not (.cells(row, "B").value Is Nothing)
                        If .cells(row, "B").value.ToString.Trim <> "" Then
                            Dim m_QDNhanSu As QDNhanSu = New QDNhanSu
                            Dim m_LuongOld As LuongCanBo = New LuongCanBo
                            Dim m_PhucapOld As Phucap = New Phucap
                            Dim vIdQDNhanSu As String = ""
                            Dim vIdLuong As String = ""
                            Dim vIdPhucap As String = ""
                            Dim _value As Integer = 0
                            vIdCB = getCanBo_ID(.cells(row, "B").value.ToString.Trim, False)
                            If vIdCB <> "" Then
                                m_QDNhanSu.IdQDNhanSu = ""
                                m_QDNhanSu.IdCanBo = vIdCB
                                If .cells(row, "D").value Is Nothing Then
                                    m_QDNhanSu.IsQD_NHCS = 0
                                Else
                                    _value = CInt(.cells(row, "D").value)
                                    If _value = 1 Then
                                        strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Chỉ áp dụng cho các quyết định nhân sự trước khi vào NHCSXH." & vbCrLf
                                        GoTo nex
                                    End If
                                End If

                                If .cells(row, "E").value Is Nothing Then
                                    m_QDNhanSu.So_QD = ""
                                Else
                                    m_QDNhanSu.So_QD = .cells(row, "E").value.ToString.Trim
                                End If

                                'kiem tra mot so dieu kien
                                If .cells(row, "I").value Is Nothing Then
                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Quyết định số " & m_QDNhanSu.So_QD & " Chưa nhập ID loại quyết định nhân sự." & vbCrLf
                                    GoTo nex
                                End If
                                _value = CInt(.cells(row, "I").value.ToString.Trim)
                                If _value = 0 Then
                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập ID Loại Quyết định nhân sự." & vbCrLf
                                    GoTo nex
                                Else
                                    If Not ((_value >= 321 And _value <= 327) Or _value = 657 Or _value = 658 Or _value = 688 Or _value = 711 Or (_value >= 714 And _value <= 717) Or _value = 748 Or _value = 768 Or (_value >= 770 And _value <= 774) Or _value = 808 Or _value = 809 Or _value = 812 Or _value = 813) Then
                                        strErr = strErr & "Giá trị ID Loại quyết định nhân sự không đúng. Hãy kiểm tra lại."
                                        GoTo nex
                                    End If
                                End If
                                If .cells(row, "L").value Is Nothing Then
                                    strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & ": Chưa nhập chức vụ, đơn vị công tác." & vbCrLf
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

                                If .cells(row, "Q").value Is Nothing Then
                                    m_QDNhanSu.NguoiKy_QD = ""
                                Else
                                    m_QDNhanSu.NguoiKy_QD = standardizeName(.cells(row, "Q").value.ToString.Trim)
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

                                m_QDNhanSu.IdDonvi_Cu = 0
                                m_QDNhanSu.IdPhong_Cu = 0
                                m_QDNhanSu.IdChucvu_Cu = 0
                                m_QDNhanSu.IdChuyenMon_Cu = 0
                                m_QDNhanSu.IdDonvi_Moi = 0
                                m_QDNhanSu.IdPhong_Moi = 0
                                m_QDNhanSu.IdChucvu_Moi = 0
                                If .cells(row, "K").value Is Nothing Then
                                    m_QDNhanSu.DenNgay = m_QDNhanSu.NgayKy_QD
                                Else
                                    If .cells(row, "K").value.GetType.Name.ToString.ToUpper = "DATETIME" Or .cells(row, "K").value.GetType.Name.ToString.ToUpper = "DATE" Then
                                        m_QDNhanSu.DenNgay = .cells(row, "K").value
                                    Else
                                        m_QDNhanSu.DenNgay = DateTimeUtil.getDate(.cells(row, "K").value)
                                    End If
                                End If

                                If .cells(row, "R").value Is Nothing Then
                                    m_QDNhanSu.CV_NguoiKy_QD = ""
                                Else
                                    m_QDNhanSu.CV_NguoiKy_QD = standardizeString(.cells(row, "R").value.ToString.Trim)
                                End If
                                m_QDNhanSu.idCV_Nguoiky_QD = 0
                                If .cells(row, "H").value Is Nothing Then
                                    m_QDNhanSu.LoaiQD = ""
                                Else
                                    m_QDNhanSu.LoaiQD = standardizeString(.cells(row, "H").value.ToString.Trim)
                                End If
                                If .cells(row, "L").value Is Nothing Then
                                    m_QDNhanSu.NoiDung = ""
                                Else
                                    m_QDNhanSu.NoiDung = standardizeString(.cells(row, "L").value.ToString.Trim)
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
                                If .cells(row, "S").value Is Nothing Then
                                    m_QDNhanSu.GhiChu = ""
                                Else
                                    m_QDNhanSu.GhiChu = standardizeString(.cells(row, "S").value.ToString.Trim)
                                End If

                                vIdLuong = m_LuongOld.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)
                                vIdPhucap = m_PhucapOld.getID(vIdCB, m_QDNhanSu.NgayHL, m_QDNhanSu.So_QD, m_QDNhanSu.NgayKy_QD)

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

                                    If Not (.cells(row, "O").value Is Nothing) Then
                                        If Not (.cells(row, "O").value.ToString.Trim = "" Or .cells(row, "O").value.ToString.Trim = "0") Then
                                            m_Luong.Ngay_Huong = DateTimeUtil.getDate(m_QDNhanSu.NgayHL)
                                            m_Luong.DVraQD = m_QDNhanSu.DVraQD

                                            'QD lương kem theo QD nhan su không thuoc NHCS
                                            If .cells(row, "P").value Is Nothing Then
                                                strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập thông tin về lương trong QĐ nhân sự." & vbCrLf
                                                m_QDNhanSu.Delete()
                                                GoTo nex
                                            End If
                                            If Not (.cells(row, "P").value.ToString.Trim <> "") Then
                                                strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập thông tin về lương trong QĐ nhân sự." & vbCrLf
                                                m_QDNhanSu.Delete()
                                                GoTo nex
                                            End If
                                            m_Luong.NgayLen_DK = Date.MinValue
                                            m_Luong.IdCV_Nguoi_QD = 0
                                            m_Luong.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                                            m_Luong.NoiDung = .cells(row, "P").value.ToString.Trim
                                            m_Luong.IsQD_NHCS = 0
                                            m_Luong.LoaiQD = m_QDNhanSu.LoaiQD

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
                                    If Not (.cells(row, "M").value Is Nothing) Then
                                        If Not (.cells(row, "M").value.ToString.Trim = "" Or .cells(row, "M").value.ToString.Trim = "0") Then
                                            If .cells(row, "N").value Is Nothing Then
                                                strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập Nội dung hưởng phụ cấp." & vbCrLf
                                                m_QDNhanSu.Delete()
                                                GoTo nex
                                            End If
                                            If Not (.cells(row, "N").value.ToString.Trim <> "") Then
                                                strErr = strErr & "Cán bộ mã: " & .cells(row, "B").value.ToString.Trim & " chưa nhập Nội dung hưởng phụ cấp." & vbCrLf
                                                m_QDNhanSu.Delete()
                                                GoTo nex
                                            End If
                                            m_PhuCap.IdCV_Nguoi_QD = 0
                                            m_PhuCap.CV_NguoiKy_QD = m_QDNhanSu.CV_NguoiKy_QD
                                            m_PhuCap.IsQD_NHCS = 0
                                            m_PhuCap.NoiDung = .cells(row, "N").value.ToString.Trim

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

    Private Sub ImportExcel_DaiMa(ByVal PathExcelFile As String)

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
                Dim col_end_CU As Integer = 0
                Dim daiDuPhong As Integer = 0
                Dim ID_CU As Integer = 0
                row = 1
                While .cells(row, "A").value.ToString.Trim <> ""
                    If row = 1 Then
                        ID_CU = .cells(row, "A").value
                        If ID_CU > 0 Then db.executeSQL("UPDATE ChiNhanh set MaDV='" & .cells(row, "C").value.ToString & "', MaCB_Begin= 1, MaCB_End=" & .cells(row, "D").value & " WHERE ID=" & .cells(row, "A").value)
                    Else
                        If .cells(row, "A").value = 99 Then
                            daiDuPhong = .cells(row, "D").value
                        Else
                            daiDuPhong = 0
                        End If
                        col_end_CU = db.getNumber("SELECT MaCB_end FROM Chinhanh WHERE id=" & ID_CU)
                        db.executeSQL("UPDATE ChiNhanh set MaDV='" & .cells(row, "C").value & "', MaCB_Begin= " & daiDuPhong + col_end_CU + 1 & ", MaCB_End=" & col_end_CU + .cells(row, "D").value & " WHERE ID=" & .cells(row, "A").value)
                        ID_CU = .cells(row, "A").value
                    End If
                    row = row + 1
                End While
            End With
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
            Excel = Nothing
        End Try
        SetProgress(10)
        If strErr <> "" Then
            MessageBox.Show(strErr, "Thông báo")
        End If

    End Sub

    Private Sub addCol(ByRef vGrid As DataGridView, ByVal vHeaderText As String, ByVal vDataPropertyName As String, ByVal vName As String)
        Dim colTxt As New DataGridViewTextBoxColumn()
        colTxt.DataPropertyName = vDataPropertyName
        colTxt.HeaderText = vHeaderText
        colTxt.Name = vName
        colTxt.ReadOnly = True
        vGrid.Columns.Add(colTxt)
    End Sub

    'Private Sub initGrid()
    '    gridDS_CB.Columns().Clear()
    '    gridDS_CB.Columns.Add("STT", "STT")
    '    gridDS_CB.Columns.Add("MaCB", "Mã Cán bộ")
    '    gridDS_CB.Columns.Add("HoTen", "Họ tên")
    '    gridDS_CB.Columns.Add("GioiTinh", "Giới tính")
    '    gridDS_CB.Columns.Add("NgaySinh", "Ngày sinh")
    '    gridDS_CB.Columns.Add("CMT_So", "Số CMT")
    '    gridDS_CB.Columns.Add("ChiNhanh", "Đơn vị")
    '    gridDS_CB.Columns.Add("Phong", "Phòng")
    '    'Căn chỉnh tiêu đề
    '    gridDS_CB.Columns("STT").Width = 50
    '    gridDS_CB.Columns("HoTen").Width = 188
    '    gridDS_CB.Columns("ChiNhanh").Width = 188
    '    gridDS_CB.Columns("Phong").Width = 188
    '    gridDS_CB.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    '    gridDS_CB.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    '    gridDS_CB.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    '    gridDS_CB.Columns(4).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter

    'End Sub

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
        End Select
        myDataGrid.Columns("STT").Width = 50
        myDataGrid.Columns("MaCB").Width = 100
        myDataGrid.Columns("HoTen").Width = 150
        myDataGrid.Columns(1).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        myDataGrid.Columns(2).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        myDataGrid.Columns(3).CellTemplate.Style.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
    End Sub

    Private Sub bntImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImport.Click
        Try
            ShowProgressBar()
            ProgressBar1.Maximum = 10
            SetProgress(1)
            Me.Refresh()

            'initGrid()
            'SetProgress(3)

            If DONVI = "000197" Then
                initGrid(1, gridDS_CB)
                SetProgress(3)
                ImportExcel_QDNS_TruocNHCS(txtPath.Text, gridDS_CB)
            Else
                initGrid(0, gridDS_CB)
                SetProgress(3)
                ImportExcel(txtPath.Text, gridDS_CB)
                'ImportExcel_DaiMa(txtPath.Text)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmImportExcelCB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If (e.KeyCode = Keys.Escape) Then
                Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Sử dụng phím tắt: " & ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button3)
            Return
        End Try
    End Sub

    Private Sub frmImportExcelCB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        HideProgressBar()
        If DONVI = gMaDonViTW Then
            bntImportMaCB.Visible = True
        Else
            bntImportMaCB.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntImportMaCB.Click
        Try
            ShowProgressBar()
            ProgressBar1.Maximum = 10
            SetProgress(1)
            Me.Refresh()

            initGrid(0, gridDS_CB)
            'initGrid()
            SetProgress(3)
            'ImportExcel(txtPath.Text, gridDS_CB)
            'ImportExcel_TCCB(txtPath.Text, gridDS_CB)
            ImportExcel_DaiMa(txtPath.Text)
        Catch ex As Exception

        End Try
    End Sub
End Class